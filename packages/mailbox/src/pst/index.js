// Outlook data files — .pst and .ost.
//
// The messaging layer: a folder tree, the messages in a folder, and one
// message with its recipients and attachments. Everything below is the node
// database ([ndb.js]) and the heap/table/property layer ([ltp.js]); this file
// only knows what a folder is.
//
// Two conventions carry most of the structure:
//
//   - A folder's tables are the folder's own node id with a different type in
//     the low five bits. The hierarchy table lists its subfolders, the contents
//     table its messages. So "list this folder" needs no search.
//   - A folder's contents table already holds the columns a message list wants
//     — subject, sender, date, size, flags. Opening 11,000 messages to draw a
//     list would be absurd; reading one table is not.
//
// Reading is lazy throughout. Pointing this at a 900 MB .ost builds two B-tree
// indexes and nothing else; message bodies are read when a message is opened.

import { Ndb, NdbError, NID, NID_TYPE, nidType, siblingNid, bufferReader, CRYPT_NAME } from './ndb.js';
import { Heap, readPc, readTc, nodePc, values } from './ltp.js';
import { MSG_FLAG, RECIPIENT_TYPE, PROPS } from '../props.js';
import { stripHtml } from '../mime.js';

export { Ndb, NdbError, bufferReader };

const NID_RECIPIENT_TABLE = 0x692;
const NID_ATTACHMENT_TABLE = 0x671;

/** Folders Outlook creates and hides; a person never asked for these. */
const HIDDEN_FOLDERS = new Set([
  'Top of Personal Folders',
  'Top of Outlook data file',
  'Root - Mailbox',
  'IPM_SUBTREE',
  'Search Root',
  'Common Views',
  'Finder',
  'Views',
  'Schedule',
  'Shortcuts',
  'Spooler Queue',
]);

const decodeText = (v) => (typeof v === 'string' ? v : v instanceof Uint8Array ? new TextDecoder('utf-8', { fatal: false }).decode(v) : null);

function addressFrom(props) {
  const name = props.senderName || props.sentRepresentingName || null;
  const smtp = props.senderSmtpAddress || props.sentRepresentingSmtpAddress || null;
  const legacy = props.senderEmailAddress || props.sentRepresentingEmailAddress || null;
  const address = smtp || (String(legacy || '').includes('@') ? legacy : '');
  if (!name && !address) return null;
  return { name: name || null, address: address || '', display: name && address ? `${name} <${address}>` : name || address };
}

export class PstFile {
  constructor(reader) {
    this.reader = reader;
    this.ndb = new Ndb(reader);
    this._folders = null;
  }

  static open(reader) {
    return new PstFile(reader);
  }

  static fromBytes(bytes) {
    return new PstFile(bufferReader(bytes));
  }

  get info() {
    const s = this.ndb.stats();
    return {
      kind: s.isOst ? 'ost' : 'pst',
      version: s.version,
      unicode: s.unicode,
      pageSize: s.pageSize,
      encoding: s.crypt,
      readable: s.readable,
      nodes: s.nodes,
      blocks: s.blocks,
      bytes: this.reader.size,
    };
  }

  /** The store's own properties — its display name, and where the tree starts. */
  store() {
    const node = this.ndb.node(NID.MESSAGE_STORE);
    if (!node) return {};
    return values(nodePc(node, this.ndb));
  }

  /** The node id the visible folder tree hangs from. */
  rootFolderNid() {
    const store = this.store();
    const entryId = store.ipmSubTreeEntryId;
    // An EntryID is 24 bytes; the node id is the last four.
    if (entryId instanceof Uint8Array && entryId.length >= 24) {
      const dv = new DataView(entryId.buffer, entryId.byteOffset, entryId.byteLength);
      const nid = dv.getUint32(20, true);
      if (this.ndb.nodeIndex().has(nid)) return nid;
    }
    return NID.ROOT_FOLDER;
  }

  #folderProps(nid) {
    const node = this.ndb.node(nid);
    if (!node) return null;
    return values(nodePc(node, this.ndb));
  }

  /**
   * The parent-child relation over folders, and which messages sit in which.
   *
   * A folder's hierarchy table is the documented place to look, and in a .pst
   * it is populated. In a .ost it usually is not: Outlook materialises those
   * view tables lazily, so a 449-folder offline store can have five of them.
   * Measured on a real 900 MB .ost, five hierarchy tables held seven rows
   * between them — while every one of the 449 folder nodes carried a correct
   * parent pointer in the node B-tree.
   *
   * So the tree is built from the node B-tree, which is always complete, and
   * the hierarchy table is used only where it exists, to order siblings the way
   * Outlook shows them.
   */
  #relations() {
    if (this._relations) return this._relations;
    const index = this.ndb.nodeIndex();
    const folders = new Set();
    const children = new Map();
    const messagesByFolder = new Map();

    for (const [nid, entry] of index) {
      const type = nidType(nid);
      if (type === NID_TYPE.NORMAL_FOLDER || type === NID_TYPE.SEARCH_FOLDER) folders.add(nid);
      else if (type === NID_TYPE.NORMAL_MESSAGE) {
        const list = messagesByFolder.get(entry.parentNid) || [];
        list.push(nid);
        messagesByFolder.set(entry.parentNid, list);
      }
    }

    const roots = [];
    for (const nid of folders) {
      const parent = index.get(nid).parentNid;
      if (parent === nid || !folders.has(parent)) {
        roots.push(nid);
        continue;
      }
      const list = children.get(parent) || [];
      list.push(nid);
      children.set(parent, list);
    }

    // A store whose folders all point at one another needs a starting point:
    // the folder nobody claims as a child.
    if (!roots.length) {
      const claimed = new Set([...children.values()].flat());
      for (const nid of folders) if (!claimed.has(nid)) roots.push(nid);
    }

    this._relations = { folders, children, messagesByFolder, roots };
    return this._relations;
  }

  /** Subfolder node ids, hierarchy-table order where that table exists. */
  #childFolderNids(nid) {
    const { children } = this.#relations();
    const kids = children.get(nid) || [];
    if (kids.length < 2) return kids;

    const node = this.ndb.node(siblingNid(nid, NID_TYPE.HIERARCHY_TABLE));
    if (!node) return kids;
    const { rows } = readTc(node, this.ndb);
    if (!rows.length) return kids;
    const order = new Map();
    rows.forEach((row, i) => order.set(row.get('rowId') >>> 0, i));
    return kids.slice().sort((a, b) => (order.get(a) ?? 1e9) - (order.get(b) ?? 1e9));
  }

  /**
   * The folder tree, as a person sees it in Outlook.
   * @param {{ includeEmpty?: boolean, includeHidden?: boolean }} [opts]
   */
  folders(opts = {}) {
    if (this._folders && !opts.force) return this._folders;
    const { messagesByFolder, roots } = this.#relations();
    const seen = new Set();

    const build = (nid, depth, path) => {
      if (seen.has(nid) || depth > 24) return null;
      seen.add(nid);
      const props = this.#folderProps(nid) || {};
      const name = decodeText(props.displayName) || `Folder ${nid}`;
      const here = depth === 0 ? [] : [...path, name];
      const children = this.#childFolderNids(nid)
        .map((child) => build(child, depth + 1, here))
        .filter(Boolean);
      // The stored count and the messages actually present can disagree in a
      // partially-synchronised .ost. The list shows what is really there.
      const present = (messagesByFolder.get(nid) || []).length;
      return {
        nid,
        name,
        path: here.join('/'),
        depth,
        messageCount: present || Number(props.contentCount || 0),
        storedCount: Number(props.contentCount || 0),
        unreadCount: Number(props.contentUnreadCount || 0),
        hasChildren: children.length > 0,
        hidden: HIDDEN_FOLDERS.has(name),
        children,
      };
    };

    const trees = roots.map((nid) => build(nid, 0, [])).filter(Boolean);
    // Any folder the walk did not reach — an orphan in a damaged file — is
    // still somebody's mail, so it is attached at the top rather than dropped.
    for (const nid of this.#relations().folders) {
      if (!seen.has(nid)) {
        const orphan = build(nid, 0, []);
        if (orphan) trees.push(orphan);
      }
    }

    const tree = trees.length === 1 ? trees[0] : { nid: 0, name: '', path: '', depth: -1, children: trees, hidden: true, messageCount: 0 };
    const list = [];
    const flatten = (node) => {
      if (!node) return;
      for (const child of node.children) {
        if (!opts.includeHidden && child.hidden && !child.children.length && !child.messageCount) continue;
        list.push(child);
        flatten(child);
      }
    };
    flatten(tree);
    this._folders = { root: tree, list };
    return this._folders;
  }

  /**
   * The message list of one folder, read from its contents table — headers
   * only, which is what a list needs and all it should cost.
   */
  messages(folderNid, { offset = 0, limit = 200 } = {}) {
    // The contents table is the cheap path: one node read gives the whole list
    // with the columns a list view wants. It is authoritative when it has rows.
    const table = this.ndb.node(siblingNid(folderNid, NID_TYPE.CONTENTS_TABLE));
    const tc = table ? readTc(table, this.ndb) : { rows: [], columns: [] };

    // A contents table is only useful if it carries the columns a list shows.
    // Outlook writes view tables per view, and plenty of them hold nothing but
    // ids and timestamps — measured on a real store, six of those tables gave
    // 200 rows with a date and 17 subjects between them.
    const hasSubject = tc.columns.some((c) => c.id === 0x0037 || c.id === 0x0e1d);
    if (tc.rows.length && hasSubject) {
      const out = [];
      for (let i = offset; i < Math.min(tc.rows.length, offset + limit); i++) out.push(this.#fromRow(tc.rows[i]));
      return out;
    }

    // Otherwise read each message's own properties. One node read per message,
    // for the page being shown and no further.
    let nids = this.#relations().messagesByFolder.get(folderNid) || [];
    if (!nids.length && tc.rows.length) {
      nids = tc.rows.map((row) => row.get('rowId') >>> 0).filter((nid) => this.ndb.nodeIndex().has(nid));
    }
    return nids
      .slice(offset, offset + limit)
      .map((nid) => this.#headerOf(nid))
      .filter(Boolean);
  }

  /** How many messages a folder holds, without reading any of them. */
  messageCount(folderNid) {
    const table = this.ndb.node(siblingNid(folderNid, NID_TYPE.CONTENTS_TABLE));
    const rows = table ? readTc(table, this.ndb).rows.length : 0;
    return rows || (this.#relations().messagesByFolder.get(folderNid) || []).length;
  }

  #fromRow(row) {
    const nid = row.get('rowId') >>> 0;
    const pick = (id) => row.get(id) ?? null;
    const flags = pick(0x0e07) || 0;
    return {
      nid,
      subject: stripSubjectPrefix(decodeText(pick(0x0037)) || decodeText(pick(0x0e1d)) || ''),
      from: {
        name: decodeText(pick(0x0c1a)) || decodeText(pick(0x0042)) || null,
        address: decodeText(pick(0x5d01)) || decodeText(pick(0x0c1f)) || '',
      },
      to: decodeText(pick(0x0e04)) || '',
      date: pick(0x0e06) || pick(0x0039) || null,
      size: pick(0x0e08) || 0,
      unread: !(flags & MSG_FLAG.READ),
      hasAttachments: Boolean(flags & MSG_FLAG.HAS_ATTACH),
      preview: (decodeText(pick(0x1000)) || '').replace(/\s+/g, ' ').slice(0, 200),
    };
  }

  #headerOf(nid) {
    const node = this.ndb.node(nid >>> 0);
    if (!node) return null;
    const p = values(nodePc(node, this.ndb));
    const flags = Number(p.messageFlags || 0);
    const body = decodeText(p.body);
    return {
      nid: nid >>> 0,
      subject: stripSubjectPrefix(decodeText(p.subject) || decodeText(p.normalizedSubject) || ''),
      from: addressFrom(p) || { name: null, address: '' },
      to: decodeText(p.displayTo) || '',
      date: p.messageDeliveryTime || p.clientSubmitTime || p.creationTime || null,
      size: Number(p.messageSize || 0),
      unread: !(flags & MSG_FLAG.READ),
      hasAttachments: Boolean(flags & MSG_FLAG.HAS_ATTACH),
      preview: (body || '').replace(/\s+/g, ' ').slice(0, 200),
    };
  }

  /** One message, in full. */
  message(nid) {
    const node = this.ndb.node(nid >>> 0);
    if (!node) return null;
    const pc = nodePc(node, this.ndb);
    const p = values(pc);
    const flags = Number(p.messageFlags || 0);

    const recipients = this.#recipients(node);
    const attachments = this.#attachments(node);
    const html = decodeText(p.bodyHtml);
    const text = decodeText(p.body);

    return {
      nid,
      subject: stripSubjectPrefix(decodeText(p.subject) || decodeText(p.normalizedSubject) || ''),
      from: [addressFrom(p)].filter(Boolean),
      to: recipients.filter((r) => r.type === 'to'),
      cc: recipients.filter((r) => r.type === 'cc'),
      bcc: recipients.filter((r) => r.type === 'bcc'),
      date: p.messageDeliveryTime || p.clientSubmitTime || p.creationTime || null,
      messageId: decodeText(p.internetMessageId) || null,
      inReplyTo: decodeText(p.inReplyToId) || null,
      text,
      html,
      preview: (text || stripHtml(html || '')).replace(/\s+/g, ' ').slice(0, 220),
      unread: !(flags & MSG_FLAG.READ),
      hasAttachments: Boolean(flags & MSG_FLAG.HAS_ATTACH) || attachments.length > 0,
      attachments,
      internetHeaders: decodeText(p.transportMessageHeaders) || null,
      messageClass: decodeText(p.messageClass) || 'IPM.Note',
      size: Number(p.messageSize || 0),
      source: 'pst',
    };
  }

  #recipients(node) {
    const sub = node.subnodes?.get(NID_RECIPIENT_TABLE);
    if (!sub) return [];
    const table = this.ndb.subnode(sub);
    if (!table) return [];
    const { rows } = readTc(table, this.ndb);
    return rows
      .map((row) => {
        const name = decodeText(row.get(0x3001)) || decodeText(row.get(0x5ff6)) || null;
        const address = decodeText(row.get(0x39fe)) || decodeText(row.get(0x3003)) || '';
        return {
          name,
          address,
          display: name && address ? `${name} <${address}>` : name || address,
          type: RECIPIENT_TYPE[row.get(0x0c15)] || 'to',
        };
      })
      .filter((r) => r.name || r.address);
  }

  #attachments(node) {
    const out = [];
    if (!node.subnodes) return out;
    for (const [subNid, sub] of node.subnodes) {
      if (nidType(subNid) !== NID_TYPE.ATTACHMENT) continue;
      const attachNode = this.ndb.subnode(sub);
      if (!attachNode) continue;
      const heap = new Heap(attachNode.parts);
      const props = values(readPc(heap, attachNode.subnodes, this.ndb));
      const data = props.attachData;
      out.push({
        nid: subNid,
        filename: decodeText(props.attachLongFilename) || decodeText(props.attachFilename) || `attachment-${out.length + 1}`,
        type: decodeText(props.attachMimeTag) || 'application/octet-stream',
        contentId: decodeText(props.attachContentId) || null,
        size: Number(props.attachSize || (data instanceof Uint8Array ? data.length : 0)),
        inline: Boolean(props.attachContentId),
        bytes: data instanceof Uint8Array ? data : null,
      });
    }
    return out;
  }

  /** Every folder with its message count — what an import preview shows. */
  summary() {
    const { list } = this.folders();
    const folders = list
      .filter((f) => !f.hidden)
      .map((f) => ({ nid: f.nid, name: f.name, path: f.path, depth: f.depth, messages: f.messageCount, unread: f.unreadCount }));
    return {
      ...this.info,
      store: decodeText(this.store().displayName) || null,
      folders,
      messages: folders.reduce((n, f) => n + f.messages, 0),
    };
  }

  close() {
    this.reader.close?.();
  }
}

/** Outlook stores "RE: " as a prefix property plus the bare subject. */
function stripSubjectPrefix(subject) {
  if (!subject) return '';
  // A control character at the front encodes the prefix length; drop it.
  if (subject.charCodeAt(0) === 1) return subject.slice(2);
  return subject;
}

export { NID, NID_TYPE, nidType, CRYPT_NAME, PROPS };
