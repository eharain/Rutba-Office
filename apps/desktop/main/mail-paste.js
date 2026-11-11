// Copy, cut and paste of messages between folders (Mail 2.25).
//
// Copy or Cut on messages, then Paste on a folder: within one account the
// server's own COPY or MOVE (a change on the line, as Move to makes), and
// between accounts the message as it is APPENDed to the target folder, its
// date and its marks with it, a cut then taking the source away once the
// copy is kept. Local folders (a mailbox brought in from a file) change here
// alone, the same way. An archive opened in place is read only: nothing is
// taken out of it or put in it.

/** The IMAP marks a message carries to its new folder: read, flagged, answered. */
export function imapFlagsOf(message) {
  const flags = [];
  if (!message.unread) flags.push('\\Seen');
  if (message.flagged) flags.push('\\Flagged');
  if (message.answered) flags.push('\\Answered');
  return flags;
}

/**
 * A message as another account's folder keeps it: its attachments' bytes read
 * from where they are now, since the store finds a moved message's files
 * only within one account, and its place (folder, id, UID) left behind.
 */
export function carriedRecord(store, accountId, folder, id, message) {
  const { folder: _folder, id: _id, uid: _uid, smime: _smime, ...rest } = message;
  const attachments = (message.attachments || []).map((a) => {
    let bytes = null;
    try {
      bytes = a.stored ? store.attachmentBytes(accountId, folder, id, a.stored) : null;
    } catch {
      bytes = null;
    }
    const { stored: _stored, ...kept } = a;
    return bytes ? { ...kept, bytes } : { ...kept, stored: false };
  });
  const original = message.smime?.original ? store.original(accountId, folder, id) : null;
  return { ...rest, attachments, uid: null, ...(original ? { smimeRaw: original } : {}) };
}

/**
 * Paste: `ids` of `folder` in `accountId`, into `to` ({ accountId, folder }),
 * copied, or moved with `cut`. Answers { copied, moved }: a message no
 * longer there is passed over, and pasting where they already are is nothing.
 *
 * `ctx` is the mail service's own: its store, its accounts and the changes it
 * sends its servers (mail.js).
 */
export async function pasteMessages(ctx, { accountId, folder, ids = [], to = null, cut = false }) {
  const { store, find, isArchive, live, move, copyWithin, expunge, record, kick, sourceBytes, listFolder } = ctx;
  const target = { accountId: to?.accountId || accountId, folder: String(to?.folder || '').trim() };
  if (!target.folder) throw new Error('Paste needs a folder to go to.');
  if (isArchive(accountId) || isArchive(target.accountId)) throw new Error('An archive opened in place is read only: messages are not copied out of it or into it.');
  find(accountId);
  find(target.accountId);
  const here = [...new Set(ids || [])].filter((id) => store.get(accountId, folder, id));
  if (!here.length) return { copied: 0, moved: 0 };

  if (target.accountId === accountId) {
    if (target.folder === folder) return { copied: 0, moved: 0 };
    // The server's MOVE, with all Move to does (the junk filter learns from it).
    if (cut) return { copied: 0, moved: move({ accountId, folder, ids: here, to: target.folder }).moved };
    return { copied: copyWithin(accountId, folder, here, target.folder), moved: 0 };
  }

  // Between accounts: each kept in the target, and APPENDed to its server, before any source goes.
  const serverOf = live(target.accountId);
  let copied = 0;
  const done = [];
  for (const id of here) {
    const message = store.get(accountId, folder, id);
    if (!message) continue;
    const kept = carriedRecord(store, accountId, folder, id, message);
    const newId = store.put(target.accountId, target.folder, kept, { force: true, flush: true });
    if (serverOf && newId) {
      const bytes = await sourceBytes(accountId, folder, id, message);
      const file = store.writeOutgoing(target.accountId, `paste-${newId}-${Date.now()}.eml`, Buffer.from(bytes));
      record(target.accountId, { kind: 'append', folder: target.folder, file, id: newId, messageId: message.messageId || null, flags: imapFlagsOf(message), date: message.date || null });
    }
    copied++;
    done.push(id);
  }
  listFolder(target.accountId, target.folder);
  if (serverOf) kick(target.accountId);
  if (cut && done.length) expunge(accountId, folder, done);
  return { copied, moved: cut ? done.length : 0 };
}
