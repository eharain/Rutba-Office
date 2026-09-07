// Diagnostic: how the folder hierarchy is actually expressed in a given file.
// Prints counts and shapes only — never message content.

import { Ndb, NID, NID_TYPE, nidType, siblingNid } from '../packages/mailbox/src/pst/ndb.js';
import { nodePc, readTc, values } from '../packages/mailbox/src/pst/ltp.js';
import { fileReader } from './pst-probe.js';

const target = process.argv[2];
const reader = fileReader(target);
const ndb = new Ndb(reader);
const index = ndb.nodeIndex();

const folders = [...index.keys()].filter((n) => nidType(n) === NID_TYPE.NORMAL_FOLDER);
console.log(`normal folders in NBT: ${folders.length}`);

// 1. Does the NBT's parent pointer form a tree over the folders?
const folderSet = new Set(folders);
let parentInSet = 0;
let parentZero = 0;
const roots = [];
for (const nid of folders) {
  const p = index.get(nid).parentNid;
  if (p === 0) parentZero++;
  else if (folderSet.has(p)) parentInSet++;
  else roots.push({ nid, parent: p });
}
console.log(`  parent is another folder: ${parentInSet}`);
console.log(`  parent is 0:              ${parentZero}`);
console.log(`  parent outside the set:   ${roots.length} (first few: ${roots.slice(0, 5).map((r) => `${r.nid}<-${r.parent}`).join(' ')})`);

// 2. Do hierarchy tables produce children?
let withTable = 0;
let totalChildren = 0;
for (const nid of folders.slice(0, 200)) {
  const t = ndb.node(siblingNid(nid, NID_TYPE.HIERARCHY_TABLE));
  if (!t) continue;
  const { rows } = readTc(t, ndb);
  if (rows.length) {
    withTable++;
    totalChildren += rows.length;
  }
}
console.log(`hierarchy tables with rows (first 200 folders): ${withTable}, ${totalChildren} child rows total`);

// 3. Do contents tables produce rows?
let contentsWithRows = 0;
let contentRows = 0;
for (const nid of folders.slice(0, 200)) {
  const t = ndb.node(siblingNid(nid, NID_TYPE.CONTENTS_TABLE));
  if (!t) continue;
  const { rows, columns } = readTc(t, ndb);
  if (rows.length) {
    contentsWithRows++;
    contentRows += rows.length;
    if (contentsWithRows === 1) console.log(`  first contents table: ${columns.length} columns, ${rows.length} rows`);
  }
}
console.log(`contents tables with rows (first 200 folders): ${contentsWithRows}, ${contentRows} rows total`);

// 4. Does the root folder we chose have children?
const store = ndb.node(NID.MESSAGE_STORE);
const props = store ? values(nodePc(store, ndb)) : {};
console.log(`store props: ${Object.keys(props).length} (${Object.keys(props).slice(0, 8).join(', ')})`);
const entry = props.ipmSubTreeEntryId;
console.log(`ipmSubTreeEntryId: ${entry instanceof Uint8Array ? `${entry.length} bytes` : entry === undefined ? 'absent' : typeof entry}`);
if (entry instanceof Uint8Array && entry.length >= 24) {
  const dv = new DataView(entry.buffer, entry.byteOffset, entry.byteLength);
  console.log(`  nid at offset 20: ${dv.getUint32(20, true)} (in NBT: ${index.has(dv.getUint32(20, true))})`);
}
console.log(`NID.ROOT_FOLDER 0x122 in NBT: ${index.has(NID.ROOT_FOLDER)}`);

// 5. Messages grouped by their NBT parent.
const messages = [...index.keys()].filter((n) => nidType(n) === NID_TYPE.NORMAL_MESSAGE);
const byParent = new Map();
for (const nid of messages) {
  const p = index.get(nid).parentNid;
  byParent.set(p, (byParent.get(p) || 0) + 1);
}
const parentsThatAreFolders = [...byParent.keys()].filter((p) => folderSet.has(p)).length;
console.log(`messages: ${messages.length}, distinct parents: ${byParent.size}, of which folders: ${parentsThatAreFolders}`);

reader.close();
