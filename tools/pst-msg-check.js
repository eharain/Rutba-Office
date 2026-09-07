// Diagnostic: why a message node's properties are or are not readable.
// Prints property names, types and value *lengths* — never values.

import { Ndb, NID_TYPE, nidType, NID } from '../packages/mailbox/src/pst/ndb.js';
import { Heap, readPc, readBth, nodePc, values } from '../packages/mailbox/src/pst/ltp.js';
import { fileReader } from './pst-probe.js';

const reader = fileReader(process.argv[2]);
const ndb = new Ndb(reader);
const index = ndb.nodeIndex();

function describe(label, nid) {
  const node = ndb.node(nid);
  if (!node) return console.log(`${label}: node ${nid} missing`);
  const heap = new Heap(node.parts);
  console.log(
    `${label} nid=${nid} parts=${node.parts.length} bytes=${node.data.length} ` +
      `subnodes=${node.subnodes.size} heapValid=${heap.valid} clientSig=0x${(heap.clientSig ?? 0).toString(16)} userRoot=0x${(heap.userRoot ?? 0).toString(16)}`
  );
  if (node.parts[0]) {
    console.log(`   first bytes: ${[...node.parts[0].subarray(0, 12)].map((b) => b.toString(16).padStart(2, '0')).join(' ')}`);
  }
  const bth = heap.valid ? readBth(heap, heap.userRoot) : { entries: [] };
  console.log(`   bth entries: ${bth.entries.length} (key ${bth.keySize}, entry ${bth.entrySize})`);
  const pc = readPc(heap, node.subnodes, ndb);
  const named = values(pc);
  console.log(`   properties: ${pc.size}; named: ${Object.keys(named).length}`);
  const shown = [...pc.values()].slice(0, 10).map((p) => `${p.name}:0x${p.type.toString(16)}${typeof p.value === 'string' ? `(${p.value.length}ch)` : p.value instanceof Uint8Array ? `(${p.value.length}b)` : ''}`);
  console.log(`   sample: ${shown.join(', ')}`);
}

describe('store  ', NID.MESSAGE_STORE);

const folders = [...index.keys()].filter((n) => nidType(n) === NID_TYPE.NORMAL_FOLDER);
describe('folder ', folders[1] ?? folders[0]);

const messages = [...index.keys()].filter((n) => nidType(n) === NID_TYPE.NORMAL_MESSAGE);
console.log(`messages in file: ${messages.length}`);

// Group by parent and sample from the folders that hold the most mail, since
// the first messages in node order are often system items.
const byParent = new Map();
for (const nid of messages) {
  const p = index.get(nid).parentNid;
  byParent.set(p, [...(byParent.get(p) || []), nid]);
}
const ranked = [...byParent.entries()].sort((a, b) => b[1].length - a[1].length);
console.log(`folders holding mail: ${ranked.length}; largest ${ranked.slice(0, 5).map(([, v]) => v.length).join(', ')}`);

for (const [parent, nids] of ranked.slice(0, 3)) {
  const folderProps = values(nodePc(ndb.node(parent) || { parts: [], subnodes: new Map() }, ndb));
  console.log(`-- folder ${parent}: ${nids.length} messages, name ${folderProps.displayName ? `${String(folderProps.displayName).length} chars` : 'none'}`);
  describe('  message', nids[Math.floor(nids.length / 2)]);
}

reader.close();
