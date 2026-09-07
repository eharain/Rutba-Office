// Structural probe for an Outlook data file.
//
//   node tools/pst-probe.js <file.pst|file.ost>
//
// Prints what the reader can see of the file's *structure* — version, page
// size, tree sizes, node type histogram, and whether the block cipher is one
// this build can decode. It deliberately prints no message content, because the
// files this is pointed at during development are somebody's real mail.

import fs from 'node:fs';
import path from 'node:path';
import { Ndb, nidType, NID_TYPE, NID } from '../packages/mailbox/src/pst/ndb.js';
import { PstFile } from '../packages/mailbox/src/pst/index.js';

export function fileReader(filePath) {
  const fd = fs.openSync(filePath, 'r');
  const size = fs.fstatSync(fd).size;
  return {
    size,
    read(offset, length) {
      const n = Math.max(0, Math.min(length, size - offset));
      const buf = Buffer.allocUnsafe(n);
      if (n) fs.readSync(fd, buf, 0, n, offset);
      return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
    },
    close: () => fs.closeSync(fd),
  };
}

const TYPE_NAMES = Object.fromEntries(Object.entries(NID_TYPE).map(([k, v]) => [v, k]));

function main() {
  const target = process.argv[2];
  if (!target) {
    console.error('usage: node tools/pst-probe.js <file.pst|file.ost>');
    process.exit(2);
  }
  const reader = fileReader(target);
  const ndb = new Ndb(reader);
  const stats = ndb.stats();
  console.log(`file        ${path.basename(target)}  (${(reader.size / 1048576).toFixed(1)} MB)`);
  console.log(`kind        ${stats.isOst ? 'OST' : 'PST'}  version ${stats.version}  ${stats.unicode ? 'unicode' : 'ANSI'}  page ${stats.pageSize}`);
  console.log(`crypt       ${stats.crypt}  ${stats.readable ? '(decodable)' : '(NOT decodable by this build)'}`);
  console.log(`btrees      ${stats.nodes} nodes, ${stats.blocks} blocks`);

  const hist = new Map();
  for (const nid of ndb.nodeIndex().keys()) {
    const t = nidType(nid);
    hist.set(t, (hist.get(t) || 0) + 1);
  }
  const rows = [...hist.entries()].sort((a, b) => b[1] - a[1]);
  console.log('node types');
  for (const [t, n] of rows) {
    console.log(`  ${String(TYPE_NAMES[t] || `0x${t.toString(16)}`).padEnd(24)} ${n}`);
  }

  const store = ndb.node(NID.MESSAGE_STORE);
  console.log(`message store node   ${store ? `${store.data.length} bytes, ${store.subnodes.size} subnodes` : 'MISSING'}`);
  if (store?.data?.length) {
    // An HN block starts with ibHnpm, then bSig 0xEC and the client signature.
    console.log(`  first bytes        ${[...store.data.subarray(0, 8)].map((b) => b.toString(16).padStart(2, '0')).join(' ')}`);
    console.log(`  HN signature 0xEC  ${store.data[2] === 0xec ? 'yes — the block decoded correctly' : 'no — the block did not decode'}`);
    console.log(`  client signature   0x${(store.data[3] ?? 0).toString(16)} (0xbc = property context)`);
  }
  const root = ndb.node(NID.ROOT_FOLDER);
  console.log(`root folder node     ${root ? `${root.data.length} bytes` : 'MISSING'}`);

  // Folder tree and one message, checked for shape only. Nothing about the
  // content is printed: during development this is pointed at real mail.
  const pst = new PstFile(reader);
  const summary = pst.summary();
  console.log(`store name           ${summary.store ? `present, ${summary.store.length} chars` : 'none'}`);
  console.log(`folders              ${summary.folders.length}, holding ${summary.messages} messages`);
  const named = summary.folders.filter((x) => x.name && !/^Folder \d/.test(x.name)).length;
  console.log(`  with real names    ${named} of ${summary.folders.length}`);

  // Sample across the biggest folders rather than one, because a store's
  // largest folder is often a system folder whose items carry no subject.
  const biggest = summary.folders.slice().sort((a, b) => b.messages - a.messages).slice(0, 6);
  let sampled = 0;
  const tally = { subject: 0, date: 0, sender: 0, preview: 0 };
  for (const folder of biggest) {
    const list = pst.messages(folder.nid, { limit: 40 });
    sampled += list.length;
    for (const m of list) {
      if (m.subject?.trim()) tally.subject++;
      if (m.date) tally.date++;
      if (m.from.name || m.from.address) tally.sender++;
      if (m.preview?.trim()) tally.preview++;
    }
    console.log(`  ${String(folder.messages).padStart(6)} messages  depth ${folder.depth}  listed ${list.length}`);
  }
  console.log(
    `sampled ${sampled} headers: ${tally.subject} with subject, ${tally.date} with date, ` +
      `${tally.sender} with sender, ${tally.preview} with preview`
  );

  const withMail = biggest.find((f) => pst.messages(f.nid, { limit: 5 }).some((m) => m.subject?.trim()));
  if (withMail) {
    const one = pst.messages(withMail.nid, { limit: 5 }).find((m) => m.subject?.trim());
    const full = pst.message(one.nid);
    console.log(
      `opened one message: body ${full.text ? `${full.text.length} chars` : 'none'}, ` +
        `html ${full.html ? `${full.html.length} chars` : 'none'}, ` +
        `${full.to.length} recipients, ${full.attachments.length} attachments, ` +
        `headers ${full.internetHeaders ? 'yes' : 'no'}, from ${full.from.length ? 'present' : 'missing'}`
    );
  }
  reader.close();
}

if (process.argv[1] && process.argv[1].endsWith("pst-probe.js")) main();