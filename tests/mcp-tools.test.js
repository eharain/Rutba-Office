// The AI tools (owner, 2026-10-16, D-2026-10-16-06: "it should be off unless
// some one opens it for single session or want to keep it always open"):
// the protocol's answers, the switch (off, for a session that is running,
// always), the folders a call may reach, never overwriting, the log, and each
// tool's file read back by the engines.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { PassThrough } from 'node:stream';
import { createServer, serve, PROTOCOL_VERSIONS } from '../apps/desktop/main/mcp/protocol.js';
import { createTools, documentText, presentationText, workbookText, pdfText } from '../apps/desktop/main/mcp/tools.js';
import { openNow, resolveIn, MODE_KEY, SESSION_KEY, FOLDERS_KEY, readLog } from '../apps/desktop/main/mcp/access.js';
import { readWorkbook } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation/build';

const scratch = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mcp-'));
function setup(mode = 'always', extra = {}) {
  const dir = scratch();
  const work = path.join(dir, 'work');
  fs.mkdirSync(work);
  let s = { [MODE_KEY]: mode, [FOLDERS_KEY]: [work], ...extra };
  const tools = createTools({ settings: () => s, dataDir: dir, isAlive: (pid) => pid === 4242 });
  return { dir, work, tools, set: (next) => { s = { ...s, ...next }; } };
}
const textOf = (r) => r.content.map((c) => c.text).join('\n');

test('the protocol: initialize names the server and its tools, tools/list lists them, a notification is not answered', async () => {
  const { tools } = setup();
  const server = createServer({ name: 'rutba-office', version: '9.9.9', tools });
  const init = await server.handle({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-03-26', capabilities: {}, clientInfo: { name: 't', version: '1' } } });
  assert.equal(init.result.protocolVersion, '2025-03-26');
  assert.deepEqual(init.result.serverInfo, { name: 'rutba-office', version: '9.9.9' });
  assert.ok(init.result.capabilities.tools);
  const unknown = await server.handle({ jsonrpc: '2.0', id: 2, method: 'initialize', params: { protocolVersion: '1999-01-01' } });
  assert.equal(unknown.result.protocolVersion, PROTOCOL_VERSIONS[0], 'an unknown version is answered with ours');
  assert.equal(await server.handle({ jsonrpc: '2.0', method: 'notifications/initialized' }), null);
  const list = await server.handle({ jsonrpc: '2.0', id: 3, method: 'tools/list' });
  const names = list.result.tools.map((t) => t.name);
  for (const n of ['create_document', 'create_workbook', 'create_presentation', 'read_file', 'export_pdf', 'combine_pdf_pages', 'read_text_in_picture', 'create_from_template']) assert.ok(names.includes(n), n);
  assert.ok(list.result.tools.every((t) => t.inputSchema?.type === 'object' && t.description));
  assert.equal((await server.handle({ jsonrpc: '2.0', id: 4, method: 'resources/list' })).error.code, -32601);
});

test('over the streams: a line in, a line out, in order', async () => {
  const { tools } = setup();
  const server = createServer({ tools });
  const input = new PassThrough();
  const output = new PassThrough();
  const lines = [];
  output.on('data', (d) => lines.push(...String(d).split('\n').filter(Boolean)));
  const done = serve(server, input, output);
  input.write('{"jsonrpc":"2.0","id":1,"method":"ping"}\n{"jsonrpc":"2.0","id":2,"method":"tools/list"}\nnot json\n');
  input.end();
  await done;
  const msgs = lines.map((l) => JSON.parse(l));
  assert.equal(msgs[0].id, 1);
  assert.equal(msgs[1].id, 2);
  assert.equal(msgs[2].error.code, -32700);
});

test('off unless turned on: off refuses with how to turn it on, a session counts only while its Rutba Office runs, always answers', async () => {
  const { tools, set, dir } = setup('off');
  const off = await tools.call('list_folders', {});
  assert.equal(off.isError, true);
  assert.match(textOf(off), /AI tools are off.*AI Tools from the app menu/);
  set({ [MODE_KEY]: 'session', [SESSION_KEY]: 999 });
  assert.equal((await tools.call('list_folders', {})).isError, true, 'the session that turned them on has ended');
  set({ [SESSION_KEY]: 4242 });
  assert.equal((await tools.call('list_folders', {})).isError, undefined, 'its Rutba Office is running');
  set({ [MODE_KEY]: 'always' });
  assert.equal((await tools.call('list_folders', {})).isError, undefined);
  assert.equal(openNow({}).open, false, 'nothing set is off');
  const log = readLog(dir);
  assert.equal(log.length, 4);
  assert.equal(log[3].ok, false, 'the refusal is in the log');
  assert.ok(log.every((e) => e.tool === 'list_folders' && typeof e.ms === 'number'));
});

test('only in the folders given: a path outside, or out through .., is refused; a name alone goes in the first folder', () => {
  const work = scratch();
  assert.equal(resolveIn([work], 'a.docx'), path.join(work, 'a.docx'));
  assert.throws(() => resolveIn([work], path.join(os.tmpdir(), 'elsewhere.docx')), /may only use these folders/);
  assert.throws(() => resolveIn([work], '..\\..\\escape.docx'), /may only use these folders/);
  assert.throws(() => resolveIn([], 'a.docx'), /No folder has been given/);
});

test('a document from blocks: title, headings, paragraphs, lists and a table, read back as written', async () => {
  const { tools, work } = setup();
  const r = await tools.call('create_document', {
    path: 'report.docx',
    title: 'Quarterly review',
    blocks: [
      { type: 'heading', text: 'Where we stand' },
      { type: 'paragraph', text: 'Sales rose in every region.' },
      { type: 'bullets', items: ['North up 12%', 'South up 8%'] },
      { type: 'numbered', items: ['Hire two', 'Open Leeds'] },
      { type: 'table', rows: [['Region', 'Sales'], ['North', 120], ['South', 98]] },
    ],
  });
  assert.equal(r.isError, undefined, textOf(r));
  const read = documentText(fs.readFileSync(path.join(work, 'report.docx')));
  assert.match(read, /^# Quarterly review$/m);
  assert.match(read, /^## Where we stand$/m);
  assert.match(read, /^Sales rose in every region\.$/m);
  assert.match(read, /^- North up 12%$/m);
  assert.match(read, /^- Open Leeds$/m);
  assert.match(read, /^Region \| Sales$/m);
  assert.match(read, /^North \| 120$/m);
  const again = await tools.call('read_file', { path: 'report.docx' });
  assert.match(textOf(again), /Where we stand/);
});

test('never overwrites: a name already there is refused, and the file is as it was', async () => {
  const { tools, work } = setup();
  fs.writeFileSync(path.join(work, 'keep.docx'), 'mine');
  const r = await tools.call('create_document', { path: 'keep.docx', blocks: [{ type: 'paragraph', text: 'x' }] });
  assert.equal(r.isError, true);
  assert.match(textOf(r), /already exists, and the AI tools never overwrite/);
  assert.equal(fs.readFileSync(path.join(work, 'keep.docx'), 'utf8'), 'mine');
  assert.equal((await tools.call('create_document', { path: 'wrong.xlsx', blocks: [] })).isError, true, 'the kind follows the name');
});

test('a workbook with a formula, and a presentation, read back', async () => {
  const { tools, work } = setup();
  assert.equal((await tools.call('create_workbook', { path: 'sales.xlsx', sheets: [{ name: 'Sales', rows: [['Region', 'Amount'], ['North', 120], ['South', 98], ['Total', '=SUM(B2:B3)']] }] })).isError, undefined);
  const book = readWorkbook(fs.readFileSync(path.join(work, 'sales.xlsx')));
  assert.equal(book.sheets[0].name, 'Sales');
  assert.equal(book.sheets[0].rows[3][1], '=SUM(B2:B3)');
  assert.match(workbookText(fs.readFileSync(path.join(work, 'sales.xlsx'))), /North\t120/);
  assert.equal((await tools.call('create_presentation', { path: 'plan.pptx', title: 'Plan', subtitle: 'Autumn', slides: [{ title: 'Goals', points: ['Grow', 'Hire'] }] })).isError, undefined);
  const deck = presentationText(fs.readFileSync(path.join(work, 'plan.pptx')));
  assert.match(deck, /Slide 1:\nPlan\nAutumn/);
  assert.match(deck, /Slide 2:\nGoals\nGrow\nHire/);
});

test('a template, and a PDF made from pages of two', async () => {
  const { tools, work } = setup();
  assert.match(textOf(await tools.call('list_templates', {})), /^word-letter\tdocx\tLetter/m);
  assert.equal((await tools.call('create_from_template', { template: 'sheet-budget', path: 'budget.xlsx' })).isError, undefined);
  assert.ok(fs.statSync(path.join(work, 'budget.xlsx')).size > 1000);
  assert.equal((await tools.call('create_from_template', { template: 'nothing', path: 'n.docx' })).isError, true);

  // Two PDFs to work with, made by the suite's own writer.
  const { renderPdf } = await import('@rutba/doc-view/export/pdf');
  const { openDocx } = await import('@rutba/doc-view/backends/ooxml');
  const tools2 = createTools({ settings: () => ({ [MODE_KEY]: 'always', [FOLDERS_KEY]: [work] }), dataDir: work, renderPdf, openDocx });
  await tools2.call('create_document', { path: 'one.docx', title: 'One', blocks: [{ type: 'paragraph', text: 'First file' }] });
  assert.equal((await tools2.call('export_pdf', { path: 'one.docx', to: 'one.pdf' })).isError, undefined);
  assert.match(pdfText(fs.readFileSync(path.join(work, 'one.pdf'))), /First file/);
  const deck = buildPptx({ title: 'Deck', slides: [{ layout: 'title', title: 'Deck', body: '' }] });
  fs.writeFileSync(path.join(work, 'deck.pptx'), deck);
  const r = await tools2.call('combine_pdf_pages', { files: [{ path: 'one.pdf' }, { path: 'one.pdf', rotate: 90 }], to: 'both.pdf' });
  assert.equal(r.isError, undefined, textOf(r));
  assert.match(textOf(r), /2 pages/);
  assert.match(pdfText(fs.readFileSync(path.join(work, 'both.pdf'))), /Page 2:\n[\s\S]*First file/);
});

test('the network only by choice: one file of the AI tools listens, remote.js, and this computer only is the default (owner, D-2026-10-17-04 "protect it from remote connections", D-2026-10-21-06 "the user should have choice to open remote connections or local access only")', () => {
  const dir = path.join(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')), '..', 'apps', 'desktop', 'main', 'mcp');
  const files = fs.readdirSync(dir).filter((f) => f.endsWith('.js'));
  assert.ok(files.includes('node.js') && files.includes('tools.js') && files.includes('protocol.js') && files.includes('remote.js'), files.join(','));
  const NET = /(?:from\s+|require\(\s*|import\(\s*)['"](?:node:)?(?:net|http|https|http2|dgram|tls|dns)['"]/;
  for (const f of files.filter((name) => name !== 'remote.js' && name !== 'channel.js')) {
    const src = fs.readFileSync(path.join(dir, f), 'utf8');
    assert.doesNotMatch(src, NET, `${f} brings in a networking module`);
    assert.doesNotMatch(src, /\.listen\s*\(/, `${f} listens on a port`);
  }
  // The channel to the running copy (channel.js) is a named pipe, or a socket file: never a port.
  const channel = fs.readFileSync(path.join(dir, 'channel.js'), 'utf8');
  assert.doesNotMatch(channel, /(?:from\s+)['"](?:node:)?(?:http|https|http2|dgram|tls|dns)['"]/);
  assert.deepEqual([...channel.matchAll(/\.listen\s*\(([^,)]*)/g)].map((m) => m[1].trim()), ['where']);
  assert.match(channel, /\\\\\\\\\.\\\\pipe\\\\rutba-office-ai-/);
  const ai = fs.readFileSync(path.join(dir, '..', 'ai-tools.js'), 'utf8');
  assert.doesNotMatch(ai, NET, 'ai-tools.js brings in a networking module');
  // Off the network unless the person chooses it.
  assert.match(ai, /settings\.get\(REACH_KEY, 'local'\)/);
});
