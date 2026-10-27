// What the windows and the AI tools offer comes from the packs loaded
// (docs/APP-PACKS.md 5.1, items 4 and 5): the launcher's Open dialog offers
// the kinds of file of the apps installed, and an AI app is offered the
// tools of the apps installed, none without AI Tools, and is refused a tool
// whose app is absent even when it names it.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { openFilters, APP_EXTENSIONS, OLD_EXTENSIONS } from '../packages/office-formats/src/registry.js';
import { createTools } from '../apps/desktop/main/mcp/tools.js';
import { MODE_KEY, FOLDERS_KEY } from '../apps/desktop/main/mcp/access.js';
import { LOADED_KEY, TOOL_PACK } from '../apps/desktop/main/pack-tools.js';

test('the launcher\'s Open dialog offers the kinds of file of the apps installed alone', () => {
  const every = openFilters(null);
  assert.deepEqual(openFilters(null, { apps: null }), every, 'no list: every app, as before');
  const two = openFilters(null, { apps: ['word', 'pdf'] });
  const nameFor = (exts) => every.find((f) => f.extensions.join() === exts.join()).name;
  assert.deepEqual(two.map((f) => f.name), ['All supported files', nameFor(APP_EXTENSIONS.word), nameFor(APP_EXTENSIONS.pdf), 'All files']);
  assert.deepEqual(new Set(two[0].extensions), new Set([...APP_EXTENSIONS.word, ...(OLD_EXTENSIONS.word || []), ...APP_EXTENSIONS.pdf, ...(OLD_EXTENSIONS.pdf || [])]));
  assert.ok(!two[0].extensions.includes('xlsx'));
  // An app's own dialog is its own whatever is installed.
  assert.deepEqual(openFilters('sheets', { apps: ['word'] }), openFilters('sheets'));
});

test('an AI app is offered the tools of the apps installed, and refused one whose app is absent', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-pack-tools-'));
  const work = path.join(dir, 'work');
  fs.mkdirSync(work);
  let s = { [MODE_KEY]: 'always', [FOLDERS_KEY]: [work] };
  const tools = createTools({ settings: () => s, dataDir: dir });
  const all = (await tools.list()).map((t) => t.name);
  assert.ok(all.includes('create_document') && all.includes('create_workbook') && all.includes('list_folders'), 'a copy from before packs: every tool');

  s = { ...s, [LOADED_KEY]: ['ai', 'word', 'mail'] };
  const some = (await tools.list()).map((t) => t.name);
  assert.ok(some.includes('create_document') && some.includes('list_folders'));
  for (const name of some) assert.ok(!TOOL_PACK[name] || TOOL_PACK[name] === 'word', name);
  assert.ok(!some.includes('create_workbook') && !some.includes('read_text_in_picture') && !some.includes('combine_pdf_pages'));
  const refused = await tools.call('create_workbook', { path: path.join(work, 'x.xlsx'), sheets: [] });
  assert.equal(refused.isError, true);
  assert.match(refused.content[0].text, /^The tool create_workbook is not offered: the app it works with is not installed in Rutba Office on this computer\.$/);
  assert.equal(fs.existsSync(path.join(work, 'x.xlsx')), false);

  s = { ...s, [LOADED_KEY]: ['word', 'sheets'] };
  assert.deepEqual(await tools.list(), [], 'without AI Tools, no tool at all');
  fs.rmSync(dir, { recursive: true, force: true });
});
