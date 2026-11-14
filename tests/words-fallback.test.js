// The main process's words and the answer it gives each window when the
// language packs' source fails: English, as the labels' definitions give it,
// never every label bare (bare: true is for a source with no words at all).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import esbuild from 'esbuild';
import { WORDS_CHANNEL } from '../packages/office-shell/src/contract.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/** office-shell's words.js bundled for node, with a stand-in for electron that keeps the windows' handlers on globalThis.ipcHandlers. */
async function words() {
  const out = await esbuild.build({
    entryPoints: [path.join(ROOT, 'packages/office-shell/src/electron/words.js')],
    bundle: true,
    format: 'esm',
    platform: 'node',
    write: false,
    logLevel: 'silent',
    plugins: [{
      name: 'electron-stand-in',
      setup(b) {
        b.onResolve({ filter: /^electron$/ }, () => ({ path: 'electron', namespace: 'stand-in' }));
        b.onLoad({ filter: /.*/, namespace: 'stand-in' }, () => ({
          contents: `const handlers = (globalThis.ipcHandlers = {});
export const ipcMain = { on(channel, fn) { handlers[channel] = fn; }, removeAllListeners(channel) { delete handlers[channel]; } };
export const app = { isReady: () => false, getLocale: () => 'en' };`,
          loader: 'js',
        }));
      },
    }],
  });
  // Beside words.js: the modules it bundles read their JSON by paths relative to themselves (createRequire).
  const file = path.join(ROOT, 'packages/office-shell/src/electron', `.words-fallback-${process.pid}.mjs`);
  fs.writeFileSync(file, out.outputFiles[0].text);
  try {
    return await import(pathToFileURL(file).href);
  } finally {
    fs.rmSync(file, { force: true });
  }
}

/** What a window asking for `tag` is answered. */
const ask = (tag) => {
  const event = { sender: { id: 1, once() {} }, returnValue: undefined };
  globalThis.ipcHandlers[WORDS_CHANNEL](event, tag);
  return event.returnValue;
};

test('when the packs fail, a window is answered in English and the main process speaks English, not every label bare', async () => {
  const w = await words();
  const source = (wordsFor) => ({ wordsFor, languages: () => [], fontsFor: () => null });
  // A source with no English pack: every label bare, in a window and in the main process.
  w.useLanguageWords(source(() => ({ tag: 'en', table: {}, english: null, bare: true })));
  assert.equal(ask('en').bare, true);
  w.speakAs({ settings: { get: () => 'en' } });
  assert.equal(w.t('common.save'), 'save');
  // A source with no words at all: still every label bare in a window.
  w.useLanguageWords(source(() => null));
  assert.equal(ask('ur').bare, true);
  // The source failing: English, as the definitions give it.
  w.useLanguageWords(source(() => { throw new Error('the packs could not be read'); }));
  const answer = ask('ur');
  assert.equal(answer.bare, null);
  assert.equal(answer.table, null);
  assert.equal(w.wordsFor('ur'), null, 'the exported lookup still says there are none');
  w.speakAs({ settings: { get: () => 'ur' } });
  assert.equal(w.t('common.save'), 'Save', 'the main process in English, not "save"');
});
