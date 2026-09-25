// The dictionaries on a thread of their own. Building one takes the best part
// of a second and a suggestion for an unusual name can take as long again;
// on the main process either would freeze every window of the suite.

import { parentPort } from 'node:worker_threads';
import { loadDictionary, checkWords, suggestWords } from './speller.js';

parentPort.on('message', async ({ id, op, lang, words, word, limit }) => {
  try {
    const d = await loadDictionary(lang);
    if (op === 'load') return parentPort.postMessage({ id, result: { lang: d.lang, builtMs: d.builtMs } });
    if (op === 'check') return parentPort.postMessage({ id, result: checkWords(d.spell, words || []) });
    if (op === 'suggest') return parentPort.postMessage({ id, result: suggestWords(d.spell, word, limit) });
    return parentPort.postMessage({ id, error: `no operation ${op}` });
  } catch (err) {
    return parentPort.postMessage({ id, error: String(err?.message || err) });
  }
});
