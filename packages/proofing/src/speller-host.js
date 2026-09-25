// The dictionaries, asked from the main process: on worker threads where
// they can be started, in this thread where they cannot (a test, a platform
// that refuses the worker). Answers are kept per language, so the second
// pass over a document asks the dictionary only about words it has not
// heard before.
//
// Two threads, each with its own copy of the dictionary: one answers "is
// this word right?", which takes milliseconds; the other finds suggestions,
// which for an unusual name can take two seconds. Kept apart, the pass
// moves on at once while the suggestions for the last word are still being
// worked out.

import { loadDictionary, checkWords, suggestWords } from './speller.js';

export function createSpellerHost({ worker = true } = {}) {
  let failed = !worker;
  let seq = 0;
  const threads = new Map(); // 'check' | 'suggest' -> Worker
  const waiting = new Map();
  const known = new Map(); // lang -> Map(word -> right?)
  const timings = {};

  const start = async (lane) => {
    if (failed) return null;
    if (threads.has(lane)) return threads.get(lane);
    try {
      const { Worker } = await import('node:worker_threads');
      const t = new Worker(new URL('./speller-worker.js', import.meta.url));
      t.unref?.();
      t.on('message', ({ id, result, error }) => {
        const w = waiting.get(id);
        if (!w) return;
        waiting.delete(id);
        if (error) w.reject(new Error(error));
        else w.resolve(result);
      });
      t.on('error', () => {
        for (const [id, w] of waiting) if (w.lane === lane) { waiting.delete(id); w.reject(new Error('the dictionary thread stopped')); }
        threads.delete(lane);
        failed = true;
      });
      threads.set(lane, t);
      return t;
    } catch {
      failed = true;
      return null;
    }
  };

  const ask = async (message, lane = message.op === 'suggest' ? 'suggest' : 'check') => {
    const t = await start(lane);
    if (t) {
      try {
        return await new Promise((resolve, reject) => {
          const id = ++seq;
          waiting.set(id, { resolve, reject, lane });
          t.postMessage({ id, ...message });
        });
      } catch {
        // The thread went away under us: answer here instead.
      }
    }
    const d = await loadDictionary(message.lang);
    if (message.op === 'load') return { lang: d.lang, builtMs: d.builtMs };
    if (message.op === 'check') return checkWords(d.spell, message.words || []);
    return suggestWords(d.spell, message.word, message.limit);
  };

  const cacheFor = (lang) => {
    if (!known.has(lang)) known.set(lang, new Map());
    return known.get(lang);
  };

  return {
    /** Read the dictionary now — both threads' copies; answers how long the checking one took. */
    async load(lang) {
      const began = Date.now();
      const [result] = await Promise.all([ask({ op: 'load', lang }, 'check'), ask({ op: 'load', lang }, 'suggest').catch(() => null)]);
      timings[lang] = timings[lang] ?? Date.now() - began;
      return { ...result, readyMs: timings[lang] };
    },

    /** The words that are not in the dictionary, of those given. */
    async misspelt(lang, words) {
      const cache = cacheFor(lang);
      const fresh = [...new Set(words.filter((w) => !cache.has(w)))];
      if (fresh.length) {
        const began = Date.now();
        const wrong = new Set(await ask({ op: 'check', lang, words: fresh }));
        if (timings[lang] == null) timings[lang] = Date.now() - began;
        for (const w of fresh) cache.set(w, !wrong.has(w));
        while (cache.size > 200000) cache.delete(cache.keys().next().value);
      }
      return new Set(words.filter((w) => cache.get(w) === false));
    },

    async suggest(lang, word, limit = 6) {
      return ask({ op: 'suggest', lang, word, limit });
    },

    timings,
    close() {
      for (const t of threads.values()) t.terminate?.();
      threads.clear();
    },
  };
}
