/**
 * The level and text of a renderer's console-message, whichever shape
 * Electron sends.
 *
 * The positional arguments — (event, level, message, line, source) — are
 * deprecated in favour of one details object whose level is a word
 * ("warning", "error"). Electron 43 still sends both. Reading only the old
 * ones means that on the day it stops, `level >= 2` becomes
 * `undefined >= 2`, and every check that asserts a window logged nothing
 * passes because nothing is ever seen again. That is the worst way for a
 * check to fail, so both shapes are read, in one place.
 */
const WORDS = { verbose: 0, debug: 0, info: 1, log: 1, warning: 2, warn: 2, error: 3 };

export function consoleMessage(args) {
  const [first, level, message, line, source] = args;
  if (typeof level === 'number') return { level, text: String(message ?? ''), line, source: String(source ?? '') };
  return {
    level: WORDS[String(first?.level ?? '').toLowerCase()] ?? 1,
    text: String(first?.message ?? ''),
    line: first?.lineNumber,
    source: String(first?.sourceId ?? ''),
  };
}
