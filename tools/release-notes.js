// The release notes, as the GitHub release page should show them.
//
// docs/releases/<tag>.md is wrapped at about 75 characters so it reads well
// in an editor and in a diff. GitHub's release page turns every line break
// in a release body into a hard break, so a body uploaded as written shows
// each paragraph and each list item broken mid-sentence. Here the wrapped
// lines of a paragraph or a list item are joined back into one, and
// everything that carries meaning in its lines is left as it is: blank
// lines, headings, the start of each list item, quotes, tables, rules and
// code blocks.

const BLOCK_START = /^\s*(?:[-*+]\s|\d+[.)]\s|#|>|\|)/;
const RULE = /^\s{0,3}(?:-{3,}|\*{3,}|_{3,})\s*$/;
const FENCE = /^\s*(?:```|~~~)/;

/** `markdown` with each paragraph's and list item's wrapped lines joined. */
export function unwrapNotes(markdown) {
  const out = [];
  let inFence = false;
  for (const line of String(markdown).replace(/\r\n/g, '\n').split('\n')) {
    if (FENCE.test(line)) {
      inFence = !inFence;
      out.push(line);
      continue;
    }
    if (inFence) {
      out.push(line);
      continue;
    }
    const prev = out.length ? out[out.length - 1] : null;
    const continues =
      line.trim() !== '' &&
      !BLOCK_START.test(line) &&
      !RULE.test(line) &&
      prev !== null &&
      prev.trim() !== '' &&
      !/^\s*(?:#|\|)/.test(prev) &&
      !RULE.test(prev) &&
      !FENCE.test(prev);
    if (continues) out[out.length - 1] = `${prev.replace(/\s+$/, '')} ${line.trim()}`;
    else out.push(line);
  }
  return out.join('\n');
}
