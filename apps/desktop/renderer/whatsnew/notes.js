// The release notes, read into blocks a dialog can draw. Pure, so the engine
// tests read it without a window.

/**
 * The bold and code inside one line, as runs. The notes are ours, so only
 * the two marks they use are understood.
 */
export function runsOf(text) {
  const runs = [];
  const re = /\*\*([^*]+)\*\*|`([^`]+)`/g;
  let at = 0;
  let m;
  while ((m = re.exec(text))) {
    if (m.index > at) runs.push({ text: text.slice(at, m.index) });
    if (m[1] != null) runs.push({ text: m[1], bold: true });
    else runs.push({ text: m[2], code: true });
    at = m.index + m[0].length;
  }
  if (at < text.length) runs.push({ text: text.slice(at) });
  return runs;
}

/**
 * The notes read into blocks: a heading, a paragraph, a bullet. The notes
 * are Markdown of the kind the releases are written in — `##` headings,
 * `-` bullets that continue on indented lines, a blank line between
 * paragraphs — and nothing else is understood, on purpose.
 */
export function parseNotes(md) {
  const blocks = [];
  let para = null;
  const flush = () => {
    if (para) blocks.push({ kind: 'p', runs: runsOf(para.join(' ')) });
    para = null;
  };
  for (const raw of String(md || '').replace(/\r\n/g, '\n').split('\n')) {
    const line = raw.replace(/\s+$/, '');
    const heading = /^(#{1,6})\s+(.*)$/.exec(line);
    if (heading) {
      flush();
      blocks.push({ kind: 'h', level: heading[1].length, text: heading[2].trim() });
      continue;
    }
    const bullet = /^\s*[-*]\s+(.*)$/.exec(line);
    if (bullet) {
      flush();
      blocks.push({ kind: 'li', text: bullet[1], runs: runsOf(bullet[1]) });
      continue;
    }
    if (line === '') {
      flush();
      continue;
    }
    const last = blocks.at(-1);
    if (/^\s+/.test(raw) && last?.kind === 'li' && !para) {
      // An indented line after a bullet continues it.
      last.text += ' ' + line.trim();
      last.runs = runsOf(last.text);
      continue;
    }
    (para ??= []).push(line.trim());
  }
  flush();
  return blocks;
}
