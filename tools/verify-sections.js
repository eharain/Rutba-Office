// The sections of tools' window checks, read from the source: each
// `await section('<app>', '<name>', async () => {` and the brace that ends it,
// found by a scanner that steps over strings, template literals (and the code
// inside their ${}), comments and regular expressions. tests/gate-sections.test.js
// holds apps/desktop/main/verify-apps.js to it: every section named once, and
// no part inside a section, so a part asked for runs alone (owner, 11 Nov 2026).

/** The sections in `src`: { app, name, start, open, close, body }. */
export function sectionsIn(src) {
  const out = [];
  const re = /\bawait section\('(\w+)', '([\w-]+)', async \(\) => \{/g;
  let m;
  while ((m = re.exec(src))) {
    const open = m.index + m[0].length - 1;
    const close = matchBrace(src, open);
    out.push({ app: m[1], name: m[2], start: m.index, open, close, body: close > 0 ? src.slice(open + 1, close) : '' });
  }
  return out;
}

/** The index of the brace that closes the one at `at`, or -1. */
export function matchBrace(s, at) {
  // Stack of contexts: 'code' (with brace depth) or 'tpl'.
  const stack = [{ k: 'code', d: 0 }];
  let prevSig = '';
  for (let i = at; i < s.length; i++) {
    const c = s[i];
    const top = stack[stack.length - 1];
    if (top.k === 'tpl') {
      if (c === '\\') { i++; continue; }
      if (c === '`') { stack.pop(); prevSig = ')'; continue; }
      if (c === '$' && s[i + 1] === '{') { stack.push({ k: 'code', d: 1, inTpl: true }); i++; continue; }
      continue;
    }
    if (c === '/' && s[i + 1] === '/') { const e = s.indexOf('\n', i); i = e < 0 ? s.length : e; continue; }
    if (c === '/' && s[i + 1] === '*') { const e = s.indexOf('*/', i + 2); i = e < 0 ? s.length : e + 1; continue; }
    if (c === '"' || c === "'") {
      let j = i + 1;
      while (j < s.length && s[j] !== c) { if (s[j] === '\\') j++; j++; }
      i = j; prevSig = ')'; continue;
    }
    if (c === '`') { stack.push({ k: 'tpl' }); continue; }
    if (c === '/' && /[(,=:[!&|?{};]|^$|return$/.test(prevSig)) {
      // A regex literal.
      let j = i + 1, cls = false;
      while (j < s.length) {
        if (s[j] === '\\') { j += 2; continue; }
        if (s[j] === '[') cls = true; else if (s[j] === ']') cls = false; else if (s[j] === '/' && !cls) break; else if (s[j] === '\n') break;
        j++;
      }
      i = j; prevSig = ')'; continue;
    }
    if (c === '{') { top.d++; }
    else if (c === '}') {
      top.d--;
      if (top.d === 0) {
        if (top.inTpl) { stack.pop(); continue; }
        if (stack.length === 1) return i;
      }
    }
    if (!/\s/.test(c)) {
      if (/[\w$]/.test(c)) {
        let j = i; while (j < s.length && /[\w$]/.test(s[j])) j++;
        prevSig = s.slice(i, j); i = j - 1;
      } else prevSig = c;
    }
  }
  return -1;
}
