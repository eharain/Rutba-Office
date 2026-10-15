// The suite's own accessibility, the parts that need no window: a control's
// name and description as they are read from its tip; a slide drawn as a
// picture named by its title; and the light theme's quieter greys and each
// app's colour as words reading 4.5:1 (WCAG 2 AA) on what they are drawn on,
// worked out from theme.css itself. The windows are audited by the checks
// a11y-audit, keyboard-pass, screen-reader and high-contrast.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { splitTip } from '../packages/office-ui/src/a11y.js';
import { renderSlide, slideTitle } from '../packages/presentation/src/render.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const CSS = fs.readFileSync(path.join(ROOT, 'packages/office-ui/src/theme.css'), 'utf8');
const LONG = String.fromCharCode(0x2014);

test("a tip's name is what comes before its long dash, the rest is its description", () => {
  assert.deepEqual(splitTip(`Rutba ${LONG} apply this theme to every slide`), { name: 'Rutba', more: 'apply this theme to every slide' });
  assert.deepEqual(splitTip('Bold (Ctrl+B)'), { name: 'Bold (Ctrl+B)', more: '' });
  assert.deepEqual(splitTip(`Bold (Ctrl+B) ${LONG} Select a text box first ${LONG} click it once`), { name: 'Bold (Ctrl+B)', more: `Select a text box first ${LONG} click it once` });
  assert.deepEqual(splitTip(''), { name: '', more: '' });
});

test('a slide drawn on its own is a picture named by its title, or by its first words', () => {
  const run = (text) => ({ paragraphs: [{ runs: [{ text }] }] });
  const slide = { size: { width: 960, height: 540 }, shapes: [
    { id: '2', geometry: { x: 0, y: 300, w: 900, h: 100 }, text: run('Sales up in three regions') },
    { id: '1', placeholder: { type: 'title' }, geometry: { x: 0, y: 0, w: 900, h: 100 }, text: run('Where we stand') },
  ] };
  assert.equal(slideTitle(slide), 'Where we stand');
  assert.match(renderSlide(slide), /<svg [^>]*role="img" aria-label="Where we stand"/);
  assert.equal(slideTitle({ shapes: [{ text: run('Only words') }] }), 'Only words');
  assert.match(renderSlide(slide, { label: 'Slide 2 of 4: "A" & B' }), /aria-label="Slide 2 of 4: &quot;A&quot; &amp; B"/);
  // Nothing to call it: no picture role with an empty name.
  assert.doesNotMatch(renderSlide({ size: { width: 960, height: 540 }, shapes: [] }), /role="img"/);
});

// Colour arithmetic, as WCAG 2 measures it.
const hex = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
const lum = ([r, g, b]) => { const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; }; return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b); };
const ratio = (a, b) => { const x = lum(a), y = lum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); };
const mix = (a, b, p) => a.map((v, i) => v * p + b[i] * (1 - p));
const token = (name) => {
  const m = new RegExp(`^\\s*--${name}:\\s*(#[0-9a-f]{6});`, 'mi').exec(CSS);
  assert.ok(m, `--${name} is a colour in theme.css`);
  return hex(m[1]);
};

test("the light theme's quieter greys read 4.5:1 on the chrome, the panes and a chosen row", () => {
  const under = { surface: token('n-0'), chrome: token('n-5'), pane: token('n-10'), sunken: token('n-20') };
  for (const ink of ['ink-2', 'ink-3']) {
    for (const [where, bg] of Object.entries(under)) assert.ok(ratio(token(ink), bg) >= 4.5, `--${ink} on ${where}: ${ratio(token(ink), bg).toFixed(2)}`);
    // A chosen row in each app: its colour at 20% over white.
    for (const app of ['mail', 'word', 'sheets', 'slides', 'pictures', 'image', 'video', 'pdf', 'calendar', 'contacts']) {
      const accent = hex(new RegExp(`\\[data-app='${app}'\\] \\{ --accent: (#[0-9a-f]{6}); \\}`).exec(CSS)[1]);
      if (ink === 'ink-3') assert.ok(ratio(token(ink), mix(accent, [255, 255, 255], 0.12)) >= 4.5, `--ink-3 on a chosen row in ${app}`);
    }
  }
});

test("by day each app's colour as words reads 4.5:1 on the chrome and a chosen row, and white reads on it as a fill", () => {
  const W = [255, 255, 255];
  const chrome = token('n-5');
  for (const app of ['mail', 'word', 'sheets', 'slides', 'pictures', 'image', 'video', 'pdf', 'calendar', 'contacts', 'home']) {
    const accent = hex(new RegExp(`\\[data-app='${app}'\\] \\{ --accent: (#[0-9a-f]{6}); \\}`).exec(CSS)[1]);
    const rule = new RegExp(`:root:not\\(\\[data-theme='dark'\\]\\) [^{]*\\[data-app='${app}'\\][^{]*\\{([^}]*)\\}`).exec(CSS)?.[1] || '';
    const pct = (name) => { const m = new RegExp(`--${name}: color-mix\\(in srgb, var\\(--accent\\) (\\d+)%, black\\)`).exec(rule); return m ? Number(m[1]) / 100 : 1; };
    const text = mix(accent, [0, 0, 0], pct('accent-text'));
    const fill = mix(accent, [0, 0, 0], pct('accent-fill'));
    for (const [where, bg] of [['the chrome', chrome], ['a chosen row', mix(accent, W, 0.2)], ['the wash', mix(accent, chrome, 0.17)]]) {
      assert.ok(ratio(text, bg) >= 4.5, `${app}: words in its colour on ${where}: ${ratio(text, bg).toFixed(2)}`);
    }
    assert.ok(ratio(W, fill) >= 4.5, `${app}: white on its fill: ${ratio(W, fill).toFixed(2)}`);
  }
});

test('in Windows High Contrast what a fill or a shadow said is said again in the system colours', () => {
  const at = CSS.indexOf('@media (forced-colors: active)');
  assert.ok(at > 0, 'theme.css has a forced-colours block');
  const block = CSS.slice(at);
  // The chosen tab's bar, a pressed or chosen control, the keyboard's ring, and the edges of what floats.
  assert.match(block, /\.rw-tab\[aria-selected='true'\]::before \{ forced-color-adjust: none; background: Highlight; \}/);
  assert.match(block, /\.rw-btn\[aria-pressed='true'\][^{]*\[aria-current='date'\][^{]*\{\s*outline: 2px solid Highlight;/);
  assert.match(block, /:focus-visible \{ outline: 2px solid Highlight;/);
  assert.match(block, /\.rw-dialog, \.rw-menu[^{]*\{ border: 1px solid CanvasText; \}/);
  assert.match(block, /\.rw-btn:disabled svg \{ color: GrayText !important; \}/);
});
