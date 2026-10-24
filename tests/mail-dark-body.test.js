// A message's own colours on Mail's dark reading frame, by the rule a
// sheet's cells follow: the sender's colour on the sender's fill is kept,
// a colour with no fill under it is lifted (keeping its hue) until it reads
// on the dark frame, and words in a box the sender filled light get the
// light theme's ink, links and quotes too. The window checks
// (RUTBA_VERIFY_ONLY=dark-pass) measure a newsletter on screen.
import test from 'node:test';
import assert from 'node:assert/strict';
import { contrast, parseColour } from '@rutba/office-ui/contrast';

const { bodyDocument, darkBody } = await import('../apps/desktop/renderer/apps/mail/parts.js');
const NIGHT = '#1e2127';
const inkOf = (html, tag) => {
  const m = new RegExp(`<${tag}\\b[^>]*style="[^"]*?color: ([^;"]+)`, 'i').exec(html);
  return m ? m[1].trim() : null;
};

test('a colour with no fill under it is lifted to read on the dark frame, keeping its hue', () => {
  const out = darkBody('<div style="font-family: Arial; color: #333333"><h2 style="color: #1F3A5F">Title</h2></div>');
  const div = inkOf(out, 'div');
  const h2 = inkOf(out, 'h2');
  assert.ok(contrast(div, NIGHT) >= 4.5, div);
  assert.ok(contrast(h2, NIGHT) >= 4.5, h2);
  const [r, g, b] = parseColour(h2);
  assert.ok(b > g && g > r, `still blue: ${h2}`);
  // The rest of the style is kept.
  assert.match(out, /font-family: Arial/);
  // A <font color> is lifted too.
  assert.ok(contrast(inkOf(darkBody('<font color="navy">Hello</font>'), 'font'), NIGHT) >= 4.5);
});

test('the sender\'s colour on the sender\'s fill is kept, and a light box\'s words without a colour of their own get the day\'s ink', () => {
  const html = '<div style="color: #333333"><div style="background: #F2F4F7; padding: 10px">Orders <a href="x">here</a></div>'
    + '<table bgcolor="#ffffff"><tr><td>In a white table</td></tr></table>'
    + '<p style="background-color: #1F3A5F; color: #ffffff">White on navy</p>'
    + '<div style="background: #FFF4CE"><span style="color: #9C5700">Amber on cream</span><blockquote>Quoted</blockquote></div></div>';
  const out = darkBody(html);
  // The pale box: the sender's own #333 (it reads there), its link the day's blue.
  assert.match(out, /background: #F2F4F7; padding: 10px; color: #333333"/);
  assert.match(out, /<a href="x" style="color: #1a56c4">/);
  // A table filled white: its words dark.
  assert.match(out, /<table bgcolor="#ffffff" style="color: #333333">/);
  // White on navy, amber on cream: kept as they are.
  assert.match(out, /<p style="background-color: #1F3A5F; color: #ffffff">/);
  assert.match(out, /<span style="color: #9C5700">/);
  // The quote in the cream box takes the day's quote grey, not the night's.
  assert.match(out, /<blockquote style="color: #5b626d">/);
});

test('a fill on the message\'s own body tag is not taken for one (the frame\'s body is drawn dark)', () => {
  const out = darkBody('<body bgcolor="#ffffff" style="color: #222222"><p>Hello</p></body>');
  const ink = inkOf(out, 'body');
  assert.ok(contrast(ink, NIGHT) >= 4.5, ink);
});

test('bodyDocument lifts colours only at night, and stays quick on a hostile message', () => {
  const html = '<p style="color: #333333">Words</p>';
  assert.match(bodyDocument({ html }, { dark: false }), /<p style="color: #333333">/);
  assert.doesNotMatch(bodyDocument({ html }, { dark: true }), /<p style="color: #333333">/);
  // Quick is the work it does, timed by this thread's own CPU time: the clock
  // on the wall also counts the turns the machine gave to the gate's other
  // lanes, and beside them passed 1.5 s on work of about 0.2 s.
  const cpu = typeof process.threadCpuUsage === 'function' ? () => process.threadCpuUsage() : () => process.cpuUsage();
  for (const hostile of ['<div>'.repeat(50000) + '</x>'.repeat(50000), '<p style="color:#333">'.repeat(30000), '<'.repeat(100000)]) {
    const started = Date.now();
    const was = cpu();
    bodyDocument({ html: hostile }, { dark: true });
    const now = cpu();
    const ms = (now.user - was.user + now.system - was.system) / 1000;
    assert.ok(ms < 1500, `${Math.round(ms)} ms of CPU time, ${Date.now() - started} ms by the clock`);
  }
});
