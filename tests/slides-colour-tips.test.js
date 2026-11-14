// Slides: a colour button that cannot act yet says why. Every control of the
// Font group is greyed until a text box is chosen, and its tip ends with
// Slides' own words for that ("Select a text box first"). The two colour
// split buttons came from the shared palette, whose tips had no room for
// those words, so they were the only two that greyed out without saying why.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { asEnglish } from './label-english.js';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (p) => asEnglish(fs.readFileSync(path.join(root, p), 'utf8'));

test('the shared colour split button passes both of its tips through the app\'s explain', () => {
  const src = read('apps/desktop/renderer/apps/colour/palette.js');
  const head = src.match(/export function ColourSplit\(\{([^}]*)\}\)/);
  assert.ok(head, 'ColourSplit is exported');
  assert.match(head[1], /\bexplain = null\b/, 'explain is a prop, off unless an app gives it');
  const face = src.match(/className="rw-split-face"[\s\S]*?title=\{([^\n]*)\}\s*onClick/);
  const arrow = src.match(/className="rw-split-arrow"[\s\S]*?title=\{([^\n]*)\}\s*onClick/);
  assert.ok(face && /^tip\(/.test(face[1]), `the face's tip goes through explain: ${face?.[1]}`);
  assert.ok(arrow && /^tip\(/.test(arrow[1]), `the arrow's tip goes through explain: ${arrow?.[1]}`);
  assert.match(src, /const tip = \(text\) => \(explain \? explain\(text\) : text\);/);
});

test('every colour split button Slides greys out until a text box is chosen says why, as the rest of its Font group does', () => {
  const src = read('apps/desktop/renderer/apps/slides/ribbon.js');
  const splits = [...src.matchAll(/<ColourSplit\b([\s\S]*?)\/>/g)].map((m) => m[1]);
  const greyed = splits.filter((s) => /disabled=\{!hasShape\}/.test(s));
  assert.ok(greyed.length >= 2, 'Font colour and Text highlight colour are greyed without a text box');
  for (const s of greyed) assert.match(s, /explain=\{need\}/, `${s.match(/kind="([^"]+)"/)?.[1]} says why it is greyed`);
  // The other controls of the group already do it the same way.
  assert.match(src, /title=\{need\(t\('Change case'\)\)\} disabled=\{!hasShape\}/);
});
