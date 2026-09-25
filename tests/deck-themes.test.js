// Design → Themes, Variants, Colours, Fonts and Effects — the deck engine's half.
//
// A theme is the theme part every master points at; applying one rewrites
// that part (and the master's background and colour map), never a slide.
// What a slide states for itself stays; what it inherits follows the theme.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide, THEMES, PALETTES, FONT_PAIRS, EFFECT_PRESETS, COLOUR_SLOTS } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({
  title: 'Themes',
  slides: [
    { layout: 'title', title: 'Quarterly review', body: 'Spring' },
    { layout: 'obj', title: [{ runs: [{ text: 'Own red title', color: '#C00000', font: 'Courier New', size: 30 }] }], body: ['One', 'Two'] },
  ],
});

const MICROSOFT_THEME_NAMES = ['Office Theme', 'Office', 'Facet', 'Integral', 'Ion', 'Ion Boardroom', 'Organic', 'Retrospect', 'Slice', 'Wisp', 'Banded', 'Basis', 'Berlin', 'Celestial', 'Circuit', 'Damask', 'Depth', 'Dividend', 'Droplet', 'Frame', 'Gallery', 'Main Event', 'Mesh', 'Metropolitan', 'Parallax', 'Quotable', 'Savon', 'View', 'Vapor Trail', 'Wood Type', 'Madison', 'Atlas', 'Badge', 'Crop', 'Feathered', 'Headlines', 'Slate', 'Madison'];

const titleOf = (deck, i) => deck.slide(i).shapes.find((s) => /title/i.test(s.placeholder?.type || ''));
const themeXml = (deck) => deck.pkg.text('ppt/theme/theme1.xml');

test('the suite ships at least eight themes of its own, each a complete palette, a font pair, effects and four variants', () => {
  assert.ok(THEMES.length >= 8, `${THEMES.length} themes`);
  const names = new Set();
  for (const t of THEMES) {
    assert.ok(!MICROSOFT_THEME_NAMES.includes(t.name), `${t.name} is not one of Microsoft's names`);
    assert.ok(!names.has(t.name), `${t.name} is unique`);
    names.add(t.name);
    for (const k of COLOUR_SLOTS) assert.match(t.palette[k], /^[0-9A-F]{6}$/, `${t.name} ${k}`);
    assert.ok(t.fonts.major && t.fonts.minor);
    assert.ok(EFFECT_PRESETS.some((p) => p.id === t.effects), `${t.name} effects`);
  }
  const deck = Deck.open(DECK);
  deck.applyTheme('harbour');
  assert.equal(deck.variants(0).length, 4);
  assert.ok(PALETTES.length > THEMES.length && FONT_PAIRS.length >= 8 && EFFECT_PRESETS.length >= 3);
});

test('applying a theme writes a schema-ordered theme part, the master background and colour map, and leaves every slide byte-identical', () => {
  const deck = Deck.open(DECK);
  const slidesBefore = deck.slideParts.map((s) => deck.pkg.text(s.part));
  assert.equal(deck.applyTheme('ember'), true);
  const xml = themeXml(deck);
  assert.match(xml, /^<\?xml[^>]*\?>\r?\n<a:theme xmlns:a="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/main" name="Ember"><a:themeElements><a:clrScheme name="Ember">/);
  const slots = [...xml.matchAll(/<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>/g)].map((m) => m[1]);
  assert.deepEqual(slots, COLOUR_SLOTS, 'twelve slots in the schema\'s order');
  assert.match(xml, /<a:fontScheme name="Ember"><a:majorFont><a:latin typeface="Georgia"\/><a:ea typeface=""\/><a:cs typeface=""\/><\/a:majorFont><a:minorFont><a:latin typeface="Segoe UI"/);
  for (const list of ['fillStyleLst', 'lnStyleLst', 'effectStyleLst', 'bgFillStyleLst']) {
    const body = new RegExp(`<a:${list}>([\\s\\S]*?)</a:${list}>`).exec(xml)?.[1] || '';
    const kids = body.match(/<a:(solidFill|gradFill|ln|effectStyle)\b/g) || [];
    assert.ok(kids.length >= 3, `${list} holds three entries`);
  }
  assert.match(xml, /<\/a:fmtScheme><\/a:themeElements><a:objectDefaults\/><a:extraClrSchemeLst\/><\/a:theme>$/);
  const master = deck.pkg.text('ppt/slideMasters/slideMaster1.xml');
  assert.match(master, /<p:cSld><p:bg><p:bgPr><a:gradFill[^>]*><a:gsLst>(<a:gs pos="\d+"><a:schemeClr val="(bg1|accent1)"\/><\/a:gs>)+<\/a:gsLst>/, 'the master takes the theme\'s background, in theme colours');
  assert.match(master, /<p:clrMap bg1="lt1" tx1="dk1"/);
  assert.deepEqual(deck.slideParts.map((s) => deck.pkg.text(s.part)), slidesBefore, 'no slide is rewritten');
  // Saved and opened again, it is the same theme.
  const again = Deck.open(deck.save());
  assert.equal(again.designInfo(0).builtIn, 'ember');
  assert.equal(again.designInfo(0).fonts.major, 'Georgia');
  assert.ok(OoxmlPackage.read(deck.save()).has('ppt/theme/theme1.xml'));
});

test('a slide keeps what it states for itself; what it inherits follows the theme', () => {
  const deck = Deck.open(DECK);
  deck.applyTheme('nightfall');
  const plain = titleOf(deck, 0);
  assert.equal(plain.textStyle[0].color, '#f4f6fb', 'a title with no colour of its own takes the dark theme\'s light text');
  assert.equal(plain.textStyle[0].font, 'Segoe UI Semibold');
  const own = titleOf(deck, 1);
  const run = own.text.paragraphs[0].runs[0];
  assert.equal(run.color.toLowerCase(), '#c00000', 'the red run stays red');
  assert.equal(run.font, 'Courier New');
  assert.equal(run.size, 30);
  const svg = renderSlide(deck.slide(1));
  assert.match(svg, /fill="#C00000"|fill="#c00000"/i, 'drawn red');
  assert.match(svg, /<rect x="0" y="0" width="1280" height="720" fill="url\(#g1\)"/, 'the dark theme\'s background is drawn');
  assert.match(renderSlide(deck.slide(0)), /fill="#f4f6fb"/, 'the inherited title is drawn in the theme\'s light text');
});

test('Variants: the fourth turns a light theme dark through the master\'s colour map, and is ticked as current', () => {
  const deck = Deck.open(DECK);
  deck.applyTheme('harbour');
  assert.equal(deck.designInfo(0).variant, 0);
  const vs = deck.variants(0);
  assert.equal(vs[3].dark, true);
  deck.applyVariant(3);
  const info = deck.designInfo(0);
  assert.equal(info.dark, true);
  assert.equal(info.variant, 3);
  assert.match(deck.pkg.text('ppt/slideMasters/slideMaster1.xml'), /<p:clrMap bg1="dk1" tx1="lt1"/);
  assert.equal(titleOf(deck, 0).textStyle[0].color, `#${vs[3].colors.lt1.toLowerCase()}`);
  deck.applyVariant(1);
  assert.equal(deck.designInfo(0).dark, false);
  assert.equal(deck.designInfo(0).colors.accent1, vs[1].colors.accent1);
});

test('Colours: a built-in palette or Customise Colours\' twelve slots rewrite only the colour scheme', () => {
  const deck = Deck.open(DECK);
  const before = themeXml(deck);
  const fontAndFormat = (xml) => /<a:fontScheme[\s\S]*<\/a:fmtScheme>/.exec(xml)[0];
  assert.equal(deck.setThemeColors('lagoon'), true);
  assert.equal(fontAndFormat(themeXml(deck)), fontAndFormat(before), 'fonts and effects untouched');
  assert.equal(deck.designInfo(0).colorName, 'Lagoon');
  assert.equal(deck.designInfo(0).colors.accent1, '0096C7');
  const custom = Object.fromEntries(COLOUR_SLOTS.map((k, i) => [k, i % 2 ? '112233' : 'AABBCC']));
  custom.accent1 = 'FF6600';
  deck.setThemeColors(custom, 'My colours');
  assert.equal(deck.designInfo(0).colorName, 'My colours');
  assert.equal(deck.designInfo(0).colors.accent1, 'FF6600');
  assert.throws(() => deck.setThemeColors({ ...custom, accent2: 'orange' }, 'Bad'), /accent2/);
  assert.throws(() => deck.setThemeColors('no-such-palette'), /no palette/);
});

test('Fonts: a pair or two faces of one\'s own — titles in the heading face, the rest in the body face', () => {
  const deck = Deck.open(DECK);
  deck.setThemeFonts('classic');
  assert.equal(deck.designInfo(0).fontName, 'Classic');
  assert.equal(titleOf(deck, 0).textStyle[0].font, 'Georgia');
  deck.setThemeFonts({ major: 'Trebuchet MS', minor: 'Verdana' }, 'Mine');
  const body = deck.slide(1).shapes.find((s) => s.placeholder?.type === 'body');
  assert.equal(titleOf(deck, 0).textStyle[0].font, 'Trebuchet MS');
  assert.equal(body.textStyle[0].font, 'Verdana');
  assert.match(renderSlide(deck.slide(1)), /font-family="Verdana, Segoe UI/);
  assert.match(renderSlide(deck.slide(0)), /font-family="Trebuchet MS, /);
  assert.throws(() => deck.setThemeFonts({ major: '', minor: 'Arial' }), /heading font/);
});

test('Effects: a shape styled from the theme is filled, outlined and lifted by the format scheme', () => {
  const pkg = OoxmlPackage.read(DECK);
  const part = 'ppt/slides/slide2.xml';
  const styled = '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Styled"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:style><a:lnRef idx="3"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Styled</a:t></a:r></a:p></p:txBody></p:sp>';
  pkg.write_(part, pkg.text(part).replace('</p:spTree>', `${styled}</p:spTree>`));
  const deck = Deck.open(pkg.write());
  const shape = () => deck.slide(1).shapes.find((s) => s.name === 'Styled');
  deck.setThemeEffects('lifted');
  assert.equal(deck.designInfo(0).effects, 'lifted');
  assert.ok(shape().effects?.shadow, 'lifted: the intense style has a shadow');
  assert.equal(shape().fill.type, 'gradient', 'lifted: the intense fill is a gradient');
  assert.ok(shape().line.width > 2, `lifted: the intense line is heavy (${shape().line.width} pt)`);
  assert.equal(shape().textStyle[0].color, '#ffffff', 'the font reference colours its words');
  assert.match(renderSlide(deck.slide(1)), /feDropShadow|<feOffset/);
  deck.setThemeEffects('flat');
  assert.equal(shape().effects, null, 'flat: no shadow');
  assert.equal(shape().fill.type, 'solid');
  assert.throws(() => deck.setThemeEffects('glitter'), /no effects/);
});

test('a theme, its colours, fonts and effects are one undo step each, and redo puts them back', () => {
  const deck = Deck.open(DECK);
  const theme0 = themeXml(deck);
  const master0 = deck.pkg.text('ppt/slideMasters/slideMaster1.xml');
  assert.equal(deck.canUndo, false);
  let snap = deck.snapshot();
  deck.applyTheme('saffron');
  deck.pushUndo(snap);
  snap = deck.snapshot();
  deck.setThemeFonts('plain');
  deck.pushUndo(snap);
  assert.equal(deck.designInfo(0).fonts.major, 'Arial');
  assert.equal(deck.undo(), true);
  assert.equal(deck.designInfo(0).name, 'Saffron');
  assert.equal(deck.designInfo(0).fonts.major, 'Trebuchet MS');
  assert.equal(deck.undo(), true);
  assert.equal(themeXml(deck), theme0, 'the theme part as it was, byte for byte');
  assert.equal(deck.pkg.text('ppt/slideMasters/slideMaster1.xml'), master0);
  assert.equal(titleOf(deck, 0).textStyle[0].font, 'Segoe UI Semibold', 'and the slides drawn from it again');
  assert.equal(deck.canUndo, false);
  assert.equal(deck.canRedo, true);
  deck.redo();
  assert.equal(deck.designInfo(0).name, 'Saffron');
  assert.equal(titleOf(deck, 0).textStyle[0].font, 'Trebuchet MS');
});

test('a preview draws the slide under another theme without writing a byte', () => {
  const deck = Deck.open(DECK);
  const before = Buffer.from(deck.save()).toString('base64');
  const scene = deck.previewSlide(0, { theme: 'tidewater' });
  assert.equal(scene.background.type, 'gradient');
  assert.equal(scene.background.stops[0].color, '#0f2a2e', 'the dark theme\'s background');
  assert.equal(scene.shapes.find((s) => s.placeholder?.type === 'ctrTitle').textStyle[0].font, 'Georgia');
  assert.ok(deck.previewSlide(0, { colors: 'berry' }));
  assert.ok(deck.previewSlide(0, { fonts: 'bookish' }));
  assert.ok(deck.previewSlide(0, { variant: 2 }));
  assert.equal(deck.dirty, false);
  assert.equal(Buffer.from(deck.save()).toString('base64'), before);
  assert.equal(deck.designInfo(0).builtIn, 'rutba');
});

test('a theme from elsewhere keeps its own fonts, effects and extensions when only its colours change', () => {
  const pkg = OoxmlPackage.read(DECK);
  const foreign = pkg.text('ppt/theme/theme1.xml')
    .replace(/name="Rutba"/g, 'name="Somebody Else"')
    .replace('<a:extraClrSchemeLst/>', '<a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Somebody Else" id="{11111111-2222-3333-4444-555555555555}" vid="{66666666-7777-8888-9999-000000000000}"/></a:ext></a:extLst>');
  pkg.write_('ppt/theme/theme1.xml', foreign);
  const deck = Deck.open(pkg.write());
  assert.equal(deck.designInfo(0).builtIn, null);
  assert.equal(deck.variants(0).length, 4, 'variants drawn from its own colours');
  deck.setThemeColors('citrus');
  const xml = themeXml(deck);
  assert.match(xml, /name="Somebody Else"><a:themeElements><a:clrScheme name="Citrus">/);
  assert.match(xml, /<thm15:themeFamily /, 'its extension kept');
  assert.equal(/<a:fontScheme[\s\S]*<\/a:fmtScheme>/.exec(xml)[0], /<a:fontScheme[\s\S]*<\/a:fmtScheme>/.exec(foreign)[0]);
});
