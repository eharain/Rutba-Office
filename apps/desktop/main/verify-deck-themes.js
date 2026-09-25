// Presentation: Design → Themes, Variants, Colours, Fonts and Effects.
//
// On a deck of its own (three slides and a shape styled from the theme,
// written beside the shared fixtures so no other block sees it): the Design
// tab's strips show this slide drawn in the themes and the variants, the
// full gallery opens with a live card per theme, and a pick restyles every
// slide — read back from the engine and seen in the stage's own drawing and
// in the strip's thumbnails. A dark variant, a palette, Customise Colours,
// a font pair, Customise Fonts and a format scheme follow, each one undo
// step; the saved file opens again with the design it was given. Run alone
// with RUTBA_VERIFY_ONLY=themes.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { consoleMessage } from './console-message.js';

/** Three slides and a rectangle whose fill, line and effect come from the theme — PowerPoint's own shape style. */
export function makeThemeFixture(file) {
  const bytes = buildPptx({
    title: 'Design checks',
    slides: [
      { layout: 'title', title: 'Northwind quarterly review', body: 'Spring results and the plan ahead' },
      { layout: 'obj', title: 'What went well', body: ['Orders up by a fifth', 'Two new regions opened', 'Returns at their lowest'] },
      { layout: 'obj', title: 'Next quarter', body: ['Hire for support', 'Ship the mobile app'] },
    ],
  });
  const pkg = OoxmlPackage.read(bytes);
  const part = 'ppt/slides/slide2.xml';
  const styled = '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Highlight"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="8200000" y="2500000"/><a:ext cx="2800000" cy="1500000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>' +
    '<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400" b="1"/><a:t>+21%</a:t></a:r></a:p></p:txBody></p:sp>';
  pkg.write_(part, pkg.text(part).replace('</p:spTree>', `${styled}</p:spTree>`));
  fs.writeFileSync(file, pkg.write());
  return file;
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where to write the block's own deck
 */
export async function verifyDeckThemes(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  try {
    const file = makeThemeFixture(path.join(dir, 'themes.pptx'));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const consoleErrors = [];
    wc.on('console-message', (...args) => { const m = consoleMessage(args); if (m.level >= 3) consoleErrors.push(m.text.split('\n')[0].slice(0, 160)); });
    const id = () => sessionFor('deck').id;
    const model = (slide = 0) => doc.model({ id: id(), slide });
    const info = () => model(0).design || {};
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn, .rw-ribbon .sl-rs-more')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.click();
      return 'clicked';
    })()`);
    const clickIn = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'missing ' + ${JSON.stringify(selector)}; b.click(); return 'clicked'; })()`);
    const setField = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no field ' + ${JSON.stringify(selector)};
      el.focus();
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'set';
    })()`);
    const stageHas = (needle) => js(`(document.querySelector('.sl-svg')?.innerHTML || '').includes(${JSON.stringify(needle)})`);
    const thumbHas = (i, needle) => js(`(document.querySelectorAll('.sl-thumb .sl-thumb-pic')[${i}]?.innerHTML || '').includes(${JSON.stringify(needle)})`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 3`), 'the slide sorter', 8000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Design')?.click(), 'tab'`);

    // 1. The strips: this slide drawn in four themes and the four variants.
    const strips = await until(() => js(`document.querySelectorAll('.sl-rs-themes .sl-rs-item svg').length >= 4 && document.querySelectorAll('.sl-rs-variants .sl-rs-item svg').length === 4`), 'the Design strips', 8000).catch(() => false);
    const stripCurrent = await js(`document.querySelector('.sl-rs-themes .sl-rs-item.current')?.dataset.theme || null`);
    check('slides: Design shows a strip of themes and the four variants, each this slide drawn live, the deck\'s own theme ringed',
      strips === true && stripCurrent === 'rutba', `strips ${strips}; current ${stripCurrent}`);
    await capture('slides-design-tab.png');

    // 2. The full gallery.
    const more = await clickRibbon('More Themes');
    const gallery = await until(() => js(`document.querySelectorAll('.sl-dg-themes .sl-dg-card svg').length >= 8`), 'the theme gallery', 8000).catch(() => false);
    const names = await js(`[...document.querySelectorAll('.sl-dg-themes .sl-dg-card .sl-dg-label')].map((c) => c.textContent.trim())`);
    const distinct = await js(`new Set([...document.querySelectorAll('.sl-dg-themes .sl-dg-card .sl-dg-pic')].map((p) => p.innerHTML.length + ':' + (p.innerHTML.match(/fill="#[0-9a-f]{6}"/gi) || []).slice(0, 3).join())).size`);
    check('slides: More Themes opens a gallery of at least eight themes of our own, each card this slide drawn in that theme',
      more === 'clicked' && gallery === true && names.length >= 8 && distinct >= 8, `${more}; ${names.join(', ')}; ${distinct} distinct drawings`);
    await capture('slides-themes-gallery.png');

    // 3. Ember, picked: every slide restyled, the slide's own words untouched.
    const wordsBefore = model(1).slide.shapes.map((s) => (s.text?.paragraphs || []).map((p) => p.plain).join('|')).join('/');
    const picked = await clickIn('.sl-dg-themes .sl-dg-card[data-theme="ember"]');
    const applied = await until(() => info().builtIn === 'ember', 'Ember applied', 6000).catch(() => false);
    const drawn = await until(() => stageHas('Georgia'), 'the stage in Georgia', 8000).catch(() => false);
    const thumbs = await until(async () => (await thumbHas(1, 'Georgia')) && (await thumbHas(2, 'Georgia')), 'the strip redrawn', 10000).catch(() => false);
    const closed = await js(`!document.querySelector('.sl-dg')`);
    const wordsAfter = model(1).slide.shapes.map((s) => (s.text?.paragraphs || []).map((p) => p.plain).join('|')).join('/');
    check('slides: picking Ember applies it to every slide — the stage and every thumbnail redrawn in its fonts, the words as they were',
      picked === 'clicked' && applied === true && drawn === true && thumbs === true && closed && wordsAfter === wordsBefore,
      `${picked}; applied ${applied}; stage ${drawn}; thumbs ${thumbs}; closed ${closed}`);
    await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'slide 2', 4000).catch(() => {});
    await capture('slides-theme-applied.png');

    // 4. The dark variant, from the strip.
    await until(() => js(`document.querySelectorAll('.sl-rs-variants .sl-rs-item').length === 4`), 'the variants', 6000).catch(() => {});
    const variant = await clickIn('.sl-rs-variants .sl-rs-item[data-variant="3"]');
    const dark = await until(() => info().dark === true && info().variant === 3, 'the dark variant', 6000).catch(() => false);
    const lightWords = await until(() => stageHas(`fill="#${String(info().colors?.lt1 || '').toLowerCase()}"`), 'light words on the stage', 8000).catch(() => false);
    check('slides: the fourth variant turns the theme dark — the master\'s colour map flipped, the words drawn in the light colour',
      variant === 'clicked' && dark === true && lightWords === true, `${variant}; ${JSON.stringify({ dark: info().dark, variant: info().variant })}; words ${lightWords}`);
    await capture('slides-theme-dark.png');

    // 5. Colours → a palette.
    const colours = await clickRibbon('Colours');
    await until(() => js(`document.querySelectorAll('.sl-dg-colours .sl-dg-row').length >= 10`), 'the Colours menu', 5000).catch(() => {});
    await capture('slides-colours-menu.png');
    const lagoon = await clickIn('.sl-dg-colours .sl-dg-row[data-palette="lagoon"]');
    const palette = await until(() => info().colorName === 'Lagoon' && info().colors?.accent1 === '0096C7', 'Lagoon', 5000).catch(() => false);
    check('slides: Colours lists the palettes with their swatches, and Lagoon rewrites the theme\'s colours', colours === 'clicked' && lagoon === 'clicked' && palette === true, `${colours}/${lagoon}; ${info().colorName}`);

    // 6. Customise Colours: Accent 1 typed, a name, Save.
    await clickRibbon('Colours');
    await until(() => js(`Boolean(document.querySelector('.sl-dg-colours .sl-dg-custom'))`), 'Customise Colours', 5000).catch(() => {});
    const custom = await clickIn('.sl-dg-colours .sl-dg-custom');
    await until(() => js(`document.querySelectorAll('.sl-cc-slot').length === 12`), 'the twelve slots', 5000).catch(() => {});
    const typed = await setField('.sl-cc-hex-accent1', '#FF6600');
    const named = await setField('.sl-cc-namefield', 'Checks orange');
    await until(() => js(`document.querySelector('.sl-cc-pick-accent1')?.value === '#ff6600'`), 'the picker to follow', 3000).catch(() => {});
    await capture('slides-customise-colours.png');
    const saved = await clickIn('.sl-cc-save');
    const customised = await until(() => info().colorName === 'Checks orange' && info().colors?.accent1 === 'FF6600', 'the custom colours', 5000).catch(() => false);
    check('slides: Customise Colours shows the twelve slots with a sample, and Save writes them under their name',
      custom === 'clicked' && typed === 'set' && named === 'set' && saved === 'clicked' && customised === true, `${custom}/${typed}/${named}/${saved}; ${info().colorName} ${info().colors?.accent1}`);

    // 7. Fonts → a pair, then Customise Fonts.
    const fonts = await clickRibbon('Fonts');
    await until(() => js(`document.querySelectorAll('.sl-dg-fonts .sl-dg-row').length >= 8`), 'the Fonts menu', 5000).catch(() => {});
    await capture('slides-fonts-menu.png');
    const classic = await clickIn('.sl-dg-fonts .sl-dg-row[data-pair="classic"]');
    const pair = await until(() => info().fontName === 'Classic' && info().fonts?.minor === 'Georgia', 'Classic', 5000).catch(() => false);
    await clickRibbon('Fonts');
    await until(() => js(`Boolean(document.querySelector('.sl-dg-fonts .sl-dg-custom'))`), 'Customise Fonts', 5000).catch(() => {});
    await clickIn('.sl-dg-fonts .sl-dg-custom');
    await until(() => js(`Boolean(document.querySelector('.sl-cf-major'))`), 'the fonts dialog', 5000).catch(() => {});
    await setField('.sl-cf-major', 'Verdana');
    await setField('.sl-cf-name', 'Checks fonts');
    await clickIn('.sl-cf-save');
    const ownFonts = await until(() => info().fontName === 'Checks fonts' && info().fonts?.major === 'Verdana' && info().fonts?.minor === 'Georgia', 'the custom fonts', 5000).catch(() => false);
    check('slides: Fonts offers the pairs in their own faces, Classic sets both, and Customise Fonts saves a heading and a body face under a name',
      fonts === 'clicked' && classic === 'clicked' && pair === true && ownFonts === true, `${fonts}/${classic}; ${JSON.stringify(info().fonts)} ${info().fontName}`);

    // 8. Effects, then Undo takes the last step back and Redo puts it again.
    const effects = await clickRibbon('Effects');
    await until(() => js(`document.querySelectorAll('.sl-dg-effects .sl-dg-card svg').length >= 3`), 'the Effects gallery', 5000).catch(() => {});
    await capture('slides-effects-menu.png');
    const outline = await clickIn('.sl-dg-effects .sl-dg-card[data-effects="outline"]');
    const set = await until(() => info().effects === 'outline', 'Outline', 5000).catch(() => false);
    const heavy = (() => { const s = model(1).slide.shapes.find((x) => x.name === 'Highlight'); return s?.line?.width; })();
    await wait(300);
    const undo = await js(`(() => { const b = document.querySelector('.rw-ribbon [data-tip^="Undo"]'); if (!b || b.disabled) return 'no undo'; b.click(); return 'clicked'; })()`);
    const undone = await until(() => info().effects === 'lifted' && info().fontName === 'Checks fonts', 'the effects undone', 5000).catch(() => false);
    await wait(300);
    await js(`(() => { document.querySelector('.sl-stage')?.focus(); return 1; })()`);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Y', modifiers: ['control'] });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Y', modifiers: ['control'] });
    const redone = await until(() => info().effects === 'outline', 'the effects redone', 5000).catch(() => false);
    check('slides: Effects → Outline outlines the theme-styled shape heavily; Undo takes it back to Ember\'s own and Ctrl+Y puts it again',
      effects === 'clicked' && outline === 'clicked' && set === true && heavy >= 2 && undo === 'clicked' && undone === true && redone === true,
      `${effects}/${outline}/${undo}; line ${heavy} pt; undone ${undone}; redone ${redone}`);

    // 9. Saved and opened again: the design is in the file.
    await js(`(() => { document.querySelector('.sl-stage')?.focus(); return 1; })()`);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'S', modifiers: ['control'] });
    const written = await until(() => { try { const d = Deck.open(fs.readFileSync(file)); return d.designInfo(0).effects === 'outline'; } catch { return false; } }, 'the saved file', 8000).catch(() => false);
    const reopened = (() => { try { return Deck.open(fs.readFileSync(file)).designInfo(0); } catch (err) { return { error: err.message }; } })();
    check('slides: saved, the deck opens again as Ember in its dark variant with the custom colours, fonts and Outline effects',
      written === true && reopened.name === 'Ember' && reopened.dark === true && reopened.colorName === 'Checks orange' && reopened.fontName === 'Checks fonts',
      JSON.stringify({ name: reopened.name, dark: reopened.dark, colours: reopened.colorName, fonts: reopened.fontName, effects: reopened.effects, error: reopened.error }));
    check('slides: the design checks raised no errors in the window', consoleErrors.length === 0, consoleErrors.slice(0, 3).join(' | '));
  } catch (err) {
    check('slides: the design checks ran', false, err.stack || err.message);
  }
}
