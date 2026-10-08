// Presentations: Design → Design Ideas.
//
// On a slide of a title and words, Design Ideas opens a pane of layouts the
// slide suits, each drawn as the slide would be; a click on the title band
// applies it — the band added behind the words in the theme's accent, the
// title in white on it — and the pane draws its ideas again for the slide
// as it is now. Run alone with RUTBA_VERIFY_ONLY=designideas.
import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';

export async function verifyDesignIdeas({ open, check, until, wait, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'design-ideas.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ slides: [{ title: 'Quarterly results', body: ['Sales rose 12%', 'Costs fell by a third'] }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('deck').id });
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Design'))`), 'the Design tab', 8000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Design').click(), 'tab'`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Design Ideas'))`), 'Design → Design Ideas', 4000);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Design Ideas'); b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-idea svg').length >= 3`), 'the ideas drawn', 8000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('.sl-idea')].map((b) => b.dataset.idea)`);
    check('slides: Design Ideas opens a pane of layouts the slide suits, each drawn as the slide would be', offered.includes('title-band') && offered.includes('side-bar') && (await js(`document.querySelectorAll('.sl-idea svg').length`)) === offered.length, offered.join(', '));
    await js(`document.querySelector('.sl-idea[data-idea="title-band"]')?.click(), 'applied'`);
    const band = () => (model().slide?.shapes || []).find((s) => s.name === 'Design idea');
    await until(() => Boolean(band()), 'the band on the slide', 6000).catch(() => {});
    const title = (model().slide?.shapes || []).find((s) => s.placeholder?.type === 'title');
    const white = (title?.text?.paragraphs?.[0]?.runs || []).every((r) => /^#?FFFFFF$/i.test(r.color || ''));
    check('slides: a click applies the idea — the band behind the words, the title in white on it', Boolean(band()) && (model().slide?.shapes || [])[0]?.name === 'Design idea' && white, JSON.stringify({ band: band()?.geometry, white }));
    await until(() => js(`document.querySelectorAll('.sl-idea svg').length >= 3`), 'the ideas drawn again', 8000).catch(() => {});
    const complaints = await errorsIn(win);
    check('slides: Design Ideas reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the Design Ideas checks ran', false, err.message);
  }
}
