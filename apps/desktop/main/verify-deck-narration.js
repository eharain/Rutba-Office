// Presentations: the Record tab, with the check run's stand-in microphone —
// From Beginning records the show: about a second on slide 1, Right, about
// a second on slide 2, Stop; each slide gets its narration and its timing;
// Preview plays this slide's narration; the show plays a slide's narration
// as it comes in without showing its speaker; the saved deck keeps the
// narration as PowerPoint does; Clear takes narration and timings away;
// and Record Audio puts a recorded sound on the slide.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

const named = (s) => /^Audio Recording\b/.test(s?.name || '') && s?.media?.kind === 'audio';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyDeckNarration(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  const file = path.join(dir, 'narration.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Talk', slides: [{ layout: 'title', title: 'Welcome' }, { layout: 'obj', title: 'Plan', body: ['One', 'Two'] }] }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
    const press = (tab, label) => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(tab)})?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!b || b.disabled) return 'no button';
      b.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 600, clientY: 110 }));
      return 'clicked';
    })()`);
    const pick = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    const key = (code) => { wc.sendInputEvent({ type: 'keyDown', keyCode: code }); wc.sendInputEvent({ type: 'keyUp', keyCode: code }); };

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    await js(`(() => { window.__plays = []; const play = HTMLMediaElement.prototype.play; HTMLMediaElement.prototype.play = function () { window.__plays.push({ src: this.src, narration: this.dataset.narration || null }); return play.call(this); }; return 1; })()`);

    // From Beginning: a second or so on each slide, then Stop.
    const pressed = await press('Record', 'From Beginning');
    const bar = await until(() => js(`Boolean(document.querySelector('.sl-recbar'))`), 'the recording bar', 6000).then(() => true).catch(() => false);
    await wait(1300);
    key('Right');
    await until(() => js(`document.querySelector('.sl-present-bar')?.textContent.trim().startsWith('2')`), 'slide 2', 4000).catch(() => {});
    await wait(1000);
    await js(`document.querySelector('.sl-recbar-stop')?.click(), 1`);
    const recorded = await until(() => model(0).slide?.shapes?.some(named) && model(1).slide?.shapes?.some(named), 'the narration on both slides', 10000).then(() => true).catch(() => false);
    const t0 = model(0).slide?.transition?.advanceAfter;
    const t1 = model(1).slide?.transition?.advanceAfter;
    check('presentations: Record → From Beginning records the show — each slide gets its narration and the time it was up as its timing',
      pressed === 'clicked' && bar && recorded && t0 > 0.9 && t0 < 3 && t1 > 0.6 && t1 < 3 && model(0).showSettings?.useTimings === true,
      JSON.stringify({ pressed, bar, recorded, t0, t1 }));

    // Preview, on slide 1.
    await js(`(() => { window.__plays = []; document.querySelectorAll('.sl-thumb')[0]?.click(); return 1; })()`);
    await wait(300);
    await press('Record', 'Preview');
    const previewed = await until(() => js(`window.__plays.length >= 1`), 'the preview', 3000).then(() => true).catch(() => false);

    // The show plays slide 1's narration as it comes in; its speaker is not seen.
    await js(`(() => { window.__plays = []; return 1; })()`);
    await press('Slide Show', 'From Beginning');
    const narrated = await until(() => js(`window.__plays.some((p) => p.narration)`), 'the narration in the show', 6000).then(() => true).catch(() => false);
    const hidden = await js(`(() => { const id = ${JSON.stringify(String(model(0).slide.shapes.find(named)?.id))}; const g = document.querySelector('.sl-show-stage [data-shape="' + id + '"]'); return Boolean(g) && getComputedStyle(g).visibility === 'hidden'; })()`);
    key('Escape');
    await until(() => js(`!document.querySelector('.sl-present-bar')`), 'the show to end', 4000).catch(() => {});
    check('presentations: Record → Preview plays this slide\'s narration, and the show plays it as the slide comes in without showing its speaker',
      previewed && narrated && hidden, JSON.stringify({ previewed, narrated, hidden }));

    // Saved: the narration as PowerPoint keeps it.
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const deck = Deck.open(fs.readFileSync(file));
        const xml = deck.pkg.text('ppt/slides/slide1.xml');
        return deck.slide(0).shapes.some(named) && /presetClass="mediacall"/.test(xml) && /<p:cMediaNode vol="80000" showWhenStopped="0">/.test(xml) && /<p:transition[^>]*advTm="\d+"/.test(xml);
      } catch { return false; }
    }, 'the saved narration', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck keeps each narration as PowerPoint does — a hidden speaker played as the slide comes in — with the slide timings', saved, saved ? 'mediacall, cMediaNode and advTm in the slide' : 'not in the file');

    // Clear narration, then timings, on every slide.
    await press('Record', 'Clear Recording');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'Clear Narration on All Slides'))`), 'the menu', 3000).catch(() => {});
    await pick('Clear Narration on All Slides');
    const noNarration = await until(() => !model(0).slide?.shapes?.some(named) && !model(1).slide?.shapes?.some(named), 'the narration cleared', 5000).then(() => true).catch(() => false);
    await press('Record', 'Clear Recording');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'Clear Timings on All Slides'))`), 'the menu', 3000).catch(() => {});
    await pick('Clear Timings on All Slides');
    const noTimings = await until(() => model(0).slide?.transition?.advanceAfter == null && model(1).slide?.transition?.advanceAfter == null, 'the timings cleared', 5000).then(() => true).catch(() => false);
    check('presentations: Clear Recording takes the narration off every slide, and then the timings', noNarration && noTimings, JSON.stringify({ noNarration, noTimings }));

    // Record Audio: a sound recorded and put on the slide.
    await press('Record', 'Audio');
    await until(() => js(`Boolean(document.querySelector('.sl-recaudio-record'))`), 'the Record Sound box', 4000).catch(() => {});
    await js(`document.querySelector('.sl-recaudio-record')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.sl-recaudio-stop'))`), 'recording', 5000).catch(() => {});
    await wait(900);
    await js(`document.querySelector('.sl-recaudio-stop')?.click(), 1`);
    await until(() => js(`!document.querySelector('.sl-recaudio-ok')?.disabled`), 'the sound', 5000).catch(() => {});
    await js(`document.querySelector('.sl-recaudio-ok')?.click(), 1`);
    const sound = await until(() => model(0).slide?.shapes?.some((s) => s.name === 'Recorded Sound' && s.media?.kind === 'audio'), 'the recorded sound', 6000).then(() => true).catch(() => false);
    check('presentations: Record → Audio records a sound and puts it on the slide', sound, sound ? 'Recorded Sound on slide 1' : 'not on the slide');
  } catch (err) {
    check('presentations: the narration checks ran', false, err.message);
  }
}
