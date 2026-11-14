// Video: smaller exports, Compact, Small and Fit under a size (the 2.23
// list, item 1).
//
// A short project, twelve seconds of detailed 1080p with a tone under it
// (blocks of colour that pan as a film's picture does, so the encoder has to
// spend what it is given), is exported through the Export box at each new preset and with a
// 5 MB fit, every file read back from the disk by our own reader: its size,
// its picture's frame and how many bits the picture took. A fit whose first
// pass is made to come out over proves the second pass on the real encoder.
// The box itself: a size on every card, the project's with its Quality; the
// aim worked out from what is typed; the warning at a size too thin for
// 1080p, and the 720 it suggests. Run alone with RUTBA_VERIFY_ONLY=video-fit.

import fs from 'node:fs';
import path from 'node:path';
import { readMp4 } from '@rutba/media/mp4-read';
import { makeHdClip } from './verify-video-mp4.js';
import { labelKeyed } from './verify-keys.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyVideoFit(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const capDir = process.env.RUTBA_VERIFY_CAPTURE || null;
  const capture = async (win, name) => {
    if (!capDir) return;
    try {
      fs.mkdirSync(capDir, { recursive: true });
      win.webContents.invalidate();
      await wait(700);
      const img = await win.webContents.capturePage();
      fs.writeFileSync(path.join(capDir, name), img.toPNG());
    } catch {
      /* a capture is a picture to look at, not a check */
    }
  };
  const root = path.join(dir, 'video-fit');
  fs.mkdirSync(root, { recursive: true });
  if (Array.isArray(globalThis.__rutbaCheckDialogAnswers)) globalThis.__rutbaCheckDialogAnswers.length = 0;
  const win = await open('video');
  const js = (code) => win.webContents.executeJavaScript(code);

  const clip = await makeHdClip(win.webContents, root, { seconds: 12, busy: true, bitrate: 24_000_000, name: 'detailed-1080p.mp4' });
  if (!clip.file) {
    check('video fit: a busy 1080p clip could be made for the checks', false, clip.error || 'no clip');
    return;
  }

  const clickTab = (name) => js(`(() => { const t = ((list) => list.find((n) => ${labelKeyed(name)}(n)) || list.find((n) => n.textContent.trim() === ${JSON.stringify(name)}))([...document.querySelectorAll('.rw-tab')]); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
  const press = async (label) => {
    const find = `((list) => list.find((n) => ${labelKeyed(label, { tip: true, start: true })}(n) && !n.disabled) || list.find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled))([...document.querySelectorAll('.rw-ribbon .rw-btn, .rw-dialog .rw-btn')])`;
    await until(() => js(`Boolean(${find})`), `the ${label} button`, 6000).catch(() => {});
    return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(label)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
  };
  const duration = () => js(`Number(document.querySelector('.vd-editor')?.dataset.duration || 0)`);
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [clip.file] });
  const added = await press('Add Media');
  await until(async () => (await js(`document.querySelectorAll('.vd-clip[data-kind="video"]').length`)) === 1 && (await duration()) > 11, 'the clip on the timeline', 30000).catch(() => {});
  const length = await duration();
  check('video fit: a short project of twelve seconds of detailed 1080p is on the timeline', length > 11.5 && length < 12.5, `${added}; ${length.toFixed(2)} s`);

  const openBox = async () => {
    await clickTab('Home');
    await press('Export');
    await until(() => js(`Boolean(document.querySelector('.vd-export-go'))`), 'the Export box', 6000).catch(() => {});
    await until(() => js(`document.querySelector('.vd-export-format')?.dataset.fast === '1'`), 'our own MP4', 8000).catch(() => {});
  };
  const choose = async (id) => {
    await js(`document.querySelector('.vd-export-card[data-preset="${id}"]')?.click(), 1`);
    await until(() => js(`document.querySelector('.vd-export-preset')?.dataset.value === '${id}'`), `the ${id} preset`, 3000).catch(() => {});
  };
  const setFit = async (on) => {
    await js(`(() => { const c = document.querySelector('.vd-export-fit-box'); if (c && c.checked !== ${on}) c.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.vd-export-chips')) === ${on}`), 'the size choices', 3000).catch(() => {});
  };
  const typeSize = async (mb) => {
    await js(`(() => {
      const n = document.querySelector('.vd-export-size');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(String(mb))});
      n.dispatchEvent(new Event('input', { bubbles: true }));
      return 1;
    })()`);
    await until(() => js(`document.querySelector('.vd-export-size')?.value === ${JSON.stringify(String(mb))}`), 'the size typed', 3000).catch(() => {});
    await wait(120);
  };
  /** Press Export with the box as it stands, wait for the file, and read it back. */
  const exportTo = async (name) => {
    const target = path.join(root, name);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: target });
    await until(() => js(`document.querySelector('.vd-export-go')?.disabled === false`), 'Export to be pressable', 6000).catch(() => {});
    await js(`document.querySelector('.vd-export-go')?.click(), 1`);
    const made = await until(() => fs.existsSync(target), `the file ${name}`, 120000).then(() => true).catch(() => false);
    await until(() => js(`!document.querySelector('.vd-export-busy')`), 'the progress to close', 15000).catch(() => {});
    await wait(300);
    if (!made) return { name, missing: true };
    const size = fs.statSync(target).size;
    let file = null;
    try { file = readMp4(new Uint8Array(fs.readFileSync(target))); } catch (err) { file = { error: err.message }; }
    const v = file?.tracks?.find((t) => t.kind === 'video');
    const a = file?.tracks?.find((t) => t.kind === 'audio');
    const pictureBytes = v ? v.samples.reduce((n, s) => n + s.size, 0) : 0;
    const said = await js(`document.querySelector('.vd-editor')?.dataset.export ? JSON.parse(document.querySelector('.vd-editor').dataset.export) : null`);
    return { name, size, width: v?.width, height: v?.height, duration: v?.duration, audio: Boolean(a), rate: v ? Math.round((pictureBytes * 8) / v.duration) : 0, said, error: file?.error };
  };
  const say = (f) => (f.missing ? `${f.name}: no file` : `${f.name}: ${f.size} bytes on disk, ${f.width}x${f.height}, ${f.duration?.toFixed(2)} s, picture ${(f.rate / 1e6).toFixed(2)} Mb/s, sound ${f.audio}; the app said ${f.said?.size} bytes${f.said?.fit ? `, ${f.said.fit.tries.length} go(es): ${f.said.fit.tries.map((x) => `${(x.video / 1e6).toFixed(2)} Mb/s -> ${x.size}`).join(', ')}` : ''}`);
  const readBack = (f) => !f.missing && !f.error && f.size === f.said?.size;

  // The box: a size on every card.
  await openBox();
  await choose('project');
  const cards = await js(`[...document.querySelectorAll('.vd-export-card')].map((c) => ({ id: c.dataset.preset, name: c.querySelector('.vd-export-card-name')?.textContent.trim(), bytes: Number(c.querySelector('.vd-export-card-size')?.dataset.bytes || 0), size: c.querySelector('.vd-export-card-size')?.textContent.trim(), detail: c.querySelector('.vd-export-card-detail')?.textContent.trim() }))`);
  const byId = Object.fromEntries(cards.map((c) => [c.id, c]));
  check('video fit: the Made for box offers every destination with a size for this project, the project\'s own included',
    cards.length === 8 && cards.every((c) => c.bytes > 0 && /^About /.test(c.size)) && /1920 × 1080/.test(byId.compact?.detail || '') && /1280 × 720/.test(byId.small?.detail || ''),
    JSON.stringify(cards.map((c) => `${c.id} ${c.size}: ${c.detail}`)));
  const high = byId.project?.bytes || 0;
  await js(`(() => { const s = document.querySelector('.vd-export-quality'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'low'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
  await until(() => js(`Number(document.querySelector('.vd-export-card[data-preset="project"] .vd-export-card-size').dataset.bytes) < ${high}`), 'the size to follow the Quality', 3000).catch(() => {});
  const low = await js(`Number(document.querySelector('.vd-export-card[data-preset="project"] .vd-export-card-size').dataset.bytes)`);
  check('video fit: the project\'s size follows its Quality', low > 0 && low < high * 0.5, `High ${high}, Low ${low}`);
  await capture(win, 'video-fit-1-box-project.png');
  await js(`(() => { const s = document.querySelector('.vd-export-quality'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'high'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);

  // Compact and Small, written and read back.
  await choose('compact');
  await capture(win, 'video-fit-2-box-compact.png');
  const compact = await exportTo('compact.mp4');
  check('video fit: Compact writes 1920 by 1080 with the picture at about 4 Mb/s, the file read back',
    readBack(compact) && compact.width === 1920 && compact.height === 1080 && Math.abs(compact.duration - length) < 0.2 && compact.rate > 3_200_000 && compact.rate < 4_800_000 && compact.audio && compact.size < 8_000_000, say(compact));
  await openBox();
  await choose('small');
  const small = await exportTo('small.mp4');
  check('video fit: Small writes 1280 by 720 with the picture at about 2 Mb/s, smaller than Compact, the file read back',
    readBack(small) && small.width === 1280 && small.height === 720 && small.rate > 1_500_000 && small.rate < 2_500_000 && small.size < compact.size, say(small));

  // Fit: the aim worked out from what is typed.
  await openBox();
  await choose('project');
  await setFit(true);
  await typeSize(5);
  await until(() => js(`/4\\.6 MB/.test(document.querySelector('.vd-export-aim')?.textContent || '')`), 'the aim', 3000).catch(() => {});
  const aim = await js(`({ text: document.querySelector('.vd-export-aim')?.textContent.trim(), capped: document.querySelector('.vd-export-aim')?.dataset.capped, video: Number(document.querySelector('.vd-export-aim')?.dataset.video || 0), pressed: document.querySelectorAll('.vd-export-chip[aria-pressed="true"]').length })`);
  check('video fit: a size typed (5 MB) says what it aims for and at what rate, and takes the chips off', aim.capped === '0' && /about 4\.6 MB/.test(aim.text || '') && /under 5 MB/.test(aim.text || '') && aim.video > 2_000_000 && aim.video < 3_400_000 && aim.pressed === 0, JSON.stringify(aim));
  await capture(win, 'video-fit-3-typed.png');
  const fit5 = await exportTo('fit-5mb.mp4');
  check('video fit: a 5 MB fit comes out under 5 MB, read from the file written', readBack(fit5) && fit5.size <= 5_000_000 && fit5.size > 3_500_000 && fit5.width === 1920 && fit5.height === 1080 && fit5.said?.fit?.over === false, say(fit5));

  // Fit: the first pass made to come out over, the second pass real.
  await js(`globalThis.__rutbaCheckFitOver = 3, 1`);
  await openBox();
  await setFit(true);
  await typeSize(5);
  const toastsBefore = await js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent)`);
  const again = await exportTo('fit-again.mp4');
  await js(`delete globalThis.__rutbaCheckFitOver, 1`);
  const toast = await until(() => js(`[...document.querySelectorAll('.rw-toast')].some((n) => /encoded again/.test(n.textContent))`), 'the toast to say so', 6000).then(() => js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).join(' | ')`)).catch(() => '');
  const passes = again.said?.fit?.tries || [];
  check('video fit: a first pass that comes out over is encoded again lower, on the real encoder, and the file is under',
    readBack(again) && passes.length >= 2 && passes[0].size > 5_000_000 && passes[passes.length - 1].size <= 5_000_000 && passes[1].video < passes[0].video && again.size <= 5_000_000 && again.said?.fit?.over === false, say(again));
  check('video fit: the person is told it was made again', /encoded again/.test(toast), `${toast}; before ${JSON.stringify(toastsBefore)}`);

  // Too thin for 1080p: the warning, the suggestion, and 720 taken.
  await openBox();
  await setFit(true);
  await typeSize(1);
  await until(() => js(`Boolean(document.querySelector('.vd-export-use-short'))`), 'the warning', 3000).catch(() => {});
  const warn = await js(`({ text: document.querySelector('.vd-export-too-low')?.textContent.trim(), rate: Number(document.querySelector('.vd-export-aim')?.dataset.video || 0), link: document.querySelector('.vd-export-use-short')?.textContent.trim() })`);
  check('video fit: a size that leaves under about 1 Mb/s at 1080p warns and suggests 720', warn.rate < 1_000_000 && warn.rate > 100_000 && /too little for a clear picture at 1080/.test(warn.text || '') && warn.link === 'Use 720 instead', JSON.stringify(warn));
  await capture(win, 'video-fit-4-too-thin.png');
  await js(`document.querySelector('.vd-export-use-short')?.click(), 1`);
  await until(() => js(`/1280 × 720/.test(document.querySelector('.vd-export-card[aria-checked="true"] .vd-export-card-detail')?.textContent || '')`), 'the card at 720', 3000).catch(() => {});
  const lowered = await js(`({ detail: document.querySelector('.vd-export-card[aria-checked="true"] .vd-export-card-detail')?.textContent.trim(), warned: Boolean(document.querySelector('.vd-export-use-short')) })`);
  check('video fit: Use 720 instead takes the frame down and the warning away', /1280 × 720/.test(lowered.detail || '') && lowered.warned === false, JSON.stringify(lowered));
  await capture(win, 'video-fit-5-at-720.png');

  // A size with no room for a picture stops the Export button.
  await typeSize(0.05);
  await until(() => js(`Boolean(document.querySelector('.vd-export-too-small'))`), 'the size refused', 3000).catch(() => {});
  const refused = await js(`({ text: document.querySelector('.vd-export-too-small')?.textContent.trim(), off: document.querySelector('.vd-export-go')?.disabled })`);
  check('video fit: a size that leaves nothing for the picture says so and Export waits', refused.off === true && /no room for the picture/.test(refused.text || ''), JSON.stringify(refused));
  // A chip again: this short project is already under 50 MB, and the box says so.
  await js(`document.querySelector('.vd-export-chip[data-mb="50"]')?.click(), 1`);
  await until(() => js(`document.querySelector('.vd-export-aim')?.dataset.capped === '1'`), 'the 50 MB chip', 3000).catch(() => {});
  const capped = await js(`({ capped: document.querySelector('.vd-export-aim')?.dataset.capped, text: document.querySelector('.vd-export-aim')?.textContent.trim(), value: document.querySelector('.vd-export-size')?.value })`);
  check('video fit: a chip clears the typed size; a size the project is already under says it stays as it is', capped.capped === '1' && /stays as it is/.test(capped.text || '') && capped.value === '', JSON.stringify(capped));
  await capture(win, 'video-fit-6-capped.png');
  await press('Cancel');
  await until(() => js(`!document.querySelector('.vd-export-go')`), 'the box to close', 3000).catch(() => {});

  // The choice stays the next time, as it always has for Made for.
  await js(`window.rutbaOffice.store.set({ key: 'video.exportPreset', value: null }).then(() => 1)`);
  const complaints = await errorsIn(win);
  check('video fit: the window reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
}
