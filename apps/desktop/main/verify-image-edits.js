// Image: the quick edits, pressed in the window and read back from the
// canvas and from the files exported. A picture of noise over a gradient
// (so a JPEG of it is large enough for quality and a size cap to show) is
// flipped upside down, resized to half, lightened by a curve, marked with a
// rectangle, straightened by five degrees, and exported as JPEG at two
// qualities, under a size cap and with its longest side capped as a PNG;
// each file's size and dimensions are read back. Copy puts the result on the
// clipboard; Paste opens a picture from it; Red eye records a spot; Auto
// enhance moves the levels. Run alone with RUTBA_VERIFY_ONLY=image-edits.
import fs from 'node:fs';
import path from 'node:path';
import { clipboard, nativeImage } from 'electron';
import { probeImage } from '@rutba/imaging/probe';
import { straightenRect } from '@rutba/imaging/pipeline';

/** A w × h picture: a diagonal gradient with grain, as BGRA for nativeImage. */
export function noisyPicture(width, height, seed = 11) {
  const buf = Buffer.alloc(width * height * 4);
  let s = seed;
  const rand = () => ((s = (s * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const t = (x / width + y / height) / 2;
      const n = (rand() - 0.5) * 70;
      const i = (y * width + x) * 4;
      // Blue at the top left, orange at the bottom right; the bottom left a green.
      const r = 30 + 200 * t + n;
      const g = 60 + 120 * (y / height) + n;
      const b = 200 - 170 * t + n;
      buf[i] = Math.max(0, Math.min(255, b));
      buf[i + 1] = Math.max(0, Math.min(255, g));
      buf[i + 2] = Math.max(0, Math.min(255, r));
      buf[i + 3] = 255;
    }
  }
  return nativeImage.createFromBitmap(buf, { width, height });
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyImageEdits({ open, check, until, wait, press, errorsIn }, { dir }) {
  const root = path.join(dir, 'image-edits');
  const keptText = clipboard.readText();
  try {
    fs.mkdirSync(root, { recursive: true });
    const file = path.join(root, 'scene.png');
    fs.writeFileSync(file, noisyPicture(1200, 800).toPNG());
    const win = await open('image', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    const result = () => js(`document.querySelector('.im-canvas-wrap')?.dataset.result || ''`);
    const drawnTurn = () => js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`);
    const ops = () => js(`[...document.querySelectorAll('.im-op')].map((n) => n.dataset.op)`);
    const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
    const click = async (label, scope = '.rw-groups') => {
      const find = `[...document.querySelectorAll(${JSON.stringify(`${scope} .rw-btn`)})].find((n) => (n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${label} button`, 4000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(label)}; b.click(); return 'clicked'; })()`);
    };
    const setRange = (selector, value) => js(`(() => { const i = document.querySelector(${JSON.stringify(selector)}); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(String(value))}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const setSelect = (selector, value) => js(`(() => { const s = document.querySelector(${JSON.stringify(selector)}); if (!s) return false; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, ${JSON.stringify(String(value))}); s.dispatchEvent(new Event('change', { bubbles: true })); return true; })()`);
    // A pixel of the preview, at a fraction of its width and height.
    const pixelAt = (fx, fy) => js(`(() => { const c = document.querySelector('.im-canvas'); const d = c.getContext('2d').getImageData(Math.floor(c.width * ${fx}), Math.floor(c.height * ${fy}), 1, 1).data; return [d[0], d[1], d[2]]; })()`);
    // The average of a 9 × 9 patch, so the grain does not decide a comparison.
    const patchAt = (fx, fy) => js(`(() => { const c = document.querySelector('.im-canvas'); const x = Math.max(0, Math.min(c.width - 9, Math.floor(c.width * ${fx}) - 4)); const y = Math.max(0, Math.min(c.height - 9, Math.floor(c.height * ${fy}) - 4)); const d = c.getContext('2d').getImageData(x, y, 9, 9).data; const s = [0, 0, 0]; for (let i = 0; i < d.length; i += 4) { s[0] += d[i]; s[1] += d[i + 1]; s[2] += d[i + 2]; } return s.map((v) => Math.round(v / 81)); })()`);
    const settled = async (what, timeout = 8000) => {
      // The preview is drawn a turn after the plan changes: wait for the turn to move, then to rest.
      const before = await drawnTurn();
      await until(async () => (await drawnTurn()) > before, what, timeout).catch(() => {});
      let last = await drawnTurn();
      for (let i = 0; i < 10; i++) {
        await wait(150);
        const now = await drawnTurn();
        if (now === last) break;
        last = now;
      }
    };

    await until(async () => (await result()) === '1200x800' && (await drawnTurn()) > 0, 'the picture on the canvas', 15000);

    // I1: Flip vertical. The top-left corner becomes what was the bottom-left.
    const topLeft = await patchAt(0.03, 0.03);
    const bottomLeft = await patchAt(0.03, 0.97);
    const flipClick = await click('Flip vertical');
    await until(async () => (await ops()).includes('flip'), 'the flip in the edit list', 4000).catch(() => {});
    await settled('the flipped picture');
    const flippedTop = await patchAt(0.03, 0.03);
    const near = (a, b, tol = 14) => a.every((v, i) => Math.abs(v - b[i]) <= tol);
    check('image: Flip vertical turns the picture upside down, the bottom-left corner now at the top-left, and records an edit',
      flipClick === 'clicked' && near(flippedTop, bottomLeft) && !near(flippedTop, topLeft, 30),
      JSON.stringify({ flipClick, topLeft, bottomLeft, flippedTop, ops: await ops() }));

    // I1: Resize to 50 percent, the shape kept.
    await click('Resize');
    await until(() => js(`Boolean(document.querySelector('[data-role="resize-width"]'))`), 'the Resize dialog', 4000).catch(() => {});
    await js(`(() => { const i = document.querySelector('[data-role="resize-width"]'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, '50'); i.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(150);
    const dialogSays = await js(`document.querySelector('.im-resize .rw-hint')?.textContent || ''`);
    await capture('image-edits-resize.png');
    await js(`document.querySelector('[data-role="resize-ok"]')?.click(), 1`);
    await until(async () => (await result()) === '600x400', 'the picture at half size', 5000).catch(() => {});
    check('image: Resize at 50 percent with the shape kept makes 1200 × 800 into 600 × 400', (await result()) === '600x400', `${await result()}; the dialog said "${dialogSays}"`);

    // I3: a curve, the Lighter preset, lifts the middle tones.
    await tab('Adjust');
    await click('Curves');
    await until(() => js(`Boolean(document.querySelector('[data-role="curve-preset"]'))`), 'the curve', 4000).catch(() => {});
    const before = await patchAt(0.5, 0.5);
    await setSelect('[data-role="curve-preset"]', 'lighter');
    await settled('the lighter picture', 10000);
    const after = await patchAt(0.5, 0.5);
    const lum = (p) => 0.2126 * p[0] + 0.7152 * p[1] + 0.0722 * p[2];
    const curvePath = await js(`document.querySelector('.im-curve-line')?.getAttribute('d')?.length || 0`);
    await capture('image-edits-curve.png');
    check('image: Curves, Lighter: the middle of the picture lightens and the curve is drawn', lum(after) > lum(before) + 12 && curvePath > 100, JSON.stringify({ before, after, curvePath }));

    // Auto enhance moves the levels from where they were.
    const levelsBefore = await js(`document.querySelector('[data-level="inBlack"] input')?.value ?? null`);
    await click('Auto enhance');
    await settled('the enhanced picture', 10000);
    await js(`[...document.querySelectorAll('.im-sec-head')].find((b) => b.textContent.includes('Levels'))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-level="inBlack"] input'))`), 'the levels', 3000).catch(() => {});
    const levelsAfter = await js(`[...document.querySelectorAll('[data-level] input')].map((i) => i.value)`);
    check('image: Auto enhance sets levels from the picture, shown on the Levels sliders', Array.isArray(levelsAfter) && levelsAfter.length === 5 && (Number(levelsAfter[0]) > 0 || Number(levelsAfter[2]) < 255), JSON.stringify({ levelsBefore, levelsAfter }));

    // I1: a rectangle drawn with the mouse, editable, then kept.
    await tab('Draw');
    await js(`document.querySelector('[data-tool="rect"]')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.im-draw'))`), 'the drawing layer', 4000).catch(() => {});
    const drag = await js(`(() => {
      const layer = document.querySelector('.im-draw');
      const r = layer.getBoundingClientRect();
      const at = (fx, fy) => ({ clientX: r.left + r.width * fx, clientY: r.top + r.height * fy, bubbles: true, button: 0 });
      layer.dispatchEvent(new MouseEvent('mousedown', at(0.2, 0.25)));
      window.dispatchEvent(new MouseEvent('mousemove', at(0.4, 0.45)));
      window.dispatchEvent(new MouseEvent('mousemove', at(0.6, 0.65)));
      window.dispatchEvent(new MouseEvent('mouseup', at(0.6, 0.65)));
      return { w: r.width, h: r.height };
    })()`);
    const handles = await until(() => js(`document.querySelectorAll('.im-handle').length === 4`), 'the rectangle being edited', 4000).then(() => true).catch(() => false);
    await settled('the rectangle drawn');
    await capture('image-edits-draw.png');
    // Drawn in red: the left edge of the rectangle.
    const edge = await js(`(() => {
      const c = document.querySelector('.im-canvas');
      const x0 = Math.floor(c.width * 0.2) - 6; const y = Math.floor(c.height * 0.45);
      const d = c.getContext('2d').getImageData(x0, y, 12, 1).data;
      for (let i = 0; i < d.length; i += 4) if (d[i] > 180 && d[i + 1] < 110 && d[i + 2] < 110) return [d[i], d[i + 1], d[i + 2]];
      return null;
    })()`);
    await click('Done');
    await until(async () => (await ops()).includes('annotate'), 'the rectangle in the edit list', 4000).catch(() => {});
    const kept = await ops();
    check('image: a rectangle dragged with the Rectangle tool is drawn in red with four handles to reshape it, and Done keeps it as an edit',
      handles && Boolean(edge) && kept.includes('annotate') && !(await js(`Boolean(document.querySelector('.im-handle'))`)),
      JSON.stringify({ drag, handles, edge, ops: kept }));

    // I2: straighten by five degrees, the crop following.
    await tab('Home');
    await until(() => js(`Boolean(document.querySelector('[data-role="straighten"]'))`), 'the straighten slider', 4000).catch(() => {});
    await setRange('[data-role="straighten"]', 5);
    const keep = straightenRect(1200, 800, 5);
    const want = `${Math.round(keep.width * 0.5)}x${Math.round(keep.height * 0.5)}`;
    await until(async () => (await result()) === want, 'the straightened size', 5000).catch(() => {});
    await settled('the straightened picture');
    await capture('image-edits-straighten.png');
    const corners = await Promise.all([[0.005, 0.005], [0.995, 0.005], [0.005, 0.995], [0.995, 0.995]].map(([x, y]) => js(`(() => { const c = document.querySelector('.im-canvas'); return c.getContext('2d').getImageData(Math.floor(c.width * ${x}), Math.floor(c.height * ${y}), 1, 1).data[3]; })()`)));
    check('image: Straighten 5 degrees keeps the largest upright rectangle inside the turned picture, with no empty corner',
      (await result()) === want && corners.every((a) => a === 255),
      JSON.stringify({ result: await result(), want, alphaAtCorners: corners, ops: await ops() }));

    // Copy: the result on the clipboard, at its size.
    clipboard.clear();
    await click('Copy');
    await until(() => !clipboard.readImage().isEmpty(), 'the picture on the clipboard', 8000).catch(() => {});
    const copied = clipboard.readImage().getSize();
    check('image: Copy puts the edited picture on the clipboard at its size', `${copied.width}x${copied.height}` === want, JSON.stringify(copied));

    // Export: JPEG at 90 and at 30; then under 100 KB; then a PNG with its longest side 320.
    await tab('Export');
    await until(() => js(`Boolean(document.querySelector('[data-format="jpeg"]'))`), 'the Export tab', 4000).catch(() => {});
    await js(`document.querySelector('[data-format="jpeg"]')?.click(), 1`);
    await setRange('[data-role="quality"]', 90);
    const exportTo = async (name) => {
      const target = path.join(root, name);
      globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: target });
      await click('Export…');
      await until(() => fs.existsSync(target) && fs.statSync(target).size > 0, `${name} written`, 20000).catch(() => {});
      await wait(200);
      if (!fs.existsSync(target)) return { name, size: 0 };
      const bytes = fs.readFileSync(target);
      const info = probeImage(bytes) || {};
      return { name, size: bytes.length, width: info.width, height: info.height, jpeg: bytes[0] === 0xff && bytes[1] === 0xd8, png: bytes[1] === 0x50 };
    };
    const q90 = await exportTo('scene-q90.jpg');
    await setRange('[data-role="quality"]', 30);
    await wait(150);
    const q30 = await exportTo('scene-q30.jpg');
    await capture('image-edits-export.png');
    const [ww, wh] = want.split('x').map(Number);
    check('image: Export as JPEG writes the edited picture at its size, and quality 30 makes a smaller file than 90',
      q90.jpeg && q30.jpeg && q90.width === ww && q90.height === wh && q30.width === ww && q30.size < q90.size * 0.75,
      JSON.stringify({ q90, q30 }));
    await setRange('[data-role="quality"]', 95);
    await wait(150);
    const q95 = await exportTo('scene-q95.jpg');
    await setSelect('[data-role="max-size"]', 100);
    await wait(150);
    const capped = await exportTo('scene-capped.jpg');
    check('image: a size cap of 100 KB brings a JPEG that is larger at quality 95 under it, at the same dimensions',
      q95.size > 100 * 1024 && capped.jpeg && capped.size > 0 && capped.size <= 100 * 1024 && capped.width === ww,
      JSON.stringify({ q95, capped }));
    await setSelect('[data-role="max-size"]', 0);
    await js(`document.querySelector('[data-format="png"]')?.click(), 1`);
    await setSelect('[data-role="long-side"]', 320);
    await wait(150);
    const small = await exportTo('scene-small.png');
    check('image: Export as PNG with the longest side at 320 writes a PNG 320 wide, the shape kept', small.png && small.width === 320 && Math.abs(small.height - Math.round((320 * wh) / ww)) <= 1, JSON.stringify(small));

    // I4: Red eye, one click on the picture, recorded as an edit.
    await tab('Filters');
    await click('Red eye');
    await until(() => js(`Boolean(document.querySelector('.im-draw.tool-redeye'))`), 'the red-eye tool', 4000).catch(() => {});
    await js(`(() => { const layer = document.querySelector('.im-draw'); const r = layer.getBoundingClientRect(); layer.dispatchEvent(new MouseEvent('mousedown', { clientX: r.left + r.width / 2, clientY: r.top + r.height / 2, bubbles: true, button: 0 })); return 1; })()`);
    await until(async () => (await ops()).includes('redeye'), 'the red-eye spot', 4000).catch(() => {});
    await setRange('[data-slider="sharpen"] input', 40);
    await settled('the sharpened picture', 10000);
    await capture('image-edits-filters.png');
    check('image: Red eye records the spot clicked as an edit, and Sharpen is on the Filters tab', (await ops()).includes('redeye'), JSON.stringify(await ops()));

    // Undo takes the last edit back.
    const count = (await ops()).length;
    await js(`document.body.focus(), 1`);
    await press(wc, 'z', { modifiers: ['control'] });
    await until(async () => (await ops()).length === count - 1, 'one edit undone', 3000).catch(() => {});
    check('image: Ctrl+Z takes the last edit back', (await ops()).length === count - 1, `${count} → ${(await ops()).length}`);

    const complaints = await errorsIn(win);
    check('image: the quick edits report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    // I5: Paste into an empty Image window opens the clipboard's picture.
    clipboard.writeImage(noisyPicture(90, 60, 3));
    const empty = await open('image');
    const ejs = (code) => empty.webContents.executeJavaScript(code);
    await until(() => ejs(`Boolean(document.querySelector('.rw-empty'))`), 'the empty window', 6000).catch(() => {});
    await ejs(`[...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Paste')?.click(), 1`);
    await until(() => ejs(`document.querySelector('.im-canvas-wrap')?.dataset.result === '90x60'`), 'the pasted picture', 8000).catch(() => {});
    const pasted = await ejs(`({ result: document.querySelector('.im-canvas-wrap')?.dataset.result || '', title: document.title })`);
    check('image: Paste opens the clipboard\'s picture as a new picture', pasted.result === '90x60', JSON.stringify(pasted));
  } catch (err) {
    check('image: the quick-edit checks ran', false, err?.message || String(err));
  } finally {
    clipboard.writeText(keptText || '');
  }
}
