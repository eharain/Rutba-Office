// Pictures: Batch edit over a folder of three pictures made here (one a JPEG
// shot on its side): resized to fit, converted, renamed by a pattern, turned
// upright and stripped; run again, the clash is asked about and Keep both
// writes beside; run into a chosen folder with each format and the metadata
// kept. Part of RUTBA_VERIFY_ONLY=image-edits.
import fs from 'node:fs';
import path from 'node:path';
import { probeImage } from '@rutba/imaging/probe';
import { readExif, withOrientation } from '@rutba/imaging/exif';
import { noisyPicture } from './verify-image-edits.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the folders are made
 */
export async function verifyImageBatch({ open, check, until, wait, errorsIn }, { dir }) {
  const root = path.join(dir, 'image-batch');
  const elsewhere = path.join(dir, 'image-batch-out');
  try {
    fs.mkdirSync(root, { recursive: true });
    const a = path.join(root, 'alpha.png');
    const b = path.join(root, 'bravo.jpg');
    const c = path.join(root, 'charlie.png');
    fs.writeFileSync(a, noisyPicture(640, 400, 5).toPNG());
    // Stored 600 × 400, shot on its side: EXIF orientation 6 says turn it a quarter clockwise.
    fs.writeFileSync(b, Buffer.from(withOrientation(new Uint8Array(noisyPicture(600, 400, 7).toJPEG(90)), 6)));
    fs.writeFileSync(c, noisyPicture(300, 300, 9).toPNG());
    const originals = [a, b, c].map((f) => fs.readFileSync(f));
    const win = await open('pictures', root);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    await until(() => js(`document.querySelectorAll('.pv-tile:not(.folder)').length === 3`), 'the folder of three', 8000).catch(() => {});

    const setValue = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return false;
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(String(value))});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    })()`);
    const closeDialog = async () => {
      await js(`document.querySelector('[data-role="batch-close"]')?.click(), 1`);
      await until(() => js(`!document.querySelector('.im-batch')`), 'the dialog closed', 4000).catch(() => {});
    };
    const runBatch = async ({ box, pattern, format = 'jpeg', strip = true, folder = null, answer = null, captureAs = null }) => {
      await js(`[...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Batch edit') && !n.disabled)?.click(), 1`);
      await until(() => js(`Boolean(document.querySelector('[data-role="batch-run"]'))`), 'the Batch edit dialog', 5000);
      // The three originals only: the folder now holds what earlier runs wrote, and all are ticked to begin with.
      await js(`[...document.querySelectorAll('.im-batch-head .rw-btn')].find((n) => n.textContent.trim() === 'None')?.click(), 1`);
      await wait(100);
      await js(`['alpha.png', 'bravo.jpg', 'charlie.png'].forEach((name) => document.querySelector('.im-batch-file[data-file="' + name + '"] input')?.click()), 1`);
      await until(() => js(`document.querySelectorAll('.im-batch-file input:checked').length === 3`), 'the three originals ticked', 3000).catch(() => {});
      await setValue('[data-role="batch-size"]', 'fit');
      await wait(100);
      await js(`(() => {
        const [w, h] = document.querySelectorAll('.im-batch-grid .rw-input[type="number"]');
        const set = (el, v) => { Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, String(v)); el.dispatchEvent(new Event('input', { bubbles: true })); };
        set(w, ${box});
        set(h, ${box});
        return 1;
      })()`);
      await setValue('[data-role="batch-format"]', format);
      await setValue('[data-role="batch-pattern"]', pattern);
      const stripped = await js(`document.querySelector('[data-role="batch-strip"]').checked`);
      if (stripped !== strip) await js(`document.querySelector('[data-role="batch-strip"]').click(), 1`);
      if (folder) {
        globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [folder] });
        await setValue('[data-role="batch-where"]', 'folder');
        await until(() => js(`document.querySelector('[data-role="batch-where"]').value === 'folder'`), 'the folder chosen', 4000).catch(() => {});
      }
      await wait(250);
      if (captureAs) await capture(captureAs);
      const example = await js(`document.querySelector('[data-role="batch-example"]')?.textContent || ''`);
      if (answer) globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer });
      await js(`document.querySelector('[data-role="batch-run"]').click(), 1`);
      await until(() => js(`Boolean(document.querySelector('.im-batch-run[data-state="done"]'))`), 'the batch to finish', 30000).catch(() => {});
      const sum = await js(`(() => { const n = document.querySelector('.im-batch-sum'); return n ? { written: Number(n.dataset.written), failed: Number(n.dataset.failed), text: n.textContent.trim() } : null; })()`);
      return { sum, example };
    };
    const read = (f) => {
      if (!fs.existsSync(f)) return null;
      const bytes = fs.readFileSync(f);
      const info = probeImage(bytes) || {};
      return { width: info.width, height: info.height, jpeg: bytes[0] === 0xff && bytes[1] === 0xd8, png: bytes[1] === 0x50, orientation: readExif(bytes).orientation ?? null };
    };

    // First run: fit inside 200 × 200, JPEG, "{name}-small", upright, stripped.
    const first = await runBatch({ box: 200, pattern: '{name}-small', captureAs: 'pictures-batch.png' });
    await capture('pictures-batch-done.png');
    const out = ['alpha', 'bravo', 'charlie'].map((n) => read(path.join(root, `${n}-small.jpg`)));
    check('pictures: Batch edit resizes three pictures to fit 200 × 200 as JPEGs named by the pattern, the one shot on its side turned upright, its metadata stripped',
      first.sum?.written === 3 && first.sum?.failed === 0
        && out[0]?.jpeg && out[0].width === 200 && out[0].height === 125
        && out[1]?.jpeg && out[1].width === 133 && out[1].height === 200 && out[1].orientation === null
        && out[2]?.jpeg && out[2].width === 200 && out[2].height === 200,
      JSON.stringify({ first, out }));
    await closeDialog();

    // Second run, the same: every name is taken. Asked, and Keep both writes beside.
    const smallBefore = fs.readFileSync(path.join(root, 'alpha-small.jpg'));
    const second = await runBatch({ box: 200, pattern: '{name}-small', answer: { response: 0 } });
    const beside = ['alpha', 'bravo', 'charlie'].map((n) => read(path.join(root, `${n}-small (2).jpg`)));
    const keptFirst = Buffer.compare(fs.readFileSync(path.join(root, 'alpha-small.jpg')), smallBefore) === 0;
    const untouched = [a, b, c].every((f, i) => Buffer.compare(fs.readFileSync(f), originals[i]) === 0);
    check('pictures: run again, Batch edit asks before replacing, and Keep both writes "(2)" beside the files already there, the originals and the first results untouched',
      second.sum?.written === 3 && beside.every((o) => o?.jpeg) && keptFirst && untouched,
      JSON.stringify({ second, beside, keptFirst, untouched }));
    await closeDialog();

    // Third run, into a folder chosen, each format kept and the metadata kept.
    const third = await runBatch({ box: 100, pattern: '{n:2} {name}', format: 'same', strip: false, folder: elsewhere });
    const listed = fs.existsSync(elsewhere) ? fs.readdirSync(elsewhere).sort() : [];
    const bravo = read(path.join(elsewhere, '02 bravo.jpg'));
    const alpha = read(path.join(elsewhere, '01 alpha.png'));
    check('pictures: Batch edit into a chosen folder keeps each format, numbers the names, and a JPEG keeps its metadata with its orientation set upright',
      third.sum?.written === 3 && listed.join('|') === '01 alpha.png|02 bravo.jpg|03 charlie.png' && alpha?.png && bravo?.orientation === 1 && bravo.width === 67 && bravo.height === 100,
      JSON.stringify({ third, listed, alpha, bravo }));
    await closeDialog();
    // The choices are remembered for the next batch: forgotten again, so a later check starts from the defaults.
    await js(`(() => { try { localStorage.removeItem('pictures.batch'); } catch {} return 1; })()`);
    const complaints = await errorsIn(win);
    check('pictures: Batch edit reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('pictures: the Batch edit checks ran', false, err?.message || String(err));
  }
}
