// The pictures of Data Files' sample businesses for office.rutba.io: each
// sample made with its records from the start screen and shown at its first
// print template (Northbank Trading's invoice, a shop's till receipt), its
// tables beside it, and captured as base-sample-<id>.png. Run on its own,
// never in the full run, as twenty-five databases take minutes:
//
//   RUTBA_VERIFY_ONLY=base-sample-shots RUTBA_VERIFY_CAPTURE=<dir> npm run verify:apps
//
// RUTBA_SAMPLE_SHOT_IDS=shop,clinic takes only those again; RUTBA_SAMPLE_SHOT_VIEW
// = data or form shows the table's records or its form instead.
//
// Without RUTBA_VERIFY_CAPTURE it still makes each sample and checks it
// opens filled, and writes nothing.

import fs from 'node:fs';
import path from 'node:path';
import { SAMPLE_IDS, sampleSpec, sampleAbout } from '@rutba/database';

/** The table a sample's picture shows: its first print template's, else its first of its own. */
function shownTable(id) {
  const spec = sampleSpec(id);
  const own = spec.tables.map((t) => t.name).filter((n) => n !== 'Start Here' && n !== 'Business');
  return spec.documents?.[0]?.table || own[0];
}

export async function verifyBaseSampleShots(h, { dir }) {
  const { open, check, until, wait } = h;
  const out = process.env.RUTBA_VERIFY_CAPTURE || null;
  const shown = [];
  const only = (process.env.RUTBA_SAMPLE_SHOT_IDS || '').split(',').filter(Boolean);
  const view = process.env.RUTBA_SAMPLE_SHOT_VIEW || 'document';
  for (const id of only.length ? SAMPLE_IDS.filter((x) => only.includes(x)) : SAMPLE_IDS) {
    const file = path.join(dir, `${sampleAbout(id).name}.rbase`);
    let win = null;
    try {
      for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
      win = await open('base');
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.bs-sample[data-sample="${id}"] .bs-sample-data'))`), `the ${id} card`, 8000).catch(() => {});
      globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: file });
      await js(`document.querySelector('.bs-sample[data-sample="${id}"] .bs-sample-data')?.click(), 1`);
      const table = shownTable(id);
      const sel = JSON.stringify(`.bs-tables:not(.bs-queries):not(.bs-forms):not(.bs-reports) [data-table="${table}"]`);
      await until(() => js(`Boolean(document.querySelector(${sel}))`), `${id}: its tables`, 30000).catch(() => {});
      await js(`document.querySelector(${sel})?.click(), 1`);
      await until(() => js(`document.querySelectorAll('.bs-row[data-rowid]').length > 5`), `${id}: the records of ${table}`, 10000).catch(() => {});
      if (view === 'form') {
        await js(`document.querySelector('.bs-forms [data-form="${table}"]')?.click(), 1`);
        await wait(1500);
      }
      if (view === 'document') {
        const doc = sampleSpec(id).documents?.[0]?.name;
        await js(`document.querySelector('.bs-documents [data-document=${JSON.stringify(doc)}]')?.click(), 1`);
        await until(() => js(`Boolean(document.querySelector('.bs-doc-page'))`), `${id}: its ${doc}`, 15000).catch(() => {});
        await wait(800);
      }
      // What is on screen: the table's records, or the printed page, filled.
      const rows = view === 'document'
        ? await js(`document.querySelector('.bs-doc-page')?.textContent.length > 200 ? 99 : 0`)
        : await js(`document.querySelectorAll('.bs-row[data-rowid]').length`);
      const dialog = await js(`Boolean(document.querySelector('[role="dialog"], [role="alertdialog"]'))`);
      if (out) {
        await wait(500);
        fs.writeFileSync(path.join(out, `base-sample-${id}.png`), (await win.webContents.capturePage()).toPNG());
      }
      shown.push({ id, table, rows, dialog });
    } catch (err) {
      shown.push({ id, error: err.message });
    } finally {
      try { win?.close(); } catch { /* closed already */ }
    }
  }
  const bad = shown.filter((s) => s.error || !(s.rows > 5) || s.dialog);
  check(`base: each sample business asked for opens with its records and its first print template showing, nothing over them${out ? ', and is captured' : ''}`,
    bad.length === 0 && shown.length === (only.length || SAMPLE_IDS.length), JSON.stringify(bad.length ? bad : shown.map((s) => `${s.id}:${s.table}:${s.rows}`)));
}
