// Data Files' sample databases, checked in the window: the start screen
// offers each sample business with its records or empty; choosing the
// shop with its records makes the file and opens it with its tables,
// queries, a form for each table and its reports, and its takings report
// shows money; choosing the clinic empty makes the same design with no
// records in it.

import fs from 'node:fs';
import path from 'node:path';
import { Base, SAMPLE_IDS, sampleSpec } from '@rutba/database';

export async function verifyBaseSamples(h, { dir }) {
  const { open, check, until, wait } = h;
  const shopFile = path.join(dir, 'Corner Store.rbase');
  const clinicFile = path.join(dir, 'Clinic.rbase');
  try {
    for (const f of [shopFile, clinicFile]) for (const s of ['', '-wal', '-shm']) fs.rmSync(f + s, { force: true });
    let win = await open('base');
    let js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.bs-sample').length === ${SAMPLE_IDS.length}`), 'the sample databases', 6000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('.bs-sample')].map((b) => ({ id: b.dataset.sample, data: Boolean(b.querySelector('.bs-sample-data')), empty: Boolean(b.querySelector('.bs-sample-empty')) }))`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await js(`document.querySelector('.bs-samples')?.scrollIntoView(), 1`);
      await wait(400);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-samples.png'), (await win.webContents.capturePage()).toPNG());
    }
    check('base: the start screen offers every sample business, each with its records or empty',
      JSON.stringify(offered.map((o) => o.id)) === JSON.stringify(SAMPLE_IDS) && offered.every((o) => o.data && o.empty),
      JSON.stringify(offered));

    // The shop, with its records.
    const shop = sampleSpec('shop');
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: shopFile });
    await js(`document.querySelector('.bs-sample[data-sample="shop"] .bs-sample-data')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-reports [data-report="Takings by month"]'))`), 'the shop', 15000).catch(() => {});
    const made = await js(`({
      tables: [...document.querySelectorAll('.bs-tables:not(.bs-queries):not(.bs-forms):not(.bs-reports) [data-table]')].map((b) => b.dataset.table),
      queries: [...document.querySelectorAll('.bs-queries [data-query]')].map((b) => b.dataset.query),
      forms: [...document.querySelectorAll('.bs-forms [data-form]')].map((b) => b.dataset.form),
      reports: [...document.querySelectorAll('.bs-reports [data-report]')].map((b) => b.dataset.report),
    })`);
    const sales = (() => { const d = Base.open(shopFile); try { return d.count('Sales'); } catch { return 0; } finally { d.close(); } })();
    check('base: the shop with its records makes its tables, queries, forms and reports, and a year of sales',
      JSON.stringify(made.tables) === JSON.stringify(shop.tables.map((t) => t.name))
        && JSON.stringify(made.queries) === JSON.stringify(Object.keys(shop.queries))
        && made.forms.length === shop.tables.length
        && JSON.stringify(made.reports) === JSON.stringify(shop.reports.map((r) => r.name))
        && sales > 5000,
      JSON.stringify({ made, sales }));

    await js(`document.querySelector('.bs-reports [data-report="Takings by month"]')?.click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-report-scroll .bs-report-page tr').length > 6`), 'the takings report', 8000).catch(() => {});
    const page = await js(`({ rows: document.querySelectorAll('.bs-report-scroll .bs-report-page tr').length, text: document.querySelector('.bs-report-scroll .bs-report-page')?.textContent || '' })`);
    check("base: the shop's takings report shows a line for each month with money in it",
      page.rows > 6 && /\d{4}-\d{2}/.test(page.text) && /\d[\d,]*\.\d{2}/.test(page.text), JSON.stringify({ rows: page.rows, text: page.text.slice(0, 300) }));
    win.close();

    // The clinic, empty: its design and nothing in it.
    win = await open('base');
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-sample[data-sample="clinic"] .bs-sample-empty'))`), 'the clinic card', 6000).catch(() => {});
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: clinicFile });
    await js(`document.querySelector('.bs-sample[data-sample="clinic"] .bs-sample-empty')?.click(), 1`);
    const clinic = sampleSpec('clinic');
    await until(() => js(`document.querySelectorAll('.bs-reports [data-report]').length === ${clinic.reports.length}`), 'the empty clinic', 8000).catch(() => {});
    const empty = (() => {
      const d = Base.open(clinicFile);
      try { return { tables: d.tables().map((t) => t.name), records: d.tables().reduce((n, t) => n + d.count(t.name), 0), queries: d.queries().length }; } finally { d.close(); }
    })();
    check('base: the clinic made empty has its whole design and no records',
      JSON.stringify(empty.tables) === JSON.stringify(clinic.tables.map((t) => t.name)) && empty.records === 0 && empty.queries === Object.keys(clinic.queries).length,
      JSON.stringify(empty));
    win.close();
  } catch (err) {
    check('base: the sample database checks ran', false, err.stack || err.message);
  }
}
