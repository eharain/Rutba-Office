// A form's lists and subforms, as the sample films found them (2026-10-22):
// a list's words fit in its box, not cut to their tops, and a subform's
// field that points at another table shows that record's words, not its
// number. On Larkspur Corner Store's Sales form. Run alone with
// RUTBA_VERIFY_ONLY=base-form-lists.

import fs from 'node:fs';
import path from 'node:path';
import { Base, buildSample } from '@rutba/database';

export async function verifyBaseFormLists({ open, check, until, wait }, { dir }) {
  const file = path.join(dir, 'Form lists.rbase');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const b = Base.open(file);
    buildSample(b, 'shop', { withData: true });
    // The Sales form: lists of its own (the customer, the payment) and the sale's lines, each pointing at a product.
    const pointing = (f) => (f.layout.subforms || []).some((s) => s.columns.some((c) => b.relationships().some((r) => r.from.table === s.table && r.from.field === c && c !== s.field)));
    const form = b.forms().find((f) => f.name === 'Sales' && pointing(f)) || b.forms().find((f) => pointing(f) && f.layout.controls.some((c) => c.kind === 'list'));
    const sub = form.layout.subforms.find((s) => s.columns.some((c) => b.relationships().some((r) => r.from.table === s.table && r.from.field === c && c !== s.field)));
    const col = sub.columns.find((c) => c !== sub.field && b.relationships().some((r) => r.from.table === sub.table && r.from.field === c));
    b.close();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('[data-form=${JSON.stringify(form.name)}]'))`), 'the database', 10000);
    await js(`document.querySelector('[data-form=${JSON.stringify(form.name)}]').click(), 1`);
    await until(() => js(`document.querySelectorAll('[data-subform=${JSON.stringify(sub.table)}] tbody tr').length > 0`), 'the subform\'s records', 8000);
    await wait(500);
    const seen = await js(`(() => {
      const list = document.querySelector('.bs-fc-list');
      const cs = list ? getComputedStyle(list) : null;
      const line = cs ? parseFloat(cs.fontSize) * 1.2 : 0;
      const room = list ? list.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom) : 0;
      const cells = [...document.querySelectorAll('[data-subform=${JSON.stringify(sub.table)}] tbody td[data-col=${JSON.stringify(col)}]')].map((n) => n.textContent);
      return { form: ${JSON.stringify(form.name)}, field: list?.dataset.field, height: list?.clientHeight, room, line, cells: cells.slice(0, 4) };
    })()`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-form-lists.png'), (await win.webContents.capturePage()).toPNG());
    }
    check('base: a form\'s list has room for its words, and a subform\'s field that points at another table shows that record\'s words, not its number',
      Boolean(seen.field) && seen.room >= seen.line && seen.cells.length > 0 && seen.cells.every((c) => c && !/^\d+$/.test(c)),
      JSON.stringify(seen));
  } catch (err) {
    check('base: the form list checks ran to the end', false, err?.message || String(err));
  }
}
