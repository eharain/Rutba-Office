// Every form of every ready-made database, used in the window as a person
// would (the owner, 2026-10-22: "testing each template is important and all
// forms should work"). For each sample business with its records and each
// simple template, each form: its first record shows in it (a list shows
// the record it points at, not "(none)"; a subform's field that points at
// another table shows words, not a number; a label is not cut); the last
// record is reached; a text box changed is kept; a new record with every
// box filled is added and deleted again; and nothing says it went wrong.
// Long: asked for by name only, never part of the full run. Run with
// RUTBA_VERIFY_ONLY=base-forms-all, and RUTBA_VERIFY_FORMS=shop,clinic
// for some only; RUTBA_VERIFY_CAPTURE=<dir> keeps a picture of each form.

import fs from 'node:fs';
import path from 'node:path';
import { Base, SAMPLE_IDS, TEMPLATE_IDS, buildSample, buildTemplate } from '@rutba/database';
import { shownField } from '@rutba/database/form-design';

// What a person types in a box of each kind.
// A date no sample holds: the samples' records run up to the day they are made, and a day a
// table keeps once (a cafe's cash-up) is taken by the time the check runs a few days on.
const TYPED = { text: 'Form check', longtext: 'Typed on the form.', number: '7', currency: '4.50', date: '2001-02-03', datetime: '2001-02-03T09:30' };

export async function verifyBaseFormsAll({ open, check, until, wait }, { dir }) {
  const wanted = (process.env.RUTBA_VERIFY_FORMS || '').split(',').map((s) => s.trim()).filter(Boolean);
  const all = [...SAMPLE_IDS, ...TEMPLATE_IDS.map((id) => `template:${id}`)].filter((id) => !wanted.length || wanted.includes(id) || wanted.includes(id.replace('template:', '')));
  const shots = process.env.RUTBA_VERIFY_CAPTURE || null;
  // Each database's window goes once the next is open: the last window closed would end the run.
  let previous = null;
  for (const id of all) {
    const file = path.join(dir, `Forms ${id.replace(':', ' ')}.rbase`);
    const faults = [];
    let formCount = 0;
    let win = null;
    try {
      for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
      // The file as the start screen makes it, and what each form should show of its first record.
      const b = Base.open(file);
      if (id.startsWith('template:')) buildTemplate(b, id.slice(9));
      else buildSample(b, id, { withData: true });
      const rels = b.relationships();
      const forms = b.forms().map(({ name, layout }) => {
        const count = b.db.prepare(`SELECT COUNT(*) AS n FROM "${layout.table.replace(/"/g, '""')}"`).get().n;
        // A subform's columns that point at another table: these show that record's words.
        const pointing = Object.fromEntries((layout.subforms || []).map((s) => [s.table, s.columns.filter((c) => c !== s.field && rels.some((r) => r.from.table === s.table && r.from.field === c))]));
        // The words each of those columns may show: a record named 2025 (a pay year) is words, not a number.
        const words = {};
        for (const [tbl, cols] of Object.entries(pointing)) for (const col of cols) {
          const rel = rels.find((r) => r.from.table === tbl && r.from.field === col);
          words[`${tbl}.${col}`] = new Set(b.lookup({ table: rel.to.table, field: rel.to.field, show: shownField(b.fields(rel.to.table)), limit: 1_000_000 }).map((o) => o.text));
        }
        // A list whose field must be filled: with nothing to choose (an empty template's), no record can be added yet.
        const required = b.fields(layout.table).filter((x) => x.required).map((x) => x.name);
        return { name, table: layout.table, first: b.rows(layout.table, { limit: 1 })[0] || null, count, pointing, words, required };
      });
      b.close();
      formCount = forms.length;

      win = await open('base', file);
      try { previous?.destroy(); } catch { /* gone already */ }
      previous = win;
      const js = (code) => win.webContents.executeJavaScript(code);
      win.setSize(1456, 910);
      await until(() => js(`document.querySelectorAll('.bs-forms [data-form]').length === ${forms.length}`), `${id}'s forms`, 15000);
      const toasts = () => js(`[...document.querySelectorAll('.rw-toast.bad')].map((n) => n.textContent)`);
      const said = new Set();
      const fault = (form, text) => faults.push(`${form}: ${text}`);
      const newToasts = async (form, when) => {
        for (const t of await toasts()) if (!said.has(t)) { said.add(t); fault(form, `${when}: "${t}"`); }
      };

      for (const f of forms) {
        const sel = JSON.stringify(f.name);
        await js(`document.querySelector('.bs-forms [data-form=${JSON.stringify(f.name).replace(/'/g, "\\'")}]')?.click(), 1`);
        const opened = await until(() => js(`Boolean(document.querySelector('.bs-form-page[data-form=${sel}]'))`), `${id} ${f.name}`, 8000).then(() => true, () => false);
        if (!opened) { fault(f.name, 'the form did not open'); continue; }
        // Its lists filled, and its subforms with the first record's.
        await until(() => js(`[...document.querySelectorAll('.bs-form-page select.bs-fc-list')].every((s) => s.options.length > 1 || s.dataset.empty === '1')`), 'lists', 2500).catch(() => {});
        await wait(500);

        // The first record, as it shows.
        const seen = await js(`(() => {
          const page = document.querySelector('.bs-form-page[data-form=${sel}]');
          const box = page.getBoundingClientRect();
          const controls = [...page.querySelectorAll('.bs-fc')].map((n) => {
            const r = n.getBoundingClientRect();
            const cs = getComputedStyle(n);
            const room = n.tagName === 'SELECT' ? n.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom) : null;
            return { field: n.dataset.field, tag: n.tagName, kind: [...n.classList].find((c) => c.startsWith('bs-fc-'))?.slice(6), value: n.classList.contains('bs-fc-search') ? n.querySelector('input')?.value ?? '' : n.tagName === 'SPAN' ? n.textContent : n.type === 'checkbox' ? n.checked : n.value, shown: n.tagName === 'SELECT' ? n.selectedOptions[0]?.textContent : null, options: n.tagName === 'SELECT' ? n.options.length : null, room, line: parseFloat(cs.fontSize) * 1.2, r: [r.left, r.top, r.right, r.bottom] };
          });
          const labels = [...page.querySelectorAll('.bs-fl')].map((n) => ({ text: n.textContent, cut: n.scrollWidth > n.clientWidth + 1 }));
          const subs = [...page.querySelectorAll('.bs-subform')].map((n) => ({ table: n.dataset.subform, r: (() => { const r = n.getBoundingClientRect(); return [r.left, r.top, r.right, r.bottom]; })(), cols: [...n.querySelectorAll('thead th')].map((t) => t.textContent), cells: [...n.querySelectorAll('tbody tr')].slice(0, 3).map((tr) => [...tr.children].map((td) => td.textContent)) }));
          return { controls, labels, subs, place: document.querySelector('.bs-nav-place')?.textContent || '', page: [box.left, box.top, box.right, box.bottom] };
        })()`);
        for (const l of seen.labels) if (l.cut) fault(f.name, `label "${l.text}" is cut`);
        // Nothing stands on anything else.
        const boxes = [...seen.controls.map((c) => ({ what: c.field, r: c.r })), ...seen.subs.map((s) => ({ what: `subform ${s.table}`, r: s.r }))];
        for (let i = 0; i < boxes.length; i++) for (let j = i + 1; j < boxes.length; j++) {
          const [a, c] = [boxes[i].r, boxes[j].r];
          if (a[0] < c[2] - 1 && c[0] < a[2] - 1 && a[1] < c[3] - 1 && c[1] < a[3] - 1) fault(f.name, `${boxes[i].what} and ${boxes[j].what} overlap`);
        }
        for (const c of seen.controls) {
          if (c.tag === 'SELECT' && c.room < c.line) fault(f.name, `list ${c.field} cuts its words (${c.room}px for a ${c.line}px line)`);
          if (!f.first) continue;
          const v = f.first[c.field];
          if (v === null || v === undefined || v === '') continue;
          // A long list is a box showing the words of the record it points at.
          if (c.kind === 'search') { if (!c.value) fault(f.name, `list ${c.field} shows nothing for ${JSON.stringify(v)}`); continue; }
          if (c.tag === 'SELECT' && (c.value === '' || c.value !== String(v))) fault(f.name, `list ${c.field} shows "${c.shown}" for ${JSON.stringify(v)} (${c.options} choices)`);
          else if (c.kind === 'date' && c.value !== String(v).slice(0, 10)) fault(f.name, `date ${c.field} shows "${c.value}" for ${JSON.stringify(v)}`);
          else if (c.kind === 'datetime' && c.value !== String(v).slice(0, 16)) fault(f.name, `date and time ${c.field} shows "${c.value}" for ${JSON.stringify(v)}`);
          else if (c.kind === 'check' && c.value !== Boolean(v)) fault(f.name, `check ${c.field} shows ${c.value} for ${JSON.stringify(v)}`);
          else if (['text', 'longtext', 'number', 'currency'].includes(c.kind) && c.value !== String(v)) fault(f.name, `${c.kind} ${c.field} shows "${c.value}" for ${JSON.stringify(v)}`);
        }
        for (const s of seen.subs) for (const row of s.cells) row.forEach((cell, i) => {
          if (/^\d+$/.test(cell) && f.pointing[s.table]?.includes(s.cols[i]) && !f.words[`${s.table}.${s.cols[i]}`]?.has(cell)) fault(f.name, `subform ${s.table} shows ${cell} for ${s.cols[i]}`);
        });
        if (f.count && seen.place !== `Record 1 of ${f.count}`) fault(f.name, `shows "${seen.place}", not record 1 of ${f.count}`);
        await newToasts(f.name, 'opening it');
        if (shots) {
          win.webContents.invalidate();
          await wait(300);
          fs.writeFileSync(path.join(shots, `forms-${id.replace(':', '-')}-${f.name.replace(/[^A-Za-z0-9]+/g, '-')}.png`), (await win.webContents.capturePage()).toPNG());
        }

        // The last record.
        if (f.count > 1) {
          await js(`document.querySelector('.bs-nav-last')?.click(), 1`);
          const last = await until(() => js(`document.querySelector('.bs-nav-place')?.textContent === 'Record ${f.count} of ${f.count}'`), 'the last record', 5000).then(() => true, () => false);
          if (!last) fault(f.name, `the last record did not come: "${await js(`document.querySelector('.bs-nav-place')?.textContent`)}"`);
          await newToasts(f.name, 'going to the last record');
        }

        // A text box changed is kept, and put back.
        if (f.count) {
          const change = (suffix) => js(`(async () => {
            const el = [...document.querySelectorAll('.bs-form-page input.bs-fc-text')].find((n) => !n.readOnly);
            if (!el) return null;
            const was = el.dataset.was ?? el.value;
            el.dataset.was = was;
            Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, was + ${JSON.stringify(suffix)});
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
            await new Promise((r) => setTimeout(r, 700));
            return el.dataset.field;
          })()`);
          const stored = (field) => { const x = Base.open(file); try { return x.rows(f.table, { offset: f.count - 1, limit: 1 })[0]?.[field] ?? null; } finally { x.close(); } };
          const edited = await change(' (checked)');
          if (edited) {
            if (!String(stored(edited) ?? '').endsWith(' (checked)')) fault(f.name, `${edited} changed on the form was not kept: ${JSON.stringify(stored(edited))}`);
            await change('');
            if (String(stored(edited) ?? '').endsWith(' (checked)')) fault(f.name, `${edited} put back on the form was not kept`);
          }
          await newToasts(f.name, 'changing a text box');
        }

        // A new record, every box filled, added and deleted again; not where a list that must be filled has nothing to choose.
        await js(`document.querySelector('.bs-nav-new')?.click(), 1`);
        const blank = await until(() => js(`document.querySelector('.bs-nav-place')?.textContent === 'New record'`), 'a new record', 5000).then(() => true, () => false);
        if (!blank) { fault(f.name, 'New Record did not give a blank record'); continue; }
        await wait(300);
        const nothingToChoose = await js(`${JSON.stringify(f.required)}.some((field) => { const el = document.querySelector('.bs-form-page select.bs-fc-list[data-field=' + JSON.stringify(field) + ']'); return el && el.options.length < 2; })`);
        if (nothingToChoose) continue;
        await js(`(async () => {
          const typed = ${JSON.stringify(TYPED)};
          const page = document.querySelector('.bs-form-page[data-form=${sel}]');
          for (const el of page.querySelectorAll('.bs-fc')) {
            const kind = [...el.classList].find((c) => c.startsWith('bs-fc-'))?.slice(6);
            if (el.classList.contains('bs-fc-search')) {
              // A long list: the box searched (empty, its first matches), the first picked.
              el.querySelector('input').dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
              let li = null;
              for (let i = 0; i < 30 && !li; i++) { await new Promise((r) => setTimeout(r, 100)); li = el.querySelector('.bs-fc-search-list li'); }
              li?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
              await new Promise((r) => setTimeout(r, 100));
              continue;
            }
            if (el.tagName === 'SPAN' || el.readOnly) continue;
            if (el.type === 'checkbox') { el.click(); continue; }
            if (el.tagName === 'SELECT') {
              const o = [...el.options].find((x) => x.value !== '');
              if (o) { el.value = o.value; el.dispatchEvent(new Event('change', { bubbles: true })); }
              continue;
            }
            let v = typed[kind] ?? 'Form check';
            // A number that should be unique (a room, a chair) is less likely to be taken.
            if (kind === 'number') v = String(900 + Math.floor(Math.random() * 90));
            const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
            Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, v);
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
            await new Promise((r) => setTimeout(r, 20));
          }
          return 1;
        })()`);
        await wait(300);
        await js(`document.querySelector('.bs-nav-add')?.click(), 1`);
        const added = await until(() => js(`document.querySelector('.bs-nav-place')?.textContent === 'Record ${f.count + 1} of ${f.count + 1}'`), 'the record added', 6000).then(() => true, () => false);
        if (!added) {
          const t = await toasts();
          fault(f.name, `a new record was not added: ${t.filter((x) => !said.has(x)).join(' | ') || await js(`document.querySelector('.bs-nav-place')?.textContent`)}`);
          for (const x of t) said.add(x);
          continue;
        }
        await newToasts(f.name, 'adding a record');
        await js(`document.querySelector('.bs-nav-delete')?.click(), 1`);
        // Delete asks first (2.18): yes.
        await until(() => js(`Boolean(document.querySelector('.bs-delete-yes'))`), 'the question', 3000).catch(() => {});
        await js(`document.querySelector('.bs-delete-yes')?.click(), 1`);
        const gone = await until(() => js(`(document.querySelector('.bs-nav-place')?.textContent || '').endsWith('of ${f.count}') || ${f.count === 0} && document.querySelector('.bs-nav-place')?.textContent === 'New record'`), 'the record deleted', 6000).then(() => true, () => false);
        if (!gone) fault(f.name, `the added record was not deleted: "${await js(`document.querySelector('.bs-nav-place')?.textContent`)}"`);
        await newToasts(f.name, 'deleting the record');
      }
    } catch (err) {
      faults.push(`the check stopped: ${err?.message || String(err)}`);
    }
    check(`base: ${id}'s ${formCount} forms each show their first record, reach the last, keep a change, and add and delete a record`, faults.length === 0, faults.slice(0, 40).join('\n        '));
  }
}
