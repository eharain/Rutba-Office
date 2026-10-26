// Databases made elsewhere in the Data Files window (A2.20): another
// program's SQLite database opens read only, its tables listed with their
// records and nothing of the design to change, with a band that says so;
// Allow changes asks first and then lets its records change, its file still
// given nothing of ours; and a database made by a newer Rutba Office is
// refused in plain words, left as it was. Run alone with
// RUTBA_VERIFY_ONLY=base-elsewhere.

import fs from 'node:fs';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { DESIGN_VERSION } from '@rutba/database';

export async function verifyBaseElsewhere({ open, check, until, wait }, { dir }) {
  const file = path.join(dir, 'Other program.db');
  const newer = path.join(dir, 'Newer.rbase');
  try {
    for (const f of [file, newer]) for (const s of ['', '-wal', '-shm']) fs.rmSync(f + s, { force: true });
    const raw = new DatabaseSync(file);
    raw.exec(`CREATE TABLE members (id INTEGER PRIMARY KEY, name TEXT NOT NULL, joined DATE);
      INSERT INTO members VALUES (1, 'Asha', '2025-04-01'), (2, 'Ben', '2026-01-15');
      CREATE TABLE settings (k TEXT PRIMARY KEY, v TEXT) WITHOUT ROWID;`);
    raw.close();
    const before = fs.readFileSync(file);
    const n = new DatabaseSync(newer);
    n.exec(`CREATE TABLE rutba_meta (key TEXT PRIMARY KEY, value TEXT); CREATE TABLE rutba_tables (name TEXT PRIMARY KEY, caption TEXT, position INTEGER);
      CREATE TABLE rutba_fields (tbl TEXT NOT NULL, name TEXT NOT NULL, type TEXT NOT NULL, size INTEGER, required INTEGER NOT NULL DEFAULT 0, uniq INTEGER NOT NULL DEFAULT 0, dflt TEXT, caption TEXT, position INTEGER NOT NULL, PRIMARY KEY (tbl, name));
      INSERT INTO rutba_meta VALUES ('version', '${DESIGN_VERSION + 1}');`);
    n.close();
    const newerBefore = fs.readFileSync(newer);

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('[data-table="members"]'))`), 'the other program\'s tables', 8000);
    const read = await js(`({
      band: document.querySelector('.bs-foreign')?.dataset.foreign || null,
      words: document.querySelector('.bs-foreign')?.textContent || '',
      tables: [...document.querySelectorAll('.bs-side [data-table]')].map((x) => x.dataset.table),
      newTable: Boolean(document.querySelector('.bs-new-table')),
      importing: Boolean(document.querySelector('.bs-import')),
      sql: Boolean(document.querySelector('[data-view="sql"]')),
      addRow: [...document.querySelectorAll('.bs-new-input')].filter((x) => !x.disabled).length,
    })`);
    // Allow changes: asked first, then the band says changes go into its file.
    await js(`document.querySelector('.bs-foreign-allow').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-foreign-allow-go'))`), 'the question', 5000);
    const asked = await js(`document.querySelector('.rw-dialog, [role="dialog"]')?.textContent || ''`);
    await js(`document.querySelector('.bs-foreign-allow-go').click(), 1`);
    await until(() => js(`document.querySelector('.bs-foreign')?.dataset.foreign === 'write'`), 'changes allowed', 5000).catch(() => {});
    const written = await js(`({ band: document.querySelector('.bs-foreign')?.dataset.foreign || null, allow: Boolean(document.querySelector('.bs-foreign-allow')), newTable: Boolean(document.querySelector('.bs-new-table')) })`);
    // A newer design: refused in words, nothing changed. (Opened before the first closes: the last window closing ends the run.)
    const w2 = await open('base', newer);
    const js2 = (code) => w2.webContents.executeJavaScript(code);
    win.close();
    await wait(600);
    const untouched = Buffer.from(fs.readFileSync(file)).equals(before) && !fs.existsSync(`${file}-wal`);
    const refused = await until(() => js2(`[...document.querySelectorAll('.rw-toast')].some((t) => /made by a newer version of Rutba Office/.test(t.textContent))`), 'the refusal', 8000).then(() => true, () => false);
    const stillThere = await js2(`Boolean(document.querySelector('.bs-new-db'))`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      w2.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-elsewhere-newer.png'), (await w2.webContents.capturePage()).toPNG());
    }
    const newerUntouched = Buffer.from(fs.readFileSync(newer)).equals(newerBefore);

    check('base: another program\'s database opens read only with a band saying so, its tables listed, nothing of the design offered, and the table it cannot show named',
      read.band === 'read' && /made by another program/.test(read.words) && /settings/.test(read.words) && JSON.stringify(read.tables) === JSON.stringify(['members'])
        && !read.newTable && !read.importing && read.sql && read.addRow === 0,
      JSON.stringify(read));
    check('base: Allow changes asks first, then its records may change; the design still is not offered, and the file is given nothing of ours',
      /goes straight into its file/.test(asked) && written.band === 'write' && !written.allow && !written.newTable && untouched,
      JSON.stringify({ asked: asked.slice(0, 120), written, untouched }));
    check('base: a database made by a newer Rutba Office is refused in plain words and left as it was',
      refused && stillThere && newerUntouched, JSON.stringify({ refused, stillThere, newerUntouched }));
  } catch (err) {
    check('base: the checks of databases made elsewhere ran to the end', false, err?.message || String(err));
  }
}
