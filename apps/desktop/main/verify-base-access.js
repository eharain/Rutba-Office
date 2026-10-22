// People, roles and protection in Data Files (D-2026-10-22-11): a protected
// database opens on its sign-in; a cashier signed in sees the tables and
// fields their role shows, nothing of the design, and no Delete or new row
// where the role only reads; Sign out locks it again; the Admin opens People
// and roles, and sees the database as the cashier does, then goes back; a
// database is protected from its side panel. Run alone with
// RUTBA_VERIFY_ONLY=base-access.

import fs from 'node:fs';
import path from 'node:path';
import { Base, buildFromSpec } from '@rutba/database';

const set = (sel, value) => `(() => { const el = document.querySelector(${JSON.stringify(sel)}); const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype; Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event(el instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true })); return 1; })()`;

export async function verifyBaseAccess({ open, check, until, wait }, { dir }) {
  const file = path.join(dir, 'Protected.rbase');
  const plain = path.join(dir, 'To protect.rbase');
  try {
    for (const f of [file, plain]) for (const s of ['', '-wal', '-shm']) fs.rmSync(f + s, { force: true });
    const b = Base.open(file);
    buildFromSpec(b, {
      tables: [
        { name: 'Products', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'currency' }, { name: 'Cost', type: 'currency' }] },
        { name: 'Sales', fields: [{ name: 'Product', type: 'integer', points: 'Products' }, { name: 'Quantity', type: 'integer' }] },
        { name: 'Wages', fields: [{ name: 'Who', type: 'text' }, { name: 'Amount', type: 'currency' }] },
      ],
    });
    b.insert('Products', { Name: 'Bread', Price: 2.5, Cost: 1.1 });
    b.insert('Wages', { Who: 'Amira', Amount: 900 });
    b.protect({ name: 'Amira', password: 'owner-pass' });
    b.saveRole('Cashier', { tables: { default: 'read', each: { Products: { access: 'read', fields: { Cost: 'hidden' } }, Sales: { access: 'add', own: true }, Wages: 'hidden' } } });
    b.addPerson({ name: 'Ben', role: 'Cashier', password: 'till-pass' });
    b.close();
    const p = Base.open(plain);
    buildFromSpec(p, { tables: [{ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] }] });
    p.close();

    let win = await open('base', file);
    let js = (code) => win.webContents.executeJavaScript(code);
    // Locked: the sign-in, the names that may sign in, and nothing of the database.
    await until(() => js(`Boolean(document.querySelector('.bs-signin'))`), 'the sign-in', 8000);
    const locked = await js(`({ names: [...document.querySelectorAll('.bs-signin-name option')].map((o) => o.value), tables: document.querySelectorAll('[data-table]').length })`);
    // A wrong password is refused, in words.
    await js(set('.bs-signin-name', 'Ben'));
    await js(set('.bs-signin-password', 'wrong'));
    await js(`document.querySelector('.bs-signin-go').click(), 1`);
    const refused = await until(() => js(`/do not match/.test(document.querySelector('.bs-signin-error')?.textContent || '')`), 'refused', 5000).then(() => true, () => false);
    // Ben, the cashier.
    await js(set('.bs-signin-password', 'till-pass'));
    await js(`document.querySelector('.bs-signin-go').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-table="Products"]'))`), 'the database as Ben', 8000);
    await until(() => js(`document.querySelectorAll('.bs-th[data-field]').length > 0`), 'the records', 5000);
    const cashier = await js(`({
      tables: [...document.querySelectorAll('.bs-tables [data-table]')].map((n) => n.dataset.table),
      fields: [...document.querySelectorAll('.bs-th[data-field]')].map((n) => n.dataset.field),
      newTable: Boolean(document.querySelector('.bs-new-table')), designTab: Boolean(document.querySelector('.bs-tab[data-view="design"]')),
      sql: Boolean(document.querySelector('[data-view="sql"]')), people: Boolean(document.querySelector('.bs-people-open')),
      del: Boolean(document.querySelector('.bs-delete')), newRow: Boolean(document.querySelector('.bs-new')),
      me: document.querySelector('.bs-me')?.dataset.me,
    })`);
    // Sales: Ben may add, so the new row is there, and no Delete.
    await js(`document.querySelector('[data-table="Sales"]').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-new'))`), 'the new row in Sales', 5000).catch(() => {});
    const sales = await js(`({ newRow: Boolean(document.querySelector('.bs-new')), del: Boolean(document.querySelector('.bs-delete')) })`);
    // Signed out: locked again.
    await js(`document.querySelector('.bs-sign-out').click(), 1`);
    const lockedAgain = await until(() => js(`Boolean(document.querySelector('.bs-signin'))`), 'locked again', 5000).then(() => true, () => false);
    check('base: a protected database opens on its sign-in; a wrong password is refused; the cashier sees the tables and fields the role shows, nothing of the design, no Delete or new row where it only reads; Sign out locks it again',
      JSON.stringify(locked.names) === JSON.stringify(['Amira', 'Ben']) && locked.tables === 0 && refused
        && JSON.stringify(cashier.tables) === JSON.stringify(['Products', 'Sales']) && !cashier.fields.includes('Cost') && cashier.fields.includes('Price')
        && !cashier.newTable && !cashier.designTab && !cashier.sql && !cashier.people && !cashier.del && !cashier.newRow && cashier.me === 'Ben'
        && sales.newRow && !sales.del && lockedAgain,
      JSON.stringify({ locked, refused, cashier, sales, lockedAgain }));

    // Amira, the Admin: People and roles, and the database seen as the cashier sees it.
    await js(set('.bs-signin-name', 'Amira'));
    await js(set('.bs-signin-password', 'owner-pass'));
    await js(`document.querySelector('.bs-signin-go').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-people-open'))`), 'the database as Amira', 8000);
    const admin = await js(`({ tables: [...document.querySelectorAll('.bs-tables [data-table]')].map((n) => n.dataset.table), newTable: Boolean(document.querySelector('.bs-new-table')) })`);
    await js(`document.querySelector('.bs-people-open').click(), 1`);
    await until(() => js(`document.querySelectorAll('[data-person]').length === 2`), 'the people', 5000);
    await until(() => js(`Boolean(document.querySelector('[data-editing="Cashier"]')) || [...document.querySelectorAll('[data-role]')].some((n) => n.dataset.role === 'Cashier')`), 'the roles', 5000);
    await js(`[...document.querySelectorAll('[data-role]')].find((n) => n.dataset.role === 'Cashier').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-editing="Cashier"]'))`), 'the cashier role', 5000);
    const grid = await js(`({
      people: [...document.querySelectorAll('[data-person]')].map((n) => n.dataset.person),
      wages: document.querySelector('[data-role-table="Wages"] .bs-role-access')?.value,
      products: document.querySelector('[data-role-table="Products"] .bs-role-access')?.value,
    })`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await js(`document.querySelector('[data-role-table="Products"] .bs-role-fields-toggle').click(), 1`);
      await wait(300);
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-access-people.png'), (await win.webContents.capturePage()).toPNG());
    }
    // Wages, read only for the cashier from now on: kept in the file.
    await js(set('[data-role-table="Wages"] .bs-role-access', 'read'));
    await until(() => js(`document.querySelector('[data-role-table="Wages"] .bs-role-access')?.value === 'read'`), 'wages read', 5000).catch(() => {});
    await wait(400);
    await js(set('[data-role-table="Wages"] .bs-role-access', 'hidden'));
    await wait(400);
    await js(`document.querySelector('.bs-role-see').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-seeing'))`), 'seeing as the cashier', 5000);
    const seeing = await js(`({ tables: [...document.querySelectorAll('.bs-tables [data-table]')].map((n) => n.dataset.table), seeing: document.querySelector('.bs-seeing')?.dataset.seeing, newTable: Boolean(document.querySelector('.bs-new-table')) })`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-access-seeing.png'), (await win.webContents.capturePage()).toPNG());
    }
    await js(`document.querySelector('.bs-seeing-back').click(), 1`);
    const back = await until(() => js(`!document.querySelector('.bs-seeing') && document.querySelectorAll('.bs-tables [data-table]').length === 3`), 'back', 5000).then(() => true, () => false);
    check('base: the Admin sees every table and the design, lists the people, gives a role what it may do in People and roles, and sees the database as that role does, then goes back',
      JSON.stringify(admin.tables) === JSON.stringify(['Products', 'Sales', 'Wages']) && admin.newTable
        && JSON.stringify(grid.people) === JSON.stringify(['Amira', 'Ben']) && grid.wages === 'hidden' && grid.products === 'read'
        && JSON.stringify(seeing.tables) === JSON.stringify(['Products', 'Sales']) && seeing.seeing === 'Cashier' && !seeing.newTable && back,
      JSON.stringify({ admin, grid, seeing, back }));

    // Protected from the side panel: the database that was not.
    const first = win;
    win = await open('base', plain);
    first.close();
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-protect'))`), 'Protect this database', 8000);
    await js(`document.querySelector('.bs-protect').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-protect-name'))`), 'the protect box', 5000);
    await js(set('.bs-protect-name', 'Chen'));
    await js(set('.bs-protect-password', 'chen-pass'));
    await js(set('.bs-protect-again', 'chen-pass'));
    await js(`document.querySelector('.bs-protect-go').click(), 1`);
    const protectedNow = await until(() => js(`document.querySelectorAll('[data-person]').length === 1`), 'protected', 6000).then(() => true, () => false);
    const after = await js(`({ me: document.querySelector('.bs-me')?.dataset.me, protect: Boolean(document.querySelector('.bs-protect')), roles: [...document.querySelectorAll('[data-role]')].map((n) => n.dataset.role) })`);
    await wait(300);
    const kept = (() => { const x = Base.open(plain); try { return { protected: x.isProtected(), names: x.signInNames() }; } finally { x.close(); } })();
    check('base: Protect this database makes its protector the Admin, opens People and roles with two roles to start from, and is kept in the file',
      protectedNow && after.me === 'Chen' && !after.protect && JSON.stringify(after.roles) === JSON.stringify(['Staff', 'Read only'])
        && kept.protected && JSON.stringify(kept.names) === JSON.stringify(['Chen']),
      JSON.stringify({ protectedNow, after, kept }));
  } catch (err) {
    check('base: the protection checks ran to the end', false, err?.message || String(err));
  }
}
