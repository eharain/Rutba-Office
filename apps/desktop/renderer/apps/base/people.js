// People, roles and protection in Data Files (D-2026-10-22-11): the sign-in
// a protected database asks for as it opens, Protect this database, your own
// password, and the Admin's People and roles, where each role is given what
// it may see and do, table by table and field by field, and Seen as it would
// be before anyone relies on it. The database itself enforces all of it
// (packages/database/src/access.js); this only asks and shows.

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Icon, Input, Select, Dialog, t } from '@rutba/office-ui';

const LEVELS = ['hidden', 'read', 'add', 'change', 'delete'];
const levelName = () => ({
  hidden: t('base.people.hidden'),
  read: t('common.read-only'),
  add: t('base.people.read-and-add'),
  change: t('base.people.read-add-change'),
  delete: t('base.people.all-deleting-too'),
});

/** Whether the person may change the design: nothing narrows what they see, and the database is not another program's (A2.20). */
export const canDesign = (db) => !db?.access?.rights && !db?.foreign;

/** Whether the person may `want` ('add', 'change', 'delete') in `table`. */
export function may(db, table, want) {
  const rights = db?.access?.rights;
  if (db?.access?.person?.seeing) return false;
  if (db?.foreign && !db.foreign.write) return false;
  if (!rights) return true;
  const right = rights.tables.each[table] || { access: rights.tables.default };
  return LEVELS.indexOf(right.access) >= LEVELS.indexOf(want);
}

/* ── signing in ────────────────────────────────────────────────────── */

/** The sign-in a protected database asks for as it opens. */
export function SignIn({ shell, db, onSigned }) {
  const names = db.access.names || [];
  const [name, setName] = useState(names[0] || '');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const go = async () => {
    if (!name || !password) return;
    setBusy(true);
    setError('');
    try {
      onSigned(await shell.base.signIn({ id: db.id, name, password }));
    } catch (e) {
      setError(e.message);
      setPassword('');
    } finally {
      setBusy(false);
    }
  };
  return (
    <div className="bs-signin" role="form" aria-label={t('base.people.sign-name', { name: db.name })}>
      <div className="bs-signin-card">
        <Icon name="lock" size={28} />
        <h1>{db.name}</h1>
        <p>{t('base.people.database-protected-sign')}</p>
        <label className="bs-field"><span>{t('base.people.who-you-are')}</span>
          <Select className="rw-select bs-signin-name" aria-label={t('base.people.who-you-are')} value={name} onChange={(e) => setName(e.target.value)}>
            {names.map((n) => <option key={n} value={n}>{n}</option>)}
          </Select>
        </label>
        <label className="bs-field"><span>{t('common.password')}</span>
          <Input className="rw-input bs-signin-password" type="password" aria-label={t('common.password')} autoFocus value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') go(); }} />
        </label>
        {error ? <p className="bs-signin-error" role="alert">{error}</p> : null}
        <Button primary label={t('common.sign-in')} className="bs-signin-go" disabled={busy || !name || !password} onClick={go} />
      </div>
    </div>
  );
}

/** Protect this database: the person protecting it becomes its first Admin. */
export function ProtectDialog({ shell, db, onDone, onClose, toast }) {
  const [name, setName] = useState('');
  const [password, setPassword] = useState('');
  const [again, setAgain] = useState('');
  const ready = name.trim() && password.length >= 4 && password === again;
  const go = async () => {
    try {
      onDone(await shell.base.protect({ id: db.id, name: name.trim(), password }));
      toast(t('base.people.database-protected-admin'), { tone: 'good', ms: 8000 });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  return (
    <Dialog title={t('base.protect-this')} onClose={onClose} actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.protect')} className="bs-protect-go" disabled={!ready} onClick={go} /></>}>
      <p className="bs-hint">{t('base.people.now-asks-opens')}</p>
      <label className="bs-field"><span>{t('common.your-name')}</span><Input className="rw-input bs-protect-name" aria-label={t('common.your-name')} autoFocus value={name} onChange={(e) => setName(e.target.value)} /></label>
      <label className="bs-field"><span>{t('common.password')}</span><Input className="rw-input bs-protect-password" type="password" aria-label={t('common.password')} value={password} onChange={(e) => setPassword(e.target.value)} /></label>
      <label className="bs-field"><span>{t('base.people.the-password-again')}</span><Input className="rw-input bs-protect-again" type="password" aria-label={t('base.people.the-password-again')} value={again} onChange={(e) => setAgain(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ready) go(); }} /></label>
      {again && password !== again ? <p className="bs-signin-error">{t('common.two-passwords-not')}</p> : null}
      <p className="bs-hint bs-protect-warn"><Icon name="info" size={14} /> {t('base.people.file-itself-not')}</p>
    </Dialog>
  );
}

/** Your own password, changed with the one you have now. */
export function PasswordDialog({ shell, db, onClose, toast }) {
  const [current, setCurrent] = useState('');
  const [password, setPassword] = useState('');
  const [again, setAgain] = useState('');
  const ready = current && password.length >= 4 && password === again;
  const go = async () => {
    try {
      await shell.base.setPassword({ id: db.id, person: db.access.person.id, password, current });
      toast(t('base.people.your-password-is-changed'), { tone: 'good' });
      onClose();
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  return (
    <Dialog title={t('base.change-my-password')} onClose={onClose} actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.change')} className="bs-password-go" disabled={!ready} onClick={go} /></>}>
      <label className="bs-field"><span>{t('base.people.your-password-now')}</span><Input className="rw-input bs-password-current" type="password" aria-label={t('base.people.your-password-now')} autoFocus value={current} onChange={(e) => setCurrent(e.target.value)} /></label>
      <label className="bs-field"><span>{t('base.people.a-new-password')}</span><Input className="rw-input bs-password-new" type="password" aria-label={t('base.people.a-new-password')} value={password} onChange={(e) => setPassword(e.target.value)} /></label>
      <label className="bs-field"><span>{t('base.people.new-password-again')}</span><Input className="rw-input bs-password-again" type="password" aria-label={t('base.people.new-password-again')} value={again} onChange={(e) => setAgain(e.target.value)} /></label>
    </Dialog>
  );
}

/** Shown while an Admin sees the database as a role would. */
export function SeeingBanner({ shell, db, onChanged }) {
  const seeing = db.access?.person?.seeing;
  if (!seeing) return null;
  return (
    <div className="bs-update bs-seeing" role="status" data-seeing={seeing}>
      <Icon name="eye" size={16} />
      <span>{t('base.people.seeing-database-role', { role: seeing })}</span>
      <Button primary label={t('base.people.back-own')} className="bs-seeing-back" onClick={async () => onChanged(await shell.base.seeAs({ id: db.id, role: null }))} />
    </div>
  );
}

/* ── People and roles, the Admin's ─────────────────────────────────── */

const KINDS = () => [
  { key: 'queries', title: t('base.queries'), names: (db) => db.queries || [] },
  { key: 'forms', title: t('base.forms'), names: (db) => db.forms || [] },
  { key: 'reports', title: t('base.reports'), names: (db) => db.reports || [] },
  { key: 'documents', title: t('base.print-templates'), names: (db) => (db.documents || []).map((d) => d.name) },
];

export function PeopleView({ shell, db, onChanged, toast }) {
  const [people, setPeople] = useState([]);
  const [roles, setRoles] = useState([]);
  const [fieldsOf, setFieldsOf] = useState({});
  const [role, setRole] = useState(null);
  const [open, setOpen] = useState(() => new Set());
  const [adding, setAdding] = useState(null);
  const [resetting, setResetting] = useState(null);
  const [newRole, setNewRole] = useState(null);
  const [unprotecting, setUnprotecting] = useState(null);

  const load = useCallback(async () => {
    try {
      const [p, r, f] = await Promise.all([shell.base.people({ id: db.id }), shell.base.roles({ id: db.id }), shell.base.allFields({ id: db.id })]);
      setPeople(p);
      setRoles(r);
      setFieldsOf(f);
      setRole((cur) => (r.some((x) => x.name === cur && !x.builtIn) ? cur : r.find((x) => !x.builtIn)?.name || null));
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  }, [shell, db.id, toast]);
  useEffect(() => { load(); }, [load, db.steps]);

  const act = async (fn, done = null) => {
    try { await fn(); done?.(); await load(); } catch (e) { toast(e.message, { tone: 'bad', ms: 7000 }); }
  };
  const editing = roles.find((r) => r.name === role) || null;
  const put = (rights) => act(() => shell.base.saveRole({ id: db.id, name: editing.name, rights }));
  const tableOf = (rights, table) => rights.tables.each[table] || { access: rights.tables.default, own: false, fields: {} };
  const setTable = (table, change) => {
    const r = editing.rights;
    put({ ...r, tables: { ...r.tables, each: { ...r.tables.each, [table]: { ...tableOf(r, table), ...change } } } });
  };
  const setField = (table, field, value) => {
    const cur = tableOf(editing.rights, table);
    const fields = { ...cur.fields };
    if (value) fields[field] = value; else delete fields[field];
    setTable(table, { fields });
  };
  const setShown = (kind, name, shown) => {
    const r = editing.rights;
    const list = r.hidden[kind].filter((n) => n !== name);
    if (!shown) list.push(name);
    put({ ...r, hidden: { ...r.hidden, [kind]: list } });
  };
  const levels = levelName();
  const me = db.access.person;

  return (
    <div className="bs-people">
      <section className="bs-people-list" aria-label={t('common.people')}>
        <header>
          <h2>{t('common.people')}</h2>
          <Button icon="plus" label={t('base.people.add-a-person')} className="bs-person-add" onClick={() => setAdding({ name: '', role: roles.find((r) => !r.builtIn)?.name || 'Admin', password: '' })} />
        </header>
        <table>
          <thead><tr><th>{t('common.name')}</th><th>{t('base.people.role')}</th><th /></tr></thead>
          <tbody>
            {people.map((p) => (
              <tr key={p.id} data-person={p.name}>
                <td>{p.name}{p.id === me.id ? <small> {t('base.people.you')}</small> : null}</td>
                <td>
                  <Select className="rw-select bs-person-role" aria-label={t('base.people.role-of-name', { name: p.name })} value={p.role} onChange={(e) => act(() => shell.base.changePerson({ id: db.id, person: p.id, role: e.target.value }))}>
                    {roles.map((r) => <option key={r.name} value={r.name}>{r.name}</option>)}
                  </Select>
                </td>
                <td>
                  <span className="bs-person-actions">
                    <Button label={t('common.new-password')} className="bs-person-reset" onClick={() => setResetting({ person: p, password: '' })} />
                    {p.id === me.id ? null : <Button icon="trash" label="" title={t('base.people.remove-name', { name: p.name })} className="bs-person-remove" onClick={() => act(() => shell.base.removePerson({ id: db.id, person: p.id }))} />}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="bs-people-foot">
          <Button ghost icon="unlock" label={t('base.people.take-protection-away')} className="bs-unprotect" onClick={() => setUnprotecting('')} />
        </div>
      </section>

      <section className="bs-roles" aria-label={t('base.people.roles')}>
        <header>
          <h2>{t('base.people.roles')}</h2>
          <div className="bs-role-tabs" role="tablist">
            {roles.filter((r) => !r.builtIn).map((r) => (
              <button key={r.name} type="button" role="tab" aria-selected={r.name === role} className={`bs-tab${r.name === role ? ' on' : ''}`} data-role={r.name} onClick={() => setRole(r.name)}>{r.name}</button>
            ))}
          </div>
          <Button icon="plus" label={t('base.people.new-role')} className="bs-role-new" onClick={() => setNewRole('')} />
        </header>
        <p className="bs-hint">{t('base.people.admin-may-do')}</p>
        {editing ? (
          <div className="bs-role" data-editing={editing.name}>
            <div className="bs-role-bar">
              <Button icon="eye" label={t('base.people.see-as-role', { role: editing.name })} className="bs-role-see" onClick={async () => { try { onChanged(await shell.base.seeAs({ id: db.id, role: editing.name })); } catch (e) { toast(e.message, { tone: 'bad' }); } }} />
              <Button ghost icon="trash" label={t('base.people.remove-this')} className="bs-role-remove" onClick={() => act(() => shell.base.removeRole({ id: db.id, name: editing.name }))} />
            </div>
            <table className="bs-role-tables">
              <thead><tr><th>{t('common.table')}</th><th>{t('base.people.may')}</th><th>{t('base.records')}</th><th /></tr></thead>
              <tbody>
                {db.tables.map((tb) => {
                  const right = tableOf(editing.rights, tb.name);
                  const isOpen = open.has(tb.name);
                  return (
                    <React.Fragment key={tb.name}>
                      <tr data-role-table={tb.name}>
                        <td>{tb.name}</td>
                        <td>
                          <Select className="rw-select bs-role-access" aria-label={t('base.people.role-may-do', { role: editing.name, table: tb.name })} value={right.access} onChange={(e) => setTable(tb.name, { access: e.target.value })}>
                            {LEVELS.map((l) => <option key={l} value={l}>{levels[l]}</option>)}
                          </Select>
                        </td>
                        <td>
                          <label className="bs-role-own"><input type="checkbox" checked={Boolean(right.own)} disabled={right.access === 'hidden'} onChange={(e) => setTable(tb.name, { own: e.target.checked })} /> {t('base.people.only-added')}</label>
                        </td>
                        <td>
                          <button type="button" className="bs-role-fields-toggle" aria-expanded={isOpen} disabled={right.access === 'hidden'} onClick={() => setOpen((s) => { const n = new Set(s); if (n.has(tb.name)) n.delete(tb.name); else n.add(tb.name); return n; })}>
                            <Icon name={isOpen ? 'chevronDown' : 'chevronRight'} size={12} /> {t('common.fields')}{Object.keys(right.fields).length ? ` (${Object.keys(right.fields).length})` : ''}
                          </button>
                        </td>
                      </tr>
                      {isOpen && right.access !== 'hidden' ? (fieldsOf[tb.name] || []).map((f) => (
                        <tr key={`${tb.name}.${f}`} className="bs-role-field" data-role-field={`${tb.name}.${f}`}>
                          <td>{f}</td>
                          <td colSpan={3}>
                            <Select className="rw-select bs-role-field-access" aria-label={t('base.people.field-for-role', { field: f, role: editing.name })} value={right.fields[f] || ''} onChange={(e) => setField(tb.name, f, e.target.value)}>
                              <option value="">{t('base.people.as-the-table')}</option>
                              <option value="read">{t('common.read-only')}</option>
                              <option value="hidden">{t('base.people.hidden')}</option>
                            </Select>
                          </td>
                        </tr>
                      )) : null}
                    </React.Fragment>
                  );
                })}
                <tr className="bs-role-default">
                  <td>{t('base.people.tables-made-now')}</td>
                  <td colSpan={3}>
                    <Select className="rw-select bs-role-default-access" aria-label={t('base.people.tables-made-now')} value={editing.rights.tables.default} onChange={(e) => put({ ...editing.rights, tables: { ...editing.rights.tables, default: e.target.value } })}>
                      {LEVELS.map((l) => <option key={l} value={l}>{levels[l]}</option>)}
                    </Select>
                  </td>
                </tr>
              </tbody>
            </table>
            <div className="bs-role-items">
              {KINDS().map((k) => (k.names(db).length ? (
                <fieldset key={k.key} data-role-kind={k.key}>
                  <legend>{k.title}</legend>
                  {k.names(db).map((n) => (
                    <label key={n} className="bs-role-item"><input type="checkbox" data-role-item={`${k.key}:${n}`} checked={!editing.rights.hidden[k.key].includes(n)} onChange={(e) => setShown(k.key, n, e.target.checked)} /> {n}</label>
                  ))}
                </fieldset>
              ) : null))}
              <p className="bs-hint">{t('base.people.ticked-ones-shown')}</p>
            </div>
          </div>
        ) : <p className="bs-hint">{t('base.people.make-role-each')}</p>}
      </section>

      {adding ? (
        <Dialog title={t('base.people.add-a-person')} onClose={() => setAdding(null)} actions={<><Button label={t('common.cancel')} onClick={() => setAdding(null)} /><Button primary label={t('common.add')} className="bs-person-add-go" disabled={!adding.name.trim() || adding.password.length < 4} onClick={() => act(() => shell.base.addPerson({ id: db.id, ...adding, name: adding.name.trim() }), () => setAdding(null))} /></>}>
          <label className="bs-field"><span>{t('common.name')}</span><Input className="rw-input bs-person-name" aria-label={t('common.name')} autoFocus value={adding.name} onChange={(e) => setAdding({ ...adding, name: e.target.value })} /></label>
          <label className="bs-field"><span>{t('base.people.role')}</span>
            <Select className="rw-select bs-person-new-role" aria-label={t('base.people.role')} value={adding.role} onChange={(e) => setAdding({ ...adding, role: e.target.value })}>
              {roles.map((r) => <option key={r.name} value={r.name}>{r.name}</option>)}
            </Select>
          </label>
          <label className="bs-field"><span>{t('base.people.their-password')}</span><Input className="rw-input bs-person-password" type="password" aria-label={t('base.people.their-password')} value={adding.password} onChange={(e) => setAdding({ ...adding, password: e.target.value })} /></label>
          <p className="bs-hint">{t('base.people.tell-password-change')}</p>
        </Dialog>
      ) : null}
      {resetting ? (
        <Dialog title={t('base.people.new-password-name', { name: resetting.person.name })} onClose={() => setResetting(null)} actions={<><Button label={t('common.cancel')} onClick={() => setResetting(null)} /><Button primary label={t('common.change')} className="bs-person-reset-go" disabled={resetting.password.length < 4} onClick={() => act(() => shell.base.setPassword({ id: db.id, person: resetting.person.id, password: resetting.password, current: null }), () => { setResetting(null); toast(t('base.people.password-changed'), { tone: 'good' }); })} /></>}>
          <label className="bs-field"><span>{t('base.people.a-new-password')}</span><Input className="rw-input bs-person-reset-password" type="password" aria-label={t('base.people.a-new-password')} autoFocus value={resetting.password} onChange={(e) => setResetting({ ...resetting, password: e.target.value })} /></label>
        </Dialog>
      ) : null}
      {newRole !== null ? (
        <Dialog title={t('base.people.new-role')} onClose={() => setNewRole(null)} actions={<><Button label={t('common.cancel')} onClick={() => setNewRole(null)} /><Button primary label={t('base.make')} className="bs-role-new-go" disabled={!newRole.trim()} onClick={() => act(() => shell.base.saveRole({ id: db.id, name: newRole.trim(), rights: { tables: { default: 'read' } } }), () => { setRole(newRole.trim()); setNewRole(null); })} /></>}>
          <label className="bs-field"><span>{t('common.name')}</span><Input className="rw-input bs-role-name" aria-label={t('common.name')} autoFocus value={newRole} onChange={(e) => setNewRole(e.target.value)} /></label>
          <p className="bs-hint">{t('base.people.starts-reading-every')}</p>
        </Dialog>
      ) : null}
      {unprotecting !== null ? (
        <Dialog title={t('base.people.take-protection-away')} onClose={() => setUnprotecting(null)} actions={<><Button label={t('common.cancel')} onClick={() => setUnprotecting(null)} /><Button primary label={t('base.people.take-it-away')} className="bs-unprotect-go" disabled={!unprotecting} onClick={async () => {
          try { onChanged(await shell.base.unprotect({ id: db.id, password: unprotecting })); toast(t('base.people.database-no-longer'), { tone: 'good', ms: 7000 }); } catch (e) { toast(e.message, { tone: 'bad', ms: 7000 }); }
        }} /></>}>
          <p className="bs-hint">{t('base.people.open-anyone-again')}</p>
          <label className="bs-field"><span>{t('common.password')}</span><Input className="rw-input bs-unprotect-password" type="password" aria-label={t('common.password')} autoFocus value={unprotecting} onChange={(e) => setUnprotecting(e.target.value)} /></label>
        </Dialog>
      ) : null}
    </div>
  );
}

export const PEOPLE_CSS = `
.bs-signin { flex: 1; display: flex; align-items: center; justify-content: center; padding: 24px; background: var(--canvas); }
.bs-signin-card { width: min(360px, 100%); display: flex; flex-direction: column; gap: 12px; padding: 28px; background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--r-3, 8px); box-shadow: var(--shadow-1); }
.bs-signin-card h1 { margin: 0; font-size: 18px; }
.bs-signin-card p { margin: 0; color: var(--ink-2); }
.bs-signin-card > svg { color: var(--accent); }
.bs-signin-error { color: var(--bad, #b42318); margin: 0; }
.bs-protect-warn { display: flex; gap: 6px; align-items: flex-start; }
.bs-people { flex: 1; min-height: 0; overflow: auto; display: grid; grid-template-columns: minmax(280px, 1fr) minmax(420px, 2fr); gap: 20px; padding: 16px; align-items: start; }
.bs-people section { background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--r-3, 8px); padding: 14px; }
.bs-people header { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 8px; }
.bs-people header h2 { margin: 0; font-size: 15px; flex: none; }
.bs-people table { width: 100%; border-collapse: collapse; }
.bs-people th { text-align: left; font-weight: 600; color: var(--ink-2); font-size: 12px; padding: 4px 6px; }
.bs-people td { padding: 4px 6px; border-top: 1px solid var(--line-soft); vertical-align: middle; }
.bs-person-actions { display: flex; gap: 4px; justify-content: flex-end; align-items: center; white-space: nowrap; }
.bs-people-foot { margin-top: 12px; }
.bs-role-tabs { display: flex; gap: 4px; flex-wrap: wrap; flex: 1; }
.bs-role-bar { display: flex; gap: 8px; margin: 6px 0 10px; }
.bs-role-field td:first-child { padding-left: 24px; color: var(--ink-2); }
.bs-role-default td { color: var(--ink-2); }
.bs-role-own { display: inline-flex; gap: 6px; align-items: center; white-space: nowrap; }
.bs-role-fields-toggle { background: none; border: 0; color: var(--accent); cursor: pointer; display: inline-flex; align-items: center; gap: 4px; }
.bs-role-fields-toggle:disabled { color: var(--ink-3); cursor: default; }
.bs-role-items { display: grid; grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); gap: 10px; margin-top: 14px; }
.bs-role-items fieldset { border: 1px solid var(--line-soft); border-radius: var(--r-2, 6px); padding: 8px 10px; margin: 0; }
.bs-role-items legend { font-weight: 600; font-size: 12px; padding: 0 4px; }
.bs-role-item { display: flex; gap: 6px; align-items: center; padding: 2px 0; }
.bs-role-items > .bs-hint { grid-column: 1 / -1; }
.bs-me { display: flex; align-items: center; gap: 6px; padding: 6px 8px; color: var(--ink-2); font-size: 12px; }
.bs-me strong { color: var(--ink-1); font-weight: 600; }
@media (max-width: 900px) { .bs-people { grid-template-columns: 1fr; } }
`;
