// Tables linked to a database engine (D-2026-10-22-13): the box that
// connects to an engine and links its tables, the bar above a linked table
// that says where its records come from and reads them again, and the box
// that asks for a link's password on a computer that has not kept it. The
// engine holds the records; main/base.js reads and writes them.

import React, { useEffect, useState } from 'react';
import { Button, Icon, Input, Select, Dialog, t, tn } from '@rutba/office-ui';

const blankSettings = (port = 5432) => ({ host: 'localhost', port, database: '', user: '', ssl: false });

/** Link to a database engine: an engine and its account, then which of its tables to link. */
export function LinkDialog({ shell, db, onDone, onClose, toast }) {
  const [engines, setEngines] = useState([]);
  const [linkId, setLinkId] = useState(db.links?.[0]?.id || '');
  const [engine, setEngine] = useState('postgres');
  const [name, setName] = useState('');
  const [settings, setSettings] = useState(blankSettings());
  const [password, setPassword] = useState('');
  const [found, setFound] = useState(null);
  const [picked, setPicked] = useState(() => new Set());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  useEffect(() => { shell.base.engines().then(setEngines).catch(() => {}); }, [shell]);

  const existing = (db.links || []).find((l) => l.id === linkId) || null;
  const set = (k, v) => setSettings((s) => ({ ...s, [k]: v }));
  const local = new Set(db.tables.map((x) => x.name.toLowerCase()));
  const nameOf = (r) => (r.schema && r.schema !== 'public' ? `${r.schema} ${r.name}` : r.name);

  const connect = async () => {
    setBusy(true);
    setError('');
    try {
      if (existing) {
        // An existing link: its tables, through the password this computer keeps.
        const r = await shell.base.engineTables({ engine: existing.engine, settings: existing.settings, password: null, link: existing.id, id: db.id });
        setFound(r.tables);
      } else {
        const r = await shell.base.engineTables({ engine, settings, password });
        setFound(r.tables);
        if (!name.trim()) setName(settings.database || engines.find((e) => e.id === engine)?.name || '');
      }
      setPicked(new Set());
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  };
  const link = async () => {
    setBusy(true);
    setError('');
    try {
      let id = existing?.id;
      if (!id) id = (await shell.base.saveLink({ id: db.id, name: name.trim(), engine, settings, password })).link;
      const remotes = found.filter((r) => picked.has(`${r.schema}.${r.name}`)).map((r) => ({ schema: r.schema, name: r.name, as: nameOf(r) }));
      const out = await shell.base.linkTables({ id: db.id, link: id, remotes });
      onDone(out.summary, out.tables[0]);
      toast(tn(out.tables.length, '{count} table linked: its records stay in the engine, read from it here.', '{count} tables linked: their records stay in the engine, read from it here.'), { tone: 'good', ms: 6000 });
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  };

  const ready = existing || (settings.host && settings.database && settings.user);
  return (
    <Dialog title={t('Link to a database engine')} onClose={onClose} actions={(
      <>
        <Button label={t('Cancel')} onClick={onClose} />
        {found
          ? <Button primary label={t('Link')} className="bs-link-go" disabled={busy || !picked.size || (!existing && !name.trim())} onClick={link} />
          : <Button primary label={t('Connect')} className="bs-link-connect" disabled={busy || !ready} onClick={connect} />}
      </>
    )}>
      <p className="bs-hint">{t('The engine keeps the records: Data Files reads them from it and saves each change straight to it. Every computer linked to the same engine works on the same records.')}</p>
      {(db.links || []).length ? (
        <label className="bs-field"><span>{t('Engine')}</span>
          <Select className="rw-select bs-link-which" aria-label={t('Engine')} value={linkId} onChange={(e) => { setLinkId(e.target.value); setFound(null); }}>
            {db.links.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
            <option value="">{t('Another engine…')}</option>
          </Select>
        </label>
      ) : null}
      {existing ? null : (
        <div className="bs-link-form">
          <label className="bs-field"><span>{t('Kind')}</span>
            <Select className="rw-select bs-link-engine" aria-label={t('Kind')} value={engine} onChange={(e) => { setEngine(e.target.value); set('port', engines.find((x) => x.id === e.target.value)?.port || 5432); setFound(null); }}>
              {engines.map((e) => <option key={e.id} value={e.id} disabled={!e.ready}>{e.ready ? e.name : t('{name} (coming)', { name: e.name })}</option>)}
            </Select>
          </label>
          <label className="bs-field bs-link-host"><span>{t('Address')}</span><Input className="rw-input bs-link-host-input" aria-label={t('Address')} value={settings.host} onChange={(e) => set('host', e.target.value)} /></label>
          <label className="bs-field bs-link-port"><span>{t('Port')}</span><Input className="rw-input bs-link-port-input" aria-label={t('Port')} value={String(settings.port || '')} onChange={(e) => set('port', e.target.value.replace(/\D/g, ''))} /></label>
          <label className="bs-field"><span>{t('Database')}</span><Input className="rw-input bs-link-database" aria-label={t('Database')} value={settings.database} onChange={(e) => set('database', e.target.value)} /></label>
          <label className="bs-field"><span>{t('Account')}</span><Input className="rw-input bs-link-user" aria-label={t('Account')} value={settings.user} onChange={(e) => set('user', e.target.value)} /></label>
          <label className="bs-field"><span>{t('Password')}</span><Input className="rw-input bs-link-password" type="password" aria-label={t('Password')} value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ready && !found) connect(); }} /></label>
          <label className="bs-link-ssl"><input type="checkbox" checked={settings.ssl} onChange={(e) => set('ssl', e.target.checked)} /> {t('An encrypted connection (SSL)')}</label>
          {found ? <label className="bs-field"><span>{t('Name of this link')}</span><Input className="rw-input bs-link-name" aria-label={t('Name of this link')} value={name} onChange={(e) => setName(e.target.value)} /></label> : null}
        </div>
      )}
      {error ? <p className="bs-signin-error" role="alert">{error}</p> : null}
      {found ? (
        <div className="bs-link-tables" role="group" aria-label={t('Tables of the engine')}>
          <h3>{t('Its tables')}</h3>
          {found.length ? found.map((r) => {
            const k = `${r.schema}.${r.name}`;
            const clash = local.has(nameOf(r).toLowerCase());
            return (
              <label key={k} className={`bs-link-table${clash ? ' clash' : ''}`} data-remote={k}>
                <input type="checkbox" disabled={clash} checked={picked.has(k)} onChange={(e) => setPicked((p) => { const n = new Set(p); if (e.target.checked) n.add(k); else n.delete(k); return n; })} />
                <Icon name={r.kind === 'view' ? 'eye' : 'sheets'} size={13} />
                <span>{nameOf(r)}</span>
                {clash ? <small>{t('a table here has that name')}</small> : null}
              </label>
            );
          }) : <p className="bs-hint">{t('This database has no tables this account can see.')}</p>}
        </div>
      ) : null}
      <p className="bs-hint">{t('The password is kept in this computer\'s protected store, never in the file. On another computer it is asked for once.')}</p>
    </Dialog>
  );
}

/** The bar above a linked table: where its records come from, when they were read, and Read again. */
export function LinkedBar({ shell, db, table, onChanged, onPassword, toast }) {
  const tb = db.tables.find((x) => x.name === table);
  if (!tb?.linked) return null;
  const link = (db.links || []).find((l) => l.id === tb.linked);
  const again = async () => {
    try {
      const s = await shell.base.refreshLinked({ id: db.id, table });
      onChanged(s);
      const err = s.tables.find((x) => x.name === table)?.linkError;
      if (err) toast(err, { tone: 'bad', ms: 7000 });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  return (
    <div className={`bs-linked-bar${tb.linkError ? ' bad' : ''}`} role="status" data-linked={table}>
      <Icon name={tb.linkError ? 'info' : 'link'} size={14} />
      {tb.linkError
        ? <span>{t('The records of {table} could not be read from {link}: {why}', { table, link: link?.name || '', why: tb.linkError })}</span>
        : <span>{t('Linked to {link}: the records are the engine\'s, read from it here.', { link: link?.name || '' })}{tb.key ? null : ` ${t('With no primary key there, they can be read, not changed.')}`}</span>}
      <span className="grow" />
      {tb.linkError && link && !link.password ? <Button label={t('Password…')} className="bs-linked-password" onClick={() => onPassword(link)} /> : null}
      <Button icon="refresh" label={t('Read again')} className="bs-linked-again" onClick={again} />
    </div>
  );
}

/** A link's password, asked for on a computer that has not kept it. */
export function LinkPasswordDialog({ shell, db, link, onDone, onClose, toast }) {
  const [password, setPassword] = useState('');
  const go = async () => {
    try {
      const r = await shell.base.saveLink({ id: db.id, link: link.id, name: link.name, engine: link.engine, settings: link.settings, password });
      onDone(r.summary);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  return (
    <Dialog title={t('Password for {link}', { link: link.name })} onClose={onClose} actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('Connect')} className="bs-link-password-go" disabled={!password} onClick={go} /></>}>
      <p className="bs-hint">{t('{user} at {host}, database {database}.', { user: link.settings.user, host: link.settings.host, database: link.settings.database })}</p>
      <label className="bs-field"><span>{t('Password')}</span><Input className="rw-input bs-link-password-input" type="password" aria-label={t('Password')} autoFocus value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && password) go(); }} /></label>
    </Dialog>
  );
}

export const LINKS_CSS = `
.bs-link-form { display: grid; grid-template-columns: 1fr 1fr; gap: 8px 12px; }
.bs-link-form .bs-link-host { grid-column: 1; }
.bs-link-ssl { grid-column: 1 / -1; display: flex; gap: 6px; align-items: center; }
.bs-link-tables { margin-top: 10px; max-height: 260px; overflow: auto; border: 1px solid var(--line-soft); border-radius: var(--r-2, 6px); padding: 6px 8px; }
.bs-link-tables h3 { margin: 2px 0 6px; font-size: 12px; color: var(--ink-2); }
.bs-link-table { display: flex; gap: 6px; align-items: center; padding: 3px 0; }
.bs-link-table.clash { color: var(--ink-3); }
.bs-link-table small { color: var(--ink-3); }
.bs-linked-bar { display: flex; align-items: center; gap: 8px; padding: 6px 12px; background: var(--accent-soft, rgba(0,0,0,.04)); border-bottom: 1px solid var(--line-soft); font-size: 12px; }
.bs-linked-bar.bad { background: rgba(180, 35, 24, .08); }
.bs-linked-bar .grow { flex: 1; }
.bs-table .bs-linked-mark { color: var(--accent); flex: none; }
.bs-table .bs-linked-mark.bad { color: var(--bad, #b42318); }
`;
