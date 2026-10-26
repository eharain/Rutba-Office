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
  // Each table read live (as 2.17) or kept as a copy in this file (D-2026-10-22-13, step 4).
  const [modes, setModes] = useState(() => ({}));
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
      const remotes = found.filter((r) => picked.has(`${r.schema}.${r.name}`)).map((r) => ({ schema: r.schema, name: r.name, as: nameOf(r), mode: modes[`${r.schema}.${r.name}`] === 'copy' ? 'copy' : 'live' }));
      const out = await shell.base.linkTables({ id: db.id, link: id, remotes });
      onDone(out.summary, out.tables[0]);
      const copies = remotes.filter((r) => r.mode === 'copy').length;
      toast(copies === remotes.length
        ? tn(out.tables.length, '{count} table linked and kept as a copy in this file: it works with the engine out of reach and catches up when it is back.', '{count} tables linked and kept as copies in this file: they work with the engine out of reach and catch up when it is back.')
        : tn(out.tables.length, '{count} table linked: its records stay in the engine, read from it here.', '{count} tables linked: their records stay in the engine, read from it here.'), { tone: 'good', ms: 6000 });
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
                {picked.has(k) ? (
                  <Select className="rw-select bs-link-mode" data-mode-for={k} aria-label={t('How {table} is kept', { table: nameOf(r) })} value={modes[k] || 'live'}
                    onChange={(e) => setModes((m) => ({ ...m, [k]: e.target.value }))}>
                    <option value="live">{t('Live')}</option>
                    <option value="copy">{t('Keep a copy in this file')}</option>
                  </Select>
                ) : null}
              </label>
            );
          }) : <p className="bs-hint">{t('This database has no tables this account can see.')}</p>}
        </div>
      ) : null}
      {found ? <p className="bs-hint">{t('Live reads the records from the engine as you work. Keep a copy in this file keeps them here too: you can work with the engine out of reach, and your changes go to it when it is back. When a record was changed in both places, the engine\'s version wins and yours is kept for you to see.')}</p> : null}
      <p className="bs-hint">{t('The password is kept in this computer\'s protected store, never in the file. On another computer it is asked for once.')}</p>
    </Dialog>
  );
}

/** A time a copy caught up, as a person reads it: the time today, the day and time before. */
export function whenWords(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  const today = new Date().toDateString() === d.toDateString();
  return today ? d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }) : d.toLocaleString(undefined, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
}

/** How often a source is read again, in words. */
export const everyWords = (every) => ({ 5: t('every 5 minutes'), 15: t('every 15 minutes'), 60: t('every hour'), 0: t('by hand') }[Number(every)] ?? t('every 15 minutes'));

/**
 * The bar above a linked table: where its records come from and when they
 * were read. Read live: Read again. Kept as a copy (D-2026-10-22-13, step
 * 4): when it caught up, how many changes wait and since when, the versions
 * the engine replaced, and Sync now. Read from a source (step 5): when it
 * was read, how often, a change to its columns, and Refresh. Not allowed on
 * this computer (A2.18): not read, the copy as it was, and Refresh asks.
 */
export function LinkedBar({ shell, db, table, design = false, onChanged, onPassword, onClashes, onSourceSignIn, onSourceKeys, onAsk, toast }) {
  const tb = db.tables.find((x) => x.name === table);
  const [fields, setFields] = useState([]);
  const copy = tb?.mode === 'copy';
  useEffect(() => {
    if (!copy || tb?.source || !design) return;
    shell.base.fields({ id: db.id, table }).then(setFields).catch(() => {});
  }, [shell, db.id, table, copy, tb?.source, design, db.steps]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!tb?.linked) return null;
  const link = (db.links || []).find((l) => l.id === tb.linked);
  const source = tb.source ? (db.sources || []).find((l) => l.id === tb.linked) : null;
  const run = async (call, say = true) => {
    try {
      const s = await call();
      onChanged(s);
      const err = s.tables.find((x) => x.name === table)?.linkError;
      if (err && say) toast(err, { tone: 'bad', ms: 7000 });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  const again = () => run(() => shell.base.refreshLinked({ id: db.id, table, now: true }));
  const modeSelect = design && !tb.source ? (
    <Select className="rw-select bs-linked-mode" aria-label={t('How {table} is kept', { table })} value={copy ? 'copy' : 'live'}
      onChange={(e) => run(() => shell.base.setLinkMode({ id: db.id, table, mode: e.target.value }), false)}>
      <option value="live">{t('Live')}</option>
      <option value="copy">{t('Keep a copy in this file')}</option>
    </Select>
  ) : null;

  if (tb.blocked) {
    const where = source?.settings?.where || '';
    const from = source ? (/^https?:/i.test(where) ? (() => { try { return new URL(where).host; } catch { return where; } })() : where.split(/[\\/]/).pop()) : link?.name || table;
    // Only someone who may change the design is asked; anyone else is told who can.
    const ask = () => (db.ask?.length ? onAsk?.() : toast(t('Someone whose role may change the design allows this database\'s sources on this computer.'), { ms: 6000 }));
    return (
      <div className="bs-linked-bar blocked" role="status" data-linked={table} data-blocked="1" data-source={tb.source || undefined} data-mode={tb.source ? undefined : copy ? 'copy' : 'live'}>
        <Icon name="shield" size={14} />
        <span className="bs-linked-words">
          {copy
            ? t('Not read: {source} waits for you to allow this database\'s sources on this computer. The copy here is as it was at {time}.', { source: from, time: whenWords(tb.synced) || t('first') })
            : t('Not read: {source} waits for you to allow this database\'s sources on this computer.', { source: from })}
        </span>
        <span className="grow" />
        <Button icon="refresh" label={tb.source ? t('Refresh') : copy ? t('Sync now') : t('Read again')} className={tb.source ? 'bs-source-refresh' : copy ? 'bs-linked-sync' : 'bs-linked-again'} onClick={ask} />
      </div>
    );
  }

  if (tb.source) {
    const where = source?.settings?.where || '';
    const shown = /^https?:/i.test(where) ? (() => { try { return new URL(where).host; } catch { return where; } })() : where.split(/[\\/]/).pop();
    const note = tb.note && (tb.note.added.length || tb.note.removed.length || tb.note.changed.length) ? [
      tb.note.added.length ? t('added {fields}', { fields: tb.note.added.join(', ') }) : null,
      tb.note.removed.length ? t('taken away {fields}', { fields: tb.note.removed.join(', ') }) : null,
      tb.note.changed.length ? t('changed {fields}', { fields: tb.note.changed.join(', ') }) : null,
    ].filter(Boolean).join('; ') : '';
    return (
      <div className={`bs-linked-bar${tb.linkError ? ' bad' : ''}`} role="status" data-linked={table} data-source={tb.source}>
        <Icon name={tb.linkError ? 'info' : 'import'} size={14} />
        <span className="bs-linked-words">
          {tb.linkError
            ? t('{source} could not be read: {why} The copy here is as it was at {time}.', { source: shown, why: tb.linkError, time: whenWords(tb.synced) || t('first') })
            : t('From {source}: read at {time}, {every}. Read only here.', { source: shown, time: whenWords(tb.synced), every: everyWords(tb.every) })}
          {note ? <span className="bs-source-note" data-note={table}> {t('Its columns changed at the source: {changes}. The design here follows.', { changes: note })}</span> : null}
        </span>
        <span className="grow" />
        {tb.who && design ? <Button label={t('Sign in there…')} className="bs-source-who" onClick={() => onSourceSignIn?.(table)} /> : null}
        {tb.token && design ? <Button label={t('Key or token…')} className="bs-source-keys" onClick={() => onSourceKeys?.(table)} /> : null}
        {design ? (
          <Select className="rw-select bs-source-every" aria-label={t('How often {table} is read', { table })} value={String(tb.every ?? 15)}
            onChange={(e) => run(() => Promise.resolve(shell.base.setSourceEvery({ id: db.id, table, every: Number(e.target.value) })), false)}>
            {[5, 15, 60, 0].map((n) => <option key={n} value={String(n)}>{n ? t('Read {every}', { every: everyWords(n) }) : t('Read by hand')}</option>)}
          </Select>
        ) : null}
        <Button icon="refresh" label={t('Refresh')} className="bs-source-refresh" onClick={again} />
      </div>
    );
  }

  if (copy) {
    const wait = tb.waiting?.count || 0;
    const words = tb.linkError && wait
      ? tn(wait, '{count} change waits to go to {link}, kept since {time}. {link} cannot be reached now: it goes when it is back.', '{count} changes wait to go to {link}, kept since {time}. {link} cannot be reached now: they go when it is back.', { link: link?.name || '', time: whenWords(tb.waiting.since) })
      : tb.linkError && tb.out
        ? t('{link} cannot be reached now. The copy here is as it was at {time}, and changes made here wait for it.', { link: link?.name || '', time: whenWords(tb.synced) || t('first') })
        : tb.linkError
          ? t('The records of {table} could not be synced with {link}: {why}', { table, link: link?.name || '', why: tb.linkError })
          : wait
            ? tn(wait, '{count} change made here goes to {link} at the next catch-up.', '{count} changes made here go to {link} at the next catch-up.', { link: link?.name || '' })
            : t('Kept as a copy of {link} in this file: caught up at {time}.', { link: link?.name || '', time: whenWords(tb.synced) });
    const stamps = fields.filter((f) => ['datetime', 'date', 'integer', 'number'].includes(f.type) && f.name !== tb.key);
    return (
      <div className={`bs-linked-bar${tb.linkError ? ' bad' : ''}${wait ? ' waiting' : ''}`} role="status" data-linked={table} data-mode="copy" data-waiting={wait}>
        <Icon name={tb.linkError ? 'info' : 'link'} size={14} />
        <span className="bs-linked-words">{words}{tb.key ? null : ` ${t('With no primary key there, they can be read, not changed.')}`}</span>
        {tb.clashes ? (
          <span className="bs-linked-clash" data-clashes={tb.clashes}>
            {tn(tb.clashes, '{count} of your changes was replaced by a newer one in the engine:', '{count} of your changes were replaced by newer ones in the engine:')}
            {' '}<button type="button" className="bs-linked-clashes" onClick={() => onClashes?.(table)}>{t('see them')}</button>
          </span>
        ) : null}
        <span className="grow" />
        {tb.linkError && link && !link.password ? <Button label={t('Password…')} className="bs-linked-password" onClick={() => onPassword(link)} /> : null}
        {design && stamps.length ? (
          <Select className="rw-select bs-linked-stamp" aria-label={t('What {table} reads at each catch-up', { table })} value={tb.stamp || ''}
            data-tip={t('For a large table: a field the engine changes with every change, so only what changed is read.')}
            onChange={(e) => run(() => Promise.resolve(shell.base.setLinkStamp({ id: db.id, table, field: e.target.value || null })), false)}>
            <option value="">{t('Read all of it')}</option>
            {stamps.map((f) => <option key={f.name} value={f.name}>{t('Read what changed by {field}', { field: f.name })}</option>)}
          </Select>
        ) : null}
        {modeSelect}
        <Button icon="refresh" label={t('Sync now')} className="bs-linked-sync" onClick={again} />
      </div>
    );
  }

  return (
    <div className={`bs-linked-bar${tb.linkError ? ' bad' : ''}`} role="status" data-linked={table} data-mode="live">
      <Icon name={tb.linkError ? 'info' : 'link'} size={14} />
      {tb.linkError
        ? <span>{t('The records of {table} could not be read from {link}: {why}', { table, link: link?.name || '', why: tb.linkError })}</span>
        : <span>{t('Linked to {link}: the records are the engine\'s, read from it here.', { link: link?.name || '' })}{tb.key ? null : ` ${t('With no primary key there, they can be read, not changed.')}`}</span>}
      <span className="grow" />
      {tb.linkError && link && !link.password ? <Button label={t('Password…')} className="bs-linked-password" onClick={() => onPassword(link)} /> : null}
      {modeSelect}
      <Button icon="refresh" label={t('Read again')} className="bs-linked-again" onClick={again} />
    </div>
  );
}

/** The versions from here a copy's engine replaced: each shown in words, kept as the engine has it or put back. */
export function ClashesDialog({ shell, db, table, onChanged, onClose, toast }) {
  const [list, setList] = useState(null);
  const load = () => shell.base.clashes({ id: db.id, table }).then(setList).catch((e) => toast(e.message, { tone: 'bad' }));
  useEffect(() => { load(); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const settle = async (c, keep) => {
    try {
      onChanged(await shell.base.settleClash({ id: db.id, table, clash: c.id, keep }));
      await load();
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  const said = (c) => ({
    changed: t('You changed it here, and it was changed in the engine too. The engine\'s version is kept.'),
    gone: t('You changed it here, and it was deleted in the engine. It is gone here too.'),
    deleted: t('You deleted it here, and it was changed in the engine. It is back.'),
    refused: t('The engine did not take it: {why}', { why: c.why || '' }),
  }[c.action] || '');
  const differs = (c) => Object.keys(c.mine || {}).filter((k) => !c.theirs || String(c.mine[k] ?? '') !== String(c.theirs[k] ?? ''));
  return (
    <Dialog title={t('Changes the engine replaced in {table}', { table })} onClose={onClose} width={620} actions={<Button primary label={t('Close')} onClick={onClose} />}>
      {list == null ? null : list.length ? (
        <ul className="bs-clash-list">
          {list.map((c) => (
            <li key={c.id} className="bs-clash" data-clash={c.id} data-action={c.action}>
              <p><strong>{c.key == null ? t('A new record') : t('Record {key}', { key: String(c.key) })}</strong> <small>{whenWords(c.at)}</small></p>
              <p>{said(c)}</p>
              {c.mine && differs(c).length ? (
                <table className="bs-clash-values">
                  <thead><tr><th>{t('Field')}</th><th>{t('Yours')}</th>{c.theirs ? <th>{t('The engine\'s')}</th> : null}</tr></thead>
                  <tbody>{differs(c).map((k) => <tr key={k}><td>{k}</td><td>{String(c.mine[k] ?? '')}</td>{c.theirs ? <td>{String(c.theirs[k] ?? '')}</td> : null}</tr>)}</tbody>
                </table>
              ) : null}
              <p className="bs-clash-actions">
                <Button label={t('Keep the engine\'s')} className="bs-clash-keep" onClick={() => settle(c, 'engine')} />
                {c.mine || c.action === 'deleted' ? <Button label={t('Put mine back')} className="bs-clash-mine" onClick={() => settle(c, 'mine')} /> : null}
              </p>
            </li>
          ))}
        </ul>
      ) : <p className="bs-hint">{t('Nothing of yours was replaced.')}</p>}
    </Dialog>
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
.bs-linked-bar.waiting:not(.bad) { background: rgba(181, 120, 0, .08); }
.bs-linked-bar .rw-select { height: 26px; font-size: 12px; }
.bs-linked-words { min-width: 0; }
.bs-linked-clash { color: var(--ink-2); }
.bs-linked-clashes { background: none; border: 0; padding: 0; color: var(--accent); text-decoration: underline; cursor: pointer; font: inherit; }
.bs-link-table .bs-link-mode { margin-left: auto; height: 24px; font-size: 12px; }
.bs-clash-list { list-style: none; margin: 0; padding: 0; max-height: 420px; overflow: auto; }
.bs-clash { border-bottom: 1px solid var(--line-soft); padding: 8px 0; }
.bs-clash p { margin: 0 0 6px; }
.bs-clash small { color: var(--ink-3); }
.bs-clash-values { border-collapse: collapse; font-size: 12px; margin-bottom: 6px; }
.bs-clash-values th, .bs-clash-values td { text-align: left; padding: 2px 10px 2px 0; }
.bs-clash-actions { display: flex; gap: 8px; }
.bs-table .bs-linked-mark { color: var(--accent); flex: none; }
.bs-table .bs-linked-mark.bad { color: var(--bad, #b42318); }
`;
