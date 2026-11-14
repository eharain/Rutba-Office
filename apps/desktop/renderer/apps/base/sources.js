// Other sources a table is read from (D-2026-10-22-13, step 5): the box that
// links a CSV or spreadsheet file (at a path, a network share or a web
// address), a table or query of another Data Files database, or a web
// address answering JSON. Each is tried first, its fields and first records
// shown, then linked as a table kept as a copy in this file, read only and
// read again on its schedule. Keys, tokens and passwords go to this
// computer's protected store, never the file; main/base-sync.js keeps them.

import React, { useState } from 'react';
import { Button, Icon, Input, Select, Dialog, t, tn } from '@rutba/office-ui';
import { everyWords } from './links.js';

const FILES = ['csv', 'tsv', 'txt', 'xlsx', 'xlsm', 'xls', 'ods'];
const isWeb = (s) => /^https?:\/\//i.test(String(s || '').trim());
const stem = (where) => {
  const s = String(where || '').trim();
  if (isWeb(s)) { try { const u = new URL(s); return (u.pathname.split('/').filter(Boolean).pop() || u.host).replace(/\.[^.]+$/, ''); } catch { return ''; } }
  return s.split(/[\\/]/).pop().replace(/\.[^.]+$/, '');
};

/** Link to a file or web address: a source tried, then linked as a table of its own. */
export function SourceDialog({ shell, db, onDone, onClose, toast }) {
  const [kind, setKind] = useState('file');
  const [where, setWhere] = useState('');
  const [header, setHeader] = useState('1');
  const [sheet, setSheet] = useState(null);
  const [person, setPerson] = useState('');
  const [password, setPassword] = useState('');
  const [needWho, setNeedWho] = useState(false);
  const [item, setItem] = useState(null);
  const [headers, setHeaders] = useState([{ name: '', value: '', secret: true }]);
  const [records, setRecords] = useState('');
  const [paging, setPaging] = useState('none');
  const [next, setNext] = useState('');
  const [param, setParam] = useState('page');
  const [start, setStart] = useState('1');
  const [every, setEvery] = useState('15');
  const [name, setName] = useState('');
  const [tried, setTried] = useState(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const taken = new Set(db.tables.map((x) => x.name.toLowerCase()));

  const settings = () => {
    if (kind === 'file') return { where, sheet, header: Number(header) || 1, every: Number(every) };
    if (kind === 'rbase') return { where, person: person.trim() || null, every: Number(every) };
    return {
      where,
      headers: headers.filter((h) => h.name.trim()).map((h) => (h.secret ? { name: h.name.trim(), secret: true } : { name: h.name.trim(), value: h.value })),
      records,
      paging: paging === 'next' ? { kind: 'next', next } : paging === 'page' ? { kind: 'page', param, start: Number(start) } : { kind: 'none' },
      every: Number(every),
    };
  };
  const secretHeaders = () => Object.fromEntries(headers.filter((h) => h.name.trim() && h.secret).map((h) => [h.name.trim(), h.value]));
  const reset = () => { setTried(null); setError(''); };

  const browse = async () => {
    const files = await shell.dialog.open({
      title: kind === 'rbase' ? t('base.sources.another-database') : t('base.sources.file-read'),
      filters: kind === 'rbase' ? [{ name: t('base.rutba-database-rbase'), extensions: ['rbase'] }] : [{ name: t('base.sources.csv-files'), extensions: FILES }],
    });
    if (files?.[0]) { setWhere(files[0]); setSheet(null); setItem(null); reset(); }
  };

  const tryIt = async (over = {}) => {
    setBusy(true);
    setError('');
    try {
      const pick = over.item !== undefined ? over.item : item;
      const s = { ...settings(), ...(over.sheet !== undefined ? { sheet: over.sheet } : {}) };
      const r = await shell.base.trySource({ kind, settings: s, password: kind === 'rbase' ? password || null : null, headers: secretHeaders(), item: pick });
      setTried(r);
      if (r.sheet) setSheet(r.sheet);
      if (r.item) setItem(r.item);
      if (!name.trim()) setName(kind === 'rbase' && r.item ? r.item.name : stem(where) || (isWeb(where) ? (() => { try { return new URL(where).host; } catch { return ''; } })() : ''));
    } catch (e) {
      setTried(null);
      setError(e.message);
      if (kind === 'rbase' && /protected|password|who you are/i.test(e.message)) setNeedWho(true);
    } finally {
      setBusy(false);
    }
  };

  const link = async () => {
    setBusy(true);
    setError('');
    try {
      const out = await shell.base.linkSource({ id: db.id, kind, name: name.trim(), settings: settings(), password: kind === 'rbase' ? password || null : null, headers: secretHeaders(), item });
      onDone(out.summary, out.table);
      toast(t('base.sources.table-linked-read', { table: out.table, every: everyWords(every) }), { tone: 'good', ms: 6000 });
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  };

  const clash = name.trim() && taken.has(name.trim().toLowerCase());
  return (
    <Dialog title={t('base.link-file-web')} width={640} onClose={onClose} actions={(
      <>
        <Button label={t('common.cancel')} onClick={onClose} />
        <Button label={t('common.read')} className="bs-source-try" disabled={busy || !where.trim()} onClick={() => tryIt()} />
        <Button primary label={t('common.link')} className="bs-source-go" disabled={busy || !tried || !tried.fields?.length || !name.trim() || clash} onClick={link} />
      </>
    )}>
      <p className="bs-hint">{t('base.sources.source-keeps-records')}</p>
      <div className="bs-source-form">
        <label className="bs-field bs-source-wide"><span>{t('base.sources.read-from')}</span>
          <Select className="rw-select bs-source-kind" aria-label={t('base.sources.read-from')} value={kind} onChange={(e) => { setKind(e.target.value); setWhere(''); setItem(null); setSheet(null); setNeedWho(false); reset(); }}>
            <option value="file">{t('base.sources.csv-file-spreadsheet')}</option>
            <option value="rbase">{t('base.sources.another-data-files')}</option>
            <option value="json">{t('base.sources.web-address-answering')}</option>
          </Select>
        </label>
        <label className="bs-field bs-source-wide"><span>{kind === 'json' ? t('base.sources.web-address') : kind === 'rbase' ? t('base.sources.where-it-is') : t('base.sources.path-network-share')}</span>
          <span className="bs-source-where-row">
            <Input className="rw-input bs-source-where" aria-label={t('base.sources.where-it-is')} value={where} placeholder={kind === 'json' ? 'https://' : ''} onChange={(e) => { setWhere(e.target.value); reset(); }} onKeyDown={(e) => { if (e.key === 'Enter' && where.trim()) tryIt(); }} />
            {kind === 'json' ? null : <Button icon="open" label={t('common.browse')} className="bs-source-browse" onClick={browse} />}
          </span>
        </label>
        {kind === 'file' ? (
          <>
            <label className="bs-field"><span>{t('base.sources.field-names-row')}</span><Input className="rw-input bs-source-header" aria-label={t('base.sources.field-names-row')} value={header} onChange={(e) => { setHeader(e.target.value.replace(/\D/g, '')); reset(); }} /></label>
            {tried?.sheets?.length ? (
              <label className="bs-field"><span>{t('common.sheet')}</span>
                <Select className="rw-select bs-source-sheet" aria-label={t('common.sheet')} value={sheet || tried.sheet} onChange={(e) => { setSheet(e.target.value); tryIt({ sheet: e.target.value }); }}>
                  {tried.sheets.map((n) => <option key={n} value={n}>{n}</option>)}
                </Select>
              </label>
            ) : null}
          </>
        ) : null}
        {kind === 'rbase' && needWho ? (
          <>
            <label className="bs-field"><span>{t('base.sources.your-name-there')}</span><Input className="rw-input bs-source-person" aria-label={t('base.sources.your-name-there')} value={person} onChange={(e) => setPerson(e.target.value)} /></label>
            <label className="bs-field"><span>{t('base.sources.your-password-there')}</span><Input className="rw-input bs-source-password" type="password" aria-label={t('base.sources.your-password-there')} value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') tryIt(); }} /></label>
          </>
        ) : null}
        {kind === 'rbase' && tried?.items?.length ? (
          <label className="bs-field bs-source-wide"><span>{t('common.table-or-query')}</span>
            <Select className="rw-select bs-source-item" aria-label={t('common.table-or-query')} value={item ? `${item.kind}:${item.name}` : ''} onChange={(e) => { const [k, ...n] = e.target.value.split(':'); const pick = { kind: k, name: n.join(':') }; setItem(pick); setName(pick.name); tryIt({ item: pick }); }}>
              {tried.items.map((x) => <option key={`${x.kind}:${x.name}`} value={`${x.kind}:${x.name}`}>{x.kind === 'query' ? t('common.query-name', { name: x.name }) : t('base.sources.table-name', { name: x.name })}</option>)}
            </Select>
          </label>
        ) : null}
        {kind === 'json' ? (
          <>
            <div className="bs-source-wide bs-source-headers" role="group" aria-label={t('base.sources.headers-sent-each')}>
              <span className="bs-source-label">{t('base.sources.headers-sent-each')}</span>
              {headers.map((h, i) => (
                <span key={i} className="bs-source-header-row">
                  <Input className="rw-input bs-source-header-name" aria-label={t('base.sources.header-name')} placeholder={t('base.sources.header-as')} value={h.name} onChange={(e) => setHeaders((l) => l.map((x, k) => (k === i ? { ...x, name: e.target.value } : x)))} />
                  <Input className="rw-input bs-source-header-value" aria-label={t('base.sources.header-value')} type={h.secret ? 'password' : 'text'} placeholder={t('common.value')} value={h.value} onChange={(e) => setHeaders((l) => l.map((x, k) => (k === i ? { ...x, value: e.target.value } : x)))} />
                  <label className="bs-source-secret"><input type="checkbox" className="bs-source-header-secret" checked={h.secret} onChange={(e) => setHeaders((l) => l.map((x, k) => (k === i ? { ...x, secret: e.target.checked } : x)))} /> {t('base.sources.key-token')}</label>
                </span>
              ))}
              <Button icon="plus" label={t('base.sources.another')} className="bs-source-header-add" onClick={() => setHeaders((l) => [...l, { name: '', value: '', secret: false }])} />
            </div>
            <label className="bs-field"><span>{t('base.sources.records-answer')}</span><Input className="rw-input bs-source-records" aria-label={t('base.sources.records-answer')} placeholder="data.items" value={records} onChange={(e) => { setRecords(e.target.value); reset(); }} /></label>
            <label className="bs-field"><span>{t('base.sources.more-pages')}</span>
              <Select className="rw-select bs-source-paging" aria-label={t('base.sources.more-pages')} value={paging} onChange={(e) => { setPaging(e.target.value); reset(); }}>
                <option value="none">{t('base.sources.one-answer-holds')}</option>
                <option value="next">{t('base.sources.each-answer-gives')}</option>
                <option value="page">{t('base.sources.pages-by-number')}</option>
              </Select>
            </label>
            {paging === 'next' ? <label className="bs-field"><span>{t('base.sources.next-address-empty')}</span><Input className="rw-input bs-source-next" aria-label={t('base.sources.next-address')} placeholder="links.next" value={next} onChange={(e) => setNext(e.target.value)} /></label> : null}
            {paging === 'page' ? (
              <>
                <label className="bs-field"><span>{t('base.sources.page-number-in')}</span><Input className="rw-input bs-source-param" aria-label={t('base.sources.page-number-in')} value={param} onChange={(e) => setParam(e.target.value)} /></label>
                <label className="bs-field"><span>{t('base.sources.first-page')}</span><Input className="rw-input bs-source-start" aria-label={t('base.sources.first-page')} value={start} onChange={(e) => setStart(e.target.value.replace(/\D/g, ''))} /></label>
              </>
            ) : null}
          </>
        ) : null}
        <label className="bs-field"><span>{t('base.read-again')}</span>
          <Select className="rw-select bs-source-every-new" aria-label={t('base.read-again')} value={every} onChange={(e) => setEvery(e.target.value)}>
            {[5, 15, 60, 0].map((n) => <option key={n} value={String(n)}>{everyWords(n)}</option>)}
          </Select>
        </label>
        {tried ? <label className="bs-field"><span>{t('base.sources.name-table')}</span><Input className="rw-input bs-source-name" aria-label={t('base.sources.name-table')} value={name} onChange={(e) => setName(e.target.value)} /></label> : null}
      </div>
      {clash ? <p className="bs-signin-error" role="alert">{t('base.sources.table-called-name', { name: name.trim() })}</p> : null}
      {error ? <p className="bs-signin-error" role="alert">{error}</p> : null}
      {tried?.fields?.length ? (
        <div className="bs-source-preview" data-count={tried.count}>
          <h3><Icon name="sheets" size={13} /> {tn(tried.count, 'base.sources.count-record-fields')}</h3>
          <div className="bs-source-grid">
            <table>
              <thead><tr>{tried.fields.map((f) => <th key={f.name} data-field={f.name}>{f.name}</th>)}</tr></thead>
              <tbody>{tried.rows.map((r, i) => <tr key={i}>{tried.fields.map((f) => <td key={f.name}>{r[f.name] == null ? '' : String(r[f.name])}</td>)}</tr>)}</tbody>
            </table>
          </div>
        </div>
      ) : null}
      {kind === 'json' ? <p className="bs-hint">{t('base.sources.key-token-kept')}</p> : null}
      {kind === 'rbase' ? <p className="bs-hint">{t('base.sources.other-database-opened')}</p> : null}
    </Dialog>
  );
}

/** Who you are in a protected database a source reads, asked for when it is not known on this computer. */
export function SourceSignInDialog({ shell, db, table, onDone, onClose, toast }) {
  const where = (db.sources || []).find((x) => x.id === db.tables.find((tb) => tb.name === table)?.linked)?.settings?.where || '';
  const [person, setPerson] = useState('');
  const [password, setPassword] = useState('');
  const go = async () => {
    try {
      const s = await shell.base.sourceSignIn({ id: db.id, table, person, password });
      const err = s.tables.find((x) => x.name === table)?.linkError;
      if (err) { toast(err, { tone: 'bad', ms: 7000 }); return; }
      onDone(s);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  return (
    <Dialog title={t('base.sources.table-read', { table })} onClose={onClose} actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.sign-in')} className="bs-source-who-go" disabled={!person || !password} onClick={go} /></>}>
      {where ? <p className="bs-hint bs-source-who-where">{t('base.sources.database', { where })}</p> : null}
      <label className="bs-field"><span>{t('base.sources.your-name-there')}</span><Input className="rw-input bs-source-who-name" aria-label={t('base.sources.your-name-there')} autoFocus value={person} onChange={(e) => setPerson(e.target.value)} /></label>
      <label className="bs-field"><span>{t('base.sources.your-password-there')}</span><Input className="rw-input bs-source-who-password" type="password" aria-label={t('base.sources.your-password-there')} value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && person && password) go(); }} /></label>
      <p className="bs-hint">{t('base.sources.password-kept-computers')}</p>
    </Dialog>
  );
}

/**
 * A web address's keys or tokens typed again: one kept for another address
 * (the file's link was changed elsewhere), or refused. Each is sent only to
 * the address named, and kept in this computer's protected store.
 */
export function SourceKeysDialog({ shell, db, table, onDone, onClose, toast }) {
  const source = (db.sources || []).find((x) => x.id === db.tables.find((tb) => tb.name === table)?.linked);
  const names = (source?.settings?.headers || []).filter((h) => h.secret).map((h) => h.name);
  const [values, setValues] = useState({});
  const where = source?.settings?.where || '';
  const go = async () => {
    try {
      const s = await shell.base.sourceKeys({ id: db.id, table, headers: values });
      const err = s.tables.find((x) => x.name === table)?.linkError;
      if (err) { toast(err, { tone: 'bad', ms: 7000 }); return; }
      onDone(s);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  const ready = names.length && names.every((n) => values[n]);
  return (
    <Dialog title={t('base.sources.key-token-table', { table })} onClose={onClose} actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.read')} className="bs-source-keys-go" disabled={!ready} onClick={go} /></>}>
      <p className="bs-hint bs-source-keys-where">{t('base.sources.sent-only', { where })}</p>
      {names.map((n, i) => (
        <label key={n} className="bs-field"><span>{n}</span><Input className="rw-input bs-source-keys-value" type="password" aria-label={n} autoFocus={i === 0} value={values[n] || ''} onChange={(e) => setValues((v) => ({ ...v, [n]: e.target.value }))} onKeyDown={(e) => { if (e.key === 'Enter' && ready) go(); }} /></label>
      ))}
      <p className="bs-hint">{t('base.sources.key-kept-computers')}</p>
    </Dialog>
  );
}

export const SOURCES_CSS = `
.bs-source-form { display: grid; grid-template-columns: 1fr 1fr; gap: 8px 12px; }
.bs-source-wide { grid-column: 1 / -1; }
.bs-source-where-row { display: flex; gap: 6px; }
.bs-source-where-row .rw-input { flex: 1; min-width: 0; }
.bs-source-headers { display: flex; flex-direction: column; gap: 6px; }
.bs-source-label { font-size: 12px; color: var(--ink-2); }
.bs-source-header-row { display: grid; grid-template-columns: 1fr 1fr auto; gap: 6px; align-items: center; }
.bs-source-secret { display: flex; gap: 4px; align-items: center; font-size: 12px; white-space: nowrap; }
.bs-source-preview { margin-top: 10px; }
.bs-source-preview h3 { margin: 0 0 6px; font-size: 12px; color: var(--ink-2); display: flex; gap: 6px; align-items: center; }
.bs-source-grid { max-height: 180px; overflow: auto; border: 1px solid var(--line-soft); border-radius: var(--r-2, 6px); }
.bs-source-grid table { border-collapse: collapse; font-size: 12px; width: 100%; }
.bs-source-grid th, .bs-source-grid td { text-align: left; padding: 3px 8px; border-bottom: 1px solid var(--line-soft); white-space: nowrap; }
.bs-source-note { font-weight: 600; }
`;
