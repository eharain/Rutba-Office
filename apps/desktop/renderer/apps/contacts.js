// Rutba Contacts.
//
// An address book that lives on this computer: every card the person keeps,
// searchable as they type, editable in place, brought in from the .vcf and
// .csv files every other program exports and written back out the same way.
// Mail asks it to complete an address, and can hand it a sender to keep.
//
// A .vcf opened from disk is shown, not kept: its cards appear with an offer
// to add them, and nothing changes until the person says so.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Ribbon, Group, Button, Separator, Icon, Spacer, Chip, Empty, Panel, Content, Search, Field, Input, Select, Dialog,
  useToast, useCommands, menuItems, useMenu, t, tn,
} from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, useFileDrop } from '../shell.js';
import { AccountsDialog } from '../dav-accounts.js';

const EMPTY_CONTACT = () => ({
  id: null,
  name: { full: '', given: '', family: '', middle: '', prefix: '', suffix: '' },
  nickname: '', org: '', department: '', title: '', role: '',
  emails: [{ value: '', type: 'work', pref: true, label: null }],
  phones: [{ value: '', type: 'mobile', pref: true, label: null }],
  addresses: [],
  birthday: '', note: '', urls: [], categories: [], kind: 'individual', extra: [],
});

const TYPES = { email: ['work', 'home', 'other'], phone: ['mobile', 'work', 'home', 'fax', 'other'], address: ['work', 'home', 'other'] };
const TYPE_NAMES = { work: t('work'), home: t('home'), other: t('other'), mobile: t('mobile'), fax: t('fax') };

const initialsOf = (c) => {
  const n = c.name || {};
  const parts = [n.given, n.family].filter(Boolean);
  const source = parts.length ? parts : String(c.display || n.full || c.org || '').split(/\s+/);
  return source.slice(0, 2).map((s) => (s[0] || '').toUpperCase()).join('') || '?';
};

/** The letter a card sorts under. */
const letterOf = (c) => {
  const ch = String(c.display || '').trim()[0] || '#';
  return /[a-z]/i.test(ch) ? ch.toUpperCase() : '#';
};

export default function Contacts({ app, shell, boot }) {
  const toast = useToast();
  // Home → Sync: every account now, and what moved said.
  const [accountsOpen, setAccountsOpen] = useState(false);
  const syncAll = useCallback(async () => {
    const results = (await shell.dav.sync({})) || [];
    if (!results.length) { setAccountsOpen(true); return; }
    const failed = results.find((r) => r.error);
    if (failed) toast(failed.error, { tone: 'warn', ms: 5500 });
    else toast(t('Synced: {received} in, {sent} out', { received: results.reduce((n, r) => n + (r.received || 0), 0), sent: results.reduce((n, r) => n + (r.sent || 0), 0) }), { tone: 'good' });
    
  }, [shell, toast]);

  const menu = useMenu();
  const [contacts, setContacts] = useState([]);
  const [query, setQuery] = useState('');
  const [selectedId, setSelectedId] = useState(null);
  const [draft, setDraft] = useState(null);
  const [peek, setPeek] = useState(null);
  const [tab, setTab] = useState('home');
  const [busy, setBusy] = useState(false);
  // A CSV being brought in: checked in a preview first (CsvImport).
  const [csv, setCsv] = useState(null);
  const searchRef = useRef(null);

  const refresh = useCallback(async () => {
    try {
      setContacts(await shell.contacts.list({ query }));
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [shell, query, toast]);

  useEffect(() => {
    refresh();
  }, [refresh]);
  useEffect(() => shell.on('contacts:changed', refresh), [shell, refresh]);

  // A file opened with the window, or dropped on it: shown, and offered.
  const showFile = useCallback(
    async (file) => {
      try {
        const seen = await shell.contacts.peek({ path: file });
        setPeek(seen);
        setSelectedId(seen.contacts[0]?.id || null);
        setDraft(null);
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 8000 });
      }
    },
    [shell, toast]
  );
  useEffect(() => {
    if (boot.file) showFile(boot.file);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useFileDrop(useCallback((files) => files.forEach((f) => (/\.(vcf|csv|vcard)$/i.test(f) ? showFile(f) : null)), [showFile]));

  const list = peek ? peek.contacts : contacts;
  const selected = useMemo(() => list.find((c) => c.id === selectedId) || null, [list, selectedId]);

  const groups = useMemo(() => {
    const out = [];
    let last = null;
    for (const c of list) {
      const letter = letterOf(c);
      if (letter !== last) {
        out.push({ letter, items: [] });
        last = letter;
      }
      out[out.length - 1].items.push(c);
    }
    return out;
  }, [list]);

  /* ── actions ──────────────────────────────────────────────────────────── */

  const startNew = useCallback(() => {
    setPeek(null);
    setSelectedId(null);
    setDraft(EMPTY_CONTACT());
  }, []);

  const edit = useCallback(() => {
    if (!selected) return;
    setDraft({
      ...selected,
      name: { ...selected.name },
      emails: selected.emails?.length ? selected.emails.map((e) => ({ ...e })) : [{ value: '', type: 'work', pref: true, label: null }],
      phones: selected.phones?.length ? selected.phones.map((p) => ({ ...p })) : [{ value: '', type: 'mobile', pref: true, label: null }],
      addresses: (selected.addresses || []).map((a) => ({ ...a })),
      birthday: selected.birthday || '',
      note: selected.note || '',
      org: selected.org || '',
      title: selected.title || '',
    });
  }, [selected]);

  const saveDraft = useCallback(async () => {
    if (!draft) return;
    const full = draft.name.full.trim() || [draft.name.given, draft.name.family].filter(Boolean).join(' ').trim();
    const emails = draft.emails.filter((e) => e.value.trim());
    if (!full && !emails.length && !draft.org) {
      toast(t('A card needs a name, an address or a company.'), { tone: 'bad' });
      return;
    }
    setBusy(true);
    try {
      const contact = {
        ...draft,
        name: full && !draft.name.given && !draft.name.family ? { ...draft.name, full, ...splitTyped(full) } : { ...draft.name, full },
        emails,
        phones: draft.phones.filter((p) => p.value.trim()),
        addresses: draft.addresses.filter((a) => a.street || a.city || a.postcode || a.country),
        birthday: draft.birthday || null,
        note: draft.note || null,
        org: draft.org || null,
        title: draft.title || null,
      };
      const saved = await shell.contacts.save({ contact });
      setDraft(null);
      setPeek(null);
      setSelectedId(saved.id);
      await refresh();
      toast(t('Kept {name}.', { name: saved.display }), { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setBusy(false);
    }
  }, [draft, shell, refresh, toast]);

  const remove = useCallback(async () => {
    if (!selected || peek) return;
    await shell.contacts.remove({ id: selected.id });
    setSelectedId(null);
    setDraft(null);
    await refresh();
    toast(t('Removed {name}.', { name: selected.display }));
  }, [selected, peek, shell, refresh, toast]);

  /** What an import did, said: and for a CSV, what it left out. */
  const imported = useCallback(async (r) => {
    setPeek(null);
    await refresh();
    const done = r.same ? t('{added} added, {updated} updated, {same} already here.', { added: r.added, updated: r.updated, same: r.same }) : t('{added} added, {updated} updated.', { added: r.added, updated: r.updated });
    const left = leftOut(r.skipped, r.unread);
    toast(left ? `${done} ${left}` : done, { tone: 'good', ms: left ? 9000 : 6000 });
  }, [refresh, toast]);

  const importFile = useCallback(
    async (file = null) => {
      const target = file || (await pickOpen(shell, 'contacts'))?.[0];
      if (!target) return;
      if (/\.(csv|tsv)$/i.test(target)) { setCsv({ path: target }); return; }
      setBusy(true);
      try {
        await imported(await shell.contacts.importFile({ path: target }));
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 8000 });
      } finally {
        setBusy(false);
      }
    },
    [shell, imported, toast]
  );

  const exportAll = useCallback(async () => {
    const target = await pickSave(shell, 'contacts', 'contacts.vcf');
    if (!target) return;
    try {
      const r = await shell.contacts.exportFile({ path: target });
      toast(tn(r.count, 'Wrote {count} card to {file}.', 'Wrote {count} cards to {file}.', { file: target.split(/[\\/]/).pop() }), { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [shell, toast]);

  const writeTo = useCallback(
    (email) => {
      if (!email) return;
      shell.win.create({ app: 'mail', query: { to: email } });
    },
    [shell]
  );

  const appMenu = useAppMenu({ shell, appKey: 'contacts', onNew: startNew, onOpen: () => importFile() });

  const commands = useMemo(
    () => ({
      'contact.new': { label: t('New contact'), icon: 'plus', key: 'Mod+N', run: startNew },
      'contact.edit': { label: t('Edit'), icon: 'textbox', run: edit },
      'contact.delete': { label: t('Delete'), icon: 'trash', key: 'Delete', run: remove },
      'contact.import': { label: t('Import…'), icon: 'import', key: 'Mod+O', run: () => importFile() },
      'contact.export': { label: t('Export…'), icon: 'export', run: exportAll },
      'contact.mail': { label: t('Send mail'), icon: 'send', run: () => writeTo(selected?.emails?.[0]?.value) },
      'contact.find': { label: t('Find'), icon: 'find', key: 'Mod+F', run: () => searchRef.current?.focus() },
    }),
    [startNew, edit, remove, importFile, exportAll, writeTo, selected]
  );
  useCommands(commands, [selected, draft]);

  const title = peek ? `${peek.name} (${peek.contacts.length})` : tn(contacts.length, '{count} contact', '{count} contacts');

  /* ── the card ─────────────────────────────────────────────────────────── */

  const card = selected && !draft ? (
    <div className="ct-card">
      <div className="ct-head">
        <Avatar contact={selected} size={72} shell={shell} />
        <div className="ct-who">
          <h2>{selected.display}</h2>
          {selected.title || selected.org ? <div className="ct-sub">{[selected.title, selected.org].filter(Boolean).join(' · ')}</div> : null}
          {selected.categories?.length ? <div className="ct-tags">{selected.categories.map((c) => <Chip key={c}>{c}</Chip>)}</div> : null}
        </div>
      </div>
      {peek ? (
        <div className="ct-offer">
          <Icon name="info" size={14} />
          <span>{t('From {name}, not yet in your contacts.', { name: peek.name })}</span>
          <Button primary label={peek.contacts.length === 1 ? t('Add this card') : t('Add all {count}', { count: peek.contacts.length })} onClick={() => importFile(peek.path)} />
        </div>
      ) : null}
      <dl className="ct-fields">
        {(selected.emails || []).map((e, i) => (
          <React.Fragment key={`e${i}`}>
            <dt>{e.label || (e.type && e.type !== 'other' ? TYPE_NAMES[e.type] || e.type : t('email'))}</dt>
            <dd><button type="button" className="ct-link" onClick={() => writeTo(e.value)} title={t('Write a message')}>{e.value}</button></dd>
          </React.Fragment>
        ))}
        {(selected.phones || []).map((p, i) => (
          <React.Fragment key={`p${i}`}>
            <dt>{p.label || (p.type && p.type !== 'other' ? TYPE_NAMES[p.type] || p.type : t('phone'))}</dt>
            <dd>{p.value}</dd>
          </React.Fragment>
        ))}
        {(selected.addresses || []).map((a, i) => (
          <React.Fragment key={`a${i}`}>
            <dt>{a.label || (a.type && a.type !== 'other' ? TYPE_NAMES[a.type] || a.type : t('address'))}</dt>
            <dd className="ct-addr">{[a.street, a.city, a.region, a.postcode, a.country].filter(Boolean).join('\n')}</dd>
          </React.Fragment>
        ))}
        {selected.birthday ? (<><dt>{t('birthday')}</dt><dd>{selected.birthday}</dd></>) : null}
        {(selected.urls || []).map((u, i) => (
          <React.Fragment key={`u${i}`}>
            <dt>{u.label || t('web')}</dt>
            <dd><a href={u.value} onClick={(e) => { e.preventDefault(); shell.shell.openExternal({ url: u.value }); }}>{u.value}</a></dd>
          </React.Fragment>
        ))}
        {selected.note ? (<><dt>{t('notes')}</dt><dd className="ct-note">{selected.note}</dd></>) : null}
      </dl>
    </div>
  ) : null;

  /* ── the editor ───────────────────────────────────────────────────────── */

  const set = (patch) => setDraft((d) => ({ ...d, ...patch }));
  const setName = (patch) => setDraft((d) => ({ ...d, name: { ...d.name, ...patch } }));
  const setRow = (list, i, patch) => setDraft((d) => ({ ...d, [list]: d[list].map((row, j) => (j === i ? { ...row, ...patch } : row)) }));
  const addRow = (list, row) => setDraft((d) => ({ ...d, [list]: [...d[list], row] }));
  const dropRow = (list, i) => setDraft((d) => ({ ...d, [list]: d[list].filter((_, j) => j !== i) }));

  const editor = draft ? (
    <div className="ct-editor">
      <div className="ct-head">
        <Avatar contact={{ ...draft, display: draft.name.full || [draft.name.given, draft.name.family].filter(Boolean).join(' ') }} size={72} />
        <div className="ct-who grow">
          <Input className="rw-input ct-name" value={draft.name.full} placeholder={t('Name')} autoFocus onChange={(e) => { const full = e.target.value; setDraft((d) => ({ ...d, name: { ...d.name, full, ...splitTyped(full) } })); }} />
          <div className="ct-row2">
            <Input value={draft.title} placeholder={t('Job title')} onChange={(e) => set({ title: e.target.value })} />
            <Input value={draft.org} placeholder={t('Company')} onChange={(e) => set({ org: e.target.value })} />
          </div>
        </div>
      </div>
      <div className="ct-section">
        <h4>{t('Email')}</h4>
        {draft.emails.map((e, i) => (
          <div className="ct-line" key={i}>
            <Select value={e.type} aria-label={t('Type')} onChange={(ev) => setRow('emails', i, { type: ev.target.value })}>{TYPES.email.map((t) => <option key={t} value={t}>{TYPE_NAMES[t] ?? t}</option>)}</Select>
            <Input value={e.value} aria-label={t('Email')} placeholder="name@example.com" onChange={(ev) => setRow('emails', i, { value: ev.target.value })} />{/* words-ok: an example address */}
            <Button icon="close" title={t('Remove')} onClick={() => dropRow('emails', i)} />
          </div>
        ))}
        <Button ghost icon="plus" label={t('Add email')} onClick={() => addRow('emails', { value: '', type: 'home', pref: false, label: null })} />
      </div>
      <div className="ct-section">
        <h4>{t('Phone')}</h4>
        {draft.phones.map((p, i) => (
          <div className="ct-line" key={i}>
            <Select value={p.type} aria-label={t('Type')} onChange={(ev) => setRow('phones', i, { type: ev.target.value })}>{TYPES.phone.map((t) => <option key={t} value={t}>{TYPE_NAMES[t] ?? t}</option>)}</Select>
            <Input value={p.value} aria-label={t('Phone')} placeholder="+44 7700 900123" onChange={(ev) => setRow('phones', i, { value: ev.target.value })} />
            <Button icon="close" title={t('Remove')} onClick={() => dropRow('phones', i)} />
          </div>
        ))}
        <Button ghost icon="plus" label={t('Add phone')} onClick={() => addRow('phones', { value: '', type: 'work', pref: false, label: null })} />
      </div>
      <div className="ct-section">
        <h4>{t('Address')}</h4>
        {draft.addresses.map((a, i) => (
          <div className="ct-address" key={i}>
            <div className="ct-line">
              <Select value={a.type} aria-label={t('Type')} onChange={(ev) => setRow('addresses', i, { type: ev.target.value })}>{TYPES.address.map((t) => <option key={t} value={t}>{TYPE_NAMES[t] ?? t}</option>)}</Select>
              <Input value={a.street} placeholder={t('Street')} onChange={(ev) => setRow('addresses', i, { street: ev.target.value })} />
              <Button icon="close" title={t('Remove')} onClick={() => dropRow('addresses', i)} />
            </div>
            <div className="ct-line four">
              <Input value={a.city} placeholder={t('City')} onChange={(ev) => setRow('addresses', i, { city: ev.target.value })} />
              <Input value={a.region} placeholder={t('Region')} onChange={(ev) => setRow('addresses', i, { region: ev.target.value })} />
              <Input value={a.postcode} placeholder={t('Postcode')} onChange={(ev) => setRow('addresses', i, { postcode: ev.target.value })} />
              <Input value={a.country} placeholder={t('Country')} onChange={(ev) => setRow('addresses', i, { country: ev.target.value })} />
            </div>
          </div>
        ))}
        <Button ghost icon="plus" label={t('Add address')} onClick={() => addRow('addresses', { street: '', city: '', region: '', postcode: '', country: '', type: 'home', pref: false, label: null })} />
      </div>
      <div className="ct-section ct-two">
        <Field label={t('Birthday')}><Input type="date" value={draft.birthday || ''} onChange={(e) => set({ birthday: e.target.value })} /></Field>
        <Field label={t('Categories')} hint={t('Comma-separated')}><Input value={(draft.categories || []).join(', ')} onChange={(e) => set({ categories: e.target.value.split(',').map((s) => s.trim()).filter(Boolean) })} /></Field>
      </div>
      <div className="ct-section">
        <Field label={t('Notes')}><textarea className="rw-input ct-notes" rows={4} value={draft.note || ''} onChange={(e) => set({ note: e.target.value })} /></Field>
      </div>
      <div className="ct-actions">
        <Button label={t('Cancel')} onClick={() => setDraft(null)} />
        <Button primary label={busy ? t('Saving…') : t('Save')} disabled={busy} onClick={saveDraft} />
      </div>
    </div>
  ) : null;

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={title}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[{ id: 'home', label: t('Home') }, { id: 'view', label: t('View') }]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="plus" title={t('New contact (Ctrl+N)')} onClick={startNew} />
              <Button icon="trash" title={t('Delete')} onClick={remove} disabled={!selected || Boolean(peek)} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label={t('New')}>
                <Button tall icon="plus" label={t('New contact')} onClick={startNew} />
              </Group>
              <Group label={t('Card')}>
                <Button tall icon="textbox" label={t('Edit')} onClick={edit} disabled={!selected || Boolean(peek)} />
                <Button tall icon="trash" label={t('Delete')} onClick={remove} disabled={!selected || Boolean(peek)} />
                <Button tall icon="send" label={t('Send mail')} onClick={() => writeTo(selected?.emails?.[0]?.value)} disabled={!selected?.emails?.length} />
              </Group>
              <Group label={t('Files')}>
                <Button tall icon="import" label={t('Import')} title={t('A .vcf or .csv from another program')} onClick={() => importFile()} />
                <Button tall icon="export" label={t('Export')} title={t('All cards as a .vcf')} onClick={exportAll} disabled={!contacts.length} />
              </Group>
              <Group label={t('Accounts')}>
                <Button tall icon="globe" label={t('Accounts')} title={t('Accounts — a CalDAV or CardDAV server (iCloud, Fastmail, Nextcloud and the like) kept in step with this computer')} onClick={() => setAccountsOpen(true)} />
                <Button tall icon="refresh" label={t('Sync')} title={t('Sync every account now')} onClick={syncAll} />
              </Group>
            </>
          ) : (
            <Group label={t('Find')}>
              <Button tall icon="find" label={t('Find')} onClick={() => searchRef.current?.focus()} />
            </Group>
          )}
        </Ribbon>
      }
    >
      <style>{CSS}</style>
      {accountsOpen ? <AccountsDialog shell={shell} kind="contacts" toast={toast} onClose={() => setAccountsOpen(false)} /> : null}
      {csv ? <CsvImport shell={shell} path={csv.path} toast={toast} onClose={() => setCsv(null)} onDone={(r) => { setCsv(null); imported(r); }} /> : null}
      <Panel width={280} resizable>
        <div className="ct-search">
          <Search value={query} onChange={(e) => setQuery(e.target.value)} placeholder={t('Search contacts')} autoFocus={false} />
        </div>
        {peek ? (
          <div className="ct-peek">
            <Icon name="file" size={14} />
            <span className="grow">{peek.name}</span>
            <Button ghost icon="close" title={t('Back to your contacts')} onClick={() => { setPeek(null); setSelectedId(null); }} />
          </div>
        ) : null}
        {/* A list a screen reader reads as one: each card an option, the chosen one selected,
            one stop in the Tab order and the arrow keys through the rest, as Outlook's People list. */}
        <div
          className="ct-list"
          role="listbox"
          aria-label={t('Contacts')}
          onKeyDown={(e) => {
            const all = [...e.currentTarget.querySelectorAll('.ct-item')];
            const at = all.indexOf(document.activeElement);
            if (at < 0) return;
            const to = e.key === 'ArrowDown' ? at + 1 : e.key === 'ArrowUp' ? at - 1 : e.key === 'Home' ? 0 : e.key === 'End' ? all.length - 1 : -1;
            if (to < 0 || to >= all.length || !/^(ArrowDown|ArrowUp|Home|End)$/.test(e.key)) return;
            e.preventDefault();
            all[to].focus();
            all[to].click();
          }}
        >
          {groups.map((g, gi) => (
            <div key={g.letter} role="group" aria-label={g.letter}>
              <div className="ct-letter" aria-hidden="true">{g.letter}</div>
              {g.items.map((c, ci) => (
                <button
                  type="button"
                  key={c.id}
                  role="option"
                  aria-selected={c.id === selectedId}
                  tabIndex={c.id === selectedId || (!list.some((x) => x.id === selectedId) && gi === 0 && ci === 0) ? 0 : -1}
                  className={`ct-item${c.id === selectedId ? ' current' : ''}`}
                  onClick={() => { setSelectedId(c.id); setDraft(null); }}
                  onContextMenu={(e) => { setSelectedId(c.id); menu.open(e, menuItems(commands, ['contact.edit', 'contact.mail', '-', 'contact.delete'])); }}
                >
                  <Avatar contact={c} size={30} shell={shell} />
                  <span className="ct-item-text">
                    <span className="name">{c.display || t('(no name)')}</span>
                    <span className="sub">{c.emails?.[0]?.value || c.org || c.phones?.[0]?.value || ''}</span>
                  </span>
                </button>
              ))}
            </div>
          ))}
          {!list.length ? <div className="ct-none" role="option" aria-selected="false" aria-disabled="true">{query ? t('Nothing matches.') : peek ? t('The file holds no cards.') : t('No contacts yet.')}</div> : null}
        </div>
      </Panel>
      <Content>
        {editor || card || (
          <Empty icon="contacts" title={contacts.length ? t('Choose a contact') : t('No contacts yet')}>
            {contacts.length ? t('Pick one on the left, or press Ctrl+N for a new card.') : t('Add a card, or bring in the .vcf or .csv another program exported.')}
            <div style={{ marginTop: 14, display: 'flex', gap: 8, justifyContent: 'center' }}>
              <Button primary icon="plus" label={t('New contact')} onClick={startNew} />
              <Button icon="import" label={t('Import…')} onClick={() => importFile()} />
            </div>
          </Empty>
        )}
      </Content>
    </AppFrame>
  );
}

/** What a CSV import left out, in a sentence; '' when nothing. */
function leftOut(skipped = 0, unread = []) {
  const parts = [];
  if (skipped) parts.push(tn(skipped, 'Left out {count} row with no name, address, company or phone.', 'Left out {count} rows with no name, address, company or phone.'));
  if (unread?.length) parts.push(tn(unread.length, 'Not read: the {names} column.', 'Not read: the {names} columns.', { names: unread.join(', ') }));
  return parts.join(' ');
}

/** What a CSV column can be read as, in the words the preview shows. */
const CSV_FIELD_NAMES = {
  '': t('Leave out'), full: t('Full name'), prefix: t('Title before the name'), first: t('First name'), middle: t('Middle name'), last: t('Last name'), suffix: t('After the name'),
  nickname: t('Nickname'), org: t('Company'), department: t('Department'), jobTitle: t('Job title'), email: t('Email'), phone: t('Phone'),
  street: t('Street'), city: t('City'), region: t('Region'), postcode: t('Postcode'), country: t('Country'), pobox: t('PO box'), extended: t('Address, second line'), formatted: t('Address on one line'),
  birthday: t('Birthday'), website: t('Website'), note: t('Notes'), categories: t('Categories'), type: t('Kind of the column before'),
};
const CSV_CHOICES = ['', 'full', 'prefix', 'first', 'middle', 'last', 'suffix', 'nickname', 'org', 'department', 'jobTitle', 'email', 'phone', 'street', 'city', 'region', 'postcode', 'country', 'birthday', 'website', 'note', 'categories'];

/**
 * A CSV checked before it comes in, as Outlook's import maps its fields:
 * each column, its first value and what it was matched to, which can be
 * changed; the first few cards as they would be kept; and what would be left
 * out. Import brings them in with the columns as they stand.
 */
function CsvImport({ shell, path, toast, onClose, onDone }) {
  const [columns, setColumns] = useState({});
  const [seen, setSeen] = useState(null);
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    let live = true;
    shell.contacts.peekCsv({ path, columns }).then((r) => { if (live) setSeen(r); }).catch((err) => { toast(err.message, { tone: 'bad', ms: 8000 }); onClose(); });
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, path, columns]);
  const go = async () => {
    setBusy(true);
    try {
      onDone(await shell.contacts.importFile({ path, columns }));
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 8000 });
      setBusy(false);
    }
  };
  const left = seen ? leftOut(seen.skipped.length, seen.unread) : '';
  return (
    <Dialog
      title={t('Import {name}', { name: seen?.name || path.split(/[\\/]/).pop() })}
      width={680}
      className="ct-csv"
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary className="ct-csv-go" label={busy ? t('Importing…') : seen ? tn(seen.count, 'Import {count} card', 'Import {count} cards') : t('Import')} disabled={!seen?.count || busy} onClick={go} />
        </>
      }
    >
      {seen ? (
        <div className="ct-csv-body">
          <p className="ct-csv-note">{t('Check what each column holds. Change one that was matched wrongly, or leave it out.')}</p>
          <div className="ct-csv-cols" role="table" aria-label={t('Columns')}>
            <div className="ct-csv-row head" role="row"><span role="columnheader">{t('Column')}</span><span role="columnheader">{t('First value')}</span><span role="columnheader">{t('Read as')}</span></div>
            {seen.columns.map((c) => (
              <div key={c.index} className={`ct-csv-row${c.field ? '' : ' out'}`} role="row" data-column={c.header}>
                <span role="cell" className="h">{c.header || t('(no heading)')}</span>
                <span role="cell" className="v" title={c.sample}>{c.sample}</span>
                <span role="cell">
                  <Select className="ct-csv-field" aria-label={t('What {column} holds', { column: c.header })} value={c.field} onChange={(e) => setColumns((was) => ({ ...was, [c.index]: e.target.value }))}>
                    {[...new Set([...(CSV_CHOICES.includes(c.auto) ? [] : [c.auto]), ...(CSV_CHOICES.includes(c.field) ? [] : [c.field]), ...CSV_CHOICES])].map((f) => <option key={f} value={f}>{CSV_FIELD_NAMES[f] ?? f}</option>)}
                  </Select>
                </span>
              </div>
            ))}
          </div>
          <div className="ct-csv-cards">
            <div className="ct-csv-cards-head">{tn(seen.count, '{count} card; the first as it will be kept:', '{count} cards; the first few as they will be kept:')}</div>
            {seen.sample.map((c) => (
              <div key={c.id} className="ct-csv-card">
                <b>{c.display || t('(no name)')}</b>
                <span>{[c.title, c.org].filter(Boolean).join(' · ')}</span>
                <span>{(c.emails || []).map((e) => e.value).join(', ')}</span>
                <span>{(c.phones || []).map((p) => `${p.label || TYPE_NAMES[p.type] || p.type} ${p.value}`).join(', ')}</span>
              </div>
            ))}
          </div>
          {left ? <p className="ct-csv-left">{left}</p> : null}
        </div>
      ) : <div className="ct-csv-note">{t('Reading the file…')}</div>}
    </Dialog>
  );
}

function splitTyped(full) {
  const words = String(full || '').trim().split(/\s+/).filter(Boolean);
  const prefixes = new Set(['mr', 'mrs', 'ms', 'miss', 'dr', 'prof', 'mr.', 'mrs.', 'ms.', 'dr.', 'prof.']);
  const prefix = words.length > 1 && prefixes.has(words[0].toLowerCase()) ? words.shift() : '';
  const given = words.shift() || '';
  const family = words.pop() || '';
  return { given, family, middle: words.join(' '), prefix };
}

/** Photos asked for once a card, by its id and when it last changed. */
const photos = new Map();

/** A card's photo, asked for when it is first drawn: a list sends only that there is one. */
function usePhoto(shell, contact) {
  const key = contact?.id && contact.hasPhoto && !contact.photoUrl ? `${contact.id}:${contact.updatedAt || ''}` : null;
  const [url, setUrl] = useState(() => (key ? photos.get(key) ?? null : null));
  useEffect(() => {
    if (!key || !shell) return undefined;
    if (photos.has(key)) { setUrl(photos.get(key)); return undefined; }
    let live = true;
    shell.contacts.get({ id: contact.id }).then((full) => {
      photos.set(key, full?.photoUrl || null);
      if (live) setUrl(full?.photoUrl || null);
    }).catch(() => {});
    return () => { live = false; };
  }, [key, shell, contact?.id]);
  return contact?.photoUrl || url;
}

function Avatar({ contact, size, shell = null }) {
  const style = { width: size, height: size, fontSize: Math.round(size * 0.38) };
  const photoUrl = usePhoto(shell, contact);
  if (photoUrl) return <img className="ct-avatar" src={photoUrl} alt="" style={style} />;
  const hue = [...String(contact.display || contact.name?.full || '')].reduce((h, ch) => (h * 31 + ch.charCodeAt(0)) % 360, 7);
  return <span className="ct-avatar" style={{ ...style, background: `hsl(${hue} 45% 82%)`, color: `hsl(${hue} 50% 28%)` }}>{initialsOf(contact)}</span>;
}

const CSS = `
.ct-search { padding: 8px 10px 4px; }
.ct-peek { display: flex; align-items: center; gap: 8px; margin: 4px 10px; padding: 6px 8px; border-radius: var(--r-2); background: var(--sunken); font-size: 12px; }
.ct-peek .grow { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ct-list { overflow: auto; flex: 1; min-height: 0; padding: 4px 6px 12px; }
.ct-letter { padding: 10px 8px 4px; font-size: 11px; font-weight: 600; color: var(--ink-3); letter-spacing: .04em; }
.ct-item { display: flex; align-items: center; gap: 10px; width: 100%; padding: 6px 8px; border: 0; border-radius: var(--r-2); background: none; text-align: left; font: inherit; color: var(--ink); cursor: pointer; }
.ct-item:hover { background: var(--hover); }
.ct-item.current { background: var(--selected); }
.ct-item-text { display: flex; flex-direction: column; min-width: 0; }
.ct-item-text .name { font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ct-item-text .sub { font-size: 12px; color: var(--ink-3); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ct-none { padding: 24px 12px; color: var(--ink-3); font-size: 13px; text-align: center; }
.ct-avatar { display: inline-flex; align-items: center; justify-content: center; flex: none; border-radius: 50%; object-fit: cover; font-weight: 600; letter-spacing: .02em; }
.ct-card, .ct-editor { max-width: 720px; margin: 0 auto; padding: 28px 32px 40px; }
.ct-head { display: flex; align-items: center; gap: 20px; margin-bottom: 18px; }
.ct-who h2 { margin: 0 0 4px; font-size: 22px; font-weight: 600; }
.ct-who.grow { flex: 1; display: flex; flex-direction: column; gap: 8px; }
.ct-sub { color: var(--ink-2); font-size: 14px; }
.ct-tags { display: flex; gap: 6px; margin-top: 8px; }
.ct-offer { display: flex; align-items: center; gap: 10px; margin: 0 0 16px; padding: 10px 12px; border-radius: var(--r-2); background: var(--sunken); font-size: 13px; }
.ct-offer span { flex: 1; }
.ct-fields { display: grid; grid-template-columns: 110px 1fr; gap: 10px 16px; margin: 0; font-size: 14px; }
.ct-fields dt { color: var(--ink-3); text-transform: lowercase; padding-top: 1px; }
.ct-fields dd { margin: 0; min-width: 0; overflow-wrap: anywhere; }
.ct-addr, .ct-note { white-space: pre-line; }
.ct-link { border: 0; background: none; padding: 0; font: inherit; color: var(--accent-text); cursor: pointer; text-align: left; }
.ct-link:hover { text-decoration: underline; }
.ct-name { font-size: 20px; font-weight: 600; }
.ct-row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
.ct-section { margin: 14px 0; display: flex; flex-direction: column; gap: 8px; }
.ct-section h4 { margin: 0; font-size: 12px; font-weight: 600; color: var(--ink-3); letter-spacing: .04em; text-transform: uppercase; }
.ct-line { display: grid; grid-template-columns: 110px 1fr 34px; gap: 8px; align-items: center; }
.ct-line.four { grid-template-columns: 1fr 1fr 1fr 1fr; }
.ct-address { display: flex; flex-direction: column; gap: 6px; }
.ct-two { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.ct-notes { width: 100%; resize: vertical; font: inherit; }
.ct-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 18px; }
.ct-csv-body { display: flex; flex-direction: column; gap: 10px; font-size: 13px; }
.ct-csv-note, .ct-csv-left { margin: 0; color: var(--ink-2); }
.ct-csv-cols { max-height: 300px; overflow: auto; border: 1px solid var(--line); border-radius: var(--r-2); }
.ct-csv-row { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) 200px; gap: 8px; align-items: center; padding: 3px 8px; border-bottom: 1px solid var(--line); }
.ct-csv-row:last-child { border-bottom: 0; }
.ct-csv-row.head { position: sticky; top: 0; z-index: 1; background: var(--surface); font-size: 11px; font-weight: 600; color: var(--ink-3); padding: 6px 8px; }
.ct-csv-row .h, .ct-csv-row .v { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ct-csv-row .v { color: var(--ink-3); }
.ct-csv-row.out .h { color: var(--ink-3); }
.ct-csv-cards { display: flex; flex-direction: column; gap: 4px; }
.ct-csv-cards-head { font-weight: 600; }
.ct-csv-card { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 8px; padding: 4px 8px; border-radius: var(--r-2); background: var(--sunken); }
.ct-csv-card > * { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
`;
