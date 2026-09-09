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
  Ribbon, Group, Button, Separator, Icon, Spacer, Chip, Empty, Panel, Content, Search, Field, Input, Select,
  useToast, useCommands, menuItems, useMenu,
} from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, useFileDrop } from '../shell.js';

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
  const menu = useMenu();
  const [contacts, setContacts] = useState([]);
  const [query, setQuery] = useState('');
  const [selectedId, setSelectedId] = useState(null);
  const [draft, setDraft] = useState(null);
  const [peek, setPeek] = useState(null);
  const [tab, setTab] = useState('home');
  const [busy, setBusy] = useState(false);
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
      toast('A card needs a name, an address or a company.', { tone: 'bad' });
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
      toast(`Kept ${saved.display}.`, { tone: 'good' });
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
    toast(`Removed ${selected.display}.`);
  }, [selected, peek, shell, refresh, toast]);

  const importFile = useCallback(
    async (file = null) => {
      const target = file || (await pickOpen(shell, 'contacts'))?.[0];
      if (!target) return;
      setBusy(true);
      try {
        const r = await shell.contacts.importFile({ path: target });
        setPeek(null);
        await refresh();
        toast(`${r.added} added, ${r.updated} updated${r.same ? `, ${r.same} already here` : ''}.`, { tone: 'good', ms: 6000 });
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 8000 });
      } finally {
        setBusy(false);
      }
    },
    [shell, refresh, toast]
  );

  const exportAll = useCallback(async () => {
    const target = await pickSave(shell, 'contacts', 'contacts.vcf');
    if (!target) return;
    try {
      const r = await shell.contacts.exportFile({ path: target });
      toast(`Wrote ${r.count} card${r.count === 1 ? '' : 's'} to ${target.split(/[\\/]/).pop()}.`, { tone: 'good' });
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
      'contact.new': { label: 'New contact', icon: 'plus', key: 'Mod+N', run: startNew },
      'contact.edit': { label: 'Edit', icon: 'textbox', run: edit },
      'contact.delete': { label: 'Delete', icon: 'trash', key: 'Delete', run: remove },
      'contact.import': { label: 'Import…', icon: 'import', key: 'Mod+O', run: () => importFile() },
      'contact.export': { label: 'Export…', icon: 'export', run: exportAll },
      'contact.mail': { label: 'Send mail', icon: 'send', run: () => writeTo(selected?.emails?.[0]?.value) },
      'contact.find': { label: 'Find', icon: 'find', key: 'Mod+F', run: () => searchRef.current?.focus() },
    }),
    [startNew, edit, remove, importFile, exportAll, writeTo, selected]
  );
  useCommands(commands, [selected, draft]);

  const title = peek ? `${peek.name} (${peek.contacts.length})` : `${contacts.length} contact${contacts.length === 1 ? '' : 's'}`;

  /* ── the card ─────────────────────────────────────────────────────────── */

  const card = selected && !draft ? (
    <div className="ct-card">
      <div className="ct-head">
        <Avatar contact={selected} size={72} />
        <div className="ct-who">
          <h2>{selected.display}</h2>
          {selected.title || selected.org ? <div className="ct-sub">{[selected.title, selected.org].filter(Boolean).join(' · ')}</div> : null}
          {selected.categories?.length ? <div className="ct-tags">{selected.categories.map((c) => <Chip key={c}>{c}</Chip>)}</div> : null}
        </div>
      </div>
      {peek ? (
        <div className="ct-offer">
          <Icon name="info" size={14} />
          <span>From {peek.name}, not yet in your contacts.</span>
          <Button primary label={`Add ${peek.contacts.length === 1 ? 'this card' : `all ${peek.contacts.length}`}`} onClick={() => importFile(peek.path)} />
        </div>
      ) : null}
      <dl className="ct-fields">
        {(selected.emails || []).map((e, i) => (
          <React.Fragment key={`e${i}`}>
            <dt>{e.label || e.type || 'email'}</dt>
            <dd><button type="button" className="ct-link" onClick={() => writeTo(e.value)} title="Write a message">{e.value}</button></dd>
          </React.Fragment>
        ))}
        {(selected.phones || []).map((p, i) => (
          <React.Fragment key={`p${i}`}>
            <dt>{p.label || p.type || 'phone'}</dt>
            <dd>{p.value}</dd>
          </React.Fragment>
        ))}
        {(selected.addresses || []).map((a, i) => (
          <React.Fragment key={`a${i}`}>
            <dt>{a.label || a.type || 'address'}</dt>
            <dd className="ct-addr">{[a.street, a.city, a.region, a.postcode, a.country].filter(Boolean).join('\n')}</dd>
          </React.Fragment>
        ))}
        {selected.birthday ? (<><dt>birthday</dt><dd>{selected.birthday}</dd></>) : null}
        {(selected.urls || []).map((u, i) => (
          <React.Fragment key={`u${i}`}>
            <dt>{u.label || 'web'}</dt>
            <dd><a href={u.value} onClick={(e) => { e.preventDefault(); shell.shell.openExternal({ url: u.value }); }}>{u.value}</a></dd>
          </React.Fragment>
        ))}
        {selected.note ? (<><dt>notes</dt><dd className="ct-note">{selected.note}</dd></>) : null}
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
          <Input className="rw-input ct-name" value={draft.name.full} placeholder="Name" autoFocus onChange={(e) => { const full = e.target.value; setDraft((d) => ({ ...d, name: { ...d.name, full, ...splitTyped(full) } })); }} />
          <div className="ct-row2">
            <Input value={draft.title} placeholder="Job title" onChange={(e) => set({ title: e.target.value })} />
            <Input value={draft.org} placeholder="Company" onChange={(e) => set({ org: e.target.value })} />
          </div>
        </div>
      </div>
      <div className="ct-section">
        <h4>Email</h4>
        {draft.emails.map((e, i) => (
          <div className="ct-line" key={i}>
            <Select value={e.type} onChange={(ev) => setRow('emails', i, { type: ev.target.value })}>{TYPES.email.map((t) => <option key={t} value={t}>{t}</option>)}</Select>
            <Input value={e.value} placeholder="name@example.com" onChange={(ev) => setRow('emails', i, { value: ev.target.value })} />
            <Button icon="close" title="Remove" onClick={() => dropRow('emails', i)} />
          </div>
        ))}
        <Button ghost icon="plus" label="Add email" onClick={() => addRow('emails', { value: '', type: 'home', pref: false, label: null })} />
      </div>
      <div className="ct-section">
        <h4>Phone</h4>
        {draft.phones.map((p, i) => (
          <div className="ct-line" key={i}>
            <Select value={p.type} onChange={(ev) => setRow('phones', i, { type: ev.target.value })}>{TYPES.phone.map((t) => <option key={t} value={t}>{t}</option>)}</Select>
            <Input value={p.value} placeholder="+44 7700 900123" onChange={(ev) => setRow('phones', i, { value: ev.target.value })} />
            <Button icon="close" title="Remove" onClick={() => dropRow('phones', i)} />
          </div>
        ))}
        <Button ghost icon="plus" label="Add phone" onClick={() => addRow('phones', { value: '', type: 'work', pref: false, label: null })} />
      </div>
      <div className="ct-section">
        <h4>Address</h4>
        {draft.addresses.map((a, i) => (
          <div className="ct-address" key={i}>
            <div className="ct-line">
              <Select value={a.type} onChange={(ev) => setRow('addresses', i, { type: ev.target.value })}>{TYPES.address.map((t) => <option key={t} value={t}>{t}</option>)}</Select>
              <Input value={a.street} placeholder="Street" onChange={(ev) => setRow('addresses', i, { street: ev.target.value })} />
              <Button icon="close" title="Remove" onClick={() => dropRow('addresses', i)} />
            </div>
            <div className="ct-line four">
              <Input value={a.city} placeholder="City" onChange={(ev) => setRow('addresses', i, { city: ev.target.value })} />
              <Input value={a.region} placeholder="Region" onChange={(ev) => setRow('addresses', i, { region: ev.target.value })} />
              <Input value={a.postcode} placeholder="Postcode" onChange={(ev) => setRow('addresses', i, { postcode: ev.target.value })} />
              <Input value={a.country} placeholder="Country" onChange={(ev) => setRow('addresses', i, { country: ev.target.value })} />
            </div>
          </div>
        ))}
        <Button ghost icon="plus" label="Add address" onClick={() => addRow('addresses', { street: '', city: '', region: '', postcode: '', country: '', type: 'home', pref: false, label: null })} />
      </div>
      <div className="ct-section ct-two">
        <Field label="Birthday"><Input type="date" value={draft.birthday || ''} onChange={(e) => set({ birthday: e.target.value })} /></Field>
        <Field label="Categories" hint="Comma-separated"><Input value={(draft.categories || []).join(', ')} onChange={(e) => set({ categories: e.target.value.split(',').map((s) => s.trim()).filter(Boolean) })} /></Field>
      </div>
      <div className="ct-section">
        <Field label="Notes"><textarea className="rw-input ct-notes" rows={4} value={draft.note || ''} onChange={(e) => set({ note: e.target.value })} /></Field>
      </div>
      <div className="ct-actions">
        <Button label="Cancel" onClick={() => setDraft(null)} />
        <Button primary label={busy ? 'Saving…' : 'Save'} disabled={busy} onClick={saveDraft} />
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
          tabs={[{ id: 'home', label: 'Home' }, { id: 'view', label: 'View' }]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="plus" title="New contact (Ctrl+N)" onClick={startNew} />
              <Button icon="trash" title="Delete" onClick={remove} disabled={!selected || Boolean(peek)} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label="New">
                <Button tall icon="plus" label="New contact" onClick={startNew} />
              </Group>
              <Group label="Card">
                <Button tall icon="textbox" label="Edit" onClick={edit} disabled={!selected || Boolean(peek)} />
                <Button tall icon="trash" label="Delete" onClick={remove} disabled={!selected || Boolean(peek)} />
                <Button tall icon="send" label="Send mail" onClick={() => writeTo(selected?.emails?.[0]?.value)} disabled={!selected?.emails?.length} />
              </Group>
              <Group label="Files">
                <Button tall icon="import" label="Import" title="A .vcf or .csv from another program" onClick={() => importFile()} />
                <Button tall icon="export" label="Export" title="All cards as a .vcf" onClick={exportAll} disabled={!contacts.length} />
              </Group>
            </>
          ) : (
            <Group label="Find">
              <Button tall icon="find" label="Find" onClick={() => searchRef.current?.focus()} />
            </Group>
          )}
        </Ribbon>
      }
    >
      <style>{CSS}</style>
      <Panel width={280} resizable>
        <div className="ct-search">
          <Search value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search contacts" autoFocus={false} />
        </div>
        {peek ? (
          <div className="ct-peek">
            <Icon name="file" size={14} />
            <span className="grow">{peek.name}</span>
            <Button ghost icon="close" title="Back to your contacts" onClick={() => { setPeek(null); setSelectedId(null); }} />
          </div>
        ) : null}
        <div className="ct-list" role="listbox">
          {groups.map((g) => (
            <React.Fragment key={g.letter}>
              <div className="ct-letter">{g.letter}</div>
              {g.items.map((c) => (
                <button
                  type="button"
                  key={c.id}
                  className={`ct-item${c.id === selectedId ? ' current' : ''}`}
                  onClick={() => { setSelectedId(c.id); setDraft(null); }}
                  onContextMenu={(e) => { setSelectedId(c.id); menu.open(e, menuItems(commands, ['contact.edit', 'contact.mail', '-', 'contact.delete'])); }}
                >
                  <Avatar contact={c} size={30} />
                  <span className="ct-item-text">
                    <span className="name">{c.display || '(no name)'}</span>
                    <span className="sub">{c.emails?.[0]?.value || c.org || c.phones?.[0]?.value || ''}</span>
                  </span>
                </button>
              ))}
            </React.Fragment>
          ))}
          {!list.length ? <div className="ct-none">{query ? 'Nothing matches.' : peek ? 'The file holds no cards.' : 'No contacts yet.'}</div> : null}
        </div>
      </Panel>
      <Content>
        {editor || card || (
          <Empty icon="contacts" title={contacts.length ? 'Choose a contact' : 'No contacts yet'}>
            {contacts.length ? 'Pick one on the left, or press Ctrl+N for a new card.' : 'Add a card, or bring in the .vcf or .csv another program exported.'}
            <div style={{ marginTop: 14, display: 'flex', gap: 8, justifyContent: 'center' }}>
              <Button primary icon="plus" label="New contact" onClick={startNew} />
              <Button icon="import" label="Import…" onClick={() => importFile()} />
            </div>
          </Empty>
        )}
      </Content>
    </AppFrame>
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

function Avatar({ contact, size }) {
  const style = { width: size, height: size, fontSize: Math.round(size * 0.38) };
  if (contact.photoUrl) return <img className="ct-avatar" src={contact.photoUrl} alt="" style={style} />;
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
.ct-link { border: 0; background: none; padding: 0; font: inherit; color: var(--accent); cursor: pointer; text-align: left; }
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
`;
