// Mailings — mail merge, laid out as Word's own Mailings tab.
//
// Start Mail Merge says what the main document is; Select Recipients gives
// it a list (a workbook, a .csv, the suite's own Contacts, or one typed in
// here); the Write & Insert Fields group puts «fields» in the letter; Preview
// Results shows any record's words in place of them; Finish & Merge makes
// the letters, prints them, or sends each one through Mail. The rules every
// field follows are `@rutba/ooxml/mailmerge`'s — the dialogs below preview
// an Address Block from the very functions the merge itself runs.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Field, Input, Select, Icon, Group, Rows, Separator, Progress } from '@rutba/office-ui';
import {
  MAIN_DOCUMENT_TYPES, NEW_LIST_FIELDS, ADDRESS_NAME_FORMATS, GREETING_NAME_FORMATS, GREETING_SALUTATIONS,
  GREETING_PUNCTUATION, GREETING_FALLBACKS, ADDRESS_FIELDS, addressBlockInstr, greetingLineInstr,
  formatAddressBlock, formatGreeting, mergeOrder, sendMergedMessages,
} from '@rutba/ooxml/mailmerge';
import { PrintDialog } from '../../print.js';

const SOURCE_FILTERS = [
  { name: 'Recipient lists', extensions: ['xlsx', 'xlsm', 'csv', 'tsv', 'txt'] },
  { name: 'Excel workbooks', extensions: ['xlsx', 'xlsm'] },
  { name: 'Comma or tab separated', extensions: ['csv', 'tsv', 'txt'] },
];

/** The country an Address Block leaves off by default: this machine's own, as Word takes it from Windows. */
function homeCountry() {
  try {
    const region = new Intl.Locale(navigator.language || 'en-GB').maximize().region || 'GB';
    return new Intl.DisplayNames(['en'], { type: 'region' }).of(region) || 'United Kingdom';
  } catch {
    return 'United Kingdom';
  }
}

const COMPARISONS = [
  ['=', 'Equal to'], ['<>', 'Not equal to'], ['<', 'Less than'], ['>', 'Greater than'],
  ['<=', 'Less than or equal'], ['>=', 'Greater than or equal'], ['blank', 'Is blank'], ['notBlank', 'Is not blank'],
];

/* ── the ribbon tab ──────────────────────────────────────────────────────── */

/**
 * The Mailings tab, group for group as Word lays it out. Envelopes and Labels
 * lead it in the Create group; everything a merge needs a list for waits,
 * greyed, until Select Recipients has given it one — as in Word.
 */
export function MailingsTab({ mm, menu, create = null }) {
  const s = mm.state;
  const has = Boolean(s?.source);
  const fields = s?.source?.fields || [];
  const [box, setBox] = useState(String(s?.record ?? 1));
  useEffect(() => setBox(String(s?.record ?? 1)), [s?.record]);
  const total = s?.included ?? 0;
  const go = (n) => mm.act('record', n);
  // Greyed until there is a list — and the tip says so, as every idle control here does.
  const need = (tip) => (has ? tip : `${tip}. Select Recipients first.`);

  return (
    <>
      <Group label="Create">
        {create || (
          <>
            <Button tall className="wd-mm-2line" icon="mail" label="Envelopes" disabled title="Envelopes — an envelope's page with the addresses where they print" />
            <Button tall icon="grid" label="Labels" disabled title="Labels — a sheet of address labels to a stock size" />
          </>
        )}
      </Group>
      <Group label="Start Mail Merge">
        <Button
          tall className="wd-mm-2line" icon="mail" label={'Start Mail\nMerge'}
          title={`Start Mail Merge — letters, e-mail messages, envelopes, labels or a directory${s?.type ? ` (now: ${(MAIN_DOCUMENT_TYPES.find((t) => t.id === s.type) || {}).label})` : ''}`}
          onClick={(e) => menu.open(e, [
            ...MAIN_DOCUMENT_TYPES.map((t) => ({
              label: t.id === 'envelopes' || t.id === 'mailingLabels' ? `${t.label}…` : t.label,
              icon: s?.type === t.id ? 'check' : undefined,
              disabled: (t.id === 'envelopes' || t.id === 'mailingLabels') && !mm.canLayout,
              run: () => mm.act('start', t.id),
            })),
            { label: 'Normal Word Document', icon: s?.type ? undefined : 'check', run: () => mm.act('start', null) },
          ])}
        />
        <Button
          tall className="wd-mm-2line" icon="contacts" label={'Select\nRecipients'}
          title="Select Recipients — type a new list, use an existing one, or choose from Contacts"
          onClick={(e) => menu.open(e, [
            { label: 'Type a New List…', icon: 'plus', run: () => mm.act('newList') },
            { label: 'Use an Existing List…', icon: 'open', run: () => mm.act('existingList') },
            { label: 'Choose from Contacts…', icon: 'contacts', run: () => mm.act('contacts') },
          ])}
        />
        <Button tall className="wd-mm-2line" icon="list" label={'Edit\nRecipient List'} disabled={!has} title={need("Edit Recipient List — tick who the merge goes to, sort, filter and find duplicates")} onClick={() => mm.act('editList')} />
      </Group>
      <Group label="Write & Insert Fields">
        <Button tall className="wd-mm-2line" icon="wand" label={'Highlight\nMerge Fields'} pressed={mm.highlight} disabled={!has} title={need("Highlight Merge Fields — shade every «field» grey, so none is missed")} onClick={() => mm.act('highlight')} />
        <Button tall className="wd-mm-2line" icon="textbox" label={'Address\nBlock'} disabled={!has} title={need("Address Block — the recipient's name and address, blank lines left out")} onClick={() => mm.act('addressBlock')} />
        <Button tall className="wd-mm-2line" icon="textbox" label={'Greeting\nLine'} disabled={!has} title={need("Greeting Line — Dear Mr. Randall, or Dear Sir or Madam, when the name is missing")} onClick={() => mm.act('greetingLine')} />
        <Button
          tall className="wd-mm-2line" icon="plus" label={'Insert Merge\nField'} disabled={!has}
          title={need("Insert Merge Field — a column of the list, at the caret")}
          onClick={(e) => menu.open(e, fields.map((f) => ({ label: f, run: () => mm.act('field', f) })))}
        />
        <Rows>
          <Button
            icon="filter" label="Rules" disabled={!has}
            title={need("Rules — If…Then…Else, Next Record, Skip Record If and the record numbers")}
            onClick={(e) => menu.open(e, [
              { label: 'Ask…', disabled: true, title: 'Asks at merge time; not in this build.' },
              { label: 'Fill-in…', disabled: true, title: 'Asks at merge time; not in this build.' },
              { label: 'If…Then…Else…', run: () => mm.act('rule', 'if') },
              { label: 'Merge Record #', run: () => mm.act('rule', 'mergerec') },
              { label: 'Merge Sequence #', run: () => mm.act('rule', 'mergeseq') },
              { label: 'Next Record', run: () => mm.act('rule', 'next') },
              { label: 'Next Record If…', run: () => mm.act('rule', 'nextif') },
              { label: 'Set Bookmark…', disabled: true, title: 'SET fields are not in this build.' },
              { label: 'Skip Record If…', run: () => mm.act('rule', 'skipif') },
            ])}
          />
          <Button icon="link" label="Match Fields" disabled={!has} title={need("Match Fields — which column is the first name, the street, the postcode")} onClick={() => mm.act('match')} />
          <Button icon="refresh" label="Update Labels" disabled={!mm.canUpdateLabels} title={mm.canUpdateLabels ? "Update Labels — copy the first label's fields to every label on the sheet" : "Update Labels — copy the first label's fields to every label on the sheet. Start Mail Merge → Labels first."} onClick={() => mm.act('updateLabels')} />
        </Rows>
      </Group>
      <Group label="Preview Results">
        <Button tall className="wd-mm-2line" icon="eye" label={'Preview\nResults'} pressed={Boolean(s?.preview)} disabled={!has} title={need("Preview Results — see each record's words in place of the «fields»")} onClick={() => mm.act('preview')} />
        <Rows>
          <span className="wd-mm-nav">
            <Button icon="skipBack" title="First record" disabled={!has || (s?.record ?? 1) <= 1} onClick={() => go(1)} />
            <Button icon="chevronLeft" title="Previous record" disabled={!has || (s?.record ?? 1) <= 1} onClick={() => go((s?.record ?? 1) - 1)} />
            <Input
              className="rw-input wd-mm-record"
              value={box}
              disabled={!has}
              title={need("Go to record")}
              aria-label="Record"
              onChange={(e) => setBox(e.target.value.replace(/[^\d]/g, ''))}
              onKeyDown={(e) => { if (e.key === 'Enter') go(Number(box) || 1); }}
              onBlur={() => { if (has && Number(box) !== s?.record) go(Number(box) || 1); }}
            />
            <Button icon="chevronRight" title="Next record" disabled={!has || (s?.record ?? 1) >= total} onClick={() => go((s?.record ?? 1) + 1)} />
            <Button icon="skipForward" title="Last record" disabled={!has || (s?.record ?? 1) >= total} onClick={() => go(total)} />
          </span>
          <Button icon="find" label="Find Recipient" disabled={!has} title={need("Find Recipient — go to the record that holds some words")} onClick={() => mm.act('find')} />
          <Button icon="check" label="Check for Errors" disabled={!has} title={need("Check for Errors — merge fields the list has no column for")} onClick={() => mm.act('errors')} />
        </Rows>
      </Group>
      <Group label="Finish">
        <Button
          tall className="wd-mm-2line" icon="send" label={'Finish &\nMerge'} disabled={!has}
          title={need("Finish & Merge — a document of every letter, print them, or send each as an e-mail")}
          onClick={(e) => menu.open(e, [
            { label: 'Edit Individual Documents…', icon: 'file', run: () => mm.act('finish', 'document') },
            { label: 'Print Documents…', icon: 'print', run: () => mm.act('finish', 'print') },
            { label: 'Send E-mail Messages…', icon: 'send', run: () => mm.act('finish', 'email') },
          ])}
        />
      </Group>
    </>
  );
}

/* ── the state and the verbs ─────────────────────────────────────────────── */

/**
 * Everything the Mailings tab does, for the Word window: the verbs the
 * ribbon names, the dialogs they open, whether the fields are highlighted.
 * `extra` lets the envelopes and labels verbs join the same dispatch.
 */
export function useMailings({ shell, doc, model, apply, toast, extra = null }) {
  const [dialog, setDialog] = useState(null);
  const [highlight, setHighlight] = useState(false);
  const state = model?.mailMerge || null;
  const refresh = useCallback(() => apply({ op: 'mergeRefresh' }), [apply]);
  const call = useCallback((action, args = {}) => shell.doc.mailMerge({ id: doc.id, action, ...args }), [shell, doc]);

  /** Use an Existing List, once a file is chosen: a workbook with several sheets asks which. */
  const useList = useCallback(async ({ path, sheet = null }) => {
    try {
      if (!sheet && /\.xls[xm]$/i.test(path)) {
        const peek = await call('sheets', { path });
        if ((peek.sheets || []).length > 1) {
          setDialog({ name: 'table', path, peek });
          return;
        }
      }
      const got = await call('attach', { path, sheet });
      await refresh();
      toast(`${got.source.count} recipient${got.source.count === 1 ? '' : 's'} from ${got.source.name}${got.source.sheet ? ` — ${got.source.sheet}` : ''}.`, { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    }
  }, [call, refresh, toast]);

  const act = useCallback(async (name, arg) => {
    if (!doc) return;
    if (extra && extra[name]) return extra[name](arg);
    switch (name) {
      case 'start':
        if ((arg === 'envelopes' || arg === 'mailingLabels') && extra?.startLayout) return extra.startLayout(arg);
        await apply({ op: 'startMailMerge', type: arg });
        if (!arg) toast('This is an ordinary document again — no mail merge.', { ms: 4000 });
        return;
      case 'newList':
        setDialog({ name: 'newList' });
        return;
      case 'existingList': {
        const [file] = await shell.dialog.open({ title: 'Select Data Source', filters: SOURCE_FILTERS });
        if (file) await useList({ path: file });
        return;
      }
      case 'useList':
        await useList(arg || {});
        return;
      case 'contacts': {
        const contacts = await shell.contacts.list({}).catch(() => []);
        if (!contacts.length) return toast('The address book is empty. Add people in Contacts first, or type a new list.', { ms: 6000 });
        const got = await call('attachContacts', { contacts });
        await refresh();
        toast(`${got.source.count} recipient${got.source.count === 1 ? '' : 's'} from Contacts.`, { tone: 'good' });
        return;
      }
      case 'editList':
        setDialog({ name: 'recipients' });
        return;
      case 'highlight':
        setHighlight((v) => !v);
        return;
      case 'addressBlock':
        setDialog({ name: 'addressBlock' });
        return;
      case 'greetingLine':
        setDialog({ name: 'greetingLine' });
        return;
      case 'field':
        await apply({ op: 'insertMergeField', name: arg });
        return;
      case 'rule':
        if (arg === 'if' || arg === 'skipif' || arg === 'nextif') setDialog({ name: 'rule', kind: arg });
        else await apply({ op: 'insertMergeRule', kind: arg });
        return;
      case 'match':
        setDialog({ name: 'match' });
        return;
      case 'preview':
        await apply({ op: 'mergePreview', on: !state?.preview });
        return;
      case 'record':
        await apply({ op: 'mergePreview', on: true, record: arg });
        return;
      case 'find':
        setDialog({ name: 'find' });
        return;
      case 'errors': {
        const bad = await call('errors');
        if (!bad.length) toast('No errors: every merge field has a column in the list.', { tone: 'good', ms: 4500 });
        else toast(`Invalid merge field${bad.length === 1 ? '' : 's'}: ${bad.map((b) => `«${b}»`).join(', ')} — the list has no such column. Match Fields or Insert Merge Field again.`, { tone: 'bad', ms: 9000 });
        return;
      }
      case 'finish':
        setDialog({ name: 'finish', to: arg });
        return;
      default:
    }
  }, [doc, extra, apply, toast, shell, call, refresh, useList, state]);

  // A list handed over by name — the OS, or a check run — goes through the
  // same door as the one chosen in the Select Data Source dialog.
  const actRef = useRef(act);
  actRef.current = act;
  useEffect(() => {
    if (!shell?.on) return undefined;
    return shell.on('app:command', ({ command, args }) => {
      if (command === 'mailings.useList' && args?.path) actRef.current('useList', args);
      if (command === 'mailings.finish' && args?.to) actRef.current('finish', args.to);
    });
  }, [shell]);

  // A merge from the address book, opened again: the cards are read afresh.
  const reattached = useRef(false);
  useEffect(() => {
    if (!doc || reattached.current || !state?.pending) return;
    reattached.current = true;
    if (state.pending.contacts) {
      shell.contacts.list({}).then((contacts) => call('attachContacts', { contacts, restore: true })).then(refresh).catch(() => {});
    } else if (state.pending.path) {
      toast(`The recipient list ${state.pending.path} was not found. Select Recipients to choose it again.`, { ms: 8000 });
    }
  }, [doc, state, shell, call, refresh, toast]);

  const close = () => setDialog(null);
  const node = !dialog ? null : (
    <>
      {dialog.name === 'table' ? (
        <SelectTableDialog peek={dialog.peek} path={dialog.path} onClose={close} onPick={(sheet) => { close(); useList({ path: dialog.path, sheet }); }} />
      ) : null}
      {dialog.name === 'newList' ? (
        <NewListDialog
          shell={shell}
          doc={doc}
          onClose={close}
          onSave={async (fields, rows) => {
            const dir = doc?.path ? doc.path.replace(/[\\/][^\\/]*$/, '') : (await shell.app.paths()).documents;
            const target = await shell.dialog.save({ title: 'Save Address List', defaultPath: `${dir}${dir.includes('\\') ? '\\' : '/'}Mailing list.csv`, filters: [{ name: 'Comma separated', extensions: ['csv'] }] });
            if (!target) return false;
            try {
              const got = await call('createList', { path: target, fields, rows });
              await refresh();
              toast(`${got.source.count} recipient${got.source.count === 1 ? '' : 's'} saved to ${got.source.name}.`, { tone: 'good' });
              close();
              return true;
            } catch (err) {
              toast(err.message, { tone: 'bad' });
              return false;
            }
          }}
        />
      ) : null}
      {dialog.name === 'recipients' ? <RecipientsDialog call={call} state={state} onClose={close} onApply={async (spec) => { await apply({ op: 'setMergeRecipients', ...spec }); close(); }} /> : null}
      {dialog.name === 'addressBlock' ? <AddressBlockDialog call={call} state={state} onClose={close} onMatch={() => setDialog({ name: 'match', back: 'addressBlock' })} onInsert={async (spec) => { close(); await apply({ op: 'insertAddressBlock', spec }); }} /> : null}
      {dialog.name === 'greetingLine' ? <GreetingLineDialog call={call} state={state} onClose={close} onInsert={async (spec) => { close(); await apply({ op: 'insertGreetingLine', spec }); }} /> : null}
      {dialog.name === 'rule' ? <RuleDialog kind={dialog.kind} fields={state?.source?.fields || []} onClose={close} onInsert={async (spec) => { close(); await apply({ op: 'insertMergeRule', kind: dialog.kind, spec }); }} /> : null}
      {dialog.name === 'match' ? (
        <MatchFieldsDialog state={state} onClose={() => setDialog(dialog.back ? { name: dialog.back } : null)} onApply={async (overrides) => { await apply({ op: 'setMergeMapping', overrides }); setDialog(dialog.back ? { name: dialog.back } : null); }} />
      ) : null}
      {dialog.name === 'find' ? (
        <FindRecipientDialog
          fields={state?.source?.fields || []}
          onClose={close}
          onFind={async (text, field) => {
            const next = await apply({ op: 'findRecipient', text, field });
            if (!next) return;
            if (!next.opResult) toast('No record holds those words.', { ms: 4000 });
          }}
        />
      ) : null}
      {dialog.name === 'finish' ? <FinishDialog to={dialog.to} shell={shell} call={call} state={state} apply={apply} toast={toast} onClose={close} /> : null}
    </>
  );

  return { state, act, highlight, node, dialog, canLayout: Boolean(extra?.startLayout), canUpdateLabels: Boolean(extra?.updateLabels) && state?.type === 'mailingLabels' };
}

/* ── the records, for a dialog that needs them ───────────────────────────── */

/** The list as Edit Recipient List and the previews read it, fetched once per dialog. */
function useRecords(call) {
  const [data, setData] = useState(null);
  useEffect(() => {
    let live = true;
    call('records').then((d) => live && setData(d)).catch(() => live && setData(null));
    return () => { live = false; };
  }, [call]);
  return data;
}

/** A record box's worth of navigation: ◀ Record 2 of 3 ▶. */
function RecordStepper({ at, total, onGo }) {
  return (
    <div className="wd-mm-stepper">
      <Button icon="chevronLeft" title="Previous record" disabled={at <= 1} onClick={() => onGo(at - 1)} />
      <span>{total ? `Record ${at} of ${total}` : 'No records'}</span>
      <Button icon="chevronRight" title="Next record" disabled={at >= total} onClick={() => onGo(at + 1)} />
    </div>
  );
}

/* ── Select Table ────────────────────────────────────────────────────────── */

function SelectTableDialog({ peek, path, onClose, onPick }) {
  const [sheet, setSheet] = useState(peek.sheet || peek.sheets[0]);
  return (
    <Dialog
      title="Select Table"
      width={440}
      onClose={onClose}
      actions={<><Button label="Cancel" onClick={onClose} /><Button primary label="OK" onClick={() => onPick(sheet)} /></>}
    >
      <p className="wd-mm-lead">{path.split(/[\\/]/).pop()} has more than one sheet. The first row of the one you choose names the fields.</p>
      <div className="wd-mm-pick" role="listbox" aria-label="Sheets">
        {peek.sheets.map((s) => (
          <button key={s} type="button" role="option" aria-selected={s === sheet} className={s === sheet ? 'on' : ''} onClick={() => setSheet(s)} onDoubleClick={() => onPick(s)}>
            <Icon name="sheets" size={14} /> <span>{s}</span>
          </button>
        ))}
      </div>
    </Dialog>
  );
}

/* ── Type a New List ─────────────────────────────────────────────────────── */

function NewListDialog({ onClose, onSave }) {
  const [rows, setRows] = useState(() => [NEW_LIST_FIELDS.map(() => '')]);
  const [at, setAt] = useState(0);
  const [busy, setBusy] = useState(false);
  const set = (r, c, v) => setRows((list) => list.map((row, i) => (i === r ? row.map((x, j) => (j === c ? v : x)) : row)));
  const filled = rows.filter((r) => r.some((v) => v.trim()));
  return (
    <Dialog
      title="New Address List"
      width={900}
      onClose={onClose}
      actions={
        <>
          <Button icon="plus" label="New Entry" onClick={() => { setRows((l) => [...l, NEW_LIST_FIELDS.map(() => '')]); setAt(rows.length); }} />
          <Button icon="trash" label="Delete Entry" disabled={rows.length <= 1} onClick={() => { setRows((l) => l.filter((_, i) => i !== at)); setAt(Math.max(0, at - 1)); }} />
          <span style={{ flex: 1 }} />
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="OK" disabled={!filled.length || busy} onClick={async () => { setBusy(true); const ok = await onSave(NEW_LIST_FIELDS, filled); if (!ok) setBusy(false); }} />
        </>
      }
    >
      <p className="wd-mm-lead">Type recipient information in the table. To add more entries, click New Entry. The list is saved as a .csv file beside your document.</p>
      <div className="wd-mm-grid-wrap">
        <table className="wd-mm-grid wd-mm-newlist">
          <thead><tr><th className="wd-mm-rowno" />{NEW_LIST_FIELDS.map((f) => <th key={f}>{f}</th>)}</tr></thead>
          <tbody>
            {rows.map((row, r) => (
              <tr key={r} className={r === at ? 'on' : ''}>
                <td className="wd-mm-rowno">{r === at ? '▸' : ''}</td>
                {row.map((v, c) => (
                  <td key={c}>
                    <input className="wd-mm-cell" value={v} autoFocus={r === 0 && c === 0} onFocus={() => setAt(r)} onChange={(e) => set(r, c, e.target.value)} aria-label={NEW_LIST_FIELDS[c]} />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="rw-hint">{filled.length} entr{filled.length === 1 ? 'y' : 'ies'}</p>
    </Dialog>
  );
}

/* ── Edit Recipient List ─────────────────────────────────────────────────── */

function RecipientsDialog({ call, state, onClose, onApply }) {
  const data = useRecords(call);
  const [off, setOff] = useState(null);
  const [sort, setSort] = useState(state?.sort || null);
  const [filter, setFilter] = useState('');
  useEffect(() => { if (data && off === null) setOff(new Set(data.excluded || [])); }, [data, off]);
  const fields = data?.fields || [];
  const dupes = useMemo(() => new Set((data?.duplicates || []).flatMap((g) => g.slice(1))), [data]);
  const rows = useMemo(() => {
    if (!data) return [];
    // The table shows every record, in the merge's order when sorted, the
    // filter narrowing what is shown — the ticks are the merge's.
    const all = mergeOrder({ fields: data.fields, records: data.records }, { excluded: [], sort });
    const q = filter.trim().toLowerCase();
    return q ? all.filter((i) => data.records[i].some((v) => String(v).toLowerCase().includes(q))) : all;
  }, [data, sort, filter]);
  const ticked = data && off ? data.records.length - off.size : 0;
  const toggle = (i) => setOff((s) => { const n = new Set(s); if (n.has(i)) n.delete(i); else n.add(i); return n; });
  const sortBy = (f) => setSort((cur) => (cur?.field === f ? (cur.descending ? null : { field: f, descending: true }) : { field: f, descending: false }));
  return (
    <Dialog
      title="Mail Merge Recipients"
      width={900}
      onClose={onClose}
      actions={
        <>
          <span className="wd-mm-foot-note">{data ? `${ticked} of ${data.records.length} recipients ticked` : ''}</span>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="OK" disabled={!data} onClick={() => onApply({ excluded: [...(off || [])], sort })} />
        </>
      }
    >
      <p className="wd-mm-lead">This is the list of recipients that will be used in your merge. Use the checkboxes to add or remove recipients; click a column heading to sort by it.</p>
      <div className="wd-mm-toolbar">
        <span className="wd-mm-filter">
          <Icon name="filter" size={14} />
          <input className="wd-mm-filter-input" placeholder="Filter — show only the records that hold…" value={filter} onChange={(e) => setFilter(e.target.value)} aria-label="Filter" />
        </span>
        <Button icon="check" label="Select all" onClick={() => setOff(new Set())} />
        <Button icon="close" label="Clear all" onClick={() => setOff(new Set(data?.records.map((_, i) => i) || []))} />
        {dupes.size ? (
          <Button icon="flag" label={`Remove ${dupes.size} duplicate${dupes.size === 1 ? '' : 's'}`} title={`Remove duplicates — untick ${dupes.size} record${dupes.size === 1 ? '' : 's'} that repeat an earlier one field for field`} onClick={() => setOff((s) => new Set([...(s || []), ...dupes]))} />
        ) : <span className="wd-mm-dupes-none">No duplicates</span>}
      </div>
      <div className="wd-mm-grid-wrap wd-mm-recipients">
        {!data ? <p className="rw-hint">Reading the list…</p> : (
          <table className="wd-mm-grid">
            <thead>
              <tr>
                <th className="wd-mm-tick"><input type="checkbox" aria-label="All" checked={off?.size === 0} onChange={(e) => setOff(e.target.checked ? new Set() : new Set(data.records.map((_, i) => i)))} /></th>
                {fields.map((f) => (
                  <th key={f} className="wd-mm-sortable" onClick={() => sortBy(f)} title={`Sort by ${f}`}>
                    {f}{sort?.field === f ? (sort.descending ? ' ▼' : ' ▲') : ''}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((i) => (
                <tr key={i} className={`${off?.has(i) ? 'off' : ''}${dupes.has(i) ? ' dupe' : ''}`} data-record={i}>
                  <td className="wd-mm-tick"><input type="checkbox" checked={!off?.has(i)} onChange={() => toggle(i)} aria-label={`Include record ${i + 1}`} /></td>
                  {data.records[i].map((v, c) => <td key={c}>{v}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      <p className="rw-hint wd-mm-source">
        Data source: {data?.name || state?.source?.name}{data?.sheet ? ` — ${data.sheet}` : ''}{dupes.size ? ` · ${dupes.size} possible duplicate${dupes.size === 1 ? '' : 's'} shaded` : ''}
      </p>
    </Dialog>
  );
}

/* ── Insert Address Block ────────────────────────────────────────────────── */

function usePreviewRecord(call, state) {
  const data = useRecords(call);
  const order = useMemo(() => (data ? mergeOrder({ fields: data.fields, records: data.records }, { excluded: data.excluded, sort: data.sort }) : []), [data]);
  const [at, setAt] = useState(state?.record || 1);
  const source = data ? { fields: data.fields, records: data.records } : null;
  const record = source && order.length ? source.records[order[Math.min(order.length, at) - 1]] : null;
  return { source, record, at: Math.min(Math.max(1, at), Math.max(1, order.length)), total: order.length, setAt };
}

function AddressBlockDialog({ call, state, onClose, onInsert, onMatch }) {
  const [name, setName] = useState('titleFirstLastSuffix');
  const [withName, setWithName] = useState(true);
  const [company, setCompany] = useState(true);
  const [postal, setPostal] = useState(true);
  const [country, setCountry] = useState('unlessEqual');
  const [except, setExcept] = useState(homeCountry);
  const p = usePreviewRecord(call, state);
  const spec = { name: withName ? name : null, company, postal, country, except };
  const lines = p.record ? formatAddressBlock(addressBlockInstr(spec), p.source, p.record, state?.mapping || {}) : [];
  return (
    <Dialog
      title="Insert Address Block"
      width={760}
      onClose={onClose}
      actions={
        <>
          <Button icon="link" label="Match Fields…" onClick={onMatch} />
          <span style={{ flex: 1 }} />
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="OK" disabled={!withName && !company && !postal} onClick={() => onInsert(spec)} />
        </>
      }
    >
      <div className="wd-mm-cols">
        <div className="wd-mm-col">
          <div className="wd-mm-section">Specify address elements</div>
          <label className="wd-mm-check"><input type="checkbox" checked={withName} onChange={(e) => setWithName(e.target.checked)} /> Insert recipient's name in this format:</label>
          <div className="wd-mm-pick wd-mm-formats" role="listbox" aria-label="Name format">
            {ADDRESS_NAME_FORMATS.map((f) => (
              <button key={f.id} type="button" role="option" disabled={!withName} aria-selected={f.id === name} className={f.id === name ? 'on' : ''} onClick={() => setName(f.id)}>{f.label}</button>
            ))}
          </div>
          <label className="wd-mm-check"><input type="checkbox" checked={company} onChange={(e) => setCompany(e.target.checked)} /> Insert company name</label>
          <label className="wd-mm-check"><input type="checkbox" checked={postal} onChange={(e) => setPostal(e.target.checked)} /> Insert postal address:</label>
          <div className="wd-mm-radios">
            <label><input type="radio" name="wd-mm-country" disabled={!postal} checked={country === 'never'} onChange={() => setCountry('never')} /> Never include the country/region in the address</label>
            <label><input type="radio" name="wd-mm-country" disabled={!postal} checked={country === 'always'} onChange={() => setCountry('always')} /> Always include the country/region in the address</label>
            <label><input type="radio" name="wd-mm-country" disabled={!postal} checked={country === 'unlessEqual'} onChange={() => setCountry('unlessEqual')} /> Only include the country/region if different than:</label>
            <Input value={except} disabled={!postal || country !== 'unlessEqual'} onChange={(e) => setExcept(e.target.value)} style={{ marginLeft: 22 }} aria-label="Leave off this country" />
          </div>
        </div>
        <div className="wd-mm-col">
          <div className="wd-mm-section">Preview</div>
          <p className="rw-hint" style={{ margin: '0 0 6px' }}>Here is a preview from your recipient list:</p>
          <RecordStepper at={p.at} total={p.total} onGo={p.setAt} />
          <div className="wd-mm-preview wd-mm-address-preview">
            {lines.length ? lines.map((l, i) => <div key={i}>{l}</div>) : <span className="rw-hint">Nothing to show for this record.</span>}
          </div>
          <p className="rw-hint">If items in your address block are missing or out of order, use Match Fields to identify the correct columns from your list.</p>
        </div>
      </div>
    </Dialog>
  );
}

/* ── Insert Greeting Line ────────────────────────────────────────────────── */

function GreetingLineDialog({ call, state, onClose, onInsert }) {
  const [salutation, setSalutation] = useState('Dear');
  const [name, setName] = useState('titleLast');
  const [punctuation, setPunctuation] = useState(',');
  const [fallback, setFallback] = useState('Dear Sir or Madam,');
  const p = usePreviewRecord(call, state);
  const spec = { salutation, name, punctuation, fallback };
  const line = p.record ? formatGreeting(greetingLineInstr(spec), p.source, p.record, state?.mapping || {}) : '';
  return (
    <Dialog
      title="Insert Greeting Line"
      width={560}
      onClose={onClose}
      actions={<><Button label="Cancel" onClick={onClose} /><Button primary label="OK" onClick={() => onInsert(spec)} /></>}
    >
      <div className="wd-mm-section">Greeting line format:</div>
      <div className="wd-mm-row3">
        <Select value={salutation} onChange={(e) => setSalutation(e.target.value)} aria-label="Salutation">
          {GREETING_SALUTATIONS.map((s) => <option key={s} value={s}>{s || '(none)'}</option>)}
        </Select>
        <Select value={name} onChange={(e) => setName(e.target.value)} aria-label="Name format">
          {GREETING_NAME_FORMATS.map((f) => <option key={f.id} value={f.id}>{f.label}</option>)}
        </Select>
        <Select value={punctuation} onChange={(e) => setPunctuation(e.target.value)} aria-label="Punctuation">
          {GREETING_PUNCTUATION.map((s) => <option key={s} value={s}>{s || '(none)'}</option>)}
        </Select>
      </div>
      <div className="wd-mm-section">Greeting line for invalid recipient names:</div>
      <Select value={fallback} onChange={(e) => setFallback(e.target.value)} aria-label="Fallback" style={{ width: '100%' }}>
        {GREETING_FALLBACKS.map((s) => <option key={s} value={s}>{s || '(none)'}</option>)}
      </Select>
      <div className="wd-mm-section">Preview</div>
      <RecordStepper at={p.at} total={p.total} onGo={p.setAt} />
      <div className="wd-mm-preview">{line || <span className="rw-hint">(nothing)</span>}</div>
    </Dialog>
  );
}

/* ── Rules: If…Then…Else, Skip Record If, Next Record If ─────────────────── */

const RULE_TITLES = { if: 'Insert Word Field: IF', skipif: 'Insert Word Field: Skip Record If', nextif: 'Insert Word Field: Next Record If' };

function RuleDialog({ kind, fields, onClose, onInsert }) {
  const [field, setField] = useState(fields[0] || '');
  const [comparison, setComparison] = useState('=');
  const [value, setValue] = useState('');
  const [then, setThen] = useState('');
  const [otherwise, setOtherwise] = useState('');
  const needsValue = comparison !== 'blank' && comparison !== 'notBlank';
  const ok = field && (kind !== 'if' || then || otherwise);
  return (
    <Dialog
      title={RULE_TITLES[kind]}
      width={560}
      onClose={onClose}
      actions={<><Button label="Cancel" onClick={onClose} /><Button primary label="OK" disabled={!ok} onClick={() => onInsert({ field, comparison, value, then, otherwise })} /></>}
    >
      <div className="wd-mm-section">{kind === 'if' ? 'IF' : 'Skip or move on when'}</div>
      <div className="wd-mm-row3">
        <Field label="Field name:">
          <Select value={field} onChange={(e) => setField(e.target.value)}>{fields.map((f) => <option key={f} value={f}>{f}</option>)}</Select>
        </Field>
        <Field label="Comparison:">
          <Select value={comparison} onChange={(e) => setComparison(e.target.value)}>{COMPARISONS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</Select>
        </Field>
        <Field label="Compare to:">
          <Input value={value} disabled={!needsValue} onChange={(e) => setValue(e.target.value)} />
        </Field>
      </div>
      {kind === 'if' ? (
        <>
          <Field label="Insert this text:">
            <textarea className="rw-input wd-mm-text" rows={3} value={then} onChange={(e) => setThen(e.target.value)} />
          </Field>
          <Field label="Otherwise insert this text:">
            <textarea className="rw-input wd-mm-text" rows={3} value={otherwise} onChange={(e) => setOtherwise(e.target.value)} />
          </Field>
        </>
      ) : (
        <p className="rw-hint">{kind === 'skipif' ? 'A record for which this holds is left out of the merge.' : 'When this holds, the fields after this one take the next record — the next label on the sheet.'}</p>
      )}
    </Dialog>
  );
}

/* ── Match Fields ────────────────────────────────────────────────────────── */

function MatchFieldsDialog({ state, onClose, onApply }) {
  const fields = state?.source?.fields || [];
  const [map, setMap] = useState(() => {
    const out = {};
    for (const f of ADDRESS_FIELDS) out[f.token] = state?.mapping?.[f.token] || '';
    return out;
  });
  return (
    <Dialog
      title="Match Fields"
      width={520}
      onClose={onClose}
      actions={<><Button label="Cancel" onClick={onClose} /><Button primary label="OK" onClick={() => onApply(map)} /></>}
    >
      <p className="wd-mm-lead">For each address field the Address Block and the Greeting Line need, choose the column of your list that holds it.</p>
      <div className="wd-mm-match">
        {ADDRESS_FIELDS.map((f) => (
          <React.Fragment key={f.token}>
            <label htmlFor={`wd-mm-m-${f.token}`}>{f.label}</label>
            <Select id={`wd-mm-m-${f.token}`} value={map[f.token]} onChange={(e) => setMap((m) => ({ ...m, [f.token]: e.target.value }))}>
              <option value="">(not matched)</option>
              {fields.map((c) => <option key={c} value={c}>{c}</option>)}
            </Select>
          </React.Fragment>
        ))}
      </div>
    </Dialog>
  );
}

/* ── Find Recipient ──────────────────────────────────────────────────────── */

function FindRecipientDialog({ fields, onClose, onFind }) {
  const [text, setText] = useState('');
  const [where, setWhere] = useState('all');
  const [field, setField] = useState(fields[0] || '');
  return (
    <Dialog
      title="Find Entry"
      width={440}
      onClose={onClose}
      actions={<><Button label="Close" onClick={onClose} /><Button primary label="Find Next" disabled={!text.trim()} onClick={() => onFind(text, where === 'field' ? field : null)} /></>}
    >
      <Field label="Find:">
        <Input value={text} autoFocus onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && text.trim()) onFind(text, where === 'field' ? field : null); }} />
      </Field>
      <div className="wd-mm-radios" style={{ marginTop: 10 }}>
        <label><input type="radio" name="wd-mm-where" checked={where === 'all'} onChange={() => setWhere('all')} /> All fields</label>
        <label className="wd-mm-inline"><input type="radio" name="wd-mm-where" checked={where === 'field'} onChange={() => setWhere('field')} /> This field:
          <Select value={field} disabled={where !== 'field'} onChange={(e) => setField(e.target.value)}>{fields.map((f) => <option key={f} value={f}>{f}</option>)}</Select>
        </label>
      </div>
    </Dialog>
  );
}

/* ── Finish & Merge ──────────────────────────────────────────────────────── */

const FINISH_TITLES = { document: 'Merge to New Document', print: 'Merge to Printer', email: 'Merge to E-mail' };

function FinishDialog({ to, shell, call, state, apply, toast, onClose }) {
  const [range, setRange] = useState('all');
  const [from, setFrom] = useState('1');
  const [until, setUntil] = useState(String(state?.included || 1));
  const [toField, setToField] = useState(state?.email?.toField || (state?.source?.fields || []).find((f) => /mail/i.test(f)) || '');
  const [subject, setSubject] = useState(state?.email?.subject || '');
  const [format, setFormat] = useState('html');
  const [accounts, setAccounts] = useState(null);
  const [account, setAccount] = useState('');
  const [progress, setProgress] = useState(null);
  const [printing, setPrinting] = useState(null);
  const savingPdf = useRef(false);
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    if (to !== 'email') return;
    shell.mail.accounts().then((list) => {
      setAccounts(list || []);
      setAccount((list || [])[0]?.id || '');
    }).catch(() => setAccounts([]));
  }, [shell, to]);
  const spec = range === 'range' ? { from: Number(from) || 1, to: Number(until) || state?.included || 1 } : range;

  const run = async () => {
    setBusy(true);
    try {
      if (to === 'document') {
        const made = await call('finish', { range: spec });
        await shell.win.create({ app: 'word', query: { session: made.id } });
        toast(`${made.name}: ${made.copies} ${made.copies === 1 ? 'copy' : 'copies'} of the ${made.typeLabel.toLowerCase()} — one section each.`, { tone: 'good', ms: 5000 });
        onClose();
      } else if (to === 'print') {
        const made = await call('finish', { range: spec });
        setPrinting(made);
      } else {
        if (!account) { toast('Add a mail account in Mail first.', { tone: 'bad' }); setBusy(false); return; }
        if (!toField) { toast('Choose the column that holds the e-mail addresses.', { tone: 'bad' }); setBusy(false); return; }
        await apply({ op: 'setMergeEmail', toField, subject });
        const messages = await call('messages', { range: spec, toField, subject, format });
        setProgress({ done: 0, total: messages.length, sent: 0, failed: [] });
        const out = await sendMergedMessages(
          messages,
          (m) => shell.mail.send({ accountId: account, draft: { to: m.to, subject: m.subject, text: m.text, ...(m.html ? { html: m.html } : {}) } }),
          (done, total) => setProgress((p) => ({ ...p, done, total })),
        );
        setProgress((p) => ({ ...p, sent: out.sent, failed: out.failed, finished: true }));
        toast(`Sent ${out.sent} of ${messages.length} message${messages.length === 1 ? '' : 's'}${out.failed.length ? ` — ${out.failed.length} could not be sent` : ''}.`, { tone: out.failed.length ? 'bad' : 'good', ms: 6000 });
      }
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    }
    setBusy(false);
  };

  if (printing) {
    // The merged document exists only to be printed: it goes when the print
    // dialog does — unless Save as PDF was pressed, which needs it a moment
    // longer (the dialog closes itself just before it asks).
    const done = () => shell.doc.close({ id: printing.id }).catch(() => {});
    return (
      <PrintDialog
        shell={shell}
        doc={{ id: printing.id, name: printing.name }}
        kind="doc"
        onClose={() => { setTimeout(() => { if (!savingPdf.current) done(); }, 0); onClose(); }}
        onSaveAs={async (options) => {
          savingPdf.current = true;
          try {
            const target = await shell.dialog.save({ title: 'Save as PDF', defaultPath: printing.name.replace(/\.[^.]+$/, '') + '.pdf', filters: [{ name: 'PDF', extensions: ['pdf'] }] });
            if (target) {
              await shell.print.pdf({ id: printing.id, path: target, options });
              toast(`Saved ${target.split(/[\\/]/).pop()}`, { tone: 'good' });
            }
          } catch (err) {
            toast(err.message, { tone: 'bad' });
          }
          done();
        }}
      />
    );
  }

  const finished = progress?.finished;
  return (
    <Dialog
      title={FINISH_TITLES[to]}
      width={to === 'email' ? 560 : 420}
      onClose={onClose}
      actions={finished ? <Button primary label="Close" onClick={onClose} /> : (
        <><Button label="Cancel" onClick={onClose} disabled={busy} /><Button primary label={to === 'email' ? 'Send' : 'OK'} disabled={busy || (to === 'email' && accounts !== null && !accounts.length)} onClick={run} /></>
      )}
    >
      {to === 'email' ? (
        <>
          <div className="wd-mm-section">Message options</div>
          <div className="wd-mm-match">
            <label>To:</label>
            <Select className="rw-select wd-mm-to" value={toField} onChange={(e) => setToField(e.target.value)}>
              <option value="">(choose a column)</option>
              {(state?.source?.fields || []).map((f) => <option key={f} value={f}>{f}</option>)}
            </Select>
            <label>Subject line:</label>
            <Input className="rw-input wd-mm-subject" value={subject} onChange={(e) => setSubject(e.target.value)} />
            <label>Mail format:</label>
            <Select value={format} onChange={(e) => setFormat(e.target.value)}>
              <option value="html">HTML</option>
              <option value="text">Plain text</option>
            </Select>
            <label>Send from:</label>
            <Select className="rw-select wd-mm-account" value={account} onChange={(e) => setAccount(e.target.value)}>
              {accounts === null ? <option>Reading accounts…</option> : accounts.length ? accounts.map((a) => <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>) : <option value="">No mail account — add one in Mail</option>}
            </Select>
          </div>
          <p className="rw-hint">Each message is the merged letter itself. The account's signature is not added: what you wrote in the letter is what each person receives.</p>
        </>
      ) : null}
      <div className="wd-mm-section">{to === 'email' ? 'Send records' : 'Merge records'}</div>
      <div className="wd-mm-radios">
        <label><input type="radio" name="wd-mm-range" checked={range === 'all'} onChange={() => setRange('all')} /> All</label>
        <label><input type="radio" name="wd-mm-range" checked={range === 'current'} onChange={() => setRange('current')} /> Current record</label>
        <label className="wd-mm-inline"><input type="radio" name="wd-mm-range" checked={range === 'range'} onChange={() => setRange('range')} /> From:
          <Input className="rw-input wd-mm-num" value={from} disabled={range !== 'range'} onChange={(e) => setFrom(e.target.value.replace(/[^\d]/g, ''))} aria-label="From record" />
          To:
          <Input className="rw-input wd-mm-num" value={until} disabled={range !== 'range'} onChange={(e) => setUntil(e.target.value.replace(/[^\d]/g, ''))} aria-label="To record" />
        </label>
      </div>
      {progress ? (
        <div className="wd-mm-progress" aria-live="polite">
          <Progress value={progress.done} max={progress.total || 1} />
          <span className="wd-mm-progress-text">{finished ? `Sent ${progress.sent} of ${progress.total}.` : `Sending ${progress.done} of ${progress.total}…`}</span>
          {finished && progress.failed.length ? (
            <ul className="wd-mm-failed">{progress.failed.map((f) => <li key={f.record}>Record {f.record}{f.to ? ` (${f.to})` : ''}: {f.error}</li>)}</ul>
          ) : null}
        </div>
      ) : null}
    </Dialog>
  );
}

/* ── the look ────────────────────────────────────────────────────────────── */

export const MAILINGS_CSS = `
/* Word's Mailings labels sit on two lines under their icons; on one, the tab
   ran past a 1280-wide window and hid Finish & Merge. */
.rw-btn.tall.wd-mm-2line > span { white-space: pre-line; line-height: 1.15; }
.wd-mm-nav { display: inline-flex; align-items: center; gap: 1px; }
.wd-mm-nav .rw-input.wd-mm-record { width: 42px; text-align: center; padding: 3px 4px; margin: 0 3px; }
.wd-mm-lead { margin: 0 0 12px; font-size: 12.5px; color: var(--ink-2); line-height: 1.45; }
.wd-mm-section { font-size: 12px; font-weight: 650; color: var(--ink); margin: 12px 0 7px; }
.wd-mm-section:first-child { margin-top: 2px; }
.wd-mm-cols { display: grid; grid-template-columns: 1.05fr 1fr; gap: 22px; }
.wd-mm-col { min-width: 0; }
.wd-mm-check { display: flex; align-items: center; gap: 7px; font-size: 12.5px; margin: 8px 0 6px; }
.wd-mm-radios { display: flex; flex-direction: column; gap: 6px; font-size: 12.5px; }
.wd-mm-radios label { display: flex; align-items: center; gap: 7px; }
.wd-mm-inline .rw-select, .wd-mm-inline .rw-input { margin-left: 4px; }
.wd-mm-num.rw-input { width: 58px; padding: 4px 6px; }
.wd-mm-pick { display: flex; flex-direction: column; border: 1px solid var(--line); border-radius: var(--r-2); max-height: 196px; overflow: auto; background: var(--surface); }
.wd-mm-pick button { display: flex; align-items: center; gap: 8px; text-align: left; font: inherit; font-size: 12.5px; border: 0; background: transparent; padding: 6px 10px; cursor: pointer; color: var(--ink); }
.wd-mm-pick button:hover:not(:disabled) { background: var(--hover); }
.wd-mm-pick button.on { background: var(--selected-strong); }
.wd-mm-pick button:disabled { color: var(--ink-3); cursor: default; }
.wd-mm-formats { margin-left: 22px; }
.wd-mm-stepper { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-2); margin-bottom: 6px; }
.wd-mm-stepper span { min-width: 104px; text-align: center; }
.wd-mm-preview { border: 1px solid var(--line); border-radius: var(--r-2); background: #fff; color: #111; padding: 12px 14px; min-height: 96px; font-family: Calibri, 'Segoe UI', sans-serif; font-size: 13.5px; line-height: 1.35; margin-bottom: 8px; }
.wd-mm-row3 { display: grid; grid-template-columns: 1fr 1.4fr 1fr; gap: 10px; align-items: end; margin-bottom: 12px; }
.wd-mm-text.rw-input { width: 100%; resize: vertical; min-height: 58px; box-sizing: border-box; font: inherit; }
.wd-mm-match { display: grid; grid-template-columns: 140px 1fr; gap: 8px 12px; align-items: center; font-size: 12.5px; max-height: 420px; overflow: auto; padding-right: 4px; }
.wd-mm-toolbar { display: flex; align-items: center; gap: 6px; margin-bottom: 8px; flex-wrap: wrap; }
.wd-mm-filter { flex: 1; min-width: 220px; display: flex; align-items: center; gap: 7px; border: 1px solid var(--line-strong); border-radius: var(--r-2); padding: 5px 9px; background: var(--surface); color: var(--ink-3); }
.wd-mm-filter:focus-within { border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }
.wd-mm-filter input { border: 0; outline: none; background: transparent; font: inherit; color: var(--ink); flex: 1; min-width: 0; }
.wd-mm-dupes-none { font-size: 11.5px; color: var(--ink-3); padding: 0 6px; }
.wd-mm-grid-wrap { border: 1px solid var(--line); border-radius: var(--r-2); overflow: auto; max-height: 360px; background: var(--surface); }
.wd-mm-grid { border-collapse: separate; border-spacing: 0; width: max-content; min-width: 100%; font-size: 12.5px; }
.wd-mm-grid th { position: sticky; top: 0; z-index: 1; background: var(--surface-2); text-align: left; font-weight: 600; color: var(--ink-2); padding: 7px 10px; border-bottom: 1px solid var(--line); white-space: nowrap; }
.wd-mm-grid td { padding: 5px 10px; border-bottom: 1px solid var(--line-soft); white-space: nowrap; max-width: 240px; overflow: hidden; text-overflow: ellipsis; }
.wd-mm-grid tr.off td:not(.wd-mm-tick) { color: var(--ink-3); }
.wd-mm-grid tr.dupe td { background: color-mix(in srgb, #e0a800 10%, transparent); }
.wd-mm-grid .wd-mm-tick { width: 30px; text-align: center; padding: 5px 6px; position: sticky; left: 0; background: inherit; }
.wd-mm-grid th.wd-mm-tick { z-index: 2; }
.wd-mm-grid td.wd-mm-tick { background: var(--surface); }
.wd-mm-sortable { cursor: pointer; user-select: none; }
.wd-mm-sortable:hover { color: var(--accent); }
.wd-mm-newlist td { padding: 0; border-right: 1px solid var(--line-soft); }
.wd-mm-newlist th { border-right: 1px solid var(--line-soft); }
.wd-mm-newlist tr.on td { background: var(--selected); }
.wd-mm-rowno { width: 22px; text-align: center; color: var(--accent); font-size: 11px; }
.wd-mm-cell { border: 0; background: transparent; font: inherit; font-size: 12.5px; padding: 6px 8px; width: 132px; color: var(--ink); }
.wd-mm-cell:focus { outline: 2px solid var(--accent); outline-offset: -2px; background: var(--surface); }
.wd-mm-source { margin: 8px 0 0; }
.wd-mm-foot-note { margin-right: auto; align-self: center; font-size: 12px; color: var(--ink-2); }
.wd-mm-progress { margin-top: 14px; display: flex; flex-direction: column; gap: 6px; }
.wd-mm-progress-text { font-size: 12.5px; color: var(--ink-2); }
.wd-mm-failed { margin: 4px 0 0; padding-left: 18px; font-size: 12px; color: var(--bad); }
.rw-dialog[aria-label="Mail Merge Recipients"], .rw-dialog[aria-label="New Address List"], .rw-dialog[aria-label="Insert Address Block"] { max-width: min(94vw, 920px); }
`;

let installed = false;
/** The Mailings look, once per window. */
export function installMailingsStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-word-mailings-css';
  style.textContent = MAILINGS_CSS;
  document.head.appendChild(style);
}
