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
import { Button, Dialog, Field, Input, Select, Icon, Group, Rows, Separator, Progress, t, tn } from '@rutba/office-ui';
import {
  MAIN_DOCUMENT_TYPES, NEW_LIST_FIELDS, ADDRESS_NAME_FORMATS, GREETING_NAME_FORMATS, GREETING_SALUTATIONS,
  GREETING_PUNCTUATION, GREETING_FALLBACKS, ADDRESS_FIELDS, addressBlockInstr, greetingLineInstr,
  formatAddressBlock, formatGreeting, mergeOrder, sendMergedMessages,
} from '@rutba/ooxml/mailmerge';
import { PrintDialog } from '../../print.js';

const SOURCE_FILTERS = [
  { name: t('word.mailings.recipient-lists'), extensions: ['xlsx', 'xlsm', 'rbase', 'csv', 'tsv', 'txt'] },
  { name: t('word.mailings.excel-workbooks'), extensions: ['xlsx', 'xlsm'] },
  { name: t('word.mailings.data-files-databases'), extensions: ['rbase'] },
  { name: t('word.mailings.comma-tab-separated'), extensions: ['csv', 'tsv', 'txt'] },
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
  ['=', t('common.equal-to')], ['<>', t('common.not-equal-to')], ['<', t('common.less-than')], ['>', t('common.greater-than')],
  ['<=', t('word.mailings.less-equal')], ['>=', t('word.mailings.greater-equal')], ['blank', t('word.mailings.is-blank')], ['notBlank', t('word.mailings.is-not-blank')],
];

const TYPE_NAMES = { formLetters: t('common.letters'), email: t('word.mailings.e-mail-messages'), envelopes: t('word.envelopes'), mailingLabels: t('common.labels'), catalog: t('word.directory') };
const TYPE_NAMES_LOWER = { formLetters: t('word.mailings.letters'), email: t('word.mailings.e-mail-messages-2'), envelopes: t('word.mailings.envelopes'), mailingLabels: t('word.mailings.labels'), catalog: t('word.mailings.directory') };

const ADDRESS_FIELD_NAMES = {
  _TITLE0_: t('word.mailings.courtesy-title'), _FIRST0_: t('word.mailings.first-name'), _MIDDLE0_: t('word.mailings.middle-name'), _LAST0_: t('word.mailings.last-name'), _SUFFIX0_: t('word.mailings.suffix'), _NICK0_: t('common.nickname'),
  _COMPANY_: t('common.company'), _STREET1_: t('word.mailings.address-1'), _STREET2_: t('word.mailings.address-2'), _CITY_: t('common.city'), _STATE_: t('word.mailings.state'), _POSTAL_: t('word.mailings.postal-code'),
  _COUNTRY_: t('word.mailings.country-or-region'), _EMAIL_: t('word.mailings.e-mail-address'),
};

const NEW_LIST_NAMES = {
  'Title': t('common.title'), 'First Name': t('word.mailings.first-name'), 'Last Name': t('word.mailings.last-name'), 'Company Name': t('word.mailings.company-name'), 'Address Line 1': t('word.mailings.address-line-1'), 'Address Line 2': t('word.mailings.address-line-2'),
  'City': t('common.city'), 'State': t('word.mailings.state'), 'ZIP Code': t('word.mailings.zip-code'), 'Country or Region': t('word.mailings.country-or-region'), 'Home Phone': t('word.mailings.home-phone'), 'Work Phone': t('word.mailings.work-phone'), 'E-mail Address': t('word.mailings.e-mail-address'),
};

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
  const need = (tip) => (has ? tip : t('word.mailings.tip-select-recipients', { tip }));

  return (
    <>
      <Group label={t('common.create')}>
        {create || (
          <>
            <Button tall icon="mail" label={t('word.envelopes')} title={t('word.mailings.envelopes-delivery')} onClick={() => mm.act('envelopes')} />
            <Button tall icon="grid" label={t('common.labels')} title={t('word.mailings.labels-sheet-same')} onClick={() => mm.act('labels')} />
          </>
        )}
      </Group>
      <Group label={t('word.mailings.start-mail-merge')}>
        <Button dropdown
          tall className="wd-mm-2line" icon="mail" label={t('word.mailings.start-mail')}
          title={s?.type ? t('word.mailings.start-mail-merge-letters-e', { type: TYPE_NAMES[s.type] || (MAIN_DOCUMENT_TYPES.find((type) => type.id === s.type) || {}).label }) : t('word.mailings.start-mail-merge-letters')}
          onClick={(e) => menu.open(e, [
            ...MAIN_DOCUMENT_TYPES.map((type) => ({
              label: type.id === 'envelopes' ? t('word.mailings.envelopes-2') : type.id === 'mailingLabels' ? t('word.mailings.labels-2') : TYPE_NAMES[type.id] || type.label,
              icon: s?.type === type.id ? 'check' : undefined,
              disabled: (type.id === 'envelopes' || type.id === 'mailingLabels') && !mm.canLayout,
              run: () => mm.act('start', type.id),
            })),
            { label: t('word.mailings.normal-word-document'), icon: s?.type ? undefined : 'check', run: () => mm.act('start', null) },
          ])}
        />
        <Button dropdown
          tall className="wd-mm-2line" icon="contacts" label={t('word.mailings.select')}
          title={t('word.mailings.select-recipients-type')}
          onClick={(e) => menu.open(e, [
            { label: t('word.mailings.type-new-list'), icon: 'plus', run: () => mm.act('newList') },
            { label: t('word.mailings.use-existing-list'), icon: 'open', run: () => mm.act('existingList') },
            { label: t('word.mailings.choose-from'), icon: 'contacts', run: () => mm.act('contacts') },
          ])}
        />
        <Button tall className="wd-mm-2line" icon="list" label={t('word.mailings.edit')} disabled={!has} title={need(t('word.mailings.edit-recipient-list'))} onClick={() => mm.act('editList')} />
      </Group>
      <Group label={t('word.mailings.write-insert')}>
        <Button tall className="wd-mm-2line" icon="wand" label={t('word.mailings.highlight')} pressed={mm.highlight} disabled={!has} title={need(t('word.mailings.highlight-merge-fields'))} onClick={() => mm.act('highlight')} />
        <Button tall className="wd-mm-2line" icon="textbox" label={t('word.mailings.address')} disabled={!has} title={need(t('word.mailings.address-block-recipients'))} onClick={() => mm.act('addressBlock')} />
        <Button tall className="wd-mm-2line" icon="textbox" label={t('word.mailings.greeting')} disabled={!has} title={need(t('word.mailings.greeting-line-dear'))} onClick={() => mm.act('greetingLine')} />
        <Button dropdown
          tall className="wd-mm-2line" icon="plus" label={t('word.mailings.insert-merge')} disabled={!has}
          title={need(t('word.mailings.insert-merge-field'))}
          onClick={(e) => menu.open(e, fields.map((f) => ({ label: f, run: () => mm.act('field', f) })))}
        />
        <Rows>
          <Button dropdown
            icon="filter" label={t('common.rules')} disabled={!has}
            title={need(t('word.mailings.rules-else-next'))}
            onClick={(e) => menu.open(e, [
              { label: t('word.mailings.ask'), run: () => mm.act('rule', 'ask'), title: t('word.mailings.ask-question-put') },
              { label: t('word.mailings.fill-in'), run: () => mm.act('rule', 'fillin'), title: t('word.mailings.fill-question-put') },
              { label: t('word.mailings.if-then-else'), run: () => mm.act('rule', 'if') },
              { label: t('word.mailings.merge-record'), run: () => mm.act('rule', 'mergerec') },
              { label: t('word.mailings.merge-sequence'), run: () => mm.act('rule', 'mergeseq') },
              { label: t('word.mailings.next-record'), run: () => mm.act('rule', 'next') },
              { label: t('word.mailings.next-record-if'), run: () => mm.act('rule', 'nextif') },
              { label: t('word.mailings.set-bookmark'), run: () => mm.act('rule', 'set'), title: t('word.mailings.set-bookmark-bookmark') },
              { label: t('word.mailings.skip-record-if'), run: () => mm.act('rule', 'skipif') },
            ])}
          />
          <Button icon="link" label={t('word.mailings.match-fields')} disabled={!has} title={need(t('word.mailings.match-fields-column'))} onClick={() => mm.act('match')} />
          <Button icon="refresh" label={t('word.mailings.update')} disabled={!mm.canUpdateLabels} title={mm.canUpdateLabels ? t('word.mailings.update-labels-copy') : t('word.mailings.update-labels-copy-first')} onClick={() => mm.act('updateLabels')} />
        </Rows>
      </Group>
      <Group label={t('word.mailings.preview-results')}>
        <Button tall className="wd-mm-2line" icon="eye" label={t('word.mailings.preview')} pressed={Boolean(s?.preview)} disabled={!has} title={need(t('word.mailings.preview-results-see'))} onClick={() => mm.act('preview')} />
        <Rows>
          <span className="wd-mm-nav">
            <Button icon="skipBack" title={t('common.first-record')} disabled={!has || (s?.record ?? 1) <= 1} onClick={() => go(1)} />
            <Button icon="chevronLeft" title={t('common.previous-record')} disabled={!has || (s?.record ?? 1) <= 1} onClick={() => go((s?.record ?? 1) - 1)} />
            <Input
              className="rw-input wd-mm-record"
              value={box}
              disabled={!has}
              title={need(t('word.mailings.go-to-record'))}
              aria-label={t('common.record')}
              onChange={(e) => setBox(e.target.value.replace(/[^\d]/g, ''))}
              onKeyDown={(e) => { if (e.key === 'Enter') go(Number(box) || 1); }}
              onBlur={() => { if (has && Number(box) !== s?.record) go(Number(box) || 1); }}
            />
            <Button icon="chevronRight" title={t('common.next-record')} disabled={!has || (s?.record ?? 1) >= total} onClick={() => go((s?.record ?? 1) + 1)} />
            <Button icon="skipForward" title={t('common.last-record')} disabled={!has || (s?.record ?? 1) >= total} onClick={() => go(total)} />
          </span>
          <Button icon="find" label={t('word.mailings.find-recipient')} disabled={!has} title={need(t('word.mailings.find-recipient-go'))} onClick={() => mm.act('find')} />
          <Button icon="check" label={t('word.mailings.check-for')} disabled={!has} title={need(t('word.mailings.check-errors-merge'))} onClick={() => mm.act('errors')} />
        </Rows>
      </Group>
      <Group label={t('common.finish')}>
        <Button dropdown
          tall className="wd-mm-2line" icon="send" label={t('word.mailings.finish-and')} disabled={!has}
          title={need(t('word.mailings.finish-merge-document'))}
          onClick={(e) => menu.open(e, [
            { label: t('word.mailings.edit-individual'), icon: 'file', run: () => mm.act('finish', 'document') },
            { label: t('word.mailings.print-documents'), icon: 'print', run: () => mm.act('finish', 'print') },
            { label: t('word.mailings.send-e-mail'), icon: 'send', run: () => mm.act('finish', 'email') },
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

  /** Use an Existing List, once a file is chosen: a workbook with several sheets, or a database with several tables and queries, asks which. */
  const useList = useCallback(async ({ path, sheet = null }) => {
    try {
      if (!sheet && /\.(xls[xm]|rbase)$/i.test(path)) {
        const peek = await call('sheets', { path });
        if ((peek.sheets || []).length > 1) {
          setDialog({ name: 'table', path, peek });
          return;
        }
      }
      const got = await call('attach', { path, sheet });
      await refresh();
      toast(got.source.sheet ? tn(got.source.count, 'word.mailings.count-recipient-name-sheet', { name: got.source.name, sheet: got.source.sheet }) : tn(got.source.count, 'word.mailings.count-recipient-name', { name: got.source.name }), { tone: 'good' });
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
        if (!arg) toast(t('word.mailings.ordinary-document-again'), { ms: 4000 });
        return;
      case 'newList':
        setDialog({ name: 'newList' });
        return;
      case 'existingList': {
        const [file] = await shell.dialog.open({ title: t('word.mailings.select-data-source'), filters: SOURCE_FILTERS });
        if (file) await useList({ path: file });
        return;
      }
      case 'useList':
        await useList(arg || {});
        return;
      case 'contacts': {
        const contacts = await shell.contacts.list({}).catch(() => []);
        if (!contacts.length) return toast(t('word.mailings.address-book-empty'), { ms: 6000 });
        const got = await call('attachContacts', { contacts });
        await refresh();
        toast(tn(got.source.count, 'word.mailings.count-recipient-contacts'), { tone: 'good' });
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
        else if (arg === 'ask' || arg === 'fillin' || arg === 'set') setDialog({ name: 'prompt', kind: arg });
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
        if (!bad.length) toast(t('word.mailings.no-errors-every'), { tone: 'good', ms: 4500 });
        else toast(tn(bad.length, 'word.mailings.invalid-merge-field', { fields: bad.map((b) => `«${b}»`).join(', ') }), { tone: 'bad', ms: 9000 });
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
    } else if (state.pending.path && state.pending.network) {
      toast(t('word.mailings.letters-recipient-list', { path: state.pending.path }), { ms: 9000 });
    } else if (state.pending.path && !state.pending.found) {
      toast(t('common.recipient-list-path', { path: state.pending.path }), { ms: 8000 });
    } else if (state.pending.path) {
      // Asked, not done: the path came from the file, as Word asks before
      // it opens a letter's data source.
      const name = state.pending.path.split(/[\\/]/).pop();
      toast(t('word.mailings.letter-takes-recipients', { name }), {
        ms: 15000,
        action: { label: t('common.attach'), run: () => call('reattach').then(refresh).catch((err) => toast(err?.message || String(err), { ms: 8000 })) },
      });
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
            const target = await shell.dialog.save({ title: t('word.mailings.save-address-list'), defaultPath: `${dir}${dir.includes('\\') ? '\\' : '/'}${t('common.mailing-list')}.csv`, filters: [{ name: t('word.mailings.comma-separated'), extensions: ['csv'] }] });
            if (!target) return false;
            try {
              const got = await call('createList', { path: target, fields, rows });
              await refresh();
              toast(tn(got.source.count, 'word.mailings.count-recipient-saved', { name: got.source.name }), { tone: 'good' });
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
      {dialog.name === 'prompt' ? <PromptRuleDialog kind={dialog.kind} onClose={close} onInsert={async (spec) => { close(); await apply({ op: 'insertMergeRule', kind: dialog.kind, spec }); }} /> : null}
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
            if (!next.opResult) toast(t('word.mailings.no-record-holds'), { ms: 4000 });
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
      <Button icon="chevronLeft" title={t('common.previous-record')} disabled={at <= 1} onClick={() => onGo(at - 1)} />
      <span>{total ? t('word.mailings.record-total', { at, total }) : t('word.mailings.no-records')}</span>
      <Button icon="chevronRight" title={t('common.next-record')} disabled={at >= total} onClick={() => onGo(at + 1)} />
    </div>
  );
}

/* ── Select Table ────────────────────────────────────────────────────────── */

function SelectTableDialog({ peek, path, onClose, onPick }) {
  const [sheet, setSheet] = useState(peek.sheet || peek.sheets[0]);
  const database = /\.rbase$/i.test(path);
  return (
    <Dialog
      title={t('word.mailings.select-table')}
      width={440}
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.ok')} onClick={() => onPick(sheet)} /></>}
    >
      <p className="wd-mm-lead">{database
        ? t('word.mailings.name-more-one-table', { name: path.split(/[\\/]/).pop() })
        : t('word.mailings.name-more-one', { name: path.split(/[\\/]/).pop() })}</p>
      <div className="wd-mm-pick" role="listbox" aria-label={database ? t('word.mailings.tables-and-queries') : t('common.sheets')}>
        {peek.sheets.map((s) => (
          <button key={s} type="button" role="option" aria-selected={s === sheet} className={s === sheet ? 'on' : ''} data-sheet={s} onClick={() => setSheet(s)} onDoubleClick={() => onPick(s)}>
            <Icon name={database ? 'table' : 'sheets'} size={14} /> <span>{s}</span>
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
      title={t('word.mailings.new-address-list')}
      width={900}
      onClose={onClose}
      className="wd-mm-wide"
      actions={
        <>
          <Button icon="plus" label={t('word.mailings.new-entry')} onClick={() => { setRows((l) => [...l, NEW_LIST_FIELDS.map(() => '')]); setAt(rows.length); }} />
          <Button icon="trash" label={t('word.mailings.delete-entry')} disabled={rows.length <= 1} onClick={() => { setRows((l) => l.filter((_, i) => i !== at)); setAt(Math.max(0, at - 1)); }} />
          <span style={{ flex: 1 }} />
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.ok')} disabled={!filled.length || busy} onClick={async () => { setBusy(true); const ok = await onSave(NEW_LIST_FIELDS, filled); if (!ok) setBusy(false); }} />
        </>
      }
    >
      <p className="wd-mm-lead">{t('word.mailings.type-recipient')}</p>
      <div className="wd-mm-grid-wrap">
        <table className="wd-mm-grid wd-mm-newlist">
          <thead><tr><th className="wd-mm-rowno" />{NEW_LIST_FIELDS.map((f) => <th key={f}>{NEW_LIST_NAMES[f] || f}</th>)}</tr></thead>
          <tbody>
            {rows.map((row, r) => (
              <tr key={r} className={r === at ? 'on' : ''}>
                <td className="wd-mm-rowno">{r === at ? '▸' : ''}</td>
                {row.map((v, c) => (
                  <td key={c}>
                    <input className="wd-mm-cell" value={v} autoFocus={r === 0 && c === 0} onFocus={() => setAt(r)} onChange={(e) => set(r, c, e.target.value)} aria-label={NEW_LIST_NAMES[NEW_LIST_FIELDS[c]] || NEW_LIST_FIELDS[c]} />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="rw-hint">{tn(filled.length, 'word.mailings.count-entry')}</p>
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
      title={t('word.mailings.mail-merge')}
      width={900}
      onClose={onClose}
      className="wd-mm-wide"
      actions={
        <>
          <span className="wd-mm-foot-note">{data ? t('word.mailings.ticked-total-recipients', { ticked, total: data.records.length }) : ''}</span>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.ok')} disabled={!data} onClick={() => onApply({ excluded: [...(off || [])], sort })} />
        </>
      }
    >
      <p className="wd-mm-lead">{t('word.mailings.list-recipients-used')}</p>
      <div className="wd-mm-toolbar">
        <span className="wd-mm-filter">
          <Icon name="filter" size={14} />
          <input className="wd-mm-filter-input" placeholder={t('word.mailings.filter-show-only')} value={filter} onChange={(e) => setFilter(e.target.value)} aria-label={t('common.filter')} />
        </span>
        <Button icon="check" label={t('common.select-all')} onClick={() => setOff(new Set())} />
        <Button icon="close" label={t('common.clear-all')} onClick={() => setOff(new Set(data?.records.map((_, i) => i) || []))} />
        {dupes.size ? (
          <Button icon="flag" label={tn(dupes.size, 'word.mailings.remove-count')} title={tn(dupes.size, 'word.mailings.remove-duplicates-untick')} onClick={() => setOff((s) => new Set([...(s || []), ...dupes]))} />
        ) : <span className="wd-mm-dupes-none">{t('word.mailings.no-duplicates')}</span>}
      </div>
      <div className="wd-mm-grid-wrap wd-mm-recipients">
        {!data ? <p className="rw-hint">{t('word.mailings.reading-the-list')}</p> : (
          <table className="wd-mm-grid">
            <thead>
              <tr>
                <th className="wd-mm-tick"><input type="checkbox" aria-label={t('common.all')} checked={off?.size === 0} onChange={(e) => setOff(e.target.checked ? new Set() : new Set(data.records.map((_, i) => i)))} /></th>
                {fields.map((f) => (
                  <th key={f} className="wd-mm-sortable" onClick={() => sortBy(f)} title={t('word.mailings.sort-by-field', { field: f })}>
                    {f}{sort?.field === f ? (sort.descending ? ' ▼' : ' ▲') : ''}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((i) => (
                <tr key={i} className={`${off?.has(i) ? 'off' : ''}${dupes.has(i) ? ' dupe' : ''}`} data-record={i}>
                  <td className="wd-mm-tick"><input type="checkbox" checked={!off?.has(i)} onChange={() => toggle(i)} aria-label={t('word.mailings.include-record-number', { number: i + 1 })} /></td>
                  {data.records[i].map((v, c) => <td key={c}>{v}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      <p className="rw-hint wd-mm-source">
        {t('word.mailings.data-source-name', { name: data?.name || state?.source?.name || '' })}{data?.sheet ? ` \u2014 ${data.sheet}` : ''}{dupes.size ? ` · ${tn(dupes.size, 'word.mailings.count-possible-duplicate')}` : ''}
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
      title={t('word.mailings.insert-address-block')}
      width={760}
      onClose={onClose}
      className="wd-mm-wide"
      actions={
        <>
          <Button icon="link" label={t('word.mailings.match-fields-2')} onClick={onMatch} />
          <span style={{ flex: 1 }} />
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.ok')} disabled={!withName && !company && !postal} onClick={() => onInsert(spec)} />
        </>
      }
    >
      <div className="wd-mm-cols">
        <div className="wd-mm-col">
          <div className="wd-mm-section">{t('word.mailings.specify-address')}</div>
          <label className="wd-mm-check"><input type="checkbox" checked={withName} onChange={(e) => setWithName(e.target.checked)} /> {t('word.mailings.insert-recipients-name')}</label>
          <div className="wd-mm-pick wd-mm-formats" role="listbox" aria-label={t('word.mailings.name-format')}>
            {ADDRESS_NAME_FORMATS.map((f) => (
              <button key={f.id} type="button" role="option" disabled={!withName} aria-selected={f.id === name} className={f.id === name ? 'on' : ''} onClick={() => setName(f.id)}>{f.label}</button>
            ))}
          </div>
          <label className="wd-mm-check"><input type="checkbox" checked={company} onChange={(e) => setCompany(e.target.checked)} /> {t('word.mailings.insert-company-name')}</label>
          <label className="wd-mm-check"><input type="checkbox" checked={postal} onChange={(e) => setPostal(e.target.checked)} /> {t('word.mailings.insert-postal')}</label>
          <div className="wd-mm-radios">
            <label><input type="radio" name="wd-mm-country" disabled={!postal} checked={country === 'never'} onChange={() => setCountry('never')} /> {t('word.mailings.never-include-country')}</label>
            <label><input type="radio" name="wd-mm-country" disabled={!postal} checked={country === 'always'} onChange={() => setCountry('always')} /> {t('word.mailings.always-include-country')}</label>
            <label><input type="radio" name="wd-mm-country" disabled={!postal} checked={country === 'unlessEqual'} onChange={() => setCountry('unlessEqual')} /> {t('word.mailings.only-include-country')}</label>
            <Input value={except} disabled={!postal || country !== 'unlessEqual'} onChange={(e) => setExcept(e.target.value)} style={{ marginLeft: 22 }} aria-label={t('word.mailings.leave-off-country')} />
          </div>
        </div>
        <div className="wd-mm-col">
          <div className="wd-mm-section">{t('common.preview')}</div>
          <p className="rw-hint" style={{ margin: '0 0 6px' }}>{t('word.mailings.preview-recipient-list')}</p>
          <RecordStepper at={p.at} total={p.total} onGo={p.setAt} />
          <div className="wd-mm-preview wd-mm-address-preview">
            {lines.length ? lines.map((l, i) => <div key={i}>{l}</div>) : <span className="rw-hint">{t('word.mailings.nothing-show-record')}</span>}
          </div>
          <p className="rw-hint">{t('word.mailings.items-address-block')}</p>
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
      title={t('word.mailings.insert-greeting-line')}
      width={560}
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.ok')} onClick={() => onInsert(spec)} /></>}
    >
      <div className="wd-mm-section">{t('word.mailings.greeting-line')}</div>
      <div className="wd-mm-row3">
        <Select value={salutation} onChange={(e) => setSalutation(e.target.value)} aria-label={t('word.mailings.salutation')}>
          {GREETING_SALUTATIONS.map((s) => <option key={s} value={s}>{s || t('common.none-2')}</option>)}
        </Select>
        <Select value={name} onChange={(e) => setName(e.target.value)} aria-label={t('word.mailings.name-format')}>
          {GREETING_NAME_FORMATS.map((f) => <option key={f.id} value={f.id}>{f.label}</option>)}
        </Select>
        <Select value={punctuation} onChange={(e) => setPunctuation(e.target.value)} aria-label={t('word.mailings.punctuation')}>
          {GREETING_PUNCTUATION.map((s) => <option key={s} value={s}>{s || t('common.none-2')}</option>)}
        </Select>
      </div>
      <div className="wd-mm-section">{t('word.mailings.greeting-line-invalid')}</div>
      <Select value={fallback} onChange={(e) => setFallback(e.target.value)} aria-label={t('word.mailings.fallback')} style={{ width: '100%' }}>
        {GREETING_FALLBACKS.map((s) => <option key={s} value={s}>{s || t('common.none-2')}</option>)}
      </Select>
      <div className="wd-mm-section">{t('common.preview')}</div>
      <RecordStepper at={p.at} total={p.total} onGo={p.setAt} />
      <div className="wd-mm-preview">{line || <span className="rw-hint">{t('word.mailings.nothing')}</span>}</div>
    </Dialog>
  );
}

/* ── Rules: If…Then…Else, Skip Record If, Next Record If ─────────────────── */

const RULE_TITLES = { if: t('word.mailings.insert-word-field'), skipif: t('word.mailings.insert-word-field-skip'), nextif: t('word.mailings.insert-word-field-next') };

/**
 * Rules → Ask, Fill-in and Set Bookmark, as Word's own dialogs: the
 * bookmark, the prompt, the default answer, Ask once; or the bookmark and
 * its value. An Ask or a Set can show its bookmark straight after it.
 */
function PromptRuleDialog({ kind, onClose, onInsert }) {
  const [name, setName] = useState('');
  const [prompt, setPrompt] = useState('');
  const [def, setDef] = useState('');
  const [once, setOnce] = useState(false);
  const [show, setShow] = useState(true);
  const needsName = kind !== 'fillin';
  const ok = (!needsName || /^[A-Za-z]/.test(name.trim())) && (kind === 'set' || prompt.trim());
  const title = { ask: t('word.mailings.insert-word-field-ask'), fillin: t('word.mailings.insert-word-field-fill'), set: t('word.mailings.insert-word-field-set') }[kind];
  return (
    <Dialog
      title={title}
      width={480}
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary className="wd-mm-prompt-ok" label={t('common.ok')} disabled={!ok} onClick={() => onInsert({ name: name.trim(), prompt, def, value: def, once, show })} /></>}
    >
      {needsName ? (
        <Field label={t('word.mailings.bookmark')} hint={t('word.mailings.letter-first-letters')}>
          <Input className="wd-mm-prompt-name" value={name} autoFocus onChange={(e) => setName(e.target.value)} />
        </Field>
      ) : null}
      {kind !== 'set' ? (
        <Field label={t('word.mailings.prompt')}>
          <Input className="wd-mm-prompt-text" value={prompt} autoFocus={!needsName} onChange={(e) => setPrompt(e.target.value)} />
        </Field>
      ) : null}
      <Field label={kind === 'set' ? t('word.mailings.value') : kind === 'ask' ? t('word.mailings.default-bookmark') : t('word.mailings.default-fill-text')}>
        <Input className="wd-mm-prompt-default" value={def} onChange={(e) => setDef(e.target.value)} />
      </Field>
      {kind !== 'set' ? <label className="wd-mm-check"><input type="checkbox" checked={once} onChange={(e) => setOnce(e.target.checked)} /> {t('word.mailings.ask-once-one')}</label> : null}
      {kind !== 'fillin' ? <label className="wd-mm-check"><input type="checkbox" className="wd-mm-prompt-show" checked={show} onChange={(e) => setShow(e.target.checked)} /> {t('word.mailings.show-bookmarks-value')}</label> : null}
    </Dialog>
  );
}

/**
 * Before a merge with Ask or Fill-in fields: each question answered — once,
 * for a question asked once, else for each record in turn, its default
 * filled in.
 */
function MergeAnswersDialog({ prompts, records, onClose, onMerge }) {
  const [answers, setAnswers] = useState(() => Object.fromEntries(prompts.map((p) => [p.key, p.once ? p.def : Object.fromEntries(records.map((r) => [String(r.number), p.def]))])));
  const set = (key, record, value) => setAnswers((a) => ({ ...a, [key]: record == null ? value : { ...a[key], [String(record)]: value } }));
  return (
    <Dialog
      title={t('word.mailings.answer-merges-questions')}
      width={620}
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary className="wd-mm-answers-ok" label={t('word.merge')} onClick={() => onMerge(answers)} /></>}
    >
      <div className="wd-mm-answers">
        {prompts.map((p) => (
          <div key={p.key} className="wd-mm-answer" data-key={p.key}>
            <div className="wd-mm-section">{p.prompt || p.name}{p.kind === 'ask' ? ` ${t('word.mailings.kept-in-name', { name: p.name })}` : ''}</div>
            {p.once ? (
              <Input className="wd-mm-answer-once" value={answers[p.key]} onChange={(e) => set(p.key, null, e.target.value)} />
            ) : records.map((r) => (
              <label key={r.number} className="wd-mm-answer-row">
                <span>{r.label}</span>
                <Input className="wd-mm-answer-record" data-record={r.number} value={answers[p.key][String(r.number)]} onChange={(e) => set(p.key, r.number, e.target.value)} />
              </label>
            ))}
          </div>
        ))}
      </div>
    </Dialog>
  );
}

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
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.ok')} disabled={!ok} onClick={() => onInsert({ field, comparison, value, then, otherwise })} /></>}
    >
      <div className="wd-mm-section">{kind === 'if' ? 'IF' : t('word.mailings.skip-move')}</div>
      <div className="wd-mm-row3">
        <Field label={t('word.mailings.field-name')}>
          <Select value={field} onChange={(e) => setField(e.target.value)}>{fields.map((f) => <option key={f} value={f}>{f}</option>)}</Select>
        </Field>
        <Field label={t('word.mailings.comparison')}>
          <Select value={comparison} onChange={(e) => setComparison(e.target.value)}>{COMPARISONS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</Select>
        </Field>
        <Field label={t('word.mailings.compare-to')}>
          <Input value={value} disabled={!needsValue} onChange={(e) => setValue(e.target.value)} />
        </Field>
      </div>
      {kind === 'if' ? (
        <>
          <Field label={t('word.mailings.insert-this-text')}>
            <textarea className="rw-input wd-mm-text" rows={3} value={then} onChange={(e) => setThen(e.target.value)} />
          </Field>
          <Field label={t('word.mailings.otherwise-insert')}>
            <textarea className="rw-input wd-mm-text" rows={3} value={otherwise} onChange={(e) => setOtherwise(e.target.value)} />
          </Field>
        </>
      ) : (
        <p className="rw-hint">{kind === 'skipif' ? t('word.mailings.record-holds-left') : t('word.mailings.holds-fields-after')}</p>
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
      title={t('word.mailings.match-fields')}
      width={520}
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.ok')} onClick={() => onApply(map)} /></>}
    >
      <p className="wd-mm-lead">{t('word.mailings.each-address-field')}</p>
      <div className="wd-mm-match">
        {ADDRESS_FIELDS.map((f) => (
          <React.Fragment key={f.token}>
            <label htmlFor={`wd-mm-m-${f.token}`}>{ADDRESS_FIELD_NAMES[f.token] || f.label}</label>
            <Select id={`wd-mm-m-${f.token}`} value={map[f.token]} onChange={(e) => setMap((m) => ({ ...m, [f.token]: e.target.value }))}>
              <option value="">{t('word.mailings.not-matched')}</option>
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
      title={t('word.mailings.find-entry')}
      width={440}
      onClose={onClose}
      actions={<><Button label={t('common.close')} onClick={onClose} /><Button primary label={t('word.mailings.find-next')} disabled={!text.trim()} onClick={() => onFind(text, where === 'field' ? field : null)} /></>}
    >
      <Field label={t('word.mailings.find')}>
        <Input value={text} autoFocus onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && text.trim()) onFind(text, where === 'field' ? field : null); }} />
      </Field>
      <div className="wd-mm-radios" style={{ marginTop: 10 }}>
        <label><input type="radio" name="wd-mm-where" checked={where === 'all'} onChange={() => setWhere('all')} /> {t('word.mailings.all-fields')}</label>
        <label className="wd-mm-inline"><input type="radio" name="wd-mm-where" checked={where === 'field'} onChange={() => setWhere('field')} /> {t('word.mailings.this-field')}
          <Select value={field} disabled={where !== 'field'} onChange={(e) => setField(e.target.value)}>{fields.map((f) => <option key={f} value={f}>{f}</option>)}</Select>
        </label>
      </div>
    </Dialog>
  );
}

/* ── Finish & Merge ──────────────────────────────────────────────────────── */

const FINISH_TITLES = { document: t('word.mailings.merge-new-document'), print: t('word.mailings.merge-to-printer'), email: t('word.mailings.merge-e-mail') };

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

  // Rules → Ask and Fill-in: the merge's questions, answered before it runs.
  const [asking, setAsking] = useState(null);
  const run = async (answers = null) => {
    if (!answers) {
      const got = await call('prompts', { range: spec }).catch(() => null);
      if (got?.prompts?.length) { setAsking(got); return; }
      answers = {};
    }
    setAsking(null);
    setBusy(true);
    try {
      if (to === 'document') {
        const made = await call('finish', { range: spec, answers });
        await shell.win.create({ app: 'word', query: { session: made.id } });
        toast(tn(made.copies, 'word.mailings.name-count-copy', { name: made.name, type: TYPE_NAMES_LOWER[made.type] || made.typeLabel.toLowerCase() }), { tone: 'good', ms: 5000 });
        onClose();
      } else if (to === 'print') {
        const made = await call('finish', { range: spec, answers });
        setPrinting(made);
      } else {
        if (!account) { toast(t('word.mailings.add-mail-account'), { tone: 'bad' }); setBusy(false); return; }
        if (!toField) { toast(t('word.mailings.choose-column-holds'), { tone: 'bad' }); setBusy(false); return; }
        await apply({ op: 'setMergeEmail', toField, subject });
        const messages = await call('messages', { range: spec, toField, subject, format, answers });
        setProgress({ done: 0, total: messages.length, sent: 0, failed: [] });
        const out = await sendMergedMessages(
          messages,
          (m) => shell.mail.send({ accountId: account, draft: { to: m.to, subject: m.subject, text: m.text, ...(m.html ? { html: m.html } : {}) } }),
          (done, total) => setProgress((p) => ({ ...p, done, total })),
        );
        setProgress((p) => ({ ...p, sent: out.sent, failed: out.failed, finished: true }));
        toast(out.failed.length ? tn(messages.length, 'word.mailings.sent-sent-count-message', { sent: out.sent, failed: out.failed.length }) : tn(messages.length, 'word.mailings.sent-sent-count', { sent: out.sent }), { tone: out.failed.length ? 'bad' : 'good', ms: 6000 });
      }
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    }
    setBusy(false);
  };

  if (asking) return <MergeAnswersDialog prompts={asking.prompts} records={asking.records} onClose={() => setAsking(null)} onMerge={(answers) => run(answers)} />;

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
            const target = await shell.dialog.save({ title: t('common.save-as-pdf'), defaultPath: printing.name.replace(/\.[^.]+$/, '') + '.pdf', filters: [{ name: 'PDF', extensions: ['pdf'] }] });
            if (target) {
              await shell.print.pdf({ id: printing.id, path: target, options });
              toast(t('common.saved-name', { name: target.split(/[\\/]/).pop() }), { tone: 'good' });
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
      actions={finished ? <Button primary label={t('common.close')} onClick={onClose} /> : (
        <><Button label={t('common.cancel')} onClick={onClose} disabled={busy} /><Button primary label={to === 'email' ? t('common.send') : t('common.ok')} disabled={busy || (to === 'email' && accounts !== null && !accounts.length)} onClick={() => run()} /></>
      )}
    >
      {to === 'email' ? (
        <>
          <div className="wd-mm-section">{t('word.mailings.message-options')}</div>
          <div className="wd-mm-match">
            <label>{t('word.mailings.to')}</label>
            <Select className="rw-select wd-mm-to" value={toField} onChange={(e) => setToField(e.target.value)}>
              <option value="">{t('word.mailings.choose-a-column')}</option>
              {(state?.source?.fields || []).map((f) => <option key={f} value={f}>{f}</option>)}
            </Select>
            <label>{t('word.mailings.subject-line')}</label>
            <Input className="rw-input wd-mm-subject" value={subject} onChange={(e) => setSubject(e.target.value)} />
            <label>{t('word.mailings.mail-format')}</label>
            <Select value={format} onChange={(e) => setFormat(e.target.value)}>
              <option value="html">HTML</option>{/* words-ok: a format's name */}
              <option value="text">{t('common.plain-text')}</option>
            </Select>
            <label>{t('word.mailings.send-from')}</label>
            <Select className="rw-select wd-mm-account" value={account} onChange={(e) => setAccount(e.target.value)}>
              {accounts === null ? <option>{t('word.mailings.reading-accounts')}</option> : accounts.length ? accounts.map((a) => <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>) : <option value="">{t('word.mailings.no-mail-account')}</option>}
            </Select>
          </div>
          <p className="rw-hint">{t('word.mailings.each-message-merged')}</p>
        </>
      ) : null}
      <div className="wd-mm-section">{to === 'email' ? t('word.mailings.send-records') : t('word.mailings.merge-records')}</div>
      <div className="wd-mm-radios">
        <label><input type="radio" name="wd-mm-range" checked={range === 'all'} onChange={() => setRange('all')} /> {t('common.all')}</label>
        <label><input type="radio" name="wd-mm-range" checked={range === 'current'} onChange={() => setRange('current')} /> {t('word.mailings.current-record')}</label>
        <label className="wd-mm-inline"><input type="radio" name="wd-mm-range" checked={range === 'range'} onChange={() => setRange('range')} /> {t('word.mailings.from')}
          <Input className="rw-input wd-mm-num" value={from} disabled={range !== 'range'} onChange={(e) => setFrom(e.target.value.replace(/[^\d]/g, ''))} aria-label={t('word.mailings.from-record')} />
          {t('word.mailings.to')}
          <Input className="rw-input wd-mm-num" value={until} disabled={range !== 'range'} onChange={(e) => setUntil(e.target.value.replace(/[^\d]/g, ''))} aria-label={t('word.mailings.to-record')} />
        </label>
      </div>
      {progress ? (
        <div className="wd-mm-progress" aria-live="polite">
          <Progress value={progress.done} max={progress.total || 1} />
          <span className="wd-mm-progress-text">{finished ? t('word.mailings.sent-sent-total', { sent: progress.sent, total: progress.total }) : t('word.mailings.sending-done-total', { done: progress.done, total: progress.total })}</span>
          {finished && progress.failed.length ? (
            <ul className="wd-mm-failed">{progress.failed.map((f) => <li key={f.record}>{f.to ? t('word.mailings.record-record-to-error', { record: f.record, to: f.to, error: f.error }) : t('word.mailings.record-record-error', { record: f.record, error: f.error })}</li>)}</ul>
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
.wd-mm-sortable:hover { color: var(--accent-text); }
.wd-mm-newlist td { padding: 0; border-right: 1px solid var(--line-soft); }
.wd-mm-newlist th { border-right: 1px solid var(--line-soft); }
.wd-mm-newlist tr.on td { background: var(--selected); }
.wd-mm-rowno { width: 22px; text-align: center; color: var(--accent-text); font-size: 11px; }
.wd-mm-cell { border: 0; background: transparent; font: inherit; font-size: 12.5px; padding: 6px 8px; width: 132px; color: var(--ink); }
.wd-mm-cell:focus { outline: 2px solid var(--accent); outline-offset: -2px; background: var(--surface); }
.wd-mm-source { margin: 8px 0 0; }
.wd-mm-foot-note { margin-right: auto; align-self: center; font-size: 12px; color: var(--ink-2); }
.wd-mm-progress { margin-top: 14px; display: flex; flex-direction: column; gap: 6px; }
.wd-mm-progress-text { font-size: 12.5px; color: var(--ink-2); }
.wd-mm-failed { margin: 4px 0 0; padding-left: 18px; font-size: 12px; color: var(--bad); }
.rw-dialog.wd-mm-wide { max-width: min(94vw, 920px); }
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
