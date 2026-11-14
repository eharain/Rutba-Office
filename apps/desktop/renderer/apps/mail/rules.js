// Rules, and the thing that makes them usable: seeing what one would do first.
//
// Every mail client has rules and almost every person is wary of them, for a
// good reason — a rule you got slightly wrong moves a thousand messages before
// you can look at it. So the editor here answers "what would this do to the
// folder I am looking at" before it offers to do anything, and the count is
// live as the rule is typed.

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { Button, Dialog, Field, Input, Select, Icon, Chip, Empty, Spinner, formatWhen, t, tn } from '@rutba/office-ui';
import { avatarFor, displayName } from './parts.js';

const FIELDS = [
  ['from', t('mail.sender')],
  ['fromAddress', t('mail.rules.senders-address')],
  ['to', t('mail.rules.recipient')],
  ['subject', t('common.subject')],
  ['body', t('mail.rules.message-text')],
  ['list', t('common.mailing-list')],
];

const OPERATORS = [
  ['contains', t('common.contains')],
  ['notContains', t('common.does-not-contain')],
  ['is', t('mail.rules.is-exactly')],
  ['isNot', t('mail.rules.is-not')],
  ['startsWith', t('mail.rules.starts-with')],
  ['endsWith', t('common.ends-with')],
  ['matches', t('mail.rules.matches-the-pattern')],
];

const FLAGS = [
  ['hasAttachment', t('mail.rules.has-an-attachment')],
  ['noAttachment', t('mail.rules.has-no-attachment')],
  ['isUnread', t('mail.rules.is-unread')],
  ['isBulk', t('mail.rules.is-bulk-mail')],
  ['isNotBulk', t('mail.rules.not-bulk-mail')],
];

const ACTIONS = [
  ['move', t('mail.rules.move-to-folder')],
  ['copy', t('mail.rules.copy-to-folder')],
  ['star', t('mail.rules.star-it')],
  ['pin', t('mail.rules.pin-top')],
  ['markRead', t('mail.rules.mark-as-read')],
  ['markUnread', t('mail.rules.mark-as-unread')],
  ['junk', t('mail.move-to-junk')],
  ['trash', t('mail.rules.move-to-trash')],
  ['delete', t('mail.rules.delete-permanently')],
];

const blankRule = () => ({
  name: '',
  enabled: true,
  all: true,
  stopOnMatch: true,
  conditions: [{ field: 'from', op: 'contains', value: '' }],
  actions: [{ type: 'move', value: '' }],
});

/* ── the list of rules ───────────────────────────────────────────────────── */

export function RulesDialog({ shell, accountId, folder, folders, onClose, toast }) {
  const [rules, setRules] = useState(null);
  const [editing, setEditing] = useState(null);

  const refresh = useCallback(() => {
    shell.mail.rules().then(setRules).catch(() => setRules([]));
  }, [shell]);

  useEffect(() => refresh(), [refresh]);

  if (editing) {
    return (
      <RuleEditor
        shell={shell}
        rule={editing}
        accountId={accountId}
        folder={folder}
        folders={folders}
        onCancel={() => setEditing(null)}
        onSave={async (rule) => {
          await shell.mail.saveRule({ rule });
          setEditing(null);
          refresh();
          toast(t('mail.rules.rule-saved'), { tone: 'good' });
        }}
      />
    );
  }

  return (
    <Dialog
      title={t('common.rules')}
      width={620}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.close')} onClick={onClose} />
          <Button
            label={t('mail.rules.run-folder')}
            disabled={!rules?.length || !folder}
            onClick={async () => {
              const r = await shell.mail.runRules({ accountId, folder });
              toast(
                r.matched
                  ? t('mail.rules.matched-matched-moved', { matched: r.matched, moved: r.moved, starred: r.starred, read: r.read, deleted: r.deleted })
                  : t('mail.rules.nothing-folder-matched'),
                { tone: r.matched ? 'good' : 'plain', ms: 7000 }
              );
              onClose();
            }}
          />
          <Button primary icon="plus" label={t('mail.rules.new-rule')} onClick={() => setEditing(blankRule())} />
        </>
      }
    >
      {!rules ? (
        <div style={{ padding: 30, display: 'grid', placeItems: 'center' }}><Spinner /></div>
      ) : rules.length ? (
        <div className="ml-import-folders" style={{ maxHeight: 330 }}>
          {rules.map((rule) => (
            <div key={rule.id} className="ml-import-folder" style={{ gap: 10 }}>
              <input
                type="checkbox"
                className="ml-check"
                checked={rule.enabled !== false}
                title={rule.enabled !== false ? t('common.on') : t('common.off')}
                onChange={async (e) => {
                  await shell.mail.saveRule({ rule: { ...rule, enabled: e.target.checked } });
                  refresh();
                }}
              />
              <button type="button" className="ml-linkchip name" style={{ flex: 1, textAlign: 'left' }} onClick={() => setEditing(rule)}>
                <strong>{rule.name || t('mail.rules.untitled-rule')}</strong>
                <span style={{ opacity: 0.7 }}> — {describe(rule)}</span>
              </button>
              <Button
                icon="trash"
                title={t('mail.rules.delete-this-rule')}
                onClick={async () => {
                  await shell.mail.deleteRule({ id: rule.id });
                  refresh();
                }}
              />
            </div>
          ))}
        </div>
      ) : (
        <Empty icon="filter" title={t('mail.rules.no-rules-yet')}>
          {t('mail.rules.rule-files-mail')}
        </Empty>
      )}
      <p className="rw-hint">{t('mail.rules.rules-run-mail')}</p>
    </Dialog>
  );
}

/** A rule in one line, for the list. */
function describe(rule) {
  const first = rule.conditions?.[0];
  const condition = !first
    ? t('mail.rules.anything')
    : first.flag
      ? FLAGS.find(([id]) => id === first.flag)?.[1] || first.flag
      : t('mail.rules.field-operator-value', {
          field: FIELDS.find(([id]) => id === first.field)?.[1] || first.field,
          operator: OPERATORS.find(([id]) => id === first.op)?.[1] || first.op,
          value: first.value,
        });
  const more = (rule.conditions?.length || 0) - 1;
  const actions = (rule.actions || []).map((a) => (a.value ? `${a.type} → ${a.value}` : a.type)).join(', ');
  return `${condition}${more > 0 ? ` ${tn(more, 'mail.rules.and-count-more')}` : ''} → ${actions || t('common.nothing')}`;
}

/* ── one rule ────────────────────────────────────────────────────────────── */

function RuleEditor({ shell, rule: initial, accountId, folder, folders, onCancel, onSave }) {
  const [rule, setRule] = useState(initial);
  const [preview, setPreview] = useState(null);
  const [checking, setChecking] = useState(false);

  const usable = useMemo(
    () => (rule.conditions || []).some((c) => c.flag || String(c.value || '').trim()) && (rule.actions || []).length > 0,
    [rule]
  );

  // What this would do, recomputed as it is typed. Debounced, because it walks
  // the folder each time and nobody needs an answer per keystroke.
  useEffect(() => {
    if (!usable || !folder) {
      setPreview(null);
      return undefined;
    }
    setChecking(true);
    const timer = setTimeout(() => {
      shell.mail
        .testRules({ accountId, folder, rules: [{ ...rule, enabled: true }] })
        .then(setPreview)
        .catch(() => setPreview(null))
        .finally(() => setChecking(false));
    }, 350);
    return () => {
      clearTimeout(timer);
      setChecking(false);
    };
  }, [rule, usable, accountId, folder, shell]);

  const setCondition = (i, patch) =>
    setRule((r) => ({ ...r, conditions: r.conditions.map((c, n) => (n === i ? { ...c, ...patch } : c)) }));
  const setAction = (i, patch) =>
    setRule((r) => ({ ...r, actions: r.actions.map((a, n) => (n === i ? { ...a, ...patch } : a)) }));

  return (
    <Dialog
      title={initial.id ? t('mail.rules.edit-rule') : t('mail.rules.new-rule')}
      width={680}
      onClose={onCancel}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onCancel} />
          <Button primary label={t('common.save')} disabled={!usable} onClick={() => onSave(rule)} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('common.name')}>
          <Input value={rule.name} onChange={(e) => setRule((r) => ({ ...r, name: e.target.value }))} placeholder={t('mail.rules.newsletters-into')} autoFocus />
        </Field>

        <div className="ml-listbar" style={{ padding: 0, border: 0, minHeight: 0 }}>
          <span>{t('mail.rules.when')}</span>
          <Select value={rule.all === false ? 'any' : 'all'} onChange={(e) => setRule((r) => ({ ...r, all: e.target.value === 'all' }))} style={{ width: 130 }}>
            <option value="all">{t('mail.rules.all-of-these')}</option>
            <option value="any">{t('mail.rules.any-of-these')}</option>
          </Select>
          <span>{t('mail.rules.are-true')}</span>
        </div>

        {(rule.conditions || []).map((c, i) => (
          <div key={i} className="ml-rule-row">
            <Select
              value={c.flag ? `flag:${c.flag}` : c.field}
              onChange={(e) => {
                const v = e.target.value;
                if (v.startsWith('flag:')) setCondition(i, { flag: v.slice(5), field: undefined, op: undefined, value: undefined });
                else setCondition(i, { field: v, flag: undefined, op: c.op || 'contains', value: c.value || '' });
              }}
              style={{ width: 150 }}
            >
              {FIELDS.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
              <optgroup label={t('mail.rules.the-message')}>
                {FLAGS.map(([id, label]) => <option key={id} value={`flag:${id}`}>{label}</option>)}
              </optgroup>
            </Select>

            {c.flag ? (
              <span className="rw-hint" style={{ flex: 1 }}>{t('mail.rules.no-value-needed')}</span>
            ) : (
              <>
                <Select value={c.op} onChange={(e) => setCondition(i, { op: e.target.value })} style={{ width: 158 }}>
                  {OPERATORS.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
                </Select>
                <span style={{ flex: 1 }}>
                  <Input value={c.value || ''} onChange={(e) => setCondition(i, { value: e.target.value })} placeholder="…" />
                </span>
              </>
            )}

            <Button
              icon="minus"
              title={t('common.remove')}
              disabled={rule.conditions.length < 2}
              onClick={() => setRule((r) => ({ ...r, conditions: r.conditions.filter((_, n) => n !== i) }))}
            />
          </div>
        ))}
        <Button
          icon="plus"
          label={t('mail.rules.add-a')}
          onClick={() => setRule((r) => ({ ...r, conditions: [...r.conditions, { field: 'subject', op: 'contains', value: '' }] }))}
        />

        <div className="ml-listbar" style={{ padding: 0, border: 0, minHeight: 0 }}>
          <span>{t('mail.rules.then')}</span>
        </div>

        {(rule.actions || []).map((a, i) => (
          <div key={i} className="ml-rule-row">
            <Select value={a.type} onChange={(e) => setAction(i, { type: e.target.value })} style={{ width: 170 }}>
              {ACTIONS.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
            </Select>
            {a.type === 'move' || a.type === 'copy' ? (
              <span style={{ flex: 1 }}>
                <Select value={a.value || ''} onChange={(e) => setAction(i, { value: e.target.value })}>
                  <option value="">{t('mail.rules.choose-a')}</option>
                  {(folders || []).map((f) => <option key={f.path} value={f.path}>{f.name || f.path}</option>)}
                </Select>
              </span>
            ) : (
              <span style={{ flex: 1 }} />
            )}
            <Button
              icon="minus"
              title={t('common.remove')}
              disabled={rule.actions.length < 2}
              onClick={() => setRule((r) => ({ ...r, actions: r.actions.filter((_, n) => n !== i) }))}
            />
          </div>
        ))}
        <Button icon="plus" label={t('mail.rules.add-an')} onClick={() => setRule((r) => ({ ...r, actions: [...r.actions, { type: 'star' }] }))} />

        <label className="about-auto">
          <input type="checkbox" checked={rule.stopOnMatch !== false} onChange={(e) => setRule((r) => ({ ...r, stopOnMatch: e.target.checked }))} />
          <span>{t('mail.rules.stop-checking-other')}</span>
        </label>

        {/*
          The part that makes a rule safe to write: what it would do, now, to
          the folder in front of you — computed without touching a message.
        */}
        <div className="ml-note" style={{ alignItems: 'flex-start', flexDirection: 'column', gap: 8 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, width: '100%' }}>
            {checking ? <Spinner /> : <Icon name={preview?.matched ? 'check' : 'info'} size={14} />}
            <span>
              {!usable
                ? t('mail.rules.fill-condition-see')
                : checking
                  ? t('mail.rules.checking-this')
                  : preview
                    ? tn(preview.of, 'mail.rules.matched-count-message', { matched: preview.matched.toLocaleString() })
                    : t('mail.rules.nothing-check-against')}
            </span>
          </div>

          {preview?.sample?.length ? (
            <div style={{ width: '100%', maxHeight: 150, overflow: 'auto', display: 'flex', flexDirection: 'column', gap: 4 }}>
              {preview.sample.slice(0, 8).map((s) => {
                const avatar = avatarFor(s.from);
                return (
                  <div key={s.id} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 11.5 }}>
                    <span className="ml-avatar-sm" style={{ background: avatar.colour, width: 18, height: 18, fontSize: 9 }}>{avatar.initial}</span>
                    <span style={{ flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {displayName(s.from)} {'\u2014'} {s.subject || t('mail.no-subject')}
                    </span>
                    <Chip>{s.actions.join(', ')}</Chip>
                    <span style={{ opacity: 0.6 }}>{formatWhen(s.date)}</span>
                  </div>
                );
              })}
              {preview.matched > 8 ? <span className="rw-hint">{tn(preview.matched - 8, 'mail.rules.and-count-more')}</span> : null}
            </div>
          ) : null}
        </div>
      </div>
    </Dialog>
  );
}
