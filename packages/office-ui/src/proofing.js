// Review's panes and dialogs, shared by Rutba Word, Worksheets and
// Presentation: the Accessibility pane (Office's "Inspection Results"), the
// Editor pane the spelling pass runs in, the status-bar verdict, and the Alt
// Text, Dictionary and one-line prompt dialogs.
//
// Presentational only. What a finding is, where it points and what a fix
// writes are the app's (renderer/review.js and each app's own review.js);
// these draw what they are handed and say what was pressed. The look is the
// suite's own — the same tokens as every other pane — and deliberately
// quiet: a list, a card for the one in hand, and words rather than colours
// to say how serious each finding is.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from './icons.js';
import { Button, Dialog } from './index.js';

const TIERS = [
  { id: 'error', label: 'Errors' },
  { id: 'warning', label: 'Warnings' },
  { id: 'tip', label: 'Tips' },
];

/** Findings grouped the way the pane lists them: tier → rule → items. */
export function groupFindings(issues = [], rules = {}) {
  return TIERS.map((tier) => {
    const inTier = issues.filter((i) => i.severity === tier.id);
    const groups = [];
    for (const issue of inTier) {
      let g = groups.find((x) => x.rule === issue.rule);
      if (!g) groups.push((g = { rule: issue.rule, title: rules[issue.rule]?.title || issue.rule, items: [] }));
      g.items.push(issue);
    }
    return { ...tier, count: inTier.length, groups };
  }).filter((t) => t.count);
}

/* ── the Accessibility pane ──────────────────────────────────────────── */

export function AccessibilityPane({ result, loading, selectedKey, onSelect, onFix, keepRunning, onKeepRunning, onRecheck, busyFix }) {
  const [closed, setClosed] = useState(() => new Set());
  const tiers = useMemo(() => groupFindings(result?.issues || [], result?.rules || {}), [result]);
  const selected = (result?.issues || []).find((i) => i.key === selectedKey) || null;
  const rule = selected ? result?.rules?.[selected.rule] : null;
  const toggle = (id) => setClosed((s) => {
    const n = new Set(s);
    if (n.has(id)) n.delete(id);
    else n.add(id);
    return n;
  });
  const listRef = useRef(null);
  useEffect(() => {
    listRef.current?.querySelector('.pf-item.on')?.scrollIntoView?.({ block: 'nearest' });
  }, [selectedKey]);

  return (
    <div className="pf-pane pf-a11y">
      <div className="pf-scroll" ref={listRef}>
        <div className="pf-heading">
          <span>Inspection Results</span>
          <button type="button" className="pf-link" onClick={onRecheck} disabled={loading} data-tip="Check again — read the document afresh">
            <Icon name="refresh" size={13} /> Check again
          </button>
        </div>
        {!result && loading ? <div className="pf-note">Checking…</div> : null}
        {result && !result.issues.length ? (
          <div className="pf-clean">
            <span className="pf-clean-mark"><Icon name="check" size={18} /></span>
            <b>No accessibility issues found</b>
            <p>People with disabilities should not have difficulty reading this document.</p>
          </div>
        ) : null}
        {tiers.map((t) => (
          <div key={t.id} className={`pf-tier tier-${t.id}`}>
            <button type="button" className="pf-tier-head" aria-expanded={!closed.has(t.id)} onClick={() => toggle(t.id)}>
              <Icon name={closed.has(t.id) ? 'chevronRight' : 'chevronDown'} size={13} />
              <span className="pf-tier-mark" aria-hidden="true" />
              <span className="pf-tier-label">{t.label}</span>
              <span className="pf-count">{t.count}</span>
            </button>
            {closed.has(t.id) ? null : t.groups.map((g) => {
              const gid = `${t.id}:${g.rule}`;
              const open = !closed.has(gid);
              return (
                <div key={gid} className="pf-rule" data-rule={g.rule}>
                  <button type="button" className="pf-rule-head" aria-expanded={open} onClick={() => toggle(gid)}>
                    <Icon name={open ? 'chevronDown' : 'chevronRight'} size={12} />
                    <span>{g.title}</span>
                    <span className="pf-rule-count">({g.items.length})</span>
                  </button>
                  {open ? g.items.map((item) => {
                    const on = item.key === selectedKey;
                    return (
                      <div key={item.key} className={`pf-item${on ? ' on' : ''}`} data-key={item.key}>
                        <button type="button" className="pf-item-main" onClick={() => onSelect?.(item)} title={item.label}>
                          <span className="pf-item-label">{item.label}</span>
                          {item.detail ? <span className="pf-item-detail">{item.detail}</span> : null}
                        </button>
                        {on && item.fixes?.length ? (
                          <div className="pf-actions">
                            <div className="pf-actions-title">Recommended actions</div>
                            {item.fixes.map((f) => (
                              <button key={f.kind + (f.label || '')} type="button" className="pf-action" data-fix={f.kind} disabled={Boolean(busyFix)} onClick={() => onFix?.(item, f)}>
                                {f.label}
                              </button>
                            ))}
                          </div>
                        ) : null}
                      </div>
                    );
                  }) : null}
                </div>
              );
            })}
          </div>
        ))}
      </div>
      {rule ? (
        <div className="pf-info">
          <div className="pf-info-title">Additional information</div>
          <div className="pf-info-h">Why fix:</div>
          <p>{rule.why}</p>
          <div className="pf-info-h">How to fix:</div>
          <p>{rule.how}</p>
        </div>
      ) : null}
      <label className="pf-keep">
        <input type="checkbox" checked={Boolean(keepRunning)} onChange={(e) => onKeepRunning?.(e.target.checked)} />
        <span>Keep accessibility checker running while I work</span>
      </label>
    </div>
  );
}

/** The status bar's verdict: a button that opens the pane. */
export function A11yStatus({ verdict, onClick }) {
  if (!verdict) return null;
  const good = verdict === 'good';
  return (
    <button type="button" className={`pf-status${good ? ' good' : ' investigate'}`} onClick={onClick} data-tip={good ? 'Accessibility: Good to go — no issues found' : 'Accessibility: Investigate — open the Accessibility pane'}>
      <Icon name={good ? 'check' : 'shield'} size={13} />
      <span>{good ? 'Accessibility: Good to go' : 'Accessibility: Investigate'}</span>
    </button>
  );
}

/* ── the Editor pane ─────────────────────────────────────────────────── */

/**
 * `state` is the pass as the app holds it:
 *   { phase: 'loading' | 'word' | 'done' | 'error', language, found, message, checking }
 * `found` is `{ word, context: { before, word, after }, suggestions, where }`.
 */
export function EditorPane({ state, options, onChange, onChangeAll, onIgnoreOnce, onIgnoreAll, onAdd, onOption, onDictionary, onClose, onRestart, restartLabel, asYouType }) {
  const found = state?.found || null;
  const [pick, setPick] = useState(0);
  const [typed, setTyped] = useState('');
  const [showOptions, setShowOptions] = useState(false);
  // A new word clears the box; its suggestions, when they arrive, fill it —
  // unless the person has typed a word of their own meanwhile.
  useEffect(() => {
    setPick(0);
    setTyped(found?.suggestions?.[0] || '');
  }, [found?.word, found?.key, found?.offset]);
  useEffect(() => {
    if (found?.suggestions?.length) setTyped((t) => t || found.suggestions[0]);
  }, [found?.suggestions]);
  const replacement = typed.trim();
  const busy = Boolean(state?.checking);

  return (
    <div className="pf-pane pf-editor">
      <div className="pf-scroll">
        <div className="pf-heading">
          <span>Spelling</span>
          {state?.language ? <span className="pf-lang" data-tip={`Checking in ${state.language}`}>{state.language}</span> : null}
        </div>

        {state?.phase === 'loading' ? (
          <div className="pf-note"><span className="rw-spinner pf-spin" /> {state.message || 'Reading the dictionary…'}</div>
        ) : null}

        {state?.phase === 'error' ? <div className="pf-note bad">{state.message}</div> : null}

        {state?.phase === 'word' && found ? (
          <>
            <div className="pf-context" data-where={found.whereLabel || ''}>
              {found.whereLabel ? <div className="pf-context-where">{found.whereLabel}</div> : null}
              <span>{found.context?.before}</span>
              <mark className="pf-miss">{found.context?.word || found.word}</mark>
              <span>{found.context?.after}</span>
            </div>
            <div className="pf-label">Not in dictionary</div>
            <div className="pf-suggest-title">Suggestions</div>
            {found.suggestions == null ? (
              <div className="pf-note small"><span className="rw-spinner pf-spin" /> Finding suggestions…</div>
            ) : found.suggestions.length ? (
              <div className="pf-suggest" role="listbox" aria-label="Suggestions">
                {found.suggestions.map((s, i) => (
                  <button
                    key={s}
                    type="button"
                    role="option"
                    aria-selected={i === pick && replacement === s}
                    className={`pf-suggestion${i === pick && replacement === s ? ' on' : ''}`}
                    onClick={() => { setPick(i); setTyped(s); }}
                    onDoubleClick={() => onChange?.(s)}
                    disabled={busy}
                  >
                    {s}
                  </button>
                ))}
              </div>
            ) : <div className="pf-note small">No suggestions — type the word as it should be.</div>}
            <label className="pf-changeto">
              <span>Change to</span>
              <input className="rw-input" value={typed} onChange={(e) => setTyped(e.target.value)} spellCheck={false} onKeyDown={(e) => { if (e.key === 'Enter' && replacement) onChange?.(replacement); }} />
            </label>
            <div className="pf-buttons">
              <Button primary label="Change" title="Change — this one" disabled={!replacement || busy} onClick={() => onChange?.(replacement)} />
              <Button label="Change All" title="Change All — every one in this document" disabled={!replacement || busy} onClick={() => onChangeAll?.(replacement)} />
            </div>
            <div className="pf-buttons">
              <Button label="Ignore Once" title="Ignore Once — leave this one and go on" disabled={busy} onClick={onIgnoreOnce} />
              <Button label="Ignore All" title="Ignore All — this word, everywhere in this document" disabled={busy} onClick={onIgnoreAll} />
            </div>
            <div className="pf-buttons">
              <Button icon="plus" label="Add to Dictionary" title="Add to Dictionary — your own list, shared by Word, Worksheets and Presentation" disabled={busy} onClick={onAdd} />
            </div>
          </>
        ) : null}

        {state?.phase === 'done' ? (
          <div className="pf-clean pf-done">
            <span className="pf-clean-mark"><Icon name="check" size={18} /></span>
            <b>Spelling check complete</b>
            <p>{state.message || 'You\'re good to go.'}</p>
            <div className="pf-buttons center">
              <Button primary label="OK" title="OK — close the pane" onClick={onClose} />
              {onRestart ? <Button label={restartLabel || 'Check again'} className="pf-more" onClick={onRestart} /> : null}
            </div>
          </div>
        ) : null}
      </div>

      <div className="pf-options">
        <button type="button" className="pf-link" aria-expanded={showOptions} onClick={() => setShowOptions((v) => !v)}>
          <Icon name={showOptions ? 'chevronDown' : 'chevronRight'} size={12} /> Options
        </button>
        {showOptions ? (
          <div className="pf-options-body">
            {asYouType ? (
              <label><input type="checkbox" checked={asYouType.on} onChange={(e) => asYouType.set(e.target.checked)} /> Check spelling as you type</label>
            ) : null}
            <label><input type="checkbox" checked={options?.ignoreUppercase !== false} onChange={(e) => onOption?.({ ignoreUppercase: e.target.checked })} /> Ignore words in UPPERCASE</label>
            <label><input type="checkbox" checked={options?.ignoreNumbers !== false} onChange={(e) => onOption?.({ ignoreNumbers: e.target.checked })} /> Ignore words that contain numbers</label>
            <label><input type="checkbox" checked={options?.ignoreAddresses !== false} onChange={(e) => onOption?.({ ignoreAddresses: e.target.checked })} /> Ignore Internet and file addresses</label>
            <button type="button" className="pf-link" onClick={onDictionary}><Icon name="list" size={12} /> Custom Dictionary…</button>
          </div>
        ) : null}
      </div>
    </div>
  );
}

/* ── dialogs ─────────────────────────────────────────────────────────── */

export function AltTextDialog({ name, initial = '', decorative: initialDecorative = false, onApply, onClose }) {
  const [text, setText] = useState(initial || '');
  const [decorative, setDecorative] = useState(Boolean(initialDecorative));
  return (
    <Dialog
      title="Alt Text"
      width={440}
      onClose={onClose}
      actions={(
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="OK" className="pf-alt-ok" disabled={!decorative && !text.trim()} onClick={() => onApply?.({ descr: decorative ? '' : text.trim(), decorative })} />
        </>
      )}
    >
      {name ? <div className="pf-alt-name">{name}</div> : null}
      <label className="pf-alt-label" htmlFor="pf-alt-text">How would you describe this object and its context to someone who is blind or has low vision?</label>
      <textarea
        id="pf-alt-text"
        className="rw-input pf-alt-text"
        rows={4}
        value={decorative ? '' : text}
        disabled={decorative}
        placeholder={decorative ? 'Decorative objects are skipped by screen readers.' : 'For example: a bar chart of orders by month, rising from 120 in January to 310 in June.'}
        onChange={(e) => setText(e.target.value)}
        autoFocus
      />
      <div className="pf-hint">A sentence or two is usually right.</div>
      <label className="pf-check">
        <input type="checkbox" className="pf-alt-decorative" checked={decorative} onChange={(e) => setDecorative(e.target.checked)} />
        <span>Mark as decorative <em>— for borders, dividers and pictures that add nothing to the words</em></span>
      </label>
    </Dialog>
  );
}

export function PromptDialog({ title, label, initial = '', placeholder = '', okLabel = 'OK', className = '', onApply, onClose }) {
  const [value, setValue] = useState(initial || '');
  const ok = () => { if (value.trim()) onApply?.(value.trim()); };
  return (
    <Dialog title={title} width={400} onClose={onClose} actions={(<><Button label="Cancel" onClick={onClose} /><Button primary label={okLabel} className={`pf-prompt-ok ${className}`} disabled={!value.trim()} onClick={ok} /></>)}>
      <label className="pf-alt-label">{label}</label>
      <input className={`rw-input pf-prompt ${className}`} value={value} placeholder={placeholder} autoFocus onFocus={(e) => e.target.select()} onChange={(e) => setValue(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') ok(); }} spellCheck={false} />
    </Dialog>
  );
}

export function DictionaryDialog({ words = [], onSave, onImport, onExport, onClose }) {
  const [list, setList] = useState(() => [...words]);
  const [draft, setDraft] = useState('');
  const [filter, setFilter] = useState('');
  useEffect(() => setList([...words]), [words]);
  const add = () => {
    const w = draft.trim();
    if (!w || /\s/.test(w)) return;
    setList((l) => [...new Set([...l, w])].sort((a, b) => a.localeCompare(b)));
    setDraft('');
  };
  const shown = list.filter((w) => !filter || w.toLowerCase().includes(filter.toLowerCase()));
  return (
    <Dialog
      title="Custom Dictionary"
      width={420}
      onClose={onClose}
      actions={(
        <>
          <Button icon="import" label="Import…" title="Import — add the words of a .dic word list" onClick={async () => { const more = await onImport?.(); if (more?.length) setList((l) => [...new Set([...l, ...more])].sort((a, b) => a.localeCompare(b))); }} />
          <Button icon="export" label="Export…" title="Export — save these words as a .dic word list" onClick={() => onExport?.(list)} />
          <span style={{ flex: 1 }} />
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="OK" className="pf-dict-ok" onClick={() => onSave?.(list)} />
        </>
      )}
    >
      <p className="pf-dict-lead">Words you have added are never marked as misspelt — in Word, Worksheets and Presentation alike.</p>
      <div className="pf-dict-add">
        <input className="rw-input pf-dict-word" value={draft} placeholder="Add a word" spellCheck={false} onChange={(e) => setDraft(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') add(); }} />
        <Button label="Add" title="Add — put the word in the dictionary" disabled={!draft.trim() || /\s/.test(draft.trim())} onClick={add} />
      </div>
      {list.length > 8 ? <input className="rw-input pf-dict-filter" value={filter} placeholder="Find a word" onChange={(e) => setFilter(e.target.value)} spellCheck={false} /> : null}
      <div className="pf-dict-list" role="list">
        {shown.length ? shown.map((w) => (
          <div key={w} className="pf-dict-row" role="listitem" data-word={w}>
            <span>{w}</span>
            <button type="button" className="pf-dict-remove" aria-label={`Remove ${w}`} data-tip={`Remove "${w}"`} onClick={() => setList((l) => l.filter((x) => x !== w))}>
              <Icon name="close" size={12} />
            </button>
          </div>
        )) : <div className="pf-note small">{list.length ? 'No word matches.' : 'No words yet. Add to Dictionary in the Editor pane puts them here.'}</div>}
      </div>
      <div className="pf-hint">{list.length} {list.length === 1 ? 'word' : 'words'}</div>
    </Dialog>
  );
}

/* ── the look ────────────────────────────────────────────────────────── */

const CSS = `
.pf-pane { display: flex; flex-direction: column; height: 100%; min-height: 0; font-size: 12.5px; color: var(--ink); background: var(--chrome); }
.pf-scroll { flex: 1; min-height: 0; overflow: auto; padding: 4px 0 10px; }
.pf-heading { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 10px 14px 8px; font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-3); font-weight: 600; }
.pf-lang { text-transform: none; letter-spacing: 0; font-weight: 500; color: var(--ink-3); font-size: 11.5px; }
.pf-link { border: 0; background: none; color: var(--accent); font: inherit; font-size: 12px; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; padding: 2px 4px; border-radius: var(--r-1); text-transform: none; letter-spacing: 0; font-weight: 500; }
.pf-link:hover:not(:disabled) { background: var(--hover); }
.pf-link:disabled { color: var(--ink-3); cursor: default; }
.pf-note { padding: 10px 14px; color: var(--ink-2); display: flex; align-items: center; gap: 8px; line-height: 1.45; }
.pf-note.small { font-size: 12px; color: var(--ink-3); padding: 6px 14px; }
.pf-note.bad { color: var(--bad); }
.pf-spin { width: 14px; height: 14px; }

.pf-clean { margin: 8px 14px; padding: 18px 16px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); text-align: center; display: flex; flex-direction: column; align-items: center; gap: 6px; }
.pf-clean b { font-size: 13px; }
.pf-clean p { margin: 0; color: var(--ink-2); line-height: 1.45; }
.pf-clean-mark { width: 32px; height: 32px; border-radius: 50%; display: grid; place-items: center; color: var(--good); background: color-mix(in srgb, var(--good) 12%, transparent); }

.pf-tier { margin: 2px 0 6px; }
.pf-tier-head { width: 100%; display: flex; align-items: center; gap: 7px; border: 0; background: none; font: inherit; font-weight: 600; color: var(--ink); padding: 7px 12px; cursor: pointer; }
.pf-tier-head:hover { background: var(--hover); }
.pf-tier-mark { width: 8px; height: 8px; border-radius: 50%; background: var(--ink-3); flex: none; }
.tier-error .pf-tier-mark { background: var(--bad); }
.tier-warning .pf-tier-mark { background: var(--warn); }
.tier-tip .pf-tier-mark { background: var(--info); }
.pf-tier-label { flex: 1; text-align: left; }
.pf-count { min-width: 20px; padding: 0 6px; height: 18px; border-radius: 9px; background: var(--sunken); color: var(--ink-2); font-size: 11px; display: inline-grid; place-items: center; font-variant-numeric: tabular-nums; }
.pf-rule-head { width: 100%; display: flex; align-items: center; gap: 6px; border: 0; background: none; font: inherit; color: var(--ink-2); padding: 5px 12px 5px 26px; cursor: pointer; text-align: left; }
.pf-rule-head:hover { background: var(--hover); color: var(--ink); }
.pf-rule-count { color: var(--ink-3); }
.pf-item { margin: 0 8px 0 40px; border-radius: var(--r-1); }
.pf-item-main { width: 100%; display: flex; flex-direction: column; align-items: flex-start; gap: 1px; border: 0; background: none; font: inherit; color: var(--ink); padding: 5px 8px; cursor: pointer; text-align: left; border-radius: var(--r-1); }
.pf-item-main:hover { background: var(--hover); }
.pf-item.on { background: var(--surface); box-shadow: inset 0 0 0 1px var(--accent-line); }
.pf-item.on .pf-item-main { color: var(--accent); }
.pf-item-label { max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.pf-item-detail { font-size: 11px; color: var(--ink-3); max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.pf-actions { padding: 2px 8px 9px; display: flex; flex-direction: column; align-items: flex-start; gap: 3px; }
.pf-actions-title { font-size: 10.5px; text-transform: uppercase; letter-spacing: .05em; color: var(--ink-3); margin: 2px 0 2px; }
.pf-action { border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; font-size: 12px; padding: 4px 10px; border-radius: var(--r-1); cursor: pointer; }
.pf-action:hover:not(:disabled) { border-color: var(--accent-line); background: var(--selected); color: var(--accent); }

.pf-info { border-top: 1px solid var(--line-soft); padding: 10px 14px 4px; background: var(--chrome); max-height: 42%; overflow: auto; }
.pf-info-title { font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-3); font-weight: 600; margin-bottom: 6px; }
.pf-info-h { font-weight: 600; font-size: 12px; margin-top: 4px; }
.pf-info p { margin: 2px 0 6px; color: var(--ink-2); line-height: 1.45; font-size: 12px; }
.pf-keep { display: flex; align-items: flex-start; gap: 8px; padding: 9px 14px 11px; border-top: 1px solid var(--line-soft); color: var(--ink-2); font-size: 12px; line-height: 1.35; cursor: pointer; }
.pf-keep input { margin: 1px 0 0; accent-color: var(--accent); }

.pf-status { border: 0; background: none; font: inherit; font-size: 11.5px; color: var(--ink-2); display: inline-flex; align-items: center; gap: 5px; padding: 0 8px; height: 22px; border-radius: var(--r-1); cursor: pointer; }
.pf-status:hover { background: var(--hover); color: var(--ink); }
.pf-status.good svg { color: var(--good); }
.pf-status.investigate svg { color: var(--warn); }

.pf-context { margin: 2px 14px 10px; padding: 10px 12px; background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--r-2); line-height: 1.55; font-size: 13px; color: var(--ink-2); overflow-wrap: anywhere; }
.pf-context-where { font-size: 11px; font-weight: 600; color: var(--ink-3); margin-bottom: 3px; }
.pf-miss { background: none; color: var(--ink); font-weight: 600; text-decoration: underline wavy var(--bad); text-underline-offset: 3px; text-decoration-thickness: 1px; }
.pf-label { margin: 0 14px 8px; font-size: 11.5px; color: var(--bad); font-weight: 600; }
.pf-suggest-title { margin: 0 14px 4px; font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-3); font-weight: 600; }
.pf-suggest { margin: 0 14px 10px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); overflow: hidden; }
.pf-suggestion { display: block; width: 100%; border: 0; border-bottom: 1px solid var(--line-soft); background: none; font: inherit; font-size: 13px; text-align: left; padding: 6px 12px; cursor: pointer; color: var(--ink); }
.pf-suggestion:last-child { border-bottom: 0; }
.pf-suggestion:hover:not(:disabled) { background: var(--hover); }
.pf-suggestion.on { background: var(--selected); color: var(--accent); font-weight: 600; }
.pf-changeto { display: flex; flex-direction: column; gap: 4px; margin: 0 14px 10px; font-size: 11.5px; color: var(--ink-2); font-weight: 600; }
.pf-changeto .rw-input { font-weight: 400; }
.pf-buttons { display: flex; gap: 6px; margin: 0 14px 6px; flex-wrap: wrap; }
.pf-buttons.center { justify-content: center; margin-top: 8px; }
.pf-buttons .rw-btn { flex: 1; justify-content: center; border: 1px solid var(--line); background: var(--surface); min-height: 30px; }
.pf-buttons .rw-btn.primary { border-color: transparent; background: var(--accent); color: var(--ink-on-accent); }
.pf-buttons .rw-btn:hover:not(:disabled) { border-color: var(--line-strong); }
.pf-buttons .rw-btn.primary:hover:not(:disabled) { background: var(--accent-hover); }
.pf-buttons.center .rw-btn { flex: none; min-width: 84px; }
.pf-done { margin-top: 14px; }
.pf-options { border-top: 1px solid var(--line-soft); padding: 6px 10px 8px; }
.pf-options-body { display: flex; flex-direction: column; gap: 6px; padding: 6px 4px 2px; color: var(--ink-2); font-size: 12px; }
.pf-options-body label { display: flex; align-items: center; gap: 7px; cursor: pointer; }
.pf-options-body input { accent-color: var(--accent); margin: 0; }
.pf-options-body .pf-link { align-self: flex-start; margin-top: 2px; }

.pf-alt-name { font-size: 11.5px; color: var(--ink-3); margin: 0 0 8px; }
.pf-alt-label { display: block; font-size: 12.5px; color: var(--ink); margin: 0 0 6px; line-height: 1.4; }
.pf-alt-text { width: 100%; box-sizing: border-box; resize: vertical; min-height: 88px; font: inherit; line-height: 1.45; padding: 8px 10px; }
.pf-hint { font-size: 11.5px; color: var(--ink-3); margin: 5px 0 10px; }
.pf-check { display: flex; align-items: flex-start; gap: 8px; font-size: 12.5px; cursor: pointer; line-height: 1.4; }
.pf-check input { margin: 2px 0 0; accent-color: var(--accent); }
.pf-check em { font-style: normal; color: var(--ink-3); }
.pf-prompt { width: 100%; box-sizing: border-box; }

.pf-dict-lead { margin: 0 0 10px; color: var(--ink-2); font-size: 12.5px; line-height: 1.45; }
.pf-dict-add { display: flex; gap: 6px; margin-bottom: 8px; }
.pf-dict-add .rw-input { flex: 1; }
.pf-dict-add .rw-btn { border: 1px solid var(--line); background: var(--surface); }
.pf-dict-filter { width: 100%; box-sizing: border-box; margin-bottom: 6px; }
.pf-dict-list { height: 190px; overflow: auto; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); }
.pf-dict-row { display: flex; align-items: center; justify-content: space-between; padding: 4px 6px 4px 10px; border-bottom: 1px solid var(--line-soft); font-size: 12.5px; }
.pf-dict-row:hover { background: var(--hover); }
.pf-dict-remove { border: 0; background: none; color: var(--ink-3); cursor: pointer; display: grid; place-items: center; width: 22px; height: 22px; border-radius: var(--r-1); }
.pf-dict-remove:hover { background: var(--active); color: var(--bad); }
`;

let installed = false;
/** Injected once per window. */
export function installProofingStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-proofing-css';
  style.textContent = CSS;
  document.head.appendChild(style);
}
