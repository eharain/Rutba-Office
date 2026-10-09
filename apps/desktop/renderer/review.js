// Review → Check Accessibility and Review → Spelling, in any of the three
// document windows.
//
// One hook holds what the two panes need — the findings and which one is in
// hand, the spelling pass and where it has got to, the dialogs — and asks
// the app for the few things only it knows through an `adapter`:
//
//   goTo(where, issue)        show and select a finding's object, cell or paragraph
//   fix(issue, fix)           the fixes that are the app's own ops
//   altText(target, props)    write an object's description or decorative mark
//   renameSheet / slideTitle  (Worksheets, Presentations) the prompts' ops
//   openLayers(target)        (Presentation) the reading-order fix
//   spellArgs()               what the pass starts from beyond the caret: the slide on screen
//   showWord(found)           select the misspelt word where it is
//   whereLabel(found)         "Footnote 2", "Sales!B4", "Slide 3 notes"
//   asYouType                 (Word) Chromium's own underline, on or off
//   thesaurus                 { word() -> { word, ... } | null, replace(found, word) } the word at the caret, and putting another in its place
//
// The panes and dialogs are office-ui's (proofing.js); the engines' side is
// main/proofing.js, behind `doc.proof` and the document's own ops.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { t, tn } from '@rutba/office-ui';
import { wordAround } from './word-around.js';
import { AccessibilityPane, EditorPane, ThesaurusPane, AltTextDialog, DictionaryDialog, PromptDialog, A11yStatus, installProofingStyles } from '@rutba/office-ui/proofing';

installProofingStyles();

const KEEP_KEY = 'proofing.keepChecking';

/** A word in the capitals of the one it replaces: Capitalised, CAPITALS or as it is. */
function caseLike(original, word) {
  const o = String(original || '');
  if (o.length > 1 && o === o.toUpperCase() && o !== o.toLowerCase()) return word.toUpperCase();
  if (o[0] && o[0] === o[0].toUpperCase() && o[0] !== o[0].toLowerCase()) return word[0].toUpperCase() + word.slice(1);
  return word;
}

export { wordAround };

export function useReview({ shell, doc, model, apply, toast, adapter }) {
  const [pane, setPane] = useState(null);
  const [a11y, setA11y] = useState({ result: null, loading: false, selectedKey: null });
  const [keepRunning, setKeepRunning] = useState(false);
  const [spell, setSpell] = useState(null);
  const [options, setOptions] = useState(null);
  const [dialog, setDialog] = useState(null);
  const [busyFix, setBusyFix] = useState(false);
  const docRef = useRef(doc);
  docRef.current = doc;
  const adapterRef = useRef(adapter);
  adapterRef.current = adapter;
  const spellRef = useRef(spell);
  spellRef.current = spell;

  const proof = useCallback(async (action, args = {}) => {
    const d = docRef.current;
    if (!d) return null;
    try {
      return await shell.doc.proof({ id: d.id, action, ...args });
    } catch (err) {
      toast?.(err.message, { tone: 'bad' });
      return null;
    }
  }, [shell, toast]);

  // A document opened: have its dictionary read in the background, a moment
  // after the window has drawn, so a right-click on a misspelt word is
  // answered at once. Nothing is shown if it cannot be.
  const docId = doc?.id;
  useEffect(() => {
    if (!docId) return undefined;
    const t = setTimeout(() => { Promise.resolve(shell.doc?.proof?.({ id: docId, action: 'spellWarm' })).catch(() => {}); }, 1500);
    return () => clearTimeout(t);
  }, [shell, docId]);

  /* ── accessibility ───────────────────────────────────────────────── */

  useEffect(() => {
    let live = true;
    Promise.resolve(shell.store?.get?.({ key: KEEP_KEY, fallback: false })).then((v) => { if (live) setKeepRunning(Boolean(v)); }).catch(() => {});
    return () => { live = false; };
  }, [shell]);

  const check = useCallback(async () => {
    setA11y((s) => ({ ...s, loading: true }));
    const result = await proof('accessibility');
    setA11y((s) => ({
      result: result || s.result,
      loading: false,
      selectedKey: result?.issues?.some((i) => i.key === s.selectedKey) ? s.selectedKey : null,
    }));
    return result;
  }, [proof]);

  // Kept running: the findings follow the document as it is edited, a
  // moment after the typing stops, and the status bar says what they are.
  const version = doc?.version;
  useEffect(() => {
    if (!doc || !(keepRunning || pane === 'accessibility')) return undefined;
    const t = setTimeout(() => { check(); }, a11y.result ? 900 : 0);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc?.id, version, keepRunning, pane]);

  const setKeep = useCallback((on) => {
    setKeepRunning(on);
    shell.store?.set?.({ key: KEEP_KEY, value: Boolean(on) }).catch?.(() => {});
    if (on) check();
  }, [shell, check]);

  const select = useCallback((issue) => {
    setA11y((s) => ({ ...s, selectedKey: issue.key }));
    if (issue.where) adapterRef.current?.goTo?.(issue.where, issue);
  }, []);

  const openAltText = useCallback(async (target, info = null) => {
    const current = info || (await proof('altTextOf', target)) || {};
    setDialog({ kind: 'altText', target: current.target || target, name: current.name || '', descr: current.descr || '', decorative: Boolean(current.decorative) });
  }, [proof]);

  const fix = useCallback(async (issue, f) => {
    const a = adapterRef.current || {};
    switch (f.kind) {
      case 'altText':
        return openAltText(f.target, null);
      case 'decorative':
        setBusyFix(true);
        await a.altText?.(f.target, { descr: '', decorative: true });
        break;
      case 'setTitle':
        return setDialog({ kind: 'prompt', title: t('Document title'), label: t('The title screen readers announce and File → Info shows:'), initial: f.suggestion || '', className: 'pf-title', run: (title) => apply({ op: 'setDocTitle', title }) });
      case 'renameSheet':
        return setDialog({ kind: 'prompt', title: t('Rename sheet'), label: t('A name that says what is on "{sheet}":', { sheet: f.target }), initial: f.target, className: 'pf-rename', run: (to) => a.renameSheet?.(f.target, to) });
      case 'slideTitle':
        return setDialog({ kind: 'prompt', title: t('Slide title'), label: t('The title of slide {number}:', { number: f.target.slide + 1 }), initial: '', className: 'pf-slidetitle', run: (title) => a.slideTitle?.(f.target, title) });
      case 'layers':
        a.openLayers?.(f.target);
        return undefined;
      default:
        setBusyFix(true);
        await a.fix?.(issue, f);
    }
    setBusyFix(false);
    await check();
    return undefined;
  }, [apply, check, openAltText]);

  const openAccessibility = useCallback(() => {
    setPane((p) => (p === 'accessibility' ? null : 'accessibility'));
  }, []);

  /* ── spelling ────────────────────────────────────────────────────── */

  const step = useCallback(async (from, stop, wrapped, extra = {}) => {
    setSpell((s) => ({ ...s, checking: true }));
    const args = adapterRef.current?.spellArgs?.() || {};
    const res = await proof('spellNext', { from, stop, wrapped, ...args, ...extra });
    if (!res) { setSpell((s) => ({ ...s, checking: false })); return; }
    if (res.done) {
      const more = adapterRef.current?.morePass?.(spellRef.current) || null;
      setSpell((s) => ({ ...s, phase: 'done', found: null, checking: false, message: more?.message || t('You\'re good to go.'), more }));
      return;
    }
    const found = { ...res.found, suggestions: null, whereLabel: adapterRef.current?.whereLabel?.(res.found) || '' };
    setSpell((s) => ({ ...s, phase: 'word', found, next: res.next, wrapped: res.wrapped, checking: false }));
    adapterRef.current?.showWord?.(found);
    // The suggestions follow — kept only if the pass is still on this word.
    const answer = await proof('spellSuggest', { word: found.word });
    setSpell((s) => (s?.found && s.found.key === found.key && s.found.offset === found.offset
      ? { ...s, found: { ...s.found, suggestions: answer?.suggestions || [] } }
      : s));
  }, [proof]);

  const startSpelling = useCallback(async (extra = {}) => {
    setPane('editor');
    setSpell({ phase: 'loading', message: t('Reading the dictionary…'), checking: true });
    const args = { ...(adapterRef.current?.spellArgs?.() || {}), ...extra };
    const res = await proof('spellStart', args);
    if (!res) return setSpell({ phase: 'error', message: t('The spelling check could not start.') });
    setOptions(res.options);
    setSpell({ phase: 'loading', language: res.languageName, start: res.start, message: t('Checking…'), checking: true, extra });
    if (!res.segments || !res.start?.key) {
      setSpell({ phase: 'done', language: res.languageName, message: t('There is no text to check.') });
      return undefined;
    }
    return step(res.start, res.start, false, extra);
  }, [proof, step]);

  const carryOn = useCallback((from) => {
    const s = spellRef.current;
    if (!s) return;
    step(from, s.start, Boolean(s.wrapped), s.extra || {});
  }, [step]);

  const change = useCallback(async (replacement) => {
    const s = spellRef.current;
    if (!s?.found) return;
    const f = s.found;
    setSpell((x) => ({ ...x, checking: true }));
    await apply({ op: 'spellReplace', edits: [{ key: f.key, from: f.offset, to: f.offset + f.length, text: replacement }] });
    carryOn({ key: f.key, offset: f.offset + replacement.length });
  }, [apply, carryOn]);

  const changeAll = useCallback(async (replacement) => {
    const s = spellRef.current;
    if (!s?.found) return;
    const f = s.found;
    setSpell((x) => ({ ...x, checking: true }));
    const next = await apply({ op: 'spellChangeAll', word: f.word, replacement, ...(s.extra || {}), ...(adapterRef.current?.spellArgs?.() || {}) });
    const n = Number(next?.opResult) || 0;
    if (n > 1) toast?.(tn(n, 'Changed {count} time.', 'Changed {count} times.'), { ms: 2500 });
    carryOn({ key: f.key, offset: f.offset + replacement.length });
  }, [apply, carryOn, toast]);

  const ignoreOnce = useCallback(() => {
    const s = spellRef.current;
    if (s?.next) carryOn(s.next);
  }, [carryOn]);

  const ignoreAll = useCallback(async () => {
    const s = spellRef.current;
    if (!s?.found) return;
    await proof('ignoreAll', { word: s.found.word });
    carryOn(s.next);
  }, [proof, carryOn]);

  const addWord = useCallback(async (word) => {
    const w = word ?? spellRef.current?.found?.word;
    if (!w) return;
    await proof('addWord', { word: w });
    if (word == null) carryOn(spellRef.current.next);
    else toast?.(t('"{word}" is in your dictionary.', { word: w }), { ms: 2500 });
  }, [proof, carryOn, toast]);

  const setOption = useCallback(async (patch) => {
    const res = await proof('options', { options: patch });
    if (res?.options) setOptions(res.options);
  }, [proof]);

  const openDictionary = useCallback(async () => {
    const res = await proof('dictionary');
    setDialog({ kind: 'dictionary', words: res?.words || [] });
  }, [proof]);

  // F7 — Word's, Excel's and PowerPoint's key for the spelling pass.
  useEffect(() => {
    const onKey = (e) => {
      if (e.key !== 'F7' || e.ctrlKey || e.altKey || e.metaKey || e.shiftKey) return;
      e.preventDefault();
      startSpelling();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [startSpelling]);

  /** A right-click on a word: is it misspelt, and what might it be? */
  const checkWord = useCallback((word) => proof('spellCheckWord', { word }), [proof]);

  /* ── thesaurus ───────────────────────────────────────────────────── */

  // { query, result, loading, found (the word in the document), trail (words looked up before) }
  const [thes, setThes] = useState(null);
  const lookUp = useCallback(async (word, { found, back = false } = {}) => {
    setThes((s) => ({
      query: word,
      result: null,
      loading: true,
      found: found !== undefined ? found : s?.found ?? null,
      trail: back ? (s?.trail || []).slice(0, -1) : found !== undefined ? [] : [...(s?.trail || []), ...(s?.query ? [s.query] : [])],
    }));
    const result = word ? await proof('thesaurus', { word }) : { word: '', meanings: [] };
    setThes((s) => (s && s.query === word ? { ...s, result: result || { word, meanings: [] }, loading: false } : s));
  }, [proof]);
  const openThesaurus = useCallback(() => {
    const found = adapterRef.current?.thesaurus?.word?.() || null;
    setPane('thesaurus');
    lookUp(found?.word || '', { found });
  }, [lookUp]);
  const insertWord = useCallback(async (word) => {
    const th = thes;
    const a = adapterRef.current?.thesaurus;
    if (!a || !th?.found) { toast?.(t('Select a word in the document to put this one in its place.'), { ms: 3500 }); return; }
    const put = caseLike(th.found.word, word);
    await a.replace(th.found, put);
    // What is in the document now is the word just put there.
    setThes((s) => (s ? { ...s, found: { ...s.found, word: put, end: (s.found.start ?? 0) + put.length } } : s));
  }, [thes, toast]);

  // Shift+F7 — Word's key for the thesaurus, on the word at the caret.
  useEffect(() => {
    const onKey = (e) => {
      if (e.key !== 'F7' || !e.shiftKey || e.ctrlKey || e.altKey || e.metaKey) return;
      e.preventDefault();
      openThesaurus();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [openThesaurus]);

  /* ── what the window draws ───────────────────────────────────────── */

  const close = useCallback(() => setPane(null), []);

  const paneNode = pane === 'accessibility' ? (
    <AccessibilityPane
      result={a11y.result}
      loading={a11y.loading}
      selectedKey={a11y.selectedKey}
      onSelect={select}
      onFix={fix}
      busyFix={busyFix}
      keepRunning={keepRunning}
      onKeepRunning={setKeep}
      onRecheck={check}
    />
  ) : pane === 'thesaurus' ? (
    <ThesaurusPane
      state={thes ? { ...thes, back: thes.trail?.length ? thes.trail[thes.trail.length - 1] : null } : null}
      onSearch={(w) => lookUp(w)}
      onBack={() => thes?.trail?.length && lookUp(thes.trail[thes.trail.length - 1], { back: true })}
      onInsert={adapter?.thesaurus ? insertWord : null}
    />
  ) : pane === 'editor' ? (
    <EditorPane
      state={spell ? { ...spell, onMore: undefined } : null}
      options={options}
      asYouType={adapter?.asYouType || null}
      onChange={change}
      onChangeAll={changeAll}
      onIgnoreOnce={ignoreOnce}
      onIgnoreAll={ignoreAll}
      onAdd={() => addWord()}
      onOption={setOption}
      onDictionary={openDictionary}
      onClose={close}
      onRestart={spell?.more ? () => startSpelling(spell.more.extra) : null}
      restartLabel={spell?.more?.label}
    />
  ) : null;

  const dialogs = dialog?.kind === 'altText' ? (
    <AltTextDialog
      name={dialog.name}
      initial={dialog.descr}
      decorative={dialog.decorative}
      onClose={() => setDialog(null)}
      onApply={async (props) => {
        setDialog(null);
        await adapterRef.current?.altText?.(dialog.target, props);
        if (pane === 'accessibility' || keepRunning) check();
      }}
    />
  ) : dialog?.kind === 'prompt' ? (
    <PromptDialog
      title={dialog.title}
      label={dialog.label}
      initial={dialog.initial}
      className={dialog.className}
      onClose={() => setDialog(null)}
      onApply={async (value) => {
        setDialog(null);
        await dialog.run(value);
        check();
      }}
    />
  ) : dialog?.kind === 'dictionary' ? (
    <DictionaryDialog
      words={dialog.words}
      onClose={() => setDialog(null)}
      onSave={async (words) => {
        setDialog(null);
        await proof('dictionary', { words });
      }}
      onImport={async () => {
        const [file] = (await shell.dialog.open({ title: t('Import a word list'), filters: [{ name: t('Word lists'), extensions: ['dic', 'txt'] }] })) || [];
        if (!file) return [];
        const res = await proof('dictionaryImport', { path: file });
        return res?.added || [];
      }}
      onExport={async (words) => {
        const file = await shell.dialog.save({ title: t('Export the dictionary'), defaultPath: 'Custom.dic', filters: [{ name: t('Word lists'), extensions: ['dic'] }] });
        if (!file) return;
        const res = await proof('dictionaryExport', { path: file, words });
        if (res) toast?.(tn(res.count, '{count} word saved.', '{count} words saved.'), { ms: 2500 });
      }}
    />
  ) : null;

  const status = a11y.result && (keepRunning || pane === 'accessibility') ? (
    <A11yStatus verdict={a11y.result.verdict} onClick={() => setPane('accessibility')} />
  ) : null;

  const paneTitle = pane === 'accessibility' ? t('Accessibility') : pane === 'editor' ? t('Editor') : pane === 'thesaurus' ? t('Thesaurus') : null;

  return {
    pane, paneTitle, paneNode, dialogs, status, close,
    openAccessibility, startSpelling, openAltText, openDictionary, checkWord, addWord, check, openThesaurus,
    ignoreWord: (word) => proof('ignoreAll', { word }),
  };
}
