// References — Citations & Bibliography, laid out as Word 365 lays them out.
//
// Insert Citation drops a source's citation at the caret, or a placeholder
// to fill in later; Manage Sources moves sources between the master list,
// kept in this profile (as Word keeps Sources.xml), and the document's own
// list; Create Source asks for the fields Word asks for each type; Style
// re-draws every citation and the bibliography; Bibliography inserts the
// gallery's building block. The formatting the dialogs preview is
// `@rutba/ooxml/bibliography`'s — the very functions the engine writes with.

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { Button, Dialog, Group, Input, Select, Icon } from '@rutba/office-ui';
import {
  BIBLIOGRAPHY_STYLES, SOURCE_TYPES, PERSON_ROLES, fieldLabel, styleById, makeTag, parseNames, namesText,
  formatCitation, formatBibliographyEntry, describeSource, parseSources, sourcesXml, newGuid, segmentsText,
} from '@rutba/ooxml/bibliography';
import { MarkEntryPanel, IndexDialog, INDEX_CSS } from './references-index.js';
import { FiguresDialog, FieldDialog, FIGURES_CSS } from './references-figures.js';

/** Where the master list lives in the profile: Word's Sources.xml, as the same XML. */
const MASTER_KEY = 'word.bibliography.master';

/** The language a citation is written in — `\l` — from the window's own. */
function lcid() {
  const lang = (typeof navigator !== 'undefined' && navigator.language) || 'en-US';
  return { 'en-GB': 2057, 'en-US': 1033, 'en-AU': 3081, 'en-CA': 4105, 'en-IE': 6153, 'en-NZ': 5129, 'en-IN': 16393, 'en-ZA': 7177 }[lang] || 1033;
}

const clone = (s) => JSON.parse(JSON.stringify(s));
const sameTag = (a, b) => String(a || '').toLowerCase() === String(b || '').toLowerCase();

/** A source as a list row reads: people; title (year). */
const rowText = (s) => describeSource(s).line || s.tag;

/** A citation or entry drawn with its italics, for a preview. */
function Segments({ segments }) {
  return (segments || []).map((s, i) => (s.italic ? <i key={i}>{s.text}</i> : <span key={i}>{s.text}</span>));
}

/* ── the ribbon group ────────────────────────────────────────────────────── */

/** References → Citations & Bibliography, as Word lays the group out. */
export function CitationsGroup({ refs, menu }) {
  const info = refs.info;
  const style = info?.style || 'apa7';
  return (
    <Group label="Citations & Bibliography">
      <Button
        tall icon="reply" className="wd-refs-2line" label={'Insert\nCitation'}
        title="Insert Citation — cite a source in this document's list, add a new source, or a placeholder to fill in later"
        onClick={(e) => menu.open(e, refs.citationMenu())}
      />
      <div className="wd-refs-col">
        <Button icon="list" label="Manage Sources" title="Manage Sources — the master list and this document's list, side by side" onClick={() => refs.open('manage')} />
        <label className="wd-refs-style" data-tip="Style — how every citation and the bibliography are written">
          <span>Style:</span>
          <Select className="rw-select wd-refs-style-select" value={style} onChange={(e) => refs.setStyle(e.target.value)} title="Citation and bibliography style">
            {BIBLIOGRAPHY_STYLES.map((s) => <option key={s.id} value={s.id} title={`${s.label} ${s.edition}`}>{s.name === 'APA' ? `APA ${s.version}th` : s.label}</option>)}
          </Select>
        </label>
        <Button
          icon="listBullet" label="Bibliography" className="wd-refs-bib"
          title="Bibliography — a list of the sources, headed Bibliography, References or Works Cited"
          onClick={(e) => menu.open(e, refs.bibliographyMenu())}
        >
          <Icon name="chevronDown" size={12} />
        </Button>
      </div>
    </Group>
  );
}

/* ── the hook ────────────────────────────────────────────────────────────── */

/**
 * The References tab's citation verbs and dialogs, for word.js: `info` is
 * the model's references (sources, style, citations), `node` the dialogs.
 */
export function useReferences({ shell, doc = null, model, apply, toast, layout = () => null, patchView = null }) {
  const info = model?.references || null;
  const [dialog, setDialog] = useState(null);
  const [marking, setMarking] = useState(false);
  const [master, setMaster] = useState([]);

  // The master list, read from the profile when a dialog needs it.
  const loadMaster = useCallback(async () => {
    try {
      const xml = await shell.store.get({ key: MASTER_KEY, fallback: '' });
      const list = xml ? parseSources(xml).sources : [];
      setMaster(list);
      return list;
    } catch {
      return [];
    }
  }, [shell]);
  const saveMaster = useCallback(async (list) => {
    setMaster(list);
    try { await shell.store.set({ key: MASTER_KEY, value: sourcesXml({ sources: list }) }); } catch { /* the profile is read-only: the list lives for this session */ }
  }, [shell]);

  const open = useCallback(async (kind, arg = null) => {
    if (kind === 'manage' || kind === 'source') await loadMaster();
    setDialog({ kind, ...(arg || {}) });
  }, [loadMaster]);
  const close = useCallback(() => setDialog(null), []);

  const setStyle = useCallback((style) => apply({ op: 'setBibliographyStyle', style }), [apply]);

  const insertCitation = useCallback(async (tags) => {
    await apply({ op: 'insertCitation', tags, lcid: lcid() });
  }, [apply]);

  const citationMenu = useCallback(() => {
    const sources = info?.sources || [];
    const items = sources.map((s) => {
      const d = describeSource(s);
      return {
        label: (
          <span className="wd-refs-menu-source">
            <b>{d.who || d.title || s.tag}</b>
            <small>{[d.who ? d.title : '', d.year ? `(${d.year})` : ''].filter(Boolean).join(' ')}</small>
          </span>
        ),
        title: `Cite ${rowText(s)}`,
        run: () => insertCitation([s.tag]),
      };
    });
    return [
      ...items,
      ...(items.length ? ['-'] : []),
      { label: 'Add New Source…', icon: 'plus', run: () => open('source', { mode: 'new', cite: true }) },
      { label: 'Add New Placeholder…', icon: 'textbox', run: () => open('placeholder') },
    ];
  }, [info, insertCitation, open]);

  const bibliographyMenu = useCallback(() => [
    { heading: true, label: 'Built-In' },
    { label: 'Bibliography', icon: 'listBullet', run: () => apply({ op: 'insertBibliography', heading: 'Bibliography' }) },
    { label: 'References', icon: 'listBullet', run: () => apply({ op: 'insertBibliography', heading: 'References' }) },
    { label: 'Works Cited', icon: 'listBullet', run: () => apply({ op: 'insertBibliography', heading: 'Works Cited' }) },
    '-',
    { label: 'Insert Bibliography', icon: 'plus', run: () => apply({ op: 'insertBibliography', heading: null }) },
    { label: 'Update Citations and Bibliography', icon: 'refresh', disabled: !info?.citations?.length && !info?.bibliography, run: async () => {
      const next = await apply({ op: 'updateCitations' });
      if (next) toast('Citations and bibliography updated', { tone: 'good' });
    } },
  ], [apply, info, toast]);

  const saveDocSources = useCallback((sources) => apply({ op: 'setSources', sources }), [apply]);

  // The words selected, when they lie in one paragraph — Mark Entry's main entry.
  const sel = model?.selection;
  const selected = (() => {
    const a = sel?.anchor;
    const f = sel?.focus;
    if (!a || !f || a.block !== f.block || a.offset === f.offset) return '';
    const text = model?.blocks?.[a.block]?.text || '';
    return text.slice(Math.min(a.offset, f.offset), Math.max(a.offset, f.offset)).replace(/[\t\n]+/g, ' ').trim();
  })();

  /** The References verbs that open something or need the page layout: Mark Entry, Insert Index, Update Index. */
  const act = useCallback(async (name) => {
    if (name === 'markEntry') {
      setMarking(true);
      return;
    }
    if (name === 'insertIndex') {
      setDialog({ kind: 'index' });
      return;
    }
    if (name === 'insertFigures') {
      setDialog({ kind: 'figures' });
      return;
    }
    if (name === 'updateFigures') {
      if (!info?.figures?.length) return;
      const next = await apply({ op: 'updateTablesOfFigures', pages: layout()?.pages || null });
      if (next) toast('Table of figures updated', { tone: 'good' });
      return;
    }
    if (name === 'field') {
      setDialog({ kind: 'field' });
      return;
    }
    if (name === 'updateIndex') {
      if (!info?.index) return;
      const next = await apply({ op: 'updateIndex', pages: layout()?.pages || null });
      if (next) toast('Index updated', { tone: 'good' });
    }
  }, [apply, info, layout, toast]);

  /** What a field is worked out from: the page each paragraph is on, how many, the file's name, now. */
  const fieldContext = useCallback(() => {
    const l = layout() || {};
    const filePath = doc?.path || null;
    const fileName = doc?.name || (filePath ? filePath.split(/[\\/]/).pop() : null);
    return { pages: l.pages || null, pageCount: l.count || null, fileName, filePath, now: new Date().toISOString() };
  }, [layout, doc]);

  /** Before printing: page numbers, dates and file names in the body worked out again. */
  const beforePrint = useCallback(async () => {
    if (!info?.docFields) return;
    await apply({ op: 'refreshReferences', ...fieldContext(), fieldsOnly: true });
  }, [apply, info, fieldContext]);

  /** Insert → Quick Parts. */
  const quickParts = useCallback(() => [
    { label: 'Field…', icon: 'formula', run: () => setDialog({ kind: 'field' }) },
    { heading: true, label: 'Document Property' },
    { label: 'Author', run: () => apply({ op: 'insertDocField', name: 'AUTHOR', ...fieldContext() }) },
    { label: 'Title', run: () => apply({ op: 'insertDocField', name: 'TITLE', ...fieldContext() }) },
  ], [apply, fieldContext]);

  const markEntry = useCallback(async (spec) => {
    const next = await apply({ op: 'markIndexEntry', ...spec });
    // Word shows the hidden XE field it just wrote: ¶ goes on.
    if (next && patchView) patchView({ marks: true });
    return next?.opResult ?? (next ? 1 : 0);
  }, [apply, patchView]);

  let node = null;
  if (dialog?.kind === 'manage') {
    node = (
      <ManageSourcesDialog
        style={info?.style || 'apa7'}
        current={info?.sources || []}
        cited={new Set((info?.citations || []).flatMap((c) => c.tags))}
        master={master}
        onClose={async ({ current, master: nextMaster, changed }) => {
          close();
          if (changed.master) await saveMaster(nextMaster);
          if (changed.current) await saveDocSources(current);
        }}
      />
    );
  } else if (dialog?.kind === 'source') {
    const all = [...(info?.sources || []), ...master];
    node = (
      <SourceDialog
        style={info?.style || 'apa7'}
        initial={dialog.source || null}
        taken={all.map((s) => s.tag).filter((t) => !dialog.source || !sameTag(t, dialog.source.tag))}
        onClose={close}
        onSave={async (source) => {
          close();
          // A new source joins both lists, as Word adds it to both.
          const inMaster = master.some((s) => sameTag(s.tag, source.tag));
          await saveMaster(inMaster ? master.map((s) => (sameTag(s.tag, source.tag) ? source : s)) : [...master, source]);
          const current = info?.sources || [];
          const inDoc = current.some((s) => sameTag(s.tag, source.tag));
          await saveDocSources(inDoc ? current.map((s) => (sameTag(s.tag, source.tag) ? source : s)) : [...current, source]);
          if (dialog.cite) await insertCitation([source.tag]);
        }}
      />
    );
  } else if (dialog?.kind === 'placeholder') {
    const taken = new Set([...(info?.sources || []).map((s) => s.tag), ...(info?.citations || []).flatMap((c) => c.tags)].map((t) => t.toLowerCase()));
    let n = 1;
    while (taken.has('placeholder' + n)) n += 1;
    node = (
      <PlaceholderDialog
        initial={'Placeholder' + n}
        onClose={close}
        onOk={async (name) => {
          close();
          await insertCitation([name]);
        }}
      />
    );
  } else if (dialog?.kind === 'index') {
    node = (
      <IndexDialog
        current={info?.index || null}
        onClose={close}
        onOk={async (opts) => {
          close();
          await apply({ op: 'insertIndex', ...opts, lcid: lcid(), pages: layout()?.pages || null });
        }}
      />
    );
  }
  if (dialog?.kind === 'figures') {
    const l = layout() || {};
    const captions = Object.fromEntries(Object.entries(info?.captions || {}).map(([k, list]) => [k, list.map((c) => ({ ...c, page: l.pages ? l.pages[c.block] : null }))]));
    const labels = ['Figure', 'Table', 'Equation'];
    const first = labels.find((k) => captions[k]?.length) || 'Figure';
    node = (
      <FiguresDialog
        labels={[first, ...labels.filter((k) => k !== first)]}
        captions={captions}
        onClose={close}
        onOk={async (opts) => {
          close();
          await apply({ op: 'insertTableOfFigures', ...opts, pages: layout()?.pages || null });
        }}
      />
    );
  } else if (dialog?.kind === 'field') {
    const ctx = fieldContext();
    const focus = model?.selection?.focus;
    node = (
      <FieldDialog
        context={{ page: ctx.pages && focus ? ctx.pages[focus.block] : 1, pages: ctx.pageCount || 1, now: new Date(), fileName: ctx.fileName, filePath: ctx.filePath, author: info?.properties?.author ?? null, title: info?.properties?.title ?? null }}
        onClose={close}
        onOk={async (spec) => {
          close();
          const c = fieldContext();
          await apply({ op: 'insertDocField', ...spec, ...c, page: c.pages && focus ? c.pages[focus.block] ?? null : null });
        }}
      />
    );
  }
  const panel = marking ? (
    <MarkEntryPanel selected={selected} bookmarks={model?.bookmarks || []} onMark={markEntry} onClose={() => setMarking(false)} />
  ) : null;

  return { info, open, setStyle, citationMenu, bibliographyMenu, act, quickParts, beforePrint, fieldContext, node: <>{node}{panel}</>, dialog };
}

/* ── Create Source / Edit Source ─────────────────────────────────────────── */

/** What each field's box says it wants, as Word's Example line does. */
const EXAMPLES = {
  Author: 'Smith, John A.; Jones, Bob',
  Editor: 'White, Eve',
  Translator: 'Grey, Al',
  BookAuthor: 'Black, Ann',
  Title: 'The Book of Things',
  BookTitle: 'Collected Papers',
  JournalName: 'Journal of Tests',
  ConferenceName: 'Proceedings of the Annual Meeting',
  InternetSiteTitle: 'Rutba',
  Year: '2020',
  Month: 'March',
  Day: '5',
  YearAccessed: '2024',
  MonthAccessed: 'June',
  DayAccessed: '1',
  City: 'London',
  Publisher: 'Penguin',
  Pages: '45-67',
  Volume: '12',
  Issue: '3',
  Edition: '2',
  URL: 'https://www.example.com',
  DOI: '10.1000/xyz123',
};

function emptySource(type = 'Book') {
  return { tag: '', type, guid: newGuid(), fields: {}, people: {}, refOrder: null, extra: '' };
}

/**
 * Create Source (or Edit Source): the type, the fields Word shows for it —
 * all of them with Show All Bibliography Fields — and the tag, made from the
 * first author and the year until it is typed over.
 */
export function SourceDialog({ style, initial, taken = [], onClose, onSave }) {
  const [source, setSource] = useState(() => (initial ? clone(initial) : emptySource()));
  const [names, setNames] = useState(() => Object.fromEntries(PERSON_ROLES.map((r) => [r, namesText(initial?.people?.[r])])));
  const [corporate, setCorporate] = useState(() => Boolean(initial?.people?.Author?.corporate));
  const [all, setAll] = useState(false);
  const [tagTouched, setTagTouched] = useState(Boolean(initial));
  const [focus, setFocus] = useState('Author');
  const type = SOURCE_TYPES.find((t) => t.id === source.type) || SOURCE_TYPES[0];
  const keys = all ? [...type.fields, ...type.more] : type.fields;

  // People from the boxes: a corporate author is one name, as typed.
  const withPeople = useCallback((s) => {
    const people = {};
    for (const role of PERSON_ROLES) {
      const text = (names[role] || '').trim();
      if (!text) continue;
      people[role] = role === 'Author' && corporate ? { corporate: text } : { list: parseNames(text) };
    }
    return { ...s, people };
  }, [names, corporate]);

  const draft = withPeople(source);
  const autoTag = makeTag(draft, taken);
  const tag = tagTouched ? source.tag : autoTag;
  const clash = taken.some((t) => sameTag(t, tag));
  const empty = !Object.values(draft.fields).some((v) => String(v || '').trim()) && !Object.keys(draft.people).length;
  const s = styleById(style);

  const setField = (key, value) => setSource((x) => ({ ...x, fields: { ...x.fields, [key]: value } }));
  const ok = () => {
    if (empty || !tag || clash) return;
    const fields = Object.fromEntries(Object.entries(draft.fields).filter(([, v]) => String(v || '').trim()).map(([k, v]) => [k, String(v).trim()]));
    onSave({ ...draft, tag, fields });
  };

  return (
    <Dialog
      title={initial ? 'Edit Source' : 'Create Source'}
      width={620}
      onClose={onClose}
      actions={
        <>
          <span className="wd-refs-foot-note">{clash ? `A source is already tagged ${tag}.` : ''}</span>
          <Button label="Cancel" onClick={onClose} />
          <Button primary className="wd-refs-source-ok" label="OK" disabled={empty || !tag || clash} onClick={ok} />
        </>
      }
    >
      <div className="wd-refs-source" onKeyDown={(e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } }}>
        <div className="wd-refs-type">
          <label htmlFor="wd-refs-type">Type of Source</label>
          <Select id="wd-refs-type" className="rw-select wd-refs-type-select" value={source.type} onChange={(e) => setSource((x) => ({ ...x, type: e.target.value }))}>
            {SOURCE_TYPES.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
          </Select>
        </div>
        <div className="wd-refs-section">Bibliography Fields for {s.label}</div>
        <div className="wd-refs-grid">
          {keys.map((key) => {
            const person = PERSON_ROLES.includes(key);
            const label = fieldLabel(source.type, key);
            return (
              <React.Fragment key={key}>
                <label htmlFor={`wd-refs-f-${key}`}>{label}</label>
                <div className="wd-refs-cell">
                  <Input
                    id={`wd-refs-f-${key}`}
                    className={`rw-input wd-refs-f wd-refs-f-${key}`}
                    value={person ? names[key] || '' : source.fields[key] || ''}
                    onFocus={() => setFocus(key)}
                    onChange={(e) => (person ? setNames((n) => ({ ...n, [key]: e.target.value })) : setField(key, e.target.value))}
                    autoFocus={key === 'Author'}
                  />
                  {key === 'Author' ? (
                    <label className="wd-refs-check">
                      <input type="checkbox" checked={corporate} onChange={(e) => setCorporate(e.target.checked)} />
                      Corporate Author
                    </label>
                  ) : null}
                </div>
              </React.Fragment>
            );
          })}
        </div>
        <label className="wd-refs-check wd-refs-all">
          <input type="checkbox" className="wd-refs-showall" checked={all} onChange={(e) => setAll(e.target.checked)} />
          Show All Bibliography Fields
        </label>
        <div className="wd-refs-bottom">
          <div className="wd-refs-tag">
            <label htmlFor="wd-refs-tag">Tag name</label>
            <Input id="wd-refs-tag" className="rw-input wd-refs-tagname" value={tag} onChange={(e) => { setTagTouched(true); setSource((x) => ({ ...x, tag: e.target.value.replace(/\s+/g, '') })); }} />
          </div>
          <div className="wd-refs-example">
            <span>Example:</span> {focus === 'Author' && corporate ? 'World Health Organization' : EXAMPLES[focus] || '—'}
          </div>
        </div>
        <div className="wd-refs-preview wd-refs-preview-one">
          <div className="wd-refs-preview-head">Preview ({s.label})</div>
          <div className="wd-refs-preview-body">
            {empty ? <span className="wd-refs-muted">Fill in the fields to see the entry.</span> : <div className="wd-refs-hang"><Segments segments={formatBibliographyEntry({ ...draft, tag }, style, { number: 1 })} /></div>}
          </div>
        </div>
      </div>
    </Dialog>
  );
}

/* ── Source Manager ──────────────────────────────────────────────────────── */

const SORTS = [
  ['author', 'Author'], ['tag', 'Tag'], ['title', 'Title'], ['year', 'Year'],
];
function sortBy(list, key) {
  const k = (s) => {
    const d = describeSource(s);
    return key === 'tag' ? s.tag : key === 'title' ? d.title : key === 'year' ? d.year : d.who || d.title;
  };
  return [...list].sort((a, b) => String(k(a)).localeCompare(String(k(b)), undefined, { sensitivity: 'base', numeric: true }));
}

/**
 * Manage Sources: the master list and the document's list, Copy between
 * them, Delete, Edit and New, a search and a sort, and a preview of the
 * selected source in the current style. Cited sources carry a tick,
 * placeholders a question mark; a cited source cannot be deleted from the
 * document's list, as in Word.
 */
export function ManageSourcesDialog({ style, current: startCurrent, master: startMaster, cited, onClose }) {
  const [current, setCurrent] = useState(() => clone(startCurrent));
  const [master, setMaster] = useState(() => clone(startMaster));
  const [changed, setChanged] = useState({ current: false, master: false });
  const [pick, setPick] = useState(() => (startCurrent[0] ? { list: 'current', tag: startCurrent[0].tag } : startMaster[0] ? { list: 'master', tag: startMaster[0].tag } : null));
  const [query, setQuery] = useState('');
  const [sort, setSort] = useState('author');
  const [editing, setEditing] = useState(null);
  const s = styleById(style);

  // Placeholders: a citation's tag with no source behind it.
  const placeholders = [...cited].filter((t) => !current.some((x) => sameTag(x.tag, t)));
  const match = (x) => !query || rowText(x).toLowerCase().includes(query.toLowerCase()) || x.tag.toLowerCase().includes(query.toLowerCase());
  const shownMaster = sortBy(master.filter(match), sort);
  const shownCurrent = sortBy(current.filter(match), sort);

  const selected = pick ? (pick.list === 'master' ? master : current).find((x) => sameTag(x.tag, pick.tag)) : null;
  const pickedPlaceholder = pick?.list === 'placeholder' ? pick.tag : null;
  const isCited = selected && pick.list === 'current' && cited.has(selected.tag);

  const setList = (which, list) => {
    if (which === 'master') setMaster(list); else setCurrent(list);
    setChanged((c) => ({ ...c, [which]: true }));
  };
  const copy = () => {
    if (!selected) return;
    const to = pick.list === 'master' ? 'current' : 'master';
    const target = to === 'master' ? master : current;
    if (target.some((x) => sameTag(x.tag, selected.tag))) {
      setList(to, target.map((x) => (sameTag(x.tag, selected.tag) ? clone(selected) : x)));
    } else setList(to, [...target, clone(selected)]);
  };
  const remove = () => {
    if (!selected || isCited) return;
    const from = pick.list;
    const list = (from === 'master' ? master : current).filter((x) => !sameTag(x.tag, selected.tag));
    setList(from, list);
    setPick(list[0] ? { list: from, tag: list[0].tag } : null);
  };
  const saveEdited = (source) => {
    const originalTag = editing?.source?.tag;
    const put = (list) => {
      const has = list.some((x) => sameTag(x.tag, originalTag || source.tag));
      return has ? list.map((x) => (sameTag(x.tag, originalTag || source.tag) ? source : x)) : [...list, source];
    };
    if (editing.mode === 'new') {
      setList('master', put(master));
      setList('current', put(current));
      setPick({ list: 'current', tag: source.tag });
    } else if (editing.mode === 'placeholder') {
      setList('current', [...current, source]);
      setPick({ list: 'current', tag: source.tag });
    } else {
      // An edit reaches both lists where the source is in both, as Word asks to.
      if (master.some((x) => sameTag(x.tag, originalTag))) setList('master', put(master));
      if (current.some((x) => sameTag(x.tag, originalTag))) setList('current', put(current));
      setPick({ list: pick.list, tag: source.tag });
    }
    setEditing(null);
  };

  const close = () => onClose({ current, master, changed });

  const row = (x, list) => {
    const on = pick && pick.list === list && sameTag(pick.tag, x.tag);
    return (
      <button key={list + x.tag} type="button" className={`wd-refs-row${on ? ' on' : ''}`} onClick={() => setPick({ list, tag: x.tag })} onDoubleClick={() => setEditing({ mode: 'edit', source: x })} title={rowText(x)}>
        <span className="wd-refs-mark">{list === 'current' && cited.has(x.tag) ? <Icon name="check" size={12} /> : null}</span>
        <span className="wd-refs-row-text">{rowText(x)}</span>
      </button>
    );
  };

  return (
    <>
      <Dialog
        title="Source Manager"
        width={860}
        onClose={close}
        actions={<Button primary className="wd-refs-manage-close" label="Close" onClick={close} />}
      >
        <div className="wd-refs-manage">
          <div className="wd-refs-toolbar">
            <label className="wd-refs-search">
              <Icon name="find" size={14} />
              <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search by author, title, tag or year" />
            </label>
            <label className="wd-refs-sort">
              Sort by
              <Select value={sort} onChange={(e) => setSort(e.target.value)}>
                {SORTS.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
              </Select>
            </label>
          </div>
          <div className="wd-refs-lists">
            <div className="wd-refs-listcol">
              <div className="wd-refs-listhead">Master List <span>{master.length}</span></div>
              <div className="wd-refs-list wd-refs-master">
                {shownMaster.map((x) => row(x, 'master'))}
                {!master.length ? <div className="wd-refs-empty">Sources you create are kept here, for every document.</div> : null}
              </div>
            </div>
            <div className="wd-refs-verbs">
              <Button className="wd-refs-copy" label={pick?.list === 'current' ? '← Copy' : 'Copy →'} disabled={!selected} title={pick?.list === 'current' ? 'Copy the source to the master list' : 'Copy the source to this document'} onClick={copy} />
              <Button className="wd-refs-delete" label="Delete" disabled={!selected || isCited} title={isCited ? 'Delete — a source cited in the document cannot be deleted from its list' : 'Delete the source from this list'} onClick={remove} />
              <Button className="wd-refs-edit" label="Edit…" disabled={!selected && !pickedPlaceholder} title="Edit the source" onClick={() => (pickedPlaceholder ? setEditing({ mode: 'placeholder', source: { ...emptySource(), tag: pickedPlaceholder } }) : setEditing({ mode: 'edit', source: selected }))} />
              <Button className="wd-refs-new" label="New…" title="Create a new source, in both lists" onClick={() => setEditing({ mode: 'new', source: null })} />
            </div>
            <div className="wd-refs-listcol">
              <div className="wd-refs-listhead">Current List <span>{current.length}</span></div>
              <div className="wd-refs-list wd-refs-current">
                {shownCurrent.map((x) => row(x, 'current'))}
                {placeholders.filter((t) => !query || t.toLowerCase().includes(query.toLowerCase())).map((t) => (
                  <button key={'p' + t} type="button" className={`wd-refs-row${pick?.list === 'placeholder' && sameTag(pick.tag, t) ? ' on' : ''}`} onClick={() => setPick({ list: 'placeholder', tag: t })} onDoubleClick={() => setEditing({ mode: 'placeholder', source: { ...emptySource(), tag: t } })}>
                    <span className="wd-refs-mark wd-refs-q">?</span>
                    <span className="wd-refs-row-text">{t}</span>
                  </button>
                ))}
                {!current.length && !placeholders.length ? <div className="wd-refs-empty">No sources in this document yet. Copy one across, or make a new one.</div> : null}
              </div>
              <div className="wd-refs-legend"><span><Icon name="check" size={11} /> cited source</span><span><b>?</b> placeholder source</span></div>
            </div>
          </div>
          <div className="wd-refs-preview">
            <div className="wd-refs-preview-head">Preview ({s.label})</div>
            <div className="wd-refs-preview-body">
              {selected ? (
                <>
                  <div><span className="wd-refs-muted">Citation:</span> <Segments segments={formatCitation([selected], style, { numbers: { [selected.tag]: 1 } })} /></div>
                  <div className="wd-refs-preview-entry"><span className="wd-refs-muted">Bibliography Entry:</span><div><Segments segments={formatBibliographyEntry(selected, style, { number: 1 })} /></div></div>
                </>
              ) : pickedPlaceholder ? (
                <span className="wd-refs-muted">{pickedPlaceholder} is a placeholder: Edit gives it a source, and every citation of it follows.</span>
              ) : (
                <span className="wd-refs-muted">Pick a source to see how it is cited.</span>
              )}
            </div>
          </div>
        </div>
      </Dialog>
      {editing ? (
        <SourceDialog
          style={style}
          initial={editing.source}
          taken={[...master, ...current].map((x) => x.tag).filter((t) => !editing.source || !sameTag(t, editing.source.tag))}
          onClose={() => setEditing(null)}
          onSave={saveEdited}
        />
      ) : null}
    </>
  );
}

/* ── placeholder ─────────────────────────────────────────────────────────── */

export function PlaceholderDialog({ initial, onClose, onOk }) {
  const [name, setName] = useState(initial);
  const valid = /^[\p{L}\p{N}_]+$/u.test(name);
  return (
    <Dialog
      title="Placeholder Name"
      width={420}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary className="wd-refs-placeholder-ok" label="OK" disabled={!valid} onClick={() => onOk(name)} />
        </>
      }
    >
      <div className="wd-refs-source" onKeyDown={(e) => { if (e.key === 'Enter' && valid) { e.preventDefault(); onOk(name); } }}>
        <p className="wd-refs-lead">Type a tag name for the placeholder. Fill it in later from Manage Sources — every citation of it follows.</p>
        <Input className="rw-input wd-refs-placeholder-name" value={name} onChange={(e) => setName(e.target.value.replace(/\s+/g, ''))} autoFocus />
      </div>
    </Dialog>
  );
}

/* ── the look ────────────────────────────────────────────────────────────── */

export const REFERENCES_CSS = `
.rw-btn.tall.wd-refs-2line > span { white-space: pre-line; line-height: 1.15; }
.wd-refs-col { display: flex; flex-direction: column; gap: 1px; justify-content: center; }
.wd-refs-style { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-2); padding: 0 6px; }
.wd-refs-style .rw-select.wd-refs-style-select { width: 132px; padding: 2px 6px; min-height: 24px; }
.rw-btn.wd-refs-bib svg:last-child { margin-left: 2px; opacity: 0.7; }
.wd-refs-menu-source { display: flex; flex-direction: column; line-height: 1.25; max-width: 340px; }
.wd-refs-menu-source b { font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.wd-refs-menu-source small { color: var(--ink-3); font-size: 11.5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.rw-dialog[aria-label="Source Manager"] { max-width: min(94vw, 900px); }
.wd-refs-source { display: flex; flex-direction: column; gap: 10px; padding-top: 2px; }
.wd-refs-type { display: flex; align-items: center; gap: 10px; font-size: 12.5px; }
.wd-refs-type label { font-weight: 600; color: var(--ink-2); }
.wd-refs-type .rw-select { min-width: 220px; }
.wd-refs-section { font-size: 12px; font-weight: 650; color: var(--ink); padding-bottom: 6px; border-bottom: 1px solid var(--line-soft); }
.wd-refs-grid { display: grid; grid-template-columns: 150px 1fr; gap: 7px 12px; align-items: start; max-height: 360px; overflow: auto; padding: 2px 4px 2px 0; }
.wd-refs-grid > label { text-align: right; font-size: 12.5px; color: var(--ink-2); padding-top: 7px; }
.wd-refs-cell { display: flex; flex-direction: column; gap: 5px; min-width: 0; }
.wd-refs-cell .rw-input { width: 100%; box-sizing: border-box; }
.wd-refs-check { display: inline-flex; align-items: center; gap: 7px; font-size: 12.5px; color: var(--ink-2); }
.wd-refs-all { margin-top: 2px; }
.wd-refs-bottom { display: grid; grid-template-columns: 200px 1fr; gap: 16px; align-items: end; }
.wd-refs-tag { display: flex; flex-direction: column; gap: 4px; }
.wd-refs-tag label { font-size: 11.5px; font-weight: 600; color: var(--ink-2); }
.wd-refs-example { font-size: 12px; color: var(--ink-3); padding-bottom: 7px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.wd-refs-example span { font-weight: 600; color: var(--ink-2); }
.wd-refs-preview { border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface-2); overflow: hidden; }
.wd-refs-preview-head { font-size: 11.5px; font-weight: 650; color: var(--ink-2); padding: 6px 10px; border-bottom: 1px solid var(--line-soft); }
.wd-refs-preview-body { padding: 9px 12px 11px; font-family: Calibri, 'Segoe UI', sans-serif; font-size: 13.5px; line-height: 1.45; color: var(--ink); display: flex; flex-direction: column; gap: 6px; min-height: 44px; }
.wd-refs-hang, .wd-refs-preview-entry > div { padding-left: 28px; text-indent: -28px; margin-top: 2px; }
.wd-refs-muted { color: var(--ink-3); font-family: var(--font-ui, 'Segoe UI', sans-serif); font-size: 12px; }
.wd-refs-lead { margin: 0; font-size: 12.5px; color: var(--ink-2); line-height: 1.45; }
.wd-refs-foot-note { margin-right: auto; align-self: center; font-size: 12px; color: var(--bad, #c0392b); }
.wd-refs-manage { display: flex; flex-direction: column; gap: 12px; padding-top: 2px; }
.wd-refs-toolbar { display: flex; align-items: center; gap: 12px; }
.wd-refs-search { flex: 1; display: flex; align-items: center; gap: 7px; border: 1px solid var(--line-strong); border-radius: var(--r-2); padding: 5px 9px; background: var(--surface); color: var(--ink-3); }
.wd-refs-search:focus-within { border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }
.wd-refs-search input { border: 0; outline: none; background: transparent; font: inherit; font-size: 12.5px; color: var(--ink); flex: 1; min-width: 0; }
.wd-refs-sort { display: inline-flex; align-items: center; gap: 7px; font-size: 12.5px; color: var(--ink-2); }
.wd-refs-lists { display: grid; grid-template-columns: 1fr 104px 1fr; gap: 12px; align-items: stretch; }
.wd-refs-listcol { display: flex; flex-direction: column; gap: 6px; min-width: 0; }
.wd-refs-listhead { font-size: 12px; font-weight: 650; color: var(--ink); display: flex; align-items: center; gap: 6px; }
.wd-refs-listhead span { font-weight: 500; font-size: 11px; color: var(--ink-3); background: var(--sunken, var(--surface-2)); border-radius: 8px; padding: 0 6px; }
.wd-refs-list { display: flex; flex-direction: column; border: 1px solid var(--line); border-radius: var(--r-2); height: 232px; overflow: auto; background: var(--surface); }
.wd-refs-row { display: flex; align-items: center; gap: 6px; text-align: left; font: inherit; font-size: 12.5px; border: 0; background: transparent; padding: 5px 8px; cursor: pointer; color: var(--ink); min-height: 28px; }
.wd-refs-row:hover { background: var(--hover); }
.wd-refs-row.on { background: var(--selected-strong); }
.wd-refs-row-text { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0; }
.wd-refs-mark { width: 14px; flex: none; display: inline-flex; justify-content: center; color: var(--accent); }
.wd-refs-q { font-weight: 700; color: #b8860b; }
.wd-refs-empty { padding: 16px 14px; font-size: 12px; color: var(--ink-3); line-height: 1.45; }
.wd-refs-verbs { display: flex; flex-direction: column; gap: 7px; justify-content: center; }
.wd-refs-verbs .rw-btn { justify-content: center; border: 1px solid var(--line); background: var(--surface); min-height: 30px; }
.wd-refs-verbs .rw-btn:hover:not(:disabled) { background: var(--hover); border-color: var(--line-strong); }
.wd-refs-legend { display: flex; gap: 16px; font-size: 11.5px; color: var(--ink-3); }
.wd-refs-legend span { display: inline-flex; align-items: center; gap: 4px; }
.wd-refs-legend b { color: #b8860b; }
`;

let installed = false;
/** The References look, once per window. */
export function installReferencesStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-word-references-css';
  style.textContent = REFERENCES_CSS + INDEX_CSS + FIGURES_CSS;
  document.head.appendChild(style);
}
