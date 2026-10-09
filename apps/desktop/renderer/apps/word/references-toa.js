// References → Table of Authorities: Mark Citation (Alt+Shift+I), Insert
// Table of Authorities and Update Table, as Word 365 lays them out. Mark
// Citation stays open beside the page, as Word's does, so one citation
// after another can be marked — Next Citation finds the next words that
// look like one; Insert Table of Authorities previews the table it will
// write. The table's words are `@rutba/ooxml/wordtoa`'s.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Group, Input, Select, Icon, t, tn } from '@rutba/office-ui';
import { TOA_CATEGORIES, buildAuthorities, authorityTail, categoryName } from '@rutba/ooxml/wordtoa';
import { LEADERS } from './references-index.js';

const LEADER_CSS = { dot: 'dotted', hyphen: 'dashed', underscore: 'solid' };

/** References → Table of Authorities, as Word lays the group out. */
export function ToaGroup({ act, hasTable }) {
  return (
    <Group label={t('Table of Authorities')}>
      <Button tall icon="flag" label={t('Mark\nCitation')} className="wd-refs-2line" title={t('Mark Citation (Alt+Shift+I) — add the selected words to the table of authorities')} onClick={() => act('markCitation')} />
      <div className="wd-refs-col">
        <Button icon="listBullet" label={t('Insert Table of Authorities')} title={t('Insert Table of Authorities — the cases, statutes and other authorities cited, and the pages they are cited on')} onClick={() => act('insertToa')} />
        <Button icon="refresh" label={t('Update Table')} className="wd-toa-update" disabled={!hasTable} title={t('Update Table — the citations and their pages as they stand now')} onClick={() => act('updateToa')} />
      </div>
    </Group>
  );
}

/**
 * Mark Citation — beside the page, not over it: the long citation starts
 * as the selected words, and follows a new selection until it is typed in;
 * the short citation follows the long one the same way. Choosing a long
 * citation from the list marks another place it is cited.
 */
export function MarkCitationPanel({ selected, authorities, onMark, onNext, onClose }) {
  const [long, setLong] = useState(selected || '');
  const [short, setShort] = useState(selected || '');
  const [category, setCategory] = useState(1);
  const [note, setNote] = useState('');
  const typedLong = useRef(false);
  const typedShort = useRef(false);

  useEffect(() => {
    // A new selection is a new citation.
    if (selected && !typedLong.current) setLong(selected);
    if (selected && !typedShort.current) setShort(selected);
  }, [selected]);
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  const inCategory = authorities.filter((a) => a.category === Number(category));
  const mark = async (all) => {
    if (!long.trim()) return;
    const n = await onMark({ long: long.trim(), short: short.trim() || long.trim(), category: Number(category), all, text: all ? selected : null });
    setNote(all ? tn(n || 0, '{count} citation marked.', '{count} citations marked.') : t('Marked. Select another citation, or press Next Citation.'));
    typedLong.current = false;
    typedShort.current = false;
  };
  const choose = (a) => {
    setLong(a.long);
    setShort(a.short);
    setCategory(a.category);
    typedLong.current = true;
    typedShort.current = true;
  };

  return (
    <div className="wd-refs-float wd-toa-float" role="dialog" aria-label={t('Mark Citation')}>
      <div className="wd-refs-float-head">
        <span>{t('Mark Citation')}</span>
        <button type="button" className="wd-refs-x" data-tip={t('Close')} aria-label={t('Close')} onClick={onClose}><Icon name="close" size={14} /></button>
      </div>
      <div className="wd-refs-float-body">
        <label className="wd-toa-label" htmlFor="wd-toa-long">{t('Selected text')}</label>
        <textarea id="wd-toa-long" className="rw-input wd-toa-long" rows={3} value={long} onChange={(e) => { typedLong.current = true; setLong(e.target.value); if (!typedShort.current) setShort(e.target.value); }} autoFocus />
        <div className="wd-refs-options">
          <label htmlFor="wd-toa-category">{t('Category')}</label>
          <Select id="wd-toa-category" className="rw-select wd-toa-category" value={String(category)} onChange={(e) => setCategory(Number(e.target.value))}>
            {TOA_CATEGORIES.map((name, i) => <option key={i} value={String(i + 1)}>{name}</option>)}
          </Select>
          <label htmlFor="wd-toa-short">{t('Short citation')}</label>
          <Input id="wd-toa-short" type="text" className="rw-input wd-toa-short" value={short} onChange={(e) => { typedShort.current = true; setShort(e.target.value); }}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); mark(false); } }} />
        </div>
        <div className="wd-toa-label">{t('Long citations in {category}', { category: categoryName(category) })}</div>
        <div className="wd-toa-list" role="listbox" aria-label={t('Long citations')}>
          {inCategory.length ? inCategory.map((a) => (
            <button key={a.long} type="button" role="option" aria-selected={a.long === long} className={`wd-toa-item${a.long === long ? ' on' : ''}`} onClick={() => choose(a)}>{a.long}</button>
          )) : <div className="wd-toa-empty">{t('None marked yet.')}</div>}
        </div>
        <p className="wd-refs-lead wd-refs-note">{note || t('This box stays open so that you can mark one citation after another.')}</p>
      </div>
      <div className="wd-refs-float-foot">
        <Button className="wd-toa-next" label={t('Next Citation')} title={t('Next Citation — find the next words that look like a citation: v., In re, Id., supra, §, a reporter')} onClick={async () => { const found = await onNext(); setNote(found ? '' : t('No more citations found.')); }} />
        <Button primary className="wd-toa-mark" label={t('Mark')} disabled={!long.trim()} onClick={() => mark(false)} />
        <Button className="wd-toa-markall" label={t('Mark All')} disabled={!long.trim() || !selected} title={t('Mark All — every other place this long or short citation appears, matching case')} onClick={() => mark(true)} />
        <Button label={t('Close')} onClick={onClose} />
      </div>
    </div>
  );
}

/** Word's own print-preview authorities, so the preview shows a page list and a passim. */
const SAMPLE = [
  { long: 'Baker v. Carr, 369 U.S. 186 (1962)', category: 1, page: 1 },
  { long: 'Baker v. Carr, 369 U.S. 186 (1962)', category: 1, page: 4 },
  ...[1, 2, 3, 5, 7].map((page) => ({ long: 'Marbury v. Madison, 5 U.S. 137 (1803)', category: 1, page })),
  { long: 'Clean Air Act, 42 U.S.C. § 7401 (1970)', category: 2, page: 6 },
];

/** Insert Table of Authorities: category, passim, the entries' formatting, the leader — and a preview of all of it. */
export function ToaDialog({ authorities, current, onClose, onOk, onMark }) {
  const [category, setCategory] = useState(current?.length === 1 ? String(current[0].category) : 'all');
  const [passim, setPassim] = useState(current?.[0] ? current[0].passim : true);
  const [keepFormatting, setKeepFormatting] = useState(current?.[0] ? current[0].keepFormatting : true);
  const [leader, setLeader] = useState(current?.[0]?.leader && current[0].leader !== 'none' ? current[0].leader : 'dot');
  const preview = useMemo(() => {
    const entries = authorities.length ? authorities.map((a, i) => ({ ...a, page: i + 1 })) : SAMPLE;
    const list = buildAuthorities(entries, { usePassim: passim });
    const cats = category === 'all' ? [...new Set(list.map((a) => a.category))].sort((a, b) => a - b) : [Number(category)];
    return cats.map((c) => ({ category: c, list: list.filter((a) => a.category === c) }));
  }, [authorities, passim, category]);

  return (
    <Dialog
      title={t('Table of Authorities')}
      width={640}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Mark Citation…')} className="wd-toa-open-mark" onClick={onMark} />
          <span style={{ flex: 1 }} />
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary className="wd-toa-ok" label={t('OK')} onClick={() => onOk({ category: category === 'all' ? 'all' : Number(category), passim, keepFormatting, leader })} />
        </>
      }
    >
      <div className="wd-refs-index">
        <div>
          <div className="wd-refs-section">{t('Print Preview')}</div>
          <div className="wd-idx-preview wd-toa-preview">
            {preview.map((g) => (
              <React.Fragment key={g.category}>
                <div className="wd-idx-letter">{categoryName(g.category)}</div>
                {g.list.length ? g.list.map((a) => (
                  <div key={a.long} className="wd-idx-line right wd-toa-line">
                    <span>{a.long}</span>
                    <span className="wd-idx-leader" style={{ borderBottomStyle: LEADER_CSS[leader] || 'none', borderBottomColor: LEADER_CSS[leader] ? 'currentColor' : 'transparent' }} />
                    <span className="wd-idx-pages">{authorityTail(a).slice(1).map((s) => s.text).join('')}</span>
                  </div>
                )) : <div className="wd-idx-line"><b>{t('No table of authorities entries found.')}</b></div>}
              </React.Fragment>
            ))}
          </div>
          <label className="wd-refs-check"><input type="checkbox" className="wd-toa-passim" checked={passim} onChange={(e) => setPassim(e.target.checked)} />{t('Use passim')}</label>
          <label className="wd-refs-check"><input type="checkbox" className="wd-toa-keep" checked={keepFormatting} onChange={(e) => setKeepFormatting(e.target.checked)} />{t('Keep original formatting')}</label>
          <label className="wd-refs-leader">
            {t('Tab leader')}
            <Select className="rw-select wd-toa-leader" value={leader} onChange={(e) => setLeader(e.target.value)}>
              {LEADERS.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
            </Select>
          </label>
        </div>
        <div className="wd-refs-index-side">
          <div className="wd-refs-section">{t('Category')}</div>
          <div className="wd-toa-list wd-toa-cats" role="listbox" aria-label={t('Category')}>
            {[['all', t('All')], ...TOA_CATEGORIES.map((name, i) => [String(i + 1), name])].map(([id, name]) => (
              <button key={id} type="button" role="option" aria-selected={category === id} data-category={id} className={`wd-toa-item${category === id ? ' on' : ''}`} onClick={() => setCategory(id)}>{name}</button>
            ))}
          </div>
        </div>
      </div>
      <p className="wd-refs-lead wd-refs-foot-lead">{t("Each category is written as Word's TOA field, its heading over it; All writes one for every category that has citations in it.")}</p>
    </Dialog>
  );
}

export const TOA_CSS = `
.wd-toa-float { top: 150px; }
.wd-toa-label { font-size: 12.5px; color: var(--ink-2); }
.wd-toa-long { width: 100%; box-sizing: border-box; resize: vertical; font: inherit; min-height: 56px; }
.wd-toa-list { border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); height: 96px; overflow: auto; }
.wd-toa-cats { height: 268px; }
.wd-toa-item { display: block; width: 100%; border: 0; border-bottom: 1px solid var(--line-soft); background: none; font: inherit; font-size: 12.5px; color: var(--ink); text-align: left; padding: 4px 8px; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.wd-toa-item:hover { background: var(--hover); }
.wd-toa-item.on { background: var(--selected); color: var(--accent); }
.wd-toa-empty { padding: 8px; color: var(--ink-3); font-size: 12px; }
.wd-toa-preview { height: 176px; margin-bottom: 8px; }
.wd-toa-line > span:first-child { white-space: normal; }
`;
