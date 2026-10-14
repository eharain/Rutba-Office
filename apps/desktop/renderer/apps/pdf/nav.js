// Finding one's way around a PDF: its bookmarks in a pane beside the pages,
// as Acrobat's Bookmarks panel lists them, each going to its page where the
// page now stands in the window (a bookmark whose page was deleted is shown
// greyed); and the label each page carries ("iv", "A-3"), which follows the
// page when pages are moved. Saving, extracting, splitting and merging carry
// bookmarks, labels and links into the new file (@rutba/pdf/navigation).

import React, { useState } from 'react';
import { Icon, t } from '@rutba/office-ui';

/** The page of the window a bookmark of `source` names: the first entry showing that page. */
export const entryFor = (pages, source, page) => (page === null || page === undefined ? null : pages.find((p) => p.source === source && p.page === page) || null);

/** A page entry's label, when its file labels its pages; null otherwise. */
export const labelOf = (sources, entry) => sources[entry?.source]?.nav?.labels?.[entry.page] ?? null;

/**
 * The bookmarks of every file the pages come from, in the order the files
 * first appear. `onGo(key)` goes to a page; `onLink(url)` opens a web
 * address a bookmark holds.
 */
export function BookmarksPane({ pages, sources, current, onGo, onLink }) {
  const seen = new Set();
  const trees = [];
  for (const p of pages) {
    if (seen.has(p.source)) continue;
    seen.add(p.source);
    const outline = sources[p.source]?.nav?.outline || [];
    if (outline.length) trees.push({ source: p.source, name: sources[p.source]?.name || '', outline });
  }
  if (!trees.length) {
    return <div className="pdf-bookmarks-empty" data-role="bookmarks-empty">{t('This PDF has no bookmarks.')}</div>;
  }
  return (
    <div className="pdf-bookmarks" role="tree" aria-label={t('Bookmarks')} data-role="bookmarks">
      {trees.map((tree) => (
        <div key={tree.source} className="pdf-bookmarks-file">
          {trees.length > 1 ? <div className="pdf-bookmarks-name">{tree.name}</div> : null}
          {tree.outline.map((node, i) => <Bookmark key={i} node={node} source={tree.source} pages={pages} current={current} onGo={onGo} onLink={onLink} depth={0} />)}
        </div>
      ))}
    </div>
  );
}

function Bookmark({ node, source, pages, current, onGo, onLink, depth }) {
  const [open, setOpen] = useState(node.open || depth === 0);
  const entry = entryFor(pages, source, node.page);
  const gone = node.page !== null && !entry;
  const here = entry && entry.key === current;
  const go = () => {
    if (entry) onGo(entry.key);
    else if (node.url) onLink(node.url);
  };
  return (
    <div className="pdf-bookmark-node" role="treeitem" aria-expanded={node.kids.length ? open : undefined}>
      <div
        className={`pdf-bookmark${here ? ' current' : ''}${gone ? ' gone' : ''}`}
        style={{ paddingInlineStart: 6 + depth * 16 }}
        data-title={node.title}
        data-page={entry ? pages.indexOf(entry) + 1 : ''}
        title={gone ? t('The page this bookmark goes to is not in the document now') : node.url || undefined}
        onClick={go}
      >
        {node.kids.length ? (
          <button type="button" className="pdf-bookmark-twist" aria-label={open ? t('Collapse') : t('Expand')} onClick={(e) => { e.stopPropagation(); setOpen((v) => !v); }}>
            <Icon name={open ? 'chevronDown' : 'chevronRight'} size={12} strokeWidth={2} />
          </button>
        ) : <span className="pdf-bookmark-twist" />}
        <Icon name={node.url ? 'link' : 'file'} size={13} />
        <span className="pdf-bookmark-title" style={{ fontWeight: node.bold ? 600 : undefined, fontStyle: node.italic ? 'italic' : undefined, color: node.color && !gone ? `rgb(${node.color.map((v) => Math.round(Math.max(0, Math.min(1, v)) * 255)).join(' ')})` : undefined }}>{node.title || t('Untitled')}</span>
      </div>
      {open && node.kids.length ? node.kids.map((k, i) => <Bookmark key={i} node={k} source={source} pages={pages} current={current} onGo={onGo} onLink={onLink} depth={depth + 1} />) : null}
    </div>
  );
}
