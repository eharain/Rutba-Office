// Pictures: the stars and tags the file itself carries, beside Pictures' own.
//
// Windows Photos, Explorer's Properties and Lightroom write a rating and
// keywords into the picture (EXIF Rating and XPKeywords, XMP xmp:Rating and
// dc:subject). Pictures keeps its own on this computer and never writes into
// the file (writing them back waits for 2.22), so the file's are shown here
// read only, under the picture's own, with Use to take them as Pictures' own
// where they differ.
import React from 'react';
import { t, tn } from '@rutba/office-ui';
import { adoptMarks } from './marks.js';

/**
 * The file's own rating and keywords, read only. Nothing when the file has
 * none. `own` is Pictures' marks for the file; `onUse(next)` takes the
 * file's as Pictures' own.
 */
export function FileMarks({ marks, own, onUse }) {
  if (!marks || (marks.rating == null && !marks.rejected && !marks.tags?.length)) return null;
  const rating = marks.rating || 0;
  const next = adoptMarks(own, marks);
  return (
    <div className="pv-file-marks" data-role="file-marks" data-rating={marks.rating ?? ''} data-tags={(marks.tags || []).join(';')}>
      <div className="pv-file-marks-head">
        <span className="pv-file-marks-title" data-tip={t('In the file: the rating and tags Windows Photos or Lightroom wrote into the picture. Pictures does not change them.')}>{t('In the file')}</span>
        {next.changed ? (
          <button type="button" className="pv-file-use" data-role="file-marks-use" data-tip={t('Use: the file\'s rating and tags taken as Pictures\' own')} onClick={() => onUse?.(next)}>{t('Use')}</button>
        ) : null}
      </div>
      {marks.rejected ? (
        <span className="pv-file-rejected">{t('Rejected')}</span>
      ) : marks.rating != null ? (
        <div className="pv-file-stars" role="img" aria-label={tn(rating, '{count} star', '{count} stars')}>
          {[1, 2, 3, 4, 5].map((n) => <span key={n} className={n <= rating ? 'on' : ''}>★</span>)}
        </div>
      ) : null}
      {marks.tags?.length ? (
        <div className="pv-file-tags">
          {marks.tags.map((name) => <span key={name} className="pv-file-tag" data-tag={name}>{name}</span>)}
        </div>
      ) : null}
    </div>
  );
}

export const FILE_MARKS_CSS = `
.pv-file-marks { display: flex; flex-direction: column; gap: 6px; padding: 8px 10px; border-radius: 6px; background: var(--sunken); border: 1px solid var(--line-soft); }
.pv-file-marks-head { display: flex; align-items: center; gap: 6px; }
.pv-file-marks-title { flex: 1; font-size: 11px; color: var(--ink-3); text-transform: none; letter-spacing: .01em; }
.pv-file-use { font: inherit; font-size: 11px; padding: 1px 8px; border-radius: 10px; border: 1px solid var(--line); background: var(--surface, #fff); color: var(--ink-2); cursor: pointer; }
.pv-file-use:hover { border-color: var(--accent); color: var(--accent-text); }
.pv-file-stars { display: flex; gap: 2px; font-size: 13px; line-height: 1; color: var(--ink-3); opacity: .55; }
.pv-file-stars .on { color: #e3a008; opacity: 1; }
.pv-file-stars:has(.on) { opacity: 1; }
.pv-file-stars span:not(.on) { opacity: .45; }
.pv-file-rejected { font-size: 11.5px; color: var(--bad, #c0392b); }
.pv-file-tags { display: flex; flex-wrap: wrap; gap: 4px; }
.pv-file-tag { font-size: 11px; padding: 1px 8px; border-radius: 10px; border: 1px dashed var(--line-strong, var(--line)); color: var(--ink-2); }
`;
