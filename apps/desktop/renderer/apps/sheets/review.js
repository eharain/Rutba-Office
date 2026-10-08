// Worksheets' side of Review → Check Accessibility and Spelling: a finding
// or a misspelt word is a cell (on its sheet), and the fixes are the grid's
// own verbs — Unmerge, a font colour, Rename. The pass reads the active
// sheet from the active cell, and offers the other sheets when it is done,
// as Excel asks. The panes, the pass and the dialogs are renderer/review.js.

import { useMemo, useRef } from 'react';
import { useReview, wordAround } from '../../review.js';

export function useSheetsReview({ shell, doc, model, dispatch, toast }) {
  const modelRef = useRef(model);
  modelRef.current = model;

  const at = (where) => {
    const ops = [];
    if (where?.sheet && where.sheet !== modelRef.current?.activeSheet) ops.push({ op: 'sheet', name: where.sheet });
    if (where?.row != null) ops.push({ op: 'select', row: where.row, col: where.col ?? 0 });
    return ops;
  };

  const adapter = useMemo(() => ({
    async goTo(where) {
      const ops = at(where);
      if (ops.length) await dispatch(...ops);
    },
    async altText(target, props) {
      await dispatch({ op: 'setAltText', sheet: target.sheet, anchor: target.anchor, descr: props.descr, decorative: props.decorative });
    },
    async fix(issue, f) {
      switch (f.kind) {
        case 'unmerge':
          return dispatch(...at(issue.where), { op: 'unmerge' });
        case 'tableHeader':
          return dispatch({ op: 'setHeaderRow', part: f.target.part });
        case 'textColour':
          return dispatch(...at(issue.where), { op: 'setFormat', delta: { fontColour: f.colour } });
        default:
          return null;
      }
    },
    renameSheet: (from, to) => dispatch({ op: 'renameSheet', from, to }),
    // Done with the active sheet: offer the others, once.
    morePass: (state) => {
      const sheets = modelRef.current?.sheets || [];
      const active = modelRef.current?.activeSheet;
      if (state?.extra?.sheets || sheets.length < 2) return null;
      const others = sheets.filter((s) => s !== active);
      return { message: `The spelling check is complete for ${active}.`, label: `Check the other ${others.length === 1 ? 'sheet' : `${others.length} sheets`}`, extra: { sheets: others }, others };
    },
    async showWord(found) {
      const ops = at(found.where);
      if (ops.length) await dispatch(...ops);
    },
    whereLabel: (found) => (found.where ? `${found.where.sheet}!${found.where.ref}` : ''),
    // The active cell's words, as Excel's thesaurus reads them; a formula's answer is not a word to replace.
    thesaurus: {
      word() {
        const m = modelRef.current;
        const cell = (m?.cells || []).find((c) => c.active);
        if (!cell || cell.isFormula || typeof cell.text !== 'string') return null;
        const text = cell.text;
        const found = /\s/.test(text.trim()) ? wordAround(text, 0) : wordAround(text, 0, text.length);
        return found ? { ...found, row: cell.row, col: cell.col, text } : null;
      },
      async replace(found, word) {
        const text = found.text.slice(0, found.start) + word + found.text.slice(found.end);
        await dispatch({ op: 'setCell', row: found.row, col: found.col, value: text });
      },
    },
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [dispatch]);

  return useReview({ shell, doc, model, apply: dispatch, toast, adapter });
}
