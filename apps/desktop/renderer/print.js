/**
 * The print dialog, for every app that has something to print.
 *
 * One dialog rather than three, because the questions are the same ones —
 * which printer, how many copies, what paper, which way round — and only the
 * middle differs: a workbook asks which sheets and whether to fit them across
 * the page, a deck asks whether it is printing slides, notes or a handout, a
 * document asks nothing more.
 *
 * Each starts from what its file keeps. A workbook's page setup and a deck's
 * print choices are kept in the file when printed, as Excel and PowerPoint
 * keep them; a document's paper, orientation and margins are the document's
 * own, so a change here changes the document, as Word's Print does.
 *
 * It says how many pages before it prints anything. That number is the whole
 * point of a print dialog: it is the difference between a report and forty
 * sheets of paper with three columns on each.
 */

import React, { useEffect, useMemo, useState } from 'react';
import { Dialog, Field, Input, Select, Button, Chip, useToast, t, tn, msg } from '@rutba/office-ui';

/** The margin presets' names, for the catalogue: the keys are the file's, the words shown are these. */
const MARGIN_NAMES = { Normal: msg('Normal'), Narrow: msg('Narrow'), Wide: msg('Wide') };

const PAPERS = ['A4', 'Letter', 'Legal', 'A3', 'A5'];
export const MARGIN_PRESETS = {
  Normal: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 },
  Narrow: { top: 6.4, right: 6.4, bottom: 6.4, left: 6.4 },
  Wide: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
};

/** Word's margins, in millimetres: a document's own presets. */
export const DOC_MARGIN_PRESETS = {
  Normal: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
  Narrow: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 },
  Wide: { top: 25.4, right: 50.8, bottom: 25.4, left: 50.8 },
};

const presetOf = (presets, m) => Object.entries(presets).find(([, p]) => ['top', 'right', 'bottom', 'left'].every((k) => Math.abs(p[k] - (m?.[k] ?? -99)) < 0.3))?.[0] ?? null;

export function defaultPrintOptions(kind) {
  if (kind === 'deck') return { paper: 'A4', orientation: 'landscape', layout: 'slides', perPage: 6, frame: true, margins: MARGIN_PRESETS.Normal };
  if (kind === 'sheet') {
    return {
      paper: 'A4',
      orientation: 'portrait',
      margins: MARGIN_PRESETS.Normal,
      fit: 'none',
      gridlines: false,
      headings: false,
      repeatRows: 0,
      sheets: null,
      area: null,
      footer: '&P of &N',
    };
  }
  return { paper: 'A4', orientation: 'portrait', margins: DOC_MARGIN_PRESETS.Normal };
}

export function PrintDialog({ shell, doc, kind, sheets = [], onClose, onSaveAs, onApply = null }) {
  const toast = useToast();
  const [options, setOptions] = useState(() => defaultPrintOptions(kind));
  const [margins, setMargins] = useState('Normal');
  const [printers, setPrinters] = useState([]);
  const [printer, setPrinter] = useState('');
  const [copies, setCopies] = useState(1);
  const [summary, setSummary] = useState(null);
  const [busy, setBusy] = useState(false);
  const set = (patch) => setOptions((o) => ({ ...o, ...patch }));
  // Through the window's own apply when it gives one, so what it shows follows.
  const applyOps = (ops) => (onApply ? onApply(ops) : shell.doc.apply({ id: doc.id, ops }));
  const presets = kind === 'doc' ? DOC_MARGIN_PRESETS : MARGIN_PRESETS;

  // A workbook carries its own page setup — Excel keeps paper, orientation,
  // margins, scaling, the print area and the repeated rows in the file — so
  // the dialog starts from what the file says rather than from what this
  // build happens to default to.
  const [keep, setKeep] = useState(kind === 'sheet' || kind === 'deck');
  useEffect(() => {
    shell.doc
      .pageSetup({ id: doc.id })
      .then((fromFile) => {
        if (!fromFile) return;
        if (kind === 'deck') {
          // An outline is not printed here: its slides are.
          const layout = fromFile.layout === 'outline' ? 'slides' : fromFile.layout;
          setOptions((o) => ({ ...o, layout, perPage: fromFile.perPage, frame: fromFile.frame, orientation: layout === 'slides' ? 'landscape' : 'portrait' }));
          return;
        }
        setOptions((o) => ({ ...o, ...fromFile, paper: fromFile.paper || o.paper }));
        setMargins(presetOf(kind === 'doc' ? DOC_MARGIN_PRESETS : MARGIN_PRESETS, fromFile.margins) || (kind === 'doc' ? 'Custom' : 'Normal'));
      })
      .catch(() => {});
  }, [shell, doc.id, kind]);

  /** A document's page changed here is changed in the document, as Word's Print changes it. */
  const change = async (patch) => {
    if (kind !== 'doc') return set(patch);
    const spec = {};
    if (patch.paper) spec.size = patch.paper;
    if (patch.orientation) spec.orientation = patch.orientation;
    if (patch.margins) spec.margins = Object.fromEntries(Object.entries(patch.margins).map(([k, mm]) => [k, Math.round((mm / 25.4) * 1440)]));
    try {
      await applyOps([{ op: 'setPageSetup', spec }]);
      set(patch);
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
    return undefined;
  };

  useEffect(() => {
    shell.print
      .printers()
      .then((list) => {
        setPrinters(list);
        setPrinter(list.find((p) => p.default)?.name || list[0]?.name || '');
      })
      .catch(() => setPrinters([]));
  }, [shell]);

  // How many pages, recomputed whenever a choice changes. The engines answer
  // this from geometry alone, so it costs nothing to ask on every keystroke.
  useEffect(() => {
    let alive = true;
    shell.print
      .summary({ id: doc.id, options })
      .then((s) => alive && setSummary(s))
      .catch((err) => alive && setSummary({ error: err.message }));
    return () => {
      alive = false;
    };
  }, [shell, doc.id, options]);

  const pages = summary?.error ? null : summary?.pages ?? null;
  const detail = useMemo(() => {
    if (summary?.error) return summary.error;
    if (pages == null) return t('Working out the pages…');
    const scale = summary?.sheets?.[0]?.scale;
    return scale && scale < 0.999
      ? tn(pages, '{count} page, shrunk to {percent}%', '{count} pages, shrunk to {percent}%', { percent: Math.round(scale * 100) })
      : tn(pages, '{count} page', '{count} pages');
  }, [summary, pages]);

  /** Keep the choices in the workbook or the deck, where the next person will find them. */
  const remember = async () => {
    if ((kind !== 'sheet' && kind !== 'deck') || !keep) return;
    try {
      if (kind === 'deck') await applyOps([{ op: 'setPrintSettings', settings: { layout: options.layout, perPage: options.perPage, frame: options.frame } }]);
      else await applyOps([{ op: 'setPageSetup', setup: options }]);
    } catch {
      /* a setup that will not save must not stop the print */
    }
  };

  const run = async () => {
    setBusy(true);
    try {
      await remember();
      const answer = await shell.print.document({ id: doc.id, options, printer, copies: Number(copies) || 1 });
      if (answer?.ok) toast(printer ? tn(pages ?? 0, 'Sent {count} page to {printer}.', 'Sent {count} pages to {printer}.', { printer }) : tn(pages ?? 0, 'Sent {count} page to the printer.', 'Sent {count} pages to the printer.'), { tone: 'good' });
      else if (answer?.reason && answer.reason !== 'cancelled') toast(answer.reason, { tone: 'bad' });
      onClose();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog
      title={t('Print')}
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Save as PDF…')} icon="pdf" disabled={busy} onClick={async () => { await remember(); onClose(); onSaveAs?.(options); }} />
          <span style={{ flex: 1 }} />
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={busy ? t('Printing…') : t('Print')} icon="print" disabled={busy || !printers.length} onClick={run} />
        </>
      }
    >
      <Field label={t('Printer')}>
        <Select value={printer} onChange={(e) => setPrinter(e.target.value)} style={{ width: '100%' }}>
          {printers.length ? null : <option value="">{t('No printer is installed')}</option>}
          {printers.map((p) => (
            <option key={p.name} value={p.name}>
              {p.name}
              {p.default ? ` ${t('(default)')}` : ''}
            </option>
          ))}
        </Select>
      </Field>

      <div style={{ display: 'flex', gap: 12 }}>
        <Field label={t('Copies')}>
          <Input type="number" min="1" max="99" value={copies} onChange={(e) => setCopies(e.target.value)} style={{ width: 80 }} />
        </Field>
        <Field label={t('Paper')}>
          <Select value={options.paper} onChange={(e) => change({ paper: e.target.value })} style={{ width: 110 }}>
            {PAPERS.map((p) => (
              <option key={p} value={p}>{p}</option>
            ))}
          </Select>
        </Field>
        <Field label={t('Orientation')}>
          <Select value={options.orientation} onChange={(e) => change({ orientation: e.target.value })} style={{ width: 130 }}>
            <option value="portrait">{t('Portrait')}</option>
            <option value="landscape">{t('Landscape')}</option>
          </Select>
        </Field>
        <Field label={t('Margins')}>
          <Select
            value={margins}
            onChange={(e) => {
              if (!presets[e.target.value]) return;
              setMargins(e.target.value);
              change({ margins: presets[e.target.value] });
            }}
            style={{ width: 110 }}
          >
            {Object.keys(presets).map((m) => (
              <option key={m} value={m}>{t(MARGIN_NAMES[m] || m)}</option>
            ))}
            {margins === 'Custom' ? <option value="Custom">{t('Custom')}</option> : null}
          </Select>
        </Field>
      </div>

      {kind === 'sheet' ? (
        <>
          <div style={{ display: 'flex', gap: 12 }}>
            <Field label={t('Print')}>
              <Select value={options.sheets === 'all' ? 'all' : 'active'} onChange={(e) => set({ sheets: e.target.value === 'all' ? 'all' : null })} style={{ width: 170 }}>
                <option value="active">{t('This sheet')}</option>
                <option value="all">{t('Every sheet ({count})', { count: sheets.length })}</option>
              </Select>
            </Field>
            <Field label={t('Scaling')}>
              <Select value={options.fit} onChange={(e) => set({ fit: e.target.value })} style={{ width: 190 }}>
                <option value="none">{t('No scaling')}</option>
                <option value="width">{t('Fit all columns on one page')}</option>
                <option value="page">{t('Fit the sheet on one page')}</option>
              </Select>
            </Field>
            <Field label={t('Repeat rows')}>
              <Input type="number" min="0" max="10" value={options.repeatRows} onChange={(e) => set({ repeatRows: Number(e.target.value) || 0 })} style={{ width: 80 }} title={t('Rows from the top of the sheet, drawn again at the top of every page')} />
            </Field>
          </div>
          <Field label={t('Print area')} hint={t('Empty prints everything with anything in it.')}>
            <Input value={options.area || ''} placeholder="A1:H60" onChange={(e) => set({ area: e.target.value.trim() || null })} style={{ width: '100%' }} />
          </Field>
          <div style={{ display: 'flex', gap: 16, marginTop: 4, flexWrap: 'wrap' }}>
            <label><input type="checkbox" checked={options.gridlines} onChange={(e) => set({ gridlines: e.target.checked })} /> {t('Gridlines')}</label>
            <label><input type="checkbox" checked={options.headings} onChange={(e) => set({ headings: e.target.checked })} /> {t('Row and column headings')}</label>
            <label title={t('Excel keeps the page setup in the file; so does this, so the next person to open it gets the same pages')}>
              <input type="checkbox" checked={keep} onChange={(e) => setKeep(e.target.checked)} /> {t('Keep these settings in the workbook')}
            </label>
          </div>
        </>
      ) : null}

      {kind === 'deck' ? (
        <div style={{ display: 'flex', gap: 12 }}>
          <Field label={t('Print')}>
            <Select
              value={options.layout}
              onChange={(e) => set({ layout: e.target.value, orientation: e.target.value === 'slides' ? 'landscape' : 'portrait' })}
              style={{ width: 190 }}
            >
              <option value="slides">{t('Full page slides')}</option>
              <option value="notes">{t('Notes pages')}</option>
              <option value="handout">{t('Handout')}</option>
            </Select>
          </Field>
          {options.layout === 'handout' ? (
            <Field label={t('Slides to a page')}>
              <Select value={options.perPage} onChange={(e) => set({ perPage: Number(e.target.value) })} style={{ width: 100 }}>
                {[1, 2, 3, 4, 6, 9].map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </Select>
            </Field>
          ) : null}
          <Field label={t('Frame')}>
            <label style={{ display: 'block', paddingTop: 6 }}>
              <input type="checkbox" checked={options.frame} onChange={(e) => set({ frame: e.target.checked })} /> {t('Draw a border round each slide')}
            </label>
          </Field>
        </div>
      ) : null}
      {kind === 'deck' ? (
        <label title={t('PowerPoint keeps what a deck prints in the file; so does this, so the next print starts from it')}>
          <input type="checkbox" checked={keep} onChange={(e) => setKeep(e.target.checked)} /> {t('Keep these settings in the presentation')}
        </label>
      ) : null}
      {kind === 'doc' ? (
        <p className="rw-hint" style={{ margin: '6px 0 0' }}>{t('Paper, orientation and margins are the document’s own: changing them here changes the document, as in Word.')}</p>
      ) : null}

      <div style={{ marginTop: 10 }}>
        <Chip>{detail}</Chip>
      </div>
    </Dialog>
  );
}
