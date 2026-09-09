/**
 * The print dialog, for every app that has something to print.
 *
 * One dialog rather than three, because the questions are the same ones —
 * which printer, how many copies, what paper, which way round — and only the
 * middle differs: a workbook asks which sheets and whether to fit them across
 * the page, a deck asks whether it is printing slides, notes or a handout, a
 * document asks nothing yet.
 *
 * It says how many pages before it prints anything. That number is the whole
 * point of a print dialog: it is the difference between a report and forty
 * sheets of paper with three columns on each.
 */

import React, { useEffect, useMemo, useState } from 'react';
import { Dialog, Field, Input, Select, Button, Chip, useToast } from '@rutba/office-ui';

const PAPERS = ['A4', 'Letter', 'Legal', 'A3', 'A5'];
const MARGIN_PRESETS = {
  Normal: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 },
  Narrow: { top: 6.4, right: 6.4, bottom: 6.4, left: 6.4 },
  Wide: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
};

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
  return { paper: 'A4', orientation: 'portrait', margins: MARGIN_PRESETS.Normal };
}

export function PrintDialog({ shell, doc, kind, sheets = [], onClose, onSaveAs }) {
  const toast = useToast();
  const [options, setOptions] = useState(() => defaultPrintOptions(kind));
  const [margins, setMargins] = useState('Normal');
  const [printers, setPrinters] = useState([]);
  const [printer, setPrinter] = useState('');
  const [copies, setCopies] = useState(1);
  const [summary, setSummary] = useState(null);
  const [busy, setBusy] = useState(false);
  const set = (patch) => setOptions((o) => ({ ...o, ...patch }));

  // A workbook carries its own page setup — Excel keeps paper, orientation,
  // margins, scaling, the print area and the repeated rows in the file — so
  // the dialog starts from what the file says rather than from what this
  // build happens to default to.
  const [keep, setKeep] = useState(kind === 'sheet');
  useEffect(() => {
    if (kind !== 'sheet') return;
    shell.doc
      .pageSetup({ id: doc.id })
      .then((fromFile) => {
        if (!fromFile) return;
        setOptions((o) => ({ ...o, ...fromFile }));
        const named = Object.entries(MARGIN_PRESETS).find(([, m]) => Math.abs(m.top - fromFile.margins.top) < 0.2 && Math.abs(m.left - fromFile.margins.left) < 0.2);
        setMargins(named ? named[0] : 'Normal');
      })
      .catch(() => {});
  }, [shell, doc.id, kind]);

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
    if (pages == null) return 'Working out the pages…';
    const scale = summary?.sheets?.[0]?.scale;
    const shrunk = scale && scale < 0.999 ? `, shrunk to ${Math.round(scale * 100)}%` : '';
    return `${pages} page${pages === 1 ? '' : 's'}${shrunk}`;
  }, [summary, pages]);

  /** Keep the choices in the workbook, where the next person will find them. */
  const remember = async () => {
    if (kind !== 'sheet' || !keep) return;
    try {
      await shell.doc.apply({ id: doc.id, ops: [{ op: 'setPageSetup', setup: options }] });
    } catch {
      /* a setup that will not save must not stop the print */
    }
  };

  const run = async () => {
    setBusy(true);
    try {
      await remember();
      const answer = await shell.print.document({ id: doc.id, options, printer, copies: Number(copies) || 1 });
      if (answer?.ok) toast(`Sent ${pages ?? ''} page${pages === 1 ? '' : 's'} to ${printer || 'the printer'}.`, { tone: 'good' });
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
      title="Print"
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label="Save as PDF…" icon="pdf" disabled={busy} onClick={async () => { await remember(); onClose(); onSaveAs?.(options); }} />
          <span style={{ flex: 1 }} />
          <Button label="Cancel" onClick={onClose} />
          <Button primary label={busy ? 'Printing…' : 'Print'} icon="print" disabled={busy || !printers.length} onClick={run} />
        </>
      }
    >
      <Field label="Printer">
        <Select value={printer} onChange={(e) => setPrinter(e.target.value)} style={{ width: '100%' }}>
          {printers.length ? null : <option value="">No printer is installed</option>}
          {printers.map((p) => (
            <option key={p.name} value={p.name}>
              {p.name}
              {p.default ? ' (default)' : ''}
            </option>
          ))}
        </Select>
      </Field>

      <div style={{ display: 'flex', gap: 12 }}>
        <Field label="Copies">
          <Input type="number" min="1" max="99" value={copies} onChange={(e) => setCopies(e.target.value)} style={{ width: 80 }} />
        </Field>
        <Field label="Paper">
          <Select value={options.paper} onChange={(e) => set({ paper: e.target.value })} style={{ width: 110 }}>
            {PAPERS.map((p) => (
              <option key={p} value={p}>{p}</option>
            ))}
          </Select>
        </Field>
        <Field label="Orientation">
          <Select value={options.orientation} onChange={(e) => set({ orientation: e.target.value })} style={{ width: 130 }}>
            <option value="portrait">Portrait</option>
            <option value="landscape">Landscape</option>
          </Select>
        </Field>
        <Field label="Margins">
          <Select
            value={margins}
            onChange={(e) => {
              setMargins(e.target.value);
              set({ margins: MARGIN_PRESETS[e.target.value] });
            }}
            style={{ width: 110 }}
          >
            {Object.keys(MARGIN_PRESETS).map((m) => (
              <option key={m} value={m}>{m}</option>
            ))}
          </Select>
        </Field>
      </div>

      {kind === 'sheet' ? (
        <>
          <div style={{ display: 'flex', gap: 12 }}>
            <Field label="Print">
              <Select value={options.sheets === 'all' ? 'all' : 'active'} onChange={(e) => set({ sheets: e.target.value === 'all' ? 'all' : null })} style={{ width: 170 }}>
                <option value="active">This sheet</option>
                <option value="all">Every sheet ({sheets.length})</option>
              </Select>
            </Field>
            <Field label="Scaling">
              <Select value={options.fit} onChange={(e) => set({ fit: e.target.value })} style={{ width: 190 }}>
                <option value="none">No scaling</option>
                <option value="width">Fit all columns on one page</option>
                <option value="page">Fit the sheet on one page</option>
              </Select>
            </Field>
            <Field label="Repeat rows">
              <Input type="number" min="0" max="10" value={options.repeatRows} onChange={(e) => set({ repeatRows: Number(e.target.value) || 0 })} style={{ width: 80 }} title="Rows from the top of the sheet, drawn again at the top of every page" />
            </Field>
          </div>
          <Field label="Print area" hint="Empty prints everything with anything in it.">
            <Input value={options.area || ''} placeholder="A1:H60" onChange={(e) => set({ area: e.target.value.trim() || null })} style={{ width: '100%' }} />
          </Field>
          <div style={{ display: 'flex', gap: 16, marginTop: 4, flexWrap: 'wrap' }}>
            <label><input type="checkbox" checked={options.gridlines} onChange={(e) => set({ gridlines: e.target.checked })} /> Gridlines</label>
            <label><input type="checkbox" checked={options.headings} onChange={(e) => set({ headings: e.target.checked })} /> Row and column headings</label>
            <label title="Excel keeps the page setup in the file; so does this, so the next person to open it gets the same pages">
              <input type="checkbox" checked={keep} onChange={(e) => setKeep(e.target.checked)} /> Keep these settings in the workbook
            </label>
          </div>
        </>
      ) : null}

      {kind === 'deck' ? (
        <div style={{ display: 'flex', gap: 12 }}>
          <Field label="Print">
            <Select
              value={options.layout}
              onChange={(e) => set({ layout: e.target.value, orientation: e.target.value === 'slides' ? 'landscape' : 'portrait' })}
              style={{ width: 190 }}
            >
              <option value="slides">Full page slides</option>
              <option value="notes">Notes pages</option>
              <option value="handout">Handout</option>
            </Select>
          </Field>
          {options.layout === 'handout' ? (
            <Field label="Slides to a page">
              <Select value={options.perPage} onChange={(e) => set({ perPage: Number(e.target.value) })} style={{ width: 100 }}>
                {[1, 2, 3, 4, 6, 9].map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </Select>
            </Field>
          ) : null}
          <Field label="Frame">
            <label style={{ display: 'block', paddingTop: 6 }}>
              <input type="checkbox" checked={options.frame} onChange={(e) => set({ frame: e.target.checked })} /> Draw a border round each slide
            </label>
          </Field>
        </div>
      ) : null}

      <div style={{ marginTop: 10 }}>
        <Chip>{detail}</Chip>
      </div>
    </Dialog>
  );
}
