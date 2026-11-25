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
 *
 * An app with no engine behind it (Mail, Calendar, Contacts, Image and
 * Pictures) gives the dialog a `job` in place of a document: the name of
 * what prints, and `html(options)`, the page as the app lays it out for the
 * paper chosen (print-page.js). The dialog counts its pages by having it laid
 * out, prints it by the same hidden window, and saves it as a PDF itself. A
 * job's own choices (which contacts, how many pictures to a sheet) are drawn
 * where a workbook's and a deck's are.
 *
 * Which way round the paper goes is Automatic until a person chooses: a sheet
 * too wide for upright paper, a slide, a picture wider than it is tall and a
 * month of a calendar go on paper turned sideways, and the rest stay upright.
 * A workbook whose file says which way is printed that way, and a document's
 * pages are the document's own. A job says its own way by `orientation(options)`.
 *
 * Pages takes the page numbers to print, as they are typed everywhere:
 * "1,2,5,6", "1-5", "2-4, 9". Empty prints every page.
 *
 * This dialog is the only one. Print sends the job straight to the printer
 * chosen here, on the paper chosen here: the system's dialog is not shown
 * after it to ask the same things again. What only a printer's own driver
 * knows (both sides of the paper, a tray) is one button away, "System
 * dialog…", which opens the system's dialog in place of printing at once.
 *
 * At the top it names what is about to print, with page one drawn small and
 * the number of sheets. The printer it opens on is the system's default, and
 * a printer that is offline, paused or out of paper says so in the list.
 *
 * The paper it opens on is the one this person last chose in this app, and
 * the region's paper (Letter or A4) until they have chosen one: each app
 * remembers its own. A workbook whose file names a paper opens on that, and
 * a document's paper is the document's. Until 2.28.0 it opened on A4 every
 * time.
 *
 * The pages are counted, and page one drawn, by having the print laid out.
 * One layout is in hand at a time: a choice changed while one runs waits for
 * it, and only the last change made meanwhile is laid out after it.
 */

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Dialog, Field, Input, Select, Button, Chip, useToast, basename, t, tn, msg } from '@rutba/office-ui';
import { printFileStem, pagesTyped, wayChosen, wayOfShape, printerToOpenOn, printOutcome, PAPERS, paperKey, startingPaper, oneAtATime } from './print-page.js';
import { openDocument, drawPage } from './pdfjs.js';

/** What a printer that is not ready says after its name in the list: the keys are the print service's states. */
const STATE_WORDS = { offline: msg('print.offline'), paused: msg('print.paused'), paper: msg('print.out-of-paper'), trouble: msg('print.needs-attention') };

/** The room the small picture of page one is drawn in, in CSS pixels: an upright sheet fills its height, one turned sideways its width. */
const FIRST_PAGE = { width: 104, height: 112 };

/** A workbook or a deck of more sheets than this is not laid out a second time for the small picture. */
const PREVIEW_UP_TO = 200;

/**
 * Page one of what is about to print, drawn small from the PDF the print service made of it: the
 * sheet as the printer gets it, turned as it will be. Drawn by the suite's own pdf.js. `fresh` is
 * false while the choices have changed and the new page has not arrived: the old one is shown dimmed.
 */
function FirstPage({ bytes, fresh }) {
  const canvas = useRef(null);
  const [shape, setShape] = useState(null);
  useEffect(() => {
    if (!bytes) return undefined;
    let gone = false;
    let draw = null;
    let pdf = null;
    (async () => {
      try {
        pdf = await openDocument(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
        const page = await pdf.getPage(1);
        const whole = page.getViewport({ scale: 1 });
        const scale = Math.min(FIRST_PAGE.width / whole.width, FIRST_PAGE.height / whole.height);
        if (gone || !canvas.current) return;
        draw = drawPage(pdf, 0, canvas.current, { scale, priority: 0 });
        await draw.done;
        // The frame takes the page's shape with the page itself, so the last one is never shown stretched.
        if (!gone) setShape({ width: Math.round(whole.width * scale), height: Math.round(whole.height * scale) });
      } catch {
        /* a page that will not draw leaves the frame empty: the print is not held up by its picture */
      }
    })();
    return () => {
      gone = true;
      draw?.cancel();
      Promise.resolve(pdf?.destroy?.()).catch(() => {});
    };
  }, [bytes]);
  return (
    <div className="rw-print-first" style={{ flex: 'none', width: FIRST_PAGE.width, height: FIRST_PAGE.height, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <canvas
        ref={canvas}
        role="img"
        aria-label={t('print.first-page-prints')}
        data-drawn={shape ? 'yes' : 'no'}
        style={{
          width: shape?.width ?? Math.round(FIRST_PAGE.height * 0.707),
          height: shape?.height ?? FIRST_PAGE.height,
          background: '#fff',
          border: '1px solid var(--line-strong)',
          boxShadow: '0 1px 3px rgba(0, 0, 0, 0.18)',
          opacity: fresh ? 1 : 0.45,
          transition: 'opacity 120ms',
        }}
      />
    </div>
  );
}

/** The margin presets' names, for the catalogue: the keys are the file's, the words shown are these. */
const MARGIN_NAMES = { Normal: msg('common.normal'), Narrow: msg('common.narrow'), Wide: msg('common.wide') };

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

/** The language and direction of the window, for a page an app lays out to read as the window does. */
export const pageWords = () => ({ lang: document.documentElement.lang || 'en', dir: document.documentElement.dir === 'rtl' ? 'rtl' : 'ltr' });

export { printFileStem, wayOfShape };

/**
 * What a print starts from. The way round is 'auto' wherever the thing printed has a shape to go by:
 * the engine (a workbook, a deck) or the job (a page an app laid out) then says which way. A document's
 * is its own.
 */
export function defaultPrintOptions(kind) {
  if (kind === 'page') return { paper: 'A4', orientation: 'auto', margins: MARGIN_PRESETS.Normal };
  if (kind === 'deck') return { paper: 'A4', orientation: 'auto', layout: 'slides', perPage: 6, frame: true, margins: MARGIN_PRESETS.Normal };
  if (kind === 'sheet') {
    return {
      paper: 'A4',
      orientation: 'auto',
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

/** The app this window is, as the window's theme wrote it on the page ('sheets', 'mail'): '' where nothing says. */
const appOfWindow = () => (typeof document !== 'undefined' && document.documentElement?.dataset?.app) || '';

/**
 * The paper a print in this app starts on: the one last chosen here, else the region's (startingPaper).
 * For a print with no dialog too (File, Export as PDF), so both go on the same paper.
 */
export async function paperToStartOn(shell, app = appOfWindow()) {
  const [remembered, about] = await Promise.all([
    Promise.resolve().then(() => (app ? shell.store.get({ key: paperKey(app), fallback: null }) : null)).catch(() => null),
    Promise.resolve().then(() => shell.app.version()).catch(() => null),
  ]);
  return startingPaper({ remembered, region: about?.region });
}

export function PrintDialog({ shell, doc = null, kind: given, sheets = [], job = null, app: givenApp = null, onClose, onSaveAs, onApply = null }) {
  const toast = useToast();
  const kind = job ? 'page' : given;
  const app = givenApp || appOfWindow();
  // False until the paper to start on, and what the file keeps, are known: nothing is laid out for a
  // paper that is about to change.
  const [started, setStarted] = useState(false);
  const lanes = useRef(null);
  lanes.current ||= { count: oneAtATime(), picture: oneAtATime() };
  const [options, setOptions] = useState(() => ({ ...defaultPrintOptions(kind), ...(job?.defaults || {}) }));
  const [margins, setMargins] = useState('Normal');
  // Null until the system has said which printers there are: an empty list is "no printer", and not yet is not.
  const [printers, setPrinters] = useState(null);
  const [printer, setPrinter] = useState('');
  const chose = useRef(false);
  // Page one of the print as a PDF: undefined until one is asked for, null when there is none to show.
  const [first, setFirst] = useState(undefined);
  const [fresh, setFresh] = useState(false);
  const [copies, setCopies] = useState(1);
  // The pages to print, as typed: empty is every page.
  const [range, setRange] = useState('');
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
    let alive = true;
    (async () => {
      // The paper to start on, unless it is a document's own or the job names one; and what the file keeps.
      const [paper, fromFile] = await Promise.all([
        kind === 'doc' || job?.defaults?.paper ? null : paperToStartOn(shell, app).catch(() => null),
        job ? null : Promise.resolve().then(() => shell.doc.pageSetup({ id: doc.id })).catch(() => null),
      ]);
      if (!alive) return;
      if (kind === 'deck') {
        // An outline is not printed here: its slides are. A deck's file keeps no paper.
        const layout = fromFile?.layout === 'outline' ? 'slides' : fromFile?.layout;
        setOptions((o) => ({ ...o, ...(paper ? { paper } : {}), ...(fromFile ? { layout, perPage: fromFile.perPage, frame: fromFile.frame } : {}) }));
      } else if (fromFile) {
        // A workbook whose file does not say which way is printed the way its shape asks, and one
        // whose file names no paper starts on the paper this person prints on.
        const { orientationStated, paperStated, ...kept } = fromFile;
        const filePaper = kind === 'sheet' && paperStated === false ? null : fromFile.paper;
        setOptions((o) => ({ ...o, ...kept, paper: filePaper || paper || o.paper, ...(kind === 'sheet' && !orientationStated ? { orientation: 'auto' } : {}) }));
        setMargins(presetOf(kind === 'doc' ? DOC_MARGIN_PRESETS : MARGIN_PRESETS, fromFile.margins) || (kind === 'doc' ? 'Custom' : 'Normal'));
      } else if (paper) {
        setOptions((o) => ({ ...o, paper }));
      }
      setStarted(true);
    })();
    return () => {
      alive = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, doc?.id, kind]);

  /**
   * A paper a person chose is the one this app's next print starts on. A document's is not kept:
   * its paper is the document's own.
   */
  const rememberPaper = (paper) => {
    if (kind === 'doc' || !app || !PAPERS.includes(paper)) return;
    Promise.resolve().then(() => shell.store.set({ key: paperKey(app), value: paper })).catch(() => {});
  };

  /** A document's page changed here is changed in the document, as Word's Print changes it. */
  const change = async (patch) => {
    if (patch.paper) rememberPaper(patch.paper);
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

  // The printers, twice: at once with the default one named, and again when the system has said which
  // are offline, paused or out of paper (that takes it a second or two, and the list is not kept waiting).
  useEffect(() => {
    let alive = true;
    (async () => {
      let list = [];
      try {
        list = (await shell.print.printers()) || [];
      } catch {
        list = [];
      }
      if (!alive) return;
      setPrinters(list);
      setPrinter(printerToOpenOn(list));
      if (!list.length) return;
      try {
        const marked = await shell.print.printers({ states: true });
        if (!alive || !Array.isArray(marked) || !marked.length) return;
        setPrinters(marked);
        // The default the quick read could not name is taken from the slow one, unless a printer was chosen meanwhile.
        if (!chose.current && !list.some((p) => p.default)) setPrinter(printerToOpenOn(marked));
      } catch {
        /* the list stays as it is, unmarked */
      }
    })();
    return () => {
      alive = false;
    };
  }, [shell]);

  // How many pages, recomputed whenever a choice changes. The engines answer
  // this from geometry alone, so it costs nothing to ask on every keystroke.
  //
  // A job's page is counted by laying it out, which is real work: it waits
  // for the choices to settle, and an answer to an older choice is dropped.
  const jobNow = useRef(job);
  jobNow.current = job;
  /**
   * A job's choices as it is laid out and printed by: with no way round chosen, the way the job says
   * its own shape asks for, and upright when it says nothing.
   */
  const laidBy = (o) => (wayChosen(o.orientation) ? o : { ...o, orientation: jobNow.current?.orientation?.(o) === 'landscape' ? 'landscape' : 'portrait' });
  //
  // Either way one layout is in hand at a time (oneAtATime), and work that
  // waited behind another and is no longer wanted when its turn comes is not
  // begun. Nothing is counted before the paper to start on is known.
  useEffect(() => {
    if (!started) return undefined;
    let alive = true;
    if (job) {
      setSummary(null);
      setFresh(false);
      const timer = setTimeout(() => {
        lanes.current
          .picture(async () => {
            if (!alive) return null;
            const html = await jobNow.current.html(laidBy(options));
            // Counted by laying it out as a PDF, and page one of that PDF is the small picture at the top.
            return shell.print.pageSummary({ html, name: jobNow.current.name, preview: true });
          })
          .then((got) => {
            if (!alive || got.dropped || !got.value) return;
            const { first: one = null, ...s } = got.value;
            setSummary(s);
            setFirst(one);
            setFresh(true);
          })
          .catch((err) => {
            if (!alive) return;
            setSummary({ error: err?.message || String(err) });
            setFirst(null);
          });
      }, 250);
      return () => {
        alive = false;
        clearTimeout(timer);
      };
    }
    lanes.current
      .count(() => (alive ? shell.print.summary({ id: doc.id, options }) : null))
      .then((got) => alive && !got.dropped && setSummary(got.value))
      .catch((err) => alive && setSummary({ error: err.message }));
    return () => {
      alive = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, doc?.id, options, started]);

  const pages = summary?.error ? null : summary?.pages ?? null;

  // Page one of a document, a workbook or a deck, asked for once its choices have settled. A document
  // is a PDF already; a workbook and a deck are laid out for it, so a very long one goes without.
  const pictured = !job && pages != null && pages > 0 && (kind === 'doc' || pages <= PREVIEW_UP_TO) && typeof shell.print.preview === 'function';
  useEffect(() => {
    if (job) return undefined;
    if (!pictured) {
      if (pages != null) setFirst(null);
      return undefined;
    }
    let alive = true;
    setFresh(false);
    const timer = setTimeout(() => {
      lanes.current
        .picture(() => (alive ? shell.print.preview({ id: doc.id, options }) : null))
        .then((got) => {
          if (!alive || got.dropped) return;
          setFirst(got.value?.first || null);
          setFresh(true);
        })
        .catch(() => alive && setFirst(null));
    }, 350);
    return () => {
      alive = false;
      clearTimeout(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, doc?.id, options, pictured]);
  // The pages typed, read against how many there are: a list to print, every page, or not pages of this print.
  const typed = useMemo(() => pagesTyped(range, pages), [range, pages]);
  const only = typed.pages || null;
  const detail = useMemo(() => {
    if (summary?.error) return summary.error;
    if (pages == null) return t('print.working-out-pages');
    if (typed.wrong) return tn(pages, 'print.not-pages-print');
    if (only) return tn(pages, 'print.chosen-count-page', { chosen: only.length });
    const scale = summary?.sheets?.[0]?.scale;
    return scale && scale < 0.999
      ? tn(pages, 'print.count-page-shrunk', { percent: Math.round(scale * 100) })
      : tn(pages, 'common.count-page');
  }, [summary, pages, typed, only]);

  // The way round in force: the one chosen, or, on Automatic, the one the job or the engine says.
  const auto = !wayChosen(options.orientation);
  const way = !auto ? options.orientation
    : job ? laidBy(options).orientation
      : wayChosen(summary?.setup?.orientation) ? summary.setup.orientation : null;

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

  /**
   * Print: straight to the printer chosen here, on the paper chosen here. With `system`, through the
   * system's own dialog instead, opened on the same printer and paper.
   */
  const run = async ({ system = false } = {}) => {
    setBusy(true);
    try {
      await remember();
      const how = { printer, copies: Math.max(1, Math.min(99, Number(copies) || 1)), silent: !system, ...(only ? { pages: only } : {}) };
      const laid = job ? laidBy(options) : options;
      const answer = job
        ? await shell.print.page({ html: await job.html(laid), name: job.name, options: laid, ...how })
        : await shell.print.document({ id: doc.id, options, ...how });
      const came = printOutcome(answer, { system, count: only ? only.length : pages });
      if (came.kind === 'sent-count') toast(tn(came.count, 'print.sent-count-page', { printer }), { tone: 'good' });
      else if (came.kind === 'sent') toast(t('print.sent-to-printer', { printer }), { tone: 'good' });
      else if (came.kind === 'sent-system') toast(t('print.sent-through-system'), { tone: 'good' });
      else if (came.kind === 'settings') toast(t('print.printer-did-not'), { tone: 'bad', ms: 8000 });
      else if (came.kind === 'failed') toast(t('print.printer-did-not-take'), { tone: 'bad', ms: 8000 });
      // A print that went closes the dialog. One that did not leaves it open, to try another printer or save a PDF.
      if (came.done) onClose();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setBusy(false);
    }
  };

  /** A job's page saved as a PDF where the person chooses: the same sheets the printer would get. */
  const savePage = async () => {
    setBusy(true);
    try {
      const html = await job.html(laidBy(options));
      const folder = (await shell.app.paths().catch(() => ({}))).documents || '';
      const file = `${printFileStem(job.name)}.pdf`;
      const target = await shell.dialog.save({ title: t('print.save-as-pdf'), defaultPath: folder ? `${folder}/${file}` : file, filters: [{ name: 'PDF', extensions: ['pdf'] }] });
      if (!target) return;
      await shell.print.pagePdf({ html, name: job.name, path: target, ...(only ? { pages: only } : {}) });
      toast(t('common.saved-name', { name: basename(target) }), { tone: 'good', ms: 3000 });
      onClose();
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
    } finally {
      setBusy(false);
    }
  };
  // A job that cannot be laid out (a picture over more sheets than a print can be) says why, and prints nothing.
  // Nor does a print whose Pages box holds something that is not pages of it, or is not counted yet.
  const stopped = Boolean(job && summary?.error) || Boolean(range.trim() && !only);
  const none = Array.isArray(printers) && !printers.length;
  const ready = Boolean(printers?.length);
  const chosen = printers?.find((p) => p.name === printer) || null;
  // What is about to print, by its name: the file's, or the one the app gave its page (a message's subject).
  const name = job?.name || doc?.name || summary?.name || '';
  const printerHint = none
    ? t('print.no-printer-installed-use')
    : chosen?.state
      ? t('print.printer-not-ready')
      : undefined;

  return (
    <Dialog
      title={t('common.print')}
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label={t('print.save-as-pdf')} icon="pdf" disabled={busy || stopped} onClick={job ? savePage : async () => { await remember(); onClose(); onSaveAs?.(only ? { ...options, pages: only } : options); }} />
          <Button
            className="rw-print-system"
            label={t('print.system-dialog')}
            title={t('print.print-through-systems')}
            disabled={busy || stopped || !ready}
            onClick={() => run({ system: true })}
          />
          <span style={{ flex: 1 }} />
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={busy ? t('print.printing') : t('common.print')} icon="print" disabled={busy || stopped || !ready} onClick={() => run()} />
        </>
      }
    >
      <div className="rw-print-what" style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 12 }}>
        {first === null ? null : <FirstPage bytes={first} fresh={fresh} />}
        <div style={{ flex: 1, minWidth: 0 }}>
          {name ? (
            <div className="rw-print-name" title={name} style={{ fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {t('print.printing-name', { name })}
            </div>
          ) : null}
          <div style={{ marginTop: 6 }}>
            <Chip>{detail}</Chip>
          </div>
        </div>
      </div>

      <Field label={t('print.printer')} hint={printerHint}>
        <Select
          className="rw-select rw-print-printer"
          value={printer}
          disabled={!ready}
          onChange={(e) => {
            chose.current = true;
            setPrinter(e.target.value);
          }}
          style={{ width: '100%' }}
        >
          {none ? <option value="">{t('print.no-printer-installed')}</option> : null}
          {(printers || []).map((p) => (
            <option key={p.name} value={p.name}>
              {[p.name, p.default ? t('print.default') : '', STATE_WORDS[p.state] ? t(STATE_WORDS[p.state]) : ''].filter(Boolean).join(' ')}
            </option>
          ))}
        </Select>
      </Field>

      <div style={{ display: 'flex', gap: 12 }}>
        <Field label={t('print.copies')}>
          <Input type="number" min="1" max="99" value={copies} onChange={(e) => setCopies(e.target.value)} style={{ width: 80 }} />
        </Field>
        <div style={{ flex: 1, minWidth: 0 }}>
          <Field label={t('common.pages')} hint={t('print.empty-prints-every')}>
            <Input className="rw-input rw-print-pages" value={range} placeholder="1-5, 8" inputMode="numeric" autoComplete="off" spellCheck={false} aria-invalid={typed.wrong ? 'true' : undefined} onChange={(e) => setRange(e.target.value)} style={{ width: '100%' }} />
          </Field>
        </div>
      </div>

      <div style={{ display: 'flex', gap: 12 }}>
        <Field label={t('common.paper')}>
          <Select value={options.paper} onChange={(e) => change({ paper: e.target.value })} style={{ width: 110 }}>
            {PAPERS.map((p) => (
              <option key={p} value={p}>{p}</option>
            ))}
          </Select>
        </Field>
        <Field label={t('common.orientation')}>
          <Select className="rw-select rw-print-way" value={auto ? 'auto' : options.orientation} onChange={(e) => change({ orientation: e.target.value })} style={{ width: 190 }}>
            {kind === 'doc' ? null : (
              <option value="auto">{way === 'landscape' ? t('print.automatic-landscape') : way === 'portrait' ? t('print.automatic-portrait') : t('common.automatic')}</option>
            )}
            <option value="portrait">{t('common.portrait')}</option>
            <option value="landscape">{t('common.landscape')}</option>
          </Select>
        </Field>
        <Field label={t('common.margins')}>
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
            {margins === 'Custom' ? <option value="Custom">{t('common.custom')}</option> : null}
          </Select>
        </Field>
      </div>

      {kind === 'sheet' ? (
        <>
          <div style={{ display: 'flex', gap: 12 }}>
            <Field label={t('common.print')}>
              <Select value={options.sheets === 'all' ? 'all' : 'active'} onChange={(e) => set({ sheets: e.target.value === 'all' ? 'all' : null })} style={{ width: 170 }}>
                <option value="active">{t('common.this-sheet')}</option>
                <option value="all">{t('print.every-sheet-count', { count: sheets.length })}</option>
              </Select>
            </Field>
            <Field label={t('print.scaling')}>
              <Select value={options.fit} onChange={(e) => set({ fit: e.target.value })} style={{ width: 190 }}>
                <option value="none">{t('common.no-scaling')}</option>
                <option value="width">{t('print.fit-all-columns')}</option>
                <option value="page">{t('print.fit-sheet-one')}</option>
              </Select>
            </Field>
            <Field label={t('print.repeat-rows')}>
              <Input type="number" min="0" max="10" value={options.repeatRows} onChange={(e) => set({ repeatRows: Number(e.target.value) || 0 })} style={{ width: 80 }} title={t('print.rows-top-sheet')} />
            </Field>
          </div>
          <Field label={t('print.print-area')} hint={t('print.empty-prints-everything')}>
            <Input value={options.area || ''} placeholder="A1:H60" onChange={(e) => set({ area: e.target.value.trim() || null })} style={{ width: '100%' }} />
          </Field>
          <div style={{ display: 'flex', gap: 16, marginTop: 4, flexWrap: 'wrap' }}>
            <label><input type="checkbox" checked={options.gridlines} onChange={(e) => set({ gridlines: e.target.checked })} /> {t('common.gridlines')}</label>
            <label><input type="checkbox" checked={options.headings} onChange={(e) => set({ headings: e.target.checked })} /> {t('print.row-column-headings')}</label>
            <label title={t('print.excel-keeps-page')}>
              <input type="checkbox" checked={keep} onChange={(e) => setKeep(e.target.checked)} /> {t('print.keep-settings-workbook')}
            </label>
          </div>
        </>
      ) : null}

      {kind === 'deck' ? (
        <div style={{ display: 'flex', gap: 12 }}>
          <Field label={t('common.print')}>
            <Select
              value={options.layout}
              onChange={(e) => set({ layout: e.target.value })}
              style={{ width: 190 }}
            >
              <option value="slides">{t('print.full-page-slides')}</option>
              <option value="notes">{t('print.notes-pages')}</option>
              <option value="handout">{t('print.handout')}</option>
            </Select>
          </Field>
          {options.layout === 'handout' ? (
            <Field label={t('print.slides-page')}>
              <Select value={options.perPage} onChange={(e) => set({ perPage: Number(e.target.value) })} style={{ width: 100 }}>
                {[1, 2, 3, 4, 6, 9].map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </Select>
            </Field>
          ) : null}
          <Field label={t('common.frame')}>
            <label style={{ display: 'block', paddingTop: 6 }}>
              <input type="checkbox" checked={options.frame} onChange={(e) => set({ frame: e.target.checked })} /> {t('print.draw-border-round')}
            </label>
          </Field>
        </div>
      ) : null}
      {kind === 'deck' ? (
        <label title={t('print.power-point-keeps')}>
          <input type="checkbox" checked={keep} onChange={(e) => setKeep(e.target.checked)} /> {t('print.keep-settings')}
        </label>
      ) : null}
      {kind === 'doc' ? (
        <p className="rw-hint" style={{ margin: '6px 0 0' }}>{t('print.paper-orientation')}</p>
      ) : null}
      {job?.choices ? job.choices({ options, set }) : null}
    </Dialog>
  );
}
