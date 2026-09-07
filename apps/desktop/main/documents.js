// The document service.
//
// Documents open, edit and save in the main process, because the engines are
// Node engines — they inflate with zlib and work in Buffers — and because a
// 200 MB workbook has no business in a window's heap. The renderer draws
// whatever view model comes back and sends back operations by name.
//
// Formats we cannot write are converted on open into ones we can: an .odt
// becomes a document, an .ods a workbook, a .csv a workbook. That way there is
// one editor per kind rather than one per format, and Save As is always honest
// about what it is about to write.

import fs from 'node:fs';
import path from 'node:path';
import { SheetView } from '@rutba/sheet-view';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildXlsx, buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { Deck, buildPptx, renderSlide, TEMPLATES as DECK_TEMPLATES } from '@rutba/presentation';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { sniff, refineOoxml } from '@rutba/office-formats/sniff';
import { readOdf } from '@rutba/office-formats/odf';
import { readRtf } from '@rutba/office-formats/rtf';
import { readDelimited, writeDelimited, readMarkdown, readPlain, writeMarkdown, writePlain, decodeText } from '@rutba/office-formats/text';
import { CompoundFile } from '@rutba/office-formats/cfb';
import { readZip } from '@rutba/ooxml/zip';

let seq = 0;
const KIND_FOR_APP = { word: 'doc', sheets: 'sheet', slides: 'deck' };

/** Blocks in the neutral reader shape → paragraphs `buildDocx` understands. */
function blocksToParagraphs(blocks) {
  const out = [];
  for (const b of blocks || []) {
    if (b.type === 'table') {
      for (const row of b.rows || []) {
        const cells = row.map((c) => (c.text ?? (c.runs || []).map((r) => r.text).join('')) || '');
        out.push({ text: cells.join('\t') });
      }
      continue;
    }
    if (b.type === 'list') {
      for (const item of b.items || []) out.push({ text: `• ${item.text ?? (item.runs || []).map((r) => r.text).join('')}` });
      continue;
    }
    const runs = (b.runs || []).map((r) => ({ text: r.text ?? '', bold: r.bold, italic: r.italic, underline: r.underline }));
    const text = b.text ?? runs.map((r) => r.text).join('');
    if (b.type === 'heading') out.push({ text, style: `Heading${Math.min(b.level || 1, 6)}`, bold: true, size: 20 - Math.min(b.level || 1, 5) * 2 });
    else if (b.type === 'code') out.push({ text: b.text, font: 'Consolas' });
    else if (b.type === 'rule') out.push({ text: '' });
    else out.push(runs.length > 1 ? { runs, text } : { text, ...(runs[0] || {}) });
  }
  return out.length ? out : [{ text: '' }];
}

/** Rows of text → a one-sheet workbook, with numbers left as numbers. */
function rowsToWorkbook(rows, name = 'Sheet1') {
  const cells = rows.map((row) =>
    row.map((value) => {
      if (value == null || value === '') return null;
      const trimmed = String(value).trim();
      const asNumber = Number(trimmed.replace(/,/g, ''));
      if (trimmed !== '' && Number.isFinite(asNumber) && /^-?[\d,]*\.?\d+%?$/.test(trimmed)) return asNumber;
      return String(value);
    })
  );
  return buildXlsx({ sheets: [{ name: name.slice(0, 31) || 'Sheet1', rows: cells }] });
}

function odfSheetsToWorkbook(sheets) {
  return buildXlsx({
    sheets: (sheets.length ? sheets : [{ name: 'Sheet1', rows: [] }]).map((s, i) => ({
      name: (s.name || `Sheet${i + 1}`).slice(0, 31),
      rows: (s.rows || []).map((row) =>
        row.map((cell) => {
          if (cell.formula) return cell.formula.startsWith('=') ? cell.formula : '=' + cell.formula;
          if (cell.type === 'float' || cell.type === 'percentage' || cell.type === 'currency') return Number(cell.value ?? cell.text) || 0;
          return cell.text || null;
        })
      ),
    })),
  });
}

function odpSlidesToDeck(slides, title) {
  return buildPptx({
    title: title || 'Presentation',
    slides: (slides.length ? slides : [{ shapes: [] }]).map((s) => {
      const texts = s.shapes.filter((sh) => sh.type === 'text');
      const [first, ...rest] = texts;
      return {
        layout: 'obj',
        title: first ? first.paragraphs.join(' ') : s.name || '',
        body: rest.flatMap((t) => t.paragraphs).filter(Boolean),
      };
    }),
  });
}

// Word 97 separates paragraphs with CR, and litters the stream with field
// and bookmark control codes. Splitting on all of them gives readable lines.
const PARAGRAPH_MARKS = new RegExp('[\\u0000-\\u0008\\u000b\\u000c\\u000d\\u000e-\\u001f]', 'g');

/** A .doc/.xls/.ppt: extract what text we can rather than refuse the file. */
function legacyText(bytes) {
  const cfb = new CompoundFile(bytes);
  const app = cfb.application();
  if (app === 'doc') {
    const stream = cfb.find(['WordDocument']);
    const raw = stream ? cfb.read(stream) : new Uint8Array(0);
    // Word 97 stores the text run at fcMin; without the piece table the
    // readable approximation is the printable Latin/Unicode runs in order.
    const text = new TextDecoder('windows-1252', { fatal: false }).decode(raw);
    const cleaned = text
      .replace(PARAGRAPH_MARKS, String.fromCharCode(10))
      .split('\n')
      .map((line) => line.trim())
      .filter((line) => line.length > 3 && /[a-zA-Z]{3}/.test(line));
    return { app, blocks: cleaned.map((line) => ({ type: 'paragraph', text: line })) };
  }
  return { app, blocks: [] };
}

class Session {
  constructor({ id, kind, filePath, engine, source, converted }) {
    this.id = id;
    this.kind = kind;
    this.path = filePath;
    this.engine = engine;
    this.source = source;
    this.converted = converted || null;
    this.dirty = false;
    this.version = 0;
    this.opened = Date.now();
  }

  get name() {
    return this.path ? path.basename(this.path) : this.kind === 'sheet' ? 'Book1.xlsx' : this.kind === 'deck' ? 'Presentation1.pptx' : 'Document1.docx';
  }

  meta() {
    return {
      id: this.id,
      kind: this.kind,
      path: this.path,
      name: this.name,
      dirty: this.dirty,
      version: this.version,
      source: this.source,
      converted: this.converted,
      readOnly: false,
      canUndo: Boolean(this.engine?.canUndo),
      canRedo: Boolean(this.engine?.canRedo),
    };
  }
}

export function createDocumentService({ holdBlob }) {
  /** @type {Map<string, Session>} */
  const sessions = new Map();
  const nextId = () => `d${++seq}`;

  const get = (id) => {
    const s = sessions.get(id);
    if (!s) throw new Error('That document is no longer open.');
    return s;
  };

  /** Decide the kind, converting the formats we do not write. */
  function load(bytes, filePath) {
    const detected = sniff(bytes, filePath ? path.basename(filePath) : '');
    let kind = detected.kind;

    if (kind === 'zip' || detected.container === 'ooxml') {
      try {
        kind = refineOoxml(readZip(Buffer.from(bytes)).entries.map((e) => e.name));
      } catch {
        /* keep what sniff said */
      }
    }

    const buf = Buffer.from(bytes);
    switch (kind) {
      case 'xlsx':
      case 'xlsm':
      case 'xltx':
        return { kind: 'sheet', bytes: buf, source: kind };
      case 'docx':
      case 'docm':
      case 'dotx':
        return { kind: 'doc', bytes: buf, source: kind };
      case 'pptx':
      case 'pptm':
      case 'potx':
      case 'ppsx':
        return { kind: 'deck', bytes: buf, source: kind };
      case 'csv':
      case 'tsv': {
        const { rows, delimiter } = readDelimited(bytes);
        return {
          kind: 'sheet',
          bytes: rowsToWorkbook(rows, filePath ? path.basename(filePath, path.extname(filePath)) : 'Sheet1'),
          source: kind,
          converted: { from: kind, delimiter },
        };
      }
      case 'ods': {
        const odf = readOdf(bytes);
        return { kind: 'sheet', bytes: odfSheetsToWorkbook(odf.sheets || []), source: 'ods', converted: { from: 'ods' } };
      }
      case 'odt': {
        const odf = readOdf(bytes);
        return { kind: 'doc', bytes: buildDocx({ paragraphs: blocksToParagraphs(odf.blocks), styles: true }), source: 'odt', converted: { from: 'odt' } };
      }
      case 'odp': {
        const odf = readOdf(bytes);
        return { kind: 'deck', bytes: odpSlidesToDeck(odf.slides || [], odf.meta?.title), source: 'odp', converted: { from: 'odp' } };
      }
      case 'rtf': {
        const rtf = readRtf(bytes);
        return { kind: 'doc', bytes: buildDocx({ paragraphs: blocksToParagraphs(rtf.blocks), styles: true }), source: 'rtf', converted: { from: 'rtf' } };
      }
      case 'md': {
        const md = readMarkdown(bytes);
        return { kind: 'doc', bytes: buildDocx({ paragraphs: blocksToParagraphs(md.blocks), styles: true }), source: 'md', converted: { from: 'md' } };
      }
      case 'txt':
      case 'html': {
        const parsed = kind === 'txt' ? readPlain(bytes) : { blocks: [{ type: 'paragraph', text: decodeText(bytes).replace(/<[^>]+>/g, ' ') }] };
        return { kind: 'doc', bytes: buildDocx({ paragraphs: blocksToParagraphs(parsed.blocks), styles: true }), source: kind, converted: { from: kind } };
      }
      case 'doc':
      case 'xls':
      case 'ppt': {
        const legacy = legacyText(bytes);
        if (kind === 'xls') {
          return {
            kind: 'sheet',
            bytes: rowsToWorkbook([['This 97-2003 workbook could not be converted in full.']]),
            source: kind,
            converted: { from: kind, partial: true },
          };
        }
        return {
          kind: kind === 'ppt' ? 'deck' : 'doc',
          bytes:
            kind === 'ppt'
              ? buildPptx({ title: 'Imported presentation', slides: [{ layout: 'title', title: 'Imported presentation', body: 'The original slides could not be converted in full.' }] })
              : buildDocx({ paragraphs: blocksToParagraphs(legacy.blocks), styles: true }),
          source: kind,
          converted: { from: kind, partial: true },
        };
      }
      default:
        throw new Error(`Rutba Office cannot open ${detected.label || kind} files yet.`);
    }
  }

  function engineFor(kind, bytes) {
    if (kind === 'sheet') return SheetView.open(Buffer.from(bytes), { viewportWidth: 1100, viewportHeight: 620 });
    if (kind === 'doc') return openDocx(Buffer.from(bytes));
    if (kind === 'deck') return Deck.open(Buffer.from(bytes));
    throw new Error(`unknown document kind: ${kind}`);
  }

  const TEMPLATES = {
    sheet: () => buildXlsx({ sheets: [{ name: 'Sheet1', rows: [] }] }),
    budget: () =>
      buildXlsx({
        sheets: [
          {
            name: 'Budget',
            rows: [
              ['Item', 'Budget', 'Actual', 'Variance'],
              ['Rent', 1200, 1200, '=B2-C2'],
              ['Utilities', 180, 0, '=B3-C3'],
              ['Supplies', 240, 0, '=B4-C4'],
              ['Total', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)'],
            ],
          },
        ],
      }),
    invoice: () =>
      buildXlsx({
        sheets: [
          {
            name: 'Invoice',
            rows: [
              ['Invoice'],
              [],
              ['Description', 'Quantity', 'Unit price', 'Amount'],
              ['', 1, 0, '=B4*C4'],
              ['', 1, 0, '=B5*C5'],
              [],
              ['', '', 'Total', '=SUM(D4:D5)'],
            ],
          },
        ],
      }),
    doc: () => buildDocx({ paragraphs: [{ text: '' }], styles: true }),
    letter: () =>
      buildDocx({
        styles: true,
        paragraphs: [
          { text: 'Your name', bold: true, size: 13 },
          { text: 'Your address' },
          { text: '' },
          { text: new Date().toLocaleDateString(undefined, { dateStyle: 'long' }) },
          { text: '' },
          { text: 'Dear …,' },
          { text: '' },
          { text: '' },
          { text: 'Yours sincerely,' },
        ],
      }),
    report: () =>
      buildDocx({
        styles: true,
        paragraphs: [
          { text: 'Report title', style: 'Heading1', bold: true, size: 20 },
          { text: 'Prepared by … · ' + new Date().toLocaleDateString() },
          { text: '' },
          { text: 'Summary', style: 'Heading2', bold: true, size: 15 },
          { text: '' },
          { text: 'Findings', style: 'Heading2', bold: true, size: 15 },
          { text: '' },
        ],
      }),
    deck: () => DECK_TEMPLATES.blank(),
    pitch: () => DECK_TEMPLATES.pitch(),
  };

  /* ── models ───────────────────────────────────────────────────────────── */

  function sheetModel(session) {
    const view = session.engine;
    const frame = view.render();
    return {
      ...frame,
      sheets: view.sheetNames(),
      activeSheet: view.activeSheet,
      canUndo: view.canUndo,
      canRedo: view.canRedo,
      names: typeof view.names === 'function' ? view.names() : [],
    };
  }

  function docModel(session) {
    const view = session.engine;
    const frame = view.render();
    // Remember what the window now holds, so the next edit can send only the
    // difference rather than the document.
    session.lastBlocks = frame.blocks.map((b) => JSON.stringify(b));
    return {
      ...frame,
      canUndo: view.canUndo,
      canRedo: view.canRedo,
      canEdit: view.canEdit,
      styles: typeof view.paragraphStyles === 'object' ? view.paragraphStyles : [],
      format: typeof view.formatAtCaret === 'function' ? view.formatAtCaret() : null,
    };
  }

  /**
   * What changed, rather than what there is.
   *
   * Typing into a four-hundred paragraph document produced a 380 KB model and
   * re-rendered every block, per keystroke — the engine was never the slow
   * part, the round trip was. Almost every edit touches one paragraph, so the
   * blocks are compared against what the window was last sent and only the run
   * that actually differs crosses: a common prefix, a common suffix, and the
   * span between them.
   *
   * When the shape of the change is not a simple splice — a style sweep, a
   * find-and-replace — this hands back the whole model instead. Correct is the
   * floor; fast is the goal above it.
   */
  function docDelta(session) {
    const view = session.engine;
    const frame = view.render();
    const next = frame.blocks.map((b) => JSON.stringify(b));
    const prev = session.lastBlocks;
    session.lastBlocks = next;

    const common = {
      selection: frame.selection,
      format: typeof view.formatAtCaret === 'function' ? view.formatAtCaret() : null,
      listLabels: frame.listLabels,
      wordCount: frame.wordCount,
      characterCount: frame.characterCount,
      section: frame.section,
      canUndo: view.canUndo,
      canRedo: view.canRedo,
    };

    if (!prev) return { model: docModel(session) };

    let head = 0;
    while (head < prev.length && head < next.length && prev[head] === next[head]) head++;
    let tail = 0;
    while (
      tail < prev.length - head &&
      tail < next.length - head &&
      prev[prev.length - 1 - tail] === next[next.length - 1 - tail]
    ) {
      tail++;
    }

    const removed = prev.length - head - tail;
    const inserted = frame.blocks.slice(head, next.length - tail);

    // A change touching more than a screenful is not worth splicing.
    if (inserted.length > 60) return { model: { ...frame, ...common, styles: view.paragraphStyles } };

    return { patch: { from: head, removed, blocks: inserted, ...common } };
  }

  function deckModel(session, { slide = 0, width = 960 } = {}) {
    const deck = session.engine;
    const count = deck.slideCount;
    const index = Math.max(0, Math.min(slide, Math.max(0, count - 1)));
    const current = count ? deck.slide(index) : null;
    const resolveImage = (shape) => {
      if (!shape.source?.part) return null;
      const bytes = deck.media(shape.source.part);
      if (!bytes) return null;
      const type = shape.source.part.endsWith('.png') ? 'image/png' : shape.source.part.endsWith('.gif') ? 'image/gif' : 'image/jpeg';
      return holdBlob(bytes, type, path.basename(shape.source.part)).url;
    };
    return {
      count,
      index,
      size: deck.size,
      outline: deck.outline().map((o) => ({ ...o, thumbnail: null })),
      slide: current
        ? {
            ...current,
            svg: renderSlide(current, { width, resolveImage }),
            shapes: current.shapes.map((s) => ({
              id: s.id,
              kind: s.kind,
              name: s.name,
              geometry: s.geometry,
              placeholder: s.placeholder,
              text: (s.text || s.inheritedText)
                ? { paragraphs: (s.text || s.inheritedText).paragraphs.map((p) => ({ ...p, plain: p.runs.map((r) => r.text).join('') })) }
                : null,
            })),
          }
        : null,
      canUndo: false,
      canRedo: false,
    };
  }

  function modelOf(session, opts) {
    if (session.kind === 'sheet') return sheetModel(session);
    if (session.kind === 'doc') return docModel(session);
    return deckModel(session, opts);
  }

  /* ── operations ───────────────────────────────────────────────────────── */

  const SHEET_OPS = {
    select: (v, a) => v.select(a.row, a.col, { extend: a.extend }),
    selectRow: (v, a) => v.selectRow(a.row, { extend: a.extend }),
    selectColumn: (v, a) => v.selectColumn(a.col, { extend: a.extend }),
    move: (v, a) => v.moveSelection(a.direction, a),
    scrollTo: (v, a) => v.scrollTo(a.x, a.y),
    viewport: (v, a) => {
      v.viewportWidth = a.width;
      v.viewportHeight = a.height;
    },
    beginEdit: (v, a) => v.beginEdit(a),
    updateDraft: (v, a) => v.updateDraft(a.text),
    cancelEdit: (v) => v.cancelEdit(),
    commitEdit: (v, a) => v.commitEdit(a),
    setCell: (v, a) => v.setCell(a.row, a.col, a.value),
    clear: (v) => v.clearSelection(),
    sheet: (v, a) => v.selectSheet(a.name),
    merge: (v) => v.mergeSelection(),
    unmerge: (v) => v.unmergeSelection(),
    colWidth: (v, a) => v.setColWidth(a.col, a.width),
    rowHeight: (v, a) => v.setRowHeight(a.row, a.height),
    insertRows: (v, a) => v.insertRows(a.at, a.count || 1),
    deleteRows: (v, a) => v.deleteRows(a.at, a.count || 1),
    insertCols: (v, a) => v.insertCols(a.at, a.count || 1),
    deleteCols: (v, a) => v.deleteCols(a.at, a.count || 1),
    autoSum: (v, a) => v.autoSum(a.fn || 'SUM'),
    sort: (v, a) => v.sortSelection({ ascending: a.ascending !== false }),
    paste: (v, a) => v.pasteText(a.text, a.html),
    copy: (v) => v.markClipboard(),
    fill: (v, a) => v.fill(a.target),
    tab: (v, a) => v.tab(a.back),
    enter: (v, a) => v.enterKey(a.back),
    insertChart: (v, a) => v.insertChart(a),
    insertShape: (v, a) => v.insertShape(a),
    formatBrush: (v) => v.markFormatBrush(),
    paintFormat: (v) => v.paintFormat(),
  };

  /** What the ribbon calls a format, and what the document engine calls it. */
  const RUN_FORMATS = {
    bold: 'b',
    italic: 'i',
    underline: 'u',
    strike: 's',
    strikethrough: 's',
  };

  const DOC_OPS = {
    setSelection: (v, a) => v.setSelection(a.anchor, a.focus ?? a.anchor),
    moveCaret: (v, a) => v.moveCaret(a.direction, { extend: a.extend }),
    selectAll: (v) => v.selectAll(),
    insertText: (v, a) => v.insertText(a.text),
    deleteBackward: (v) => v.deleteBackward(),
    deleteForward: (v) => v.deleteForward(),
    deleteSelection: (v) => v.deleteSelection(),
    splitParagraph: (v) => v.splitParagraph(),
    // The renderer speaks in words and the engine in OOXML's letters. The
    // translation belongs here rather than in the ribbon, which should not have
    // to know that bold is called "b".
    toggleFormat: (v, a) => v.toggleFormat(RUN_FORMATS[a.tag] ?? a.tag),
    setRunFormat: (v, a) => v.setRunFormat(a.delta),
    clearFormat: (v) => v.clearFormat(),
    setParagraphFormat: (v, a) => v.setParagraphFormat(a.delta),
    setLink: (v, a) => v.setLink(a.url),
    insertTable: (v, a) => v.insertTable(a),
    insertImage: (v, a) => v.insertImage(a),
    insertPageBreak: (v) => v.insertPageBreak(),
    insertChart: (v, a) => v.insertChart(a),
    insertShape: (v, a) => v.insertShape(a),
    replaceAll: (v, a) => v.replaceAll(a.find, a.replace, { matchCase: a.matchCase }),
    pasteText: (v, a) => v.pasteText(a.text),
    setPageSetup: (v, a) => v.setPageSetup(a.spec),
    tableOp: (v, a) => v.tableOp(a.op, a.arg),
    addComment: (v, a) => v.addComment(a.text, { author: a.author }),
    tabCell: (v, a) => v.tabCell({ back: a.back }),
  };

  const DECK_OPS = {
    setText: (d, a) => d.setText(a.slide, a.shape, a.paragraphs),
    setGeometry: (d, a) => d.setGeometry(a.slide, a.shape, a),
    removeShape: (d, a) => d.removeShape(a.slide, a.shape),
    addTextBox: (d, a) => d.addTextBox(a.slide, a),
    duplicateSlide: (d, a) => d.duplicateSlide(a.slide),
    removeSlide: (d, a) => d.removeSlide(a.slide),
    moveSlide: (d, a) => d.moveSlide(a.from, a.to),
  };

  const OPS = { sheet: SHEET_OPS, doc: DOC_OPS, deck: DECK_OPS };

  // Operations that only move the cursor or the viewport do not make a file
  // dirty; a document that says "unsaved changes" because somebody scrolled is
  // a document nobody trusts.
  const CLEAN_OPS = new Set(['select', 'selectRow', 'selectColumn', 'move', 'scrollTo', 'viewport', 'beginEdit', 'cancelEdit', 'setSelection', 'moveCaret', 'selectAll', 'copy', 'formatBrush', 'sheet']);

  /* ── the namespace ────────────────────────────────────────────────────── */

  return {
    new: ({ kind = 'doc', template }) => {
      const make = TEMPLATES[template] || TEMPLATES[KIND_FOR_APP[kind] || kind] || TEMPLATES.doc;
      const bytes = make();
      const resolved = template && TEMPLATES[template] ? (['budget', 'invoice', 'sheet'].includes(template) ? 'sheet' : ['pitch', 'deck'].includes(template) ? 'deck' : 'doc') : KIND_FOR_APP[kind] || kind;
      const session = new Session({ id: nextId(), kind: resolved, filePath: null, engine: engineFor(resolved, bytes), source: 'new' });
      sessions.set(session.id, session);
      return { ...session.meta(), model: modelOf(session) };
    },

    open: ({ path: filePath, width }) => {
      const bytes = fs.readFileSync(filePath);
      const loaded = load(bytes, filePath);
      const session = new Session({
        id: nextId(),
        kind: loaded.kind,
        filePath,
        engine: engineFor(loaded.kind, loaded.bytes),
        source: loaded.source,
        converted: loaded.converted,
      });
      sessions.set(session.id, session);
      return { ...session.meta(), model: modelOf(session, { width }) };
    },

    close: ({ id }) => {
      sessions.delete(id);
      return true;
    },

    meta: ({ id }) => get(id).meta(),

    model: ({ id, ...opts }) => modelOf(get(id), opts),

    apply: ({ id, ops, width, slide, delta = true }) => {
      const session = get(id);
      const table = OPS[session.kind];
      let touched = false;
      for (const op of ops || []) {
        const fn = table[op.op];
        if (!fn) throw new Error(`${session.kind} documents have no operation "${op.op}"`);
        fn(session.engine, op);
        if (!CLEAN_OPS.has(op.op)) touched = true;
      }
      if (touched) {
        session.dirty = true;
        session.version++;
      }
      // A document answers with the difference; the other kinds are already
      // small — a sheet sends only the viewport, a deck one slide.
      if (session.kind === 'doc' && delta) return { ...session.meta(), ...docDelta(session) };
      return { ...session.meta(), model: modelOf(session, { width, slide }) };
    },

    viewport: ({ id, width, height, x, y }) => {
      const session = get(id);
      if (session.kind !== 'sheet') return modelOf(session);
      const view = session.engine;
      if (width) view.viewportWidth = width;
      if (height) view.viewportHeight = height;
      if (x != null || y != null) view.scrollTo(x ?? view.scrollX, y ?? view.scrollY);
      return sheetModel(session);
    },

    undo: ({ id, width, slide }) => {
      const session = get(id);
      session.engine.undo?.();
      session.version++;
      return { ...session.meta(), model: modelOf(session, { width, slide }) };
    },

    redo: ({ id, width, slide }) => {
      const session = get(id);
      session.engine.redo?.();
      session.version++;
      return { ...session.meta(), model: modelOf(session, { width, slide }) };
    },

    save: ({ id, path: target }) => {
      const session = get(id);
      const to = target || session.path;
      if (!to) throw new Error('This document has never been saved, so it needs a name.');
      const ext = path.extname(to).toLowerCase();

      // Saving to a format the engine does not write is an export, and is
      // routed as one rather than writing a mislabelled file.
      const native = { sheet: ['.xlsx', '.xlsm', '.xltx'], doc: ['.docx', '.docm', '.dotx'], deck: ['.pptx', '.pptm', '.potx', '.ppsx'] }[session.kind];
      if (!native.includes(ext)) {
        return exportTo(session, to, ext.replace('.', ''));
      }

      const bytes = session.kind === 'deck' ? session.engine.save() : session.engine.save();
      fs.writeFileSync(to, Buffer.from(bytes));
      session.path = to;
      session.dirty = false;
      session.converted = null;
      return { ...session.meta(), path: to, stat: { size: fs.statSync(to).size } };
    },

    export: ({ id, format, path: target }) => exportTo(get(id), target, format),

    search: ({ id, query, options }) => {
      const session = get(id);
      if (session.kind === 'doc' && typeof session.engine.render === 'function') {
        const frame = session.engine.render();
        const hits = [];
        const needle = options?.matchCase ? query : query.toLowerCase();
        (frame.blocks || []).forEach((block, index) => {
          const text = (block.runs || []).map((r) => r.text).join('') || block.text || '';
          const hay = options?.matchCase ? text : text.toLowerCase();
          let at = hay.indexOf(needle);
          while (at >= 0 && needle) {
            hits.push({ block: index, offset: at, length: needle.length, preview: text.slice(Math.max(0, at - 30), at + needle.length + 30) });
            at = hay.indexOf(needle, at + needle.length);
          }
        });
        return hits;
      }
      if (session.kind === 'sheet') {
        const view = session.engine;
        const hits = [];
        const bounds = view.bounds;
        const needle = (options?.matchCase ? query : query.toLowerCase()) || '';
        for (let r = 0; r <= Math.min(bounds.maxRow, 5000) && needle; r++) {
          for (let c = 0; c <= Math.min(bounds.maxCol, 200); c++) {
            const text = String(view.displayValue(r, c)?.text ?? '');
            const hay = options?.matchCase ? text : text.toLowerCase();
            if (text && hay.includes(needle)) hits.push({ row: r, col: c, preview: text });
            if (hits.length >= 500) return hits;
          }
        }
        return hits;
      }
      return [];
    },

    asset: ({ id, ref }) => {
      const session = get(id);
      if (session.kind === 'deck') {
        const bytes = session.engine.media(ref);
        if (!bytes) return null;
        const type = ref.endsWith('.png') ? 'image/png' : ref.endsWith('.gif') ? 'image/gif' : ref.endsWith('.svg') ? 'image/svg+xml' : 'image/jpeg';
        return holdBlob(bytes, type, path.basename(ref));
      }
      const pkg = OoxmlPackage.read(Buffer.from(session.engine.serialize ? session.engine.serialize() : session.engine.save()));
      if (!pkg.has(ref)) return null;
      return holdBlob(pkg.read(ref), 'application/octet-stream', path.basename(ref));
    },

    sessions: () => [...sessions.values()].map((s) => s.meta()),
  };

  function exportTo(session, target, format) {
    const ext = (format || path.extname(target).replace('.', '')).toLowerCase();

    if (ext === 'pdf') {
      if (session.kind === 'doc') {
        // The document engine has its own PDF writer, which lays the document
        // out rather than photographing a screen.
        const bytes = renderPdf(session.engine, { title: session.name });
        if (bytes) {
          fs.writeFileSync(target, Buffer.from(bytes));
          return { path: target, format: 'pdf' };
        }
      }
      throw new Error('PDF export for this document is done from the window, so the page can be laid out first.');
    }

    if (session.kind === 'sheet' && (ext === 'csv' || ext === 'tsv')) {
      const view = session.engine;
      const bounds = view.bounds;
      const rows = [];
      for (let r = 0; r <= bounds.maxRow; r++) {
        const row = [];
        for (let c = 0; c <= bounds.maxCol; c++) row.push(view.displayValue(r, c)?.text ?? '');
        rows.push(row);
      }
      fs.writeFileSync(target, writeDelimited(rows, { delimiter: ext === 'tsv' ? '\t' : ',' }), 'utf8');
      return { path: target, format: ext, rows: rows.length };
    }

    if (session.kind === 'doc' && (ext === 'txt' || ext === 'md' || ext === 'html')) {
      const frame = session.engine.render();
      const blocks = (frame.blocks || []).map((b) => ({
        type: b.style && /heading/i.test(b.style) ? 'heading' : 'paragraph',
        level: Number(String(b.style || '').replace(/\D/g, '')) || 1,
        text: (b.runs || []).map((r) => r.text).join('') || b.text || '',
        runs: b.runs,
      }));
      const text = ext === 'md' ? writeMarkdown(blocks) : ext === 'txt' ? writePlain(blocks) : `<!doctype html>\n<meta charset="utf-8">\n<title>${session.name}</title>\n${blocks.map((b) => (b.type === 'heading' ? `<h${b.level}>${b.text}</h${b.level}>` : `<p>${b.text}</p>`)).join('\n')}\n`;
      fs.writeFileSync(target, text, 'utf8');
      return { path: target, format: ext };
    }

    throw new Error(`Rutba Office cannot export this document as ${ext.toUpperCase()} yet.`);
  }
}
