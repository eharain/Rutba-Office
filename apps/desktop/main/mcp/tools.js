// What an AI app may ask Rutba Office to do (D-2026-10-16-07): make and read
// documents, workbooks and presentations, start from a template, make a PDF
// of a document, put pages of PDFs together, and read the words in a picture.
// Built on the engines the windows use, so a file made here opens in Rutba
// Office, and in Word, Excel and PowerPoint, as one made by hand does.
//
// Every call is asked first whether the tools are on (access.js openNow), and
// works only in the folders the person gave them, each with its own step
// (reach.js): read only, or read and make new files. Nothing is overwritten
// or deleted: a new file must have a name of its own, so there is nothing to
// undo and nothing to confirm. Every call goes in the log.
//
// Mail, calendars, address books and Data Files databases are reached
// through the running Rutba Office of the profile (reach-tools.js, through
// channel.js), each asked its own setting.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { buildXlsx, readWorkbook } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation/build';
import { compose, para, table } from '@rutba/templates/word';
import { TEMPLATES, templateById } from '@rutba/templates/catalogue';
import { themeById } from '@rutba/office-formats/themes';
import { makeTemplate } from '@rutba/templates/make';
import { openPdf } from '@rutba/pdf/reader';
import { assemble, parsePageRange } from '@rutba/pdf/assemble';
import { openNow, foldersOf, resolveIn, newFileIn, logCall, FOLDERS_KEY } from './access.js';
import { folderEntries } from './reach.js';
import { REACH_TOOL_LIST, REACH_TOOLS } from './reach-tools.js';

const KINDS = {
  '.docx': 'document', '.xlsx': 'workbook', '.pptx': 'presentation', '.pdf': 'pdf',
  '.png': 'picture', '.jpg': 'picture', '.jpeg': 'picture', '.txt': 'text', '.md': 'text', '.csv': 'text',
};
const MAX_CHARS = 60000;

const text = (t) => ({ content: [{ type: 'text', text: String(t) }] });
const refusal = (t) => ({ content: [{ type: 'text', text: String(t) }], isError: true });
const decodeXml = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(Number(n))).replace(/&amp;/g, '&');
const trimTo = (s, max) => (s.length > max ? `${s.slice(0, max)}\n\n[… cut at ${max} characters of ${s.length}]` : s);
const writeNew = (file, bytes) => {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, bytes, { flag: 'wx' });
};

/* ── reading ─────────────────────────────────────────────────────────── */

/** A document's words, a paragraph a line, headings marked with #, tables a row a line. */
export function documentText(bytes) {
  const pkg = OoxmlPackage.read(Buffer.from(bytes));
  const xml = pkg.read('word/document.xml').toString('utf8');
  const words = (x) => decodeXml([...x.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join(''));
  const lines = [];
  for (const m of xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>|<w:p\b[\s\S]*?<\/w:p>/g)) {
    const block = m[0];
    if (block.startsWith('<w:tbl')) {
      for (const row of block.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)) {
        lines.push([...row[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((c) => words(c[0]).trim()).join(' | '));
      }
      continue;
    }
    const style = /<w:pStyle w:val="([^"]+)"/.exec(block)?.[1] || '';
    const t = words(block);
    if (!t.trim()) continue;
    const level = style === 'Title' ? 1 : Number(/^Heading(\d)$/.exec(style)?.[1] || 0) + (/^Heading/.test(style) ? 1 : 0);
    lines.push(level ? `${'#'.repeat(Math.min(6, level))} ${t}` : /<w:numPr>/.test(block) ? `- ${t}` : t);
  }
  return lines.join('\n');
}

/** A presentation's words, slide by slide. */
export function presentationText(bytes) {
  const pkg = OoxmlPackage.read(Buffer.from(bytes));
  const slides = [...pkg.byName.keys()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => Number(/(\d+)\.xml$/.exec(a)[1]) - Number(/(\d+)\.xml$/.exec(b)[1]));
  return slides.map((name, i) => {
    const xml = pkg.read(name).toString('utf8');
    const paras = [...xml.matchAll(/<a:p>[\s\S]*?<\/a:p>|<a:p\s[\s\S]*?<\/a:p>/g)].map((m) => decodeXml([...m[0].matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g)].map((t) => t[1]).join(''))).filter((t) => t.trim());
    return `Slide ${i + 1}:\n${paras.join('\n')}`;
  }).join('\n\n');
}

/** A workbook's values, sheet by sheet, a row a line, cells by tabs. */
export function workbookText(bytes, { maxRows = 500 } = {}) {
  const book = readWorkbook(Buffer.from(bytes));
  return book.sheets.map((s) => `Sheet ${s.name}:\n${(s.rows || []).slice(0, maxRows).map((r) => (r || []).map((v) => (v == null ? '' : String(v))).join('\t')).join('\n')}`).join('\n\n');
}

/** A PDF's words, page by page. */
export function pdfText(bytes) {
  const file = openPdf(Buffer.from(bytes));
  const out = [];
  for (let i = 0; i < file.pageCount; i++) out.push(`Page ${i + 1}:\n${file.pageText(i) || ''}`);
  return out.join('\n\n');
}

/* ── making ──────────────────────────────────────────────────────────── */

const WIDTH = 9026; // the text width of an A4 page with Word's margins, in twentieths of a point

/** A document from blocks: headings, paragraphs, bulleted and numbered lists, tables. */
export function documentFrom({ title = '', blocks = [], theme = 'rutba' }) {
  const pieces = [];
  const lists = { bullet: [], number: [] };
  if (title) pieces.push(para(String(title), { style: 'Title' }));
  for (const b of blocks || []) {
    if (!b || typeof b !== 'object') continue;
    if (b.type === 'heading') pieces.push(para(String(b.text || ''), { style: Number(b.level) >= 2 ? 'Heading2' : 'Heading1' }));
    else if (b.type === 'paragraph') pieces.push(para(String(b.text || '')));
    else if (b.type === 'bullets' || b.type === 'numbered') {
      for (const item of b.items || []) {
        lists[b.type === 'bullets' ? 'bullet' : 'number'].push(pieces.length);
        pieces.push(para(String(item)));
      }
    } else if (b.type === 'table' && Array.isArray(b.rows) && b.rows.length) {
      const cols = Math.max(...b.rows.map((r) => (Array.isArray(r) ? r.length : 0)), 1);
      const widths = Array.from({ length: cols }, () => Math.floor(WIDTH / cols));
      const rows = b.rows.map((r) => Array.from({ length: cols }, (_, i) => String((Array.isArray(r) ? r[i] : '') ?? '')));
      const header = b.header !== false;
      pieces.push(table(rows, { widths, borders: 'grid', header, headerFill: header ? (themeById(theme)?.palette?.accent1 || '2B5FD9') : null }));
      pieces.push(para(''));
    }
  }
  if (!pieces.length) pieces.push(para(''));
  return compose({ theme, pieces, lists: Object.fromEntries(Object.entries(lists).filter(([, v]) => v.length)) });
}

/** A presentation: a title slide, then a slide of points each. */
export function presentationFrom({ title = '', subtitle = '', slides = [] }) {
  const out = [];
  if (title) out.push({ layout: 'title', title: String(title), body: String(subtitle || '') });
  for (const s of slides || []) out.push({ layout: 'obj', title: String(s?.title || ''), body: (s?.points || []).map(String) });
  return buildPptx({ title: String(title || out[0]?.title || 'Presentation'), slides: out });
}

/* ── the tools ───────────────────────────────────────────────────────── */

const FILE = { type: 'string', description: 'A file name or path inside one of the folders the person gave the AI tools. A name alone goes in the first folder.' };

export const TOOL_LIST = [
  {
    name: 'list_folders',
    description: 'The folders on this computer the person has let the Rutba Office AI tools use. Every file the tools read or make is in one of them.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
  },
  {
    name: 'list_files',
    description: 'The documents, workbooks, presentations, PDFs, pictures and text files in a folder the tools may use (and the folders one level below), with their kind, size and when they last changed.',
    inputSchema: { type: 'object', properties: { folder: { type: 'string', description: 'One of the folders, or a folder inside one. The first folder when left out.' } }, additionalProperties: false },
  },
  {
    name: 'read_file',
    description: 'The words of a document (.docx: headings marked with #, list items with -, table rows with |), a workbook (.xlsx: each sheet\'s values, cells separated by tabs), a presentation (.pptx: slide by slide), a PDF (page by page) or a text file.',
    inputSchema: { type: 'object', properties: { path: FILE }, required: ['path'], additionalProperties: false },
  },
  {
    name: 'create_document',
    description: 'Make a new Word document (.docx) with a title, headings, paragraphs, bulleted and numbered lists and tables, in a designed theme. Never overwrites: the name must be new.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { ...FILE, description: `${FILE.description} Must end in .docx and not exist yet.` },
        title: { type: 'string' },
        theme: { type: 'string', description: 'A Rutba Office theme: rutba, harbour, meadow, ember, paperkite, basalt, saffron, orchard or studio. rutba when left out.' },
        blocks: {
          type: 'array',
          description: 'The body, in order.',
          items: {
            type: 'object',
            properties: {
              type: { type: 'string', enum: ['heading', 'paragraph', 'bullets', 'numbered', 'table'] },
              text: { type: 'string', description: 'For heading and paragraph.' },
              level: { type: 'integer', minimum: 1, maximum: 2, description: 'For heading: 1 or 2.' },
              items: { type: 'array', items: { type: 'string' }, description: 'For bullets and numbered.' },
              rows: { type: 'array', items: { type: 'array', items: { type: ['string', 'number'] } }, description: 'For table: rows of cells, the first row the header.' },
              header: { type: 'boolean', description: 'For table: whether the first row is shaded as a header (yes when left out).' },
            },
            required: ['type'],
          },
        },
      },
      required: ['path', 'blocks'],
      additionalProperties: false,
    },
  },
  {
    name: 'create_workbook',
    description: 'Make a new Excel workbook (.xlsx) from sheets of rows. A cell that starts with = is a formula (for example =SUM(B2:B9)). Never overwrites: the name must be new.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { ...FILE, description: `${FILE.description} Must end in .xlsx and not exist yet.` },
        sheets: {
          type: 'array',
          items: {
            type: 'object',
            properties: { name: { type: 'string' }, rows: { type: 'array', items: { type: 'array', items: { type: ['string', 'number', 'boolean', 'null'] } } } },
            required: ['name', 'rows'],
          },
        },
      },
      required: ['path', 'sheets'],
      additionalProperties: false,
    },
  },
  {
    name: 'create_presentation',
    description: 'Make a new PowerPoint presentation (.pptx): a title slide, then a slide for each item with its title and points. Never overwrites: the name must be new.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { ...FILE, description: `${FILE.description} Must end in .pptx and not exist yet.` },
        title: { type: 'string' },
        subtitle: { type: 'string' },
        slides: { type: 'array', items: { type: 'object', properties: { title: { type: 'string' }, points: { type: 'array', items: { type: 'string' } } }, required: ['title'] } },
      },
      required: ['path', 'slides'],
      additionalProperties: false,
    },
  },
  {
    name: 'list_templates',
    description: 'The templates Rutba Office comes with: letters, a CV, a report, invoices, budgets, a timesheet and others, and decks.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
  },
  {
    name: 'create_from_template',
    description: 'Make a new file from one of Rutba Office\'s templates (see list_templates), to be filled in. Never overwrites: the name must be new.',
    inputSchema: { type: 'object', properties: { template: { type: 'string', description: 'The template\'s id, for example word-letter or sheet-budget.' }, path: FILE }, required: ['template', 'path'], additionalProperties: false },
  },
  {
    name: 'export_pdf',
    description: 'Make a PDF of a Word document (.docx), laid out as Rutba Office prints it. Never overwrites: the PDF\'s name must be new.',
    inputSchema: { type: 'object', properties: { path: { ...FILE, description: 'The .docx to make a PDF of.' }, to: { ...FILE, description: 'The new PDF, ending in .pdf.' } }, required: ['path', 'to'], additionalProperties: false },
  },
  {
    name: 'combine_pdf_pages',
    description: 'Make a new PDF from pages of one or more PDFs, in the order given: to put PDFs together, take some pages out, reorder them or turn them. Never overwrites: the new PDF\'s name must be new.',
    inputSchema: {
      type: 'object',
      properties: {
        files: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              path: FILE,
              pages: { type: 'string', description: 'The pages to take, like 1-3,5 or 4-; all of them when left out.' },
              rotate: { type: 'integer', enum: [0, 90, 180, 270], description: 'Turn the pages taken clockwise by this many degrees.' },
            },
            required: ['path'],
          },
        },
        to: { ...FILE, description: 'The new PDF, ending in .pdf.' },
        title: { type: 'string' },
      },
      required: ['files', 'to'],
      additionalProperties: false,
    },
  },
  {
    name: 'read_text_in_picture',
    description: 'Read the words in a PNG picture, on this computer, with Rutba Office\'s text recognition. Languages: eng (English), urd (Urdu, beta), ara (Arabic), hin (Hindi), chi_sim (Chinese), jpn (Japanese).',
    inputSchema: { type: 'object', properties: { path: FILE, languages: { type: 'array', items: { type: 'string' }, description: 'eng when left out.' } }, required: ['path'], additionalProperties: false },
  },
];

/** The words of a document's bytes, by kind: what an attachment is read with too (reach-tools.js). */
export const READERS = { docx: documentText, xlsx: workbookText, pptx: presentationText, pdf: pdfText };

/** What an AI app is told when a tool needs the running Rutba Office and there is none. */
const NOT_OPEN = 'Rutba Office is not open in this profile. The AI tools reach mail, calendars, address books and Data Files databases through it: open Rutba Office, then ask again.';
const NOT_OPEN_TO_CONFIRM = 'This needs the person to see it and press a button in Rutba Office, which is not open in this profile, so nothing was done and nothing waits. Open Rutba Office, then ask again.';
const CONFIRMED = new Set(['send_message', 'send_invitation', 'answer_invitation', 'confirmation_answer']);

/**
 * The tools, answering for real: `settings()` reads the settings as saved,
 * afresh for each call; `dataDir` is where the log is kept; `ocr` the text
 * recognition service (ocr.js) and `pictureToPng(file)` a picture's PNG
 * bytes; `unicodeFont()` and `renderPdf` the document PDF writer's needs.
 * `reach` answers the tools for mail, calendars, address books and
 * databases (reach-tools.js): in the running copy itself, or through its
 * channel (node.js), throwing { notRunning } when there is none.
 * `locked()` answers a refusal while the profile has a password and is not
 * open and unlocked in a running Rutba Office, or null: then nothing at all
 * is reached, files included.
 */
export function createTools({ settings, dataDir, ocr = null, pictureToPng = null, renderPdf = null, openDocx = null, unicodeFont = () => null, isAlive, reach = null, locked = null } = {}) {
  const handlers = {
    list_folders: async (_args, folders, made, steps) => text(steps.length ? steps.map((f) => `${f.path}\t${f.step === 'prepare' ? 'read and make new files' : 'read only'}`).join('\n') : 'No folder is set for the AI tools.'),

    list_files: async ({ folder }, folders) => {
      const root = folder ? resolveIn(folders, folder) : folders[0];
      const found = [];
      const walk = (dir, depth) => {
        let entries = [];
        try {
          entries = fs.readdirSync(dir, { withFileTypes: true });
        } catch {
          return;
        }
        for (const e of entries) {
          if (found.length >= 300) return;
          const p = path.join(dir, e.name);
          if (e.isDirectory() && depth < 1 && !e.name.startsWith('.')) walk(p, depth + 1);
          else if (e.isFile() && KINDS[path.extname(e.name).toLowerCase()]) {
            const st = fs.statSync(p);
            found.push(`${KINDS[path.extname(e.name).toLowerCase()]}\t${st.size}\t${st.mtime.toISOString()}\t${p}`);
          }
        }
      };
      walk(root, 0);
      return text(found.length ? `kind\tbytes\tchanged\tpath\n${found.join('\n')}` : `No documents, workbooks, presentations, PDFs, pictures or text files in ${root}.`);
    },

    read_file: async ({ path: given }, folders) => {
      const file = resolveIn(folders, given);
      const ext = path.extname(file).toLowerCase();
      const bytes = fs.readFileSync(file);
      let out;
      if (ext === '.docx') out = documentText(bytes);
      else if (ext === '.xlsx') out = workbookText(bytes);
      else if (ext === '.pptx') out = presentationText(bytes);
      else if (ext === '.pdf') out = pdfText(bytes);
      else if (['.txt', '.md', '.csv'].includes(ext)) out = bytes.toString('utf8');
      else return refusal(`read_file reads .docx, .xlsx, .pptx, .pdf, .txt, .md and .csv files; ${path.basename(file)} is none of them.`);
      return text(trimTo(out, MAX_CHARS));
    },

    create_document: async ({ path: given, title, theme, blocks }, folders, made) => {
      const file = made(given, ['.docx']);
      writeNew(file, documentFrom({ title, blocks, theme: theme || 'rutba' }));
      return text(`Made ${file}.`);
    },

    create_workbook: async ({ path: given, sheets }, folders, made) => {
      const file = made(given, ['.xlsx']);
      const clean = (sheets || []).map((s, i) => ({ name: String(s?.name || `Sheet${i + 1}`).slice(0, 31), rows: (s?.rows || []).map((r) => (Array.isArray(r) ? r : [r])) }));
      if (!clean.length) return refusal('A workbook needs at least one sheet.');
      writeNew(file, buildXlsx({ sheets: clean }));
      return text(`Made ${file}.`);
    },

    create_presentation: async ({ path: given, title, subtitle, slides }, folders, made) => {
      const file = made(given, ['.pptx']);
      writeNew(file, presentationFrom({ title, subtitle, slides }));
      return text(`Made ${file}.`);
    },

    list_templates: async () => text(TEMPLATES.map((t) => `${t.id}\t${t.ext}\t${t.name}: ${t.description}`).join('\n')),

    create_from_template: async ({ template, path: given }, folders, made) => {
      const entry = templateById(String(template || ''));
      if (!entry) return refusal(`There is no template ${template}. list_templates names them.`);
      const file = made(given, [`.${entry.ext}`]);
      writeNew(file, makeTemplate(entry.id));
      return text(`Made ${file} from the ${entry.name} template.`);
    },

    export_pdf: async ({ path: given, to }, folders, made) => {
      if (!renderPdf || !openDocx) return refusal('This copy of Rutba Office cannot make PDFs here.');
      const from = resolveIn(folders, given);
      if (path.extname(from).toLowerCase() !== '.docx') return refusal('export_pdf makes a PDF of a .docx document.');
      const file = made(to, ['.pdf']);
      const { buffer, pages } = renderPdf(openDocx(fs.readFileSync(from)), { title: path.basename(from, '.docx'), unicodeFont: unicodeFont() }) || {};
      if (!buffer) return refusal(`${path.basename(from)} could not be laid out as a PDF.`);
      writeNew(file, Buffer.from(buffer));
      return text(`Made ${file}${pages ? `, ${pages} ${pages === 1 ? 'page' : 'pages'}` : ''}.`);
    },

    combine_pdf_pages: async ({ files, to, title }, folders, made) => {
      if (!Array.isArray(files) || !files.length) return refusal('Name at least one PDF.');
      const items = [];
      for (const f of files) {
        const source = resolveIn(folders, f?.path);
        const pdf = openPdf(fs.readFileSync(source));
        const pages = f?.pages ? parsePageRange(String(f.pages), pdf.pageCount) : Array.from({ length: pdf.pageCount }, (_, i) => i + 1);
        for (const n of pages) items.push({ file: pdf, page: Number(n) - 1, rotate: Number(f?.rotate) || 0 });
      }
      if (!items.length) return refusal('Those page numbers take no pages.');
      const file = made(to, ['.pdf']);
      writeNew(file, Buffer.from(assemble(items, { title: title || path.basename(file, '.pdf') })));
      return text(`Made ${file}, ${items.length} ${items.length === 1 ? 'page' : 'pages'}.`);
    },

    read_text_in_picture: async ({ path: given, languages }, folders) => {
      if (!ocr || !pictureToPng) return refusal('This copy of Rutba Office cannot read pictures here.');
      const file = resolveIn(folders, given);
      if (path.extname(file).toLowerCase() !== '.png') return refusal('read_text_in_picture reads PNG pictures; save the picture as a PNG first.');
      const png = await pictureToPng(file);
      const result = await ocr.recognize({ image: png, languages: Array.isArray(languages) && languages.length ? languages : ['eng'] });
      return text(result?.text?.trim() || 'No words were found in the picture.');
    },
  };

  /** Paths a call named, for the log: never a file's contents. */
  const pathsOf = (args) => [args?.path, args?.to, args?.folder, ...(Array.isArray(args?.files) ? args.files.map((f) => f?.path) : [])].filter((p) => typeof p === 'string');
  /** Files a reach call attached, for the log; the thing it used is logged beside them. */
  const attachedOf = (args) => (Array.isArray(args?.attachments) ? args.attachments.filter((p) => typeof p === 'string') : []);

  /** A reach tool: asked of the running copy, its answer and the thing and step it used for the log. */
  async function reachCall(name, args, entry) {
    if (!reach) return { result: refusal(NOT_OPEN) };
    try {
      const { result, log } = await reach.call(name, args);
      if (log) Object.assign(entry, { kind: log.kind, thing: String(log.name || '').slice(0, 200), step: log.step });
      return { result };
    } catch (err) {
      if (err?.notRunning) return { result: refusal(CONFIRMED.has(name) ? NOT_OPEN_TO_CONFIRM : NOT_OPEN) };
      return { result: refusal(err?.message || String(err)) };
    }
  }

  return {
    list: () => [...TOOL_LIST, ...REACH_TOOL_LIST].map((t) => ({ ...t })),
    /** One call; `ctx.where` is where it came from when not this computer (an address on the network). */
    async call(name, args = {}, ctx = {}) {
      const started = Date.now();
      const isReach = REACH_TOOLS.has(name);
      const entry = { at: new Date(started).toISOString(), tool: String(name), paths: isReach ? attachedOf(args) : pathsOf(args), ...(ctx.where ? { where: String(ctx.where).slice(0, 64) } : {}) };
      const done = (result) => {
        logCall(dataDir, { ...entry, ok: !result.isError, ms: Date.now() - started, ...(result.isError ? { error: result.content?.[0]?.text?.slice(0, 300) } : {}) });
        return result;
      };
      const s = settings();
      const gate = openNow(s, isAlive ? { isAlive } : undefined);
      if (!gate.open) return done(refusal(gate.reason));
      // A profile with a password: nothing is reached while it is not open and unlocked.
      const shut = locked ? await locked() : null;
      if (shut) return done(refusal(shut));
      if (isReach) return done((await reachCall(name, args || {}, entry)).result);
      const handler = handlers[name];
      if (!handler) return done(refusal(`Rutba Office has no tool called ${name}.`));
      const read = foldersOf(s);
      const write = foldersOf(s, 'prepare');
      const made = (given, extensions) => newFileIn(write, given, extensions, read.filter((f) => !write.includes(f)));
      // Every folder a call uses is one it may read; one only read is listed, and refused for new files.
      if (name !== 'list_folders' && name !== 'list_templates' && !read.length) return done(refusal('No folder has been given to the AI tools yet. In Rutba Office, choose AI Tools from the app menu and add a folder.'));
      try {
        return done(await handler(args || {}, read, made, folderEntries(s[FOLDERS_KEY]).filter((f) => f.step !== 'off')));
      } catch (err) {
        return done(refusal(err?.message || String(err)));
      }
    },
  };
}
