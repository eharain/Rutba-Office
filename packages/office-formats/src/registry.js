// One list of what the suite opens, used by four things that otherwise drift:
// the Open dialog's filters, the installer's file associations, the launcher's
// "open with" menu, and the Save As type list.
//
// Keeping them in one table is the difference between a suite that claims to
// open .odt on its download page and one that actually offers it in the dialog.

import { KINDS } from './sniff.js';

/**
 * Where the product lives. One definition, because these appear in the About
 * dialog, the Help menu, the launcher's footer, the installer metadata and the
 * README, and a stale URL in any of them is a support request.
 */
export const SITE = {
  home: 'https://office.rutba.io',
  contact: 'https://office.rutba.io/contact',
  help: 'https://office.rutba.io/help',
  source: 'https://github.com/eharain/Rutba-Office',
  releases: 'https://github.com/eharain/Rutba-Office/releases',
  company: 'https://tech-style.co',
};

/** @typedef {'mail'|'word'|'sheets'|'slides'|'pictures'|'image'|'video'} AppKey */

export const APPS = {
  mail: {
    key: 'mail',
    name: 'Rutba Mail',
    short: 'Mail',
    tagline: 'Every account, every archive, one inbox',
    accent: '#3b7de0',
    icon: 'mail',
  },
  word: {
    key: 'word',
    name: 'Rutba Word',
    short: 'Word',
    tagline: 'Documents that open the same everywhere',
    accent: '#2b5fd9',
    icon: 'word',
  },
  sheets: {
    key: 'sheets',
    name: 'Rutba Worksheets',
    short: 'Worksheets',
    tagline: 'Real formulas, real recalculation',
    accent: '#0f9d58',
    icon: 'sheets',
  },
  slides: {
    key: 'slides',
    name: 'Rutba Presentation',
    short: 'Presentation',
    tagline: 'Slides that survive the round trip',
    accent: '#d9534f',
    icon: 'slides',
  },
  pictures: {
    key: 'pictures',
    name: 'Rutba Pictures',
    short: 'Pictures',
    tagline: 'A viewer that opens before you blink',
    accent: '#7b5cd6',
    icon: 'pictures',
  },
  image: {
    key: 'image',
    name: 'Rutba Image',
    short: 'Image',
    tagline: 'Crop, correct, annotate, export',
    accent: '#e08b2b',
    icon: 'image',
  },
  video: {
    key: 'video',
    name: 'Rutba Video',
    short: 'Video',
    tagline: 'Trim and export without a render farm',
    accent: '#c2408f',
    icon: 'video',
  },
};

/** Extensions each app claims, in the order a dialog should show them. */
export const APP_EXTENSIONS = {
  // `markdown` alongside `md`: GitHub accepts both for a README, and an editor
  // that claims one and not the other is the default for half your files.
  word: ['docx', 'docm', 'dotx', 'doc', 'odt', 'rtf', 'txt', 'md', 'markdown', 'html', 'htm'],
  sheets: ['xlsx', 'xlsm', 'xltx', 'xls', 'ods', 'csv', 'tsv'],
  slides: ['pptx', 'pptm', 'potx', 'ppsx', 'ppt', 'odp'],
  pictures: ['jpg', 'jpeg', 'png', 'gif', 'webp', 'avif', 'heic', 'bmp', 'tif', 'tiff', 'ico', 'svg', 'pdf'],
  image: ['jpg', 'jpeg', 'png', 'webp', 'bmp', 'gif', 'avif'],
  video: ['mp4', 'm4v', 'mov', 'webm', 'mkv', 'ogv', 'mp3', 'm4a', 'wav', 'flac', 'ogg'],
  mail: ['eml', 'emlx', 'msg', 'mbox', 'pst', 'ost', 'olm', 'vcf', 'ics'],
};

/** Formats we can write. Everything else is read-only, and says so. */
export const WRITABLE = {
  word: [
    { kind: 'docx', label: 'Word Document (.docx)' },
    { kind: 'pdf', label: 'PDF Document (.pdf)' },
    { kind: 'txt', label: 'Plain Text (.txt)' },
    { kind: 'md', label: 'Markdown (.md)' },
    { kind: 'html', label: 'Web Page (.html)' },
  ],
  sheets: [
    { kind: 'xlsx', label: 'Worksheet (.xlsx)' },
    { kind: 'csv', label: 'Comma Separated Values (.csv)' },
    { kind: 'tsv', label: 'Tab Separated Values (.tsv)' },
    { kind: 'pdf', label: 'PDF Document (.pdf)' },
  ],
  slides: [
    { kind: 'pptx', label: 'Presentation (.pptx)' },
    { kind: 'pdf', label: 'PDF Document (.pdf)' },
  ],
  image: [
    { kind: 'png', label: 'PNG Image (.png)' },
    { kind: 'jpeg', label: 'JPEG Image (.jpg)' },
    { kind: 'webp', label: 'WebP Image (.webp)' },
  ],
  video: [
    { kind: 'webm', label: 'WebM Video (.webm)' },
    { kind: 'mp4', label: 'MP4 Video (.mp4)' },
  ],
  mail: [
    { kind: 'eml', label: 'Email Message (.eml)' },
    { kind: 'mbox', label: 'Mbox Archive (.mbox)' },
  ],
};

const TITLES = {
  word: 'Documents',
  sheets: 'Worksheets',
  slides: 'Presentations',
  pictures: 'Pictures and PDFs',
  image: 'Images',
  video: 'Video and audio',
  mail: 'Mail files',
};

/** Electron `dialog.showOpenDialog` filters for one app, or for everything. */
export function openFilters(appKey) {
  if (appKey && APP_EXTENSIONS[appKey]) {
    return [
      { name: TITLES[appKey], extensions: APP_EXTENSIONS[appKey] },
      { name: 'All files', extensions: ['*'] },
    ];
  }
  const all = [...new Set(Object.values(APP_EXTENSIONS).flat())];
  return [
    { name: 'All supported files', extensions: all },
    ...Object.entries(APP_EXTENSIONS).map(([k, exts]) => ({ name: TITLES[k], extensions: exts })),
    { name: 'All files', extensions: ['*'] },
  ];
}

/** Filters for Save As, given the app doing the saving. */
export function saveFilters(appKey) {
  return (WRITABLE[appKey] || []).map((w) => ({
    name: w.label,
    extensions: [KINDS[w.kind]?.ext.replace('.', '') || w.kind],
  }));
}

/**
 * Extensions that are a second name for a kind rather than a kind of their own.
 * Mirrors the alias list in sniff.js, which is where a file's identity is
 * decided; this is only about what the operating system should call it.
 */
const EXTENSION_ALIASES = {
  markdown: 'md',
  htm: 'html',
  jpeg: 'jpeg',
  jpe: 'jpeg',
  tif: 'tiff',
  m4v: 'mp4',
  emlx: 'eml',
  mbx: 'mbox',
};

/**
 * File associations for electron-builder. Every extension the suite can open
 * is claimed, so double-clicking works for the formats we advertise — and only
 * for those.
 */
export function fileAssociations() {
  const seen = new Set();
  const out = [];
  for (const [appKey, exts] of Object.entries(APP_EXTENSIONS)) {
    if (appKey === 'image') continue; // Pictures already claims these
    for (const ext of exts) {
      if (seen.has(ext)) continue;
      seen.add(ext);
      // A kind's canonical extension, or the one it is also known by: `.markdown`
      // is `md`, `.htm` is `html`, and both need a real name and MIME type or
      // the operating system offers to open them with "MARKDOWN file".
      const kind = Object.entries(KINDS).find(([, d]) => d.ext === `.${ext}`)?.[0] ?? EXTENSION_ALIASES[ext];
      out.push({
        ext,
        name: KINDS[kind]?.label || `${ext.toUpperCase()} file`,
        description: KINDS[kind]?.label || `${ext.toUpperCase()} file`,
        mimeType: KINDS[kind]?.mime || 'application/octet-stream',
        role: 'Editor',
      });
    }
  }
  return out;
}

/** The template list the launcher offers under New. */
export const NEW_DOCUMENTS = [
  { app: 'word', kind: 'docx', label: 'Blank document', template: 'blank' },
  { app: 'word', kind: 'docx', label: 'Letter', template: 'letter' },
  { app: 'word', kind: 'docx', label: 'Report', template: 'report' },
  { app: 'sheets', kind: 'xlsx', label: 'Blank worksheet', template: 'blank' },
  { app: 'sheets', kind: 'xlsx', label: 'Budget', template: 'budget' },
  { app: 'sheets', kind: 'xlsx', label: 'Invoice', template: 'invoice' },
  { app: 'slides', kind: 'pptx', label: 'Blank presentation', template: 'blank' },
  { app: 'slides', kind: 'pptx', label: 'Pitch deck', template: 'pitch' },
];
