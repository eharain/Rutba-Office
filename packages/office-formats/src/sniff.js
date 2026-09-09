// What is this file, actually?
//
// Extensions lie. A .doc that is really RTF, a .xls that is really a tab-
// separated export, a .jpg that a phone wrote as HEIC — every office suite
// meets these on its first day, and the ones that trust the extension show the
// user a corrupt-file dialog for a file that is perfectly fine.
//
// So: magic bytes first, container inspection second, extension only as a
// tie-breaker. `sniff()` needs the first 4 KB; container formats (zip, CFB) get
// a second look with the whole buffer when one is available.

const dec = new TextDecoder('latin1');

function bytes(input) {
  if (input instanceof Uint8Array) return input;
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  throw new TypeError('sniff() wants bytes');
}

function starts(b, sig, at = 0) {
  if (b.length < at + sig.length) return false;
  for (let i = 0; i < sig.length; i++) if (b[at + i] !== sig[i]) return false;
  return true;
}

function ascii(b, at, len) {
  return dec.decode(b.subarray(at, at + len));
}

const ZIP = [0x50, 0x4b, 0x03, 0x04];
const CFB = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const PNG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const GIF = [0x47, 0x49, 0x46, 0x38];
const BMP = [0x42, 0x4d];
const PDF = [0x25, 0x50, 0x44, 0x46];
const EBML = [0x1a, 0x45, 0xdf, 0xa3];
const OGG = [0x4f, 0x67, 0x67, 0x53];
const FLAC = [0x66, 0x4c, 0x61, 0x43];
const PST = [0x21, 0x42, 0x44, 0x4e]; // "!BDN"
const RAR = [0x52, 0x61, 0x72, 0x21];
const SEVENZ = [0x37, 0x7a, 0xbc, 0xaf, 0x27, 0x1c];
const GZ = [0x1f, 0x8b];
const ICO = [0x00, 0x00, 0x01, 0x00];
const PSD = [0x38, 0x42, 0x50, 0x53];

/** Everything the suite knows how to open, and which app owns it. */
export const KINDS = {
  docx: { app: 'word', label: 'Word Document', ext: '.docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', edit: true },
  docm: { app: 'word', label: 'Word Macro-Enabled Document', ext: '.docm', mime: 'application/vnd.ms-word.document.macroEnabled.12', edit: true },
  dotx: { app: 'word', label: 'Word Template', ext: '.dotx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template', edit: true },
  doc: { app: 'word', label: 'Word 97-2003 Document', ext: '.doc', mime: 'application/msword', edit: false },
  odt: { app: 'word', label: 'OpenDocument Text', ext: '.odt', mime: 'application/vnd.oasis.opendocument.text', edit: true },
  rtf: { app: 'word', label: 'Rich Text Format', ext: '.rtf', mime: 'application/rtf', edit: false },
  txt: { app: 'word', label: 'Plain Text', ext: '.txt', mime: 'text/plain', edit: true },
  md: { app: 'word', label: 'Markdown', ext: '.md', mime: 'text/markdown', edit: true },
  html: { app: 'word', label: 'Web Page', ext: '.html', mime: 'text/html', edit: false },

  xlsx: { app: 'sheets', label: 'Worksheet', ext: '.xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', edit: true },
  xlsm: { app: 'sheets', label: 'Macro-Enabled Worksheet', ext: '.xlsm', mime: 'application/vnd.ms-excel.sheet.macroEnabled.12', edit: true },
  xltx: { app: 'sheets', label: 'Worksheet Template', ext: '.xltx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template', edit: true },
  xls: { app: 'sheets', label: 'Excel 97-2003 Worksheet', ext: '.xls', mime: 'application/vnd.ms-excel', edit: false },
  ods: { app: 'sheets', label: 'OpenDocument Spreadsheet', ext: '.ods', mime: 'application/vnd.oasis.opendocument.spreadsheet', edit: true },
  csv: { app: 'sheets', label: 'Comma Separated Values', ext: '.csv', mime: 'text/csv', edit: true },
  tsv: { app: 'sheets', label: 'Tab Separated Values', ext: '.tsv', mime: 'text/tab-separated-values', edit: true },

  pptx: { app: 'slides', label: 'Presentation', ext: '.pptx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', edit: true },
  pptm: { app: 'slides', label: 'Macro-Enabled Presentation', ext: '.pptm', mime: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12', edit: true },
  potx: { app: 'slides', label: 'Presentation Template', ext: '.potx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.template', edit: true },
  ppsx: { app: 'slides', label: 'Presentation Slideshow', ext: '.ppsx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow', edit: true },
  ppt: { app: 'slides', label: 'PowerPoint 97-2003 Presentation', ext: '.ppt', mime: 'application/vnd.ms-powerpoint', edit: false },
  odp: { app: 'slides', label: 'OpenDocument Presentation', ext: '.odp', mime: 'application/vnd.oasis.opendocument.presentation', edit: true },

  png: { app: 'pictures', label: 'PNG Image', ext: '.png', mime: 'image/png', edit: true },
  jpeg: { app: 'pictures', label: 'JPEG Image', ext: '.jpg', mime: 'image/jpeg', edit: true },
  gif: { app: 'pictures', label: 'GIF Image', ext: '.gif', mime: 'image/gif', edit: true },
  webp: { app: 'pictures', label: 'WebP Image', ext: '.webp', mime: 'image/webp', edit: true },
  avif: { app: 'pictures', label: 'AVIF Image', ext: '.avif', mime: 'image/avif', edit: true },
  heic: { app: 'pictures', label: 'HEIC Image', ext: '.heic', mime: 'image/heic', edit: false },
  bmp: { app: 'pictures', label: 'Bitmap Image', ext: '.bmp', mime: 'image/bmp', edit: true },
  tiff: { app: 'pictures', label: 'TIFF Image', ext: '.tiff', mime: 'image/tiff', edit: false },
  ico: { app: 'pictures', label: 'Icon', ext: '.ico', mime: 'image/x-icon', edit: false },
  svg: { app: 'pictures', label: 'SVG Image', ext: '.svg', mime: 'image/svg+xml', edit: false },
  psd: { app: 'pictures', label: 'Photoshop Document', ext: '.psd', mime: 'image/vnd.adobe.photoshop', edit: false },

  mp4: { app: 'video', label: 'MP4 Video', ext: '.mp4', mime: 'video/mp4', edit: true },
  mov: { app: 'video', label: 'QuickTime Video', ext: '.mov', mime: 'video/quicktime', edit: true },
  webm: { app: 'video', label: 'WebM Video', ext: '.webm', mime: 'video/webm', edit: true },
  mkv: { app: 'video', label: 'Matroska Video', ext: '.mkv', mime: 'video/x-matroska', edit: false },
  ogv: { app: 'video', label: 'Ogg Video', ext: '.ogv', mime: 'video/ogg', edit: false },
  mp3: { app: 'video', label: 'MP3 Audio', ext: '.mp3', mime: 'audio/mpeg', edit: true },
  wav: { app: 'video', label: 'WAV Audio', ext: '.wav', mime: 'audio/wav', edit: true },
  flac: { app: 'video', label: 'FLAC Audio', ext: '.flac', mime: 'audio/flac', edit: false },
  m4a: { app: 'video', label: 'AAC Audio', ext: '.m4a', mime: 'audio/mp4', edit: true },
  ogg: { app: 'video', label: 'Ogg Audio', ext: '.ogg', mime: 'audio/ogg', edit: false },

  eml: { app: 'mail', label: 'Email Message', ext: '.eml', mime: 'message/rfc822', edit: false },
  msg: { app: 'mail', label: 'Outlook Message', ext: '.msg', mime: 'application/vnd.ms-outlook', edit: false },
  mbox: { app: 'mail', label: 'Mbox Mail Archive', ext: '.mbox', mime: 'application/mbox', edit: false },
  pst: { app: 'mail', label: 'Outlook Data File', ext: '.pst', mime: 'application/vnd.ms-outlook-pst', edit: false },
  ost: { app: 'mail', label: 'Outlook Offline Data File', ext: '.ost', mime: 'application/vnd.ms-outlook-pst', edit: false },
  olm: { app: 'mail', label: 'Outlook for Mac Archive', ext: '.olm', mime: 'application/vnd.ms-outlook-olm', edit: false },
  vcf: { app: 'mail', label: 'Contact Card', ext: '.vcf', mime: 'text/vcard', edit: false },
  ics: { app: 'mail', label: 'Calendar', ext: '.ics', mime: 'text/calendar', edit: false },

  pdf: { app: 'pictures', label: 'PDF Document', ext: '.pdf', mime: 'application/pdf', edit: false },
  zip: { app: null, label: 'Zip Archive', ext: '.zip', mime: 'application/zip', edit: false },
  unknown: { app: null, label: 'Unknown', ext: '', mime: 'application/octet-stream', edit: false },
};

const BY_EXT = (() => {
  const m = new Map();
  const add = (ext, kind) => m.set(ext, kind);
  for (const [kind, def] of Object.entries(KINDS)) if (def.ext) add(def.ext, kind);
  add('.jpg', 'jpeg');
  add('.jpeg', 'jpeg');
  add('.jpe', 'jpeg');
  add('.tif', 'tiff');
  add('.htm', 'html');
  add('.markdown', 'md');
  add('.text', 'txt');
  add('.log', 'txt');
  add('.json', 'txt');
  add('.xml', 'txt');
  add('.yml', 'txt');
  add('.yaml', 'txt');
  add('.heif', 'heic');
  add('.m4v', 'mp4');
  add('.emlx', 'eml');
  add('.mbx', 'mbox');
  return m;
})();

export function kindFromExtension(name = '') {
  const i = name.lastIndexOf('.');
  if (i < 0) return null;
  return BY_EXT.get(name.slice(i).toLowerCase()) || null;
}

/** ISO-BMFF brand → kind. mp4, mov, heic and avif share one container. */
function fromFtyp(b) {
  const brand = ascii(b, 8, 4);
  const table = {
    'ftyp': null,
    'qt  ': 'mov',
    'isom': 'mp4',
    'iso2': 'mp4',
    'mp41': 'mp4',
    'mp42': 'mp4',
    'M4V ': 'mp4',
    'M4A ': 'm4a',
    'M4B ': 'm4a',
    'avif': 'avif',
    'avis': 'avif',
    'heic': 'heic',
    'heix': 'heic',
    'hevc': 'heic',
    'mif1': 'heic',
    'msf1': 'heic',
  };
  if (table[brand]) return table[brand];
  // Unknown brand: look through the compatible-brand list before giving up.
  const size = (b[0] << 24) | (b[1] << 16) | (b[2] << 8) | b[3];
  const end = Math.min(size, b.length, 64);
  for (let at = 16; at + 4 <= end; at += 4) {
    const compat = ascii(b, at, 4);
    if (table[compat]) return table[compat];
  }
  return 'mp4';
}

/** Zip container: OOXML, ODF, OLM or a plain archive. */
function fromZip(b) {
  // The local file header names the first entry. ODF puts "mimetype" first by
  // specification, which makes it identifiable from the first 100 bytes.
  const nameLen = b[26] | (b[27] << 8);
  const extraLen = b[28] | (b[29] << 8);
  const first = ascii(b, 30, nameLen);
  if (first === 'mimetype') {
    const at = 30 + nameLen + extraLen;
    const mime = ascii(b, at, 60);
    if (mime.startsWith('application/vnd.oasis.opendocument.text')) return 'odt';
    if (mime.startsWith('application/vnd.oasis.opendocument.spreadsheet')) return 'ods';
    if (mime.startsWith('application/vnd.oasis.opendocument.presentation')) return 'odp';
    if (mime.startsWith('application/vnd.oasis.opendocument.graphics')) return 'odp';
  }
  // OOXML always begins with [Content_Types].xml written first by every writer
  // that matters, but not by specification — so this is a hint, not a verdict.
  if (first === '[Content_Types].xml') return 'ooxml';
  if (first.startsWith('word/')) return 'docx';
  if (first.startsWith('xl/')) return 'xlsx';
  if (first.startsWith('ppt/')) return 'pptx';
  return 'zip';
}

/** Text that is probably a mail message rather than prose. */
function looksLikeRfc822(text) {
  const head = text.slice(0, 2048);
  if (/^From \S+.*\r?\n/.test(head)) return 'mbox';
  const headerish = /^(Received|Return-Path|Message-ID|From|To|Subject|Date|MIME-Version|Content-Type):/im;
  if (!headerish.test(head)) return null;
  const firstLine = head.split(/\r?\n/, 1)[0];
  return /^[A-Za-z-]+:\s/.test(firstLine) ? 'eml' : null;
}

function looksLikeText(b) {
  const n = Math.min(b.length, 1024);
  let control = 0;
  for (let i = 0; i < n; i++) {
    const c = b[i];
    if (c === 0) return false;
    if (c < 9 || (c > 13 && c < 32)) control++;
  }
  return control / Math.max(1, n) < 0.02;
}

/**
 * The UTF-16 encoding a byte-order mark names, or null.
 *
 * UTF-32's mark begins with UTF-16LE's, so it is ruled out first — and left
 * alone, because nothing here reads UTF-32 and calling it text would be worse
 * than calling it unknown.
 */
function utf16Encoding(b) {
  const has = (...m) => b.length >= m.length && m.every((v, i) => b[i] === v);
  if (has(0xff, 0xfe, 0x00, 0x00) || has(0x00, 0x00, 0xfe, 0xff)) return 'utf-32';
  if (has(0xff, 0xfe)) return 'utf-16le';
  if (has(0xfe, 0xff)) return 'utf-16be';
  return null;
}

function delimitedGuess(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim()).slice(0, 8);
  if (lines.length < 2) return null;
  const count = (l, ch) => l.split(ch).length - 1;
  for (const [ch, kind] of [['\t', 'tsv'], [',', 'csv'], [';', 'csv']]) {
    const counts = lines.map((l) => count(l, ch));
    if (counts[0] >= 1 && counts.every((c) => c === counts[0])) return kind;
  }
  return null;
}

/**
 * Identify a file.
 * @param {Uint8Array|ArrayBuffer} input first bytes (4 KB is plenty)
 * @param {string} [name] file name, used only to break ties
 * @returns {{ kind: string, app: string|null, label: string, mime: string, confidence: 'magic'|'container'|'text'|'extension'|'none', container?: string }}
 */
export function sniff(input, name = '') {
  const b = bytes(input);
  const byExt = kindFromExtension(name);
  const answer = (kind, confidence, extra) => ({
    kind,
    app: KINDS[kind]?.app ?? null,
    label: KINDS[kind]?.label ?? 'Unknown',
    mime: KINDS[kind]?.mime ?? 'application/octet-stream',
    editable: Boolean(KINDS[kind]?.edit),
    confidence,
    ...extra,
  });

  if (b.length === 0) return answer(byExt || 'unknown', byExt ? 'extension' : 'none');

  // A UTF-16 byte-order mark, before the binary magics — because it collides
  // with one. FF FE is a valid MPEG frame sync, so a text file saved as
  // Unicode on Windows (Notepad's own "Unicode", and what PowerShell writes
  // when it redirects) was identified as an MP3 and the document apps refused
  // it: "Rutba Office cannot open MP3 Audio files yet." The content is decoded
  // and identified as the text it is; `decodeText` reads the same mark when
  // the file is opened.
  const wide = utf16Encoding(b);
  // UTF-32 begins with UTF-16LE's mark. Nothing here reads it, but it is
  // certainly not an MP3, and saying so is the difference between a refusal a
  // person can act on and a wrong one.
  if (wide === 'utf-32') return answer(byExt && KINDS[byExt] ? byExt : 'unknown', byExt ? 'extension' : 'magic', { encoding: 'utf-32' });
  if (wide) {
    const text = new TextDecoder(wide, { fatal: false }).decode(b.subarray(2));
    return { ...sniff(new TextEncoder().encode(text), name), encoding: wide };
  }

  if (starts(b, PNG)) return answer('png', 'magic');
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return answer('jpeg', 'magic');
  if (starts(b, GIF)) return answer('gif', 'magic');
  if (starts(b, BMP)) return answer('bmp', 'magic');
  if (starts(b, PSD)) return answer('psd', 'magic');
  if (starts(b, ICO)) return answer('ico', 'magic');
  if (ascii(b, 0, 4) === 'RIFF') {
    const form = ascii(b, 8, 4);
    if (form === 'WEBP') return answer('webp', 'magic');
    if (form === 'WAVE') return answer('wav', 'magic');
    if (form === 'AVI ') return answer('mp4', 'magic', { container: 'avi' });
  }
  if ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0x00)) {
    return answer('tiff', 'magic');
  }
  if (starts(b, PDF)) return answer('pdf', 'magic');
  if (starts(b, EBML)) {
    // WebM and Matroska share EBML; the DocType string separates them.
    const head = ascii(b, 0, Math.min(b.length, 256));
    return answer(head.includes('webm') ? 'webm' : 'mkv', 'magic');
  }
  if (starts(b, OGG)) {
    const head = ascii(b, 0, Math.min(b.length, 128));
    return answer(head.includes('theora') || head.includes('video') ? 'ogv' : 'ogg', 'magic');
  }
  if (starts(b, FLAC)) return answer('flac', 'magic');
  if (starts(b, PST)) return answer(byExt === 'ost' ? 'ost' : 'pst', 'magic');
  if (starts(b, RAR) || starts(b, SEVENZ) || starts(b, GZ)) return answer('unknown', 'magic', { container: 'archive' });
  if (ascii(b, 4, 4) === 'ftyp') return answer(fromFtyp(b), 'magic');
  if (starts(b, [0x49, 0x44, 0x33]) || (b[0] === 0xff && (b[1] & 0xe0) === 0xe0)) return answer('mp3', 'magic');

  if (starts(b, CFB)) {
    // Compound File: Word, Excel, PowerPoint or an Outlook message. The
    // directory decides, and that needs the whole file — so callers with only a
    // prefix get the extension's opinion, marked as such.
    const guess = byExt && ['doc', 'xls', 'ppt', 'msg'].includes(byExt) ? byExt : 'doc';
    return answer(guess, byExt ? 'container' : 'magic', { container: 'cfb' });
  }

  if (starts(b, ZIP)) {
    const z = fromZip(b);
    if (z === 'ooxml') {
      const guess = byExt && ['docx', 'docm', 'dotx', 'xlsx', 'xlsm', 'xltx', 'pptx', 'pptm', 'potx', 'ppsx'].includes(byExt) ? byExt : null;
      return guess
        ? answer(guess, 'container', { container: 'ooxml' })
        : answer('zip', 'container', { container: 'ooxml', needsDeepScan: true });
    }
    if (z === 'zip' && byExt === 'olm') return answer('olm', 'container', { container: 'zip' });
    if (z === 'zip' && byExt) return answer(byExt, 'extension', { container: 'zip' });
    return answer(z, 'container', { container: 'zip' });
  }

  if (looksLikeText(b)) {
    const text = new TextDecoder('utf-8', { fatal: false }).decode(b.subarray(0, 4096));
    const trimmed = text.replace(/^﻿/, '').trimStart();
    if (trimmed.startsWith('{\\rtf')) return answer('rtf', 'magic');
    if (/^<\?xml/i.test(trimmed) && /<svg[\s>]/i.test(text)) return answer('svg', 'magic');
    if (/^<svg[\s>]/i.test(trimmed)) return answer('svg', 'magic');
    if (/^<!doctype html|^<html[\s>]/i.test(trimmed)) return answer('html', 'magic');
    if (trimmed.startsWith('BEGIN:VCALENDAR')) return answer('ics', 'magic');
    if (trimmed.startsWith('BEGIN:VCARD')) return answer('vcf', 'magic');
    const mailish = looksLikeRfc822(trimmed);
    if (mailish) return answer(mailish, 'text');
    if (byExt && ['csv', 'tsv', 'md', 'txt', 'html'].includes(byExt)) return answer(byExt, 'extension');
    const delim = delimitedGuess(text);
    if (delim) return answer(delim, 'text');
    return answer(byExt && KINDS[byExt] ? byExt : 'txt', byExt ? 'extension' : 'text');
  }

  return answer(byExt || 'unknown', byExt ? 'extension' : 'none');
}

/**
 * Second pass for a zip whose first entry did not settle it: read the part
 * names. `entries` is any iterable of names — the caller already has the zip
 * open, so we do not open it again.
 */
export function refineOoxml(entryNames) {
  const names = new Set(entryNames);
  if (names.has('word/document.xml')) return 'docx';
  if (names.has('xl/workbook.xml')) return 'xlsx';
  if (names.has('ppt/presentation.xml')) return 'pptx';
  for (const n of names) {
    if (n.startsWith('word/')) return 'docx';
    if (n.startsWith('xl/')) return 'xlsx';
    if (n.startsWith('ppt/')) return 'pptx';
  }
  return 'zip';
}

/** Which app should open this kind — the one place that mapping is decided. */
export function appFor(kind) {
  return KINDS[kind]?.app ?? null;
}
