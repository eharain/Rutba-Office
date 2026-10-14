/**
 * Forms in a document, content controls (`w:sdt`) and the legacy form
 * fields (FORMTEXT, FORMCHECKBOX, FORMDROPDOWN), read, written and filled
 * as Word writes them.
 *
 * A form field in a paragraph is ONE character of its text, as an equation
 * is: the run the editor holds carries the field's XML whole (`control.xml`)
 * and a rebuilt paragraph writes it back byte for byte. Filling a field is a
 * new XML for that run, made here: the words or the box's glyph in its
 * content, `w:showingPlcHdr` taken off (or put back when it is emptied), a
 * date's `w:fullDate`, a list's `w:lastValue`, a check box's `w14:checked`;
 * for a legacy field the `w:ffData` and the field's result. Nothing here
 * knows about packages: a picture's bytes and a custom placeholder's words
 * (Word keeps those in the glossary) are the document's to resolve.
 */
import zlib from 'node:zlib';
import { esc, attrs } from './package.js';
import { unesc } from './workbook.js';
import { crc32 } from './zip.js';

/** The one character a form field is in its paragraph's text: U+FFFC, as an equation's. */
export const CONTROL_MARK = '￼';

export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';

/** The kinds of content control the Developer tab inserts. */
export const CONTROL_KINDS = ['richText', 'plainText', 'picture', 'checkBox', 'comboBox', 'dropDownList', 'date'];

/** The legacy fields' kinds, by their field code. */
export const LEGACY_KINDS = { FORMTEXT: 'formText', FORMCHECKBOX: 'formCheckBox', FORMDROPDOWN: 'formDropDown' };

/**
 * Word's own placeholders, by the glossary entry a control names, in the
 * words Word writes into the file when there is no glossary to read: a
 * document opened without its glossary shows these.
 */
export const DEFAULT_PLACEHOLDERS = {
  'DefaultPlaceholder_-1854013440': 'Click or tap here to enter text.',
  'DefaultPlaceholder_-1854013439': 'Choose an item.',
  'DefaultPlaceholder_-1854013438': 'Choose an item.',
  'DefaultPlaceholder_-1854013437': 'Click or tap to enter a date.',
  'DefaultPlaceholder_-1854013436': 'Click or tap here to enter text.',
};

/** The glossary entry Word names for each kind's default placeholder. */
const DEFAULT_DOC_PART = {
  richText: 'DefaultPlaceholder_-1854013440',
  plainText: 'DefaultPlaceholder_-1854013440',
  comboBox: 'DefaultPlaceholder_-1854013439',
  dropDownList: 'DefaultPlaceholder_-1854013438',
  date: 'DefaultPlaceholder_-1854013437',
};

/** The words a kind shows until it is filled, as Word 365 writes them. */
export function defaultPlaceholder(kind) {
  if (kind === 'comboBox' || kind === 'dropDownList') return 'Choose an item.';
  if (kind === 'date') return 'Click or tap to enter a date.';
  return 'Click or tap here to enter text.';
}

/** Word's own check box glyphs: ☒ and ☐ in MS Gothic. */
const CHECKED = { val: '2612', font: 'MS Gothic' };
const UNCHECKED = { val: '2610', font: 'MS Gothic' };

/** The run properties Word gives a control's placeholder words. */
const PLACEHOLDER_RPR = '<w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>';

/** The Placeholder Text character style, as Word 365 defines it. */
export const PLACEHOLDER_STYLE_XML = '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';

const SDT_OPEN = /<w:sdt\b[^>]*>/g;

/**
 * The end of the element whose opening tag starts at `at`, `<w:sdt>` holding
 * `<w:sdt>`s of its own included, or -1 when it never closes.
 */
function closeOf(xml, at, name) {
  const re = new RegExp('<' + name + '\\b[^>]*?(/?)>|</' + name + '>', 'g');
  re.lastIndex = at;
  let depth = 0;
  let m;
  while ((m = re.exec(xml))) {
    if (m[0].startsWith('</')) {
      depth -= 1;
      if (depth === 0) return m.index + m[0].length;
    } else if (!m[1]) depth += 1;
    else if (depth === 0) return m.index + m[0].length;
  }
  return -1;
}

/** The wrappers a form field must not sit inside to be folded: a link, a tracked change, a smart tag. */
const WRAPPERS = /<(\/?)w:(hyperlink|ins|del|smartTag|customXml|fldSimple|moveTo|moveFrom)\b[^>]*?(\/?)>/g;

/** For each position, is it inside one of `WRAPPERS`? Answers a test over offsets. */
function wrapperDepth(xml) {
  const spans = [];
  let depth = 0;
  let start = -1;
  for (const m of xml.matchAll(WRAPPERS)) {
    if (m[3]) continue;
    if (m[1]) {
      depth -= 1;
      if (depth === 0 && start >= 0) spans.push([start, m.index + m[0].length]);
    } else {
      if (depth === 0) start = m.index;
      depth += 1;
    }
  }
  return (at) => spans.some(([a, b]) => at > a && at < b);
}

/** The field code tokens of a paragraph, for the legacy fields: begin, instr, separate, end, by run. */
function fieldRuns(xml) {
  const out = [];
  const re = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g;
  let m;
  while ((m = re.exec(xml))) {
    const inner = m[1];
    const begin = /<w:fldChar\b[^>]*\bw:fldCharType="begin"/.test(inner);
    const sep = /<w:fldChar\b[^>]*\bw:fldCharType="separate"/.test(inner);
    const end = /<w:fldChar\b[^>]*\bw:fldCharType="end"/.test(inner);
    const instr = [...inner.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)].map((x) => unesc(x[1])).join('');
    if (begin || sep || end || instr) out.push({ start: m.index, end: m.index + m[0].length, begin, sep, endChar: end, instr });
  }
  return out;
}

/** The complete legacy form fields at the top of a paragraph: `{ start, end }` from the begin run to the end run. */
function legacySpans(xml) {
  const runs = fieldRuns(xml);
  const out = [];
  for (let i = 0; i < runs.length; i++) {
    if (!runs[i].begin) continue;
    let depth = 0;
    let instr = '';
    let j = i;
    for (; j < runs.length; j++) {
      if (runs[j].begin) depth += 1;
      if (depth === 1 && runs[j].instr) instr += runs[j].instr;
      if (runs[j].endChar) {
        depth -= 1;
        if (depth === 0) break;
      }
    }
    if (j >= runs.length) break;
    const code = instr.trim().split(/\s+/)[0]?.toUpperCase();
    if (LEGACY_KINDS[code]) out.push({ start: runs[i].start, end: runs[j].end, code });
    i = j;
  }
  return out;
}

/**
 * Every form field at the top of a paragraph's XML, each run-level content
 * control not wrapped in a link or a tracked change (a citation's control is
 * a field the citation code owns, and stays), and each complete legacy form
 * field, replaced by `<w:rutbaControl w:k="n"/>`, its XML kept in
 * `controls[n]`. `parseRuns` reads the stand-in as the field's one run.
 */
export function foldControls(xml) {
  const source = String(xml);
  const controls = [];
  if (!source.includes('<w:sdt') && !/FORM(?:TEXT|CHECKBOX|DROPDOWN)/.test(source)) return { xml: source, controls };
  const inWrapper = wrapperDepth(source);
  const spans = [];
  SDT_OPEN.lastIndex = 0;
  let m;
  while ((m = SDT_OPEN.exec(source))) {
    const end = closeOf(source, m.index, 'w:sdt');
    if (end < 0) break;
    const whole = source.slice(m.index, end);
    const pr = /<w:sdtPr\b[^>]*>([\s\S]*?)<\/w:sdtPr>/.exec(whole)?.[1] ?? '';
    // A citation's control is folded by the citation code; a paragraph inside a control is not a run's.
    if (!/<w:citation\s*\/>/.test(pr) && !/<w:p\b/.test(whole) && !inWrapper(m.index)) spans.push({ start: m.index, end });
    SDT_OPEN.lastIndex = end;
  }
  for (const f of legacySpans(source)) {
    if (inWrapper(f.start) || spans.some((s) => f.start < s.end && f.end > s.start)) continue;
    const between = source.slice(f.start, f.end);
    if (/<\/?w:(?:hyperlink|ins|del|p)\b/.test(between)) continue;
    // The bookmark Word puts round a form field, by the field's name, goes with it.
    const name = /<w:ffData\b[\s\S]*?<w:name\b[^>]*w:val="([^"]*)"/.exec(between)?.[1];
    const mark = name != null ? new RegExp('<w:bookmarkStart\\b[^>]*w:name="' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>$').exec(source.slice(0, f.start)) : null;
    const id = mark ? /w:id="([^"]*)"/.exec(mark[0])?.[1] : null;
    const close = id != null ? new RegExp('^<w:bookmarkEnd\\b[^>]*w:id="' + id + '"[^>]*/>').exec(source.slice(f.end)) : null;
    spans.push(mark && close ? { start: f.start - mark[0].length, end: f.end + close[0].length } : f);
  }
  if (!spans.length) return { xml: source, controls };
  spans.sort((a, b) => a.start - b.start);
  let out = '';
  let at = 0;
  for (const s of spans) {
    out += source.slice(at, s.start);
    controls.push(source.slice(s.start, s.end));
    out += '<w:rutbaControl w:k="' + (controls.length - 1) + '"/>';
    at = s.end;
  }
  return { xml: out + source.slice(at), controls };
}

/** A paragraph's XML with its form fields each written as the one character it is in the text. */
export function withControlsAsMarks(xml) {
  const folded = foldControls(xml);
  if (!folded.controls.length) return String(xml);
  return folded.xml.replace(/<w:rutbaControl w:k="\d+"\/>/g, () => '<w:r><w:t>' + CONTROL_MARK + '</w:t></w:r>');
}

/* ── reading ─────────────────────────────────────────────────────────────── */

const prOf = (xml) => /<w:sdtPr\b[^>]*>([\s\S]*?)<\/w:sdtPr>/.exec(xml)?.[1] ?? '';
const contentOf = (xml) => {
  const open = /<w:sdtContent\b[^>]*>/.exec(xml);
  if (!open) return '';
  const close = xml.lastIndexOf('</w:sdtContent>');
  return close > open.index ? xml.slice(open.index + open[0].length, close) : '';
};
const valOf = (xml, tag) => {
  const m = new RegExp('<' + tag + '\\b([^>]*?)/?>').exec(xml);
  if (!m) return null;
  const a = attrs(m[1]);
  return a['w:val'] != null ? unesc(a['w:val']) : '';
};

/** The words of some runs, as Word shows them: text, tabs, breaks; deleted words left out. */
function wordsOf(xml) {
  let out = '';
  const re = /<w:t\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:t>)|<w:tab\s*\/>|<w:(?:br|cr)\b[^>]*\/>|<w:sym\b([^>]*)\/>/g;
  for (const m of String(xml).replace(/<w:del\b[\s\S]*?<\/w:del>/g, '').matchAll(re)) {
    if (m[0].startsWith('<w:tab')) out += '\t';
    else if (m[0].startsWith('<w:sym')) {
      const c = attrs(m[2])['w:char'];
      if (c) out += String.fromCodePoint(parseInt(c, 16) >= 0xF000 ? parseInt(c, 16) - 0xF000 : parseInt(c, 16));
    } else if (m[0].startsWith('<w:t')) out += unesc(m[1] ?? '');
    else out += '\n';
  }
  return out;
}

/** The first run's properties of a stretch of XML, without the placeholder style, the look the words are written in. */
function firstRPr(xml) {
  const run = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/.exec(xml);
  const rPr = run ? /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/.exec(run[1])?.[0] ?? null : null;
  return rPr;
}

const withoutPlaceholderStyle = (rPr) => {
  if (!rPr) return null;
  const out = rPr.replace(/<w:rStyle\b[^>]*w:val="PlaceholderText"[^>]*\/>/, '');
  return /<w:rPr\b[^>]*>\s*<\/w:rPr>/.test(out) || /<w:rPr\b[^>]*\/>/.test(out) ? null : out;
};

/** A run's look, for drawing a field's words the way the file has them. */
function lookOf(rPr) {
  if (!rPr) return {};
  const on = (tag) => {
    const m = new RegExp('<w:' + tag + '\\b([^>]*)/?>').exec(rPr);
    if (!m) return false;
    const v = /w:val="([^"]*)"/.exec(m[1]);
    return !v || !['0', 'false', 'off', 'none'].includes(v[1]);
  };
  const colour = /<w:color\b[^>]*w:val="([0-9A-Fa-f]{6})"/.exec(rPr)?.[1];
  const size = /<w:sz\b[^>]*w:val="(\d+)"/.exec(rPr)?.[1];
  const font = /<w:rFonts\b[^>]*w:ascii="([^"]*)"/.exec(rPr)?.[1];
  return {
    ...(on('b') ? { bold: true } : {}), ...(on('i') ? { italic: true } : {}), ...(on('u') ? { underline: true } : {}),
    ...(colour ? { colour: colour.toUpperCase() } : {}), ...(size ? { sizePt: Number(size) / 2 } : {}), ...(font ? { font: unesc(font) } : {}),
  };
}

/** What a content control or a legacy form field is, read from its XML. */
export function readControl(xml) {
  const source = String(xml || '');
  if (!source.startsWith('<w:sdt')) return readLegacy(source);
  const pr = prOf(source);
  const content = contentOf(source);
  const kind = /<w14:checkbox\b/.test(pr) ? 'checkBox'
    : /<w:text\b/.test(pr) ? 'plainText'
      : /<w:comboBox\b/.test(pr) ? 'comboBox'
        : /<w:dropDownList\b/.test(pr) ? 'dropDownList'
          : /<w:date\b/.test(pr) ? 'date'
            : /<w:picture\b/.test(pr) ? 'picture'
              : /<w:(?:docPartObj|docPartList|bibliography|equation|group)\b|<w15:repeatingSection\b/.test(pr) ? 'other'
                : 'richText';
  const lock = valOf(pr, 'w:lock');
  const showing = /<w:showingPlcHdr\b(?![^>]*w:val="(?:0|false)")/.test(pr);
  const words = wordsOf(content);
  const out = {
    kind,
    id: valOf(pr, 'w:id'),
    title: valOf(pr, 'w:alias') ?? '',
    tag: valOf(pr, 'w:tag') ?? '',
    docPart: /<w:placeholder\b[^>]*>[\s\S]*?<w:docPart\b[^>]*w:val="([^"]*)"/.exec(pr)?.[1] ?? null,
    showingPlaceholder: showing,
    noDelete: lock === 'sdtLocked' || lock === 'sdtContentLocked',
    noEdit: lock === 'contentLocked' || lock === 'sdtContentLocked',
    temporary: /<w:temporary\b/.test(pr),
    text: words,
    value: showing ? '' : words,
    placeholder: showing ? words : null,
    look: lookOf(withoutPlaceholderStyle(firstRPr(content)) || /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/.exec(pr)?.[0] || null),
    ...(/<w:dataBinding\b/.test(pr) ? { bound: true } : {}),
  };
  if (kind === 'plainText') out.multiLine = /<w:text\b[^>]*w:multiLine="(?:1|true)"/.test(pr);
  if (kind === 'checkBox') {
    const box = /<w14:checkbox\b[^>]*>([\s\S]*?)<\/w14:checkbox>/.exec(pr)?.[1] ?? '';
    const checked = /<w14:checked\b[^>]*w14:val="([^"]*)"/.exec(box)?.[1];
    out.checked = checked === '1' || checked === 'true';
    const state = (tag, fallback) => {
      const a = attrs(new RegExp('<w14:' + tag + '\\b([^>]*)/?>').exec(box)?.[1] ?? '');
      return { val: a['w14:val'] || fallback.val, font: a['w14:font'] ? unesc(a['w14:font']) : fallback.font };
    };
    out.checkedState = state('checkedState', CHECKED);
    out.uncheckedState = state('uncheckedState', UNCHECKED);
    out.value = out.checked;
    out.placeholder = null;
  }
  if (kind === 'comboBox' || kind === 'dropDownList') {
    const list = new RegExp('<w:' + kind + '\\b([^>]*?)(?:/>|>([\\s\\S]*?)</w:' + kind + '>)').exec(pr);
    out.items = [...(list?.[2] ?? '').matchAll(/<w:listItem\b([^>]*)\/?>/g)].map((x) => {
      const a = attrs(x[1]);
      const text = a['w:displayText'] != null ? unesc(a['w:displayText']) : unesc(a['w:value'] ?? '');
      return { text, value: a['w:value'] != null ? unesc(a['w:value']) : text };
    });
    const last = attrs(list?.[1] ?? '')['w:lastValue'];
    out.lastValue = last != null ? unesc(last) : null;
  }
  if (kind === 'date') {
    const d = /<w:date\b([^>]*?)(?:\/>|>([\s\S]*?)<\/w:date>)/.exec(pr);
    const full = attrs(d?.[1] ?? '')['w:fullDate'];
    out.date = full ? full.slice(0, 10) : null;
    out.dateFormat = valOf(d?.[2] ?? '', 'w:dateFormat') || 'dd/MM/yyyy';
    out.lid = valOf(d?.[2] ?? '', 'w:lid') || 'en-GB';
    out.calendar = valOf(d?.[2] ?? '', 'w:calendar') || 'gregorian';
  }
  if (kind === 'picture') {
    const blip = /<a:blip\b[^>]*r:embed="([^"]*)"/.exec(content);
    out.rId = blip ? blip[1] : null;
    const ext = /<wp:extent\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(content);
    out.widthPx = ext ? Number(ext[1]) / 9525 : 120;
    out.heightPx = ext ? Number(ext[2]) / 9525 : 120;
    out.value = showing ? null : out.rId;
    out.placeholder = null;
  }
  return out;
}

/** A legacy form field's `w:ffData`, read: its name, its kind, and what it holds. */
function readLegacy(xml) {
  const ff = /<w:ffData\b[^>]*>([\s\S]*?)<\/w:ffData>/.exec(xml)?.[1] ?? '';
  const instr = [...xml.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)].map((m) => unesc(m[1])).join('').trim().split(/\s+/)[0]?.toUpperCase();
  const kind = LEGACY_KINDS[instr] || 'formText';
  const name = valOf(ff, 'w:name') ?? '';
  const out = {
    kind, legacy: true, id: name, title: name, tag: '',
    enabled: !/<w:enabled\b[^>]*w:val="(?:0|false)"/.test(ff),
    help: valOf(ff, 'w:helpText') ?? '',
    status: valOf(ff, 'w:statusText') ?? '',
    noDelete: false, noEdit: false, showingPlaceholder: false, placeholder: null,
  };
  if (kind === 'formCheckBox') {
    const box = /<w:checkBox\b[^>]*>([\s\S]*?)<\/w:checkBox>/.exec(ff)?.[1] ?? '';
    const checked = valOf(box, 'w:checked');
    const def = valOf(box, 'w:default');
    const on = (v) => v === '' || v === '1' || v === 'true' || v === 'on';
    out.checked = checked != null ? on(checked) : on(def ?? '0');
    out.value = out.checked;
    out.text = out.checked ? '☒' : '☐';
    return out;
  }
  if (kind === 'formDropDown') {
    const dd = /<w:ddList\b[^>]*>([\s\S]*?)<\/w:ddList>/.exec(ff)?.[1] ?? '';
    out.items = [...dd.matchAll(/<w:listEntry\b[^>]*w:val="([^"]*)"/g)].map((m) => ({ text: unesc(m[1]), value: unesc(m[1]) }));
    const at = Number(valOf(dd, 'w:result') ?? valOf(dd, 'w:default') ?? 0) || 0;
    out.selected = Math.min(Math.max(0, at), Math.max(0, out.items.length - 1));
    out.value = out.items[out.selected]?.text ?? '';
    out.text = out.value;
    return out;
  }
  // A text field's words are its field's result.
  const sep = xml.search(/<w:fldChar\b[^>]*w:fldCharType="separate"/);
  const end = xml.search(/<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*?<w:fldChar\b[^>]*w:fldCharType="end"/);
  const result = sep >= 0 && end > sep ? xml.slice(xml.indexOf('</w:r>', sep) + 6, end) : '';
  const input = /<w:textInput\b[^>]*>([\s\S]*?)<\/w:textInput>/.exec(ff)?.[1] ?? '';
  out.text = wordsOf(result);
  out.value = out.text.replace(/\u2002/g, '').trim() ? out.text : '';
  out.default = valOf(input, 'w:default') ?? '';
  out.maxLength = Number(valOf(input, 'w:maxLength') ?? 0) || 0;
  out.textType = valOf(input, 'w:type') || 'regular';
  out.format = valOf(input, 'w:format') ?? '';
  out.look = lookOf(firstRPr(result));
  return out;
}

/* ── dates ───────────────────────────────────────────────────────────────── */

/**
 * A date in one of Word's date formats, `dd/MM/yyyy`, `dddd, MMMM d, yyyy`,
 * `d-MMM-yy`, `h:mm am/pm`, its names in the control's language. Words in
 * single quotes are written as they are.
 */
export function formatControlDate(iso, format = 'dd/MM/yyyy', lid = 'en-GB') {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(String(iso || ''));
  if (!m) return '';
  const [y, mo, d, h = '00', mi = '00', s = '00'] = m.slice(1);
  const when = new Date(Date.UTC(Number(y), Number(mo) - 1, Number(d), Number(h), Number(mi), Number(s)));
  let locale = 'en-GB';
  try { locale = Intl.getCanonicalLocales(lid)[0] || 'en-GB'; } catch { /* a language tag Word knows and Intl does not */ }
  const name = (opts) => new Intl.DateTimeFormat(locale, { timeZone: 'UTC', ...opts }).format(when);
  const hour = Number(h);
  const tokens = {
    yyyy: y, yy: y.slice(2), MMMM: name({ month: 'long' }), MMM: name({ month: 'short' }), MM: mo, M: String(Number(mo)),
    dddd: name({ weekday: 'long' }), ddd: name({ weekday: 'short' }), dd: d, d: String(Number(d)),
    HH: h, H: String(hour), hh: String(((hour + 11) % 12) + 1).padStart(2, '0'), h: String(((hour + 11) % 12) + 1),
    mm: mi, m: String(Number(mi)), ss: s, s: String(Number(s)),
  };
  let out = '';
  const re = /'([^']*)'|am\/pm|AM\/PM|yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|m|ss|s|[\s\S]/g;
  for (const t of String(format).matchAll(re)) {
    if (t[1] !== undefined) out += t[1];
    else if (/^am\/pm$/i.test(t[0])) out += hour < 12 ? (t[0] === 'am/pm' ? 'am' : 'AM') : (t[0] === 'am/pm' ? 'pm' : 'PM');
    else out += tokens[t[0]] ?? t[0];
  }
  return out;
}

/** Word's date formats offered in Properties, its default first. */
export const DATE_FORMATS = ['dd/MM/yyyy', 'd/M/yyyy', 'dd-MM-yy', 'yyyy-MM-dd', 'd MMMM yyyy', 'dddd, d MMMM yyyy', 'MMMM d, yyyy', 'M/d/yyyy', 'd-MMM-yy', 'MMMM yy'];

/* ── writing ─────────────────────────────────────────────────────────────── */

/** A run of words in some run properties; a tab and a break are elements. */
function runOf(rPr, text) {
  const body = String(text).split(/([\t\n])/).map((piece) => {
    if (piece === '\t') return '<w:tab/>';
    if (piece === '\n') return '<w:br/>';
    return piece ? '<w:t xml:space="preserve">' + esc(piece) + '</w:t>' : '';
  }).join('');
  return '<w:r>' + (rPr || '') + body + '</w:r>';
}

/** A new control id, as Word makes them: a signed 32-bit number. */
export function newControlId(taken = new Set()) {
  let id;
  do id = (Math.floor(Math.random() * 0xFFFFFFFF) | 0) || 1;
  while (taken.has(String(id)));
  return String(id);
}

/** The `w:lock` a pair of Properties ticks makes, or none. */
function lockOf({ noDelete = false, noEdit = false } = {}) {
  const v = noDelete && noEdit ? 'sdtContentLocked' : noDelete ? 'sdtLocked' : noEdit ? 'contentLocked' : null;
  return v ? '<w:lock w:val="' + v + '"/>' : '';
}

function listXml(items = []) {
  return items.map((it) => {
    const text = typeof it === 'string' ? it : it.text ?? '';
    const value = typeof it === 'string' ? it : it.value ?? text;
    return '<w:listItem w:displayText="' + esc(text) + '" w:value="' + esc(value) + '"/>';
  }).join('');
}

/** The type element a kind writes in its `w:sdtPr`, last among the properties, as the schema orders them. */
function typeXml(kind, props) {
  switch (kind) {
    case 'plainText': return props.multiLine ? '<w:text w:multiLine="1"/>' : '<w:text/>';
    case 'comboBox': return '<w:comboBox>' + listXml(props.items) + '</w:comboBox>';
    case 'dropDownList': return '<w:dropDownList' + (props.lastValue != null ? ' w:lastValue="' + esc(props.lastValue) + '"' : '') + '>' + listXml(props.items) + '</w:dropDownList>';
    case 'date':
      return '<w:date' + (props.date ? ' w:fullDate="' + esc(props.date) + 'T00:00:00Z"' : '') + '><w:dateFormat w:val="' + esc(props.dateFormat || 'dd/MM/yyyy') + '"/>'
        + '<w:lid w:val="' + esc(props.lid || 'en-GB') + '"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="' + esc(props.calendar || 'gregorian') + '"/></w:date>';
    case 'picture': return '<w:picture/>';
    case 'checkBox': {
      const c = props.checkedState || CHECKED;
      const u = props.uncheckedState || UNCHECKED;
      return '<w14:checkbox><w14:checked w14:val="' + (props.checked ? 1 : 0) + '"/><w14:checkedState w14:val="' + esc(c.val) + '" w14:font="' + esc(c.font) + '"/>'
        + '<w14:uncheckedState w14:val="' + esc(u.val) + '" w14:font="' + esc(u.font) + '"/></w14:checkbox>';
    }
    default: return '';
  }
}

/** A check box's glyph run: the box's character in its font, as Word writes it. */
function checkRun(state, rPr = null) {
  const font = esc(state.font || 'MS Gothic');
  const own = withoutPlaceholderStyle(rPr);
  const fonts = '<w:rFonts w:ascii="' + font + '" w:eastAsia="' + font + '" w:hAnsi="' + font + '" w:hint="eastAsia"/>';
  const props = own ? own.replace(/<w:rFonts\b[^>]*\/>/, '').replace(/^<w:rPr\b[^>]*>/, (o) => o + fonts) : '<w:rPr>' + fonts + '</w:rPr>';
  return runOf(props, String.fromCodePoint(parseInt(state.val, 16)));
}

/**
 * A new content control, inline, as Word's Developer tab inserts one: its
 * title, tag and id, a lock, the placeholder it shows until filled (Word's
 * own default glossary entry, or `docPart` for a placeholder of one's own),
 * and the kind's own element. `picture` is the run holding a picture
 * control's drawing; `placeholder` the words shown until it is filled.
 */
export function controlXml(kind, props = {}, { id, picture = null } = {}) {
  if (!CONTROL_KINDS.includes(kind)) throw new Error('not a kind of content control: ' + kind);
  const placeholder = props.placeholder || props.placeholderDefault || defaultPlaceholder(kind);
  const docPart = props.docPart || DEFAULT_DOC_PART[kind] || null;
  const value = props.value;
  const filled = kind === 'checkBox' ? true : kind === 'picture' ? false : value != null && value !== '';
  const pr = '<w:sdtPr>'
    + (props.rPr || '')
    + (props.title ? '<w:alias w:val="' + esc(props.title) + '"/>' : '')
    + (props.tag ? '<w:tag w:val="' + esc(props.tag) + '"/>' : '')
    + '<w:id w:val="' + esc(id) + '"/>'
    + lockOf(props)
    + (docPart && kind !== 'checkBox' && kind !== 'picture' ? '<w:placeholder><w:docPart w:val="' + esc(docPart) + '"/></w:placeholder>' : '')
    + (!filled && kind !== 'checkBox' ? '<w:showingPlcHdr/>' : '')
    + typeXml(kind, props)
    + '</w:sdtPr>';
  let content;
  if (kind === 'checkBox') content = checkRun(props.checked ? props.checkedState || CHECKED : props.uncheckedState || UNCHECKED);
  else if (kind === 'picture') content = picture || '';
  else if (filled) content = runOf(null, value);
  else content = runOf(PLACEHOLDER_RPR, placeholder);
  return '<w:sdt>' + pr + '<w:sdtContent>' + content + '</w:sdtContent></w:sdt>';
}

/** The control's `w:sdtPr` element replaced by `next(inner)`. */
function editPr(xml, next) {
  return xml.replace(/<w:sdtPr\b[^>]*>([\s\S]*?)<\/w:sdtPr>/, (whole, inner) => '<w:sdtPr>' + next(inner) + '</w:sdtPr>');
}

/** The control's content replaced. */
function withContent(xml, content) {
  const open = /<w:sdtContent\b[^>]*>/.exec(xml);
  if (!open) return xml.replace(/<\/w:sdt>$/, () => '<w:sdtContent>' + content + '</w:sdtContent></w:sdt>');
  const close = xml.lastIndexOf('</w:sdtContent>');
  return xml.slice(0, open.index + open[0].length) + content + xml.slice(close);
}

/** `w:showingPlcHdr` on or off, where the schema puts it (after the placeholder, before the type). */
function withShowing(inner, on) {
  const off = inner.replace(/<w:showingPlcHdr\b[^>]*\/>/g, '');
  if (!on) return off;
  const after = /<w:placeholder\b[\s\S]*?<\/w:placeholder>|<w:lock\b[^>]*\/>|<w:id\b[^>]*\/>/g;
  let at = -1;
  for (const m of off.matchAll(after)) at = Math.max(at, m.index + m[0].length);
  return at < 0 ? '<w:showingPlcHdr/>' + off : off.slice(0, at) + '<w:showingPlcHdr/>' + off.slice(at);
}

/**
 * A content control or a legacy field, filled: `value` a string for words,
 * a list's choice (by its words or its value) or a date (`yyyy-mm-dd`), true
 * or false for a check box. An empty value puts the placeholder back,
 * `placeholder` its words, which the document read from its glossary.
 */
export function withControlValue(xml, value, { placeholder = null } = {}) {
  const source = String(xml);
  if (!source.startsWith('<w:sdt')) return withLegacyValue(source, value);
  const c = readControl(source);
  if (c.kind === 'picture' || c.kind === 'other') throw new Error('that content control is not filled with words');
  if (c.noEdit) throw new Error('This content control cannot be edited: its contents are locked.');
  const content = contentOf(source);
  // The words are written in the control's own look: its first run's, or what its properties keep.
  const ownRPr = withoutPlaceholderStyle(firstRPr(content)) || /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/.exec(prOf(source))?.[0] || null;
  if (c.kind === 'checkBox') {
    const on = value === true || value === 'true' || value === 1 || value === '1';
    const next = editPr(source, (inner) => inner.replace(/<w14:checked\b[^>]*\/>/, () => '<w14:checked w14:val="' + (on ? 1 : 0) + '"/>'));
    return withContent(next, checkRun(on ? c.checkedState : c.uncheckedState, firstRPr(content)));
  }
  let text = value == null ? '' : String(value);
  if (c.kind === 'plainText' && !c.multiLine) text = text.replace(/\r?\n/g, ' ');
  const empty = text === '';
  const shown = placeholder || c.placeholder || DEFAULT_PLACEHOLDERS[c.docPart] || defaultPlaceholder(c.kind);
  let next = editPr(source, (inner) => withShowing(inner, empty));
  if (c.kind === 'dropDownList' || c.kind === 'comboBox') {
    const item = c.items.find((it) => it.text === text) || c.items.find((it) => it.value === text) || null;
    if (c.kind === 'dropDownList' && !empty && !item) throw new Error('"' + text + '" is not one of the list\'s items.');
    if (item) text = item.text;
    next = next.replace(new RegExp('<w:' + c.kind + '\\b([^>]*?)(/?)>'), (whole, a, selfClose) => {
      const kept = a.replace(/\s+w:lastValue="[^"]*"/, '');
      return '<w:' + c.kind + kept + (!empty && item ? ' w:lastValue="' + esc(item.value) + '"' : '') + selfClose + '>';
    });
  }
  if (c.kind === 'date') {
    const iso = empty ? null : /^\d{4}-\d{2}-\d{2}/.test(text) ? text.slice(0, 10) : null;
    if (!empty && !iso) throw new Error('A date is written yyyy-mm-dd.');
    next = next.replace(/<w:date\b([^>]*?)(\/?)>/, (whole, a, selfClose) => '<w:date' + a.replace(/\s+w:fullDate="[^"]*"/, '') + (iso ? ' w:fullDate="' + iso + 'T00:00:00Z"' : '') + selfClose + '>');
    if (iso) text = formatControlDate(iso, c.dateFormat, c.lid);
  }
  return withContent(next, empty ? runOf(PLACEHOLDER_RPR, shown) : runOf(ownRPr, text));
}

/**
 * Properties, as Word's Content Control Properties box sets them: the
 * title, tag and locks; a plain text control's multiple paragraphs; a
 * list's items; a date's format and language. `placeholder` and `docPart`
 * change the words shown while it is empty (the document keeps them in its
 * glossary). A date shown is written again in its new format.
 */
export function withControlProps(xml, props = {}, { placeholder = null } = {}) {
  const source = String(xml);
  if (!source.startsWith('<w:sdt')) return withLegacyProps(source, props);
  const c = readControl(source);
  let next = editPr(source, (inner) => {
    let out = inner;
    const setVal = (tag, v, before) => {
      const el = v ? '<' + tag + ' w:val="' + esc(v) + '"/>' : '';
      if (new RegExp('<' + tag + '\\b').test(out)) out = out.replace(new RegExp('<' + tag + '\\b[^>]*/>'), () => el);
      else if (el) {
        const at = out.search(before);
        out = at < 0 ? out + el : out.slice(0, at) + el + out.slice(at);
      }
    };
    if ('title' in props) setVal('w:alias', props.title, /<w:(?:tag|id)\b/);
    if ('tag' in props) setVal('w:tag', props.tag, /<w:id\b/);
    if ('noDelete' in props || 'noEdit' in props) {
      const lock = lockOf({ noDelete: props.noDelete ?? c.noDelete, noEdit: props.noEdit ?? c.noEdit });
      out = out.replace(/<w:lock\b[^>]*\/>/, '');
      const id = /<w:id\b[^>]*\/>/.exec(out);
      out = id ? out.slice(0, id.index + id[0].length) + lock + out.slice(id.index + id[0].length) : lock + out;
    }
    if (props.docPart) {
      const ph = '<w:placeholder><w:docPart w:val="' + esc(props.docPart) + '"/></w:placeholder>';
      if (/<w:placeholder\b/.test(out)) out = out.replace(/<w:placeholder\b[\s\S]*?<\/w:placeholder>/, () => ph);
      else {
        const lockOrId = /<w:lock\b[^>]*\/>|<w:id\b[^>]*\/>/g;
        let at = -1;
        for (const m of out.matchAll(lockOrId)) at = Math.max(at, m.index + m[0].length);
        out = at < 0 ? ph + out : out.slice(0, at) + ph + out.slice(at);
      }
    }
    if (c.kind === 'plainText' && 'multiLine' in props) out = out.replace(/<w:text\b[^>]*\/>/, () => (props.multiLine ? '<w:text w:multiLine="1"/>' : '<w:text/>'));
    if ((c.kind === 'comboBox' || c.kind === 'dropDownList') && Array.isArray(props.items)) {
      out = out.replace(new RegExp('<w:' + c.kind + '\\b([^>]*?)(?:/>|>[\\s\\S]*?</w:' + c.kind + '>)'), (whole, a) => '<w:' + c.kind + a + '>' + listXml(props.items) + '</w:' + c.kind + '>');
    }
    if (c.kind === 'date' && (props.dateFormat || props.lid)) {
      out = out.replace(/<w:date\b([^>]*?)(?:\/>|>([\s\S]*?)<\/w:date>)/, (whole, a) => '<w:date' + a + '>'
        + '<w:dateFormat w:val="' + esc(props.dateFormat || c.dateFormat) + '"/><w:lid w:val="' + esc(props.lid || c.lid) + '"/>'
        + '<w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="' + esc(c.calendar) + '"/></w:date>');
    }
    return out;
  });
  // What is shown follows: the new placeholder while empty, a date in its new format.
  if (c.showingPlaceholder && (props.placeholder != null || placeholder)) {
    next = withContent(next, runOf(PLACEHOLDER_RPR, props.placeholder || placeholder));
  } else if (c.kind === 'date' && c.date && (props.dateFormat || props.lid)) {
    next = withControlValue(next, c.date);
  } else if (c.kind === 'dropDownList' && Array.isArray(props.items) && !c.showingPlaceholder && !props.items.some((it) => (it.text ?? it) === c.value)) {
    next = withControlValue(next, '', { placeholder });
  }
  return next;
}

/** A picture control holding `pictureRun` (a run with an inline drawing), its placeholder flag taken off. */
export function withControlPicture(xml, pictureRun) {
  const source = String(xml);
  if (readControl(source).kind !== 'picture') throw new Error('that content control does not hold a picture');
  return withContent(editPr(source, (inner) => withShowing(inner, false)), pictureRun);
}

/* ── legacy form fields ──────────────────────────────────────────────────── */

/** Five en spaces: what an empty text form field shows, as Word writes it. */
const EMPTY_TEXT_FIELD = '\u2002\u2002\u2002\u2002\u2002';

/**
 * A legacy form field, as Developer → Legacy Tools inserts one: a complex
 * field whose `begin` carries `w:ffData` (its name, enabled, and the text
 * input, check box or drop-down list), with a bookmark of its name round it.
 */
export function legacyFieldXml(kind, { name, items = [], checked = false, bookmarkId = 0, rPr = null } = {}) {
  const code = Object.entries(LEGACY_KINDS).find(([, k]) => k === kind)?.[0];
  if (!code) throw new Error('not a legacy form field: ' + kind);
  const own = rPr || '';
  const head = '<w:ffData><w:name w:val="' + esc(name) + '"/><w:enabled/><w:calcOnExit w:val="0"/>';
  const ff = kind === 'formText' ? head + '<w:textInput/></w:ffData>'
    : kind === 'formCheckBox' ? head + '<w:checkBox><w:sizeAuto/><w:default w:val="' + (checked ? 1 : 0) + '"/></w:checkBox></w:ffData>'
      : head + '<w:ddList>' + items.map((it) => '<w:listEntry w:val="' + esc(typeof it === 'string' ? it : it.text) + '"/>').join('') + '</w:ddList></w:ffData>';
  const begin = '<w:r>' + own + '<w:fldChar w:fldCharType="begin">' + ff + '</w:fldChar></w:r>';
  const instr = '<w:r>' + own + '<w:instrText xml:space="preserve"> ' + code + ' </w:instrText></w:r>';
  const result = kind === 'formText'
    ? '<w:r>' + own + '<w:fldChar w:fldCharType="separate"/></w:r><w:r>' + (own ? own.replace(/<\/w:rPr>$/, '<w:noProof/></w:rPr>') : '<w:rPr><w:noProof/></w:rPr>') + '<w:t xml:space="preserve">' + EMPTY_TEXT_FIELD + '</w:t></w:r>'
    : '';
  const end = '<w:r>' + own + '<w:fldChar w:fldCharType="end"/></w:r>';
  return '<w:bookmarkStart w:id="' + bookmarkId + '" w:name="' + esc(name) + '"/>' + begin + instr + result + end + '<w:bookmarkEnd w:id="' + bookmarkId + '"/>';
}

/** The field span of a legacy field's XML: its begin run to its end run, leaving a bookmark round it where it was. */
function legacyParts(xml) {
  const begin = xml.search(/<w:r\b(?:(?!<\/w:r>)[\s\S])*?<w:fldChar\b[^>]*w:fldCharType="begin"/);
  const endTag = /<w:r\b(?:(?!<\/w:r>)[\s\S])*?<w:fldChar\b[^>]*w:fldCharType="end"[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g;
  let last = null;
  for (const m of xml.matchAll(endTag)) last = m;
  return { begin, end: last ? last.index + last[0].length : xml.length, endRun: last ? last[0] : '' };
}

/** A legacy field filled: words into a text field's result, a check box ticked, a list's entry chosen. */
export function withLegacyValue(xml, value) {
  const source = String(xml);
  const f = readLegacy(source);
  if (!f.enabled) throw new Error('This form field is not enabled for filling in.');
  if (f.kind === 'formCheckBox') {
    const on = value === true || value === 'true' || value === 1 || value === '1';
    return source.replace(/<w:checkBox\b[^>]*>([\s\S]*?)<\/w:checkBox>/, (whole, inner) => {
      const without = inner.replace(/<w:checked\b[^>]*\/>/g, '');
      return '<w:checkBox>' + without + '<w:checked w:val="' + (on ? 1 : 0) + '"/></w:checkBox>';
    });
  }
  if (f.kind === 'formDropDown') {
    const at = f.items.findIndex((it) => it.text === String(value));
    if (at < 0) throw new Error('"' + value + '" is not one of the list\'s entries.');
    return source.replace(/<w:ddList\b[^>]*>([\s\S]*?)<\/w:ddList>/, (whole, inner) => {
      const without = inner.replace(/<w:result\b[^>]*\/>/g, '');
      return '<w:ddList><w:result w:val="' + at + '"/>' + without + '</w:ddList>';
    });
  }
  let text = value == null ? '' : String(value);
  if (f.maxLength > 0) text = text.slice(0, f.maxLength);
  if (text === '') text = f.default || EMPTY_TEXT_FIELD;
  const { begin, end, endRun } = legacyParts(source);
  if (begin < 0) return source;
  const span = source.slice(begin, end);
  const sep = /<w:r\b(?:(?!<\/w:r>)[\s\S])*?<w:fldChar\b[^>]*w:fldCharType="separate"[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/.exec(span);
  const resultRPr = (() => {
    if (!sep) return '<w:rPr><w:noProof/></w:rPr>';
    const after = span.slice(sep.index + sep[0].length, span.length - endRun.length);
    return firstRPr(after) || '<w:rPr><w:noProof/></w:rPr>';
  })();
  const head = sep ? span.slice(0, sep.index + sep[0].length) : span.slice(0, span.length - endRun.length) + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
  return source.slice(0, begin) + head + runOf(resultRPr, text) + endRun + source.slice(end);
}

/** A legacy field's own properties: a list's entries, a text field's default and length, help words. */
function withLegacyProps(xml, props = {}) {
  let out = String(xml);
  if (Array.isArray(props.items)) {
    out = out.replace(/<w:ddList\b[^>]*>[\s\S]*?<\/w:ddList>|<w:ddList\b[^>]*\/>/, () => '<w:ddList>' + props.items.map((it) => '<w:listEntry w:val="' + esc(typeof it === 'string' ? it : it.text) + '"/>').join('') + '</w:ddList>');
  }
  if (props.title) {
    const name = String(props.title).replace(/[^\w]/g, '').slice(0, 20);
    if (name) out = out.replace(/<w:name\b[^>]*\/>/, () => '<w:name w:val="' + esc(name) + '"/>').replace(/(<w:bookmarkStart\b[^>]*w:name=")[^"]*"/, (m, a) => a + esc(name) + '"');
  }
  return out;
}

/* ── the picture a picture control shows until it is filled ─────────────── */

const pngChunk = (type, data) => {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const head = Buffer.concat([Buffer.from(type, 'latin1'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(head) >>> 0, 0);
  return Buffer.concat([len, head, crc]);
};

/**
 * Word's picture control placeholder, drawn here: a pale square with a
 * picture's sign in it (a frame, a sun and two hills), as a PNG.
 */
export function placeholderPicturePng(size = 120) {
  const n = Math.max(24, Math.round(size));
  const rows = [];
  // A frame two fifths of the square wide, a sun at its top right, two hills along its foot.
  const left = n * 0.27;
  const right = n * 0.73;
  const top = n * 0.31;
  const bottom = n * 0.69;
  const line = Math.max(1.5, n / 60);
  const ink = [166, 166, 166];
  for (let y = 0; y < n; y++) {
    const row = Buffer.alloc(1 + n * 3);
    for (let x = 0; x < n; x++) {
      let rgb = [242, 242, 242];
      const inside = x >= left && x <= right && y >= top && y <= bottom;
      if (inside) {
        const edge = x - left < line || right - x < line || y - top < line || bottom - y < line;
        if (edge) rgb = ink;
        else {
          rgb = [255, 255, 255];
          if ((x - n * 0.61) ** 2 + (y - n * 0.41) ** 2 < (n * 0.045) ** 2) rgb = ink;
          const hill = (cx, h, w) => y > bottom - h * (1 - Math.abs(x - cx) / w);
          if (hill(n * 0.4, n * 0.2, n * 0.17) || hill(n * 0.58, n * 0.13, n * 0.14)) rgb = ink;
        }
      }
      row.set(rgb, 1 + x * 3);
    }
    rows.push(row);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(n, 0);
  ihdr.writeUInt32BE(n, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([sig, pngChunk('IHDR', ihdr), pngChunk('IDAT', zlib.deflateSync(Buffer.concat(rows))), pngChunk('IEND', Buffer.alloc(0))]);
}

/** Every form field in a stretch of XML, in order, read, Restrict Editing's Filling in forms walks these. */
export function formFieldsIn(xml) {
  return foldControls(xml).controls.map((x) => readControl(x));
}
