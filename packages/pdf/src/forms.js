'use strict';

/**
 * A PDF's form: the fields of its AcroForm read, and filled as an
 * incremental update with appearance streams any reader shows.
 *
 * Fields are a tree under the catalog's /AcroForm /Fields; a field's name is
 * its ancestors' partial names joined with dots, and what it is (/FT), its
 * flags (/Ff), its value (/V) and its default look (/DA) are inherited from
 * the nodes above it. The widgets that draw a field on its pages are its
 * kids without names, or the field itself when the two are one dictionary.
 *
 * Filling writes the value where the field keeps it and draws each widget
 * again, as Acrobat does: a text or choice field gets a new normal
 * appearance with its words in the font its /DA names (taken from the
 * form's /DR), sized to fit when /DA says 0; a check box or a radio button
 * is turned to the state that means "on" (/AS), whose appearance the file
 * already holds. Words Windows-1252 cannot spell (Urdu, Arabic, Hindi) are
 * written as the value with no appearance of ours, and the form is marked
 * /NeedAppearances so the reader draws them in a font it has, rather than
 * drawn here as question marks.
 */

const { Name, Ref, Dict, PdfString } = require('./reader');
const { textString } = require('./assemble');
const { PdfUpdate, UpdateError, copyDict, flateStream } = require('./update');
const { encode, literal, fitsWinAnsi } = require('./encoding');
const { widthOfBytes } = require('./metrics');

// Field flags (/Ff), PDF 32000 table 221 onward.
const FF = {
  readOnly: 1 << 0, required: 1 << 1, multiline: 1 << 12, password: 1 << 13, noToggleToOff: 1 << 14,
  radio: 1 << 15, pushbutton: 1 << 16, combo: 1 << 17, edit: 1 << 18, multiSelect: 1 << 21, comb: 1 << 24,
};

const num = (v, fallback = 0) => (typeof v === 'number' && Number.isFinite(v) ? v : fallback);
const n = (v) => {
  if (!Number.isFinite(v)) return '0';
  const s = (Math.round(v * 1000) / 1000).toString();
  return s === '-0' ? '0' : s;
};

/** A field's own value or one it inherits, walking up its parents. */
function inherited(file, dict, key) {
  let node = dict;
  for (let i = 0; i < 64 && node instanceof Dict; i++) {
    if (node.has(key)) return file.resolve(node.get(key));
    node = file.resolve(node.get('Parent'));
  }
  return undefined;
}

const textOf = (v) => (v instanceof PdfString ? v.text() : v instanceof Name ? v.name : typeof v === 'string' ? v : '');

/** The names of a widget's "on" appearances: every normal appearance but Off. */
function onStates(file, widget) {
  const ap = file.resolve(widget.get('AP'));
  const normal = ap instanceof Dict ? file.resolve(ap.get('N')) : null;
  if (!(normal instanceof Dict) || normal.get('Subtype')) return [];
  return [...normal.keys()].filter((k) => k !== 'Off');
}

/** Each widget's page, by the pages' own lists of annotations. */
function widgetPages(file) {
  const map = new Map();
  file.pages().forEach((page, p) => {
    const list = file.resolve(page.dict.get('Annots'));
    if (!Array.isArray(list)) return;
    for (const item of list) if (item instanceof Ref) map.set(item.num, p);
  });
  return map;
}

/**
 * The form's fields, each a plain value a window can hold:
 * { id, name, type: 'text'|'checkbox'|'radio'|'button'|'combo'|'list'|'signature',
 *   value, options: [{ value, label }], multiline, password, readOnly, required,
 *   multiSelect, maxLength, align, fontSize,
 *   widgets: [{ id, page, rect, onState, state }] }.
 * `id` is the field dictionary's object number; widgets on pages the file
 * does not list are left out.
 */
function readFields(file) {
  const catalog = file.catalogOrNull();
  const form = catalog ? file.resolve(catalog.get('AcroForm')) : null;
  if (!(form instanceof Dict)) return [];
  const roots = file.resolve(form.get('Fields'));
  if (!Array.isArray(roots)) return [];
  const pages = widgetPages(file);
  const out = [];
  const seen = new Set();

  const visit = (ref, prefix, depth) => {
    if (depth > 32) return;
    if (ref instanceof Ref) {
      if (seen.has(ref.num)) return;
      seen.add(ref.num);
    }
    const dict = file.resolve(ref);
    if (!(dict instanceof Dict)) return;
    const partial = textOf(file.resolve(dict.get('T')));
    const name = partial ? (prefix ? `${prefix}.${partial}` : partial) : prefix;
    const kids = file.resolve(dict.get('Kids'));
    const fieldKids = Array.isArray(kids) ? kids.filter((k) => { const d = file.resolve(k); return d instanceof Dict && d.has('T'); }) : [];
    if (fieldKids.length) {
      for (const k of fieldKids) visit(k, name, depth + 1);
      return;
    }
    if (!(ref instanceof Ref)) return;
    const widgetRefs = Array.isArray(kids) && kids.length ? kids : [ref];
    const ft = inherited(file, dict, 'FT');
    const ff = num(inherited(file, dict, 'Ff'));
    const kind = ft instanceof Name ? ft.name : '';
    let type = 'text';
    if (kind === 'Btn') type = ff & FF.pushbutton ? 'button' : ff & FF.radio ? 'radio' : 'checkbox';
    else if (kind === 'Ch') type = ff & FF.combo ? 'combo' : 'list';
    else if (kind === 'Sig') type = 'signature';
    else if (kind !== 'Tx') return;
    const v = inherited(file, dict, 'V');
    let value;
    if (type === 'checkbox' || type === 'radio') value = v instanceof Name ? v.name : textOf(v) || 'Off';
    else if (type === 'list' && Array.isArray(v)) value = v.map((x) => textOf(file.resolve(x)));
    else value = textOf(v);
    const opt = file.resolve(inherited(file, dict, 'Opt'));
    const options = Array.isArray(opt) ? opt.map((o) => {
      const item = file.resolve(o);
      if (Array.isArray(item)) return { value: textOf(file.resolve(item[0])), label: textOf(file.resolve(item[1])) };
      return { value: textOf(item), label: textOf(item) };
    }) : [];
    const da = textOf(inherited(file, dict, 'DA')) || textOf(file.resolve(form.get('DA')));
    const size = /([\d.]+)\s+Tf/.exec(da);
    const widgets = [];
    for (const w of widgetRefs) {
      const wd = file.resolve(w);
      if (!(wd instanceof Dict) || !(w instanceof Ref)) continue;
      const page = pages.get(w.num);
      if (page === undefined) continue;
      const r = file.resolve(wd.get('Rect'));
      const rect = Array.isArray(r) && r.length >= 4 ? (() => { const x = r.map((q) => num(file.resolve(q))); return [Math.min(x[0], x[2]), Math.min(x[1], x[3]), Math.max(x[0], x[2]), Math.max(x[1], x[3])]; })() : null;
      const as = wd.get('AS');
      widgets.push({ id: String(w.num), page, rect, onState: onStates(file, wd)[0] || null, state: as instanceof Name ? as.name : null });
    }
    out.push({
      id: String(ref.num),
      name,
      type,
      value,
      options,
      multiline: Boolean(ff & FF.multiline),
      password: Boolean(ff & FF.password),
      readOnly: Boolean(ff & FF.readOnly),
      required: Boolean(ff & FF.required),
      multiSelect: Boolean(ff & FF.multiSelect),
      editable: Boolean(ff & FF.edit),
      maxLength: num(inherited(file, dict, 'MaxLen'), 0) || null,
      align: num(inherited(file, dict, 'Q'), 0),
      fontSize: size ? Number(size[1]) : 0,
      widgets,
    });
  };
  for (const r of roots) visit(r, '', 0);
  return out;
}

// ── appearances ───────────────────────────────────────────────────────────

/** The font, size and colour a /DA string names. */
function parseDA(da) {
  const font = /\/([^\s/]+)\s+([\d.]+)\s+Tf/.exec(da || '');
  const colour = /((?:[\d.]+\s+){0,3}[\d.]+)\s+(rg|g|k)\b/.exec(da || '');
  return { font: font ? font[1] : 'Helv', size: font ? Number(font[2]) : 0, colour: colour ? `${colour[1]} ${colour[2]}` : '0 g' };
}

/** Lines of words that fit `width` points at `size` in Helvetica's widths. */
function wrap(text, size, width) {
  const out = [];
  for (const para of String(text).split(/\r\n|\r|\n/)) {
    let line = '';
    for (const w of para.split(/(\s+)/).filter(Boolean)) {
      const trial = line + w;
      if (line && widthOfBytes(encode(trial.trimEnd()), 'Helvetica', size) > width) {
        out.push(line.trimEnd());
        line = /^\s+$/.test(w) ? '' : w;
      } else line = trial;
    }
    out.push(line.trimEnd());
  }
  return out;
}

/**
 * A text or choice widget's normal appearance: its words in the field's
 * font, inside a 2-point margin, as /Tx BMC ... EMC so a reader that draws
 * fields itself knows where the variable text is.
 */
function textAppearance({ width, height, text, da, align = 0, multiline = false, comb = 0, list = null }) {
  const { font, size: daSize, colour } = parseDA(da);
  const pad = 2;
  const ops = ['/Tx BMC', 'q', `${n(pad / 2)} ${n(pad / 2)} ${n(width - pad)} ${n(height - pad)} re W n`];
  if (list) {
    // A list box: its options one a line, those chosen on a blue band, as Acrobat draws one.
    const size = daSize || 12;
    const lead = size * 1.15;
    list.options.forEach((o, i) => {
      const top = height - pad - i * lead;
      if (list.chosen.includes(o.value)) ops.push(`0.6 0.75 0.86 rg ${n(pad / 2)} ${n(top - lead)} ${n(width - pad)} ${n(lead)} re f`);
    });
    ops.push('BT', `/${font} ${n(size)} Tf`, colour);
    list.options.forEach((o, i) => {
      const y = height - pad - (i + 1) * lead + size * 0.25;
      ops.push(`1 0 0 1 ${n(pad + 1)} ${n(y)} Tm ${literal(encode(o.label))} Tj`);
    });
    ops.push('ET', 'Q', 'EMC');
    return { content: ops.join('\n'), font };
  }
  let size = daSize;
  if (!size) {
    // Auto size: as tall as the box allows, and narrow enough for one line.
    size = multiline ? 12 : Math.max(4, Math.min(12, (height - pad * 2) * 0.72));
    if (!multiline) {
      const w = widthOfBytes(encode(text), 'Helvetica', 1);
      if (w > 0) size = Math.max(4, Math.min(size, (width - pad * 4) / w));
    }
  }
  ops.push('BT', `/${font} ${n(size)} Tf`, colour);
  if (comb > 0) {
    const cell = width / comb;
    [...String(text)].slice(0, comb).forEach((ch, i) => {
      const w = widthOfBytes(encode(ch), 'Helvetica', size);
      ops.push(`1 0 0 1 ${n(cell * i + (cell - w) / 2)} ${n((height - size * 0.7) / 2)} Tm ${literal(encode(ch))} Tj`);
    });
  } else {
    const lines = multiline ? wrap(text, size, width - pad * 4) : [String(text).replace(/\r?\n/g, ' ')];
    const lead = size * 1.15;
    lines.forEach((line, i) => {
      const w = widthOfBytes(encode(line), 'Helvetica', size);
      const x = align === 1 ? (width - w) / 2 : align === 2 ? width - pad * 2 - w : pad * 2;
      const y = multiline ? height - pad - size * 0.95 - i * lead : (height - size * 0.72) / 2;
      ops.push(`1 0 0 1 ${n(x)} ${n(y)} Tm ${literal(encode(line))} Tj`);
    });
  }
  ops.push('ET', 'Q', 'EMC');
  return { content: ops.join('\n'), font };
}

// ── filling ───────────────────────────────────────────────────────────────

/**
 * Fill a PDF's form, as one incremental update.
 *
 * @param {Buffer|Uint8Array} bytes the whole file
 * @param {Object<string, string|boolean|string[]>} values by field id or full name:
 *   text for a text field or a combo box, true/false or the state's name for
 *   a check box, the chosen state's name for a radio group, a value or a
 *   list of values for a list box
 * @returns {{ bytes: Buffer, changed: string[] }} the file with the update, and the ids of the fields changed
 */
function fillFields(bytes, values = {}) {
  const up = new PdfUpdate(bytes, { what: 'A form value' });
  const { file } = up;
  const fields = readFields(file);
  const catalog = file.catalog();
  const formRef = catalog.get('AcroForm');
  const form = file.resolve(formRef);
  if (!(form instanceof Dict)) throw new UpdateError('This PDF has no form to fill.', 'missing');
  const dr = file.resolve(form.get('DR'));
  const drFonts = dr instanceof Dict ? file.resolve(dr.get('Font')) : null;
  let needAppearances = false;
  const changed = [];
  // Objects as this update leaves them, so a field that is its own widget is written once.
  const pending = new Map();
  const dictFor = (num) => {
    if (!pending.has(num)) pending.set(num, copyDict(file.getObject(num)));
    return pending.get(num);
  };
  let helv = null;
  const fontRef = (name) => {
    const own = drFonts instanceof Dict ? drFonts.get(name) : null;
    if (own) return own;
    if (!helv) {
      helv = up.add(new Dict([['Type', new Name('Font')], ['Subtype', new Name('Type1')], ['BaseFont', new Name('Helvetica')], ['Encoding', new Name('WinAnsiEncoding')]]));
    }
    return helv;
  };

  for (const [key, wanted] of Object.entries(values)) {
    const field = fields.find((f) => f.id === String(key)) || fields.find((f) => f.name === key);
    if (!field) throw new UpdateError(`This form has no field named ${key}.`, 'missing');
    if (field.readOnly || field.type === 'button' || field.type === 'signature') continue;
    const fieldDict = dictFor(Number(field.id));
    const da = textOf(inherited(file, file.getObject(Number(field.id)), 'DA')) || textOf(file.resolve(form.get('DA'))) || '/Helv 0 Tf 0 g';

    if (field.type === 'checkbox' || field.type === 'radio') {
      // The state each widget is turned to: its own "on" name when chosen, else Off.
      let state;
      if (field.type === 'checkbox') {
        const on = field.widgets.find((w) => w.onState)?.onState || 'Yes';
        state = wanted === true || wanted === on || wanted === 'Yes' || wanted === 'On' ? on : 'Off';
      } else {
        state = wanted && field.widgets.some((w) => w.onState === wanted) ? String(wanted) : 'Off';
      }
      fieldDict.set('V', new Name(state));
      for (const w of field.widgets) {
        const wd = w.id === field.id ? fieldDict : dictFor(Number(w.id));
        wd.set('AS', new Name(w.onState && w.onState === state ? state : 'Off'));
      }
      changed.push(field.id);
      continue;
    }

    let text;
    let list = null;
    if (field.type === 'list') {
      const chosen = (Array.isArray(wanted) ? wanted : [wanted]).filter((x) => x !== null && x !== undefined && x !== '').map(String);
      fieldDict.set('V', chosen.length > 1 ? chosen.map((c) => textString(c)) : textString(chosen[0] || ''));
      fieldDict.set('I', chosen.map((c) => field.options.findIndex((o) => o.value === c)).filter((i) => i >= 0).sort((a, b) => a - b));
      list = { options: field.options, chosen };
      text = chosen.join(', ');
    } else {
      text = String(wanted ?? '');
      if (field.maxLength) text = text.slice(0, field.maxLength);
      fieldDict.set('V', textString(text));
      if (field.type === 'combo') {
        const idx = field.options.findIndex((o) => o.value === text);
        if (idx >= 0) fieldDict.set('I', [idx]); else fieldDict.delete('I');
      }
    }
    // The words shown: a combo box shows its option's label.
    const shown = field.type === 'combo' ? (field.options.find((o) => o.value === text)?.label ?? text) : field.password ? '*'.repeat(text.length) : text;
    const drawable = fitsWinAnsi(list ? list.options.map((o) => o.label).join(' ') : shown);
    const ff = num(inherited(file, file.getObject(Number(field.id)), 'Ff'));
    for (const w of field.widgets) {
      const wd = w.id === field.id ? fieldDict : dictFor(Number(w.id));
      if (!drawable) { wd.delete('AP'); continue; }
      const width = w.rect ? w.rect[2] - w.rect[0] : 100;
      const height = w.rect ? w.rect[3] - w.rect[1] : 20;
      const look = textAppearance({ width, height, text: shown, da: textOf(file.resolve(wd.get('DA'))) || da, align: field.align, multiline: field.multiline, comb: ff & FF.comb ? field.maxLength || 0 : 0, list });
      const stream = flateStream([
        ['Type', new Name('XObject')], ['Subtype', new Name('Form')], ['BBox', [0, 0, Math.round(width * 1000) / 1000, Math.round(height * 1000) / 1000]],
        ['Resources', new Dict([['Font', new Dict([[look.font, fontRef(look.font)]])]])],
      ], look.content);
      wd.set('AP', new Dict([['N', up.add(stream)]]));
    }
    if (!drawable) needAppearances = true;
    changed.push(field.id);
  }
  for (const [num, dict] of pending) up.set(num, dict);
  if (needAppearances && form.get('NeedAppearances') !== true) {
    const next = copyDict(form);
    next.set('NeedAppearances', true);
    if (formRef instanceof Ref) up.set(formRef, next);
    else { const cat = copyDict(catalog); cat.set('AcroForm', next); up.set(file.trailer.get('Root'), cat); }
  }
  return { bytes: up.toBuffer(), changed };
}

module.exports = { readFields, fillFields, textAppearance, parseDA, FF };
