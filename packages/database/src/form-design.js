// Forms in Data Files: a record at a time, laid out on a grid. The form
// wizard makes a good one from a table in one step: a label and a control
// for each field, the control the field's type calls for, a field that
// points at another table's record shown as a list of that table's records
// by their first words, and the records of other tables that point at this
// one in a subform below. A form is kept in the file as its layout:
//
// { name, table, title, grid: 8, width,
//   controls: [{ kind, field, label, x, y, w, h, readOnly?, lookup?: { table, field, show } }],
//   subforms: [{ table, field, columns: [field] , x, y, w, h }] }
//
// `kind` is 'text', 'longtext', 'number', 'currency', 'date', 'datetime',
// 'check', 'list' (a lookup) or 'auto' (an AutoNumber, shown, not typed).
// A control's label stands LABEL_W to the left of it, so it moves with it.
//
// A form's layout is changed in the window's Design: a control moved and
// resized on the grid (placeControl), taken off (takeOff) or put back
// (putBack), its label renamed (relabel); the form's size follows (fitForm).

export const GRID = 8;
const LABEL_W = 136;
const ROW_H = 32;
const GAP = 8;

const KINDS = { text: 'text', longtext: 'longtext', number: 'number', integer: 'number', currency: 'currency', date: 'date', datetime: 'datetime', yesno: 'check', autonumber: 'auto' };
const WIDTHS = { text: 280, longtext: 400, number: 120, currency: 140, date: 150, datetime: 200, check: 24, list: 280, auto: 90 };
const snap = (v) => Math.round(v / GRID) * GRID;

/** Where a control's label stands: LABEL_W to its left, as the wizard puts it. */
export const labelLeft = (c) => Math.max(0, c.x - LABEL_W);
/** The width a label has: up to its control, less a gap. */
export const LABEL_WIDTH = LABEL_W - GRID;

/** A field's name as a label: "first_name" and "FirstName" read "First name". */
export function labelFor(name) {
  // An underscore is a space; a hyphen inside a word stays (E-mail, Follow-up), one at either end goes.
  const words = String(name).replace(/_+/g, ' ').replace(/(^|\s)-+|-+(\s|$)/g, ' ').replace(/([a-z])([A-Z])/g, '$1 $2').trim().split(/\s+/).filter(Boolean);
  if (!words.length) return String(name);
  // A word in capitals (ID, VAT, UK) stays so; the others read as a sentence.
  return words.map((w, i) => (/^[A-Z0-9:]+$/.test(w) && w.length > 1 ? w : i ? w.toLowerCase() : w[0].toUpperCase() + w.slice(1).toLowerCase())).join(' ');
}

/**
 * The field a table's records are best known by: its Name or Title, or a
 * field whose name ends in Name; else its first text field that is not a
 * code (a barcode, an SKU, a reference); else its first text field, else its
 * first field. A person's First Name and Last Name are both, as a list of
 * two, so a borrower is Amira Haddad and not Amira.
 */
export function shownField(fields) {
  const named = (re) => fields.find((f) => f.type === 'text' && re.test(f.name))?.name;
  const first = named(/^(first|given)[ _]?name$/i);
  const last = named(/^(last|family)[ _]?name$|^surname$/i);
  if (first && last) return [first, last];
  const code = /(^|[ _])(bar ?code|code|sku|ref|reference|number|no|id|plate|registration|reg)$/i;
  return named(/^(name|title)$/i) || named(/[ _]?name$/i)
    || fields.find((f) => f.type === 'text' && !code.test(f.name))?.name
    || (fields.find((f) => f.type === 'text') || fields.find((f) => f.type !== 'autonumber') || fields[0])?.name;
}

/**
 * A form for `table`, made as the wizard makes it. `fieldsOf(name)` answers a
 * table's fields; `relationships` are the file's.
 */
export function wizardForm(table, fieldsOf, relationships = [], { name = null } = {}) {
  const fields = fieldsOf(table);
  if (!fields.length) throw new Error(`${table} has no fields to put on a form.`);
  const controls = [];
  let y = GRID * 2;
  for (const f of fields) {
    const c = controlFor(table, f, fieldsOf, relationships, y);
    controls.push(c);
    y += c.h + GAP;
  }
  // The records elsewhere that point at this one: a subform each, below.
  const subforms = relationships.filter((r) => r.to.table === table).map((r) => {
    const cols = fieldsOf(r.from.table).filter((f) => f.name !== r.from.field && f.type !== 'autonumber').slice(0, 5).map((f) => f.name);
    const sub = { table: r.from.table, field: r.from.field, to: r.to.field, columns: cols, label: labelFor(r.from.table), x: GRID * 2, y: snap(y + GRID * 2), w: 560, h: 200 };
    y = sub.y + sub.h + GAP;
    return sub;
  });
  const width = snap(Math.max(LABEL_W + GRID * 4 + Math.max(...controls.map((c) => c.w)), ...subforms.map((s) => s.w + GRID * 4)));
  return { name: name || table, table, title: labelFor(table), grid: GRID, width, height: snap(y + GRID * 2), controls, subforms };
}

/** The control the wizard makes for field `f` of `table`, its top at `y`. */
function controlFor(table, f, fieldsOf, relationships, y) {
  const rel = relationships.find((r) => r.from.table === table && r.from.field === f.name);
  const kind = rel ? 'list' : KINDS[f.type] || 'text';
  return {
    kind,
    field: f.name,
    label: f.caption || labelFor(f.name),
    x: LABEL_W + GRID * 2,
    y,
    w: WIDTHS[kind],
    h: kind === 'longtext' ? ROW_H * 3 - GAP : ROW_H - GAP,
    ...(kind === 'auto' ? { readOnly: true } : {}),
    ...(rel ? { lookup: { table: rel.to.table, field: rel.to.field, show: shownField(fieldsOf(rel.to.table)) } } : {}),
  };
}

const bottom = (layout) => Math.max(0, ...(layout.controls || []).map((c) => c.y + c.h), ...(layout.subforms || []).map((s) => s.y + s.h));

/** The form's width and height worked out again from what is on it, as the wizard works them out. */
export function fitForm(layout) {
  const right = Math.max(LABEL_W + GRID * 4, ...(layout.controls || []).map((c) => c.x + c.w + GRID * 2), ...(layout.subforms || []).map((s) => s.x + s.w + GRID * 2));
  return { ...layout, width: snap(right), height: snap(bottom(layout) + GAP + GRID * 2) };
}

/**
 * The layout with `field`'s control at `x`, `y`, `w` by `h` (any of them;
 * the rest as they were), snapped to the grid: room left for its label, a
 * tick box its own size, nothing smaller than the grid allows.
 */
export function placeControl(layout, field, to = {}) {
  const c = layout.controls.find((x) => x.field === field);
  if (!c) throw new Error(`${field} is not on the form.`);
  const next = { ...c };
  if (to.x !== undefined) next.x = Math.max(LABEL_W, snap(to.x));
  if (to.y !== undefined) next.y = Math.max(0, snap(to.y));
  if (c.kind !== 'check') {
    if (to.w !== undefined) next.w = Math.min(1600, Math.max(GRID * 4, snap(to.w)));
    if (to.h !== undefined) next.h = Math.min(800, Math.max(GRID * 2, snap(to.h)));
  }
  return fitForm({ ...layout, controls: layout.controls.map((x) => (x.field === field ? next : x)) });
}

/** The layout without `field`'s control: the field stays in the table, only the form no longer shows it. */
export function takeOff(layout, field) {
  if (!layout.controls.some((c) => c.field === field)) throw new Error(`${field} is not on the form.`);
  if (layout.controls.length === 1) throw new Error('A form shows one field at least. Put another on before taking this one off.');
  return fitForm({ ...layout, controls: layout.controls.filter((c) => c.field !== field) });
}

/** The fields of the form's table that it does not show, to put back. */
export function fieldsOff(layout, fields) {
  const on = new Set(layout.controls.map((c) => c.field));
  return fields.filter((f) => !on.has(f.name));
}

/**
 * The layout with field `f` (one of the table's) put back on: below the
 * other fields, as the wizard puts it, the subforms below moved down.
 */
export function putBack(layout, f, fieldsOf, relationships = []) {
  if (layout.controls.some((c) => c.field === f.name)) throw new Error(`${f.name} is on the form already.`);
  const y = snap(Math.max(GRID * 2 - GAP, ...layout.controls.map((c) => c.y + c.h)) + GAP);
  const c = controlFor(layout.table, f, fieldsOf, relationships, y);
  // A subform it would stand on moves down, the others with it.
  const clash = (layout.subforms || []).filter((s) => s.y < y + c.h + GAP).map((s) => y + c.h + GAP + GRID * 2 - s.y);
  const down = clash.length ? snap(Math.max(...clash)) : 0;
  const subforms = (layout.subforms || []).map((s) => (down ? { ...s, y: s.y + down } : s));
  return fitForm({ ...layout, controls: [...layout.controls, c], subforms });
}

/** The layout with `field`'s label reading `label`; an empty one reads as the field's name does. */
export function relabel(layout, field, label) {
  if (!layout.controls.some((c) => c.field === field)) throw new Error(`${field} is not on the form.`);
  const words = String(label ?? '').trim().slice(0, 80) || labelFor(field);
  return { ...layout, controls: layout.controls.map((c) => (c.field === field ? { ...c, label: words } : c)) };
}

/** A value from a form's control as the table takes it: text trimmed, numbers as numbers, a check as yes or no, empty as nothing. */
export function valueFromControl(kind, raw) {
  if (kind === 'check') return Boolean(raw);
  if (raw === '' || raw === null || raw === undefined) return null;
  if (kind === 'number' || kind === 'currency' || kind === 'list') {
    const n = Number(String(raw).replace(/,/g, ''));
    return Number.isFinite(n) ? n : raw;
  }
  return typeof raw === 'string' ? raw.trimEnd() : raw;
}
