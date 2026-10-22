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

export const GRID = 8;
const LABEL_W = 136;
const ROW_H = 32;
const GAP = 8;

const KINDS = { text: 'text', longtext: 'longtext', number: 'number', integer: 'number', currency: 'currency', date: 'date', datetime: 'datetime', yesno: 'check', autonumber: 'auto' };
const WIDTHS = { text: 280, longtext: 400, number: 120, currency: 140, date: 150, datetime: 200, check: 24, list: 280, auto: 90 };
const snap = (v) => Math.round(v / GRID) * GRID;

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
    const rel = relationships.find((r) => r.from.table === table && r.from.field === f.name);
    const kind = rel ? 'list' : KINDS[f.type] || 'text';
    const h = kind === 'longtext' ? ROW_H * 3 - GAP : ROW_H - GAP;
    controls.push({
      kind,
      field: f.name,
      label: f.caption || labelFor(f.name),
      x: LABEL_W + GRID * 2,
      y,
      w: WIDTHS[kind],
      h,
      ...(kind === 'auto' ? { readOnly: true } : {}),
      ...(rel ? { lookup: { table: rel.to.table, field: rel.to.field, show: shownField(fieldsOf(rel.to.table)) } } : {}),
    });
    y += h + GAP;
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
