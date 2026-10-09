// Worksheets: Record Actions — the script lines written for what the
// window sends while recording. Kept apart from the pane so the engine's
// tests can read it without a window.

const q = (s) => JSON.stringify(String(s));

/**
 * Record Actions: a script line for each operation the window sends, where
 * one is known — a cell typed into, a format set, another sheet chosen —
 * given the selection it applied to. Others are left out.
 */
export function recordLines(ops, model) {
  const out = [];
  const sel = model?.selection;
  const a1 = (row, col) => { let s = ''; for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s + (row + 1); };
  const rangeOf = () => (sel ? (sel.top === sel.bottom && sel.left === sel.right ? a1(sel.top, sel.left) : `${a1(sel.top, sel.left)}:${a1(sel.bottom, sel.right)}`) : 'A1');
  let draft = null;
  for (const op of ops) {
    if (op.op === 'updateDraft') draft = op.text;
    else if (op.op === 'commitEdit' && draft != null && sel?.active) {
      const at = a1(sel.active.row, sel.active.col);
      out.push(String(draft).startsWith('=') ? `  sheet.getRange(${q(at)}).setFormula(${q(draft)});` : `  sheet.getRange(${q(at)}).setValue(${/^-?\d+(\.\d+)?$/.test(String(draft).trim()) ? Number(draft) : q(draft)});`);
      draft = null;
    } else if (op.op === 'setFormat' && op.delta) {
      const d = op.delta;
      const r = `sheet.getRange(${q(rangeOf())}).getFormat()`;
      if (d.bold !== undefined) out.push(`  ${r}.getFont().setBold(${d.bold === 'toggle' ? 'true' : Boolean(d.bold)});`);
      if (d.italic !== undefined) out.push(`  ${r}.getFont().setItalic(${d.italic === 'toggle' ? 'true' : Boolean(d.italic)});`);
      if (d.fontColour) out.push(`  ${r}.getFont().setColor(${q(d.fontColour)});`);
      if (d.fill) out.push(`  ${r}.getFill().setColor(${q(d.fill)});`);
      if (d.fontSize) out.push(`  ${r}.getFont().setSize(${Number(d.fontSize)});`);
      if (d.numberFormat) out.push(`  sheet.getRange(${q(rangeOf())}).setNumberFormat(${q(d.numberFormat)});`);
      if (d.align) out.push(`  ${r}.setHorizontalAlignment(${q(d.align)});`);
      if (d.underline !== undefined) out.push(`  ${r}.getFont().setUnderline(${d.underline === 'toggle' || d.underline ? '"Single"' : '"None"'});`);
      if (d.strike !== undefined) out.push(`  ${r}.getFont().setStrikethrough(${d.strike === 'toggle' ? 'true' : Boolean(d.strike)});`);
      if (d.fontName) out.push(`  ${r}.getFont().setName(${q(d.fontName)});`);
      if (d.wrap !== undefined) out.push(`  ${r}.setWrapText(${Boolean(d.wrap)});`);
      if (d.valign) out.push(`  ${r}.setVerticalAlignment(${q(d.valign)});`);
    } else if (op.op === 'sheet' && op.name) {
      out.push(`  sheet = workbook.getWorksheet(${q(op.name)});`, '  sheet.activate();');
    } else if (op.op === 'colWidth' && Number.isInteger(op.col)) {
      // A width or height in points, as Office Scripts take them.
      out.push(`  sheet.getRange(${q(a1(0, op.col))}).getFormat().setColumnWidth(${Math.round((Number(op.width) || 0) * 0.75)});`);
    } else if (op.op === 'rowHeight' && Number.isInteger(op.row)) {
      out.push(`  sheet.getRange(${q(a1(op.row, 0))}).getFormat().setRowHeight(${Math.round((Number(op.height) || 0) * 0.75)});`);
    } else if (op.op === 'merge') {
      out.push(`  sheet.getRange(${q(rangeOf())}).merge();`);
    } else if (op.op === 'unmerge') {
      out.push(`  sheet.getRange(${q(rangeOf())}).unmerge();`);
    } else if (op.op === 'clear') {
      out.push(`  sheet.getRange(${q(rangeOf())}).clear();`);
    } else if (op.op === 'addSheet') {
      out.push(`  workbook.addWorksheet(${op.name ? q(op.name) : ''});`);
    } else if (op.op === 'renameSheet' && op.from && op.to) {
      out.push(`  workbook.getWorksheet(${q(op.from)}).setName(${q(op.to)});`);
    }
  }
  return out;
}

/** The script Record Actions writes from its lines. */
export const recordedScript = (lines) => `function main(workbook) {\n  let sheet = workbook.getActiveWorksheet();\n${lines.join('\n')}\n}\n`;
