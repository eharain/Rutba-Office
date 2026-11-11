// Find in Sheets: what the box and the Find and Replace panel read back
// from the engine, and which keys open them. No React here, so the rules
// are tested on their own (tests/sheet-find.test.js).

/** The engine's answer to a step: { found, index, total, sheet, ref }. */
export function readFindResult(opResult) {
  let r = null;
  try { r = typeof opResult === 'string' ? JSON.parse(opResult) : opResult; } catch { r = null; }
  if (!r || typeof r !== 'object' || !r.found) return { found: false, index: 0, total: 0, sheet: r?.sheet ?? null, ref: null };
  return { found: true, index: Number(r.index) || 0, total: Number(r.total) || 0, sheet: r.sheet ?? null, ref: r.ref ?? null };
}

/** Find All's list: [{ sheet, row, col, ref, value, formula }]. */
export function readFindList(opResult) {
  let list = null;
  try { list = typeof opResult === 'string' ? JSON.parse(opResult) : opResult; } catch { list = null; }
  return Array.isArray(list) ? list.filter((m) => m && Number.isInteger(m.row) && Number.isInteger(m.col)) : [];
}

/** What the box says after a step: "3 of 12", or that nothing matched. `say` is t(). */
export function findCountText(result, say) {
  if (!result?.found || !result.total) return say('No matches');
  return say('{index} of {total}', { index: result.index, total: result.total });
}

/** The choices sent with each search, as the engine reads them. */
export function findOptions({ matchCase = false, wholeCell = false, within = 'sheet' } = {}) {
  return { matchCase: Boolean(matchCase), wholeCell: Boolean(wholeCell), within: within === 'workbook' ? 'workbook' : 'sheet' };
}

/**
 * Ctrl+F is Find and Ctrl+H is Replace, as in Excel (Cmd on a Mac); with
 * Alt or Shift held they are something else.
 */
export function findKey(e) {
  if (!e || !(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return null;
  const key = String(e.key || '').toLowerCase();
  return key === 'f' ? 'find' : key === 'h' ? 'replace' : null;
}

/** The operations that go to a Find All row: its sheet first when it is another, then its cell. */
export function goToMatch(match, activeSheet) {
  if (!match) return [];
  const ops = [];
  if (match.sheet && match.sheet !== activeSheet) ops.push({ op: 'sheet', name: match.sheet });
  ops.push({ op: 'select', row: match.row, col: match.col });
  return ops;
}
