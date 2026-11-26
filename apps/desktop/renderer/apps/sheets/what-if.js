/**
 * Data → What-If Analysis: what the engine's answer says in the dialog.
 */

/** A number as the result line says it: two decimals at most, grouped once it is in the thousands. */
const said = (n) => (Math.abs(n) >= 1000
  ? n.toLocaleString(undefined, { maximumFractionDigits: 2 })
  : String(Math.round(n * 100) / 100));

/**
 * Goal Seek's result line, from what the operation sent back: the value put
 * in the changing cell and the value the formula now reaches. With nothing
 * sent back (the search was refused or found nothing) it did not converge.
 * The line used to show the changing cell's name where its value belongs.
 */
export function goalSeekSaid(opResult, to) {
  let found = null;
  try { found = JSON.parse(opResult || 'null')?.goalSeek ?? null; } catch { found = null; }
  if (!found || typeof found.value !== 'number' || !Number.isFinite(found.value)) return { converged: false, reached: to };
  return {
    converged: true,
    value: said(found.value),
    reached: typeof found.reached === 'number' && Number.isFinite(found.reached) ? said(found.reached) : to,
  };
}
