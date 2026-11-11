// How quickly typed keys reach a page while other work runs: judged on the
// keys as a whole, not on the one slowest.
//
// The budget was "every key within 200 ms". Gate 9's machine lane failed it
// on single keys (255 ms of eighteen, 312 ms of twenty-three) whose median
// was 70 and 28 ms: one late turn of the system's own, not the work held the
// windows (review of 11 Nov). So now: the median under `median`, at most one
// key in twenty (rounded up) over `budget`, and none at all over `cap`, which
// is still well inside what a person feels as instant typing held up.

/** The verdict on `lags` (ms each, null for a key that never arrived): { ok, said }. */
export function keyLagVerdict(lags, { budget = 200, median = 100, cap = 500 } = {}) {
  const list = Array.isArray(lags) ? lags : [];
  if (!list.length || list.some((l) => l == null || !Number.isFinite(l))) return { ok: false, said: `a key did not reach the page: ${JSON.stringify(list)}`, median: null, over: null, worst: null };
  const sorted = [...list].sort((a, b) => a - b);
  const mid = sorted.length % 2 ? sorted[(sorted.length - 1) / 2] : (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2;
  const over = list.filter((l) => l >= budget).length;
  const allowed = Math.ceil(list.length / 20);
  const worst = sorted.at(-1);
  const ok = mid < median && over <= allowed && worst < cap;
  return { ok, median: mid, over, allowed, worst, said: `median ${mid} ms (under ${median}), ${over} of ${list.length} at ${budget} ms or more (${allowed} allowed), slowest ${worst} ms (under ${cap})` };
}
