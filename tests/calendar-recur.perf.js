/**
 * Performance budget: a repeating rule that never lands on a day (the 30th
 * of February, every seventh day from a Tuesday but only on Mondays) gives
 * up at once rather than walking for ever. tests/calendar.test.js holds how
 * far the walk goes as a count; this holds the time, so it runs alone, with
 * `npm run test:perf`, where other work does not slow it.
 *
 * Measured on a four-core laptop (2026-11-11), alone: the longest, every
 * seventh day with no end to the range, 63 ms; the rest under 7 ms. Under
 * the full unit run in parallel the same walk took 2.1 s. The budget is
 * generous on purpose: it exists to catch a walk gone without an end, not
 * to time a machine.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readCalendar, occurrences } from '../packages/calendar/src/index.js';

const CRLF = (s) => s.replace(/\r?\n/g, '\r\n');
const BUDGET_MS = 1500;

test('a rule that never lands on a day gives up well inside its budget', () => {
  for (const rule of ['FREQ=MONTHLY;BYMONTHDAY=30;BYMONTH=2', 'FREQ=YEARLY;BYMONTH=4;BYMONTHDAY=31', 'FREQ=DAILY;INTERVAL=7;BYDAY=MO', 'FREQ=MONTHLY;INTERVAL=12;BYMONTH=6']) {
    const event = readCalendar(CRLF(`BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:never
DTSTART:20260106T090000Z
DTEND:20260106T100000Z
RRULE:${rule}
SUMMARY:Never
END:VEVENT
END:VCALENDAR
`)).events[0];
    for (const range of [{ from: Date.UTC(2026, 0, 1), to: Date.UTC(2026, 11, 31) }, {}]) {
      const t0 = performance.now();
      const found = occurrences(event, range);
      const ms = Math.round(performance.now() - t0);
      assert.deepEqual(found, []);
      assert.ok(ms < BUDGET_MS, `${rule} gave up in ${ms} ms${range.to ? '' : ' with no end to the range'}`);
    }
  }
});
