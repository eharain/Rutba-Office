// Send later's own presets — Outlook and Gmail's short list, computed from
// whatever moment is handed in rather than from the clock, which is what lets
// an engine test see "Monday morning" land correctly on a Saturday without
// waiting for one.
//
// Nothing here touches a store, a window or the network: it turns "now" into
// a short list of { label, at }, and the composer offers them as buttons. A
// person can always go around this list with their own date and time instead
// — see `parseCustomSchedule` — which is the one choice this list cannot
// enumerate for them.

/**
 * The short list a compose window offers under Send later, in local time —
 * the zone `base` (or the real clock, if none is given) is already in.
 *
 * "Later today" only appears when there is a later today left to offer: past
 * 9pm, tomorrow morning is the soonest sensible choice. "Monday morning" is
 * always the *next* Monday, including when today already is one — Send
 * later is for later, and a button that can mean "later today" wearing a
 * Monday label would be a trap.
 */
export function scheduleChoices(base = new Date()) {
  const now = new Date(base);
  const at = (days, hour) => {
    const d = new Date(now);
    d.setDate(d.getDate() + days);
    d.setHours(hour, 0, 0, 0);
    return d;
  };

  const options = [];
  const tonight = at(0, 21);
  if (tonight > now) options.push({ id: 'tonight', label: 'Later today, 9pm', at: tonight });
  options.push({ id: 'tomorrow-morning', label: 'Tomorrow morning, 8am', at: at(1, 8) });
  options.push({ id: 'tomorrow-afternoon', label: 'Tomorrow afternoon, 1pm', at: at(1, 13) });

  const monday = new Date(now);
  // Days until the *next* Monday: 0 only lands on today, which `|| 7` turns
  // into a full week instead, so a Monday given this list always means one
  // still to come.
  monday.setDate(monday.getDate() + ((8 - monday.getDay()) % 7 || 7));
  monday.setHours(8, 0, 0, 0);
  options.push({ id: 'monday', label: 'Monday morning, 8am', at: monday });

  return options;
}

/**
 * A person's own date and time, from a `<input type="datetime-local">` —
 * read as local time, the way the field itself shows it — or `null` when
 * what came back does not parse as one, which an empty field and a browser
 * that rejects the value both look like.
 */
export function parseCustomSchedule(value) {
  if (!value) return null;
  const at = new Date(value);
  return Number.isNaN(at.getTime()) ? null : at;
}
