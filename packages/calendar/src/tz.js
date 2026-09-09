// Time zones, without a table of our own.
//
// An iCalendar file names a zone by TZID and gives local wall-clock times in
// it. Turning those into instants needs the zone's rules, which every copy of
// this suite already carries: the runtime's Intl data knows every IANA zone
// and every rule change in it, so a VTIMEZONE block in the file is read past
// rather than trusted — Outlook's are wrong for years outside the one it
// wrote them for. Outlook also names zones its own way ("GMT Standard Time");
// the table below turns the common ones into IANA names.

/** Windows zone names, as Outlook writes them, to the IANA zones Intl knows. */
export const WINDOWS_ZONES = {
  'GMT Standard Time': 'Europe/London',
  'Greenwich Standard Time': 'Atlantic/Reykjavik',
  'W. Europe Standard Time': 'Europe/Berlin',
  'Central Europe Standard Time': 'Europe/Budapest',
  'Central European Standard Time': 'Europe/Warsaw',
  'Romance Standard Time': 'Europe/Paris',
  'E. Europe Standard Time': 'Europe/Chisinau',
  'FLE Standard Time': 'Europe/Kiev',
  'GTB Standard Time': 'Europe/Bucharest',
  'Russian Standard Time': 'Europe/Moscow',
  'Turkey Standard Time': 'Europe/Istanbul',
  'Israel Standard Time': 'Asia/Jerusalem',
  'Arab Standard Time': 'Asia/Riyadh',
  'Arabian Standard Time': 'Asia/Dubai',
  'Pakistan Standard Time': 'Asia/Karachi',
  'India Standard Time': 'Asia/Kolkata',
  'Bangladesh Standard Time': 'Asia/Dhaka',
  'SE Asia Standard Time': 'Asia/Bangkok',
  'Singapore Standard Time': 'Asia/Singapore',
  'China Standard Time': 'Asia/Shanghai',
  'Taipei Standard Time': 'Asia/Taipei',
  'Tokyo Standard Time': 'Asia/Tokyo',
  'Korea Standard Time': 'Asia/Seoul',
  'AUS Eastern Standard Time': 'Australia/Sydney',
  'AUS Central Standard Time': 'Australia/Darwin',
  'W. Australia Standard Time': 'Australia/Perth',
  'New Zealand Standard Time': 'Pacific/Auckland',
  'South Africa Standard Time': 'Africa/Johannesburg',
  'Egypt Standard Time': 'Africa/Cairo',
  'W. Central Africa Standard Time': 'Africa/Lagos',
  'E. Africa Standard Time': 'Africa/Nairobi',
  'Atlantic Standard Time': 'America/Halifax',
  'Eastern Standard Time': 'America/New_York',
  'US Eastern Standard Time': 'America/Indiana/Indianapolis',
  'Central Standard Time': 'America/Chicago',
  'Mountain Standard Time': 'America/Denver',
  'US Mountain Standard Time': 'America/Phoenix',
  'Pacific Standard Time': 'America/Los_Angeles',
  'Alaskan Standard Time': 'America/Anchorage',
  'Hawaiian Standard Time': 'Pacific/Honolulu',
  'Central America Standard Time': 'America/Guatemala',
  'SA Pacific Standard Time': 'America/Bogota',
  'SA Western Standard Time': 'America/La_Paz',
  'E. South America Standard Time': 'America/Sao_Paulo',
  'Argentina Standard Time': 'America/Argentina/Buenos_Aires',
  'Canada Central Standard Time': 'America/Regina',
  'Mexico Standard Time': 'America/Mexico_City',
  'UTC': 'UTC',
  'Coordinated Universal Time': 'UTC',
};

const formatters = new Map();

/** Whether Intl knows the zone. */
export function knownZone(tzid) {
  if (!tzid) return false;
  try {
    formatterFor(tzid);
    return true;
  } catch {
    return false;
  }
}

/** An IANA name for a TZID as a file wrote it, or null when nothing knows it. */
export function ianaZone(tzid) {
  if (!tzid) return null;
  const bare = String(tzid).replace(/^\/?(mozilla\.org|apple\.com|microsoft)?[^/]*\/?/i, (m) => (m.includes('/') && !/^[A-Z][a-z]+\/[A-Z]/.test(tzid) ? '' : m)).trim();
  for (const candidate of [tzid, bare, WINDOWS_ZONES[tzid], WINDOWS_ZONES[bare]]) {
    if (candidate && knownZone(candidate)) return candidate;
  }
  // "(UTC+01:00) Amsterdam, Berlin…" and the like: nothing to do but the offset, if there is one.
  return null;
}

function formatterFor(tzid) {
  let f = formatters.get(tzid);
  if (!f) {
    f = new Intl.DateTimeFormat('en-US', {
      timeZone: tzid,
      hourCycle: 'h23',
      year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', second: 'numeric',
    });
    formatters.set(tzid, f);
  }
  return f;
}

/** The zone's wall-clock parts at an instant. */
export function partsAt(tzid, epochMs) {
  const parts = {};
  for (const p of formatterFor(tzid).formatToParts(new Date(epochMs))) if (p.type !== 'literal') parts[p.type] = Number(p.value);
  return { y: parts.year, m: parts.month, d: parts.day, H: parts.hour === 24 ? 0 : parts.hour, M: parts.minute, S: parts.second };
}

/** Minutes east of UTC that the zone keeps at an instant. */
export function offsetMinutes(tzid, epochMs) {
  const p = partsAt(tzid, epochMs);
  return Math.round((Date.UTC(p.y, p.m - 1, p.d, p.H, p.M, p.S) - epochMs) / 60000);
}

/**
 * The instant of a wall-clock time in a zone. Twice through the offset,
 * because the offset depends on the answer across a clock change; a time
 * that never happened (the spring-forward gap) lands after the gap, as every
 * calendar does.
 */
export function localToUtc({ y, m, d, H = 0, M = 0, S = 0 }, tzid) {
  const naive = Date.UTC(y, m - 1, d, H, M, S);
  const zone = ianaZone(tzid);
  if (!zone) return new Date(y, m - 1, d, H, M, S).getTime(); // floating: this machine's clock
  const first = offsetMinutes(zone, naive);
  let at = naive - first * 60000;
  const second = offsetMinutes(zone, at);
  if (second !== first) at = naive - second * 60000;
  return at;
}

/** The wall-clock parts of an instant in a zone, or on this machine when the zone is unknown. */
export function utcToLocal(epochMs, tzid) {
  const zone = ianaZone(tzid);
  if (zone) return partsAt(zone, epochMs);
  const d = new Date(epochMs);
  return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate(), H: d.getHours(), M: d.getMinutes(), S: d.getSeconds() };
}

/** This machine's zone, as Intl names it. */
export function localZone() {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  } catch {
    return 'UTC';
  }
}
