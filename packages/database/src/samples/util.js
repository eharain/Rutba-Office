// What every sample database is filled from: a seeded random source, so a
// sample is the same each time it is made on the same day; dates counted
// back from the day it is made, so "due this week" and "not seen for six
// months" always have something to show; and invented people, companies
// and addresses. Every name is made up. E-mail addresses are on example
// domains, and phone numbers come only from the ranges kept for fiction:
// the UK's drama numbers (07700 900xxx, 020 7946 0xxx) and the US's 555-01xx.

/** A seeded random source: next() in [0, 1). */
export function random(seed = 1) {
  let a = seed >>> 0;
  const next = () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const int = (lo, hi) => lo + Math.floor(next() * (hi - lo + 1));
  const pick = (list) => list[Math.floor(next() * list.length)];
  const chance = (p) => next() < p;
  /** `k` different items of `list`. */
  const some = (list, k) => {
    const copy = [...list];
    const out = [];
    while (out.length < k && copy.length) out.push(copy.splice(Math.floor(next() * copy.length), 1)[0]);
    return out;
  };
  /** An item of `list` chosen by weight: [[item, weight], ...]. */
  const weighted = (pairs) => {
    const total = pairs.reduce((n, [, w]) => n + w, 0);
    let r = next() * total;
    for (const [item, w] of pairs) { r -= w; if (r < 0) return item; }
    return pairs[pairs.length - 1][0];
  };
  return { next, int, pick, chance, some, weighted };
}

const pad = (n) => String(n).padStart(2, '0');

/** Dates relative to `today` (a Date): day(-30) is thirty days before it, as YYYY-MM-DD. */
export function calendar(today = new Date()) {
  const base = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const at = (offset) => new Date(base.getFullYear(), base.getMonth(), base.getDate() + offset);
  const iso = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const day = (offset) => iso(at(offset));
  /** A working day near `offset`: a Saturday or Sunday moves to the Monday after. */
  const workday = (offset) => {
    const d = at(offset);
    const w = d.getDay();
    return w === 6 ? offset + 2 : w === 0 ? offset + 1 : offset;
  };
  const time = (offset, hour, minute = 0) => `${day(offset)}T${pad(hour)}:${pad(minute)}:00`;
  const weekday = (offset) => at(offset).getDay();
  return { day, workday, time, weekday, iso, at };
}

/** First names from many places, and family names, for invented people. */
export const FIRST = [
  'Amira', 'Ben', 'Chen', 'Daniela', 'Emeka', 'Fatima', 'Giorgio', 'Hana', 'Ivan', 'Jaya', 'Kofi', 'Leila', 'Mateo', 'Nadia',
  'Oskar', 'Priya', 'Rhys', 'Sofia', 'Tariq', 'Ursula', 'Viktor', 'Wen', 'Yusuf', 'Zara', 'Aisha', 'Bilal', 'Camille', 'Dmitri',
  'Elena', 'Farhan', 'Grace', 'Hamza', 'Imogen', 'Javier', 'Keiko', 'Lukas', 'Maryam', 'Noah', 'Olu', 'Paulo', 'Rania', 'Samir',
  'Tomasz', 'Uzma', 'Valentina', 'Wale', 'Ximena', 'Yara', 'Zain', 'Anika', 'Callum', 'Divya', 'Eamon', 'Freya', 'Hiroshi', 'Ines',
];
export const LAST = [
  'Haddad', 'Cole', 'Wang', 'Moreno', 'Okafor', 'Rahman', 'Bianchi', 'Sato', 'Petrov', 'Iyer', 'Mensah', 'Karimi', 'Silva', 'Hussain',
  'Lindqvist', 'Sharma', 'Morgan', 'Marchetti', 'Mahmood', 'Keller', 'Novak', 'Liu', 'Demir', 'Qureshi', 'Siddiqui', 'Dubois', 'Volkov',
  'Fraser', 'Hartley', 'Garcia', 'Nakamura', 'Fischer', 'Abbasi', 'Adeyemi', 'Costa', 'Nasser', 'Kowalski', 'Price', 'Davies', 'Osei',
  'Chaudhry', 'Byrne', 'Brennan', 'Tanaka', 'Rossi', 'Jensen', 'Patel', 'Ali', 'Khan', 'Ahmed', 'Lopez', 'Murphy', 'Walsh', 'Sutherland',
];

/** Cities by country, with how an address there is written: a postcode or ZIP pattern and a phone maker. */
export const PLACES = [
  { country: 'United Kingdom', cities: ['London', 'Manchester', 'Leeds', 'Bristol', 'Glasgow', 'Birmingham', 'Cardiff'], post: (r) => `${r.pick(['LS', 'M', 'BS', 'G', 'B', 'CF', 'N', 'SE', 'EC'])}${r.int(1, 19)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU'])}` },
  { country: 'United States', cities: ['Boston', 'Chicago', 'Denver', 'Seattle', 'Austin', 'Atlanta'], post: (r) => String(r.int(10000, 98999)) },
  { country: 'Canada', cities: ['Toronto', 'Vancouver', 'Montreal', 'Calgary'], post: (r) => `${r.pick(['M5V', 'V6B', 'H2X', 'T2P'])} ${r.int(1, 9)}${r.pick(['A', 'B', 'C'])}${r.int(1, 9)}` },
  { country: 'Ireland', cities: ['Dublin', 'Cork', 'Galway'], post: (r) => `${r.pick(['D02', 'T12', 'H91'])} ${r.pick(['X', 'Y', 'K'])}${r.int(100, 999)}` },
  { country: 'Pakistan', cities: ['Lahore', 'Karachi', 'Islamabad', 'Faisalabad'], post: (r) => String(r.int(44000, 75999)) },
  { country: 'United Arab Emirates', cities: ['Dubai', 'Abu Dhabi', 'Sharjah'], post: () => '' },
  { country: 'Germany', cities: ['Hamburg', 'Munich', 'Cologne', 'Berlin'], post: (r) => String(r.int(10115, 81929)) },
  { country: 'Netherlands', cities: ['Rotterdam', 'Utrecht', 'Amsterdam'], post: (r) => `${r.int(1011, 3584)} ${r.pick(['AB', 'CD', 'KL', 'ST'])}` },
  { country: 'Australia', cities: ['Sydney', 'Melbourne', 'Brisbane', 'Perth'], post: (r) => String(r.int(2000, 6999)) },
  { country: 'Nigeria', cities: ['Lagos', 'Abuja'], post: (r) => String(r.int(100001, 900999)) },
];
const STREETS = ['Harbour Street', 'Mill Lane', 'Station Road', 'Orchard Way', 'Church Street', 'Canal Walk', 'Market Square', 'Beech Grove', 'Riverside Drive', 'Quarry Road', 'Cedar Avenue', 'Linden Close'];

/** A phone number only from a range kept for fiction, or '' where the place has none we can use. */
export function phoneFor(r, country) {
  if (country === 'United Kingdom') return r.chance(0.5) ? `07700 900${String(r.int(0, 999)).padStart(3, '0')}` : `020 7946 0${String(r.int(0, 999)).padStart(3, '0')}`;
  if (country === 'United States' || country === 'Canada') return `(${r.pick(['617', '312', '303', '206', '416', '604'])}) 555-01${String(r.int(0, 99)).padStart(2, '0')}`;
  return '';
}

/** An invented address somewhere among `places` (all of them by default). */
export function address(r, places = PLACES) {
  const p = r.pick(places);
  return { Address: `${r.int(1, 240)} ${r.pick(STREETS)}`, City: r.pick(p.cities), Postcode: p.post(r), Country: p.country };
}

/** An invented person: names, and an e-mail address on an example domain. */
export function person(r, { domain = 'example.com' } = {}) {
  const first = r.pick(FIRST);
  const last = r.pick(LAST);
  return { 'First Name': first, 'Last Name': last, 'E-mail': `${first}.${last}@${domain}`.toLowerCase() };
}

/** Invented company names: two words that sound like a business and are not one we know of. */
const COMPANY_A = ['Harbourside', 'Meridian', 'Alder', 'Copperfield', 'Bluewater', 'Kestrel', 'Lantern', 'Saffron', 'Northgate', 'Willow', 'Granite', 'Silverline', 'Orchard', 'Juniper', 'Cobalt', 'Riverbend', 'Sandstone', 'Larkspur', 'Thistle', 'Falcon', 'Amberly', 'Crescent', 'Hollybrook', 'Marlow'];
const COMPANY_B = ['Foods', 'Provisions', 'Grocers', 'Trading', 'Merchants', 'Wholesale', 'Supplies', 'Pantry', 'Deli', 'Market', 'Stores', 'Kitchen', 'Imports', 'Hospitality', 'Cafe', 'Hotels'];

/** `n` different invented company names. */
export function companies(r, n, tails = COMPANY_B) {
  const out = new Set();
  let guard = 0;
  while (out.size < n && guard++ < n * 50) out.add(`${r.pick(COMPANY_A)} ${r.pick(tails)}`);
  return [...out];
}

/** Money rounded to cents, as a till rounds it. */
export const money = (n) => Math.round(n * 100) / 100;
