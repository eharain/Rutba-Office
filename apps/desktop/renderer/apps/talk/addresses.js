// Addresses as the Invite by email dialog takes them: typed, pasted, or
// picked from Contacts. Plain code, with no words of its own, so the tests
// reach it without a window.

export const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
/** What the introducer keeps of a note, and so what Rutba's email carries. */
export const NOTE_CAP = 300;
/** People in one go: what the introducer takes from one person in a day. */
export const INVITE_BATCH = 20;

/**
 * Addresses typed or pasted, split at commas, semicolons, spaces and new
 * lines; "Name <address>" keeps the name. Each address once.
 */
export function parseAddresses(text) {
  const out = [];
  const seen = new Set();
  const add = (email, name = '') => {
    const e = String(email || '').trim().replace(/^mailto:/i, '');
    if (!EMAIL.test(e) || seen.has(e.toLowerCase())) return;
    seen.add(e.toLowerCase());
    out.push({ email: e, name: String(name || '').trim().replace(/^"|"$/g, '') });
  };
  let rest = String(text || '');
  rest = rest.replace(/([^,;<>\n]*)<([^<>\s]+)>/g, (m, name, email) => { add(email, name); return ' '; });
  for (const piece of rest.split(/[\s,;]+/)) add(piece);
  return out;
}
