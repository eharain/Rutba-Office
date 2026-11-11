// Links that open Rutba Office from elsewhere: a web page, an email, a chat.
// Only the shapes Office acts on are taken, so a page cannot use a link to
// reach anything else in it.

/** The scheme (electron-builder.config.cjs registers it with the system). */
export const LINK_SCHEME = 'rutba-office';

/**
 * A link Office acts on, as the app to open and its query, or null:
 * rutba-office://talk/invite/<code> opens Talk on that invitation.
 */
export function parseLink(text) {
  const asked = String(text || '').trim();
  const m = /^rutba-office:\/\/talk\/invite\/([A-Za-z0-9_-]{8,128})\/?$/.exec(asked);
  if (m) return { app: 'talk', query: { invite: m[1] } };
  // A mailto: link, Office Mail being the mail app Windows opens them with
  // (Mail 2.25): Mail opens a message to write, filled from the link, and
  // sends nothing by itself. Nothing over a line or past a sensible length.
  if (/^mailto:[^\s]/i.test(asked) && asked.length <= 8000 && !/[\u0000-\u001f\u007f]/.test(asked)) return { app: 'mail', query: { mailto: asked } };
  return null;
}
