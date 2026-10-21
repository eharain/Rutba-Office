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
  const m = /^rutba-office:\/\/talk\/invite\/([A-Za-z0-9_-]{8,128})\/?$/.exec(String(text || '').trim());
  return m ? { app: 'talk', query: { invite: m[1] } } : null;
}
