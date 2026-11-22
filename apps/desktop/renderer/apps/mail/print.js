// The open message, on paper.
//
// What prints is what the reading pane shows: who it is from and to, when,
// what is attached, and the body. The same rules hold as on the screen:
//
// - a message shown as plain text prints as plain text;
// - remote pictures print only when the reader has let them be shown, and
//   the page's policy fetches nothing else either way (print-page.js);
// - the only pictures of Office's own that a message draws are the ones it
//   carries, held by the main process: an address of a file on this computer
//   written by the sender draws nothing, and is not read as one;
// - with Safety on, a link prints as its words: the PDF a print can be saved
//   as carries no address to click, as none can be copied from the screen.
//
// The body is the sender's HTML, so it is kept apart from the heading: it is
// written into a shadow tree of its own, where its styles reach nothing but
// itself. A message cannot restyle or hide the lines that say who sent it.
// The hidden window runs no script, whatever the message holds.

import { t, numberTag } from '@rutba/office-ui/messages';
import { pageHtml, esc, sizeWords } from '../../print-page.js';
import { BLOCKED_NOTE, stripTags } from './parts.js';

const person = (p) => (p?.name && p?.address ? `${p.name} <${p.address}>` : p?.name || p?.address || '');
const people = (list) => (list || []).map(person).filter(Boolean).join(', ');

/** When the message was sent, in full, in the window's language and digits. */
export function sentWords(date, tag = numberTag()) {
  if (!date) return '';
  const d = new Date(date);
  if (Number.isNaN(d.getTime())) return '';
  try {
    return d.toLocaleString(tag, { dateStyle: 'full', timeStyle: 'short' });
  } catch {
    return d.toLocaleString(undefined, { dateStyle: 'full', timeStyle: 'short' });
  }
}

/**
 * The body as it prints: the sender's HTML with its inline pictures drawn
 * from where the main process holds them, remote pictures left out unless
 * shown, and links as words unless `links`. A closing template tag in the
 * message is written as text, so nothing in it leaves the tree it is drawn in.
 */
export function messagePrintBody(message, { remote = false, plain = false, links = false } = {}) {
  let html = plain ? '' : String(message?.html || '');
  if (!html) {
    const text = message?.text || (plain ? stripTags(message?.html || '') : '');
    return `<pre class="plain">${esc(text)}</pre>`;
  }
  // An address of Office's own that the sender wrote is theirs, not ours: its slashes are written as
  // a reference, so the words read the same and the main process finds no file named in the page.
  html = html.replace(/rutba:\/\//gi, 'rutba:&#47;/');
  for (const [cid, url] of Object.entries(message.inlineUrls || {})) html = html.split(`cid:${cid}`).join(url);
  if (!remote) html = html.replace(/(<img\b[^<>]*?\bsrc=)(["'])(https?:[^"']*)\2/gi, `$1$2$2 data-${BLOCKED_NOTE}=$2$3$2`);
  if (!links) {
    const HREF = /(<(?:a|area)\b[^<>]*?)\s+href\s*=\s*(?:"[^"]*"|'[^']*'|[^\s<>]+)/gi;
    let was;
    do {
      was = html;
      html = html.replace(HREF, '$1');
    } while (html !== was);
  }
  return html.replace(/<\/(template)/gi, '&lt;/$1');
}

const BODY_CSS = `
  :host { display: block; font: 10.5pt/1.5 "Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif; color: #111; overflow-wrap: anywhere; }
  img { max-width: 100%; height: auto; }
  pre.plain { white-space: pre-wrap; font: inherit; margin: 0; }
  blockquote { margin: 0 0 0 12px; padding-left: 12px; border-left: 3px solid #d5d9e0; color: #5b626d; }
  a { color: #1a56c4; }
  table { max-width: 100%; }
`;

const CSS = `
  .head { margin: 0 0 4mm; padding-bottom: 3mm; border-bottom: 0.3mm solid #9aa0a6; break-inside: avoid; }
  .who th { text-align: start; vertical-align: top; font-weight: 600; color: #444; padding: 0.4mm 4mm 0.4mm 0; white-space: nowrap; }
  .who td { vertical-align: top; padding: 0.4mm 0; }
`;

/**
 * The page for the open message. `options` are the print dialog's (paper,
 * orientation, margins); `remote`, `plain` and `links` are the reading
 * pane's own state (links: Safety is off).
 */
export function messagePrintHtml(message, { options = {}, remote = false, plain = false, links = false, lang = 'en', dir = 'ltr', tag = numberTag() } = {}) {
  const subject = message?.subject || t('mail.no-subject');
  const attachments = (message?.attachments || []).filter((a) => !a.inline);
  const rows = [
    [t('common.from'), people(message?.from)],
    [t('common.to'), people(message?.to)],
    [t('common.cc'), people(message?.cc)],
    [t('common.bcc'), people(message?.bcc)],
    [t('common.date'), sentWords(message?.date, tag)],
    [t('common.attachments'), attachments.map((a) => `${a.filename || t('common.file')} (${sizeWords(a.size)})`).join(', ')],
  ].filter(([, value]) => value);
  const body = `<header class="head">
<h1>${esc(subject)}</h1>
<table class="who">${rows.map(([name, value]) => `<tr><th scope="row">${esc(name)}</th><td>${esc(value)}</td></tr>`).join('')}</table>
</header>
<div class="msg"><template shadowrootmode="open"><style>${BODY_CSS}</style>${messagePrintBody(message, { remote, plain, links })}</template></div>`;
  return pageHtml({ title: subject, options, css: CSS, body, remote: remote && !plain, files: false, lang, dir });
}
