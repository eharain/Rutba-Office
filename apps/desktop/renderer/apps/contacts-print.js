// Contacts, on paper: the chosen card in full, or the shown list.
//
// A card prints what the window shows of it: the name, the title and
// organisation, a room's line, its categories, then every address, number
// and note under the same names. The list prints the cards the window lists
// (a search's results, or a file's cards), under their letters, a line each:
// the name, where the person works, and how to reach them.

import { t, tn } from '@rutba/office-ui/messages';
import { pageHtml, esc } from '../print-page.js';

const typeNames = () => ({ work: t('contacts.work'), home: t('contacts.home'), other: t('contacts.other'), mobile: t('contacts.mobile'), fax: t('contacts.fax') });
const nameOf = (row, names, fallback) => row.label || (row.type && row.type !== 'other' ? names[row.type] || row.type : fallback);
const bookable = (c) => c?.kind === 'room' || c?.kind === 'resource';
const roomLine = (c) => [c.kind === 'room' ? t('common.room') : t('common.resource'), c.room?.capacity ? tn(c.room.capacity, 'common.count-seat') : null, c.room?.where || null].filter(Boolean).join(', ');

/** The letter a card sorts under, as the window's list has it. */
export const letterOf = (c) => {
  const ch = String(c?.display || '').trim()[0] || '#';
  return /[a-z]/i.test(ch) ? ch.toUpperCase() : '#';
};

/** A card's rows as the window names them: [name, value] pairs, in the window's order. */
export function cardRows(c) {
  const names = typeNames();
  const rows = [];
  for (const e of c?.emails || []) rows.push([nameOf(e, names, t('contacts.email')), e.value]);
  for (const p of c?.phones || []) rows.push([nameOf(p, names, t('contacts.phone')), p.value]);
  for (const a of c?.addresses || []) rows.push([nameOf(a, names, t('contacts.address')), [a.street, a.city, a.region, a.postcode, a.country].filter(Boolean).join('\n')]);
  if (c?.birthday) rows.push([t('contacts.birthday-2'), c.birthday]);
  for (const u of c?.urls || []) rows.push([u.label || t('contacts.web'), u.value]);
  if (c?.note) rows.push([t('contacts.notes'), c.note]);
  return rows.filter(([, value]) => value);
}

const CARD_CSS = `
  .sub { color: #444; margin: 0 0 1mm; }
  .tags { color: #444; font-size: 9.5pt; margin: 0 0 1mm; }
  .fields { margin-top: 4mm; padding-top: 3mm; border-top: 0.3mm solid #9aa0a6; width: 100%; }
  .fields th { text-align: start; vertical-align: top; font-weight: 600; color: #444; width: 34mm; padding: 1mm 4mm 1mm 0; }
  .fields td { vertical-align: top; padding: 1mm 0; white-space: pre-wrap; }
  .fields tr { break-inside: avoid; }
`;

function cardBody(c) {
  const sub = [c.title, c.org].filter(Boolean).join(' · ');
  const rows = cardRows(c);
  return `<h1>${esc(c.display || t('contacts.no-name'))}</h1>
${sub ? `<p class="sub">${esc(sub)}</p>` : ''}
${bookable(c) ? `<p class="sub">${esc(roomLine(c))}</p>` : ''}
${bookable(c) && c.room?.equipment?.length ? `<p class="tags">${esc(c.room.equipment.join(', '))}</p>` : ''}
${c.categories?.length ? `<p class="tags">${esc(c.categories.join(', '))}</p>` : ''}
${rows.length ? `<table class="fields">${rows.map(([name, value]) => `<tr><th scope="row">${esc(name)}</th><td>${esc(value)}</td></tr>`).join('')}</table>` : ''}`;
}

/** The page for one card. */
export function contactPrintHtml(contact, { options = {}, lang = 'en', dir = 'ltr' } = {}) {
  return pageHtml({ title: contact?.display || t('contacts.no-name'), options, css: CARD_CSS, body: cardBody(contact || {}), lang, dir });
}

const LIST_CSS = `
  .count { color: #444; margin: 0 0 4mm; }
  .list { width: 100%; table-layout: fixed; font-size: 9.5pt; }
  .list td { vertical-align: top; padding: 1.1mm 3mm 1.1mm 0; border-bottom: 0.2mm solid #d5d9e0; }
  .list tr { break-inside: avoid; }
  .list .letter td { font-size: 11pt; font-weight: 700; padding-top: 4mm; border-bottom: 0.3mm solid #9aa0a6; break-after: avoid; }
  .list .n { width: 30%; font-weight: 600; }
  .list .o { width: 22%; color: #444; }
  .list .e { width: 28%; }
  .list .p { width: 20%; }
  .list .v { white-space: pre-wrap; }
`;

/** The page for the shown list. `title` is the heading the window gives it (a file's name, or nothing). */
export function contactListPrintHtml(list, { title = '', options = {}, lang = 'en', dir = 'ltr' } = {}) {
  const names = typeNames();
  const cards = list || [];
  const lines = [];
  let last = null;
  for (const c of cards) {
    const letter = letterOf(c);
    if (letter !== last) {
      lines.push(`<tr class="letter"><td colspan="4">${esc(letter)}</td></tr>`);
      last = letter;
    }
    const sub = bookable(c) ? roomLine(c) : [c.title, c.org].filter(Boolean).join(' · ');
    const emails = (c.emails || []).map((e) => e.value).filter(Boolean).join('\n');
    const phones = (c.phones || []).filter((p) => p.value).map((p) => `${p.label || names[p.type] || p.type || ''} ${p.value}`.trim()).join('\n');
    lines.push(`<tr><td class="n">${esc(c.display || t('contacts.no-name'))}</td><td class="o">${esc(sub)}</td><td class="e v">${esc(emails)}</td><td class="p v">${esc(phones)}</td></tr>`);
  }
  const heading = title || t('common.contacts');
  const body = `<h1>${esc(heading)}</h1>
<p class="count">${esc(tn(cards.length, 'contacts.count-contact'))}</p>
<table class="list"><tbody>${lines.join('')}</tbody></table>`;
  return pageHtml({ title: heading, options, css: LIST_CSS, body, lang, dir });
}
