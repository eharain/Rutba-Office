// The registry's words in the window's language.
//
// The app names, their taglines, the kinds of file and the New templates
// live in @rutba/office-formats/registry, which the main process and the
// installer read too — Windows is told the file types in English — so they
// stay English there. They are marked for the catalogue here, each written
// out as msg() needs it, and the window's own copy of the registry is put
// into its language once, as language.js settles it, before any app reads it.

import { APPS, WRITABLE, NEW_DOCUMENTS } from '@rutba/office-formats/registry';
import { msg, t } from '@rutba/office-ui';

/** Every word of the registry a window shows, marked for the catalogue. */
export const REGISTRY_WORDS = [
  msg('Rutba Mail'), msg('Mail'), msg('Every account, every archive, one inbox'),
  msg('Rutba Documents'), msg('Documents'), msg('document'), msg('Documents that open the same everywhere'),
  msg('Rutba Worksheets'), msg('Worksheets'), msg('workbook'), msg('Real formulas, real recalculation'),
  msg('Rutba Presentations'), msg('Presentations'), msg('presentation'), msg('Slides that survive the round trip'),
  msg('Rutba Pictures'), msg('Pictures'), msg('A viewer that opens before you blink'),
  msg('Rutba Image'), msg('Image'), msg('Crop, correct, annotate, export'),
  msg('Rutba Video'), msg('Video'), msg('Trim and export without a render farm'),
  msg('Rutba Calendar'), msg('Calendar'), msg('Your days, on this computer'),
  msg('Rutba Contacts'), msg('Contacts'), msg('Everyone you write to, in one book'),
  msg('Word Document (.docx)'), msg('OpenDocument Text (.odt)'), msg('PDF Document (.pdf)'), msg('Plain Text (.txt)'),
  msg('Markdown (.md)'), msg('Web Page (.html)'), msg('Worksheet (.xlsx)'), msg('OpenDocument Spreadsheet (.ods)'),
  msg('Comma Separated Values (.csv)'), msg('Tab Separated Values (.tsv)'), msg('Presentation (.pptx)'),
  msg('OpenDocument Presentation (.odp)'), msg('PNG Image (.png)'), msg('JPEG Image (.jpg)'), msg('WebP Image (.webp)'),
  msg('WebM Video (.webm)'), msg('MP4 Video (.mp4)'), msg('Email Message (.eml)'), msg('Mbox Archive (.mbox)'),
  msg('Calendar (.ics)'), msg('vCard (.vcf)'),
  msg('Pictures and PDFs'), msg('Images'), msg('Video and audio'), msg('Mail files'), msg('Calendars'), msg('Address books'),
  msg('All files'), msg('All supported files'),
  msg('Blank document'), msg('Letter'), msg('Report'), msg('Blank worksheet'), msg('Budget'), msg('Invoice'),
  msg('Blank presentation'), msg('Pitch deck'),
];

/** The window's copy of the registry, its words put into the window's language. */
export function registryInLanguage() {
  for (const app of Object.values(APPS)) {
    for (const key of ['name', 'short', 'noun', 'tagline']) if (app[key]) app[key] = t(app[key]);
  }
  for (const list of Object.values(WRITABLE)) for (const w of list) w.label = t(w.label);
  for (const d of NEW_DOCUMENTS) d.label = t(d.label);
}

/** A dialog's file filters, their names in the window's language. */
export const filtersInLanguage = (filters) => filters.map((f) => ({ ...f, name: t(f.name) }));
