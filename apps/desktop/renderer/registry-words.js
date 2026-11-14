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
  msg('registry.rutba-mail'), msg('common.mail'), msg('registry.every-account-every'),
  msg('registry.rutba-documents'), msg('common.documents'), msg('common.document'), msg('registry.documents-open-same'),
  msg('registry.rutba-worksheets'), msg('common.worksheets'), msg('common.workbook-2'), msg('registry.real-formulas-real'),
  msg('common.rutba-presentations'), msg('common.presentations'), msg('common.presentation-2'), msg('registry.slides-survive-round'),
  msg('registry.rutba-pictures'), msg('common.pictures'), msg('registry.viewer-opens-before'),
  msg('registry.rutba-image'), msg('common.image'), msg('registry.crop-correct-annotate'),
  msg('registry.rutba-video'), msg('common.video'), msg('registry.trim-export-without'),
  msg('registry.rutba-recorder'), msg('common.recorder'), msg('registry.record-screen-camera'),
  msg('registry.rutba-talk'), msg('common.talk'), msg('registry.messages-calls-contacts'),
  msg('registry.rutba-data-files'), msg('common.data-files'), msg('common.databases'), msg('registry.tables-relationships'),
  msg('registry.rutba-calendar'), msg('common.calendar'), msg('registry.days-computer'),
  msg('registry.rutba-contacts'), msg('common.contacts'), msg('registry.everyone-write-one'),
  msg('common.word-document-docx'), msg('registry.open-document-text'), msg('common.pdf-document-pdf'), msg('registry.plain-text-txt'),
  msg('registry.markdown-md'), msg('registry.web-page-html'), msg('registry.worksheet-xlsx'), msg('registry.open-document'),
  msg('registry.comma-separated'), msg('registry.tab-separated-values'), msg('registry.presentation-pptx'),
  msg('registry.open-document-presentation'), msg('registry.png-image-png'), msg('registry.jpeg-image-jpg'), msg('registry.web-p-image'),
  msg('common.rutba-video-project'), msg('common.rutba-image-project'), msg('common.web-m-video'), msg('common.mp4-video-mp4'), msg('registry.email-message-eml'), msg('registry.mbox-archive-mbox'),
  msg('registry.calendar-ics'), msg('registry.v-card-vcf'),
  msg('common.pictures-and-pdfs'), msg('common.images'), msg('registry.video-and-audio'), msg('registry.mail-files'), msg('common.calendars'), msg('common.address-books'),
  msg('common.all-files'), msg('registry.all-supported-files'),
  msg('common.blank-document'), msg('common.letter'), msg('common.report'), msg('common.blank-worksheet'), msg('registry.budget'), msg('common.invoice'),
  msg('common.blank-presentation'), msg('common.pitch-deck'),
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
