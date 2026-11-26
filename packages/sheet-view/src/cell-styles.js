/**
 * Home → Cell Styles: the styles Excel ships, by the names and the built-in
 * ids Excel knows them by, so a file saved here opens there with the same
 * styles in its own gallery.
 *
 * The headings, the title and the total take their colours from the theme
 * (the dark text colour and the first accent), as Excel's do, so they change
 * with Page Layout → Themes and Colours. Good, Bad, Neutral and the data
 * styles are fixed colours in Excel too.
 */

const TEXT_2 = { theme: 3 };
const ACCENT_1 = { theme: 4 };
const THIN = { style: 'thin', colour: '#7F7F7F' };
const MEDIUM = { style: 'medium', colour: '#3F3F3F' };

export const CELL_STYLES = [
  { name: 'Normal', builtinId: 0, format: { bold: false, italic: false, fill: null, fontColour: null, border: { top: null, bottom: null, left: null, right: null }, numberFormat: 'General' } },
  { name: 'Heading 1', builtinId: 16, format: { bold: true, fontSize: 15, fontColour: TEXT_2, border: { bottom: { style: 'medium', colour: ACCENT_1 } } } },
  { name: 'Heading 2', builtinId: 17, format: { bold: true, fontSize: 13, fontColour: TEXT_2, border: { bottom: { style: 'medium', colour: { theme: 4, tint: 0.499984740745262 } } } } },
  { name: 'Heading 3', builtinId: 18, format: { bold: true, fontSize: 11, fontColour: TEXT_2, border: { bottom: { style: 'medium', colour: { theme: 4, tint: 0.3999755851924192 } } } } },
  { name: 'Title', builtinId: 15, format: { bold: true, fontSize: 18, fontColour: TEXT_2 } },
  { name: 'Total', builtinId: 25, format: { bold: true, border: { top: { style: 'thin', colour: ACCENT_1 }, bottom: { style: 'double', colour: ACCENT_1 } } } },
  { name: 'Good', builtinId: 26, format: { fill: '#C6EFCE', fontColour: '#006100' } },
  { name: 'Bad', builtinId: 27, format: { fill: '#FFC7CE', fontColour: '#9C0006' } },
  { name: 'Neutral', builtinId: 28, format: { fill: '#FFEB9C', fontColour: '#9C5700' } },
  { name: 'Input', builtinId: 20, format: { fill: '#FFCC99', fontColour: '#3F3F76', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } } },
  { name: 'Calculation', builtinId: 22, format: { bold: true, fill: '#F2F2F2', fontColour: '#FA7D00', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } } },
  { name: 'Check Cell', builtinId: 23, format: { bold: true, fill: '#A5A5A5', fontColour: '#FFFFFF', border: { top: MEDIUM, bottom: MEDIUM, left: MEDIUM, right: MEDIUM } } },
  { name: 'Note', builtinId: 10, format: { fill: '#FFFFCC', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } } },
  { name: 'Warning Text', builtinId: 11, format: { fontColour: '#FF0000' } },
  { name: 'Currency', builtinId: 4, format: { numberFormat: '"£"#,##0.00' } },
  { name: 'Percent', builtinId: 5, format: { numberFormat: '0%' } },
  { name: 'Comma', builtinId: 3, format: { numberFormat: '#,##0.00' } },
];

/** The built-in style of that name, whatever its capitals. */
export function builtinCellStyle(name) {
  const wanted = String(name ?? '').trim().toLowerCase();
  return CELL_STYLES.find((s) => s.name.toLowerCase() === wanted) ?? null;
}
