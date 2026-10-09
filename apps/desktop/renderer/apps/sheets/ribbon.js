// The Worksheets ribbon.
//
// Laid out the way Excel lays it out — Home, Insert, Draw, Page Layout,
// Formulas, Data, Review, View, Automate, Help — with every group Excel has,
// because a person arriving from Excel should not have to learn where
// anything is. What the engine can do is wired to it; what it cannot is
// drawn where Excel draws it, disabled, with a title that says why. A ribbon
// that hides its gaps is a ribbon nobody fixes.

import React from 'react';
import { Ribbon, Group, Rows, Button, Separator, Select, t, tn } from '@rutba/office-ui';
import { WARP_PRESETS, WARP_MORE, warpLabel } from '@rutba/drawing/warp';
import { catalogByCategory } from '@rutba/formula';
import { numberFormats } from './dialogs.js';
import { recentSources } from './queries.js';
import { MARGIN_PRESETS as PRINT_MARGINS } from '../../print.js';
import { wordArtMenu } from '../../wordart.js';
import { PEN_COLOURS, PEN_WIDTHS } from '../slides/ink-geometry.js';

/** The palette a toolbar offers before it offers a colour picker. */
const SWATCHES = [
  ['#000000', t('Black')], ['#444444', t('Dark grey')], ['#888888', t('Grey')], ['#ffffff', t('White')],
  ['#c00000', t('Dark red')], ['#e03131', t('Red')], ['#e08b2b', t('Orange')], ['#e0a800', t('Amber')],
  ['#0f9d58', t('Green')], ['#0d8f6f', t('Teal')], ['#2b5fd9', t('Blue')], ['#7b5cd6', t('Purple')],
];

const FILLS = [
  [null, t('No fill')], ['FFF3BF', t('Light amber')], ['D3F9D8', t('Light green')], ['D0EBFF', t('Light blue')],
  ['FFE3E3', t('Light red')], ['E9ECEF', t('Light grey')], ['FFD8A8', t('Light orange')], ['E5DBFF', t('Light purple')],
];

const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Consolas'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];
const AUTOSUM_NAMES = { SUM: t('Sum'), AVERAGE: t('Average'), COUNT: t('Count'), MAX: t('Max'), MIN: t('Min') };

/**
 * A border delta names edges. The engine keeps the edges it is not told about,
 * writes the ones it is, and deletes the ones set to null — so "all" is four
 * thin black edges and "none" is four nulls.
 */
const THIN = { style: 'thin', colour: '#000000' };
const MEDIUM = { style: 'medium', colour: '#000000' };
const BORDERS = [
  [t('All borders'), { top: THIN, bottom: THIN, left: THIN, right: THIN }],
  [t('Outside borders'), { top: THIN, bottom: THIN, left: THIN, right: THIN }],
  [t('Thick outside borders'), { top: MEDIUM, bottom: MEDIUM, left: MEDIUM, right: MEDIUM }],
  [t('Top border'), { top: THIN }],
  [t('Bottom border'), { bottom: THIN }],
  [t('Left border'), { left: THIN }],
  [t('Right border'), { right: THIN }],
  [t('Thick bottom border'), { bottom: MEDIUM }],
  [t('Double bottom border'), { bottom: { style: 'double', colour: '#000000' } }],
  [t('Top and bottom border'), { top: THIN, bottom: THIN }],
  [t('No border'), { top: null, bottom: null, left: null, right: null }],
];

/** What the engine can draw, by the names it draws them under. */
/** A PivotChart can be any of these; never a scatter, as in Excel. */
const PIVOT_CHARTS = [
  ['column', t('Column')], ['bar', t('Bar')], ['line', t('Line')], ['area', t('Area')], ['pie', t('Pie')], ['doughnut', t('Doughnut')],
];

/** The pivot the active cell is in, from the frame's list of pivots. */
export function pivotAround(model, sel) {
  const at = sel?.active;
  if (!at || !model?.pivots) return null;
  const cell = (t) => {
    const m = /^([A-Z]+)(\d+)$/.exec(t || '');
    if (!m) return null;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { row: Number(m[2]) - 1, col: col - 1 };
  };
  return model.pivots.find((p) => {
    if (p.sheet !== model.sheet || !p.ref) return false;
    const [a, b] = p.ref.split(':').map(cell);
    return a && b && at.row >= a.row && at.row <= b.row && at.col >= a.col && at.col <= b.col;
  }) || null;
}

const CHARTS = [
  ['column', t('Column')], ['bar', t('Bar')], ['line', t('Line')], ['area', t('Area')], ['pie', t('Pie')], ['doughnut', t('Doughnut')],
];
const SHAPES = [
  ['rect', t('Rectangle')], ['roundRect', t('Rounded rectangle')], ['ellipse', t('Ellipse')], ['line', t('Line')],
  ['triangle', t('Triangle')], ['diamond', t('Diamond')], ['rightArrow', t('Arrow right')], ['leftArrow', t('Arrow left')],
  ['upArrow', t('Arrow up')], ['downArrow', t('Arrow down')], ['pentagon', t('Pentagon')], ['hexagon', t('Hexagon')],
  ['star5', t('Star')], ['plus', t('Plus')], ['chevron', t('Chevron')], ['parallelogram', t('Parallelogram')], ['trapezoid', t('Trapezoid')],
];

/**
 * Excel's cell styles, as the formatting they apply. Each is one `setFormat`
 * delta, so "Good" is exactly the green Excel means and "Normal" clears it.
 */
/**
 * Excel's Orientation menu, as the textRotation each writes: 0 horizontal,
 * 1..90 anticlockwise, 91..180 clockwise by the value less 90, 255 stacked.
 */
const ORIENTATIONS = [
  [t('Horizontal text'), 0],
  [t('Angle anticlockwise'), 45],
  [t('Angle clockwise'), 135],
  [t('Vertical text'), 255],
  [t('Rotate text up'), 90],
  [t('Rotate text down'), 180],
];

/** The print dialog's three margin presets by their lower-case names, so the tab and the dialog agree. */
export const MARGIN_PRESETS = Object.fromEntries(Object.entries(PRINT_MARGINS).map(([name, m]) => [name.toLowerCase(), m]));

/** Which preset a file's margins are, or null when they are its own. */
export function marginsName(margins) {
  if (!margins) return null;
  for (const [name, m] of Object.entries(MARGIN_PRESETS)) {
    if (['top', 'right', 'bottom', 'left'].every((side) => Math.abs((Number(margins[side]) || 0) - m[side]) < 0.3)) return name;
  }
  return null;
}

/** How many page breaks a person has put in the sheet, for the Breaks tip. */
function breaksSaid(page) {
  const n = (page?.rowBreaks?.length || 0) + (page?.colBreaks?.length || 0);
  return n ? tn(n, '{count} page break put by hand', '{count} page breaks put by hand') : t('no page breaks put by hand');
}

export const CELL_STYLES = [
  [t('Normal'), { bold: false, italic: false, fill: null, fontColour: null, border: { top: null, bottom: null, left: null, right: null }, numberFormat: 'General' }],
  [t('Heading 1'), { bold: true, fontSize: 15, fontColour: '#1F3864', border: { bottom: { style: 'medium', colour: '#4472C4' } } }],
  [t('Heading 2'), { bold: true, fontSize: 13, fontColour: '#1F3864', border: { bottom: { style: 'medium', colour: '#A9C4E9' } } }],
  [t('Heading 3'), { bold: true, fontSize: 11, fontColour: '#1F3864', border: { bottom: { style: 'medium', colour: '#B4C6E7' } } }],
  [t('Title'), { bold: true, fontSize: 18, fontColour: '#1F3864' }],
  [t('Total'), { bold: true, border: { top: THIN, bottom: { style: 'double', colour: '#4472C4' } } }],
  [t('Good'), { fill: '#C6EFCE', fontColour: '#006100' }],
  [t('Bad'), { fill: '#FFC7CE', fontColour: '#9C0006' }],
  [t('Neutral'), { fill: '#FFEB9C', fontColour: '#9C5700' }],
  [t('Input'), { fill: '#FFCC99', fontColour: '#3F3F76', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }],
  [t('Calculation'), { bold: true, fill: '#F2F2F2', fontColour: '#FA7D00', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }],
  [t('Check Cell'), { bold: true, fill: '#A5A5A5', fontColour: '#FFFFFF', border: { top: MEDIUM, bottom: MEDIUM, left: MEDIUM, right: MEDIUM } }],
  [t('Note'), { fill: '#FFFFCC', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }],
  [t('Warning Text'), { fontColour: '#FF0000' }],
  [t('Currency'), { numberFormat: '"£"#,##0.00' }],
  [t('Percent'), { numberFormat: '0%' }],
  [t('Comma'), { numberFormat: '#,##0.00' }],
];

/**
 * The functions Excel lists by category. Picking one starts an edit in the
 * active cell with `=NAME(` typed, which is what Excel's Insert Function does
 * after its wizard: the arguments are the person's to type.
 */
export const FUNCTIONS = (() => {
  // From the engine's own catalogue, which a test holds in lock-step with
  // the functions that exist — so nothing offered here can evaluate to
  // #NAME?. "Recently used" is Excel's first category; a fixed short list
  // stands in until use is remembered.
  const out = { [t('Recently used')]: ['SUM', 'AVERAGE', 'IF', 'COUNT', 'MAX', 'MIN', 'VLOOKUP', 'ROUND'] };
  for (const { category, functions } of catalogByCategory()) out[t(category)] = functions.map((f) => f.name);
  return out;
})();

/**
 * One more or one fewer decimal on a number format: General goes to `0.0`,
 * `0.00` goes to `0.000` or `0.0`, and a format with no decimals gains one.
 */
export function withDecimals(code, delta) {
  const c = !code || code === 'General' ? '0' : code;
  const m = /0(\.0+)?/.exec(c);
  if (!m) return delta > 0 ? c + '.0' : c;
  const current = m[1] ? m[1].length - 1 : 0;
  const next = Math.max(0, Math.min(10, current + delta));
  const decimals = next ? '.' + '0'.repeat(next) : '';
  return c.slice(0, m.index) + '0' + decimals + c.slice(m.index + m[0].length);
}

/** A control that is drawn where Excel draws it, and says why it is not live. */
const Soon = ({ icon, label, tall, why }) => (
  <Button tall={tall} icon={icon} label={label} disabled title={t('{label} — not built yet. {why}', { label, why })} />
);

export default function SheetsRibbon({
  tab, setTab, model, dispatch, commands, shell, menu, save, openFile, exportAs, doc, sel, openDialog, act, view = {}, review = null, arrange = { picked: [], pane: false }, ink = null,
}) {
  // Page Layout → Arrange acts on the drawings picked on the sheet.
  const nPicked = arrange.picked.length;
  const need = nPicked ? null : t('select a picture, shape, chart or slicer first');
  // A gallery hangs under the button that opened it.
  const openGallery = (e, kind) => {
    const r = e.currentTarget.getBoundingClientRect();
    act('themeGallery', { kind, anchor: { left: r.left, bottom: r.bottom + 4 } });
  };
  const turnable = arrange.picked.some((o) => ['shape', 'image', 'group'].includes(o.kind));
  const groupPicked = arrange.picked.some((o) => o.kind === 'group');
  const format = model?.format || {};
  const inPivot = pivotAround(model, sel);
  const frozen = model?.frozen || { rows: 0, cols: 0 };
  const isFrozen = frozen.rows > 0 || frozen.cols > 0;
  const protectedSheet = Boolean(model?.protection?.sheet);
  const protectedBook = Boolean(model?.workbookProtection?.structure);
  const size = Number(format.fontSize || 11);
  const nearer = (dir) => {
    const bigger = SIZES.filter((s) => (dir > 0 ? s > size : s < size));
    return dir > 0 ? bigger[0] ?? size : bigger[bigger.length - 1] ?? size;
  };

  const setFormat = (delta) => dispatch({ op: 'setFormat', delta });
  // Format as Table: Excel's own gallery names, so Excel shows the style it
  // asked for; this window paints every table in its own palette.
  const tableMenu = (event) => menu.open(event, [
    { label: t('Light, banded rows'), icon: 'table', run: () => act('table', { style: 'TableStyleLight9', stripes: true }) },
    { label: t('Medium, banded rows'), icon: 'table', run: () => act('table', { style: 'TableStyleMedium2', stripes: true }) },
    { label: t('Medium, plain rows'), icon: 'table', run: () => act('table', { style: 'TableStyleMedium2', stripes: false }) },
    { label: t('Dark, banded rows'), icon: 'table', run: () => act('table', { style: 'TableStyleDark1', stripes: true }) },
  ]);
  const swatchMenu = (event, key, list) =>
    menu.open(
      event,
      list.map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => setFormat({ [key]: value }) }))
    );
  const functionMenu = (event, names) =>
    menu.open(event, names.map((fn) => ({ label: fn, icon: 'formula', run: () => act('insertFunction', fn) })));

  return (
    <Ribbon
      tabs={[
        { id: 'home', label: t('Home') },
        { id: 'insert', label: t('Insert') },
        { id: 'draw', label: t('Draw') },
        { id: 'layout', label: t('Page Layout') },
        { id: 'formulas', label: t('Formulas') },
        { id: 'data', label: t('Data') },
        { id: 'review', label: t('Review') },
        { id: 'view', label: t('View') },
        { id: 'automate', label: t('Automate') },
        { id: 'help', label: t('Help') },
        // Shape Format, as Excel's: there while one shape with words is picked.
        ...(arrange.picked.length === 1 && arrange.picked[0].kind === 'shape' && arrange.picked[0].hasText ? [{ id: 'shapeFormat', label: t('Shape Format') }] : []),
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title={t('Save (Ctrl+S)')} onClick={() => save(false)} />
          <Button icon="undo" title={t('Undo (Ctrl+Z)')} disabled={!doc?.canUndo} onClick={() => commands['edit.undo'].run()} />
          <Button icon="redo" title={t('Redo (Ctrl+Y)')} disabled={!doc?.canRedo} onClick={() => commands['edit.redo'].run()} />
          <Button icon="sum" title={t('AutoSum')} onClick={() => commands['sheet.autoSum'].run()} />
          <Button icon="print" title={t('Print')} onClick={() => exportAs('pdf')} />
        </>
      }
    >
      {/* ── Shape Format (contextual): WordArt's Transform ──────────────── */}
      {tab === 'shapeFormat' && arrange.picked.length === 1 && arrange.picked[0].hasText ? (
        <>
          <Group label={t('Text Effects: Transform')}>
            {WARP_PRESETS.map((p) => (
              <Button key={p.id} tall icon="wand" label={t(p.label)} className="sh-warp" data-preset={p.id} pressed={(arrange.picked[0].textWarp || 'textNoShape') === p.id} title={p.id === 'textNoShape' ? t('No Transform — the words in straight lines') : p.label === 'Button' ? t('Transform — the words along a button: an arc, a line and an arc') : p.label === 'Circle' ? t('Transform — the words along the circle of the shape') : t('Transform — the words along the arc of the shape')} onClick={() => act('textWarp', { id: arrange.picked[0].id, preset: p.id })} />
            ))}
            {/* The warps, the rest of Office's gallery, from a menu. */}
            <Button tall icon="wand" label={t('More')} className="sh-warp-more" pressed={WARP_MORE.some((p) => p.id === arrange.picked[0].textWarp)} title={WARP_MORE.some((p) => p.id === arrange.picked[0].textWarp) ? t('Transform — now {warp}; the warps: the words stretched between two curves', { warp: t(warpLabel(arrange.picked[0].textWarp)) }) : t('More Transforms — the warps: the words stretched between two curves, a wave, a slant, a chevron and the rest')} onClick={(e) => menu.open(e, WARP_MORE.map((p) => ({ label: t(p.label), icon: arrange.picked[0].textWarp === p.id ? 'check' : 'wand', run: () => act('textWarp', { id: arrange.picked[0].id, preset: p.id }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Home ─────────────────────────────────────────────────────────── */}
      {tab === 'home' ? (
        <>
          <Group label={t('Clipboard')}>
            <Button tall icon="paste" label={t('Paste')} title={t('Paste (Ctrl+V)')} onClick={() => commands['edit.paste']?.run?.()} />
            <Button icon="chevronDown" className="sh-paste-options" title={t('Paste options — the formulas, the values or the formatting alone, turned rows to columns, or Paste Special (Ctrl+Alt+V)')} onClick={(e) => menu.open(e, [
              { label: t('Paste'), icon: 'paste', run: () => commands['edit.paste']?.run?.() },
              { label: t('Formulas'), run: () => act('pasteSpecial', { what: 'formulas' }) },
              { label: t('Values'), run: () => act('pasteSpecial', { what: 'values' }) },
              { label: t('Formatting'), run: () => act('pasteSpecial', { what: 'formats' }) },
              { label: t('Transpose'), run: () => act('pasteSpecial', { what: 'all', transpose: true }) },
              { label: t('Paste Special…'), run: () => commands['edit.pasteSpecial']?.run?.() },
            ])} />
            <Rows>
              <>
                <Button icon="cut" label={t('Cut')} onClick={() => commands['edit.cut']?.run?.()} />
                <Button icon="copy" label={t('Copy')} onClick={() => commands['edit.copy'].run()} />
              </>
              <Button icon="wand" label={t('Format Painter')} title={t('Copy the formatting here to the next selection')} onClick={() => dispatch({ op: 'formatBrush' })} />
            </Rows>
          </Group>

          <Group label={t('Font')}>
            <Rows>
              <>
                <Select value={format.fontName || 'Calibri'} onChange={(e) => setFormat({ fontName: e.target.value })} style={{ width: 118 }} title={t('Font')}>
                  {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
                  {format.fontName && !FONTS.includes(format.fontName) ? <option value={format.fontName}>{format.fontName}</option> : null}
                </Select>
                <Select value={String(size)} onChange={(e) => setFormat({ fontSize: Number(e.target.value) })} style={{ width: 56 }} title={t('Font size')}>
                  {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
                  {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
                </Select>
                <Button icon="chevronUp" title={t('Increase font size')} onClick={() => setFormat({ fontSize: nearer(1) })} />
                <Button icon="chevronDown" title={t('Decrease font size')} onClick={() => setFormat({ fontSize: nearer(-1) })} />
              </>
              <>
                <Button icon="bold" title={t('Bold (Ctrl+B)')} pressed={format.bold} onClick={() => setFormat({ bold: 'toggle' })} />
                <Button icon="italic" title={t('Italic (Ctrl+I)')} pressed={format.italic} onClick={() => setFormat({ italic: 'toggle' })} />
                <Button icon="underline" title={t('Underline (Ctrl+U)')} pressed={format.underline} onClick={() => setFormat({ underline: 'toggle' })} />
                <Button icon="strike" title={t('Strikethrough')} pressed={format.strike} onClick={() => setFormat({ strike: 'toggle' })} />
                <Separator />
                <Button icon="grid" title={t('Borders')} onClick={(e) => menu.open(e, BORDERS.map(([label, edges]) => ({ label, run: () => setFormat({ border: edges }) })))} />
                <Button icon="wand" title={t('Fill colour')} onClick={(e) => swatchMenu(e, 'fill', FILLS.map(([c, l]) => [c ? `#${c}` : null, l]))} />
                <Button icon="contrast" title={t('Font colour')} onClick={(e) => swatchMenu(e, 'fontColour', SWATCHES)} />
              </>
            </Rows>
          </Group>

          <Group label={t('Alignment')}>
            <Rows>
              <>
                <Button icon="chevronUp" title={t('Top align')} pressed={format.valign === 'top'} onClick={() => setFormat({ valign: 'top' })} />
                <Button icon="minus" title={t('Middle align')} pressed={format.valign === 'center'} onClick={() => setFormat({ valign: 'center' })} />
                <Button icon="chevronDown" title={t('Bottom align')} pressed={format.valign === 'bottom'} onClick={() => setFormat({ valign: 'bottom' })} />
                <Button icon="rotate" title={t('Text orientation')} pressed={Boolean(format.rotation)} onClick={(e) => menu.open(e, ORIENTATIONS.map(([label, rotation]) => ({ label, icon: (format.rotation ?? -1) === rotation ? 'check' : undefined, run: () => act('orientation', rotation) })))} />
                <Separator />
                <Button icon="listBullet" label={t('Wrap Text')} pressed={format.wrap} onClick={() => setFormat({ wrap: !format.wrap })} />
              </>
              <>
                <Button icon="alignLeft" title={t('Align left')} pressed={format.align === 'left'} onClick={() => setFormat({ align: 'left' })} />
                <Button icon="alignCenter" title={t('Centre')} pressed={format.align === 'center'} onClick={() => setFormat({ align: 'center' })} />
                <Button icon="alignRight" title={t('Align right')} pressed={format.align === 'right'} onClick={() => setFormat({ align: 'right' })} />
                <Button icon="chevronLeft" title={t('Decrease indent')} disabled={format.indent === 0} onClick={() => setFormat({ indentBy: -1 })} />
                <Button icon="chevronRight" title={t('Increase indent')} onClick={() => setFormat({ indentBy: 1 })} />
                <Separator />
                <Button icon="table" label={t('Merge & Centre')} title={t('Merge the selected cells into one, and centre it')} onClick={() => act('mergeCentre')} />
                <Button icon="table" title={t('Merge cells')} onClick={() => dispatch({ op: 'merge' })} />
                <Button icon="minus" title={t('Unmerge cells')} onClick={() => dispatch({ op: 'unmerge' })} />
              </>
            </Rows>
          </Group>

          <Group label={t('Number')}>
            <Rows>
              <Select value={format.numberFormat || 'General'} onChange={(e) => setFormat({ numberFormat: e.target.value })} style={{ width: 148 }} title={t('Number format')}>
                {numberFormats().map((f) => <option key={f.label} value={f.code}>{f.label}</option>)}
                {numberFormats().some((f) => f.code === (format.numberFormat || 'General')) ? null : (
                  <option value={format.numberFormat}>{format.numberFormat}</option>
                )}
              </Select>
              <>
                {/* The mark is the button. An icon beside a currency symbol says nothing. */}
                <Button title={t('Currency')} onClick={() => setFormat({ numberFormat: '"£"#,##0.00' })} label="£" />
                <Button title={t('Percent style')} onClick={() => setFormat({ numberFormat: '0%' })} label="%" />
                <Button title={t('Comma style')} onClick={() => setFormat({ numberFormat: '#,##0.00' })} label="," />
                <Button title={t('Increase decimal')} onClick={() => setFormat({ numberFormat: withDecimals(format.numberFormat, 1) })} label=".0" />
                <Button title={t('Decrease decimal')} onClick={() => setFormat({ numberFormat: withDecimals(format.numberFormat, -1) })} label=".00" />
              </>
            </Rows>
          </Group>

          <Group label={t('Styles')}>
            <Button tall icon="wand" label={t('Conditional Formatting')} onClick={() => openDialog('conditional')} />
            <Button tall icon="table" label={t('Format as Table')} title={t('Format as Table — a header row, banded rows and a style Excel knows by name, over the selection or the block of data round the cell')} onClick={tableMenu} />
            <Button tall icon="grid" label={t('Cell Styles')} title={t("Excel's cell styles: headings, totals, good, bad, input")} onClick={(e) => menu.open(e, CELL_STYLES.map(([label, delta]) => ({ label, run: () => setFormat(delta) })))} />
          </Group>

          <Group label={t('Cells')}>
            <Button tall icon="plus" label={t('Insert')} onClick={(e) => menu.open(e, [
              { label: t('Insert sheet rows'), icon: 'plus', run: () => commands['sheet.insertRow'].run() },
              { label: t('Insert sheet columns'), icon: 'plus', run: () => commands['sheet.insertCol'].run() },
            ])} />
            <Button tall icon="minus" label={t('Delete')} onClick={(e) => menu.open(e, [
              { label: t('Delete sheet rows'), icon: 'minus', run: () => commands['sheet.deleteRow'].run() },
              { label: t('Delete sheet columns'), icon: 'minus', run: () => commands['sheet.deleteCol'].run() },
            ])} />
            <Button tall icon="settings" label={t('Format')} onClick={(e) => menu.open(e, [
              { label: t('Row height…'), run: () => openDialog('rowHeight') },
              { label: t('Column width…'), run: () => openDialog('colWidth') },
              { label: t('AutoFit column width'), run: () => act('autoFit') },
              { label: '-' },
              { label: protectedSheet ? t('Unprotect sheet') : t('Protect sheet'), icon: 'lock', run: () => dispatch({ op: protectedSheet ? 'unprotect' : 'protect' }) },
              { label: format.locked === false ? t('Lock cells') : t('Unlock cells'), icon: 'lock', run: () => setFormat({ locked: format.locked === false }) },
            ])} />
          </Group>

          <Group label={t('Editing')}>
            <Button icon="sum" label={t('AutoSum')} onClick={(e) =>
              menu.open(e, ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((fn) => ({ label: AUTOSUM_NAMES[fn], icon: 'sum', run: () => dispatch({ op: 'autoSum', fn }) })))
            } />
            <Button icon="chevronDown" label={t('Fill')} onClick={(e) => menu.open(e, [
              { label: t('Down (Ctrl+D)'), run: () => act('fill', 'down') },
              { label: t('Right (Ctrl+R)'), run: () => act('fill', 'right') },
              { label: t('Up'), run: () => act('fill', 'up') },
              { label: t('Left'), run: () => act('fill', 'left') },
              { label: t('Series…'), run: () => act('seriesDialog') },
            ])} />
            <Button icon="close" label={t('Clear')} onClick={(e) => menu.open(e, [
              { label: t('Clear contents'), run: () => dispatch({ op: 'clear' }) },
              { label: t('Clear formats'), run: () => setFormat(CELL_STYLES[0][1]) },
            ])} />
            <Button icon="sort" label={t('Sort & Filter')} onClick={(e) => menu.open(e, [
              { label: t('Sort A to Z'), icon: 'sort', run: () => commands['sheet.sortAsc'].run() },
              { label: t('Sort Z to A'), icon: 'sort', run: () => commands['sheet.sortDesc'].run() },
              { label: model?.filtered ? t('Remove filter') : t('Filter'), icon: 'filter', run: () => dispatch({ op: 'autoFilter' }) },
            ])} />
            <Button icon="find" label={t('Find & Select')} onClick={(e) => menu.open(e, [
              { label: t('Find…'), icon: 'find', run: () => openDialog('find') },
              { label: t('Replace…'), icon: 'find', run: () => openDialog('find') },
              { label: t('Go To…'), run: () => openDialog('goto') },
            ])} />
          </Group>
        </>
      ) : null}

      {/* ── Insert ───────────────────────────────────────────────────────── */}
      {tab === 'insert' ? (
        <>
          <Group label={t('Tables')}>
            <Button tall icon="table" label={t('PivotTable')} title={t('PivotTable — summarise the list round the cell: rows, columns and values')} onClick={() => act('pivotTable')} />
            {inPivot && !inPivot.unsupported ? <Button icon="list" label={t('Field List')} title={t('Field List — the fields of {name}: what is on its rows, columns and values', { name: inPivot.name })} onClick={() => act('pivotFields')} /> : null}
            <Button tall icon="table" label={t('Table')} title={t('Table — the selection or the block of data round the cell, with a header row, banded rows and filters')} onClick={tableMenu} />
          </Group>
          <Group label={t('Illustrations')}>
            <Button tall icon="picture" label={t('Pictures')} title={t('Pictures — a picture from a file, at the cell, at its own proportions')} onClick={() => act('picture')} />
            <Button tall icon="shape" label={t('Shapes')} onClick={(e) => menu.open(e, SHAPES.map(([geometry, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', geometry, text: '' }) })))} />
            <Button icon="star" label={t('Icons')} title={t("Icons — one of the suite's own icons, in the colour you choose, as a picture at the cell")} onClick={() => act('icons')} />
            <Button icon="shape" label={t('SmartArt')} title={t('SmartArt — a list, a process, a cycle or a hierarchy, drawn from lines you type, as a group of shapes at the cell')} onClick={() => act('smartArt')} />
          </Group>
          <Group label={t('Charts')}>
            {CHARTS.map(([kind, label]) => (
              <Button key={kind} icon="chart" label={label} title={t('{chart} chart from the data around the selection', { chart: label })} onClick={() => dispatch({ op: 'insertChart', kind })} />
            ))}
            <Button icon="chart" label={t('Scatter')} title={t('Scatter — X and Y values from the data around the selection: markers only, or joined by straight or smooth lines')} onClick={(e) => menu.open(e, [
              ['markers', t('Scatter')], ['lines', t('Scatter with Straight Lines and Markers')], ['smooth', t('Scatter with Smooth Lines and Markers')],
            ].map(([scatterStyle, label]) => ({ label, icon: 'chart', run: () => dispatch({ op: 'insertChart', kind: 'scatter', scatterStyle }) })))} />
            <Button icon="chart" label={t('PivotChart')} title={inPivot ? t('PivotChart — a chart of {name}, following it when it is refreshed or filtered', { name: inPivot.name }) : t('PivotChart — PivotChart & PivotTable from the list round the cell')}
              onClick={(e) => (inPivot
                ? menu.open(e, PIVOT_CHARTS.map(([kind, label]) => ({ label: t('{chart} PivotChart', { chart: label }), icon: 'chart', run: () => act('pivotChart', { kind }) })))
                : act('pivotChart'))} />
          </Group>
          <Group label={t('Sparklines')}>
            <Button tall icon="chart" label={t('Line')} title={t('Line sparkline — a small line, in the cell after the selection, from the numbers in it')} onClick={() => openDialog('sparklineLine')} />
            <Button icon="chart" label={t('Column')} title={t('Column sparkline — small bars, in the cell after the selection, from the numbers in it')} onClick={() => openDialog('sparklineColumn')} />
          </Group>
          <Group label={t('Filters')}>
            <Button tall icon="filter" label={t('Filter')} pressed={model?.filtered} onClick={() => dispatch({ op: 'autoFilter' })} />
            <Button icon="filter" label={t('Slicer')} title={t('Slicer — buttons that filter the table or pivot table the cell is in, one panel per field')} onClick={() => act('slicer')} />
          </Group>
          <Group label={t('Links')}>
            <Button tall icon="link" label={t('Link')} title={t('A link on this cell: an address, or a place in the workbook (Ctrl+K)')} onClick={() => act('link')} />
          </Group>
          <Group label={t('Comments')}>
            <Button tall icon="reply" label={t('Comment')} title={t('Comment — a conversation on this cell: replies, resolve, reopen (Ctrl+Alt+M)')} onClick={() => act('newComment')} />
            <Button tall icon="reply" label={t('Note')} title={t('A note on this cell, shown when the pointer rests on it (Shift+F2)')} onClick={() => act('note')} />
            <Button icon="close" label={t('Delete')} title={t('Take the note off this cell')} onClick={() => act('removeNote')} />
          </Group>
          <Group label={t('Text')}>
            <Button tall icon="textbox" label={t('Text Box')} title={t('A rectangle with words in it')} onClick={() => act('textBox')} />
            <Button icon="file" label={t('Header & Footer')} title={t('Header & Footer — what prints at the top and the foot of every page')} onClick={() => act('headerFooter')} />
            <Button icon="wand" label={t('WordArt')} title={t('WordArt — big words in a style of their own: a fill, an outline round the letters, a shadow or a glow')} onClick={(e) => menu.open(e, wordArtMenu((style) => act('wordArt', style)))} />
          </Group>
          <Group label={t('Symbols')}>
            <Button icon="formula" label={t('Equation')} title={t('Equation — typed in its linear form (x^2+y^2=r^2), set as math over the selection; double-click one to change it')} onClick={() => act('equation')} />
            <Button icon="plus" label={t('Symbol')} onClick={() => openDialog('symbol')} />
          </Group>
          <Group label={t('Names')}>
            <Button tall icon="find" label={t('Define Name')} onClick={() => openDialog('names')} />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label={t('Drawing Tools')}>
            <Button tall icon="mouse" label={t('Select')} pressed={!ink?.tool} title={t('Select — put the pen down and work with the cells again (Esc)')} onClick={() => act('inkTool', null)} />
            <Button tall icon="wand" label={t('Lasso')} pressed={ink?.tool === 'lasso'} title={t('Lasso Select — draw a loop round strokes to select them')} onClick={() => act('inkTool', 'lasso')} />
            <Button tall icon="close" label={t('Eraser')} pressed={ink?.tool === 'eraser'} title={t('Eraser — take away each stroke the pointer passes over')} onClick={() => act('inkTool', 'eraser')} />
            {(ink?.pens || []).map((p) => {
              const on = ink.tool === 'pen' && ink.penId === p.id;
              const label = { pen: t('Pen'), pencil: t('Pencil'), highlighter: t('Highlighter') }[p.tool];
              return (
                <Button key={p.id} tall icon="wand" label={label} pressed={on} className={`sl-pen sl-pen-${p.tool}`} data-pen={p.id} style={{ '--pen': p.color }}
                  title={on ? t('{pen} — click again for its colour and thickness', { pen: label }) : t('{pen} — draw over the cells', { pen: label })}
                  onClick={(e) => {
                    if (!on) { act('inkTool', p.id); return; }
                    menu.open(e, [
                      { heading: true, label: t('Thickness') },
                      ...PEN_WIDTHS[p.tool].map((w) => ({ label: t('{width} pt', { width: w }), icon: p.width === w ? 'check' : undefined, run: () => act('inkPen', { id: p.id, width: w }) })),
                      { heading: true, label: t('Colour') },
                      ...PEN_COLOURS.map((c) => ({ label: c, icon: p.color === c ? 'check' : undefined, preview: <span style={{ display: 'inline-block', width: 16, height: 16, borderRadius: 8, background: c, border: '1px solid rgba(0,0,0,.25)' }} />, run: () => act('inkPen', { id: p.id, color: c }) })),
                    ]);
                  }} />
              );
            })}
            <Button tall icon="plus" label={t('Add')} title={t('Add Pen — another pen or highlighter in the gallery')} onClick={(e) => menu.open(e, [['pen', t('Pen')], ['highlighter', t('Highlighter')]].map(([tool, label]) => ({ label, icon: 'plus', run: () => act('inkAdd', tool) })))} />
          </Group>
          <Group label={t('Convert')}>
            <Button tall icon="shape" label={t('Ink to Shape')} pressed={Boolean(ink?.toShape)} title={t('Ink to Shape — a rectangle, oval or triangle drawn becomes that shape')} onClick={() => act('inkToShape')} />
            <Soon tall icon="formula" label={t('Ink to Math')} why={t('Turning handwriting into an equation needs handwriting recognition, which this suite does not have.')} />
          </Group>
          <Group label={t('Shapes')}>
            <Button tall icon="shape" label={t('Shapes')} title={t("Shapes — a rectangle, an oval, an arrow, in the theme's colours")} onClick={(e) => menu.open(e, SHAPES.map(([geometry, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', geometry, text: '' }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Page Layout ──────────────────────────────────────────────────── */}
      {tab === 'layout' ? (
        <>
          <Group label={t('Themes')}>
            <Button tall icon="wand" label={t('Themes')} pressed={view.gallery === 'themes'} title={t("Themes — the suite's themes, the same as Presentation's; this workbook wears {theme}", { theme: model?.design?.name || 'Office Theme' })} onClick={(e) => openGallery(e, 'themes')} />
            <Button icon="contrast" label={t('Colours')} title={t("Colours — the theme's twelve colours, or your own; now {name}", { name: model?.design?.colorName || 'Office' })} onClick={(e) => openGallery(e, 'colours')} />
            <Button icon="textbox" label={t('Fonts')} title={t('Fonts — the heading and body faces; now {major} and {minor}', model?.design?.fonts ? { major: model.design.fonts.major, minor: model.design.fonts.minor } : { major: 'Calibri Light', minor: 'Calibri' })} onClick={(e) => openGallery(e, 'fonts')} />
            <Button icon="wand" label={t('Effects')} title={t('Effects — how shapes styled from the theme are filled, outlined and lifted; now {name}', { name: model?.design?.effectName || 'Office' })} onClick={(e) => openGallery(e, 'effects')} />
          </Group>
          <Group label={t('Page Setup')}>
            <Button tall icon="file" label={t('Margins')} title={{ normal: t('Margins — now normal'), narrow: t('Margins — now narrow'), wide: t('Margins — now wide') }[marginsName(view.page?.margins)] || (view.page?.margins ? t('Margins — now {top} mm top and bottom, {left} mm at the sides', { top: view.page.margins.top, left: view.page.margins.left }) : t('Margins — now normal'))} onClick={(e) => menu.open(e, [['normal', t('Normal')], ['narrow', t('Narrow')], ['wide', t('Wide')]].map(([m, label]) => ({ label, icon: marginsName(view.page?.margins) === m ? 'check' : undefined, run: () => act('page', { margins: m }) })))} />
            <Button tall icon="rotate" label={t('Orientation')} title={view.page?.orientation === 'landscape' ? t('Orientation — now landscape') : t('Orientation — now portrait')} onClick={(e) => menu.open(e, [['portrait', t('Portrait')], ['landscape', t('Landscape')]].map(([o, label]) => ({ label, icon: (view.page?.orientation || 'portrait') === o ? 'check' : undefined, run: () => act('page', { orientation: o }) })))} />
            <Button tall icon="file" label={t('Size')} title={t('Size — now {paper}', { paper: view.page?.paper || 'A4' })} onClick={(e) => menu.open(e, ['A4', 'Letter', 'Legal', 'A3'].map((s) => ({ label: s, icon: (view.page?.paper || 'A4') === s ? 'check' : undefined, run: () => act('page', { size: s }) })))} />
            <Button icon="grid" label={t('Print Area')} title={t('Print Area — the selection becomes what prints, or the print area is cleared')} onClick={(e) => menu.open(e, [
              { label: t('Set print area (the selection)'), icon: 'grid', run: () => act('printArea', 'set') },
              { label: t('Clear print area'), run: () => act('printArea', 'clear') },
            ])} />
            <Button icon="minus" label={t('Breaks')} title={t('Breaks — {breaks}', { breaks: breaksSaid(view.page) })} onClick={(e) => menu.open(e, [
              { label: t('Insert page break — above the row and left of the column of the cell'), icon: 'minus', run: () => act('page', { breaks: 'insert' }) },
              { label: t('Remove page break at the cell'), run: () => act('page', { breaks: 'remove' }) },
              { label: t('Reset all page breaks'), run: () => act('page', { breaks: 'reset' }) },
            ])} />
            <Button icon="picture" label={model?.background ? t('Delete Background') : t('Background')}
              title={model?.background ? t('Delete Background — take the picture from behind the cells') : t('Background — a picture tiled behind the cells, shown on screen and not printed, as in Excel')}
              onClick={() => act('background')} />
            <Button icon="table" label={t('Print Titles')} title={view.page?.repeatRows ? t('Print Titles — rows 1 to {rows} repeat at the top of every page', { rows: view.page.repeatRows }) : t('Print Titles — no rows repeat yet')} onClick={(e) => menu.open(e, [
              { label: t('Repeat row 1 at the top of every page'), icon: view.page?.repeatRows === 1 ? 'check' : undefined, run: () => act('page', { repeatRows: 1 }) },
              { label: t('Repeat rows 1 to 2'), icon: view.page?.repeatRows === 2 ? 'check' : undefined, run: () => act('page', { repeatRows: 2 }) },
              { label: t('Repeat rows 1 to 3'), icon: view.page?.repeatRows === 3 ? 'check' : undefined, run: () => act('page', { repeatRows: 3 }) },
              { label: t('Repeat the selected rows, from row 1'), run: () => act('page', { repeatRows: 'selection' }) },
              { label: t('No repeated rows'), icon: !view.page?.repeatRows ? 'check' : undefined, run: () => act('page', { repeatRows: 0 }) },
            ])} />
          </Group>
          <Group label={t('Scale to Fit')}>
            <Button icon="minus" label={t('Width')} title={view.page?.fit === 'width' || view.page?.fit === 'page' ? t('Width — all the columns on one page across') : t('Width — automatic')} onClick={(e) => menu.open(e, [
              { label: t('Automatic'), icon: view.page?.fit === 'width' || view.page?.fit === 'page' ? undefined : 'check', run: () => act('page', { fit: view.page?.fit === 'page' ? 'height' : 'none' }) },
              { label: t('1 page'), icon: view.page?.fit === 'width' || view.page?.fit === 'page' ? 'check' : undefined, run: () => act('page', { fit: view.page?.fit === 'height' ? 'page' : 'width' }) },
            ])} />
            <Button icon="minus" label={t('Height')} title={view.page?.fit === 'height' || view.page?.fit === 'page' ? t('Height — all the rows on one page down') : t('Height — automatic')} onClick={(e) => menu.open(e, [
              { label: t('Automatic'), icon: view.page?.fit === 'height' || view.page?.fit === 'page' ? undefined : 'check', run: () => act('page', { fit: view.page?.fit === 'page' ? 'width' : 'none' }) },
              { label: t('1 page'), icon: view.page?.fit === 'height' || view.page?.fit === 'page' ? 'check' : undefined, run: () => act('page', { fit: view.page?.fit === 'width' ? 'page' : 'height' }) },
            ])} />
            <Button icon="zoomIn" label={t('Scale')} title={view.page?.fit && view.page.fit !== 'none' ? t('Scale — now {percent}%, set aside while fitting', { percent: Math.round((view.page?.scale ?? 1) * 100) }) : t('Scale — now {percent}%', { percent: Math.round((view.page?.scale ?? 1) * 100) })} onClick={(e) => menu.open(e, [50, 75, 100, 125, 150].map((p) => ({ label: `${p}%`, icon: Math.round((view.page?.scale ?? 1) * 100) === p && (view.page?.fit ?? 'none') === 'none' ? 'check' : undefined, run: () => act('page', { scale: p / 100 }) })))} />
          </Group>
          <Group label={t('Sheet Options')}>
            <Button icon="grid" label={t('Gridlines')} pressed={view.gridlines !== false} title={t('Show the gridlines on screen')} onClick={() => act('toggleGridlines')} />
            <Button icon="list" label={t('Headings')} pressed={view.headings !== false} title={t('Show the row and column headings')} onClick={() => act('toggleHeadings')} />
            <Button icon="textRtl" label={t('Sheet Right-to-Left')} pressed={Boolean(model?.rtl)} title={t('Sheet Right-to-Left — column A at the right and the row headings with it, as an Arabic or Hebrew sheet reads')} onClick={() => dispatch({ op: 'setRightToLeft', on: !model?.rtl })} />
          </Group>
          <Group label={t('Arrange')}>
            <Button icon="chevronUp" label={t('Bring Forward')} disabled={!nPicked} title={need ? t('Bring Forward — {need}', { need }) : t('Bring Forward — the picked object one step, or all the way, to the front')} onClick={(e) => menu.open(e, [
              { label: t('Bring Forward'), icon: 'chevronUp', run: () => act('arrange', { op: 'order', to: 'forward' }) },
              { label: t('Bring to Front'), run: () => act('arrange', { op: 'order', to: 'front' }) },
            ])} />
            <Button icon="chevronDown" label={t('Send Backward')} disabled={!nPicked} title={need ? t('Send Backward — {need}', { need }) : t('Send Backward — the picked object one step, or all the way, to the back')} onClick={(e) => menu.open(e, [
              { label: t('Send Backward'), icon: 'chevronDown', run: () => act('arrange', { op: 'order', to: 'backward' }) },
              { label: t('Send to Back'), run: () => act('arrange', { op: 'order', to: 'back' }) },
            ])} />
            <Button icon="list" label={t('Selection Pane')} pressed={arrange.pane} title={t('Selection Pane — every picture, shape, chart and slicer on the sheet, to pick, hide, show and rename')} onClick={() => act('arrange', { op: 'pane' })} />
            <Button icon="alignLeft" label={t('Align')} disabled={nPicked < 2} title={nPicked < 2 ? t('Align — select two or more objects first (Ctrl+click adds one)') : t('Align — line the picked objects up, or space them evenly')} onClick={(e) => menu.open(e, [
              ...[[t('Align Left'), 'left'], [t('Align Center'), 'center'], [t('Align Right'), 'right'], [t('Align Top'), 'top'], [t('Align Middle'), 'middle'], [t('Align Bottom'), 'bottom']]
                .map(([label, edge]) => ({ label, run: () => act('arrange', { op: 'align', edge }) })),
              { label: t('Distribute Horizontally'), disabled: nPicked < 3, title: nPicked < 3 ? t('Select three or more objects to distribute') : undefined, run: () => act('arrange', { op: 'distribute', axis: 'horizontal' }) },
              { label: t('Distribute Vertically'), disabled: nPicked < 3, title: nPicked < 3 ? t('Select three or more objects to distribute') : undefined, run: () => act('arrange', { op: 'distribute', axis: 'vertical' }) },
            ])} />
            <Button icon="grid" label={t('Group')} disabled={!nPicked} title={need ? t('Group — {need}', { need }) : t('Group — gather the picked objects into one, or take a group apart')} onClick={(e) => menu.open(e, [
              { label: t('Group'), icon: 'grid', disabled: nPicked < 2, title: nPicked < 2 ? t('Select two or more objects to group') : undefined, run: () => act('arrange', { op: 'group' }) },
              { label: t('Ungroup'), disabled: !groupPicked, title: !groupPicked ? t('Select a group to ungroup') : undefined, run: () => act('arrange', { op: 'ungroup' }) },
            ])} />
            <Button icon="rotate" label={t('Rotate')} disabled={!turnable} title={nPicked ? (turnable ? t('Rotate — turn or flip the picked shapes and pictures') : t('Rotate — charts and slicers do not turn')) : t('Rotate — {need}', { need })} onClick={(e) => menu.open(e, [
              { label: t('Rotate Right 90°'), icon: 'rotate', run: () => act('arrange', { op: 'rotate', by: 90 }) },
              { label: t('Rotate Left 90°'), run: () => act('arrange', { op: 'rotate', by: -90 }) },
              { label: t('Flip Vertical'), icon: 'flip', run: () => act('arrange', { op: 'rotate', flip: 'vertical' }) },
              { label: t('Flip Horizontal'), run: () => act('arrange', { op: 'rotate', flip: 'horizontal' }) },
            ])} />
          </Group>
        </>
      ) : null}

      {/* ── Formulas ─────────────────────────────────────────────────────── */}
      {tab === 'formulas' ? (
        <>
          <Group label={t('Function Library')}>
            <Button tall icon="formula" label={t('Insert Function')} onClick={() => openDialog('function')} />
            <Button tall icon="sum" label={t('AutoSum')} onClick={(e) =>
              menu.open(e, ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((fn) => ({ label: AUTOSUM_NAMES[fn], icon: 'sum', run: () => dispatch({ op: 'autoSum', fn }) })))
            } />
            {Object.entries(FUNCTIONS).map(([category, names]) => (
              <Button key={category} icon="formula" label={category} onClick={(e) => functionMenu(e, names)} />
            ))}
          </Group>
          <Group label={t('Defined Names')}>
            <Button tall icon="find" label={t('Name Manager')} onClick={() => openDialog('names')} />
            <Button icon="plus" label={t('Define Name')} onClick={() => openDialog('names')} />
            <Button icon="formula" label={t('Use in Formula')} onClick={(e) => menu.open(e, (model?.names || []).length
              ? model.names.map((n) => ({ label: n.name, run: () => act('insertFunction', n.name, { bare: true }) }))
              : [{ label: t('No names defined yet'), run: () => openDialog('names') }])} />
            <Button icon="table" label={t('Create from Selection')} title={t('Create from Selection — a name for each column of the block, from its header')} onClick={() => act('namesFromSelection')} />
          </Group>
          <Group label={t('Formula Auditing')}>
            <Button icon="chevronRight" label={t('Trace Precedents')} title={t('Trace Precedents — arrows from the cells this formula reads')} onClick={() => act('trace', 'precedents')} />
            <Button icon="chevronLeft" label={t('Trace Dependents')} title={t('Trace Dependents — arrows to the formulas that read this cell')} onClick={() => act('trace', 'dependents')} />
            <Button icon="close" label={t('Remove Arrows')} title={t('Remove Arrows — take the tracing arrows off the grid')} onClick={() => act('removeArrows')} />
            <Button icon="formula" label={t('Show Formulas')} pressed={Boolean(view.formulas)} title={t('Show every formula instead of its result (Ctrl+`)')} onClick={() => act('toggleFormulas')} />
            <Button icon="check" label={t('Error Checking')} pressed={Boolean(model?.errors)} title={t('Error Checking — every cell whose value is an error, or part of a circular reference')} onClick={() => act('errorCheck')} />
            <Button icon="eye" label={t('Evaluate Formula')} title={t("Evaluate Formula — the active cell's formula worked out a part at a time, with Step In to the cells it reads")} onClick={() => act('evaluateFormula')} />
            <Button icon="eye" label={t('Watch Window')} pressed={Boolean(model?.watches)} title={t('Watch Window — a list of chosen cells whose value stays visible wherever you scroll')} onClick={() => act('watchOpen')} />
          </Group>
          <Group label={t('Calculation')}>
            <Button tall icon="settings" label={t('Calculation Options')} title={{ auto: t('Calculation Options — now automatic'), autoNoTable: t('Calculation Options — now automatic except for data tables'), manual: t('Calculation Options — now manual') }[model?.calc?.mode || 'auto']} onClick={(e) => menu.open(e, [
              ['auto', t('Automatic')], ['autoNoTable', t('Automatic except for data tables')], ['manual', t('Manual')],
            ].map(([mode, label]) => ({ label, icon: (model?.calc?.mode || 'auto') === mode ? 'check' : undefined, run: () => act('calcMode', mode) })))} />
            <Rows>
              <Button icon="refresh" label={t('Calculate Now')} title={t('Calculate Now — every formula an edit has reached, on every sheet (F9)')} onClick={() => act('calculate', 'workbook')} />
              <Button icon="refresh" label={t('Calculate Sheet')} title={t("Calculate Sheet — this sheet's formulas only (Shift+F9)")} onClick={() => act('calculate', 'sheet')} />
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Data ─────────────────────────────────────────────────────────── */}
      {tab === 'data' ? (
        <>
          <Group label={t('Get & Transform Data')}>
            <Button tall icon="import" label={t('From Text/CSV')} title={t('From Text/CSV — a CSV or TSV file opened as a workbook, or read into this one by a query')} onClick={(e) => menu.open(e, [
              { label: t('Into this workbook, as a query…'), icon: 'table', run: () => act('queryFromCsv') },
              { label: t('Open as a workbook…'), icon: 'open', run: openFile },
            ])} />
            <Soon tall icon="globe" label={t('From Web')} why={t('Fetching a table from a web page is a network feature this suite does not do on its own.')} />
            <Button icon="table" label={t('From Table/Range')} title={t('From Table/Range — a query on the table or the list round the cell, shaped in the Power Query Editor and loaded on a sheet of its own')} onClick={() => act('queryFromRange')} />
            <Button icon="clock" label={t('Recent Sources')} title={t('Recent Sources — a file a query read lately, in a new query')} onClick={(e) => {
              const list = recentSources();
              menu.open(e, list.length ? list.map((r) => ({ label: r.sourceText, title: r.source.path, icon: 'file', run: () => act('recentSource', r) })) : [{ label: t('No recent sources yet'), disabled: true }]);
            }} />
          </Group>
          <Group label={t('Queries & Connections')}>
            <Button tall icon="refresh" label={t('Refresh All')} title={t('Refresh All — every query run again on its source, every pivot table refreshed, the workbook recalculated')} onClick={() => act('refreshAll')} />
            <Button icon="list" label={t('Queries & Connections')} title={t("Queries & Connections — the workbook's queries, each refreshed, edited or deleted")} onClick={() => act('queriesPane')} />
          </Group>
          <Group label={t('Sort & Filter')}>
            <Button tall icon="sort" label={t('A → Z')} title={t('Sort the selection ascending by its first column')} onClick={() => commands['sheet.sortAsc'].run()} />
            <Button tall icon="sort" label={t('Z → A')} title={t('Sort the selection descending by its first column')} onClick={() => commands['sheet.sortDesc'].run()} />
            <Button tall icon="sort" label={t('Sort')} title={t('Sort by up to three columns, each A to Z or Z to A')} onClick={() => act('sortDialog')} />
            <Button tall icon="filter" label={t('Filter')} pressed={model?.filtered} onClick={() => dispatch({ op: 'autoFilter' })} />
            <Button icon="close" label={t('Clear')} title={t('Clear — show every row an advanced filter hid')} onClick={() => act('clearFilter')} />
            <Button icon="filter" label={t('Advanced')} title={t('Advanced — filter the list by a criteria range, in place or copied to another place')} onClick={() => act('advancedDialog')} />
          </Group>
          <Group label={t('Data Tools')}>
            <Button tall icon="table" label={t('Text to Columns')} title={t("Text to Columns — split the selected column's cells on a delimiter into the cells to the right")} onClick={(e) => menu.open(e, [[t('Comma'), 'comma'], [t('Tab'), 'tab'], [t('Semicolon'), 'semicolon'], [t('Space'), 'space']].map(([label, delimiter]) => ({ label, run: () => act('textToColumns', delimiter) })))} />
            <Button icon="wand" label={t('Flash Fill')} title={t('Flash Fill — fill the column from an example or two typed in it (Ctrl+E)')} onClick={() => act('flashFill')} />
            <Button icon="minus" label={t('Remove Duplicates')} title={t('Remove Duplicates — rows that repeat an earlier one in the selection, or the block round the cell, go')} onClick={() => act('removeDuplicates')} />
            <Button icon="check" label={t('Data Validation')} onClick={() => openDialog('validation')} />
            <Button icon="sum" label={t('Consolidate')} title={t('Consolidate — ranges on this sheet or others summed (or averaged, counted…) at the active cell, by position or by their labels')} onClick={() => act('consolidateDialog')} />
          </Group>
          <Group label={t('Forecast')}>
            <Button tall icon="wand" label={t('What-If Analysis')} onClick={(e) => menu.open(e, [
              { label: t('Goal Seek…'), run: () => openDialog('goalSeek') },
              { label: t('Data Table…'), run: () => openDialog('dataTable') },
            ])} />
            <Button tall icon="chart" label={t('Forecast Sheet')} title={t("Forecast Sheet — a new sheet carrying a timeline's values forward, with FORECAST.ETS, confidence bounds and a chart")} onClick={() => act('forecastDialog')} />
          </Group>
          <Group label={t('Outline')}>
            <Button tall icon="plus" label={t('Group')} title={t('Group — the selected rows or columns one outline level deeper (Shift+Alt+Right)')} onClick={() => act('group')} />
            <Button tall icon="minus" label={t('Ungroup')} title={t('Ungroup — the selected rows or columns one level shallower, or clear the whole outline (Shift+Alt+Left)')} onClick={(e) => menu.open(e, [
              { label: t('Ungroup…'), run: () => act('ungroup') },
              { label: t('Clear Outline'), run: () => act('clearOutline') },
            ])} />
            <Button tall icon="sum" label={t('Subtotal')} title={t('Subtotal — a total row at each change in a column, a Grand Total, and the outline round them')} onClick={() => act('subtotalDialog')} />
            <Rows>
              <Button icon="plus" label={t('Show Detail')} title={t('Show Detail — open the folded group at the active cell')} onClick={() => act('showDetail')} />
              <Button icon="minus" label={t('Hide Detail')} title={t('Hide Detail — fold the group the active cell is in')} onClick={() => act('hideDetail')} />
              <Button icon="grid" label={t('Auto Outline')} title={t('Auto Outline — group the rows each total adds up above it and the columns each row total adds up to its left, nested as deep as the totals go')} onClick={() => act('autoOutline')} />
            </Rows>
          </Group>
          <Group label={t('Export')}>
            <Button icon="export" label="CSV" onClick={() => exportAs('csv')} />{/* words-ok: a file format's name */}
            <Button icon="export" label="TSV" onClick={() => exportAs('tsv')} />{/* words-ok: a file format's name */}
            <Button icon="pdf" label="PDF" onClick={() => exportAs('pdf')} />{/* words-ok: a file format's name */}
          </Group>
        </>
      ) : null}

      {/* ── Review ───────────────────────────────────────────────────────── */}
      {tab === 'review' ? (
        <>
          <Group label={t('Proofing')}>
            <Button tall icon="check" label={t('Spelling')} pressed={review?.pane === 'editor'} title={t("Spelling (F7) — check the text in this sheet's cells from the active cell, then the other sheets")} onClick={() => review?.startSpelling()} />
            <Button icon="find" label={t('Thesaurus')} title={t('Thesaurus (Shift+F7) — words of like meaning for the word in the active cell')} onClick={() => review?.openThesaurus()} />
            <Button icon="info" label={t('Workbook Statistics')} onClick={() => openDialog('statistics')} />
          </Group>
          <Group label={t('Accessibility')}>
            <Button tall icon="shield" label={t('Check Accessibility')} pressed={review?.pane === 'accessibility'} title={t('Check Accessibility — alt text, sheet names, merged cells, table headers, contrast and links, with a fix for each')} onClick={() => review?.openAccessibility()} />
          </Group>
          <Group label={t('Language')}>
            <Soon tall icon="globe" label={t('Translate')} why={t('Translation is an online service this suite does not call.')} />
          </Group>
          <Group label={t('Comments')}>
            <Button tall icon="reply" label={t('New Comment')} title={model?.thread ? t('New Comment — a reply at the end of this cell’s thread (Ctrl+Alt+M)') : t('New Comment — a conversation on this cell: replies, resolve, reopen (Ctrl+Alt+M)')} onClick={() => act('newComment')} />
            <Rows>
              <Button icon="close" label={t('Delete')} title={model?.thread ? t('Delete — the comment thread on this cell, replies and all') : t('Delete — this cell has no comment thread to delete')} disabled={!model?.thread} onClick={() => act('deleteThread')} />
              <Button icon="chevronLeft" label={t('Previous')} title={t('Previous — the comment before this cell, across the sheets')} onClick={() => act('stepComment', 'prev')} />
              <Button icon="chevronRight" label={t('Next')} title={t('Next — the comment after this cell, across the sheets')} onClick={() => act('stepComment', 'next')} />
            </Rows>
            <Button tall icon="eye" label={t('Show Comments')} pressed={Boolean(model?.comments)} title={t('Show Comments — every comment thread in the workbook in a pane, open or resolved')} onClick={() => act('commentsOpen')} />
          </Group>
          <Group label={t('Protect')}>
            <Button tall icon="lock" label={protectedSheet ? t('Unprotect Sheet') : t('Protect Sheet')} pressed={protectedSheet}
              title={protectedSheet
                ? (model?.protection?.hasPassword ? t('Unprotect Sheet — locked cells take edits again; it asks for the password') : t('Unprotect Sheet — locked cells take edits again'))
                : t('Protect Sheet — locked cells refuse edits, with an optional password')}
              onClick={() => act('protectSheet')} />
            <Button tall icon="lock" label={protectedBook ? t('Unprotect Workbook') : t('Protect Workbook')} pressed={protectedBook}
              title={protectedBook
                ? (model?.workbookProtection?.hasPassword ? t('Unprotect Workbook — sheets can be added, deleted, renamed, moved and hidden again; it asks for the password') : t('Unprotect Workbook — sheets can be added, deleted, renamed, moved and hidden again'))
                : t('Protect Workbook — no sheet added, deleted, renamed, moved or hidden, with an optional password')}
              onClick={() => act('protectWorkbook')} />
            <Button icon="lock" label={format.locked === false ? t('Unlocked') : t('Locked')} title={t('Whether these cells are locked when the sheet is protected')} pressed={format.locked !== false} onClick={() => setFormat({ locked: format.locked === false })} />
            <Button icon="lock" label={t('Allow Edit Ranges')} title={(model?.editRanges || []).length ? t('Allow Edit Ranges — ranges that stay editable when the sheet is protected, each with an optional password ({count} now)', { count: model.editRanges.length }) : t('Allow Edit Ranges — ranges that stay editable when the sheet is protected, each with an optional password')} onClick={() => act('editRanges')} />
          </Group>
          <Group label={t('Rules')}>
            <Button tall icon="wand" label={t('Conditional')} onClick={() => openDialog('conditional')} />
            <Button tall icon="check" label={t('Validation')} onClick={() => openDialog('validation')} />
          </Group>
          <Group label={t('Ink')}>
            <Button icon="eye" label={t('Hide Ink')} pressed={Boolean(ink?.hide)} title={t("Hide Ink — the sheet's strokes out of sight while you work; they stay in the file")} onClick={() => act('hideInk')} />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
      {tab === 'view' ? (
        <>
          <Group label={t('Workbook Views')}>
            <Button tall icon="grid" label={t('Normal')} title={t('Normal — the sheet as a grid')} pressed={(model?.viewMode || 'normal') === 'normal'} onClick={() => act('view', 'normal')} />
            <Button tall icon="file" label={t('Page Break Preview')} title={t('Page Break Preview — where the pages will break when printed; drag a break to move it')} pressed={model?.viewMode === 'pageBreakPreview'} onClick={() => act('view', 'pageBreakPreview')} />
            <Button tall icon="file" label={t('Page Layout')} title={t('Page Layout — the sheet on the pages it prints on, with margins, header and footer; click a header to write it')} pressed={model?.viewMode === 'pageLayout'} onClick={() => act('view', 'pageLayout')} />
            <Button icon="list" label={t('Custom Views')} disabled={Boolean(model?.customViewsBlocked)}
              title={model?.customViewsBlocked ? t('Custom Views — {why}', { why: model.customViewsBlocked }) : (model?.customViews || []).length ? t('Custom Views — keep the way the workbook looks under a name, and show it again ({count} kept)', { count: model.customViews.length }) : t('Custom Views — keep the way the workbook looks under a name, and show it again')}
              onClick={() => act('customViews')} />
          </Group>
          <Group label={t('Show')}>
            <Button icon="grid" label={t('Gridlines')} pressed={view.gridlines !== false} onClick={() => act('toggleGridlines')} />
            <Button icon="formula" label={t('Formula Bar')} pressed={view.formulaBar !== false} onClick={() => act('toggleFormulaBar')} />
            <Button icon="list" label={t('Headings')} pressed={view.headings !== false} onClick={() => act('toggleHeadings')} />
            <Button icon="minus" label={t('Ruler')} pressed={model?.viewMode === 'pageLayout' && model?.showRuler !== false} disabled={model?.viewMode !== 'pageLayout'}
              title={model?.viewMode === 'pageLayout' ? t('Ruler — centimetres along the top and side of the pages') : t('Ruler — shown in Page Layout view')}
              onClick={() => act('toggleRuler')} />
          </Group>
          <Group label={t('Zoom')}>
            <Button tall icon="zoomIn" label={t('Zoom')} onClick={(e) => menu.open(e, [50, 75, 100, 125, 150, 200].map((z) => ({ label: `${z}%`, run: () => act('zoom', z / 100) })))} />
            <Button icon="check" label="100%" onClick={() => act('zoom', 1)} />
            <Button icon="zoomIn" title={t('Zoom in')} onClick={() => act('zoom', (view.zoom ?? 1) + 0.1)} />
            <Button icon="zoomOut" title={t('Zoom out')} onClick={() => act('zoom', (view.zoom ?? 1) - 0.1)} />
          </Group>
          <Group label={t('Window')}>
            <Button tall icon="new" label={t('New Window')} title={t('This workbook in a second window')} onClick={() => act('newWindow')} />
            <Button tall icon="freeze" label={t('Freeze Panes')} pressed={isFrozen} onClick={(e) => menu.open(e, [
              { label: isFrozen ? t('Unfreeze panes') : t('Freeze panes at the selection'), icon: 'freeze', run: () => act('freeze', isFrozen ? 'none' : 'here') },
              { label: t('Freeze top row'), run: () => act('freeze', 'row') },
              { label: t('Freeze first column'), run: () => act('freeze', 'col') },
              { label: t('Choose…'), run: () => openDialog('freeze') },
            ])} />
            <Button icon="grid" label={t('Arrange All')} title={t('Arrange All — every Worksheets window, tiled')} onClick={() => act('arrangeWindows', 'tile')} />
            <Button icon="minus" label={t('Split')} title={t('Split — the window in four panes at the active cell (two in its first row or column), each scrolling on its own; press again to take it away')} pressed={Boolean(model?.split)} onClick={() => act('split')} />
            <Button icon="eye" label={t('Hide')} title={t('Hide — this window put away, its work kept; Unhide brings it back')} onClick={() => act('hideWindow')} />
            <Button icon="eye" label={t('Unhide')} title={t('Unhide — a hidden Worksheets window, back on screen')} onClick={(e) => act('unhideMenu', e)} />
            <Button icon="maximize" label={t('Full Screen')} onClick={() => shell.win.fullscreen({})} />
          </Group>
          <Group label={t('Macros')}>
            <Soon tall icon="settings" label={t('Macros')} why={t('VBA is preserved in the file and never run: a spreadsheet that runs code it received in an email is how ransomware starts.')} />
          </Group>
        </>
      ) : null}

      {/* ── Automate ─────────────────────────────────────────────────────── */}
      {tab === 'automate' ? (
        <>
          <Group label={t('Scripting Tools')}>
            <Button tall icon="textbox" label={t('New Script')} className="sh-new-script" title={t('New Script — a script in the shape of Office Scripts, function main(workbook), written in the Code Editor and run on this workbook')} onClick={() => act('scripts', 'new')} />
            <Button tall icon="video" label={t('Record Actions')} className="sh-record-actions" pressed={Boolean(view.recordingScript)} title={view.recordingScript ? t('Record Actions — recording what you type and format; press again to stop and see the script') : t('Record Actions — what you type and format becomes a script')} onClick={() => act('scriptRecord')} />
            <Button icon="list" label={t('All Scripts')} className="sh-all-scripts" title={t('All Scripts — the scripts kept on this computer, to run, edit or delete')} onClick={() => act('scripts')} />
          </Group>
          <Group label={t('Power Automate')}>
            <Soon tall icon="refresh" label={t('Automate a Task')} why={t('Power Automate is a Microsoft cloud service.')} />
          </Group>
        </>
      ) : null}

      {/* ── Help ─────────────────────────────────────────────────────────── */}
      {tab === 'help' ? (
        <>
          <Group label={t('Help & Support')}>
            <Button tall icon="info" label={t('Help')} title={t('The Worksheets guide on office.rutba.io')} onClick={() => act('help')} />
            <Button tall icon="send" label={t('Feedback')} title={t('Tell us what is wrong or missing')} onClick={() => act('feedback')} />
            <Button icon="list" label={t('Keyboard Shortcuts')} onClick={() => openDialog('shortcuts')} />
            <Button icon="info" label={t('About')} onClick={() => act('about')} />
          </Group>
          <Group label={t('File')}>
            <Button tall icon="new" label={t('New')} onClick={() => shell.win.create({ app: 'sheets' })} />
            <Button tall icon="open" label={t('Open')} onClick={openFile} />
            <Button tall icon="save" label={t('Save')} onClick={() => save(false)} />
          </Group>
        </>
      ) : null}
    </Ribbon>
  );
}
