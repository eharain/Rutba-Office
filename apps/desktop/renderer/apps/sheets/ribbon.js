// The Worksheets ribbon.
//
// Laid out the way Excel lays it out — Home, Insert, Draw, Page Layout,
// Formulas, Data, Review, View, Automate, Help — with every group Excel has,
// because a person arriving from Excel should not have to learn where
// anything is. What the engine can do is wired to it; what it cannot is
// drawn where Excel draws it, disabled, with a title that says why. A ribbon
// that hides its gaps is a ribbon nobody fixes.

import React from 'react';
import { Ribbon, Group, Rows, Button, Separator, Select, Input, t, tn } from '@rutba/office-ui';
import { WARP_PRESETS, WARP_MORE, warpLabel } from '@rutba/drawing/warp';
import { catalogByCategory } from '@rutba/formula';
import { numberFormats } from './dialogs.js';
import { recentSources } from './queries.js';
import { MARGIN_PRESETS as PRINT_MARGINS } from '../../print.js';
import { wordArtMenu } from '../../wordart.js';
import { PEN_COLOURS, PEN_WIDTHS } from '../../ink-geometry.js';
import { ColourSplit } from '../colour/palette.js';

/** The palette a toolbar offers before it offers a colour picker. */
const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Consolas'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];
const AUTOSUM_NAMES = { SUM: t('common.sum'), AVERAGE: t('common.average'), COUNT: t('common.count'), MAX: t('common.max'), MIN: t('common.min') };

/**
 * A border delta names edges. The engine keeps the edges it is not told about,
 * writes the ones it is, and deletes the ones set to null — so "all" is four
 * thin black edges and "none" is four nulls.
 */
const THIN = { style: 'thin', colour: '#000000' };
const MEDIUM = { style: 'medium', colour: '#000000' };
const BORDERS = [
  [t('sheets.ribbon.all-borders'), { top: THIN, bottom: THIN, left: THIN, right: THIN }],
  [t('common.outside-borders'), { top: THIN, bottom: THIN, left: THIN, right: THIN }],
  [t('sheets.ribbon.thick-outside'), { top: MEDIUM, bottom: MEDIUM, left: MEDIUM, right: MEDIUM }],
  [t('common.top-border'), { top: THIN }],
  [t('common.bottom-border'), { bottom: THIN }],
  [t('common.left-border'), { left: THIN }],
  [t('common.right-border'), { right: THIN }],
  [t('sheets.ribbon.thick-bottom-border'), { bottom: MEDIUM }],
  [t('sheets.ribbon.double-bottom-border'), { bottom: { style: 'double', colour: '#000000' } }],
  [t('sheets.ribbon.top-bottom-border'), { top: THIN, bottom: THIN }],
  [t('common.no-border'), { top: null, bottom: null, left: null, right: null }],
];

/** What the engine can draw, by the names it draws them under. */
/** A PivotChart can be any of these; never a scatter, as in Excel. */
const PIVOT_CHARTS = [
  ['column', t('common.column')], ['bar', t('common.bar')], ['line', t('common.line')], ['area', t('common.area')], ['pie', t('common.pie')], ['doughnut', t('sheets.doughnut')],
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
  ['column', t('common.column')], ['bar', t('common.bar')], ['line', t('common.line')], ['area', t('common.area')], ['pie', t('common.pie')], ['doughnut', t('sheets.doughnut')],
];
const SHAPES = [
  ['rect', t('common.rectangle')], ['roundRect', t('common.rounded-rectangle')], ['ellipse', t('common.ellipse')], ['line', t('common.line')],
  ['triangle', t('common.triangle')], ['diamond', t('common.diamond')], ['rightArrow', t('common.arrow-right')], ['leftArrow', t('common.arrow-left')],
  ['upArrow', t('common.arrow-up')], ['downArrow', t('common.arrow-down')], ['pentagon', t('common.pentagon')], ['hexagon', t('common.hexagon')],
  ['star5', t('common.star')], ['plus', t('common.plus')], ['chevron', t('common.chevron')], ['parallelogram', t('common.parallelogram')], ['trapezoid', t('common.trapezoid')],
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
  [t('sheets.ribbon.horizontal-text'), 0],
  [t('sheets.ribbon.angle-anticlockwise'), 45],
  [t('sheets.ribbon.angle-clockwise'), 135],
  [t('sheets.ribbon.vertical-text'), 255],
  [t('sheets.ribbon.rotate-text-up'), 90],
  [t('sheets.ribbon.rotate-text-down'), 180],
];

/** The print dialog's three margin presets by their lower-case names, so the tab and the dialog agree. */
export const MARGIN_PRESETS = Object.fromEntries(Object.entries(PRINT_MARGINS).map(([name, m]) => [name.toLowerCase(), m]));

/** Which preset a file's margins are, or null when they are its own. */
/** A column's letters, as the grid heads it: 0 is A, 26 is AA. */
function letterOf(c) {
  let s = '';
  for (let k = Number(c) + 1; k > 0; k = Math.floor((k - 1) / 26)) s = String.fromCharCode(65 + ((k - 1) % 26)) + s;
  return s;
}

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
  return n ? tn(n, 'sheets.ribbon.count-page-break') : t('sheets.ribbon.no-page-breaks');
}

export const CELL_STYLES = [
  [t('common.normal'), { bold: false, italic: false, fill: null, fontColour: null, border: { top: null, bottom: null, left: null, right: null }, numberFormat: 'General' }],
  [t('sheets.ribbon.heading-1'), { bold: true, fontSize: 15, fontColour: '#1F3864', border: { bottom: { style: 'medium', colour: '#4472C4' } } }],
  [t('sheets.ribbon.heading-2'), { bold: true, fontSize: 13, fontColour: '#1F3864', border: { bottom: { style: 'medium', colour: '#A9C4E9' } } }],
  [t('sheets.ribbon.heading-3'), { bold: true, fontSize: 11, fontColour: '#1F3864', border: { bottom: { style: 'medium', colour: '#B4C6E7' } } }],
  [t('common.title'), { bold: true, fontSize: 18, fontColour: '#1F3864' }],
  [t('common.total'), { bold: true, border: { top: THIN, bottom: { style: 'double', colour: '#4472C4' } } }],
  [t('sheets.ribbon.good'), { fill: '#C6EFCE', fontColour: '#006100' }],
  [t('sheets.ribbon.bad'), { fill: '#FFC7CE', fontColour: '#9C0006' }],
  [t('sheets.ribbon.neutral'), { fill: '#FFEB9C', fontColour: '#9C5700' }],
  [t('sheets.ribbon.input'), { fill: '#FFCC99', fontColour: '#3F3F76', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }],
  [t('sheets.ribbon.calculation'), { bold: true, fill: '#F2F2F2', fontColour: '#FA7D00', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }],
  [t('sheets.ribbon.check-cell'), { bold: true, fill: '#A5A5A5', fontColour: '#FFFFFF', border: { top: MEDIUM, bottom: MEDIUM, left: MEDIUM, right: MEDIUM } }],
  [t('sheets.note'), { fill: '#FFFFCC', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }],
  [t('sheets.ribbon.warning-text'), { fontColour: '#FF0000' }],
  [t('common.currency'), { numberFormat: '"£"#,##0.00' }],
  [t('common.percent'), { numberFormat: '0%' }],
  [t('sheets.comma'), { numberFormat: '#,##0.00' }],
];

/** Chart Design's title field: the picked chart's title, set on Enter or on leaving the field. */
function ChartTitleField({ chart, onSet }) {
  const [text, setText] = React.useState(chart.title || '');
  React.useEffect(() => setText(chart.title || ''), [chart.id, chart.title]);
  const commit = () => { if (text.trim() !== (chart.title || '')) onSet(text.trim()); };
  return (
    <Input
      className="rw-input sh-chart-title"
      value={text}
      placeholder={t('sheets.ribbon.chart-title')}
      aria-label={t('sheets.ribbon.chart-title')}
      style={{ width: 220 }}
      onChange={(e) => setText(e.target.value)}
      onKeyDown={(e) => {
        e.stopPropagation();
        if (e.key === 'Enter') { e.preventDefault(); commit(); }
        if (e.key === 'Escape') setText(chart.title || '');
      }}
      onBlur={commit}
    />
  );
}

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
  const out = { [t('sheets.ribbon.recently-used')]: ['SUM', 'AVERAGE', 'IF', 'COUNT', 'MAX', 'MIN', 'VLOOKUP', 'ROUND'] };
  for (const { category, functions } of catalogByCategory()) out[t(category)] = functions.map((f) => f.name);
  return out;
})();

/** Each category of the Function Library its own picture, as Excel gives them. */
const CATEGORY_ICONS = {
  [t('sheets.ribbon.recently-used')]: 'clock', [t('common.financial')]: 'financial', [t('common.logical')]: 'logical', [t('common.text')]: 'textbox',
  [t('common.date-and')]: 'calendar', [t('common.lookup-and')]: 'find', [t('common.math-and')]: 'mathTrig', [t('common.statistical')]: 'chartColumn',
  [t('common.information')]: 'info',
};

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


export default function SheetsRibbon({
  tab, setTab, model, dispatch, commands, shell, menu, save, openFile, exportAs, doc, sel, openDialog, act, view = {}, review = null, arrange = { picked: [], pane: false }, ink = null,
  redaction = null,
}) {
  // Page Layout → Arrange acts on the drawings picked on the sheet.
  const nPicked = arrange.picked.length;
  const need = nPicked ? null : t('sheets.ribbon.select-picture-shape');
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
    { label: t('sheets.ribbon.light'), icon: 'table', run: () => act('table', { style: 'TableStyleLight9', stripes: true }) },
    { label: t('sheets.ribbon.medium'), icon: 'table', run: () => act('table', { style: 'TableStyleMedium2', stripes: true }) },
    { label: t('sheets.ribbon.medium-plain'), icon: 'table', run: () => act('table', { style: 'TableStyleMedium2', stripes: false }) },
    { label: t('sheets.ribbon.dark-banded'), icon: 'table', run: () => act('table', { style: 'TableStyleDark1', stripes: true }) },
  ]);
  const functionMenu = (event, names) =>
    menu.open(event, names.map((fn) => ({ label: fn, icon: 'formula', run: () => act('insertFunction', fn) })));

  return (
    <Ribbon
      tabs={[
        { id: 'home', label: t('common.home') },
        { id: 'insert', label: t('common.insert') },
        { id: 'draw', label: t('common.draw') },
        { id: 'layout', label: t('sheets.page-layout') },
        { id: 'formulas', label: t('sheets.formulas') },
        { id: 'data', label: t('common.data') },
        { id: 'review', label: t('common.review') },
        { id: 'view', label: t('common.view') },
        { id: 'automate', label: t('sheets.ribbon.automate') },
        { id: 'help', label: t('common.help') },
        // Shape Format, as Excel's: there while one shape with words is picked.
        ...(arrange.picked.length === 1 && arrange.picked[0].kind === 'shape' && arrange.picked[0].hasText ? [{ id: 'shapeFormat', label: t('common.shape-format'), contextual: true }] : []),
        // Picture Format, as Excel's: there while one picture is picked.
        ...(arrange.picked.length === 1 && arrange.picked[0].kind === 'image' ? [{ id: 'pictureFormat', label: t('common.picture-format'), contextual: true }] : []),
        // Chart Design, as Excel's: there while one chart is picked.
        ...(arrange.picked.length === 1 && arrange.picked[0].kind === 'chart' ? [{ id: 'chartDesign', label: t('sheets.ribbon.chart-design'), contextual: true }] : []),
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title={t('common.save-ctrl-s')} onClick={() => save(false)} />
          <Button icon="undo" title={t('common.undo-ctrl-z')} disabled={!doc?.canUndo} onClick={() => commands['edit.undo'].run()} />
          <Button icon="redo" title={t('common.redo-ctrl-y')} disabled={!doc?.canRedo} onClick={() => commands['edit.redo'].run()} />
          <Button icon="sum" title={t('sheets.auto-sum')} onClick={() => commands['sheet.autoSum'].run()} />
          <Button icon="print" title={t('common.print')} onClick={() => exportAs('pdf')} />
        </>
      }
    >
      {/* ── Shape Format (contextual): WordArt's Transform ──────────────── */}
      {tab === 'shapeFormat' && arrange.picked.length === 1 && arrange.picked[0].hasText ? (
        <>
          <Group label={t('common.text-effects')}>
            {WARP_PRESETS.map((p) => (
              <Button key={p.id} tall icon="wand" label={t(p.label)} className="sh-warp" data-preset={p.id} pressed={(arrange.picked[0].textWarp || 'textNoShape') === p.id} title={p.id === 'textNoShape' ? t('common.no-transform-words') : p.label === 'Button' ? t('common.transform-words-along-button') : p.label === 'Circle' ? t('common.transform-words-along-circle') : t('common.transform-words-along')} onClick={() => act('textWarp', { id: arrange.picked[0].id, preset: p.id })} />
            ))}
            {/* The warps, the rest of Office's gallery, from a menu. */}
            <Button dropdown tall icon="wand" label={t('common.more')} className="sh-warp-more" pressed={WARP_MORE.some((p) => p.id === arrange.picked[0].textWarp)} title={WARP_MORE.some((p) => p.id === arrange.picked[0].textWarp) ? t('common.transform-now-warp', { warp: t(warpLabel(arrange.picked[0].textWarp)) }) : t('common.more-transforms-warps')} onClick={(e) => menu.open(e, WARP_MORE.map((p) => ({ label: t(p.label), icon: arrange.picked[0].textWarp === p.id ? 'check' : 'wand', run: () => act('textWarp', { id: arrange.picked[0].id, preset: p.id }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Picture Format (contextual): Edit in Image ──────────────────── */}
      {tab === 'pictureFormat' && arrange.picked.length === 1 && arrange.picked[0].kind === 'image' ? (
        <>
          <Group label={t('common.edit')}>
            <Button tall icon="image" label={t('common.edit-in')} title={t('common.edit-image-crop')} onClick={() => act('editInImage', { id: arrange.picked[0].id })} data-role="edit-in-image" />
          </Group>
          <Group label={t('common.arrange')}>
            <Rows>
              <Button icon="layerUp" label={t('common.bring')} onClick={() => dispatch({ op: 'reorderDrawings', ids: [arrange.picked[0].id], to: 'forward' })} />
              <Button icon="layerDown" label={t('common.send-backward')} onClick={() => dispatch({ op: 'reorderDrawings', ids: [arrange.picked[0].id], to: 'backward' })} />
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Chart Design (contextual): the chart's title ────────────────── */}
      {tab === 'chartDesign' && arrange.picked.length === 1 && arrange.picked[0].kind === 'chart' ? (
        <>
          <Group label={t('sheets.ribbon.chart')}>
            <Rows>
              <ChartTitleField chart={arrange.picked[0]} onSet={(title) => dispatch({ op: 'chartTitle', id: arrange.picked[0].id, title })} />
              <Button icon="close" label={t('sheets.ribbon.no-title')} className="sh-chart-notitle" disabled={!arrange.picked[0].title} title={t('sheets.ribbon.no-title-chart')} onClick={() => dispatch({ op: 'chartTitle', id: arrange.picked[0].id, title: '' })} />
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Home ─────────────────────────────────────────────────────────── */}
      {tab === 'home' ? (
        <>
          <Group label={t('common.clipboard')}>
            <Button tall icon="paste" label={t('common.paste')} title={t('common.paste-ctrl-v')} onClick={() => commands['edit.paste']?.run?.()} />
            <Button icon="chevronDown" className="sh-paste-options" title={t('sheets.ribbon.paste-options-formulas')} onClick={(e) => menu.open(e, [
              { label: t('common.paste'), icon: 'paste', run: () => commands['edit.paste']?.run?.() },
              { label: t('sheets.formulas'), run: () => act('pasteSpecial', { what: 'formulas' }) },
              { label: t('common.values'), run: () => act('pasteSpecial', { what: 'values' }) },
              { label: t('sheets.ribbon.formatting'), run: () => act('pasteSpecial', { what: 'formats' }) },
              { label: t('sheets.transpose'), run: () => act('pasteSpecial', { what: 'all', transpose: true }) },
              { label: t('sheets.paste'), run: () => commands['edit.pasteSpecial']?.run?.() },
            ])} />
            <Rows>
              <>
                <Button icon="cut" label={t('common.cut')} onClick={() => commands['edit.cut']?.run?.()} />
                <Button icon="copy" label={t('common.copy')} onClick={() => commands['edit.copy'].run()} />
              </>
              <Button icon="brush" label={t('common.format-painter')} title={t('sheets.ribbon.copy-formatting-next')} onClick={() => dispatch({ op: 'formatBrush' })} />
            </Rows>
          </Group>

          <Group label={t('common.font')}>
            <Rows>
              <>
                <Select value={format.fontName || 'Calibri'} onChange={(e) => setFormat({ fontName: e.target.value })} style={{ width: 118 }} title={t('common.font')}>
                  {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
                  {format.fontName && !FONTS.includes(format.fontName) ? <option value={format.fontName}>{format.fontName}</option> : null}
                </Select>
                <Select value={String(size)} onChange={(e) => setFormat({ fontSize: Number(e.target.value) })} style={{ width: 56 }} title={t('common.font-size')}>
                  {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
                  {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
                </Select>
                <Button icon="fontGrow" title={t('common.increase')} onClick={() => setFormat({ fontSize: nearer(1) })} />
                <Button icon="fontShrink" title={t('common.decrease')} onClick={() => setFormat({ fontSize: nearer(-1) })} />
              </>
              <>
                <Button icon="bold" title={t('common.bold-ctrl-b')} pressed={format.bold} onClick={() => setFormat({ bold: 'toggle' })} />
                <Button icon="italic" title={t('common.italic-ctrl-i')} pressed={format.italic} onClick={() => setFormat({ italic: 'toggle' })} />
                <Button icon="underline" title={t('common.underline-ctrl-u')} pressed={format.underline} onClick={() => setFormat({ underline: 'toggle' })} />
                <Button icon="strike" title={t('common.strikethrough')} pressed={format.strike} onClick={() => setFormat({ strike: 'toggle' })} />
                <Separator />
                <Button dropdown icon="borders" title={t('common.borders')} onClick={(e) => menu.open(e, BORDERS.map(([label, edges]) => ({ label, run: () => setFormat({ border: edges }) })))} />
                {/* Excel's two split buttons: the face puts on the colour last chosen, the arrow opens the palette. */}
                <ColourSplit kind="sheets.fill" icon="bucketFill" name={t('sheets.ribbon.fill-colour')} first="FFFF00" theme={model?.design?.colors} current={format.fill ?? null}
                  empty={{ label: t('common.no-fill-2'), value: null }} onPick={(value) => setFormat({ fill: value ? `#${value}` : null })} />
                <ColourSplit kind="sheets.font" icon="fontColour" name={t('common.font-colour')} first="FF0000" theme={model?.design?.colors} current={format.fontColour ?? null}
                  empty={{ label: t('common.automatic'), value: null, swatch: '#000000' }} onPick={(value) => setFormat({ fontColour: value ? `#${value}` : null })} />
              </>
            </Rows>
          </Group>

          <Group label={t('common.alignment')}>
            <Rows>
              <>
                <Button icon="alignTop" title={t('sheets.ribbon.top-align')} pressed={format.valign === 'top'} onClick={() => setFormat({ valign: 'top' })} />
                <Button icon="alignMiddle" title={t('sheets.ribbon.middle-align')} pressed={format.valign === 'center'} onClick={() => setFormat({ valign: 'center' })} />
                <Button icon="alignBottom" title={t('sheets.ribbon.bottom-align')} pressed={format.valign === 'bottom'} onClick={() => setFormat({ valign: 'bottom' })} />
                <Button dropdown icon="rotate" title={t('sheets.ribbon.text-orientation')} pressed={Boolean(format.rotation)} onClick={(e) => menu.open(e, ORIENTATIONS.map(([label, rotation]) => ({ label, icon: (format.rotation ?? -1) === rotation ? 'check' : undefined, run: () => act('orientation', rotation) })))} />
                <Separator />
                <Button icon="listBullet" label={t('common.wrap-text')} pressed={format.wrap} onClick={() => setFormat({ wrap: !format.wrap })} />
              </>
              <>
                <Button icon="alignLeft" title={t('common.align-left-2')} pressed={format.align === 'left'} onClick={() => setFormat({ align: 'left' })} />
                <Button icon="alignCenter" title={t('common.centre')} pressed={format.align === 'center'} onClick={() => setFormat({ align: 'center' })} />
                <Button icon="alignRight" title={t('common.align-right-2')} pressed={format.align === 'right'} onClick={() => setFormat({ align: 'right' })} />
                <Button icon="indentLess" title={t('common.decrease-indent')} disabled={format.indent === 0} onClick={() => setFormat({ indentBy: -1 })} />
                <Button icon="indentMore" title={t('common.increase-indent')} onClick={() => setFormat({ indentBy: 1 })} />
                <Separator />
                <Button icon="table" label={t('sheets.ribbon.merge-and')} title={t('sheets.ribbon.merge-selected-cells')} onClick={() => act('mergeCentre')} />
                <Button icon="table" title={t('common.merge-cells')} onClick={() => dispatch({ op: 'merge' })} />
                <Button icon="minus" title={t('sheets.ribbon.unmerge-cells')} onClick={() => dispatch({ op: 'unmerge' })} />
              </>
            </Rows>
          </Group>

          <Group label={t('common.number')}>
            <Rows>
              <Select value={format.numberFormat || 'General'} onChange={(e) => setFormat({ numberFormat: e.target.value })} style={{ width: 148 }} title={t('sheets.ribbon.number-format')}>
                {numberFormats().map((f) => <option key={f.label} value={f.code}>{f.label}</option>)}
                {numberFormats().some((f) => f.code === (format.numberFormat || 'General')) ? null : (
                  <option value={format.numberFormat}>{format.numberFormat}</option>
                )}
              </Select>
              <>
                {/* The mark is the button. An icon beside a currency symbol says nothing. */}
                <Button title={t('common.currency')} onClick={() => setFormat({ numberFormat: '"£"#,##0.00' })} label="£" />
                <Button title={t('sheets.ribbon.percent-style')} onClick={() => setFormat({ numberFormat: '0%' })} label="%" />
                <Button title={t('sheets.ribbon.comma-style')} onClick={() => setFormat({ numberFormat: '#,##0.00' })} label="," />
                <Button title={t('sheets.ribbon.increase-decimal')} onClick={() => setFormat({ numberFormat: withDecimals(format.numberFormat, 1) })} label=".0" />
                <Button title={t('sheets.ribbon.decrease-decimal')} onClick={() => setFormat({ numberFormat: withDecimals(format.numberFormat, -1) })} label=".00" />
              </>
            </Rows>
          </Group>

          <Group label={t('common.styles')}>
            <Button tall icon="wand" label={t('sheets.ribbon.conditional-formatting')} onClick={() => openDialog('conditional')} />
            <Button tall icon="table" label={t('sheets.ribbon.format-as')} title={t('sheets.ribbon.format-table-header')} onClick={tableMenu} />
            <Button dropdown tall icon="grid" label={t('sheets.ribbon.cell-styles')} title={t('sheets.ribbon.excels-cell-styles')} onClick={(e) => menu.open(e, CELL_STYLES.map(([label, delta]) => ({ label, run: () => setFormat(delta) })))} />
          </Group>

          <Group label={t('common.cells')}>
            <Button dropdown tall icon="plus" label={t('common.insert')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.insert-sheet-rows'), icon: 'plus', run: () => commands['sheet.insertRow'].run() },
              { label: t('sheets.ribbon.insert-sheet-columns'), icon: 'plus', run: () => commands['sheet.insertCol'].run() },
            ])} />
            <Button dropdown tall icon="minus" label={t('common.delete')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.delete-sheet-rows'), icon: 'minus', run: () => commands['sheet.deleteRow'].run() },
              { label: t('sheets.ribbon.delete-sheet-columns'), icon: 'minus', run: () => commands['sheet.deleteCol'].run() },
            ])} />
            <Button dropdown tall icon="settings" label={t('common.format')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.row-height'), run: () => openDialog('rowHeight') },
              { label: t('sheets.ribbon.column-width'), run: () => openDialog('colWidth') },
              { label: t('sheets.ribbon.auto-fit'), run: () => act('autoFit') },
              '-',
              { label: protectedSheet ? t('sheets.ribbon.unprotect') : t('sheets.ribbon.protect'), icon: 'lock', run: () => dispatch({ op: protectedSheet ? 'unprotect' : 'protect' }) },
              { label: format.locked === false ? t('sheets.ribbon.lock-cells') : t('sheets.ribbon.unlock-cells'), icon: 'lock', run: () => setFormat({ locked: format.locked === false }) },
            ])} />
          </Group>

          <Group label={t('common.editing')}>
            <Button dropdown icon="sum" label={t('sheets.auto-sum')} onClick={(e) =>
              menu.open(e, ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((fn) => ({ label: AUTOSUM_NAMES[fn], icon: 'sum', run: () => dispatch({ op: 'autoSum', fn }) })))
            } />
            <Button icon="chevronDown" label={t('common.fill')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.down-ctrl-d'), run: () => act('fill', 'down') },
              { label: t('sheets.ribbon.right-ctrl-r'), run: () => act('fill', 'right') },
              { label: t('common.up'), run: () => act('fill', 'up') },
              { label: t('common.left'), run: () => act('fill', 'left') },
              { label: t('sheets.ribbon.series'), run: () => act('seriesDialog') },
            ])} />
            <Button dropdown icon="close" label={t('common.clear')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.clear-contents'), run: () => dispatch({ op: 'clear' }) },
              { label: t('sheets.ribbon.clear-formats'), run: () => setFormat(CELL_STYLES[0][1]) },
            ])} />
            <Button dropdown icon="sort" label={t('sheets.ribbon.sort-and')} onClick={(e) => menu.open(e, [
              { label: t('sheets.sort-z'), icon: 'sort', run: () => commands['sheet.sortAsc'].run() },
              { label: t('sheets.sort-z-to-a'), icon: 'sort', run: () => commands['sheet.sortDesc'].run() },
              { label: model?.filtered ? t('sheets.ribbon.remove-filter') : t('common.filter'), icon: 'filter', run: () => dispatch({ op: 'autoFilter' }) },
            ])} />
            <Button dropdown icon="find" label={t('sheets.ribbon.find-and')} onClick={(e) => menu.open(e, [
              { label: t('common.find-2'), icon: 'find', run: () => openDialog('find') },
              { label: t('common.replace-2'), icon: 'find', run: () => openDialog('replace') },
              { label: t('sheets.ribbon.go-to'), run: () => openDialog('goto') },
            ])} />
          </Group>
        </>
      ) : null}

      {/* ── Insert ───────────────────────────────────────────────────────── */}
      {tab === 'insert' ? (
        <>
          <Group label={t('common.tables')}>
            <Button tall icon="table" label={t('sheets.pivot-table')} title={t('sheets.ribbon.pivot-table-summarise')} onClick={() => act('pivotTable')} />
            {inPivot && !inPivot.unsupported ? <Button icon="list" label={t('sheets.ribbon.field-list')} title={t('sheets.ribbon.field-list-fields', { name: inPivot.name })} onClick={() => act('pivotFields')} /> : null}
            <Button tall icon="table" label={t('common.table')} title={t('sheets.ribbon.table-selection-block')} onClick={tableMenu} />
          </Group>
          <Group label={t('common.illustrations')}>
            <Button tall icon="picture" label={t('common.pictures')} title={t('sheets.ribbon.pictures-picture-file')} onClick={() => act('picture')} />
            <Button dropdown tall icon="shape" label={t('common.shapes')} onClick={(e) => menu.open(e, SHAPES.map(([geometry, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', geometry, text: '' }) })))} />
            <Button icon="qrcode" label={t('common.qr-code')} className="sh-qr-insert" title={t('sheets.ribbon.qr-code-code')} onClick={() => act('qrcode')} />
            <Button icon="star" label={t('common.icons')} title={t('sheets.ribbon.icons-one-suites')} onClick={() => act('icons')} />
            <Button icon="shape" label={t('common.smart-art')} title={t('sheets.ribbon.smart-art-list')} onClick={() => act('smartArt')} />
          </Group>
          <Group label={t('common.charts')}>
            {CHARTS.map(([kind, label]) => (
              <Button key={kind} icon="chart" label={label} title={t('sheets.ribbon.chart-chart-data', { chart: label })} onClick={() => dispatch({ op: 'insertChart', kind })} />
            ))}
            <Button dropdown icon="chartScatter" label={t('sheets.ribbon.scatter')} title={t('sheets.ribbon.scatter-x-y')} onClick={(e) => menu.open(e, [
              ['markers', t('sheets.ribbon.scatter')], ['lines', t('sheets.ribbon.scatter-straight-lines')], ['smooth', t('sheets.ribbon.scatter-smooth-lines')],
            ].map(([scatterStyle, label]) => ({ label, icon: 'chart', run: () => dispatch({ op: 'insertChart', kind: 'scatter', scatterStyle }) })))} />
            <Button icon="chart" label={t('sheets.pivot-chart')} title={inPivot ? t('sheets.ribbon.pivot-chart-chart', { name: inPivot.name }) : t('sheets.ribbon.pivot-chart-pivot')}
              onClick={(e) => (inPivot
                ? menu.open(e, PIVOT_CHARTS.map(([kind, label]) => ({ label: t('sheets.ribbon.chart-pivot-chart', { chart: label }), icon: 'chart', run: () => act('pivotChart', { kind }) })))
                : act('pivotChart'))} />
          </Group>
          <Group label={t('sheets.ribbon.sparklines')}>
            <Button tall icon="chartLine" label={t('common.line')} title={t('sheets.ribbon.line-sparkline-small')} onClick={() => openDialog('sparklineLine')} />
            <Button icon="chartColumn" label={t('common.column')} title={t('sheets.ribbon.column-sparkline-small')} onClick={() => openDialog('sparklineColumn')} />
          </Group>
          <Group label={t('common.filters')}>
            <Button tall icon="filter" label={t('common.filter')} pressed={model?.filtered} onClick={() => dispatch({ op: 'autoFilter' })} />
            <Button icon="filter" label={t('sheets.slicer')} title={t('sheets.ribbon.slicer-buttons-filter')} onClick={() => act('slicer')} />
          </Group>
          <Group label={t('common.links')}>
            <Button tall icon="link" label={t('common.link')} title={t('sheets.ribbon.link-cell-address')} onClick={() => act('link')} />
          </Group>
          <Group label={t('common.comments')}>
            <Button tall icon="comment" label={t('common.comment')} title={t('sheets.ribbon.comment-conversation')} onClick={() => act('newComment')} />
            <Button tall icon="note" label={t('sheets.note')} title={t('sheets.ribbon.note-cell-shown')} onClick={() => act('note')} />
            <Button icon="close" label={t('common.delete')} title={t('sheets.ribbon.take-note-off')} onClick={() => act('removeNote')} />
          </Group>
          <Group label={t('common.text')}>
            <Button tall icon="textbox" label={t('common.text-box')} title={t('sheets.ribbon.rectangle-words')} onClick={() => act('textBox')} />
            <Button icon="header" label={t('common.header-and')} title={t('sheets.ribbon.header-footer-prints')} onClick={() => act('headerFooter')} />
            <Button dropdown icon="wordArt" label={t('common.word-art')} title={t('common.word-art-big')} onClick={(e) => menu.open(e, wordArtMenu((style) => act('wordArt', style)))} />
          </Group>
          <Group label={t('common.symbols')}>
            <Button icon="formula" label={t('common.equation')} title={t('sheets.ribbon.equation-typed-linear')} onClick={() => act('equation')} />
            <Button icon="plus" label={t('common.symbol')} onClick={() => openDialog('symbol')} />
          </Group>
          <Group label={t('sheets.ribbon.names')}>
            <Button tall icon="find" label={t('sheets.ribbon.define-name')} onClick={() => openDialog('names')} />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label={t('common.drawing-tools')}>
            <Button tall icon="mouse" label={t('common.select')} pressed={!ink?.tool} title={t('sheets.ribbon.select-put-pen')} onClick={() => act('inkTool', null)} />
            <Button tall icon="lasso" label={t('common.lasso')} pressed={ink?.tool === 'lasso'} title={t('common.lasso-select-draw')} onClick={() => act('inkTool', 'lasso')} />
            <Button tall icon="eraser" label={t('common.eraser')} pressed={ink?.tool === 'eraser'} title={t('common.eraser-take-away')} onClick={() => act('inkTool', 'eraser')} />
            {(ink?.pens || []).map((p) => {
              const on = ink.tool === 'pen' && ink.penId === p.id;
              const label = { pen: t('common.pen'), pencil: t('common.pencil'), highlighter: t('common.highlighter') }[p.tool];
              return (
                <Button key={p.id} tall icon="wand" label={label} pressed={on} className={`sl-pen sl-pen-${p.tool}`} data-pen={p.id} style={{ '--pen': p.color }}
                  title={on ? t('sheets.ribbon.pen-click-again', { pen: label }) : t('sheets.ribbon.pen-draw-over', { pen: label })}
                  onClick={(e) => {
                    if (!on) { act('inkTool', p.id); return; }
                    menu.open(e, [
                      { heading: true, label: t('common.thickness') },
                      ...PEN_WIDTHS[p.tool].map((w) => ({ label: t('common.width-pt', { width: w }), icon: p.width === w ? 'check' : undefined, run: () => act('inkPen', { id: p.id, width: w }) })),
                      { heading: true, label: t('common.colour') },
                      ...PEN_COLOURS.map((c) => ({ label: c, icon: p.color === c ? 'check' : undefined, preview: <span style={{ display: 'inline-block', width: 16, height: 16, borderRadius: 8, background: c, border: '1px solid rgba(0,0,0,.25)' }} />, run: () => act('inkPen', { id: p.id, color: c }) })),
                    ]);
                  }} />
              );
            })}
            <Button dropdown tall icon="plus" label={t('common.add')} title={t('common.add-pen-another')} onClick={(e) => menu.open(e, [['pen', t('common.pen')], ['highlighter', t('common.highlighter')]].map(([tool, label]) => ({ label, icon: 'plus', run: () => act('inkAdd', tool) })))} />
          </Group>
          <Group label={t('common.convert')}>
            <Button tall icon="shape" label={t('common.ink-to-shape')} pressed={Boolean(ink?.toShape)} title={t('common.ink-shape-rectangle')} onClick={() => act('inkToShape')} />
          </Group>
          <Group label={t('common.shapes')}>
            <Button dropdown tall icon="shape" label={t('common.shapes')} title={t('sheets.ribbon.shapes-rectangle-oval')} onClick={(e) => menu.open(e, SHAPES.map(([geometry, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', geometry, text: '' }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Page Layout ──────────────────────────────────────────────────── */}
      {tab === 'layout' ? (
        <>
          <Group label={t('common.themes')} icon="theme">
            <Button tall icon="wand" label={t('common.themes')} pressed={view.gallery === 'themes'} title={t('sheets.ribbon.themes-suites-themes', { theme: model?.design?.name || 'Office Theme' })} onClick={(e) => openGallery(e, 'themes')} />
            <Button icon="contrast" label={t('common.colours')} title={t('sheets.ribbon.colours-themes-twelve', { name: model?.design?.colorName || 'Office' })} onClick={(e) => openGallery(e, 'colours')} />
            <Button icon="textbox" label={t('common.fonts')} title={t('sheets.ribbon.fonts-heading-body', model?.design?.fonts ? { major: model.design.fonts.major, minor: model.design.fonts.minor } : { major: 'Calibri Light', minor: 'Calibri' })} onClick={(e) => openGallery(e, 'fonts')} />
            <Button icon="wand" label={t('common.effects')} title={t('sheets.ribbon.effects-how-shapes', { name: model?.design?.effectName || 'Office' })} onClick={(e) => openGallery(e, 'effects')} />
          </Group>
          <Group label={t('common.page-setup')}>
            <Button dropdown tall icon="file" label={t('common.margins')} title={{ normal: t('sheets.ribbon.margins-now-normal'), narrow: t('sheets.ribbon.margins-now-narrow'), wide: t('sheets.ribbon.margins-now') }[marginsName(view.page?.margins)] || (view.page?.margins ? t('sheets.ribbon.margins-now-top', { top: view.page.margins.top, left: view.page.margins.left }) : t('sheets.ribbon.margins-now-normal'))} onClick={(e) => menu.open(e, [['normal', t('common.normal')], ['narrow', t('common.narrow')], ['wide', t('common.wide')]].map(([m, label]) => ({ label, icon: marginsName(view.page?.margins) === m ? 'check' : undefined, run: () => act('page', { margins: m }) })))} />
            <Button dropdown tall icon="rotate" label={t('common.orientation')} title={view.page?.orientation === 'landscape' ? t('common.orientation-now-landscape') : t('common.orientation-now')} onClick={(e) => menu.open(e, [['portrait', t('common.portrait')], ['landscape', t('common.landscape')]].map(([o, label]) => ({ label, icon: (view.page?.orientation || 'portrait') === o ? 'check' : undefined, run: () => act('page', { orientation: o }) })))} />
            <Button dropdown tall icon="file" label={t('common.size')} title={t('sheets.ribbon.size-now-paper', { paper: view.page?.paper || 'A4' })} onClick={(e) => menu.open(e, ['A4', 'Letter', 'Legal', 'A3'].map((s) => ({ label: s, icon: (view.page?.paper || 'A4') === s ? 'check' : undefined, run: () => act('page', { size: s }) })))} />
            <Button dropdown icon="grid" label={t('sheets.ribbon.print-area')} title={t('sheets.ribbon.print-area-selection')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.set-print-area'), icon: 'grid', run: () => act('printArea', 'set') },
              { label: t('sheets.ribbon.clear-print-area'), run: () => act('printArea', 'clear') },
            ])} />
            <Button dropdown icon="minus" label={t('common.breaks')} title={t('sheets.ribbon.breaks-breaks', { breaks: breaksSaid(view.page) })} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.insert-page-break'), icon: 'minus', run: () => act('page', { breaks: 'insert' }) },
              { label: t('sheets.ribbon.remove-page-break'), run: () => act('page', { breaks: 'remove' }) },
              { label: t('sheets.ribbon.reset-all'), run: () => act('page', { breaks: 'reset' }) },
            ])} />
            <Button icon="picture" label={model?.background ? t('sheets.ribbon.delete') : t('common.background')}
              title={model?.background ? t('sheets.ribbon.delete-background-take') : t('sheets.ribbon.background-picture-tiled')}
              onClick={() => act('background')} />
            <Button dropdown icon="table" label={t('sheets.ribbon.print-titles')} title={view.page?.repeatRows ? t('sheets.ribbon.print-titles-rows', { rows: view.page.repeatRows }) : t('sheets.ribbon.print-titles-no')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.repeat-row-1'), icon: view.page?.repeatRows === 1 ? 'check' : undefined, run: () => act('page', { repeatRows: 1 }) },
              { label: t('sheets.ribbon.repeat-rows-1'), icon: view.page?.repeatRows === 2 ? 'check' : undefined, run: () => act('page', { repeatRows: 2 }) },
              { label: t('sheets.ribbon.repeat-rows'), icon: view.page?.repeatRows === 3 ? 'check' : undefined, run: () => act('page', { repeatRows: 3 }) },
              { label: t('sheets.ribbon.repeat-selected'), run: () => act('page', { repeatRows: 'selection' }) },
              { label: t('sheets.ribbon.no-repeated'), icon: !view.page?.repeatRows ? 'check' : undefined, run: () => act('page', { repeatRows: 0 }) },
            ])} />
          </Group>
          <Group label={t('sheets.ribbon.scale-to-fit')}>
            <Button dropdown icon="minus" label={t('common.width')} title={view.page?.fit === 'width' || view.page?.fit === 'page' ? t('sheets.ribbon.width-all-columns') : t('sheets.ribbon.width')} onClick={(e) => menu.open(e, [
              { label: t('common.automatic'), icon: view.page?.fit === 'width' || view.page?.fit === 'page' ? undefined : 'check', run: () => act('page', { fit: view.page?.fit === 'page' ? 'height' : 'none' }) },
              { label: t('sheets.ribbon.1-page'), icon: view.page?.fit === 'width' || view.page?.fit === 'page' ? 'check' : undefined, run: () => act('page', { fit: view.page?.fit === 'height' ? 'page' : 'width' }) },
            ])} />
            <Button dropdown icon="minus" label={t('common.height')} title={view.page?.fit === 'height' || view.page?.fit === 'page' ? t('sheets.ribbon.height-all-rows') : t('sheets.ribbon.height')} onClick={(e) => menu.open(e, [
              { label: t('common.automatic'), icon: view.page?.fit === 'height' || view.page?.fit === 'page' ? undefined : 'check', run: () => act('page', { fit: view.page?.fit === 'page' ? 'width' : 'none' }) },
              { label: t('sheets.ribbon.1-page'), icon: view.page?.fit === 'height' || view.page?.fit === 'page' ? 'check' : undefined, run: () => act('page', { fit: view.page?.fit === 'width' ? 'page' : 'height' }) },
            ])} />
            <Button dropdown icon="zoomIn" label={t('common.scale')} title={view.page?.fit && view.page.fit !== 'none' ? t('sheets.ribbon.scale-now-percent-set', { percent: Math.round((view.page?.scale ?? 1) * 100) }) : t('sheets.ribbon.scale-now-percent', { percent: Math.round((view.page?.scale ?? 1) * 100) })} onClick={(e) => menu.open(e, [50, 75, 100, 125, 150].map((p) => ({ label: `${p}%`, icon: Math.round((view.page?.scale ?? 1) * 100) === p && (view.page?.fit ?? 'none') === 'none' ? 'check' : undefined, run: () => act('page', { scale: p / 100 }) })))} />
          </Group>
          <Group label={t('sheets.ribbon.sheet')}>
            <Button icon="grid" label={t('common.gridlines')} pressed={view.gridlines !== false} title={t('sheets.ribbon.show-gridlines-screen')} onClick={() => act('toggleGridlines')} />
            <Button icon="list" label={t('sheets.ribbon.headings')} pressed={view.headings !== false} title={t('sheets.ribbon.show-row-column')} onClick={() => act('toggleHeadings')} />
            <Button icon="textRtl" label={t('sheets.ribbon.sheet-right')} pressed={Boolean(model?.rtl)} title={t('sheets.ribbon.sheet-right-left')} onClick={() => dispatch({ op: 'setRightToLeft', on: !model?.rtl })} />
          </Group>
          <Group label={t('common.arrange')}>
            <Button dropdown icon="chevronUp" label={t('common.bring')} disabled={!nPicked} title={need ? t('sheets.ribbon.bring-forward', { need }) : t('sheets.ribbon.bring-forward-picked')} onClick={(e) => menu.open(e, [
              { label: t('common.bring'), icon: 'chevronUp', run: () => act('arrange', { op: 'order', to: 'forward' }) },
              { label: t('common.bring-to'), run: () => act('arrange', { op: 'order', to: 'front' }) },
            ])} />
            <Button icon="chevronDown" label={t('common.send-backward')} disabled={!nPicked} title={need ? t('sheets.ribbon.send-backward', { need }) : t('sheets.ribbon.send-backward-picked')} onClick={(e) => menu.open(e, [
              { label: t('common.send-backward'), icon: 'chevronDown', run: () => act('arrange', { op: 'order', to: 'backward' }) },
              { label: t('common.send-to-back'), run: () => act('arrange', { op: 'order', to: 'back' }) },
            ])} />
            <Button icon="list" label={t('common.selection-pane')} pressed={arrange.pane} title={t('sheets.ribbon.selection-pane-every')} onClick={() => act('arrange', { op: 'pane' })} />
            <Button dropdown icon="alignLeft" label={t('common.align')} disabled={nPicked < 2} title={nPicked < 2 ? t('sheets.ribbon.align-select-two') : t('sheets.ribbon.align-line-picked')} onClick={(e) => menu.open(e, [
              ...[[t('common.align-left'), 'left'], [t('common.align-center'), 'center'], [t('common.align-right'), 'right'], [t('common.align-top'), 'top'], [t('common.align-middle'), 'middle'], [t('common.align-bottom'), 'bottom']]
                .map(([label, edge]) => ({ label, run: () => act('arrange', { op: 'align', edge }) })),
              { label: t('common.distribute-horizontally'), disabled: nPicked < 3, title: nPicked < 3 ? t('sheets.ribbon.select-three-more') : undefined, run: () => act('arrange', { op: 'distribute', axis: 'horizontal' }) },
              { label: t('common.distribute'), disabled: nPicked < 3, title: nPicked < 3 ? t('sheets.ribbon.select-three-more') : undefined, run: () => act('arrange', { op: 'distribute', axis: 'vertical' }) },
            ])} />
            <Button dropdown icon="grid" label={t('common.group')} disabled={!nPicked} title={need ? t('sheets.ribbon.group-need', { need }) : t('sheets.ribbon.group-gather-picked')} onClick={(e) => menu.open(e, [
              { label: t('common.group'), icon: 'grid', disabled: nPicked < 2, title: nPicked < 2 ? t('sheets.ribbon.select-two-more') : undefined, run: () => act('arrange', { op: 'group' }) },
              { label: t('common.ungroup'), disabled: !groupPicked, title: !groupPicked ? t('common.select-group') : undefined, run: () => act('arrange', { op: 'ungroup' }) },
            ])} />
            <Button dropdown icon="rotate" label={t('common.rotate')} disabled={!turnable} title={nPicked ? (turnable ? t('sheets.ribbon.rotate-turn-flip') : t('sheets.ribbon.rotate-charts')) : t('sheets.ribbon.rotate-need', { need })} onClick={(e) => menu.open(e, [
              { label: t('common.rotate-right-90'), icon: 'rotate', run: () => act('arrange', { op: 'rotate', by: 90 }) },
              { label: t('common.rotate-left-90'), run: () => act('arrange', { op: 'rotate', by: -90 }) },
              { label: t('common.flip'), icon: 'flip', run: () => act('arrange', { op: 'rotate', flip: 'vertical' }) },
              { label: t('common.flip-horizontal'), run: () => act('arrange', { op: 'rotate', flip: 'horizontal' }) },
              { label: t('common.more-rotation'), run: () => act('arrange', { op: 'rotationDialog' }) },
            ])} />
          </Group>
        </>
      ) : null}

      {/* ── Formulas ─────────────────────────────────────────────────────── */}
      {tab === 'formulas' ? (
        <>
          <Group label={t('sheets.ribbon.function')}>
            <Button tall icon="formula" label={t('sheets.ribbon.insert')} onClick={() => openDialog('function')} />
            <Button dropdown tall icon="sum" label={t('sheets.auto-sum')} onClick={(e) =>
              menu.open(e, ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((fn) => ({ label: AUTOSUM_NAMES[fn], icon: 'sum', run: () => dispatch({ op: 'autoSum', fn }) })))
            } />
            {Object.entries(FUNCTIONS).map(([category, names]) => (
              <Button key={category} dropdown icon={CATEGORY_ICONS[category] || 'formula'} label={category} onClick={(e) => functionMenu(e, names)} />
            ))}
          </Group>
          <Group label={t('sheets.ribbon.defined')}>
            <Button tall icon="find" label={t('sheets.ribbon.name-manager')} onClick={() => openDialog('names')} />
            <Button icon="plus" label={t('sheets.ribbon.define-name')} onClick={() => openDialog('names')} />
            <Button dropdown icon="formula" label={t('sheets.ribbon.use-in')} onClick={(e) => menu.open(e, (model?.names || []).length
              ? model.names.map((n) => ({ label: n.name, run: () => act('insertFunction', n.name, { bare: true }) }))
              : [{ label: t('sheets.ribbon.no-names-defined'), run: () => openDialog('names') }])} />
            <Button icon="table" label={t('sheets.ribbon.create-from')} title={t('sheets.ribbon.create-selection-name')} onClick={() => act('namesFromSelection')} />
          </Group>
          <Group label={t('sheets.ribbon.formula')}>
            <Button icon="chevronRight" label={t('sheets.ribbon.trace-precedents')} title={t('sheets.ribbon.trace-precedents-arrows')} onClick={() => act('trace', 'precedents')} />
            <Button icon="chevronLeft" label={t('sheets.ribbon.trace')} title={t('sheets.ribbon.trace-dependents-arrows')} onClick={() => act('trace', 'dependents')} />
            <Button icon="close" label={t('sheets.ribbon.remove')} title={t('sheets.ribbon.remove-arrows-take')} onClick={() => act('removeArrows')} />
            <Button icon="formula" label={t('sheets.ribbon.show')} pressed={Boolean(view.formulas)} title={t('sheets.ribbon.show-every-formula')} onClick={() => act('toggleFormulas')} />
            <Button icon="check" label={t('sheets.ribbon.error')} pressed={Boolean(model?.errors)} title={t('sheets.ribbon.error-checking-every')} onClick={() => act('errorCheck')} />
            <Button icon="eye" label={t('sheets.evaluate-formula')} title={t('sheets.ribbon.evaluate-formula-active')} onClick={() => act('evaluateFormula')} />
            <Button icon="eye" label={t('sheets.ribbon.watch-window')} pressed={Boolean(model?.watches)} title={t('sheets.ribbon.watch-window-list')} onClick={() => act('watchOpen')} />
          </Group>
          <Group label={t('sheets.ribbon.calculation')}>
            <Button dropdown tall icon="settings" label={t('sheets.ribbon.calculation-options')} title={{ auto: t('sheets.ribbon.calculation-options-now-automatic'), autoNoTable: t('sheets.ribbon.calculation-options-now-automatic-except'), manual: t('sheets.ribbon.calculation-options-now') }[model?.calc?.mode || 'auto']} onClick={(e) => menu.open(e, [
              ['auto', t('common.automatic')], ['autoNoTable', t('sheets.ribbon.automatic-except-data')], ['manual', t('common.manual')],
            ].map(([mode, label]) => ({ label, icon: (model?.calc?.mode || 'auto') === mode ? 'check' : undefined, run: () => act('calcMode', mode) })))} />
            <Rows>
              <Button icon="refresh" label={t('sheets.ribbon.calculate')} title={t('sheets.ribbon.calculate-now-every')} onClick={() => act('calculate', 'workbook')} />
              <Button icon="refresh" label={t('sheets.ribbon.calculate-sheet')} title={t('sheets.ribbon.calculate-sheet-sheets')} onClick={() => act('calculate', 'sheet')} />
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Data ─────────────────────────────────────────────────────────── */}
      {tab === 'data' ? (
        <>
          <Group label={t('sheets.ribbon.get')}>
            <Button dropdown tall icon="import" label={t('sheets.ribbon.from-text')} title={t('sheets.ribbon.text-csv-csv')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.workbook-query'), icon: 'table', run: () => act('queryFromCsv') },
              { label: t('sheets.ribbon.open-workbook'), icon: 'open', run: openFile },
            ])} />
            <Button dropdown tall icon="scanTable" label={t('sheets.ribbon.from-picture')} title={t('sheets.ribbon.picture-table-picture')} data-role="data-from-picture" onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.picture-from-file'), icon: 'picture', run: () => act('dataFromPicture', 'file') },
              { label: t('sheets.ribbon.picture-from'), icon: 'paste', run: () => act('dataFromPicture', 'clipboard') },
              '-',
              { label: t('sheets.ribbon.receipt-or'), icon: 'scanReceipt', title: t('sheets.ribbon.receipt-invoice-date'), run: () => act('receiptRow') },
            ])} />
            <Button icon="table" label={t('sheets.ribbon.from-table')} title={t('sheets.ribbon.table-range-query')} onClick={() => act('queryFromRange')} />
            <Button icon="clock" label={t('sheets.ribbon.recent')} title={t('sheets.ribbon.recent-sources-file')} onClick={(e) => {
              const list = recentSources();
              menu.open(e, list.length ? list.map((r) => ({ label: r.sourceText, title: r.source.path, icon: 'file', run: () => act('recentSource', r) })) : [{ label: t('sheets.ribbon.no-recent-sources'), disabled: true }]);
            }} />
          </Group>
          <Group label={t('sheets.queries-and')}>
            <Button tall icon="refresh" label={t('sheets.refresh-all')} title={t('sheets.ribbon.refresh-all-every')} onClick={() => act('refreshAll')} />
            <Button icon="list" label={t('sheets.queries-and')} title={t('sheets.ribbon.queries-connections')} onClick={() => act('queriesPane')} />
          </Group>
          <Group label={t('sheets.ribbon.sort-and')}>
            <Button tall icon="sort" label={t('sheets.ribbon.a-z')} title={t('sheets.ribbon.sort-list-z', { name: letterOf(sel?.active?.col ?? 0) })} onClick={() => commands['sheet.sortAsc'].run()} />
            <Button tall icon="sort" label={t('sheets.ribbon.z-a')} title={t('sheets.ribbon.sort-list-z-column', { name: letterOf(sel?.active?.col ?? 0) })} onClick={() => commands['sheet.sortDesc'].run()} />
            <Button tall icon="sort" label={t('common.sort')} title={t('sheets.ribbon.sort-three-columns')} onClick={() => act('sortDialog')} />
            <Button tall icon="filter" label={t('common.filter')} pressed={model?.filtered} onClick={() => dispatch({ op: 'autoFilter' })} />
            <Button icon="close" label={t('common.clear')} title={t('sheets.ribbon.clear-show-every')} onClick={() => act('clearFilter')} />
            <Button icon="filter" label={t('sheets.ribbon.advanced')} title={t('sheets.ribbon.advanced-filter-list')} onClick={() => act('advancedDialog')} />
          </Group>
          <Group label={t('sheets.ribbon.data-tools')}>
            <Button dropdown tall icon="table" label={t('sheets.ribbon.text-to')} title={t('sheets.ribbon.text-columns-split')} onClick={(e) => menu.open(e, [[t('sheets.comma'), 'comma'], [t('sheets.ribbon.tab'), 'tab'], [t('sheets.ribbon.semicolon'), 'semicolon'], [t('common.space'), 'space']].map(([label, delimiter]) => ({ label, run: () => act('textToColumns', delimiter) })))} />
            <Button icon="wand" label={t('sheets.ribbon.flash-fill')} title={t('sheets.ribbon.flash-fill-fill')} onClick={() => act('flashFill')} />
            <Button icon="minus" label={t('sheets.remove')} title={t('sheets.ribbon.remove-duplicates-rows')} onClick={() => act('removeDuplicates')} />
            <Button icon="check" label={t('sheets.ribbon.data')} onClick={() => openDialog('validation')} />
            <Button icon="sum" label={t('sheets.consolidate')} title={t('sheets.ribbon.consolidate-ranges-sheet')} onClick={() => act('consolidateDialog')} />
          </Group>
          <Group label={t('sheets.ribbon.forecast')}>
            <Button dropdown tall icon="wand" label={t('sheets.ribbon.what-if')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.goal-seek'), run: () => openDialog('goalSeek') },
              { label: t('sheets.ribbon.data-table'), run: () => openDialog('dataTable') },
            ])} />
            <Button tall icon="chart" label={t('sheets.ribbon.forecast-sheet')} title={t('sheets.ribbon.forecast-sheet-new')} onClick={() => act('forecastDialog')} />
          </Group>
          <Group label={t('common.outline')}>
            <Button tall icon="plus" label={t('common.group')} title={t('sheets.ribbon.group-selected-rows')} onClick={() => act('group')} />
            <Button dropdown tall icon="minus" label={t('common.ungroup')} title={t('sheets.ribbon.ungroup-selected-rows')} onClick={(e) => menu.open(e, [
              { label: t('sheets.ribbon.ungroup'), run: () => act('ungroup') },
              { label: t('sheets.ribbon.clear-outline'), run: () => act('clearOutline') },
            ])} />
            <Button tall icon="sum" label={t('common.subtotal')} title={t('sheets.ribbon.subtotal-total-row')} onClick={() => act('subtotalDialog')} />
            <Rows>
              <Button icon="plus" label={t('sheets.ribbon.show-detail')} title={t('sheets.ribbon.show-detail-open')} onClick={() => act('showDetail')} />
              <Button icon="minus" label={t('sheets.ribbon.hide-detail')} title={t('sheets.ribbon.hide-detail-fold')} onClick={() => act('hideDetail')} />
              <Button icon="grid" label={t('sheets.ribbon.auto-outline')} title={t('sheets.ribbon.auto-outline-group')} onClick={() => act('autoOutline')} />
            </Rows>
          </Group>
          <Group label={t('common.export')}>
            <Button icon="export" label="CSV" onClick={() => exportAs('csv')} />{/* words-ok: a file format's name */}
            <Button icon="export" label="TSV" onClick={() => exportAs('tsv')} />{/* words-ok: a file format's name */}
            <Button icon="pdf" label="PDF" onClick={() => exportAs('pdf')} />{/* words-ok: a file format's name */}
          </Group>
        </>
      ) : null}

      {/* ── Review ───────────────────────────────────────────────────────── */}
      {tab === 'review' ? (
        <>
          <Group label={t('common.proofing')}>
            <Button tall icon="check" label={t('common.spelling')} pressed={review?.pane === 'editor'} title={t('sheets.ribbon.spelling-f7-check')} onClick={() => review?.startSpelling()} />
            <Button icon="find" label={t('common.thesaurus')} title={t('sheets.ribbon.thesaurus-shift-f7')} onClick={() => review?.openThesaurus()} />
            <Button icon="info" label={t('sheets.ribbon.workbook-statistics')} onClick={() => openDialog('statistics')} />
          </Group>
          <Group label={t('common.accessibility')}>
            <Button tall icon="shield" label={t('common.check-accessibility')} pressed={review?.pane === 'accessibility'} title={t('sheets.ribbon.check-accessibility-alt')} onClick={() => review?.openAccessibility()} />
          </Group>
          <Group label={t('common.comments')}>
            <Button tall icon="comment" label={t('common.new-comment')} title={model?.thread ? t('sheets.ribbon.new-comment-reply') : t('sheets.ribbon.new-comment-conversation')} onClick={() => act('newComment')} />
            <Rows>
              <Button icon="close" label={t('common.delete')} title={model?.thread ? t('sheets.ribbon.delete-comment-thread') : t('sheets.ribbon.delete-cell-no')} disabled={!model?.thread} onClick={() => act('deleteThread')} />
              <Button icon="chevronLeft" label={t('common.previous')} title={t('sheets.ribbon.previous-comment-before')} onClick={() => act('stepComment', 'prev')} />
              <Button icon="chevronRight" label={t('common.next')} title={t('sheets.ribbon.next-comment-after')} onClick={() => act('stepComment', 'next')} />
            </Rows>
            <Button tall icon="eye" label={t('common.show-comments')} pressed={Boolean(model?.comments)} title={t('sheets.ribbon.show-comments-every')} onClick={() => act('commentsOpen')} />
          </Group>
          <Group label={t('common.protect')}>
            <Button tall icon="lock" label={protectedSheet ? t('sheets.unprotect') : t('sheets.protect-sheet')} pressed={protectedSheet}
              title={protectedSheet
                ? (model?.protection?.hasPassword ? t('sheets.ribbon.unprotect-sheet-locked-cells') : t('sheets.ribbon.unprotect-sheet-locked'))
                : t('sheets.ribbon.protect-sheet-locked')}
              onClick={() => act('protectSheet')} />
            <Button tall icon="lock" label={protectedBook ? t('sheets.unprotect-workbook') : t('common.protect-workbook')} pressed={protectedBook}
              title={protectedBook
                ? (model?.workbookProtection?.hasPassword ? t('sheets.ribbon.unprotect-workbook-sheets') : t('sheets.ribbon.unprotect-workbook'))
                : t('sheets.ribbon.protect-workbook-no')}
              onClick={() => act('protectWorkbook')} />
            <Button icon="lock" label={format.locked === false ? t('sheets.ribbon.unlocked') : t('sheets.ribbon.locked')} title={t('sheets.ribbon.whether-cells-locked')} pressed={format.locked !== false} onClick={() => setFormat({ locked: format.locked === false })} />
            <Button icon="lock" label={t('sheets.ribbon.allow-edit')} title={(model?.editRanges || []).length ? t('sheets.ribbon.allow-edit-ranges-ranges', { count: model.editRanges.length }) : t('sheets.ribbon.allow-edit-ranges')} onClick={() => act('editRanges')} />
            <Button tall icon="shield" label={t('common.redact')} pressed={Boolean(redaction?.open)} title={t('sheets.ribbon.redact-mark-words')} onClick={() => redaction?.toggle()} />
          </Group>
          <Group label={t('common.rules')}>
            <Button tall icon="wand" label={t('sheets.ribbon.conditional')} onClick={() => openDialog('conditional')} />
            <Button tall icon="check" label={t('sheets.ribbon.validation')} onClick={() => openDialog('validation')} />
          </Group>
          <Group label={t('common.ink')}>
            <Button icon="eye" label={t('common.hide-ink')} pressed={Boolean(ink?.hide)} title={t('sheets.ribbon.hide-ink-sheets')} onClick={() => act('hideInk')} />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
      {tab === 'view' ? (
        <>
          <Group label={t('sheets.ribbon.workbook')}>
            <Button tall icon="grid" label={t('common.normal')} title={t('sheets.ribbon.normal-sheet-grid')} pressed={(model?.viewMode || 'normal') === 'normal'} onClick={() => act('view', 'normal')} />
            <Button tall icon="file" label={t('sheets.page-break')} title={t('sheets.ribbon.page-break-preview')} pressed={model?.viewMode === 'pageBreakPreview'} onClick={() => act('view', 'pageBreakPreview')} />
            <Button tall icon="file" label={t('sheets.page-layout')} title={t('sheets.ribbon.page-layout-sheet')} pressed={model?.viewMode === 'pageLayout'} onClick={() => act('view', 'pageLayout')} />
            <Button icon="list" label={t('sheets.custom-views')} disabled={Boolean(model?.customViewsBlocked)}
              title={model?.customViewsBlocked ? t('sheets.ribbon.custom-views-why', { why: model.customViewsBlocked }) : (model?.customViews || []).length ? t('sheets.ribbon.custom-views-keep-way', { count: model.customViews.length }) : t('sheets.ribbon.custom-views-keep')}
              onClick={() => act('customViews')} />
          </Group>
          <Group label={t('common.show')}>
            <Button icon="grid" label={t('common.gridlines')} pressed={view.gridlines !== false} onClick={() => act('toggleGridlines')} />
            <Button icon="formula" label={t('sheets.formula-bar')} pressed={view.formulaBar !== false} onClick={() => act('toggleFormulaBar')} />
            <Button icon="list" label={t('sheets.ribbon.headings')} pressed={view.headings !== false} onClick={() => act('toggleHeadings')} />
            <Button icon="minus" label={t('common.ruler')} pressed={model?.viewMode === 'pageLayout' && model?.showRuler !== false} disabled={model?.viewMode !== 'pageLayout'}
              title={model?.viewMode === 'pageLayout' ? t('sheets.ribbon.ruler-centimetres-along') : t('sheets.ribbon.ruler-shown-page')}
              onClick={() => act('toggleRuler')} />
          </Group>
          <Group label={t('common.zoom')}>
            <Button dropdown tall icon="zoomIn" label={t('common.zoom')} onClick={(e) => menu.open(e, [50, 75, 100, 125, 150, 200].map((z) => ({ label: `${z}%`, run: () => act('zoom', z / 100) })))} />
            <Button icon="check" label="100%" onClick={() => act('zoom', 1)} />
            <Button icon="zoomIn" title={t('common.zoom-in-2')} onClick={() => act('zoom', (view.zoom ?? 1) + 0.1)} />
            <Button icon="zoomOut" title={t('common.zoom-out-2')} onClick={() => act('zoom', (view.zoom ?? 1) - 0.1)} />
          </Group>
          <Group label={t('common.window')}>
            <Button tall icon="new" label={t('common.new-window')} title={t('sheets.ribbon.workbook-second-window')} onClick={() => act('newWindow')} />
            <Button dropdown tall icon="freeze" label={t('sheets.ribbon.freeze-panes')} pressed={isFrozen} onClick={(e) => menu.open(e, [
              { label: isFrozen ? t('sheets.ribbon.unfreeze-panes') : t('sheets.ribbon.freeze-panes-selection'), icon: 'freeze', run: () => act('freeze', isFrozen ? 'none' : 'here') },
              { label: t('sheets.ribbon.freeze-top-row'), run: () => act('freeze', 'row') },
              { label: t('sheets.ribbon.freeze-first'), run: () => act('freeze', 'col') },
              { label: t('common.choose'), run: () => openDialog('freeze') },
            ])} />
            <Button icon="grid" label={t('common.arrange-all')} title={t('sheets.ribbon.arrange-all-every')} onClick={() => act('arrangeWindows', 'tile')} />
            <Button icon="minus" label={t('common.split')} title={t('sheets.ribbon.split-window-four')} pressed={Boolean(model?.split)} onClick={() => act('split')} />
            <Button icon="eye" label={t('common.hide')} title={t('sheets.ribbon.hide-window-put')} onClick={() => act('hideWindow')} />
            <Button icon="eye" label={t('sheets.ribbon.unhide')} title={t('sheets.ribbon.unhide-hidden-worksheets')} onClick={(e) => act('unhideMenu', e)} />
            <Button icon="maximize" label={t('common.full-screen')} onClick={() => shell.win.fullscreen({})} />
          </Group>
        </>
      ) : null}

      {/* ── Automate ─────────────────────────────────────────────────────── */}
      {tab === 'automate' ? (
        <>
          <Group label={t('sheets.ribbon.scripting')}>
            <Button tall icon="textbox" label={t('sheets.new-script')} className="sh-new-script" title={t('sheets.ribbon.new-script-script')} onClick={() => act('scripts', 'new')} />
            <Button tall icon="video" label={t('sheets.ribbon.record')} className="sh-record-actions" pressed={Boolean(view.recordingScript)} title={view.recordingScript ? t('sheets.ribbon.record-actions-recording') : t('sheets.record-actions-type')} onClick={() => act('scriptRecord')} />
            <Button icon="list" label={t('sheets.all-scripts')} className="sh-all-scripts" title={t('sheets.ribbon.all-scripts-scripts')} onClick={() => act('scripts')} />
          </Group>
        </>
      ) : null}

      {/* ── Help ─────────────────────────────────────────────────────────── */}
      {tab === 'help' ? (
        <>
          <Group label={t('common.help-and')}>
            <Button tall icon="info" label={t('common.help')} title={t('sheets.ribbon.worksheets-guide-office')} onClick={() => act('help')} />
            <Button tall icon="send" label={t('common.feedback')} title={t('common.tell-wrong-missing')} onClick={() => act('feedback')} />
            <Button icon="list" label={t('common.keyboard')} onClick={() => openDialog('shortcuts')} />
            <Button icon="info" label={t('sheets.ribbon.about')} onClick={() => act('about')} />
          </Group>
          <Group label={t('common.file')}>
            <Button tall icon="new" label={t('common.new')} onClick={() => shell.win.create({ app: 'sheets' })} />
            <Button tall icon="open" label={t('common.open')} onClick={openFile} />
            <Button tall icon="save" label={t('common.save')} onClick={() => save(false)} />
          </Group>
        </>
      ) : null}
    </Ribbon>
  );
}
