// The Worksheets ribbon.
//
// Laid out the way Excel lays it out — Home, Insert, Draw, Page Layout,
// Formulas, Data, Review, View, Automate, Help — with every group Excel has,
// because a person arriving from Excel should not have to learn where
// anything is. What the engine can do is wired to it; what it cannot is
// drawn where Excel draws it, disabled, with a title that says why. A ribbon
// that hides its gaps is a ribbon nobody fixes.

import React from 'react';
import { Ribbon, Group, Button, Separator, Select } from '@rutba/office-ui';
import { NUMBER_FORMATS } from './dialogs.js';

/** The palette a toolbar offers before it offers a colour picker. */
const SWATCHES = [
  ['#000000', 'Black'], ['#444444', 'Dark grey'], ['#888888', 'Grey'], ['#ffffff', 'White'],
  ['#c00000', 'Dark red'], ['#e03131', 'Red'], ['#e08b2b', 'Orange'], ['#e0a800', 'Amber'],
  ['#0f9d58', 'Green'], ['#0d8f6f', 'Teal'], ['#2b5fd9', 'Blue'], ['#7b5cd6', 'Purple'],
];

const FILLS = [
  [null, 'No fill'], ['FFF3BF', 'Light amber'], ['D3F9D8', 'Light green'], ['D0EBFF', 'Light blue'],
  ['FFE3E3', 'Light red'], ['E9ECEF', 'Light grey'], ['FFD8A8', 'Light orange'], ['E5DBFF', 'Light purple'],
];

const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Consolas'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];

/**
 * A border delta names edges. The engine keeps the edges it is not told about,
 * writes the ones it is, and deletes the ones set to null — so "all" is four
 * thin black edges and "none" is four nulls.
 */
const THIN = { style: 'thin', colour: '#000000' };
const MEDIUM = { style: 'medium', colour: '#000000' };
const BORDERS = [
  ['All borders', { top: THIN, bottom: THIN, left: THIN, right: THIN }],
  ['Outside borders', { top: THIN, bottom: THIN, left: THIN, right: THIN }],
  ['Thick outside borders', { top: MEDIUM, bottom: MEDIUM, left: MEDIUM, right: MEDIUM }],
  ['Top border', { top: THIN }],
  ['Bottom border', { bottom: THIN }],
  ['Left border', { left: THIN }],
  ['Right border', { right: THIN }],
  ['Thick bottom border', { bottom: MEDIUM }],
  ['Double bottom border', { bottom: { style: 'double', colour: '#000000' } }],
  ['Top and bottom border', { top: THIN, bottom: THIN }],
  ['No border', { top: null, bottom: null, left: null, right: null }],
];

/** What the engine can draw, by the names it draws them under. */
const CHARTS = [
  ['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie'], ['doughnut', 'Doughnut'],
];
const SHAPES = [
  ['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Ellipse'], ['line', 'Line'],
  ['triangle', 'Triangle'], ['diamond', 'Diamond'], ['rightArrow', 'Arrow right'], ['leftArrow', 'Arrow left'],
  ['upArrow', 'Arrow up'], ['downArrow', 'Arrow down'], ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'],
  ['star5', 'Star'], ['plus', 'Plus'], ['chevron', 'Chevron'], ['parallelogram', 'Parallelogram'], ['trapezoid', 'Trapezoid'],
];

/**
 * Excel's cell styles, as the formatting they apply. Each is one `setFormat`
 * delta, so "Good" is exactly the green Excel means and "Normal" clears it.
 */
export const CELL_STYLES = [
  ['Normal', { bold: false, italic: false, fill: null, fontColour: null, border: { top: null, bottom: null, left: null, right: null }, numberFormat: 'General' }],
  ['Heading 1', { bold: true, fontSize: 15, fontColour: '#1F3864', border: { bottom: { style: 'medium', colour: '#4472C4' } } }],
  ['Heading 2', { bold: true, fontSize: 13, fontColour: '#1F3864', border: { bottom: { style: 'medium', colour: '#A9C4E9' } } }],
  ['Heading 3', { bold: true, fontSize: 11, fontColour: '#1F3864', border: { bottom: { style: 'medium', colour: '#B4C6E7' } } }],
  ['Title', { bold: true, fontSize: 18, fontColour: '#1F3864' }],
  ['Total', { bold: true, border: { top: THIN, bottom: { style: 'double', colour: '#4472C4' } } }],
  ['Good', { fill: '#C6EFCE', fontColour: '#006100' }],
  ['Bad', { fill: '#FFC7CE', fontColour: '#9C0006' }],
  ['Neutral', { fill: '#FFEB9C', fontColour: '#9C5700' }],
  ['Input', { fill: '#FFCC99', fontColour: '#3F3F76', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }],
  ['Calculation', { bold: true, fill: '#F2F2F2', fontColour: '#FA7D00', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }],
  ['Check Cell', { bold: true, fill: '#A5A5A5', fontColour: '#FFFFFF', border: { top: MEDIUM, bottom: MEDIUM, left: MEDIUM, right: MEDIUM } }],
  ['Note', { fill: '#FFFFCC', border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }],
  ['Warning Text', { fontColour: '#FF0000' }],
  ['Currency', { numberFormat: '"£"#,##0.00' }],
  ['Percent', { numberFormat: '0%' }],
  ['Comma', { numberFormat: '#,##0.00' }],
];

/**
 * The functions Excel lists by category. Picking one starts an edit in the
 * active cell with `=NAME(` typed, which is what Excel's Insert Function does
 * after its wizard: the arguments are the person's to type.
 */
export const FUNCTIONS = {
  'Recently used': ['SUM', 'AVERAGE', 'IF', 'COUNT', 'MAX', 'MIN', 'VLOOKUP', 'ROUND'],
  Financial: ['PMT', 'FV', 'PV', 'NPV', 'IRR', 'RATE', 'NPER'],
  Logical: ['IF', 'AND', 'OR', 'NOT', 'IFERROR', 'IFS', 'SWITCH', 'TRUE', 'FALSE'],
  Text: ['CONCAT', 'LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM', 'UPPER', 'LOWER', 'PROPER', 'TEXT', 'FIND', 'SUBSTITUTE'],
  'Date & Time': ['TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'DATEDIF', 'EOMONTH', 'WEEKDAY'],
  'Lookup & Reference': ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'OFFSET', 'INDIRECT', 'ROW', 'COLUMN'],
  'Math & Trig': ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'ABS', 'MOD', 'POWER', 'SQRT', 'INT'],
  Statistical: ['AVERAGE', 'AVERAGEIF', 'COUNT', 'COUNTA', 'COUNTIF', 'COUNTIFS', 'MAX', 'MIN', 'MEDIAN', 'STDEV', 'RANK'],
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

/** A control that is drawn where Excel draws it, and says why it is not live. */
const Soon = ({ icon, label, tall, why }) => (
  <Button tall={tall} icon={icon} label={label} disabled title={`${label} — not built yet. ${why}`} />
);

export default function SheetsRibbon({
  tab, setTab, model, dispatch, commands, shell, menu, save, openFile, exportAs, doc, sel, openDialog, act, view = {},
}) {
  const format = model?.format || {};
  const frozen = model?.frozen || { rows: 0, cols: 0 };
  const isFrozen = frozen.rows > 0 || frozen.cols > 0;
  const protectedSheet = Boolean(model?.protection?.sheet);
  const size = Number(format.fontSize || 11);
  const nearer = (dir) => {
    const bigger = SIZES.filter((s) => (dir > 0 ? s > size : s < size));
    return dir > 0 ? bigger[0] ?? size : bigger[bigger.length - 1] ?? size;
  };

  const setFormat = (delta) => dispatch({ op: 'setFormat', delta });
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
        { id: 'home', label: 'Home' },
        { id: 'insert', label: 'Insert' },
        { id: 'draw', label: 'Draw' },
        { id: 'layout', label: 'Page Layout' },
        { id: 'formulas', label: 'Formulas' },
        { id: 'data', label: 'Data' },
        { id: 'review', label: 'Review' },
        { id: 'view', label: 'View' },
        { id: 'automate', label: 'Automate' },
        { id: 'help', label: 'Help' },
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title="Save (Ctrl+S)" onClick={() => save(false)} />
          <Button icon="undo" title="Undo (Ctrl+Z)" disabled={!doc?.canUndo} onClick={() => commands['edit.undo'].run()} />
          <Button icon="redo" title="Redo (Ctrl+Y)" disabled={!doc?.canRedo} onClick={() => commands['edit.redo'].run()} />
          <Button icon="sum" title="AutoSum" onClick={() => commands['sheet.autoSum'].run()} />
          <Button icon="print" title="Print" onClick={() => exportAs('pdf')} />
        </>
      }
    >
      {/* ── Home ─────────────────────────────────────────────────────────── */}
      {tab === 'home' ? (
        <>
          <Group label="Clipboard">
            <Button tall icon="paste" label="Paste" onClick={async () => dispatch({ op: 'paste', text: await shell.clipboard.readText() })} />
            <Button icon="cut" label="Cut" onClick={() => commands['edit.cut']?.run?.()} />
            <Button icon="copy" label="Copy" onClick={() => commands['edit.copy'].run()} />
            <Button icon="wand" label="Format Painter" title="Copy the formatting here to the next selection" onClick={() => dispatch({ op: 'formatBrush' })} />
          </Group>

          <Group label="Font">
            <Select value={format.fontName || 'Calibri'} onChange={(e) => setFormat({ fontName: e.target.value })} style={{ width: 118 }} title="Font">
              {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
              {format.fontName && !FONTS.includes(format.fontName) ? <option value={format.fontName}>{format.fontName}</option> : null}
            </Select>
            <Select value={String(size)} onChange={(e) => setFormat({ fontSize: Number(e.target.value) })} style={{ width: 56 }} title="Font size">
              {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
              {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
            </Select>
            <Button icon="chevronUp" title="Increase font size" onClick={() => setFormat({ fontSize: nearer(1) })} />
            <Button icon="chevronDown" title="Decrease font size" onClick={() => setFormat({ fontSize: nearer(-1) })} />
            <Separator />
            <Button icon="bold" title="Bold (Ctrl+B)" pressed={format.bold} onClick={() => setFormat({ bold: 'toggle' })} />
            <Button icon="italic" title="Italic (Ctrl+I)" pressed={format.italic} onClick={() => setFormat({ italic: 'toggle' })} />
            <Button icon="underline" title="Underline (Ctrl+U)" pressed={format.underline} onClick={() => setFormat({ underline: 'toggle' })} />
            <Button icon="strike" title="Strikethrough" pressed={format.strike} onClick={() => setFormat({ strike: 'toggle' })} />
            <Separator />
            <Button icon="grid" title="Borders" onClick={(e) => menu.open(e, BORDERS.map(([label, edges]) => ({ label, run: () => setFormat({ border: edges }) })))} />
            <Button icon="wand" title="Fill colour" onClick={(e) => swatchMenu(e, 'fill', FILLS.map(([c, l]) => [c ? `#${c}` : null, l]))} />
            <Button icon="contrast" title="Font colour" onClick={(e) => swatchMenu(e, 'fontColour', SWATCHES)} />
          </Group>

          <Group label="Alignment">
            <Button icon="chevronUp" title="Top align" pressed={format.valign === 'top'} onClick={() => setFormat({ valign: 'top' })} />
            <Button icon="minus" title="Middle align" pressed={format.valign === 'center'} onClick={() => setFormat({ valign: 'center' })} />
            <Button icon="chevronDown" title="Bottom align" pressed={format.valign === 'bottom'} onClick={() => setFormat({ valign: 'bottom' })} />
            <Soon icon="rotate" label="" why="Text orientation (angled or vertical text) is a cell alignment the engine does not write yet." />
            <Button icon="listBullet" label="Wrap Text" pressed={format.wrap} onClick={() => setFormat({ wrap: !format.wrap })} />
            <Separator />
            <Button icon="alignLeft" title="Align left" pressed={format.align === 'left'} onClick={() => setFormat({ align: 'left' })} />
            <Button icon="alignCenter" title="Centre" pressed={format.align === 'center'} onClick={() => setFormat({ align: 'center' })} />
            <Button icon="alignRight" title="Align right" pressed={format.align === 'right'} onClick={() => setFormat({ align: 'right' })} />
            <Soon icon="chevronLeft" label="" why="Decrease indent is a cell alignment the engine does not write yet." />
            <Soon icon="chevronRight" label="" why="Increase indent is a cell alignment the engine does not write yet." />
            <Separator />
            <Button icon="table" label="Merge & Centre" title="Merge the selected cells into one, and centre it" onClick={() => act('mergeCentre')} />
            <Button icon="table" title="Merge cells" onClick={() => dispatch({ op: 'merge' })} />
            <Button icon="minus" title="Unmerge cells" onClick={() => dispatch({ op: 'unmerge' })} />
          </Group>

          <Group label="Number">
            <Select value={format.numberFormat || 'General'} onChange={(e) => setFormat({ numberFormat: e.target.value })} style={{ width: 148 }} title="Number format">
              {NUMBER_FORMATS.map((f) => <option key={f.label} value={f.code}>{f.label}</option>)}
              {NUMBER_FORMATS.some((f) => f.code === (format.numberFormat || 'General')) ? null : (
                <option value={format.numberFormat}>{format.numberFormat}</option>
              )}
            </Select>
            <Separator />
            {/* The mark is the button. An icon beside a currency symbol says nothing. */}
            <Button title="Currency" onClick={() => setFormat({ numberFormat: '"£"#,##0.00' })} label="£" />
            <Button title="Percent style" onClick={() => setFormat({ numberFormat: '0%' })} label="%" />
            <Button title="Comma style" onClick={() => setFormat({ numberFormat: '#,##0.00' })} label="," />
            <Button title="Increase decimal" onClick={() => setFormat({ numberFormat: withDecimals(format.numberFormat, 1) })} label=".0" />
            <Button title="Decrease decimal" onClick={() => setFormat({ numberFormat: withDecimals(format.numberFormat, -1) })} label=".00" />
          </Group>

          <Group label="Styles">
            <Button tall icon="wand" label="Conditional Formatting" onClick={() => openDialog('conditional')} />
            <Soon tall icon="table" label="Format as Table" why="A styled table needs the table part written with a style; the engine reads tables and does not yet write one." />
            <Button tall icon="grid" label="Cell Styles" title="Excel's cell styles: headings, totals, good, bad, input" onClick={(e) => menu.open(e, CELL_STYLES.map(([label, delta]) => ({ label, run: () => setFormat(delta) })))} />
          </Group>

          <Group label="Cells">
            <Button tall icon="plus" label="Insert" onClick={(e) => menu.open(e, [
              { label: 'Insert sheet rows', icon: 'plus', run: () => commands['sheet.insertRow'].run() },
              { label: 'Insert sheet columns', icon: 'plus', run: () => commands['sheet.insertCol'].run() },
            ])} />
            <Button tall icon="minus" label="Delete" onClick={(e) => menu.open(e, [
              { label: 'Delete sheet rows', icon: 'minus', run: () => commands['sheet.deleteRow'].run() },
              { label: 'Delete sheet columns', icon: 'minus', run: () => commands['sheet.deleteCol'].run() },
            ])} />
            <Button tall icon="settings" label="Format" onClick={(e) => menu.open(e, [
              { label: 'Row height…', run: () => openDialog('rowHeight') },
              { label: 'Column width…', run: () => openDialog('colWidth') },
              { label: 'AutoFit column width', run: () => act('autoFit') },
              { label: '-' },
              { label: protectedSheet ? 'Unprotect sheet' : 'Protect sheet', icon: 'lock', run: () => dispatch({ op: protectedSheet ? 'unprotect' : 'protect' }) },
              { label: format.locked === false ? 'Lock cells' : 'Unlock cells', icon: 'lock', run: () => setFormat({ locked: format.locked === false }) },
            ])} />
          </Group>

          <Group label="Editing">
            <Button icon="sum" label="AutoSum" onClick={(e) =>
              menu.open(e, ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((fn) => ({ label: fn[0] + fn.slice(1).toLowerCase(), icon: 'sum', run: () => dispatch({ op: 'autoSum', fn }) })))
            } />
            <Button icon="chevronDown" label="Fill" onClick={(e) => menu.open(e, [
              { label: 'Down (Ctrl+D)', run: () => act('fill', 'down') },
              { label: 'Right (Ctrl+R)', run: () => act('fill', 'right') },
            ])} />
            <Button icon="close" label="Clear" onClick={(e) => menu.open(e, [
              { label: 'Clear contents', run: () => dispatch({ op: 'clear' }) },
              { label: 'Clear formats', run: () => setFormat(CELL_STYLES[0][1]) },
            ])} />
            <Button icon="sort" label="Sort & Filter" onClick={(e) => menu.open(e, [
              { label: 'Sort A to Z', icon: 'sort', run: () => commands['sheet.sortAsc'].run() },
              { label: 'Sort Z to A', icon: 'sort', run: () => commands['sheet.sortDesc'].run() },
              { label: model?.filtered ? 'Remove filter' : 'Filter', icon: 'filter', run: () => dispatch({ op: 'autoFilter' }) },
            ])} />
            <Button icon="find" label="Find & Select" onClick={(e) => menu.open(e, [
              { label: 'Find…', icon: 'find', run: () => openDialog('find') },
              { label: 'Replace…', icon: 'find', run: () => openDialog('find') },
              { label: 'Go To…', run: () => openDialog('goto') },
            ])} />
          </Group>
        </>
      ) : null}

      {/* ── Insert ───────────────────────────────────────────────────────── */}
      {tab === 'insert' ? (
        <>
          <Group label="Tables">
            <Button tall icon="table" label="PivotTable" onClick={() => openDialog('pivot')} />
            <Soon tall icon="table" label="Table" why="A styled table needs the table part written with a style; the engine reads tables and does not yet write one." />
          </Group>
          <Group label="Illustrations">
            <Soon tall icon="picture" label="Pictures" why="A picture in a worksheet is a drawing part with an anchor; the engine draws them and does not yet write one." />
            <Button tall icon="shape" label="Shapes" onClick={(e) => menu.open(e, SHAPES.map(([geometry, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', geometry, text: '' }) })))} />
            <Soon icon="star" label="Icons" why="Icons are an online library." />
            <Soon icon="shape" label="SmartArt" why="SmartArt is a diagram part the engine does not write." />
          </Group>
          <Group label="Charts">
            {CHARTS.map(([kind, label]) => (
              <Button key={kind} icon="chart" label={label} title={`${label} chart from the data around the selection`} onClick={() => dispatch({ op: 'insertChart', kind })} />
            ))}
            <Soon icon="chart" label="Scatter" why="A scatter chart needs an XY series the chart writer does not build yet." />
            <Soon icon="chart" label="PivotChart" why="Comes with a pivot table that can be charted." />
          </Group>
          <Group label="Sparklines">
            <Soon icon="chart" label="Line" why="Sparklines are an extension part (x14) the engine does not write." />
            <Soon icon="chart" label="Column" why="Comes with sparklines." />
          </Group>
          <Group label="Filters">
            <Button tall icon="filter" label="Filter" pressed={model?.filtered} onClick={() => dispatch({ op: 'autoFilter' })} />
            <Soon icon="filter" label="Slicer" why="A slicer is a separate part tied to a table or pivot; not written yet." />
          </Group>
          <Group label="Links">
            <Soon tall icon="link" label="Link" why="A cell hyperlink needs the hyperlinks list and a relationship the engine does not write yet." />
          </Group>
          <Group label="Comments">
            <Soon tall icon="reply" label="Comment" why="Cell comments are a comments part plus VML the engine does not write yet." />
          </Group>
          <Group label="Text">
            <Button tall icon="textbox" label="Text Box" title="A rectangle with words in it" onClick={() => act('textBox')} />
            <Soon icon="file" label="Header & Footer" why="Page headers and footers are print settings the engine does not write yet." />
            <Soon icon="wand" label="WordArt" why="WordArt is DrawingML text effects the engine does not write." />
          </Group>
          <Group label="Symbols">
            <Soon icon="formula" label="Equation" why="Equations are OMML the engine does not write." />
            <Button icon="plus" label="Symbol" onClick={() => openDialog('symbol')} />
          </Group>
          <Group label="Names">
            <Button tall icon="find" label="Define Name" onClick={() => openDialog('names')} />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label="Drawing Tools">
            <Soon tall icon="wand" label="Select" why="Ink needs a pointer surface the grid does not have." />
            <Soon tall icon="wand" label="Lasso" why="Comes with ink." />
            <Soon tall icon="close" label="Eraser" why="Comes with ink." />
            <Soon tall icon="wand" label="Pen" why="Ink is a drawing part (ink ML) the engine does not write." />
            <Soon tall icon="wand" label="Highlighter" why="Comes with ink." />
          </Group>
          <Group label="Convert">
            <Soon tall icon="shape" label="Ink to Shape" why="Comes with ink." />
            <Soon tall icon="formula" label="Ink to Math" why="Comes with ink." />
          </Group>
          <Group label="Shapes">
            <Button tall icon="shape" label="Shapes" title="A shape, since a pen is not here yet" onClick={(e) => menu.open(e, SHAPES.map(([geometry, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', geometry, text: '' }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Page Layout ──────────────────────────────────────────────────── */}
      {tab === 'layout' ? (
        <>
          <Group label="Themes">
            <Soon tall icon="wand" label="Themes" why="A theme is the theme part; the engine reads it and does not yet swap it." />
            <Soon icon="contrast" label="Colours" why="Comes with themes." />
            <Soon icon="textbox" label="Fonts" why="Comes with themes." />
            <Soon icon="wand" label="Effects" why="Comes with themes." />
          </Group>
          <Group label="Page Setup">
            <Button tall icon="file" label="Margins" title={`Now ${view.page?.margins || 'normal'}`} onClick={(e) => menu.open(e, [
              { label: 'Normal', run: () => act('page', { margins: 'normal' }) },
              { label: 'Narrow', run: () => act('page', { margins: 'narrow' }) },
              { label: 'Wide', run: () => act('page', { margins: 'wide' }) },
            ])} />
            <Button tall icon="rotate" label="Orientation" title={`Now ${view.page?.orientation || 'portrait'}`} onClick={(e) => menu.open(e, [
              { label: 'Portrait', run: () => act('page', { orientation: 'portrait' }) },
              { label: 'Landscape', run: () => act('page', { orientation: 'landscape' }) },
            ])} />
            <Button tall icon="file" label="Size" title={`Now ${view.page?.size || 'A4'}`} onClick={(e) => menu.open(e, ['A4', 'Letter', 'Legal', 'A3'].map((s) => ({ label: s, run: () => act('page', { size: s }) })))} />
            <Soon icon="grid" label="Print Area" why="A print area is a defined name (_xlnm.Print_Area); the name manager can make one, this button will soon." />
            <Soon icon="minus" label="Breaks" why="Manual page breaks are print settings the engine does not write yet." />
            <Soon icon="picture" label="Background" why="A sheet background is a picture part the engine does not write yet." />
            <Soon icon="table" label="Print Titles" why="Print titles are a defined name (_xlnm.Print_Titles) not written yet." />
          </Group>
          <Group label="Scale to Fit">
            <Soon icon="minus" label="Width" why="Scaling is a print setting the engine does not write yet." />
            <Soon icon="minus" label="Height" why="Comes with scaling." />
            <Soon icon="zoomIn" label="Scale" why="Comes with scaling." />
          </Group>
          <Group label="Sheet Options">
            <Button icon="grid" label="Gridlines" pressed={view.gridlines !== false} title="Show the gridlines on screen" onClick={() => act('toggleGridlines')} />
            <Button icon="list" label="Headings" pressed={view.headings !== false} title="Show the row and column headings" onClick={() => act('toggleHeadings')} />
          </Group>
          <Group label="Arrange">
            <Soon icon="chevronUp" label="Bring Forward" why="Shape ordering is a drawing-part edit not written yet." />
            <Soon icon="chevronDown" label="Send Backward" why="Comes with ordering." />
            <Soon icon="list" label="Selection Pane" why="Comes with ordering." />
            <Soon icon="alignLeft" label="Align" why="Comes with ordering." />
            <Soon icon="shape" label="Group" why="Comes with ordering." />
            <Soon icon="rotate" label="Rotate" why="Comes with ordering." />
          </Group>
        </>
      ) : null}

      {/* ── Formulas ─────────────────────────────────────────────────────── */}
      {tab === 'formulas' ? (
        <>
          <Group label="Function Library">
            <Button tall icon="formula" label="Insert Function" onClick={() => openDialog('function')} />
            <Button tall icon="sum" label="AutoSum" onClick={(e) =>
              menu.open(e, ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((fn) => ({ label: fn[0] + fn.slice(1).toLowerCase(), icon: 'sum', run: () => dispatch({ op: 'autoSum', fn }) })))
            } />
            {Object.entries(FUNCTIONS).map(([category, names]) => (
              <Button key={category} icon="formula" label={category} onClick={(e) => functionMenu(e, names)} />
            ))}
          </Group>
          <Group label="Defined Names">
            <Button tall icon="find" label="Name Manager" onClick={() => openDialog('names')} />
            <Button icon="plus" label="Define Name" onClick={() => openDialog('names')} />
            <Button icon="formula" label="Use in Formula" onClick={(e) => menu.open(e, (model?.names || []).length
              ? model.names.map((n) => ({ label: n.name, run: () => act('insertFunction', n.name, { bare: true }) }))
              : [{ label: 'No names defined yet', run: () => openDialog('names') }])} />
            <Soon icon="table" label="Create from Selection" why="Naming a range from its header row is on the list; the name manager does it one at a time today." />
          </Group>
          <Group label="Formula Auditing">
            <Soon icon="chevronRight" label="Trace Precedents" why="Tracing arrows need the dependency graph drawn on the grid; the engine has the graph, the grid does not draw it yet." />
            <Soon icon="chevronLeft" label="Trace Dependents" why="Comes with tracing." />
            <Soon icon="close" label="Remove Arrows" why="Comes with tracing." />
            <Button icon="formula" label="Show Formulas" pressed={Boolean(view.formulas)} title="Show every formula instead of its result (Ctrl+`)" onClick={() => act('toggleFormulas')} />
            <Soon icon="check" label="Error Checking" why="Error checking walks the graph for #REF! and circularity; on the list." />
            <Soon icon="eye" label="Evaluate Formula" why="Step evaluation needs the engine to expose each step; on the list." />
            <Soon icon="eye" label="Watch Window" why="Comes with evaluation." />
          </Group>
          <Group label="Calculation">
            <Button tall icon="refresh" label="Calculate Now" title="The engine recalculates on every edit; this forces a full pass" onClick={() => act('recalculate')} />
            <Soon icon="settings" label="Calculation Options" why="Manual calculation is a workbook setting the engine does not honour yet — it always calculates." />
          </Group>
        </>
      ) : null}

      {/* ── Data ─────────────────────────────────────────────────────────── */}
      {tab === 'data' ? (
        <>
          <Group label="Get & Transform Data">
            <Button tall icon="import" label="From Text/CSV" title="Open a CSV or TSV file as a workbook" onClick={openFile} />
            <Soon tall icon="globe" label="From Web" why="Fetching a table from a web page is a network feature this suite does not do on its own." />
            <Soon icon="table" label="From Table/Range" why="Power Query is not here." />
            <Soon icon="clock" label="Recent Sources" why="Power Query is not here." />
          </Group>
          <Group label="Queries & Connections">
            <Button tall icon="refresh" label="Refresh All" title="Recalculate and refresh every pivot table" onClick={() => act('refreshAll')} />
            <Soon icon="list" label="Queries & Connections" why="Power Query is not here." />
          </Group>
          <Group label="Sort & Filter">
            <Button tall icon="sort" label="A → Z" title="Sort the selection ascending by its first column" onClick={() => commands['sheet.sortAsc'].run()} />
            <Button tall icon="sort" label="Z → A" title="Sort the selection descending by its first column" onClick={() => commands['sheet.sortDesc'].run()} />
            <Soon tall icon="sort" label="Sort" why="Sorting by several columns at once needs a keys list the engine does not take yet." />
            <Button tall icon="filter" label="Filter" pressed={model?.filtered} onClick={() => dispatch({ op: 'autoFilter' })} />
            <Soon icon="close" label="Clear" why="Comes with the filter dialog." />
            <Soon icon="filter" label="Advanced" why="Criteria ranges are not built." />
          </Group>
          <Group label="Data Tools">
            <Soon tall icon="table" label="Text to Columns" why="Splitting a column on a delimiter is on the list; Paste already splits tabs and commas." />
            <Soon icon="wand" label="Flash Fill" why="Pattern completion is not built." />
            <Soon icon="minus" label="Remove Duplicates" why="Needs the engine to read the whole selection, not the part on screen; on the list." />
            <Button icon="check" label="Data Validation" onClick={() => openDialog('validation')} />
            <Soon icon="sum" label="Consolidate" why="Consolidating ranges across sheets is not built." />
          </Group>
          <Group label="Forecast">
            <Button tall icon="wand" label="What-If Analysis" onClick={(e) => menu.open(e, [
              { label: 'Goal Seek…', run: () => openDialog('goalSeek') },
              { label: 'Data Table…', run: () => openDialog('dataTable') },
            ])} />
            <Soon tall icon="chart" label="Forecast Sheet" why="Forecasting (FORECAST.ETS) is not built." />
          </Group>
          <Group label="Outline">
            <Soon icon="plus" label="Group" why="Row and column outlining is a sheet property the engine does not write yet." />
            <Soon icon="minus" label="Ungroup" why="Comes with outlining." />
            <Soon icon="sum" label="Subtotal" why="Comes with outlining." />
          </Group>
          <Group label="Export">
            <Button icon="export" label="CSV" onClick={() => exportAs('csv')} />
            <Button icon="export" label="TSV" onClick={() => exportAs('tsv')} />
            <Button icon="pdf" label="PDF" onClick={() => exportAs('pdf')} />
          </Group>
        </>
      ) : null}

      {/* ── Review ───────────────────────────────────────────────────────── */}
      {tab === 'review' ? (
        <>
          <Group label="Proofing">
            <Soon tall icon="check" label="Spelling" why="The system spell checker underlines as you type in a cell; a pass over the sheet is not built." />
            <Soon icon="find" label="Thesaurus" why="A thesaurus is a dictionary this suite does not ship." />
            <Button icon="info" label="Workbook Statistics" onClick={() => openDialog('statistics')} />
          </Group>
          <Group label="Accessibility">
            <Soon tall icon="check" label="Check Accessibility" why="An accessibility pass is on the list." />
          </Group>
          <Group label="Language">
            <Soon tall icon="globe" label="Translate" why="Translation is an online service this suite does not call." />
          </Group>
          <Group label="Comments">
            <Soon tall icon="reply" label="New Comment" why="Cell comments are a comments part plus VML the engine does not write yet." />
            <Soon icon="close" label="Delete" why="Comes with comments." />
            <Soon icon="chevronLeft" label="Previous" why="Comes with comments." />
            <Soon icon="chevronRight" label="Next" why="Comes with comments." />
            <Soon icon="eye" label="Show Comments" why="Comes with comments." />
          </Group>
          <Group label="Protect">
            <Button tall icon="lock" label={protectedSheet ? 'Unprotect Sheet' : 'Protect Sheet'} pressed={protectedSheet} onClick={() => dispatch({ op: protectedSheet ? 'unprotect' : 'protect' })} />
            <Soon tall icon="lock" label="Protect Workbook" why="Workbook structure protection is a workbook setting the engine does not write yet." />
            <Button icon="lock" label={format.locked === false ? 'Unlocked' : 'Locked'} title="Whether these cells are locked when the sheet is protected" pressed={format.locked !== false} onClick={() => setFormat({ locked: format.locked === false })} />
            <Soon icon="lock" label="Allow Edit Ranges" why="Editable ranges on a protected sheet are not built." />
          </Group>
          <Group label="Rules">
            <Button tall icon="wand" label="Conditional" onClick={() => openDialog('conditional')} />
            <Button tall icon="check" label="Validation" onClick={() => openDialog('validation')} />
          </Group>
          <Group label="Ink">
            <Soon icon="eye" label="Hide Ink" why="Comes with ink." />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
      {tab === 'view' ? (
        <>
          <Group label="Workbook Views">
            <Button tall icon="grid" label="Normal" pressed onClick={() => act('view', 'normal')} />
            <Soon tall icon="file" label="Page Break Preview" why="Page breaks need the print layout the paginator does not compute for a sheet yet." />
            <Soon tall icon="file" label="Page Layout" why="Comes with page breaks." />
            <Soon icon="list" label="Custom Views" why="Saved views are a workbook part not written yet." />
          </Group>
          <Group label="Show">
            <Button icon="grid" label="Gridlines" pressed={view.gridlines !== false} onClick={() => act('toggleGridlines')} />
            <Button icon="formula" label="Formula Bar" pressed={view.formulaBar !== false} onClick={() => act('toggleFormulaBar')} />
            <Button icon="list" label="Headings" pressed={view.headings !== false} onClick={() => act('toggleHeadings')} />
          </Group>
          <Group label="Zoom">
            <Button tall icon="zoomIn" label="Zoom" onClick={(e) => menu.open(e, [50, 75, 100, 125, 150, 200].map((z) => ({ label: `${z}%`, run: () => act('zoom', z / 100) })))} />
            <Button icon="check" label="100%" onClick={() => act('zoom', 1)} />
            <Button icon="zoomIn" title="Zoom in" onClick={() => shell.win.zoom({ delta: 0.1 })} />
            <Button icon="zoomOut" title="Zoom out" onClick={() => shell.win.zoom({ delta: -0.1 })} />
          </Group>
          <Group label="Window">
            <Button tall icon="new" label="New Window" title="This workbook in a second window" onClick={() => act('newWindow')} />
            <Button tall icon="freeze" label="Freeze Panes" pressed={isFrozen} onClick={(e) => menu.open(e, [
              { label: isFrozen ? 'Unfreeze panes' : 'Freeze panes at the selection', icon: 'freeze', run: () => act('freeze', isFrozen ? 'none' : 'here') },
              { label: 'Freeze top row', run: () => act('freeze', 'row') },
              { label: 'Freeze first column', run: () => act('freeze', 'col') },
              { label: 'Choose…', run: () => openDialog('freeze') },
            ])} />
            <Soon icon="grid" label="Arrange All" why="Window tiling is the operating system's; not built." />
            <Soon icon="minus" label="Split" why="A split grid is on the list; freeze panes covers most of it." />
            <Soon icon="eye" label="Hide" why="Hiding a window is not built." />
            <Button icon="maximize" label="Full Screen" onClick={() => shell.win.fullscreen({})} />
          </Group>
          <Group label="Macros">
            <Soon tall icon="settings" label="Macros" why="VBA is preserved in the file and never run: a spreadsheet that runs code it received in an email is how ransomware starts." />
          </Group>
        </>
      ) : null}

      {/* ── Automate ─────────────────────────────────────────────────────── */}
      {tab === 'automate' ? (
        <>
          <Group label="Scripting Tools">
            <Soon tall icon="settings" label="New Script" why="Office Scripts are a cloud runtime; a local, open scripting surface is on the list." />
            <Soon tall icon="settings" label="Record Actions" why="Comes with scripts." />
            <Soon icon="list" label="All Scripts" why="Comes with scripts." />
          </Group>
          <Group label="Power Automate">
            <Soon tall icon="refresh" label="Automate a Task" why="Power Automate is a Microsoft cloud service." />
          </Group>
        </>
      ) : null}

      {/* ── Help ─────────────────────────────────────────────────────────── */}
      {tab === 'help' ? (
        <>
          <Group label="Help & Support">
            <Button tall icon="info" label="Help" title="The Worksheets guide on office.rutba.io" onClick={() => act('help')} />
            <Button tall icon="send" label="Feedback" title="Tell us what is wrong or missing" onClick={() => act('feedback')} />
            <Button icon="list" label="Keyboard Shortcuts" onClick={() => openDialog('shortcuts')} />
            <Button icon="info" label="About" onClick={() => act('about')} />
          </Group>
          <Group label="File">
            <Button tall icon="new" label="New" onClick={() => shell.win.create({ app: 'sheets' })} />
            <Button tall icon="open" label="Open" onClick={openFile} />
            <Button tall icon="save" label="Save" onClick={() => save(false)} />
          </Group>
        </>
      ) : null}
    </Ribbon>
  );
}
