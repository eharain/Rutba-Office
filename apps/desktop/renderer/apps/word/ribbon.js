// The Word ribbon, to Word's own layout.
//
// Home, Insert, Draw, Design, Layout, References, Mailings, Review, View, Help,
// PDF — every tab and every group Microsoft Word shows, in the same order, so a
// person arriving from Word finds things where their hands already go.
//
// Two rules hold this file together:
//
//   Every control speaks the engine's vocabulary. `fontName`, `fontSize`,
//   `fontColour`, `highlight`; `align`, `indentDelta`, `styleId`, `list`,
//   `lineSpacing`, `spaceBefore`, `spaceAfter`; `A4`, `Letter`, `Legal`;
//   `normal`, `narrow`, `wide`. A control that sends anything else is a button
//   that presses and does nothing, and the checks now press them all.
//
//   A control that is not wired says so. It is drawn where Word draws it,
//   disabled, and its tooltip says what it will do and why it does not yet.
//   That is the build order the owner asked for — all the ribbons first, then
//   wired one after another — and it is also more honest than hiding the gap.
//
// State comes from `model.format` (what the engine reports at the caret) and
// from the `view` the page keeps (mode, panes, marks).

import React from 'react';
import { TABLE_STYLES } from '@rutba/ooxml/table-styles';
import { WARP_PRESETS, WARP_MORE, warpLabel } from '@rutba/drawing/warp';
import { Ribbon, Group, Rows, Button, Separator, Select, Input } from '@rutba/office-ui';
import { EQUATION_GALLERY } from '@rutba/ooxml/math-linear';
import { THEMES, PALETTES, FONT_PAIRS, EFFECT_PRESETS } from '@rutba/office-formats/themes';
import { MailingsTab } from './mailings.js';
import { CitationsGroup } from './references.js';
import { IndexGroup } from './references-index.js';
import { PEN_COLOURS, PEN_WIDTHS } from '../slides/ink-geometry.js';
import { ToaGroup } from './references-toa.js';
import { captionsExtra } from './references-figures.js';
import { MODEL_VIEWS } from '@rutba/imaging/model3d';
import { untrackedOk } from './untracked.js';

/* ── vocabularies ────────────────────────────────────────────────────────── */

const FONTS = ['Calibri', 'Calibri Light', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Tahoma', 'Garamond', 'Consolas', 'Courier New'];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

// Immersive Reader's choices, as Word names them.
const IR_WIDTHS = [['veryNarrow', 'Very Narrow'], ['narrow', 'Narrow'], ['moderate', 'Moderate'], ['wide', 'Wide']];
const IR_COLOURS = [['none', 'None'], ['sepia', 'Sepia'], ['inverse', 'Inverse']];
const IR_FOCUS = [[0, 'None'], [1, 'One Line'], [3, 'Three Lines'], [5, 'Five Lines']];

const TEXT_COLOURS = [
  [null, 'Automatic'], ['000000', 'Black'], ['444444', 'Dark grey'], ['767171', 'Grey'], ['FFFFFF', 'White'],
  ['C00000', 'Dark red'], ['E03131', 'Red'], ['E08B2B', 'Orange'], ['E0A800', 'Gold'],
  ['0F9D58', 'Green'], ['0D8F6F', 'Teal'], ['2B5FD9', 'Blue'], ['1F3864', 'Dark blue'], ['7B5CD6', 'Purple'],
];
const HIGHLIGHTS = [
  [null, 'No highlight'], ['yellow', 'Yellow'], ['green', 'Bright green'], ['cyan', 'Turquoise'], ['magenta', 'Pink'],
  ['blue', 'Blue'], ['red', 'Red'], ['darkBlue', 'Dark blue'], ['darkCyan', 'Teal'], ['darkGreen', 'Green'],
  ['darkMagenta', 'Violet'], ['darkRed', 'Dark red'], ['darkYellow', 'Dark yellow'], ['darkGray', 'Grey 50%'], ['lightGray', 'Grey 25%'], ['black', 'Black'],
];
const SPACING = [[1, '1.0'], [1.15, '1.15'], [1.5, '1.5'], [2, '2.0'], [2.5, '2.5'], [3, '3.0']];
/** Word's half-point rule, for the Borders menu. */
const LINE = { style: 'single', widthPx: 1, colour: null, spacePt: 1 };
const SHADES = [['#FFF2CC', 'Light yellow'], ['#DEEBF7', 'Light blue'], ['#E2EFDA', 'Light green'], ['#FCE4D6', 'Light orange'], ['#EDEDED', 'Light grey'], [null, 'No colour']];
const BORDERS = [
  ['Bottom border', { bottom: LINE }], ['Top border', { top: LINE }], ['Left border', { left: LINE }], ['Right border', { right: LINE }],
  ['Outside borders', { top: LINE, bottom: LINE, left: LINE, right: LINE }], ['No border', null],
];
/** Page Borders: Word's Box setting in four lines, measured 24 pt in from the page edge as Word measures it. */
const FRAME = (style, widthPx) => ({ style, widthPx, colour: null, spacePt: 24 });
const BOX = (line) => ({ offsetFrom: 'page', top: line, left: line, bottom: line, right: line });
const PAGE_BORDERS = [
  ['Box — thin line', BOX(FRAME('single', 1))], ['Box — thick line', BOX(FRAME('single', 3))],
  ['Box — double line', BOX(FRAME('double', 3))], ['Box — dashed line', BOX(FRAME('dashed', 1))],
  ['No page border', null],
];
/** Watermark: Word's own stock words, then a custom one. */
const WATERMARKS = ['DRAFT', 'CONFIDENTIAL', 'SAMPLE', 'DO NOT COPY', 'URGENT'];
const PAGE_SIZES = [['A4', 'A4 — 21 × 29.7 cm'], ['Letter', 'Letter — 8.5 × 11 in'], ['Legal', 'Legal — 8.5 × 14 in']];
/** Review → Display for Review: the button's own label per mode. */
const MARKUP_LABELS = { all: 'All Markup', simple: 'Simple Markup', final: 'No Markup', original: 'Original' };
const MARGINS = [['normal', 'Normal — 2.54 cm all round'], ['narrow', 'Narrow — 1.27 cm all round'], ['wide', 'Wide — 5.08 cm at the sides']];
const SHAPES = [
  ['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Oval'], ['triangle', 'Triangle'], ['diamond', 'Diamond'],
  ['rightArrow', 'Arrow: right'], ['leftArrow', 'Arrow: left'], ['upArrow', 'Arrow: up'], ['downArrow', 'Arrow: down'],
  ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'], ['star5', 'Star: 5 points'], ['plus', 'Cross'], ['chevron', 'Chevron'],
  ['parallelogram', 'Parallelogram'], ['trapezoid', 'Trapezoid'], ['line', 'Line'],
];
const CHARTS = [['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['pie', 'Pie']];
const CASES = [['sentence', 'Sentence case.'], ['lower', 'lowercase'], ['upper', 'UPPERCASE'], ['title', 'Capitalize Each Word'], ['toggle', 'tOGGLE cASE']];
/** The "A" button's glow colours — Word's own accent palette, 4pt radius. */
const GLOWS = [['FFC000', 'gold'], ['4472C4', 'blue'], ['70AD47', 'green'], ['FF0000', 'red']];

/** Shape Fill and Shape Outline: Word's standard colours, in the suite's own names. */
const SHAPE_COLOURS = [
  ['FFFFFF', 'White'], ['000000', 'Black'], ['E7E6E6', 'Light grey'], ['767171', 'Grey'], ['1F3864', 'Dark blue'],
  ['2B5FD9', 'Blue'], ['DEEBF7', 'Light blue'], ['0D8F6F', 'Teal'], ['0F9D58', 'Green'], ['E2EFDA', 'Light green'],
  ['E0A800', 'Gold'], ['FFF2CC', 'Light yellow'], ['E08B2B', 'Orange'], ['C00000', 'Dark red'], ['7B5CD6', 'Purple'],
];
/** Shape Outline → Weight, in Word's points and the pixels the engine writes. */
const WEIGHTS = [['½ pt', 0.67], ['1 pt', 1.33], ['1½ pt', 2], ['2¼ pt', 3], ['3 pt', 4], ['6 pt', 8]];
/** A text box's Margins, Word's four: its insets in px, left/top/right/bottom. */
const BOX_MARGINS = [
  ['None', { l: 0, t: 0, r: 0, b: 0 }],
  ['Narrow', { l: 4.8, t: 4.8, r: 4.8, b: 4.8 }],
  ['Normal', { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }],
  ['Wide', { l: 14.4, t: 14.4, r: 14.4, b: 14.4 }],
];
/** Insert → WordArt: four styles of our own, made of the text effects the engine writes. */
const WORDART = [
  ['Fill: blue, shadow', { colour: '2B5FD9', effects: { shadow: true } }],
  ['Outline: blue', { colour: '2B5FD9', effects: { outline: true } }],
  ['Fill: gold, glow', { colour: 'C98A00', effects: { glow: { colour: 'FFC000', radiusPt: 4 } } }],
  ['Fill: black, shadow', { colour: '262626', effects: { shadow: true } }],
];
const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);
/** 1 cm in CSS px, for the Size boxes. */
const CM = 96 / 2.54;

/** Table Design's pen: Word's line styles, and its weights in eighths of a point. */
const PEN_STYLES = [['single', 'Single'], ['double', 'Double'], ['dotted', 'Dotted'], ['dashed', 'Dashed'], ['thick', 'Thick']];
const PEN_WEIGHTS = [[2, '¼ pt'], [4, '½ pt'], [6, '¾ pt'], [8, '1 pt'], [12, '1½ pt'], [18, '2¼ pt'], [24, '3 pt'], [36, '4½ pt'], [48, '6 pt']];

/** A control that is drawn where Word draws it, and says why it is not live. */
const Soon = ({ icon, label, tall, why }) => (
  <Button tall={tall} icon={icon} label={label} disabled title={`${label} — not built yet. ${why}`} />
);

export default function WordRibbon({
  tab, setTab, doc, model, dispatch, commands, shell, menu, save, openFile, exportAs, openDialog, insertPicture, act, view = {}, picked = null, mailings = null, review = null, drawing = null, references = null, ink = null, blocks = null, table = null,
}) {
  // Table Layout: an operation on the caret's table, and whether the selection runs across its cells.
  // A merge or a split while recording is asked about first, as Word asks (see untracked.js).
  const tableOp = async (kind, arg) => { if (await untrackedOk(shell, model, kind)) dispatch({ op: 'tableOp', kind, ...(arg ? { arg } : {}) }); };
  const innermost = (block) => /(t\d+):r(\d+):c(\d+)(?!.*:t\d+:)/.exec(String(model?.blocks?.[block]?.container || ''));
  const anchorCell = innermost(model?.selection?.anchor?.block ?? -1);
  const focusCell = innermost(model?.selection?.focus?.block ?? -1);
  const acrossCells = Boolean(anchorCell && focusCell && anchorCell[1] === focusCell[1] && (anchorCell[2] !== focusCell[2] || anchorCell[3] !== focusCell[3]));
  // The caret's cell: where its words sit, and whether its row is a header row.
  const caretCell = model?.blocks?.[model?.selection?.focus?.block ?? -1] || null;
  // A merged caret cell: across columns, or carried on in the row below it.
  const cellBelow = caretCell?.container ? caretCell.container.replace(/:r(\d+):c(\d+)$/, (m, r, c) => `:r${Number(r) + 1}:c${c}`) : null;
  const mergedCaret = Boolean(caretCell?.cellSpan > 1 || (cellBelow && model?.blocks?.some((b) => b.hiddenCell && b.container === cellBelow)));
  // Table Design: the caret's table's style and which of its parts it shows (Word's for a new table when it says nothing).
  const tableStyleId = caretCell?.tableStyle?.id || null;
  const styleLook = caretCell?.tableStyle?.look || { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true };
  const setLook = (change) => tableOp('styleOptions', { look: { ...styleLook, ...change } });
  // Table Design → Borders: the pen the borders are drawn in — Word's single half-point automatic line to begin with.
  const [pen, setPen] = React.useState({ val: 'single', sz: 4, color: 'auto' });
  const cellAlign = caretCell?.cellVAlign || 'top';
  const format = model?.format || {};
  const design = model?.design || null;
  const styles = Array.isArray(model?.styles) ? model.styles : [];
  const section = model?.section || null;
  const comments = model?.comments || [];
  // Review → Accept/Reject/Previous/Next read the caret's own paragraph and
  // whether anything in the document is tracked at all.
  const atBlock = model?.selection?.focus?.block ?? 0;
  const trackedHere = Boolean(model?.blocks?.[atBlock]?.tracked);
  const anyTracked = Boolean(model?.blocks?.some((b) => b.tracked));
  const landscape = section?.orientation === 'landscape';
  // Layout → Columns: the section's own `columns` (count 1 with no
  // `w:cols` at all), and Word's five presets in its own vocabulary — the
  // engine works out `w:col` widths for Left and Right from the content
  // width it already knows in twips.
  const columns = section?.columns || { count: 1, spacePx: 0, separator: false, widths: null };
  const colCount = columns.count || 1;
  const contentTwips = section ? Math.round(section.contentWidthPx * 15) : 9360;
  const setColumns = (spec) => dispatch({ op: 'setPageSetup', spec: { columns: spec } });
  const unevenSpec = (leftNarrow) => {
    const spaceTwips = 720;
    const usable = Math.max(2, contentTwips - spaceTwips);
    const narrow = Math.round(usable / 3);
    const wide = usable - narrow;
    return { count: 2, spaceTwips, separator: Boolean(columns.separator), widths: leftNarrow ? [narrow, wide] : [wide, narrow] };
  };
  const isColumnPreset = (name) => {
    if (name === 'one') return colCount <= 1;
    if (!columns.widths) return (name === 'two' && colCount === 2) || (name === 'three' && colCount === 3);
    if (colCount !== 2) return false;
    return name === 'left' ? columns.widths[0] < columns.widths[1] : name === 'right' && columns.widths[0] > columns.widths[1];
  };

  // Drawings: what is selected, and what Arrange may do with it.
  const pickedIds = picked?.ids || [];
  const drawings = model?.drawings || [];
  const hasDrawing = Boolean(drawing) || pickedIds.length > 0 || picked?.image != null;
  const floatingSel = drawings.filter((d) => (pickedIds.length ? pickedIds.includes(d.id) : drawing?.id === d.id) && d.anchored);
  const needDrawing = 'Click a picture, a shape or a text box first';
  const [alignTo, setAlignTo] = React.useState(null);
  const alignMode = alignTo || (floatingSel.length > 1 ? 'selected' : 'margin');
  const isGroup = drawing?.kind === 'group';
  const formatTab = drawing ? (drawing.kind === 'picture' ? 'pictureFormat' : 'shapeFormat') : pickedIds.length ? (drawings.find((d) => d.id === pickedIds[0])?.kind === 'picture' ? 'pictureFormat' : 'shapeFormat') : null;
  // Insert → 3D Models: the picked picture, when it is a model, and the view it is drawn at.
  const pickedModel = picked?.image != null && !pickedIds.length ? (model?.blocks?.[picked.block]?.images?.[picked.image]?.model3d || null) : null;
  const look = drawing?.look || {};
  const shapeLike = drawing && (drawing.kind === 'textbox' || drawing.kind === 'shape');
  const sizeBox = (label, key) => {
    const px = key === 'w' ? (look.widthPx ?? drawing?.widthPx) : (look.heightPx ?? drawing?.heightPx);
    const cm = px ? Math.round((px / CM) * 100) / 100 : '';
    return (
      <div className="wd-fields">
        <label>{label}</label>
        <Input
          key={`${drawing?.id}:${key}:${cm}`}
          type="number" min="0.1" step="0.1" defaultValue={cm} disabled={!drawing}
          title={`${label} (cm)`}
          style={{ width: 64 }}
          onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }}
          onBlur={(e) => {
            const v = Number(e.target.value);
            if (!(v > 0) || Math.abs(v - cm) < 0.005) return;
            act('drawingSize', key === 'w' ? { widthPx: Math.round(v * CM) } : { heightPx: Math.round(v * CM) });
          }}
        />
        <span className="wd-field-value" style={{ minWidth: 0 }}>cm</span>
      </div>
    );
  };
  const arrange = (
    <Group label="Arrange">
      <Button tall icon="grid" label="Position" disabled={!hasDrawing} title={hasDrawing ? 'Position — where the drawing sits, with the words round it' : needDrawing} onClick={(e) => menu.open(e, [
        { label: 'Left, words round it', icon: 'alignLeft', run: () => act('position', 'left') },
        { label: 'Centre, words above and below', icon: 'alignCenter', run: () => act('position', 'center') },
        { label: 'Right, words round it', icon: 'alignRight', run: () => act('position', 'right') },
        { heading: true, label: 'With text wrapping' },
        ...['top', 'middle', 'bottom'].flatMap((v) => ['left', 'center', 'right'].map((h) => ({ label: `${cap(v)} ${h === 'center' ? 'Centre' : cap(h)}`, run: () => act('position', `${v}-${h}`) }))),
      ])} />
      <Button tall icon="picture" label="Wrap Text" disabled={!hasDrawing} title={hasDrawing ? 'Wrap Text — how the words treat the drawing' : needDrawing} onClick={(e) => menu.open(e, [
        { label: 'In line with text', icon: drawing && !drawing.anchored ? 'check' : undefined, run: () => act('wrap', 'inline') },
        { label: 'Square', icon: drawing?.anchored && drawing.wrap === 'square' ? 'check' : undefined, run: () => act('wrap', 'square') },
        { label: 'Tight', icon: drawing?.anchored && drawing.wrap === 'tight' ? 'check' : undefined, run: () => act('wrap', 'tight') },
        { label: 'Through', icon: drawing?.anchored && drawing.wrap === 'through' ? 'check' : undefined, run: () => act('wrap', 'through') },
        { label: 'Top and bottom', icon: drawing?.anchored && drawing.wrap === 'topAndBottom' ? 'check' : undefined, run: () => act('wrap', 'topAndBottom') },
        '-',
        { label: 'Behind text', icon: drawing?.anchored && drawing.wrap === 'none' && drawing.behind ? 'check' : undefined, run: () => act('wrap', 'behind') },
        { label: 'In front of text', icon: drawing?.anchored && drawing.wrap === 'none' && !drawing.behind ? 'check' : undefined, run: () => act('wrap', 'front') },
      ])} />
      <Button tall icon="chevronUp" label="Bring Forward" disabled={!floatingSel.length} title={floatingSel.length ? 'Bring Forward — one place nearer the front; also Bring to Front and Bring in Front of Text' : 'Bring Forward — select a floating drawing first'} onClick={(e) => menu.open(e, [
        { label: 'Bring Forward', icon: 'chevronUp', run: () => act('order', 'forward') },
        { label: 'Bring to Front', run: () => act('order', 'front') },
        { label: 'Bring in Front of Text', run: () => act('order', 'inFront') },
      ])} />
      <Button tall icon="chevronDown" label="Send Backward" disabled={!floatingSel.length} title={floatingSel.length ? 'Send Backward — one place further back; also Send to Back and Send Behind Text' : 'Send Backward — select a floating drawing first'} onClick={(e) => menu.open(e, [
        { label: 'Send Backward', icon: 'chevronDown', run: () => act('order', 'backward') },
        { label: 'Send to Back', run: () => act('order', 'back') },
        { label: 'Send Behind Text', run: () => act('order', 'behind') },
      ])} />
      <Rows>
        <Button icon="list" label="Selection Pane" pressed={Boolean(view.selectionPane)} title="Selection Pane — every drawing listed: select, hide, rename, reorder" onClick={() => act('selectionPane')} />
        <>
          <Button icon="alignLeft" label="Align" disabled={!floatingSel.length} title={floatingSel.length ? `Align — to the ${alignMode === 'selected' ? 'selected drawings' : alignMode}` : 'Align — select a floating drawing first'} onClick={(e) => menu.open(e, [
            { label: 'Align Left', icon: 'alignLeft', run: () => act('align', { edge: 'left', to: alignMode }) },
            { label: 'Align Centre', icon: 'alignCenter', run: () => act('align', { edge: 'center', to: alignMode }) },
            { label: 'Align Right', icon: 'alignRight', run: () => act('align', { edge: 'right', to: alignMode }) },
            { label: 'Align Top', run: () => act('align', { edge: 'top', to: alignMode }) },
            { label: 'Align Middle', run: () => act('align', { edge: 'middle', to: alignMode }) },
            { label: 'Align Bottom', run: () => act('align', { edge: 'bottom', to: alignMode }) },
            '-',
            { label: 'Distribute Horizontally', disabled: floatingSel.length < 3, title: floatingSel.length < 3 ? 'Select three or more drawings to distribute' : undefined, run: () => act('distribute', { axis: 'horizontal' }) },
            { label: 'Distribute Vertically', disabled: floatingSel.length < 3, title: floatingSel.length < 3 ? 'Select three or more drawings to distribute' : undefined, run: () => act('distribute', { axis: 'vertical' }) },
            '-',
            { label: 'Align to Page', icon: alignMode === 'page' ? 'check' : undefined, run: () => setAlignTo('page') },
            { label: 'Align to Margin', icon: alignMode === 'margin' ? 'check' : undefined, run: () => setAlignTo('margin') },
            { label: 'Align Selected Objects', icon: alignMode === 'selected' ? 'check' : undefined, disabled: floatingSel.length < 2, run: () => setAlignTo('selected') },
          ])} />
          <Button icon="grid" label="Group" disabled={floatingSel.length < 2 && !isGroup} title={floatingSel.length >= 2 ? 'Group — the selected drawings as one' : isGroup ? 'Group — ungroup this group' : 'Group — select two or more floating drawings (Shift+click adds one)'} onClick={(e) => menu.open(e, [
            { label: 'Group', icon: 'grid', disabled: floatingSel.length < 2, run: () => act('group') },
            { label: 'Ungroup', disabled: !isGroup, run: () => act('ungroup') },
          ])} />
        </>
        <Button icon="rotate" label="Rotate" disabled={!hasDrawing} title={hasDrawing ? 'Rotate — turn or flip the drawing' : needDrawing} onClick={(e) => menu.open(e, [
          { label: 'Rotate Right 90°', icon: 'rotate', run: () => act('rotate', 90) },
          { label: 'Rotate Left 90°', run: () => act('rotate', -90) },
          { label: 'Flip Vertical', icon: 'flip', run: () => act('rotate', 'flipV') },
          { label: 'Flip Horizontal', run: () => act('rotate', 'flipH') },
        ])} />
      </Rows>
    </Group>
  );

  const run = (delta) => dispatch({ op: 'setRunFormat', delta });
  const para = (delta) => dispatch({ op: 'setParagraphFormat', delta });
  const toggleList = (kind) => para({ list: format.listType === kind ? null : kind });
  const colourMenu = (event, key, list) =>
    menu.open(event, list.map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => run({ [key]: value }) })));
  const size = Number(format.fontSize || 11);
  const nearer = (dir) => {
    const at = SIZES.findIndex((s) => s >= size);
    const next = dir > 0 ? SIZES[Math.min(SIZES.length - 1, (at < 0 ? SIZES.length - 1 : at) + 1)] : SIZES[Math.max(0, (at < 0 ? SIZES.length : at) - 1)];
    return next ?? size;
  };

  return (
    <Ribbon
      tabs={[
        { id: 'home', label: 'Home' },
        { id: 'insert', label: 'Insert' },
        { id: 'draw', label: 'Draw' },
        { id: 'design', label: 'Design' },
        { id: 'layout', label: 'Layout' },
        { id: 'references', label: 'References' },
        { id: 'mailings', label: 'Mailings' },
        { id: 'review', label: 'Review' },
        { id: 'view', label: 'View' },
        { id: 'help', label: 'Help' },
        { id: 'pdf', label: 'PDF' },
        // The contextual tab, as Word's: there while a drawing is selected.
        ...(formatTab === 'shapeFormat' ? [{ id: 'shapeFormat', label: 'Shape Format' }] : []),
        ...(formatTab === 'pictureFormat' ? [{ id: 'pictureFormat', label: 'Picture Format' }] : []),
        // 3D Model, as Word's: there while a 3D model is picked.
        ...(pickedModel ? [{ id: 'model3d', label: '3D Model' }] : []),
        // Table Design and Table Layout, as Word's: there while the caret is in a table.
        ...(table ? [{ id: 'tableDesign', label: 'Table Design' }, { id: 'tableLayout', label: 'Table Layout' }] : []),
        // View → Immersive Reader's own tab, there while the reader is open.
        ...(view.immersive ? [{ id: 'immersive', label: 'Immersive Reader' }] : []),
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title="Save (Ctrl+S)" onClick={() => save(false)} />
          <Button icon="undo" title="Undo (Ctrl+Z)" disabled={!doc?.canUndo} onClick={() => commands['edit.undo'].run()} />
          <Button icon="redo" title="Redo (Ctrl+Y)" disabled={!doc?.canRedo} onClick={() => commands['edit.redo'].run()} />
          <Button icon="print" title="Print" onClick={() => commands['file.print']?.run?.()} />
        </>
      }
    >
      {/* ── Home ─────────────────────────────────────────────────────────── */}
      {tab === 'home' ? (
        <>
          <Group label="Clipboard">
            <Button tall icon="paste" label="Paste" title="Paste (Ctrl+V)" onClick={() => commands['edit.paste']?.run?.()} />
            <Rows>
              <>
                <Button icon="cut" label="Cut" title="Cut (Ctrl+X)" onClick={() => commands['edit.cut']?.run?.()} />
                <Button icon="copy" label="Copy" title="Copy (Ctrl+C)" onClick={() => commands['edit.copy']?.run?.()} />
              </>
              <Button icon="wand" label="Format Painter" pressed={Boolean(view.painting)} title="Format Painter — copy the formatting here, then select the words to paint" onClick={() => act('formatPainter')} />
            </Rows>
          </Group>

          <Group label="Font">
            <Rows>
              <>
                <Select value={format.fontName || 'Calibri'} onChange={(e) => run({ fontName: e.target.value })} style={{ width: 132 }} title="Font">
                  {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
                  {format.fontName && !FONTS.includes(format.fontName) ? <option value={format.fontName}>{format.fontName}</option> : null}
                </Select>
                <Select value={String(size)} onChange={(e) => run({ fontSize: Number(e.target.value) })} style={{ width: 58 }} title="Size">
                  {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
                  {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
                </Select>
                <Button icon="chevronUp" title="Grow font" onClick={() => run({ fontSize: nearer(1) })} />
                <Button icon="chevronDown" title="Shrink font" onClick={() => run({ fontSize: nearer(-1) })} />
                <Button icon="textbox" title="Change case" onClick={(e) => menu.open(e, CASES.map(([mode, label]) => ({ label, run: () => act('changeCase', mode) })))} />
                <Button icon="undo" title="Clear all formatting" onClick={() => dispatch({ op: 'clearFormat' })} />
              </>
              <>
                <Button icon="bold" title="Bold (Ctrl+B)" pressed={format.bold} onClick={() => dispatch({ op: 'toggleFormat', tag: 'bold' })} />
                <Button icon="italic" title="Italic (Ctrl+I)" pressed={format.italic} onClick={() => dispatch({ op: 'toggleFormat', tag: 'italic' })} />
                <Button icon="underline" title="Underline (Ctrl+U)" pressed={format.underline} onClick={() => dispatch({ op: 'toggleFormat', tag: 'underline' })} />
                <Button icon="strike" title="Strikethrough" pressed={format.strike} onClick={() => dispatch({ op: 'toggleFormat', tag: 'strike' })} />
                <Button icon="chevronDown" label="x₂" title="Subscript" pressed={format.vertAlign === 'subscript'} onClick={() => run({ vertAlign: format.vertAlign === 'subscript' ? null : 'subscript' })} />
                <Button icon="chevronUp" label="x²" title="Superscript" pressed={format.vertAlign === 'superscript'} onClick={() => run({ vertAlign: format.vertAlign === 'superscript' ? null : 'superscript' })} />
                <Separator />
                <Button icon="wand" label="A" title="Text effects — an outline, a shadow or a glow on the selected words" onClick={(e) => menu.open(e, [
                  { label: 'Outline', icon: format.outline ? 'check' : undefined, run: () => run({ outline: !format.outline }) },
                  { label: 'Shadow', icon: format.shadow ? 'check' : undefined, run: () => run({ shadow: !format.shadow }) },
                  '-',
                  ...GLOWS.map(([colour, label]) => ({ label: `Glow: ${label}`, icon: format.glow?.colour === colour ? 'check' : undefined, run: () => run({ glow: { colour, radiusPt: 4 } }) })),
                  { label: 'No glow', run: () => run({ glow: null }) },
                  '-',
                  { label: 'Clear effects', run: () => run({ outline: false, shadow: false, glow: null }) },
                ])} />
                <Button icon="wand" title="Text highlight colour" pressed={Boolean(format.highlight)} onClick={(e) => colourMenu(e, 'highlight', HIGHLIGHTS)} />
                <Button icon="contrast" title="Font colour" onClick={(e) => colourMenu(e, 'fontColour', TEXT_COLOURS)} />
              </>
            </Rows>
          </Group>

          <Group label="Paragraph">
            <Rows>
              <>
                <Button icon="listBullet" title="Bullets" pressed={format.listType === 'bullet'} onClick={() => toggleList('bullet')} />
                <Button icon="listNumber" title="Numbering" pressed={format.listType === 'number'} onClick={() => toggleList('number')} />
                <Button icon="listNumber" title="Multilevel list — 1. 1.1. 1.1.1., deeper with Increase indent" pressed={format.listType === 'number' && (format.listLevel || 0) > 0} onClick={() => para({ list: 'outline' })} />
                <Separator />
                <Button icon="chevronLeft" title="Decrease indent" disabled={!format.indentLevel} onClick={() => para({ indentDelta: -1 })} />
                <Button icon="chevronRight" title="Increase indent" onClick={() => para({ indentDelta: 1 })} />
                <Button icon="sort" title="Sort paragraphs — A to Z or Z to A; the whole document when nothing is selected" onClick={(e) => menu.open(e, [
                  { label: 'A to Z', run: () => dispatch({ op: 'sortParagraphs', descending: false }) },
                  { label: 'Z to A', run: () => dispatch({ op: 'sortParagraphs', descending: true }) },
                ])} />
                <Button icon="formula" title="Show formatting marks (¶)" pressed={Boolean(view.marks)} onClick={() => act('toggleMarks')} />
                <Separator />
                <Button icon="textLtr" title="Left-to-right text direction" pressed={!format.rtl} onClick={() => para({ rtl: false })} />
                <Button icon="textRtl" title="Right-to-left text direction — the paragraph runs from the right margin, as Arabic, Hebrew, Persian and Urdu do" pressed={Boolean(format.rtl)} onClick={() => para({ rtl: true })} />
              </>
              <>
                <Button icon="alignLeft" title="Align left (Ctrl+L)" pressed={format.paragraphAlign === 'left'} onClick={() => para({ align: 'left' })} />
                <Button icon="alignCenter" title="Centre (Ctrl+E)" pressed={format.paragraphAlign === 'center'} onClick={() => para({ align: 'center' })} />
                <Button icon="alignRight" title="Align right (Ctrl+R)" pressed={format.paragraphAlign === 'right'} onClick={() => para({ align: 'right' })} />
                <Button icon="alignJustify" title="Justify (Ctrl+J)" pressed={format.paragraphAlign === 'justify' || format.paragraphAlign === 'both'} onClick={() => para({ align: 'both' })} />
                <Separator />
                <Button icon="listNumber" title={`Line and paragraph spacing${format.lineSpacing ? ` — ${format.lineSpacing}` : ''}`} onClick={(e) =>
                  menu.open(e, [
                    ...SPACING.map(([v, l]) => ({ label: l, icon: format.lineSpacing === v ? 'check' : undefined, run: () => para({ lineSpacing: v }) })),
                    '-',
                    { label: 'Add space before paragraph', run: () => para({ spaceBefore: 12 }) },
                    { label: 'Remove space before paragraph', run: () => para({ spaceBefore: 0 }) },
                    { label: 'Add space after paragraph', run: () => para({ spaceAfter: 8 }) },
                    { label: 'Remove space after paragraph', run: () => para({ spaceAfter: 0 }) },
                  ])
                } />
                <Button icon="wand" title="Shading — a colour behind the paragraph" onClick={(e) => menu.open(e, SHADES.map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => para({ shading: value }) })))} />
                <Button icon="grid" title="Borders — lines round the paragraph" onClick={(e) => menu.open(e, BORDERS.map(([label, value]) => ({ label, icon: value ? undefined : 'close', run: () => para({ borders: value }) })))} />
              </>
            </Rows>
          </Group>

          <Group label="Styles">
            <div className="wd-styles">
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).slice(0, 8).map((s) => (
                <button
                  key={s.id}
                  type="button"
                  className={`wd-style${(format.paragraphStyle || 'Normal') === s.id ? ' on' : ''}`}
                  title={`Apply ${s.name || s.id}`}
                  onClick={() => para({ styleId: s.id === 'Normal' && !styles.some((x) => x.id === 'Normal') ? null : s.id })}
                  style={STYLE_LOOK[s.id] || undefined}
                >
                  {s.name || s.id}
                </button>
              ))}
            </div>
            <Select
              value={format.paragraphStyle || 'Normal'}
              onChange={(e) => para({ styleId: e.target.value === 'Normal' && !styles.some((s) => s.id === 'Normal') ? null : e.target.value })}
              style={{ width: 120 }}
              title="Paragraph style"
            >
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).map((s) => <option key={s.id} value={s.id}>{s.name || s.id}</option>)}
            </Select>
          </Group>

          <Group label="Editing">
            <Rows>
              <>
                <Button icon="find" label="Find" title="Find (Ctrl+F)" onClick={() => openDialog('find')} />
                <Button icon="refresh" label="Replace" title="Replace (Ctrl+H)" onClick={() => openDialog('find')} />
              </>
              <Button icon="check" label="Select" title="Select all (Ctrl+A)" onClick={() => dispatch({ op: 'selectAll' })} />
            </Rows>
          </Group>

          <Group label="Voice">
            <Button tall icon="volume" label="Read aloud" pressed={Boolean(view.reading)} title="Read the document aloud, from the caret" onClick={() => act('readAloud')} />
          </Group>
        </>
      ) : null}

      {/* ── Insert ───────────────────────────────────────────────────────── */}
      {tab === 'insert' ? (
        <>
          <Group label="Pages">
            <Button tall icon="file" label="Cover Page" title="A title page at the front" onClick={() => act('coverPage')} />
            <Button tall icon="file" label="Blank Page" title="A page break here, and another after — an empty page" onClick={() => act('blankPage')} />
            <Button tall icon="file" label="Page Break" title="Page break (Ctrl+Enter)" onClick={() => dispatch({ op: 'insertPageBreak' })} />
          </Group>
          <Group label="Tables">
            <Button tall icon="table" label="Table" onClick={() => openDialog('table')} />
            <Button icon="table" label="Text to Table" disabled={Boolean(table)} title={table ? 'Convert Text to Table — the caret is in a table already' : 'Convert Text to Table — the selected paragraphs as rows, split at their tabs or commas'} onClick={(e) => menu.open(e, [
              { label: 'Split at tabs', run: () => dispatch({ op: 'textToTable', separator: 'tab' }) },
              { label: 'Split at commas', run: () => dispatch({ op: 'textToTable', separator: 'comma' }) },
            ])} />
          </Group>
          <Group label="Illustrations">
            <Button tall icon="picture" label="Pictures" onClick={insertPicture} />
            <Button tall icon="shape" label="Shapes" onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', preset, widthPx: 200, heightPx: 120 }) })))} />
            <Button tall icon="star" label="Icons" title="Icons — one of the suite's own icons, in the colour you choose, as a picture at the caret" onClick={() => act('icons')} />
            <Button tall icon="shape" label="3D Models" className="wd-model3d-insert" title="3D Models — a model from a .glb or .gltf file on this computer, drawn as a picture at the caret; turn it from the 3D Model tab" onClick={() => act('model3d')} />
            <Button tall icon="grid" label="SmartArt" title="SmartArt — a list, a process, a cycle or a hierarchy, drawn from lines you type, as a group of shapes after this paragraph" onClick={() => act('insertSmartArt')} />
            <Button tall icon="chart" label="Chart" title="A chart drawn from the table the caret is in — put the caret in a table first" onClick={(e) => menu.open(e, CHARTS.map(([k, l]) => ({ label: l, icon: 'chart', run: () => dispatch({ op: 'insertChart', kind: k }) })))} />
            <Button tall icon="picture" label="Screenshot" title="Screenshot — a picture of another open window, or of a whole screen, put in at the caret" onClick={() => act('screenshot')} />
          </Group>
          <Group label="Media">
            <Soon tall icon="video" label="Online Videos" why="Embedding a video is a network request; this suite makes none it has not declared." />
          </Group>
          <Group label="Links">
            <Button tall icon="link" label="Link" title="Link (Ctrl+K)" pressed={Boolean(format.link)} onClick={() => openDialog('link')} />
            <Button icon="flag" label="Bookmark" title="Bookmark — name this place in the document, so Go To and a cross-reference can find it" onClick={() => openDialog('bookmark')} />
            <Button icon="link" label="Cross-reference" title="Cross-reference — a REF field to a bookmark, its words kept live by Update Fields" onClick={() => openDialog('crossReference')} />
          </Group>
          <Group label="Comments">
            <Button tall icon="reply" label="Comment" onClick={() => openDialog('comment')} />
          </Group>
          <Group label="Header & Footer">
            <Button tall icon="chevronUp" label="Header" onClick={() => openDialog('header')} />
            <Button tall icon="chevronDown" label="Footer" onClick={() => openDialog('footer')} />
            <Button tall icon="file" label="Page Number" onClick={() => openDialog('pageNumber')} />
          </Group>
          <Group label="Text">
            <Button tall icon="textbox" label="Text Box" title="Text Box — a box of words that floats on the page, the text wrapping round it" onClick={(e) => menu.open(e, [
              { heading: true, label: 'Built-in' },
              { label: 'Simple Text Box', icon: 'textbox', run: () => act('textBox', 'simple') },
              { label: 'Sidebar', icon: 'textbox', run: () => act('textBox', 'sidebar') },
              { label: 'Pull Quote', icon: 'textbox', run: () => act('textBox', 'quote') },
              '-',
              { label: 'Draw Text Box', icon: 'shape', run: () => act('drawTextBox') },
            ])} />
            {references ? <Button icon="file" label="Quick Parts" title="Quick Parts — words you keep to put in again, AutoText, a field: a page number, the date, the file's name" onClick={(e) => menu.open(e, [...(blocks?.head() || []), ...references.quickParts(), ...(blocks?.tail() || [])])} /> : null}
            <Button icon="wand" label="WordArt" title="WordArt — decorative words in a box of their own that floats on the page" onClick={(e) => menu.open(e, WORDART.map(([label, spec]) => ({ label, icon: 'wand', run: () => act('wordArt', spec) })))} />
            <Button icon="textbox" label="Drop Cap" title="Drop Cap — the first letter, framed to stand tall beside the words that follow it" onClick={(e) => menu.open(e, [
              { label: 'None', icon: !format.dropCap ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: null }) },
              { label: 'Dropped', icon: format.dropCap?.kind === 'drop' && format.dropCap?.lines === 3 ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: { kind: 'drop', lines: 3 } }) },
              { label: 'Dropped, two lines', icon: format.dropCap?.kind === 'drop' && format.dropCap?.lines === 2 ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: { kind: 'drop', lines: 2 } }) },
              { label: 'In margin', icon: format.dropCap?.kind === 'margin' ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: { kind: 'margin', lines: 3 } }) },
            ])} />
            <Button icon="check" label="Signature Line" title="Signature Line — a line for someone to sign in Word, with their name, title and e-mail under it" onClick={() => act('signatureLine')} />
            <Button icon="clock" label="Date & Time" onClick={() => openDialog('dateTime')} />
            <Button icon="file" label="Object" title="Object — a Word, Excel or PowerPoint document embedded as an icon; a double-click opens it" onClick={() => act('insertObject')} />
          </Group>
          <Group label="Symbols">
            <Button tall icon="formula" label="Equation" title="Equation — type one in Word's linear format, or build it from structures and symbols (Alt+=)" onClick={() => act('equation')} />
            <Button
              icon="chevronDown"
              label="Built-in"
              title="Built-in equations — the quadratic formula, the area of a circle, the binomial theorem and more, put in at the caret"
              onClick={(e) => menu.open(e, EQUATION_GALLERY.map((g) => ({ label: g.name, icon: 'formula', run: () => act('equation', { linear: g.linear }) })))}
            />
            <Button tall icon="star" label="Symbol" onClick={() => openDialog('symbol')} />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label="Drawing Tools">
            <Button tall icon="mouse" label="Select" pressed={!ink?.tool} title="Select — put the pen down and type again (Esc)" onClick={() => act('inkTool', null)} />
            <Button tall icon="shape" label="Lasso" pressed={ink?.tool === 'lasso'} title="Lasso Select — draw a loop round strokes to select them" onClick={() => act('inkTool', 'lasso')} />
            <Button tall icon="close" label="Eraser" pressed={ink?.tool === 'eraser'} title="Eraser — take away each stroke the pointer passes over" onClick={() => act('inkTool', 'eraser')} />
            {(ink?.pens || []).map((p) => {
              const on = ink.tool === 'pen' && ink.penId === p.id;
              const label = { pen: 'Pen', pencil: 'Pencil', highlighter: 'Highlighter' }[p.tool];
              return (
                <Button key={p.id} tall icon="wand" label={label} pressed={on} className={`sl-pen sl-pen-${p.tool}`} data-pen={p.id} style={{ '--pen': p.color }}
                  title={on ? `${label} — click again for its colour and thickness` : `${label} — draw on the page`}
                  onClick={(e) => {
                    if (!on) { act('inkTool', p.id); return; }
                    menu.open(e, [
                      { heading: true, label: 'Thickness' },
                      ...PEN_WIDTHS[p.tool].map((w) => ({ label: `${w} pt`, icon: p.width === w ? 'check' : undefined, run: () => act('inkPen', { id: p.id, width: w }) })),
                      { heading: true, label: 'Colour' },
                      ...PEN_COLOURS.map((c) => ({ label: c, icon: p.color === c ? 'check' : undefined, preview: <span style={{ display: 'inline-block', width: 16, height: 16, borderRadius: 8, background: c, border: '1px solid rgba(0,0,0,.25)' }} />, run: () => act('inkPen', { id: p.id, color: c }) })),
                    ]);
                  }} />
              );
            })}
            <Button tall icon="plus" label="Add" title="Add Pen — another pen or highlighter in the gallery" onClick={(e) => menu.open(e, [['pen', 'Pen'], ['highlighter', 'Highlighter']].map(([tool, label]) => ({ label, icon: 'plus', run: () => act('inkAdd', tool) })))} />
          </Group>
          <Group label="Touch">
            <Button tall icon="wand" label="Draw with Touch" pressed={ink?.touch !== false} title="Draw with Touch — a finger draws with the pen in hand; off, only a pen or the mouse does" onClick={() => act('inkTouch')} />
          </Group>
          <Group label="Stencils">
            <Button tall icon="minus" label="Ruler" pressed={Boolean(ink?.ruler)} title="Ruler — a straight edge on the page: a stroke begun along it follows it; drag it to move, the wheel to turn it" onClick={() => act('inkRuler')} />
          </Group>
          <Group label="Convert">
            <Button tall icon="shape" label="Ink to Shape" pressed={Boolean(ink?.toShape)} title="Ink to Shape — a rectangle, oval or triangle drawn becomes that shape" onClick={() => act('inkToShape')} />
            <Soon tall icon="formula" label="Ink to Maths" why="Turning handwriting into an equation needs handwriting recognition, which this suite does not have." />
          </Group>
          <Group label="Insert">
            <Button tall icon="shape" label="Shape" title="The shapes the engine can draw" onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', preset, widthPx: 200, heightPx: 120 }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Design ───────────────────────────────────────────────────────── */}
      {tab === 'design' ? (
        <>
          <Group label="Document Formatting">
            <Button
              tall
              icon="wand"
              label="Themes"
              title={`Themes — colours, fonts and effects for the whole document, in one; now ${design?.name || 'Office Theme'}`}
              onClick={(e) => menu.open(e, THEMES.map((t) => ({ label: t.name, icon: design?.builtIn === t.id ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { theme: t.id } }) })))}
            />
            <div className="wd-styles">
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).slice(0, 6).map((s) => (
                <button key={s.id} type="button" className={`wd-style${(format.paragraphStyle || 'Normal') === s.id ? ' on' : ''}`} style={STYLE_LOOK[s.id] || undefined} title={`Apply ${s.name || s.id} to this paragraph`} onClick={() => para({ styleId: s.id })}>
                  {s.name || s.id}
                </button>
              ))}
            </div>
            <Button
              icon="contrast"
              label="Colours"
              title={`Colours — the theme's twelve colours, which headings and theme-coloured text follow; now ${design?.colorName || 'Office'}`}
              onClick={(e) => menu.open(e, PALETTES.map((p) => ({ label: p.name, icon: design?.colorName === p.name ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { colors: p.id } }) })))}
            />
            <Button
              icon="word"
              label="Fonts"
              title={`Fonts — the theme's heading and body faces, which the styles follow; now ${design?.fonts ? `${design.fonts.major} / ${design.fonts.minor}` : 'Calibri Light / Calibri'}`}
              onClick={(e) => menu.open(e, FONT_PAIRS.map((p) => ({ label: `${p.name} — ${p.major} / ${p.minor}`, icon: design?.fontName === p.name ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { fonts: p.id } }) })))}
            />
            <Button icon="listNumber" label="Paragraph Spacing" title="Spacing for the paragraphs you have selected" onClick={(e) =>
              menu.open(e, [
                { label: 'No paragraph space', run: () => para({ spaceBefore: 0, spaceAfter: 0, lineSpacing: 1 }) },
                { label: 'Compact', run: () => para({ spaceBefore: 0, spaceAfter: 4, lineSpacing: 1 }) },
                { label: 'Tight', run: () => para({ spaceBefore: 0, spaceAfter: 6, lineSpacing: 1.15 }) },
                { label: 'Open', run: () => para({ spaceBefore: 0, spaceAfter: 10, lineSpacing: 1.15 }) },
                { label: 'Relaxed', run: () => para({ spaceBefore: 0, spaceAfter: 6, lineSpacing: 1.5 }) },
                { label: 'Double', run: () => para({ spaceBefore: 0, spaceAfter: 8, lineSpacing: 2 }) },
              ])
            } />
            <Button
              icon="wand"
              label="Effects"
              title={`Effects — the fills, lines and shadows that shapes taking their look from the theme are drawn with; now ${design?.effectName || 'Office'}`}
              onClick={(e) => menu.open(e, EFFECT_PRESETS.map((p) => ({ label: p.name, title: p.description, icon: design?.effects === p.id ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { effects: p.id } }) })))}
            />
            <Button icon="check" label="Set as Default" title="Set as Default — new blank documents start in this document's theme and styles" onClick={() => act('setDefaultDesign')} />
          </Group>
          <Group label="Page Background">
            <Button tall icon="shield" label="Watermark" title="Watermark — faint words, or a picture, behind every page" onClick={(e) => menu.open(e, [
              ...WATERMARKS.map((text) => ({ label: text, run: () => dispatch({ op: 'setWatermark', text }) })),
              { label: 'Custom watermark…', icon: 'textbox', run: () => openDialog('watermark') },
              { label: 'Picture watermark…', icon: 'picture', run: () => act('pictureWatermark') },
              { label: 'Remove watermark', icon: 'close', run: () => dispatch({ op: 'setWatermark', text: null }) },
            ])} />
            <Button tall icon="contrast" label="Page Colour" title="Page Colour — a colour behind every page" onClick={(e) => menu.open(e, SHADES.map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => dispatch({ op: 'setPageColour', colour: value }) })))} />
            <Button tall icon="grid" label="Page Borders" title="Page Borders — a frame around every page" onClick={(e) => menu.open(e, PAGE_BORDERS.map(([label, borders]) => ({ label, icon: borders ? undefined : 'close', run: () => dispatch({ op: 'setPageBorders', borders }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Layout ───────────────────────────────────────────────────────── */}
      {tab === 'layout' ? (
        <>
          <Group label="Page Setup">
            <Button tall icon="crop" label="Margins" onClick={(e) => menu.open(e, MARGINS.map(([v, l]) => ({ label: l, run: () => dispatch({ op: 'setPageSetup', spec: { margins: v } }) })))} />
            <Button tall icon="rotate" label="Orientation" title={`Orientation — now ${landscape ? 'landscape' : 'portrait'}`} onClick={(e) =>
              menu.open(e, [
                { label: 'Portrait', icon: landscape ? undefined : 'check', run: () => dispatch({ op: 'setPageSetup', spec: { orientation: 'portrait' } }) },
                { label: 'Landscape', icon: landscape ? 'check' : undefined, run: () => dispatch({ op: 'setPageSetup', spec: { orientation: 'landscape' } }) },
              ])
            } />
            <Button tall icon="file" label="Size" onClick={(e) => menu.open(e, PAGE_SIZES.map(([v, l]) => ({ label: l, run: () => dispatch({ op: 'setPageSetup', spec: { size: v } }) })))} />
            <Button tall icon="grid" label="Columns" title={`Columns — now ${colCount > 1 ? colCount : 'one'}`} onClick={(e) => menu.open(e, [
              { label: 'One', icon: isColumnPreset('one') ? 'check' : undefined, run: () => setColumns(null) },
              { label: 'Two', icon: isColumnPreset('two') ? 'check' : undefined, run: () => setColumns({ count: 2, spaceTwips: 720, separator: Boolean(columns.separator) }) },
              { label: 'Three', icon: isColumnPreset('three') ? 'check' : undefined, run: () => setColumns({ count: 3, spaceTwips: 720, separator: Boolean(columns.separator) }) },
              { label: 'Left', icon: isColumnPreset('left') ? 'check' : undefined, run: () => setColumns(unevenSpec(true)) },
              { label: 'Right', icon: isColumnPreset('right') ? 'check' : undefined, run: () => setColumns(unevenSpec(false)) },
              '-',
              // The section right to left: its columns run from the right, as an Arabic or Hebrew page's do.
              { label: 'Right to left', icon: section?.rtl ? 'check' : undefined, run: () => dispatch({ op: 'setPageSetup', spec: { rtl: !section?.rtl } }) },
              {
                label: 'Line between', icon: columns.separator ? 'check' : undefined, disabled: colCount <= 1,
                run: () => setColumns({
                  count: colCount, spaceTwips: Math.round((columns.spacePx || 36) * 15), separator: !columns.separator,
                  ...(columns.widths ? { widths: columns.widths.map((w) => Math.round(w * 15)) } : {}),
                }),
              },
            ])} />
            <Button icon="file" label="Breaks" onClick={(e) =>
              menu.open(e, [
                { label: 'Page break', icon: 'file', run: () => dispatch({ op: 'insertPageBreak' }) },
                { label: 'Section break — not built yet', disabled: true },
              ])
            } />
            <Button icon="listNumber" label="Line Numbers" title="Line Numbers — a number beside every line, down the left margin" onClick={(e) => menu.open(e, [
              { label: 'None', icon: 'close', run: () => dispatch({ op: 'setLineNumbers', spec: null }) },
              { label: 'Continuous', run: () => dispatch({ op: 'setLineNumbers', spec: { countBy: 1, restart: 'continuous' } }) },
              { label: 'Restart each page', run: () => dispatch({ op: 'setLineNumbers', spec: { countBy: 1, restart: 'newPage' } }) },
              { label: 'Every fifth line', run: () => dispatch({ op: 'setLineNumbers', spec: { countBy: 5, restart: 'continuous' } }) },
            ])} />
            <Button icon="minus" label="Hyphenation" title={`Hyphenation — words broken at the ends of lines: now ${model?.hyphenation?.auto ? 'automatic' : 'none'}`} onClick={(e) => menu.open(e, [
              { label: 'None', icon: !model?.hyphenation?.auto ? 'check' : undefined, run: () => act('hyphenation', 'none') },
              { label: 'Automatic', icon: model?.hyphenation?.auto ? 'check' : undefined, run: () => act('hyphenation', 'auto') },
              { label: 'Manual', run: () => act('hyphenation', 'manual') },
              '-',
              { label: 'Hyphenation Options…', run: () => act('hyphenation', 'options') },
              { label: "Don't hyphenate this paragraph", icon: format.noHyphens ? 'check' : undefined, run: () => act('hyphenation', 'paragraph') },
            ])} />
          </Group>
          <Group label="Paragraph">
            <div className="wd-fields">
              <label>Indent</label>
              <Button icon="chevronLeft" title="Decrease left indent" disabled={!format.indentLevel} onClick={() => para({ indentDelta: -1 })} />
              <span className="wd-field-value">{(format.indentLevel || 0) * 1.27} cm</span>
              <Button icon="chevronRight" title="Increase left indent" onClick={() => para({ indentDelta: 1 })} />
            </div>
            <div className="wd-fields">
              <label>Spacing</label>
              <Input type="number" min="0" max="200" value={format.spaceBefore ?? 0} title="Before (pt)" onChange={(e) => para({ spaceBefore: Number(e.target.value) })} style={{ width: 56 }} />
              <Input type="number" min="0" max="200" value={format.spaceAfter ?? 0} title="After (pt)" onChange={(e) => para({ spaceAfter: Number(e.target.value) })} style={{ width: 56 }} />
            </div>
          </Group>
          {arrange}
        </>
      ) : null}

      {/* ── Shape Format (contextual) ─────────────────────────────────────── */}
      {tab === 'shapeFormat' ? (
        <>
          <Group label="Insert Shapes">
            <Button tall icon="textbox" label="Draw Text Box" pressed={Boolean(view.drawBox)} title="Draw Text Box — drag on the page to draw one" onClick={() => act('drawTextBox')} />
          </Group>
          <Group label="Shape Styles">
            <Button tall icon="wand" label="Shape Fill" disabled={!shapeLike} title={shapeLike ? `Shape Fill — now ${look.fill || 'no fill'}` : 'Shape Fill — select a text box or a shape first'} onClick={(e) => menu.open(e, [
              { label: 'No Fill', icon: 'close', run: () => act('boxFormat', { fill: null }) },
              '-',
              ...SHAPE_COLOURS.map(([hex, label]) => ({ label, icon: look.fill && look.fill.replace('#', '').toUpperCase() === hex ? 'check' : undefined, run: () => act('boxFormat', { fill: hex }) })),
            ])} />
            <Button tall icon="shape" label="Shape Outline" disabled={!shapeLike} title={shapeLike ? `Shape Outline — now ${look.line || 'no outline'}` : 'Shape Outline — select a text box or a shape first'} onClick={(e) => menu.open(e, [
              { label: 'No Outline', icon: 'close', run: () => act('boxFormat', { line: null }) },
              '-',
              ...SHAPE_COLOURS.map(([hex, label]) => ({ label, icon: look.line && look.line.replace('#', '').toUpperCase() === hex ? 'check' : undefined, run: () => act('boxFormat', { line: { colour: hex, widthPx: look.lineWidthPx || 1 } }) })),
              { heading: true, label: 'Weight' },
              ...WEIGHTS.map(([label, px]) => ({ label, icon: look.line && Math.abs((look.lineWidthPx || 1) - px) < 0.2 ? 'check' : undefined, run: () => act('boxFormat', { line: { colour: (look.line || '#000000').replace('#', ''), widthPx: px } }) })),
            ])} />
          </Group>
          <Group label="Text">
            <Button tall icon="rotate" label="Text Direction" disabled={drawing?.kind !== 'textbox'} title="Text Direction — across, or turned to read down or up" onClick={(e) => menu.open(e, [
              { label: 'Horizontal', icon: (look.vert || 'horz') === 'horz' ? 'check' : undefined, run: () => act('boxFormat', { vert: 'horz' }) },
              { label: 'Rotate all text 90°', icon: look.vert === 'vert' ? 'check' : undefined, run: () => act('boxFormat', { vert: 'vert' }) },
              { label: 'Rotate all text 270°', icon: look.vert === 'vert270' ? 'check' : undefined, run: () => act('boxFormat', { vert: 'vert270' }) },
            ])} />
            <Button tall icon="alignCenter" label="Align Text" disabled={drawing?.kind !== 'textbox'} title="Align Text — the words at the top, middle or bottom of the box" onClick={(e) => menu.open(e, [
              { label: 'Top', icon: (look.vAnchor || 'top') === 'top' ? 'check' : undefined, run: () => act('boxFormat', { vAnchor: 'top' }) },
              { label: 'Middle', icon: look.vAnchor === 'middle' ? 'check' : undefined, run: () => act('boxFormat', { vAnchor: 'middle' }) },
              { label: 'Bottom', icon: look.vAnchor === 'bottom' ? 'check' : undefined, run: () => act('boxFormat', { vAnchor: 'bottom' }) },
            ])} />
            <Button tall icon="crop" label="Margins" disabled={drawing?.kind !== 'textbox'} title="Margins — the room between the box's edge and its words" onClick={(e) => menu.open(e, BOX_MARGINS.map(([label, insets]) => ({
              label, icon: look.insets && ['l', 't', 'r', 'b'].every((k) => Math.abs((look.insets[k] || 0) - insets[k]) < 0.6) ? 'check' : undefined, run: () => act('boxFormat', { insets }),
            })))} />
          </Group>
          {/* WordArt Styles → Text Effects → Transform, as Word's: a text box's words along a path. */}
          <Group label="Text Effects: Transform">
            {WARP_PRESETS.map((p) => (
              <Button
                key={p.id} tall icon="wand" label={p.label} className="wd-warp" data-preset={p.id}
                disabled={drawing?.kind !== 'textbox'}
                pressed={drawing?.kind === 'textbox' && (look.warp || 'textNoShape') === p.id}
                title={drawing?.kind !== 'textbox' ? `${p.label} — select a text box or WordArt first` : p.id === 'textNoShape' ? 'No Transform — the words in straight lines' : `Transform — the words along ${p.label === 'Button' ? 'a button: an arc, a line and an arc' : `the ${p.label === 'Circle' ? 'circle' : 'arc'} of the box`}`}
                onClick={() => act('boxFormat', { warp: p.id === 'textNoShape' ? null : p.id })}
              />
            ))}
            {/* The warps, the rest of Office's gallery, from a menu. */}
            <Button tall disabled={drawing?.kind !== 'textbox'} icon="wand" label="More" className="wd-warp-more" pressed={WARP_MORE.some((p) => p.id === look.warp)} title={WARP_MORE.some((p) => p.id === look.warp) ? `Transform — now ${warpLabel(look.warp)}; the warps: the words stretched between two curves` : 'More Transforms — the warps: the words stretched between two curves, a wave, a slant, a chevron and the rest'} onClick={(e) => menu.open(e, WARP_MORE.map((p) => ({ label: p.label, icon: look.warp === p.id ? 'check' : 'wand', run: () => act('boxFormat', { warp: p.id }) })))} />
          </Group>
          {arrange}
          <Group label="Size">
            <Rows>
              {sizeBox('Height', 'h')}
              {sizeBox('Width', 'w')}
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Table Design (contextual) ─────────────────────────────────────── */}
      {tab === 'tableDesign' && table ? (
        <>
          <Group label="Table Style Options">
            <Rows>
              <Button icon={styleLook.firstRow ? 'check' : undefined} label="Header Row" pressed={styleLook.firstRow} title="Header Row — the first row in its style's header look" onClick={() => setLook({ firstRow: !styleLook.firstRow })} />
              <Button icon={styleLook.lastRow ? 'check' : undefined} label="Total Row" pressed={styleLook.lastRow} title="Total Row — the last row in its style's total look" onClick={() => setLook({ lastRow: !styleLook.lastRow })} />
              <Button icon={!styleLook.noHBand ? 'check' : undefined} label="Banded Rows" pressed={!styleLook.noHBand} title="Banded Rows — every other row shaded, as the style bands them" onClick={() => setLook({ noHBand: !styleLook.noHBand })} />
            </Rows>
            <Rows>
              <Button icon={styleLook.firstColumn ? 'check' : undefined} label="First Column" pressed={styleLook.firstColumn} title="First Column — the first column in its style's look" onClick={() => setLook({ firstColumn: !styleLook.firstColumn })} />
              <Button icon={styleLook.lastColumn ? 'check' : undefined} label="Last Column" pressed={styleLook.lastColumn} title="Last Column — the last column in its style's look" onClick={() => setLook({ lastColumn: !styleLook.lastColumn })} />
              <Button icon={!styleLook.noVBand ? 'check' : undefined} label="Banded Columns" pressed={!styleLook.noVBand} title="Banded Columns — every other column shaded, as the style bands them" onClick={() => setLook({ noVBand: !styleLook.noVBand })} />
            </Rows>
          </Group>
          <Group label="Table Styles">
            <Button tall icon="table" label="Table Styles" title={`Table Styles — Word's own, in this document's theme colours; now ${TABLE_STYLES.find((st) => st.id === tableStyleId)?.name || (tableStyleId ? tableStyleId : 'none')}`} onClick={(e) => menu.open(e, [
              { heading: true, label: 'Plain Tables' },
              ...TABLE_STYLES.filter((st) => st.family === 'plain').map((st) => ({ label: st.name, icon: st.id === tableStyleId ? 'check' : undefined, run: () => tableOp('style', { id: st.id }) })),
              { heading: true, label: 'Grid Tables' },
              ...TABLE_STYLES.filter((st) => st.family === 'grid4').map((st) => ({ label: st.name, icon: st.id === tableStyleId ? 'check' : undefined, run: () => tableOp('style', { id: st.id }) })),
              { heading: true, label: 'List Tables' },
              ...TABLE_STYLES.filter((st) => st.family === 'list4').map((st) => ({ label: st.name, icon: st.id === tableStyleId ? 'check' : undefined, run: () => tableOp('style', { id: st.id }) })),
              '-',
              { label: 'Clear', icon: 'close', disabled: !tableStyleId, run: () => tableOp('style', { id: null }) },
            ])} />
          </Group>
          <Group label="Borders">
            <Rows>
              <Button icon="minus" label={`Line: ${PEN_STYLES.find(([v]) => v === pen.val)?.[1] || 'Single'}`} title="Line Style — the pen's line for the borders it draws" onClick={(e) => menu.open(e, PEN_STYLES.map(([val, label]) => ({ label, icon: pen.val === val ? 'check' : undefined, run: () => setPen({ ...pen, val }) })))} />
              <Button icon="sliders" label={`Weight: ${PEN_WEIGHTS.find(([sz]) => sz === pen.sz)?.[1] || '½ pt'}`} title="Pen Weight — how thick a line the pen draws" onClick={(e) => menu.open(e, PEN_WEIGHTS.map(([sz, label]) => ({ label, icon: pen.sz === sz ? 'check' : undefined, run: () => setPen({ ...pen, sz }) })))} />
              <Button icon="wand" label="Pen Colour" title={`Pen Colour — now ${SHAPE_COLOURS.find(([hex]) => hex === pen.color)?.[1] || 'Automatic'}`} onClick={(e) => menu.open(e, [
                { label: 'Automatic', icon: pen.color === 'auto' ? 'check' : undefined, run: () => setPen({ ...pen, color: 'auto' }) },
                '-',
                ...SHAPE_COLOURS.map(([hex, label]) => ({ label, icon: pen.color === hex ? 'check' : undefined, run: () => setPen({ ...pen, color: hex }) })),
              ])} />
            </Rows>
            <Button tall icon="grid" label="Borders" title="Borders — lines round, between or through the selected cells, in the pen's line" onClick={(e) => menu.open(e, [
              ['bottom', 'Bottom Border'], ['top', 'Top Border'], ['left', 'Left Border'], ['right', 'Right Border'], '-',
              ['none', 'No Border'], ['all', 'All Borders'], ['outside', 'Outside Borders'], ['inside', 'Inside Borders'], '-',
              ['insideH', 'Inside Horizontal Border'], ['insideV', 'Inside Vertical Border'],
            ].map((item) => (item === '-' ? '-' : { label: item[1], run: () => tableOp('borders', { kind: item[0], pen }) })))} />
          </Group>
          <Group label="Shading">
            <Button tall icon="wand" label="Shading" title="Shading — the selected cells' background" onClick={(e) => menu.open(e, [
              { label: 'No Colour', icon: 'close', run: () => tableOp('shading', { fill: null }) },
              '-',
              ...SHAPE_COLOURS.map(([hex, label]) => ({ label, run: () => tableOp('shading', { fill: hex }) })),
            ])} />
          </Group>
        </>
      ) : null}

      {/* ── Table Layout (contextual) ─────────────────────────────────────── */}
      {tab === 'tableLayout' && table ? (
        <>
          <Group label="Draw">
            <Button tall icon="pen" label="Draw Table" pressed={view.tableDraw === 'pen'} title="Draw Table — draw a line down a cell to split it into two columns there, or across it to split it into two rows; Escape puts the pen down" onClick={() => act('tableDraw', 'pen')} />
            <Button tall icon="eraser" label="Eraser" pressed={view.tableDraw === 'eraser'} title="Eraser — click a line between two cells to join them into one; Escape puts the eraser down" onClick={() => act('tableDraw', 'eraser')} />
          </Group>
          <Group label="Table">
            <Button tall icon="settings" label="Properties" title="Table Properties — the table's width, alignment and alt text, the caret's row and column" onClick={() => openDialog('tableProperties')} />
            <Button tall icon="alignCenter" label="Align Table" title="Align Table — the table at the left margin, centred or at the right" onClick={(e) => menu.open(e, [
              { label: 'Left', run: () => tableOp('align', { align: 'left' }) },
              { label: 'Centre', run: () => tableOp('align', { align: 'center' }) },
              { label: 'Right', run: () => tableOp('align', { align: 'right' }) },
            ])} />
            <Button tall icon="grid" label="View Gridlines" pressed={!view.noTableGridlines} title="View Gridlines — a table's faint dashes where it has no lines of its own; on screen only, never printed" onClick={() => act('toggleTableGridlines')} />
          </Group>
          <Group label="Rows & Columns">
            <Button tall icon="minus" label="Delete" title="Delete — the row, the column or the whole table at the caret" onClick={(e) => menu.open(e, [
              { label: 'Delete columns', icon: 'minus', run: () => tableOp('deleteColumn') },
              { label: 'Delete rows', icon: 'minus', run: () => tableOp('deleteRow') },
              { label: 'Delete table', icon: 'trash', run: () => tableOp('deleteTable') },
            ])} />
            <Button tall icon="rowAbove" label="Insert Above" title="Insert a row above the caret's" onClick={() => tableOp('insertRowAbove')} />
            <Button tall icon="rowBelow" label="Insert Below" title="Insert a row below the caret's" onClick={() => tableOp('insertRowBelow')} />
            <Rows>
              <Button icon="colLeft" label="Insert Left" title="Insert a column to the left of the caret's" onClick={() => tableOp('insertColumnLeft')} />
              <Button icon="colRight" label="Insert Right" title="Insert a column to the right of the caret's" onClick={() => tableOp('insertColumnRight')} />
            </Rows>
          </Group>
          <Group label="Merge">
            <Button tall icon="mergeCells" label="Merge Cells" title={acrossCells ? 'Merge the selected cells into one' : 'Merge Cells — select from one cell to another first'} disabled={!acrossCells} onClick={() => tableOp('mergeCells')} />
            <Button tall icon="splitCells" label="Split Cells" title="Split Cells — a merged cell back into the cells it covers, or the caret's cell into columns or rows" onClick={(e) => mergedCaret ? tableOp('splitCell') : menu.open(e, [
              { label: 'Into 2 columns', run: () => tableOp('splitInto', { columns: 2 }) },
              { label: 'Into 3 columns', run: () => tableOp('splitInto', { columns: 3 }) },
              { label: 'Into 2 rows', run: () => tableOp('splitInto', { rows: 2 }) },
              { label: 'Into 2 columns and 2 rows', run: () => tableOp('splitInto', { columns: 2, rows: 2 }) },
              { label: 'A merged cell back into its cells', run: () => tableOp('splitCell') },
            ])} />
            <Button tall icon="splitCells" label="Split Table" title="Split Table — the caret's row starts a table of its own, an empty paragraph between" onClick={() => tableOp('splitTable')} />
          </Group>
          <Group label="Cell Size">
            <Rows>
              <div className="wd-fields" title="The caret's column width, in centimetres">
                <label>Width</label>
                <Input
                  key={`${table.id}:${table.col}:${Math.round((table.gridPx?.[table.col] ?? 0) * 10)}`}
                  type="number" min="0.3" max="50" step="0.1"
                  defaultValue={table.gridPx?.[table.col] ? (table.gridPx[table.col] / CM).toFixed(2) : ''}
                  style={{ width: 64 }}
                  onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }}
                  onBlur={(e) => {
                    const cm = Number(e.currentTarget.value);
                    if (Number.isFinite(cm) && cm > 0 && Math.abs(cm - (table.gridPx[table.col] / CM)) > 0.004) tableOp('columnWidth', { cm });
                  }}
                />
                <span className="wd-field-value" style={{ minWidth: 0 }}>cm</span>
              </div>
            </Rows>
            <Button tall icon="table" label="AutoFit" title="AutoFit — the table to the window, to fixed column widths, or to its contents" onClick={(e) => menu.open(e, [
              { label: 'AutoFit Contents', run: () => tableOp('autoFit', { mode: 'contents' }) },
              { label: 'AutoFit Window', run: () => tableOp('autoFit', { mode: 'window' }) },
              { label: 'Fixed Column Width', run: () => tableOp('autoFit', { mode: 'fixed' }) },
            ])} />
            <Button tall icon="sliders" label="Distribute Columns" title="Distribute Columns — every column the same width, the table as wide as before" onClick={() => tableOp('distributeColumns')} />
          </Group>
          <Group label="Alignment">
            <Button tall icon="rotate" label="Text Direction" title="Text Direction — the selected cells' words across, or turned to read down or up" onClick={(e) => menu.open(e, [
              { label: 'Horizontal', run: () => tableOp('textDirection', { dir: null }) },
              { label: 'Rotate all text 90°', run: () => tableOp('textDirection', { dir: 'down' }) },
              { label: 'Rotate all text 270°', run: () => tableOp('textDirection', { dir: 'up' }) },
            ])} />
            <Button tall icon="crop" label="Cell Margins" title="Cell Margins — the room inside every cell of the table" onClick={(e) => menu.open(e, [
              { label: 'Narrow', run: () => tableOp('cellMargins', { margins: { top: 0, left: 57, bottom: 0, right: 57 } }) },
              { label: 'Normal', run: () => tableOp('cellMargins', { margins: { top: 0, left: 108, bottom: 0, right: 108 } }) },
              { label: 'Wide', run: () => tableOp('cellMargins', { margins: { top: 72, left: 216, bottom: 72, right: 216 } }) },
            ])} />
            <Button tall icon="alignCenter" label="Cell Alignment" title={`Cell Alignment — the words at the top, centre or bottom of the selected cells; now ${cellAlign === 'center' ? 'centre' : cellAlign}`} onClick={(e) => menu.open(e, [
              { label: 'Top', icon: cellAlign === 'top' ? 'check' : undefined, run: () => tableOp('cellVAlign', { v: 'top' }) },
              { label: 'Centre', icon: cellAlign === 'center' ? 'check' : undefined, run: () => tableOp('cellVAlign', { v: 'center' }) },
              { label: 'Bottom', icon: cellAlign === 'bottom' ? 'check' : undefined, run: () => tableOp('cellVAlign', { v: 'bottom' }) },
            ])} />
          </Group>
          <Group label="Direction">
            <Button tall icon="textRtl" label="Right to Left" title="Right to left — the table's columns run from the right, as an Arabic or Hebrew table's do" pressed={table.rtl} onClick={() => tableOp('direction', { rtl: !table.rtl })} />
          </Group>
          <Group label="Data">
            <Button tall icon="sort" label="Sort" title="Sort — the table's rows by the caret's column, its header row kept at the top" onClick={(e) => menu.open(e, [
              { label: 'Sort A to Z by this column', run: () => tableOp('sort', { descending: false }) },
              { label: 'Sort Z to A by this column', run: () => tableOp('sort', { descending: true }) },
              '-',
              { label: 'A to Z, the first row a header', run: () => tableOp('sort', { descending: false, header: true }) },
              { label: 'Z to A, the first row a header', run: () => tableOp('sort', { descending: true, header: true }) },
            ])} />
            <Button tall icon="file" label="Convert to Text" title="Convert to Text — each row a paragraph, its cells between tabs or commas" onClick={(e) => menu.open(e, [
              { label: 'Separated by tabs', run: () => tableOp('toText', { separator: 'tab' }) },
              { label: 'Separated by commas', run: () => tableOp('toText', { separator: 'comma' }) },
              { label: 'Each cell its own paragraph', run: () => tableOp('toText', { separator: 'paragraph' }) },
            ])} />
            <Button tall icon="formula" label="Formula" title="Formula — a field in the caret's cell worked out from the numbers above or to the left of it, or from cells like B2 and B2:B4, as Word's =SUM(ABOVE)" onClick={(e) => menu.open(e, [
              { label: 'Sum above', run: () => tableOp('formula', { formula: '=SUM(ABOVE)' }) },
              { label: 'Sum left', run: () => tableOp('formula', { formula: '=SUM(LEFT)' }) },
              { label: 'Average above', run: () => tableOp('formula', { formula: '=AVERAGE(ABOVE)', format: '#,##0.00' }) },
              { label: 'Count above', run: () => tableOp('formula', { formula: '=COUNT(ABOVE)' }) },
              { label: 'Largest above', run: () => tableOp('formula', { formula: '=MAX(ABOVE)' }) },
              { label: 'Smallest above', run: () => tableOp('formula', { formula: '=MIN(ABOVE)' }) },
              { label: 'Your own formula…', run: () => openDialog('tableFormula') },
              { label: 'Update all formulas', run: () => dispatch({ op: 'updateTableFormulas' }) },
            ])} />
            <Button tall icon="refresh" label="Repeat Header Rows" pressed={Boolean(caretCell?.rowHeader)} title="Repeat Header Rows — the rows from the top through this one drawn again at the top of every page the table runs onto" onClick={() => tableOp('headerRows', { on: !caretCell?.rowHeader })} />
          </Group>
        </>
      ) : null}

      {/* ── 3D Model (contextual) ─────────────────────────────────────────── */}
      {tab === 'model3d' && pickedModel ? (() => {
        const v = pickedModel.view || {};
        const at = (view) => Math.abs((v.yaw || 0) - view.yaw) < 0.5 && Math.abs((v.pitch || 0) - view.pitch) < 0.5;
        const rows = [MODEL_VIEWS.slice(0, 4), MODEL_VIEWS.slice(4, 7), MODEL_VIEWS.slice(7)];
        return (
          <>
            <Group label="3D Model Views">
              {rows.map((row, i) => (
                <Rows key={i}>
                  {row.map(([key, label, view]) => <Button key={key} icon={at(view) ? 'check' : 'shape'} label={label} className="wd-model3d-view" data-view={key} pressed={at(view)} title={`${label} — the model turned to show it from there`} onClick={() => act('model3dView', view)} />)}
                </Rows>
              ))}
            </Group>
            <Group label="Turn">
              <Rows>
                <Button icon="chevronLeft" label="Turn Left" className="wd-model3d-turn-left" title="Turn Left — fifteen degrees about its upright" onClick={() => act('model3dView', { turn: { yaw: -15, pitch: 0 } })} />
                <Button icon="chevronRight" label="Turn Right" title="Turn Right — fifteen degrees about its upright" onClick={() => act('model3dView', { turn: { yaw: 15, pitch: 0 } })} />
              </Rows>
              <Rows>
                <Button icon="chevronUp" label="Tip Back" title="Tip Back — fifteen degrees, the top towards you" onClick={() => act('model3dView', { turn: { yaw: 0, pitch: 15 } })} />
                <Button icon="chevronDown" label="Tip Forward" title="Tip Forward — fifteen degrees, the bottom towards you" onClick={() => act('model3dView', { turn: { yaw: 0, pitch: -15 } })} />
              </Rows>
            </Group>
            <Group label="Adjust">
              <Button tall icon="undo" label="Reset 3D Model" className="wd-model3d-reset" title="Reset 3D Model — back to the view it was put in at" onClick={() => act('model3dView', 'reset')} />
            </Group>
          </>
        );
      })() : null}

      {/* ── Picture Format (contextual) ───────────────────────────────────── */}
      {tab === 'pictureFormat' ? (
        <>
          {arrange}
          <Group label="Size">
            <Rows>
              {sizeBox('Height', 'h')}
              {sizeBox('Width', 'w')}
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── References ───────────────────────────────────────────────────── */}
      {tab === 'references' ? (
        <>
          <Group label="Table of Contents">
            <Button tall icon="listBullet" label="Table of Contents" title="Table of Contents — a live field built from the headings, with page numbers" onClick={() => act('tableOfContents')} />
            <Button icon="plus" label="Add Text" title="Make this paragraph a heading, so it appears in the table of contents" onClick={(e) =>
              menu.open(e, [
                { label: 'Do not show in table of contents', run: () => para({ styleId: null }) },
                { label: 'Level 1 (Heading 1)', run: () => para({ styleId: 'Heading1' }) },
                { label: 'Level 2 (Heading 2)', run: () => para({ styleId: 'Heading2' }) },
                { label: 'Level 3 (Heading 3)', run: () => para({ styleId: 'Heading3' }) },
              ])
            } />
            <Button icon="refresh" label="Update Table" disabled={!model?.tableOfContents} title="Update Table — rebuild the entries from the current headings and pages" onClick={() => act('updateTableOfContents')} />
            <Button icon="close" label="Remove Table of Contents" disabled={!model?.tableOfContents} title="Remove Table of Contents" onClick={() => act('removeTableOfContents')} />
          </Group>
          <Group label="Footnotes">
            <Button tall icon="file" label="Insert Footnote" title="A raised number at the caret, and its words under the body" onClick={() => act('insertNote', 'footnote')} />
            <Button icon="file" label="Insert Endnote" title="A raised number at the caret, and its words at the end of the document" onClick={() => act('insertNote', 'endnote')} />
            <Button icon="chevronDown" label="Next Footnote" title="Go to the next footnote reference" onClick={() => act('nextNote', 1)} />
            <Button icon="eye" label="Show Notes" title="Scroll to the footnotes under the body" onClick={() => act('showNotes')} />
          </Group>
          {references ? <CitationsGroup refs={references} menu={menu} /> : null}
          <Group label="Captions">
            <Button tall icon="textbox" label="Insert Caption" title="Insert Caption — a label, a running number kept live by Update Fields, and your own words" onClick={() => openDialog('caption')} />
            {references ? captionsExtra(references, <Button icon="link" label="Cross-reference" title="Cross-reference — a REF field to a bookmark, its words kept live by Update Fields" onClick={() => openDialog('crossReference')} />) : <Button icon="link" label="Cross-reference" title="Cross-reference — a REF field to a bookmark, its words kept live by Update Fields" onClick={() => openDialog('crossReference')} />}
          </Group>
          <Group label="Fields">
            <Button tall icon="refresh" label="Update Fields" title="Update Fields (F9) — refresh every cross-reference to its bookmark's current words" onClick={() => commands['field.update']?.run?.()} />
          </Group>
          {references ? <IndexGroup act={references.act} hasIndex={Boolean(references.info?.index)} /> : null}
          {references ? <ToaGroup act={references.act} hasTable={Boolean(references.info?.toa?.length)} /> : null}
        </>
      ) : null}

      {/* ── Mailings ─────────────────────────────────────────────────────── */}
      {tab === 'mailings' && mailings ? (
        <MailingsTab mm={mailings} menu={menu} />
      ) : null}

      {/* ── Review ───────────────────────────────────────────────────────── */}
      {tab === 'review' ? (
        <>
          <Group label="Proofing">
            <Button tall icon="check" label="Spelling" pressed={review?.pane === 'editor'} title="Spelling (F7) — check the whole document from the caret in the Editor pane: body, tables, text boxes, notes, headers and footers" onClick={() => review?.startSpelling()} />
            <Button tall icon="word" label="Thesaurus" pressed={review?.pane === 'thesaurus'} title="Thesaurus (Shift+F7) — words of like meaning for the word at the caret, one of them put in its place" onClick={() => review?.openThesaurus()} />
            <Button tall icon="listNumber" label="Word Count" onClick={() => openDialog('wordCount')} />
          </Group>
          <Group label="Speech">
            <Button tall icon="volume" label="Read Aloud" pressed={Boolean(view.reading)} onClick={() => act('readAloud')} />
          </Group>
          <Group label="Accessibility">
            <Button tall icon="shield" label="Check Accessibility" pressed={review?.pane === 'accessibility'} title="Check Accessibility — alt text, headings, tables, contrast, links and spacing, with a fix for each" onClick={() => review?.openAccessibility()} />
            <Button icon="textbox" label="Alt Text" disabled={!picked} title={picked ? 'Alt Text — describe the selected picture for people who cannot see it' : 'Alt Text — click a picture first, then describe it'} onClick={() => picked && review?.openAltText({ block: picked.block, image: picked.image })} />
          </Group>
          <Group label="Language">
            <Soon tall icon="globe" label="Translate" why="Translation is a network service; this suite makes no requests it has not declared." />
            <Button tall icon="globe" label="Language" title="Language — mark the selected words as a language, or as not to be checked, for Spelling here and proofing in Office" onClick={() => act('language')} />
          </Group>
          <Group label="Comments">
            <Button tall icon="reply" label="New Comment" onClick={() => openDialog('comment')} />
            <Button
              tall
              icon="trash"
              label="Delete"
              disabled={!comments.length}
              title={comments.length ? 'Delete — the comment on the caret\'s paragraph, or every comment in the document' : 'Delete — this document has no comments'}
              onClick={(e) => menu.open(e, [
                { label: 'Delete', icon: 'trash', disabled: !comments.some((c) => c.blockIndex === (model?.selection?.focus?.block ?? -1)), run: () => act('deleteComment') },
                { label: 'Delete All Comments in Document', icon: 'trash', run: () => act('deleteComment', 'all') },
              ])}
            />
            <Button icon="chevronUp" label="Previous" disabled={!comments.length} onClick={() => act('comment', -1)} />
            <Button icon="chevronDown" label="Next" disabled={!comments.length} onClick={() => act('comment', 1)} />
            <Button icon="listBullet" label={comments.length ? `Show (${comments.length})` : 'Show Comments'} disabled={!comments.length} onClick={() => openDialog('comments')} />
          </Group>
          <Group label="Tracking">
            <Button tall icon="eye" label="Track Changes" pressed={Boolean(model?.trackRevisions)} disabled={Boolean(model?.protection?.lockedTracking)} title={model?.protection?.lockedTracking ? 'Track Changes — locked on: the document is protected for tracked changes' : 'Track Changes — record every insertion and deletion as w:ins / w:del while you edit'} onClick={() => act('toggleTrackChanges')} />
            <Button icon="eye" label={MARKUP_LABELS[view.markupMode] || 'Simple Markup'} title="Display for Review — how tracked changes are shown" onClick={(e) => menu.open(e, [
              { label: 'All Markup', run: () => act('markupMode', 'all') },
              { label: 'Simple Markup', run: () => act('markupMode', 'simple') },
              { label: 'No Markup', run: () => act('markupMode', 'final') },
              { label: 'Original', run: () => act('markupMode', 'original') },
            ])} />
            <Button icon="list" label="Reviewing Pane" title="Every tracked change in this document" disabled={!model?.blocks?.some((b) => b.tracked)} onClick={() => openDialog('tracked')} />
          </Group>
          <Group label="Changes">
            <Button tall icon="check" label="Accept" disabled={!trackedHere || Boolean(model?.protection?.lockedTracking)} title="Accept — keep this change" onClick={(e) => menu.open(e, [
              { label: 'Accept This Change', run: () => act('acceptChanges') },
              { label: 'Accept All Changes', run: () => act('acceptChanges', 'all') },
            ])} />
            <Button tall icon="close" label="Reject" disabled={!trackedHere || Boolean(model?.protection?.lockedTracking)} title="Reject — undo this change" onClick={(e) => menu.open(e, [
              { label: 'Reject This Change', run: () => act('rejectChanges') },
              { label: 'Reject All Changes', run: () => act('rejectChanges', 'all') },
            ])} />
            <Button icon="chevronUp" label="Previous" disabled={!anyTracked} title="Previous tracked change" onClick={() => act('nextChange', -1)} />
            <Button icon="chevronDown" label="Next" disabled={!anyTracked} title="Next tracked change" onClick={() => act('nextChange', 1)} />
          </Group>
          <Group label="Compare">
            <Button tall icon="copy" label="Compare" title="Compare — two versions of a document side by side, as a new document with what changed marked as revisions" onClick={() => act('compare')} />
          </Group>
          <Group label="Protect">
            <Button tall icon="lock" label="Restrict Editing" pressed={Boolean(view.restrict)} title={model?.protection?.enforced ? 'Restrict Editing — the document is protected; the pane says what you may do and stops protection' : 'Restrict Editing — limit formatting and editing, with exceptions, and enforce it with an optional password'} onClick={() => act('restrictPane')} />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
      {tab === 'immersive' && view.immersive ? (
        <>
          <Group label="Immersive Reader">
            <Button
              tall
              icon="alignJustify"
              label="Column Width"
              title="Column Width — how long the lines are"
              onClick={(e) => menu.open(e, IR_WIDTHS.map(([k, label]) => ({ label, icon: view.immersive.width === k ? 'check' : undefined, run: () => act('immersiveSet', { width: k }) })))}
            />
            <Button
              tall
              icon="contrast"
              label="Page Color"
              title="Page Color — the colour behind the words"
              onClick={(e) => menu.open(e, IR_COLOURS.map(([k, label]) => ({ label, icon: view.immersive.colour === k ? 'check' : undefined, run: () => act('immersiveSet', { colour: k }) })))}
            />
            <Button
              tall
              icon="eye"
              label="Line Focus"
              title="Line Focus — one, three or five lines in view at a time; the arrow keys move the words through them"
              onClick={(e) => menu.open(e, IR_FOCUS.map(([n, label]) => ({ label, icon: view.immersive.focus === n ? 'check' : undefined, run: () => act('immersiveSet', { focus: n }) })))}
            />
            <Button tall icon="sliders" label="Text Spacing" pressed={Boolean(view.immersive.spacing)} title="Text Spacing — more room between letters, words and lines" onClick={() => act('immersiveSet', { spacing: !view.immersive.spacing })} />
            <Button tall icon="scissors" label="Syllables" pressed={Boolean(view.immersive.syllables)} title="Syllables — long words shown in their syllables, a dot between each" onClick={() => act('immersiveSet', { syllables: !view.immersive.syllables })} />
            <Button tall icon="volume" label="Read Aloud" pressed={Boolean(view.reading)} onClick={() => act('readAloud')} />
          </Group>
          <Group label="Close">
            <Button tall icon="close" label="Close Immersive Reader" onClick={() => act('immersiveClose')} />
          </Group>
        </>
      ) : null}

      {tab === 'view' ? (
        <>
          <Group label="Views">
            <Button tall icon="eye" label="Read Mode" pressed={view.mode === 'read'} onClick={() => act('mode', 'read')} />
            <Button tall icon="file" label="Print Layout" pressed={!view.mode || view.mode === 'print'} onClick={() => act('mode', 'print')} />
            <Button tall icon="globe" label="Web Layout" pressed={view.mode === 'web'} onClick={() => act('mode', 'web')} />
            <Button icon="listBullet" label="Outline" pressed={view.mode === 'outline'} onClick={() => act('mode', 'outline')} />
            <Button icon="textbox" label="Draft" pressed={view.mode === 'draft'} onClick={() => act('mode', 'draft')} />
          </Group>
          <Group label="Immersive">
            <Button tall icon="maximize" label="Focus" pressed={Boolean(view.focus)} title="Just the page, full screen" onClick={() => act('focus')} />
            <Button tall icon="eye" label="Immersive Reader" pressed={Boolean(view.immersive)} title="Immersive Reader — the words in a column of their own, with line focus, wider spacing, syllables and a page colour that is easier to read" onClick={() => act(view.immersive ? 'immersiveClose' : 'immersive')} />
          </Group>
          <Group label="Show">
            <Button icon="minus" label="Ruler" pressed={Boolean(view.ruler)} onClick={() => act('toggleRuler')} />
            <Button icon="grid" label="Gridlines" pressed={Boolean(view.gridlines)} title="Gridlines — a quarter-inch grid over the page, to line drawings up by eye; on screen only" onClick={() => act('toggleGridlines')} />
            <Button icon="list" label="Navigation Pane" pressed={Boolean(view.navigation)} title="Headings, to move around a long document" onClick={() => act('toggleNavigation')} />
          </Group>
          <Group label="Zoom">
            <Button tall icon="find" label="Zoom" onClick={(e) => menu.open(e, [50, 75, 100, 125, 150, 200].map((z) => ({ label: `${z}%`, run: () => act('zoom', z / 100) })))} />
            <Button icon="check" label="100%" onClick={() => act('zoom', 1)} />
            <Button icon="file" label="One Page" onClick={() => act('zoom', 'page')} />
            <Button icon="copy" label="Multiple Pages" onClick={() => act('zoom', 'pages')} />
            <Button icon="maximize" label="Page Width" onClick={() => act('zoom', 'width')} />
          </Group>
          <Group label="Window">
            <Button tall icon="new" label="New Window" title="This document again, in another window" onClick={() => act('newWindow')} />
            <Button icon="grid" label="Arrange All" title="Arrange All — every Documents window, one above another" onClick={() => act('arrange', 'stack')} />
            <Button icon="minus" label="Split" pressed={Boolean(view.split)} title={view.split ? 'Remove Split — the document in one pane again' : 'Split — the document in two panes, one over the other, each scrolled on its own'} onClick={() => act('toggleSplit')} />
            <Button icon="copy" label="Side by Side" title="View Side by Side — this document and the one before it, half the screen each" onClick={() => act('arrange', 'sideBySide')} />
          </Group>
          <Group label="Macros">
            <Soon tall icon="settings" label="Macros" why="A macro engine is not built, and VBA in a file is preserved untouched rather than run." />
          </Group>
          <Group label="Properties">
            <Button tall icon="info" label="Properties" onClick={() => openDialog('properties')} />
          </Group>
        </>
      ) : null}

      {/* ── Help ─────────────────────────────────────────────────────────── */}
      {tab === 'help' ? (
        <>
          <Group label="Help">
            <Button tall icon="info" label="Help" onClick={() => act('open', 'help')} />
            <Button tall icon="reply" label="Feedback" onClick={() => act('open', 'contact')} />
            <Button tall icon="globe" label="What's New" onClick={() => act('open', 'releases')} />
            <Button tall icon="download" label="Check for Updates" onClick={() => act('open', 'about')} />
          </Group>
          <Group label="Shortcuts">
            <Button tall icon="listBullet" label="Keyboard Shortcuts" onClick={() => openDialog('shortcuts')} />
          </Group>
        </>
      ) : null}

      {/* ── PDF ──────────────────────────────────────────────────────────── */}
      {tab === 'pdf' ? (
        <>
          <Group label="Create">
            <Button tall icon="pdf" label="Create a PDF" onClick={() => exportAs('pdf')} />
            <Button tall icon="export" label="Markdown" onClick={() => exportAs('md')} />
            <Button tall icon="export" label="Plain Text" onClick={() => exportAs('txt')} />
            <Button tall icon="export" label="Web Page" onClick={() => exportAs('html')} />
          </Group>
          <Group label="Print">
            <Button tall icon="print" label="Print" onClick={() => commands['file.print']?.run?.()} />
          </Group>
          <Group label="File">
            <Button tall icon="new" label="New" onClick={() => shell.win.create({ app: 'word' })} />
            <Button tall icon="open" label="Open" onClick={openFile} />
            <Button tall icon="save" label="Save" onClick={() => save(false)} />
            <Button tall icon="save" label="Save As" onClick={() => save(true)} />
          </Group>
        </>
      ) : null}
    </Ribbon>
  );
}

/** How the style gallery previews a style: enough to tell them apart at a glance. */
const STYLE_LOOK = {
  Title: { fontSize: 16, fontWeight: 600, fontFamily: 'Cambria, Georgia, serif' },
  Heading1: { fontSize: 14, fontWeight: 600, color: '#2b5fd9' },
  Heading2: { fontSize: 13, fontWeight: 600, color: '#2b5fd9' },
  Heading3: { fontSize: 12, fontWeight: 600, color: '#1f3864' },
  Subtitle: { fontSize: 12, fontStyle: 'italic', color: '#767171' },
  Quote: { fontStyle: 'italic' },
};
