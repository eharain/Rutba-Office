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
import { Ribbon, Group, Rows, Button, Separator, Select, Input, t, msg } from '@rutba/office-ui';
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
/** Word's three Arabic justifications, each stretching the letters with kashida a little more. */
const KASHIDA = [['lowKashida', msg('Justify Low')], ['mediumKashida', msg('Justify Medium')], ['highKashida', msg('Justify High')]];

// Immersive Reader's choices, as Word names them.
const IR_WIDTHS = [['veryNarrow', t('Very Narrow')], ['narrow', t('Narrow')], ['moderate', t('Moderate')], ['wide', t('Wide')]];
const IR_COLOURS = [['none', t('None')], ['sepia', t('Sepia')], ['inverse', t('Inverse')]];
const IR_FOCUS = [[0, t('None')], [1, t('One Line')], [3, t('Three Lines')], [5, t('Five Lines')]];

const TEXT_COLOURS = [
  [null, t('Automatic')], ['000000', t('Black')], ['444444', t('Dark grey')], ['767171', t('Grey')], ['FFFFFF', t('White')],
  ['C00000', t('Dark red')], ['E03131', t('Red')], ['E08B2B', t('Orange')], ['E0A800', t('Gold')],
  ['0F9D58', t('Green')], ['0D8F6F', t('Teal')], ['2B5FD9', t('Blue')], ['1F3864', t('Dark blue')], ['7B5CD6', t('Purple')],
];
const HIGHLIGHTS = [
  [null, t('No highlight')], ['yellow', t('Yellow')], ['green', t('Bright green')], ['cyan', t('Turquoise')], ['magenta', t('Pink')],
  ['blue', t('Blue')], ['red', t('Red')], ['darkBlue', t('Dark blue')], ['darkCyan', t('Teal')], ['darkGreen', t('Green')],
  ['darkMagenta', t('Violet')], ['darkRed', t('Dark red')], ['darkYellow', t('Dark yellow')], ['darkGray', t('Grey 50%')], ['lightGray', t('Grey 25%')], ['black', t('Black')],
];
const SPACING = [[1, '1.0'], [1.15, '1.15'], [1.5, '1.5'], [2, '2.0'], [2.5, '2.5'], [3, '3.0']];
/** Word's half-point rule, for the Borders menu. */
const LINE = { style: 'single', widthPx: 1, colour: null, spacePt: 1 };
const SHADES = [['#FFF2CC', t('Light yellow')], ['#DEEBF7', t('Light blue')], ['#E2EFDA', t('Light green')], ['#FCE4D6', t('Light orange')], ['#EDEDED', t('Light grey')], [null, t('No colour')]];
const BORDERS = [
  [t('Bottom border'), { bottom: LINE }], [t('Top border'), { top: LINE }], [t('Left border'), { left: LINE }], [t('Right border'), { right: LINE }],
  [t('Outside borders'), { top: LINE, bottom: LINE, left: LINE, right: LINE }], [t('No border'), null],
];
/** Page Borders: Word's Box setting in four lines, measured 24 pt in from the page edge as Word measures it. */
const FRAME = (style, widthPx) => ({ style, widthPx, colour: null, spacePt: 24 });
const BOX = (line) => ({ offsetFrom: 'page', top: line, left: line, bottom: line, right: line });
const PAGE_BORDERS = [
  [t('Box — thin line'), BOX(FRAME('single', 1))], [t('Box — thick line'), BOX(FRAME('single', 3))],
  [t('Box — double line'), BOX(FRAME('double', 3))], [t('Box — dashed line'), BOX(FRAME('dashed', 1))],
  [t('No page border'), null],
];
/** Watermark: Word's own stock words, then a custom one. */
const WATERMARKS = [t('DRAFT'), t('CONFIDENTIAL'), t('SAMPLE'), t('DO NOT COPY'), t('URGENT')];
const PAGE_SIZES = [['A4', t('A4 — 21 × 29.7 cm')], ['Letter', t('Letter — 8.5 × 11 in')], ['Legal', t('Legal — 8.5 × 14 in')]];
/** Review → Display for Review: the button's own label per mode. */
const MARKUP_LABELS = { all: t('All Markup'), simple: t('Simple Markup'), final: t('No Markup'), original: t('Original') };
const MARGINS = [['normal', t('Normal — 2.54 cm all round')], ['narrow', t('Narrow — 1.27 cm all round')], ['wide', t('Wide — 5.08 cm at the sides')]];
const SHAPES = [
  ['rect', t('Rectangle')], ['roundRect', t('Rounded rectangle')], ['ellipse', t('Oval')], ['triangle', t('Triangle')], ['diamond', t('Diamond')],
  ['rightArrow', t('Arrow: right')], ['leftArrow', t('Arrow: left')], ['upArrow', t('Arrow: up')], ['downArrow', t('Arrow: down')],
  ['pentagon', t('Pentagon')], ['hexagon', t('Hexagon')], ['star5', t('Star: 5 points')], ['plus', t('Cross')], ['chevron', t('Chevron')],
  ['parallelogram', t('Parallelogram')], ['trapezoid', t('Trapezoid')], ['line', t('Line')],
];
const CHARTS = [['column', t('Column')], ['bar', t('Bar')], ['line', t('Line')], ['pie', t('Pie')]];
const CASES = [['sentence', t('Sentence case.')], ['lower', t('lowercase')], ['upper', t('UPPERCASE')], ['title', t('Capitalize Each Word')], ['toggle', t('tOGGLE cASE')]];
/** The "A" button's glow colours — Word's own accent palette, 4pt radius. */
const GLOWS = [['FFC000', t('Glow: gold')], ['4472C4', t('Glow: blue')], ['70AD47', t('Glow: green')], ['FF0000', t('Glow: red')]];

/** Shape Fill and Shape Outline: Word's standard colours, in the suite's own names. */
const SHAPE_COLOURS = [
  ['FFFFFF', t('White')], ['000000', t('Black')], ['E7E6E6', t('Light grey')], ['767171', t('Grey')], ['1F3864', t('Dark blue')],
  ['2B5FD9', t('Blue')], ['DEEBF7', t('Light blue')], ['0D8F6F', t('Teal')], ['0F9D58', t('Green')], ['E2EFDA', t('Light green')],
  ['E0A800', t('Gold')], ['FFF2CC', t('Light yellow')], ['E08B2B', t('Orange')], ['C00000', t('Dark red')], ['7B5CD6', t('Purple')],
];
/** Shape Outline → Weight, in Word's points and the pixels the engine writes. */
const WEIGHTS = [[t('½ pt'), 0.67], [t('1 pt'), 1.33], [t('1½ pt'), 2], [t('2¼ pt'), 3], [t('3 pt'), 4], [t('6 pt'), 8]];
/** A text box's Margins, Word's four: its insets in px, left/top/right/bottom. */
const BOX_MARGINS = [
  [t('None'), { l: 0, t: 0, r: 0, b: 0 }],
  [t('Narrow'), { l: 4.8, t: 4.8, r: 4.8, b: 4.8 }],
  [t('Normal'), { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }],
  [t('Wide'), { l: 14.4, t: 14.4, r: 14.4, b: 14.4 }],
];
/** Home → Multilevel List's library, as Word's gallery shows each: its first three levels. */
const MULTILEVEL = [
  ['outline', '1.  1.1.  1.1.1.'],
  ['outlineParen', '1)  a)  i)'],
  ['outlineRoman', 'I.  A.  1.'],
  ['legal', t('Article I.  Section 1.01  (a)')],
  ['bullet', '•  ○  ▪'],
];

/** Insert → WordArt: four styles of our own, made of the text effects the engine writes. */
const WORDART = [
  [t('Fill: blue, shadow'), { colour: '2B5FD9', effects: { shadow: true } }],
  [t('Outline: blue'), { colour: '2B5FD9', effects: { outline: true } }],
  [t('Fill: gold, glow'), { colour: 'C98A00', effects: { glow: { colour: 'FFC000', radiusPt: 4 } } }],
  [t('Fill: black, shadow'), { colour: '262626', effects: { shadow: true } }],
];
/** Arrange → Position's nine places with the words wrapped round, by `${vertical}-${horizontal}`. */
const POSITIONS = {
  'top-left': t('Top Left'), 'top-center': t('Top Centre'), 'top-right': t('Top Right'),
  'middle-left': t('Middle Left'), 'middle-center': t('Middle Centre'), 'middle-right': t('Middle Right'),
  'bottom-left': t('Bottom Left'), 'bottom-center': t('Bottom Centre'), 'bottom-right': t('Bottom Right'),
};
/** 1 cm in CSS px, for the Size boxes. */
const CM = 96 / 2.54;

/** Table Design's pen: Word's line styles, and its weights in eighths of a point. */
const PEN_STYLES = [['single', t('Single')], ['double', t('Double')], ['dotted', t('Dotted')], ['dashed', t('Dashed')], ['thick', t('Thick')]];
const PEN_WEIGHTS = [[2, t('¼ pt')], [4, t('½ pt')], [6, t('¾ pt')], [8, t('1 pt')], [12, t('1½ pt')], [18, t('2¼ pt')], [24, t('3 pt')], [36, t('4½ pt')], [48, t('6 pt')]];

/** A control that is drawn where Word draws it, and says why it is not live. */
const Soon = ({ icon, label, tall, why }) => (
  <Button tall={tall} icon={icon} label={label} disabled title={t('{label} — not built yet. {why}', { label, why })} />
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
  const needDrawing = t('Click a picture, a shape or a text box first');
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
          title={t('{label} (cm)', { label })}
          style={{ width: 64 }}
          onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }}
          onBlur={(e) => {
            const v = Number(e.target.value);
            if (!(v > 0) || Math.abs(v - cm) < 0.005) return;
            act('drawingSize', key === 'w' ? { widthPx: Math.round(v * CM) } : { heightPx: Math.round(v * CM) });
          }}
        />
        <span className="wd-field-value" style={{ minWidth: 0 }}>{t('cm')}</span>
      </div>
    );
  };
  const arrange = (
    <Group label={t('Arrange')}>
      <Button tall icon="grid" label={t('Position')} disabled={!hasDrawing} title={hasDrawing ? t('Position — where the drawing sits, with the words round it') : needDrawing} onClick={(e) => menu.open(e, [
        { label: t('Left, words round it'), icon: 'alignLeft', run: () => act('position', 'left') },
        { label: t('Centre, words above and below'), icon: 'alignCenter', run: () => act('position', 'center') },
        { label: t('Right, words round it'), icon: 'alignRight', run: () => act('position', 'right') },
        { heading: true, label: t('With text wrapping') },
        ...['top', 'middle', 'bottom'].flatMap((v) => ['left', 'center', 'right'].map((h) => ({ label: POSITIONS[`${v}-${h}`], run: () => act('position', `${v}-${h}`) }))),
      ])} />
      <Button tall icon="picture" label={t('Wrap Text')} disabled={!hasDrawing} title={hasDrawing ? t('Wrap Text — how the words treat the drawing') : needDrawing} onClick={(e) => menu.open(e, [
        { label: t('In line with text'), icon: drawing && !drawing.anchored ? 'check' : undefined, run: () => act('wrap', 'inline') },
        { label: t('Square'), icon: drawing?.anchored && drawing.wrap === 'square' ? 'check' : undefined, run: () => act('wrap', 'square') },
        { label: t('Tight'), icon: drawing?.anchored && drawing.wrap === 'tight' ? 'check' : undefined, run: () => act('wrap', 'tight') },
        { label: t('Through'), icon: drawing?.anchored && drawing.wrap === 'through' ? 'check' : undefined, run: () => act('wrap', 'through') },
        { label: t('Top and bottom'), icon: drawing?.anchored && drawing.wrap === 'topAndBottom' ? 'check' : undefined, run: () => act('wrap', 'topAndBottom') },
        '-',
        { label: t('Behind text'), icon: drawing?.anchored && drawing.wrap === 'none' && drawing.behind ? 'check' : undefined, run: () => act('wrap', 'behind') },
        { label: t('In front of text'), icon: drawing?.anchored && drawing.wrap === 'none' && !drawing.behind ? 'check' : undefined, run: () => act('wrap', 'front') },
      ])} />
      <Button tall icon="chevronUp" label={t('Bring Forward')} disabled={!floatingSel.length} title={floatingSel.length ? t('Bring Forward — one place nearer the front; also Bring to Front and Bring in Front of Text') : t('Bring Forward — select a floating drawing first')} onClick={(e) => menu.open(e, [
        { label: t('Bring Forward'), icon: 'chevronUp', run: () => act('order', 'forward') },
        { label: t('Bring to Front'), run: () => act('order', 'front') },
        { label: t('Bring in Front of Text'), run: () => act('order', 'inFront') },
      ])} />
      <Button tall icon="chevronDown" label={t('Send Backward')} disabled={!floatingSel.length} title={floatingSel.length ? t('Send Backward — one place further back; also Send to Back and Send Behind Text') : t('Send Backward — select a floating drawing first')} onClick={(e) => menu.open(e, [
        { label: t('Send Backward'), icon: 'chevronDown', run: () => act('order', 'backward') },
        { label: t('Send to Back'), run: () => act('order', 'back') },
        { label: t('Send Behind Text'), run: () => act('order', 'behind') },
      ])} />
      <Rows>
        <Button icon="list" label={t('Selection Pane')} pressed={Boolean(view.selectionPane)} title={t('Selection Pane — every drawing listed: select, hide, rename, reorder')} onClick={() => act('selectionPane')} />
        <>
          <Button icon="alignLeft" label={t('Align')} disabled={!floatingSel.length} title={floatingSel.length ? (alignMode === 'selected' ? t('Align — to the selected drawings') : alignMode === 'page' ? t('Align — to the page') : t('Align — to the margin')) : t('Align — select a floating drawing first')} onClick={(e) => menu.open(e, [
            { label: t('Align Left'), icon: 'alignLeft', run: () => act('align', { edge: 'left', to: alignMode }) },
            { label: t('Align Centre'), icon: 'alignCenter', run: () => act('align', { edge: 'center', to: alignMode }) },
            { label: t('Align Right'), icon: 'alignRight', run: () => act('align', { edge: 'right', to: alignMode }) },
            { label: t('Align Top'), run: () => act('align', { edge: 'top', to: alignMode }) },
            { label: t('Align Middle'), run: () => act('align', { edge: 'middle', to: alignMode }) },
            { label: t('Align Bottom'), run: () => act('align', { edge: 'bottom', to: alignMode }) },
            '-',
            { label: t('Distribute Horizontally'), disabled: floatingSel.length < 3, title: floatingSel.length < 3 ? t('Select three or more drawings to distribute') : undefined, run: () => act('distribute', { axis: 'horizontal' }) },
            { label: t('Distribute Vertically'), disabled: floatingSel.length < 3, title: floatingSel.length < 3 ? t('Select three or more drawings to distribute') : undefined, run: () => act('distribute', { axis: 'vertical' }) },
            '-',
            { label: t('Align to Page'), icon: alignMode === 'page' ? 'check' : undefined, run: () => setAlignTo('page') },
            { label: t('Align to Margin'), icon: alignMode === 'margin' ? 'check' : undefined, run: () => setAlignTo('margin') },
            { label: t('Align Selected Objects'), icon: alignMode === 'selected' ? 'check' : undefined, disabled: floatingSel.length < 2, run: () => setAlignTo('selected') },
          ])} />
          <Button icon="grid" label={t('Group')} disabled={floatingSel.length < 2 && !isGroup} title={floatingSel.length >= 2 ? t('Group — the selected drawings as one') : isGroup ? t('Group — ungroup this group') : t('Group — select two or more floating drawings (Shift+click adds one)')} onClick={(e) => menu.open(e, [
            { label: t('Group'), icon: 'grid', disabled: floatingSel.length < 2, run: () => act('group') },
            { label: t('Ungroup'), disabled: !isGroup, run: () => act('ungroup') },
          ])} />
        </>
        <Button icon="rotate" label={t('Rotate')} disabled={!hasDrawing} title={hasDrawing ? t('Rotate — turn or flip the drawing') : needDrawing} onClick={(e) => menu.open(e, [
          { label: t('Rotate Right 90°'), icon: 'rotate', run: () => act('rotate', 90) },
          { label: t('Rotate Left 90°'), run: () => act('rotate', -90) },
          { label: t('Flip Vertical'), icon: 'flip', run: () => act('rotate', 'flipV') },
          { label: t('Flip Horizontal'), run: () => act('rotate', 'flipH') },
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
        { id: 'home', label: t('Home') },
        { id: 'insert', label: t('Insert') },
        { id: 'draw', label: t('Draw') },
        { id: 'design', label: t('Design') },
        { id: 'layout', label: t('Layout') },
        { id: 'references', label: t('References') },
        { id: 'mailings', label: t('Mailings') },
        { id: 'review', label: t('Review') },
        { id: 'view', label: t('View') },
        { id: 'help', label: t('Help') },
        { id: 'pdf', label: 'PDF' }, // words-ok: a file format's name
        // The contextual tab, as Word's: there while a drawing is selected.
        ...(formatTab === 'shapeFormat' ? [{ id: 'shapeFormat', label: t('Shape Format') }] : []),
        ...(formatTab === 'pictureFormat' ? [{ id: 'pictureFormat', label: t('Picture Format') }] : []),
        // 3D Model, as Word's: there while a 3D model is picked.
        ...(pickedModel ? [{ id: 'model3d', label: t('3D Model') }] : []),
        // Table Design and Table Layout, as Word's: there while the caret is in a table.
        ...(table ? [{ id: 'tableDesign', label: t('Table Design') }, { id: 'tableLayout', label: t('Table Layout') }] : []),
        // View → Immersive Reader's own tab, there while the reader is open.
        ...(view.immersive ? [{ id: 'immersive', label: t('Immersive Reader') }] : []),
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title={t('Save (Ctrl+S)')} onClick={() => save(false)} />
          <Button icon="undo" title={t('Undo (Ctrl+Z)')} disabled={!doc?.canUndo} onClick={() => commands['edit.undo'].run()} />
          <Button icon="redo" title={t('Redo (Ctrl+Y)')} disabled={!doc?.canRedo} onClick={() => commands['edit.redo'].run()} />
          <Button icon="print" title={t('Print')} onClick={() => commands['file.print']?.run?.()} />
        </>
      }
    >
      {/* ── Home ─────────────────────────────────────────────────────────── */}
      {tab === 'home' ? (
        <>
          <Group label={t('Clipboard')}>
            <Button tall icon="paste" label={t('Paste')} title={t('Paste (Ctrl+V)')} onClick={() => commands['edit.paste']?.run?.()} />
            <Rows>
              <>
                <Button icon="cut" label={t('Cut')} title={t('Cut (Ctrl+X)')} onClick={() => commands['edit.cut']?.run?.()} />
                <Button icon="copy" label={t('Copy')} title={t('Copy (Ctrl+C)')} onClick={() => commands['edit.copy']?.run?.()} />
              </>
              <Button icon="wand" label={t('Format Painter')} pressed={Boolean(view.painting)} title={t('Format Painter — copy the formatting here, then select the words to paint')} onClick={() => act('formatPainter')} />
            </Rows>
          </Group>

          <Group label={t('Font')}>
            <Rows>
              <>
                <Select value={format.fontName || 'Calibri'} onChange={(e) => run({ fontName: e.target.value })} style={{ width: 132 }} title={t('Font')}>
                  {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
                  {format.fontName && !FONTS.includes(format.fontName) ? <option value={format.fontName}>{format.fontName}</option> : null}
                </Select>
                <Select value={String(size)} onChange={(e) => run({ fontSize: Number(e.target.value) })} style={{ width: 58 }} title={t('Size')}>
                  {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
                  {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
                </Select>
                <Button icon="chevronUp" title={t('Grow font')} onClick={() => run({ fontSize: nearer(1) })} />
                <Button icon="chevronDown" title={t('Shrink font')} onClick={() => run({ fontSize: nearer(-1) })} />
                <Button icon="textbox" title={t('Change case')} onClick={(e) => menu.open(e, CASES.map(([mode, label]) => ({ label, run: () => act('changeCase', mode) })))} />
                <Button icon="undo" title={t('Clear all formatting')} onClick={() => dispatch({ op: 'clearFormat' })} />
              </>
              <>
                <Button icon="bold" title={t('Bold (Ctrl+B)')} pressed={format.bold} onClick={() => dispatch({ op: 'toggleFormat', tag: 'bold' })} />
                <Button icon="italic" title={t('Italic (Ctrl+I)')} pressed={format.italic} onClick={() => dispatch({ op: 'toggleFormat', tag: 'italic' })} />
                <Button icon="underline" title={t('Underline (Ctrl+U)')} pressed={format.underline} onClick={() => dispatch({ op: 'toggleFormat', tag: 'underline' })} />
                <Button icon="strike" title={t('Strikethrough')} pressed={format.strike} onClick={() => dispatch({ op: 'toggleFormat', tag: 'strike' })} />
                <Button icon="chevronDown" label="x₂" title={t('Subscript')} pressed={format.vertAlign === 'subscript'} onClick={() => run({ vertAlign: format.vertAlign === 'subscript' ? null : 'subscript' })} />
                <Button icon="chevronUp" label="x²" title={t('Superscript')} pressed={format.vertAlign === 'superscript'} onClick={() => run({ vertAlign: format.vertAlign === 'superscript' ? null : 'superscript' })} />
                <Separator />
                <Button icon="wand" label="A" title={t('Text effects — an outline, a shadow or a glow on the selected words')} onClick={(e) => menu.open(e, [
                  { label: t('Outline'), icon: format.outline ? 'check' : undefined, run: () => run({ outline: !format.outline }) },
                  { label: t('Shadow'), icon: format.shadow ? 'check' : undefined, run: () => run({ shadow: !format.shadow }) },
                  '-',
                  ...GLOWS.map(([colour, label]) => ({ label, icon: format.glow?.colour === colour ? 'check' : undefined, run: () => run({ glow: { colour, radiusPt: 4 } }) })),
                  { label: t('No glow'), run: () => run({ glow: null }) },
                  '-',
                  { label: t('Clear effects'), run: () => run({ outline: false, shadow: false, glow: null }) },
                ])} />
                <Button icon="wand" title={t('Text highlight colour')} pressed={Boolean(format.highlight)} onClick={(e) => colourMenu(e, 'highlight', HIGHLIGHTS)} />
                <Button icon="contrast" title={t('Font colour')} onClick={(e) => colourMenu(e, 'fontColour', TEXT_COLOURS)} />
              </>
            </Rows>
          </Group>

          <Group label={t('Paragraph')}>
            <Rows>
              <>
                <Button icon="listBullet" title={t('Bullets')} pressed={format.listType === 'bullet'} onClick={() => toggleList('bullet')} />
                <Button icon="listNumber" title={t('Numbering')} pressed={format.listType === 'number'} onClick={() => toggleList('number')} />
                <Button icon="listNumber" className="wd-multilevel" title={t('Multilevel List — a list library: 1. 1.1. 1.1.1., 1) a) i), I. A. 1. or Article and Section; deeper with Increase indent')} pressed={format.listType === 'number' && (format.listLevel || 0) > 0} onClick={(e) => menu.open(e, [
                  { heading: true, label: t('List Library') },
                  ...MULTILEVEL.map(([list, label]) => ({ label, icon: 'listNumber', run: () => para({ list }) })),
                  '-',
                  { label: t('None'), icon: 'close', run: () => para({ list: null }) },
                  { label: t('Define New Multilevel List…'), icon: 'listNumber', run: () => openDialog('defineList') },
                ])} />
                <Separator />
                <Button icon="chevronLeft" title={t('Decrease indent')} disabled={!format.indentLevel} onClick={() => para({ indentDelta: -1 })} />
                <Button icon="chevronRight" title={t('Increase indent')} onClick={() => para({ indentDelta: 1 })} />
                <Button icon="sort" title={t('Sort paragraphs — A to Z or Z to A; the whole document when nothing is selected')} onClick={(e) => menu.open(e, [
                  { label: t('A to Z'), run: () => dispatch({ op: 'sortParagraphs', descending: false }) },
                  { label: t('Z to A'), run: () => dispatch({ op: 'sortParagraphs', descending: true }) },
                ])} />
                <Button icon="formula" title={t('Show formatting marks (¶)')} pressed={Boolean(view.marks)} onClick={() => act('toggleMarks')} />
                <Separator />
                <Button icon="textLtr" title={t('Left-to-right text direction')} pressed={!format.rtl} onClick={() => para({ rtl: false })} />
                <Button icon="textRtl" title={t('Right-to-left text direction — the paragraph runs from the right margin, as Arabic, Hebrew, Persian and Urdu do')} pressed={Boolean(format.rtl)} onClick={() => para({ rtl: true })} />
              </>
              <>
                <Button icon="alignLeft" title={t('Align left (Ctrl+L)')} pressed={format.paragraphAlign === 'left'} onClick={() => para({ align: 'left' })} />
                <Button icon="alignCenter" title={t('Centre (Ctrl+E)')} pressed={format.paragraphAlign === 'center'} onClick={() => para({ align: 'center' })} />
                <Button icon="alignRight" title={t('Align right (Ctrl+R)')} pressed={format.paragraphAlign === 'right'} onClick={() => para({ align: 'right' })} />
                <Button icon="alignJustify" title={t('Justify (Ctrl+J)')} pressed={format.paragraphAlign === 'justify' || format.paragraphAlign === 'both'} onClick={() => para({ align: 'both' })} />
                {format.rtl ? (
                  // Word's Arabic justifications: the line filled by stretching the letters with kashida.
                  <Button icon="chevronDown" title={t('Justify with kashida — stretch the letters to fill the line, as Arabic is justified')} pressed={KASHIDA.some(([v]) => v === format.paragraphAlign)} onClick={(e) =>
                    menu.open(e, KASHIDA.map(([v, label]) => ({ label: t(label), icon: format.paragraphAlign === v ? 'check' : undefined, run: () => para({ align: v }) })))} />
                ) : null}
                <Separator />
                <Button icon="listNumber" title={format.lineSpacing ? t('Line and paragraph spacing — {spacing}', { spacing: format.lineSpacing }) : t('Line and paragraph spacing')} onClick={(e) =>
                  menu.open(e, [
                    ...SPACING.map(([v, l]) => ({ label: l, icon: format.lineSpacing === v ? 'check' : undefined, run: () => para({ lineSpacing: v }) })),
                    '-',
                    { label: t('Add space before paragraph'), run: () => para({ spaceBefore: 12 }) },
                    { label: t('Remove space before paragraph'), run: () => para({ spaceBefore: 0 }) },
                    { label: t('Add space after paragraph'), run: () => para({ spaceAfter: 8 }) },
                    { label: t('Remove space after paragraph'), run: () => para({ spaceAfter: 0 }) },
                  ])
                } />
                <Button icon="wand" title={t('Shading — a colour behind the paragraph')} onClick={(e) => menu.open(e, SHADES.map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => para({ shading: value }) })))} />
                <Button icon="grid" title={t('Borders — lines round the paragraph')} onClick={(e) => menu.open(e, BORDERS.map(([label, value]) => ({ label, icon: value ? undefined : 'close', run: () => para({ borders: value }) })))} />
              </>
            </Rows>
          </Group>

          <Group label={t('Styles')}>
            <div className="wd-styles">
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).slice(0, 8).map((s) => (
                <button
                  key={s.id}
                  type="button"
                  className={`wd-style${(format.paragraphStyle || 'Normal') === s.id ? ' on' : ''}`}
                  title={t('Apply {style}', { style: s.name || s.id })}
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
              title={t('Paragraph style')}
            >
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).map((s) => <option key={s.id} value={s.id}>{s.name || s.id}</option>)}
            </Select>
          </Group>

          <Group label={t('Editing')}>
            <Rows>
              <>
                <Button icon="find" label={t('Find')} title={t('Find (Ctrl+F)')} onClick={() => openDialog('find')} />
                <Button icon="refresh" label={t('Replace')} title={t('Replace (Ctrl+H)')} onClick={() => openDialog('find')} />
              </>
              <Button icon="check" label={t('Select')} title={t('Select all (Ctrl+A)')} onClick={() => dispatch({ op: 'selectAll' })} />
            </Rows>
          </Group>

          <Group label={t('Voice')}>
            <Button tall icon="volume" label={t('Read aloud')} pressed={Boolean(view.reading)} title={t('Read the document aloud, from the caret')} onClick={() => act('readAloud')} />
          </Group>
        </>
      ) : null}

      {/* ── Insert ───────────────────────────────────────────────────────── */}
      {tab === 'insert' ? (
        <>
          <Group label={t('Pages')}>
            <Button tall icon="file" label={t('Cover Page')} title={t('A title page at the front')} onClick={() => act('coverPage')} />
            <Button tall icon="file" label={t('Blank Page')} title={t('A page break here, and another after — an empty page')} onClick={() => act('blankPage')} />
            <Button tall icon="file" label={t('Page Break')} title={t('Page break (Ctrl+Enter)')} onClick={() => dispatch({ op: 'insertPageBreak' })} />
          </Group>
          <Group label={t('Tables')}>
            <Button tall icon="table" label={t('Table')} onClick={() => openDialog('table')} />
            <Button icon="table" label={t('Text to Table')} disabled={Boolean(table)} title={table ? t('Convert Text to Table — the caret is in a table already') : t('Convert Text to Table — the selected paragraphs as rows, split at their tabs or commas')} onClick={(e) => menu.open(e, [
              { label: t('Split at tabs'), run: () => dispatch({ op: 'textToTable', separator: 'tab' }) },
              { label: t('Split at commas'), run: () => dispatch({ op: 'textToTable', separator: 'comma' }) },
            ])} />
          </Group>
          <Group label={t('Illustrations')}>
            <Button tall icon="picture" label={t('Pictures')} onClick={insertPicture} />
            <Button tall icon="shape" label={t('Shapes')} onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', preset, widthPx: 200, heightPx: 120 }) })))} />
            <Button tall icon="star" label={t('Icons')} title={t("Icons — one of the suite's own icons, in the colour you choose, as a picture at the caret")} onClick={() => act('icons')} />
            <Button tall icon="shape" label={t('3D Models')} className="wd-model3d-insert" title={t('3D Models — a model from a .glb or .gltf file on this computer, drawn as a picture at the caret; turn it from the 3D Model tab')} onClick={() => act('model3d')} />
            <Button tall icon="grid" label={t('SmartArt')} title={t('SmartArt — a list, a process, a cycle or a hierarchy, drawn from lines you type, as a group of shapes after this paragraph')} onClick={() => act('insertSmartArt')} />
            <Button tall icon="chart" label={t('Chart')} title={t('A chart drawn from the table the caret is in — put the caret in a table first')} onClick={(e) => menu.open(e, CHARTS.map(([k, l]) => ({ label: l, icon: 'chart', run: () => dispatch({ op: 'insertChart', kind: k }) })))} />
            <Button tall icon="picture" label={t('Screenshot')} title={t('Screenshot — a picture of another open window, or of a whole screen, put in at the caret')} onClick={() => act('screenshot')} />
          </Group>
          <Group label={t('Media')}>
            <Soon tall icon="video" label={t('Online Videos')} why={t('Embedding a video is a network request; this suite makes none it has not declared.')} />
          </Group>
          <Group label={t('Links')}>
            <Button tall icon="link" label={t('Link')} title={t('Link (Ctrl+K)')} pressed={Boolean(format.link)} onClick={() => openDialog('link')} />
            <Button icon="flag" label={t('Bookmark')} title={t('Bookmark — name this place in the document, so Go To and a cross-reference can find it')} onClick={() => openDialog('bookmark')} />
            <Button icon="link" label={t('Cross-reference')} title={t('Cross-reference — a REF field to a bookmark, its words kept live by Update Fields')} onClick={() => openDialog('crossReference')} />
          </Group>
          <Group label={t('Comments')}>
            <Button tall icon="reply" label={t('Comment')} onClick={() => openDialog('comment')} />
          </Group>
          <Group label={t('Header & Footer')}>
            <Button tall icon="chevronUp" label={t('Header')} onClick={() => openDialog('header')} />
            <Button tall icon="chevronDown" label={t('Footer')} onClick={() => openDialog('footer')} />
            <Button tall icon="file" label={t('Page Number')} onClick={() => openDialog('pageNumber')} />
          </Group>
          <Group label={t('Text')}>
            <Button tall icon="textbox" label={t('Text Box')} title={t('Text Box — a box of words that floats on the page, the text wrapping round it')} onClick={(e) => menu.open(e, [
              { heading: true, label: t('Built-in') },
              { label: t('Simple Text Box'), icon: 'textbox', run: () => act('textBox', 'simple') },
              { label: t('Sidebar'), icon: 'textbox', run: () => act('textBox', 'sidebar') },
              { label: t('Pull Quote'), icon: 'textbox', run: () => act('textBox', 'quote') },
              '-',
              { label: t('Draw Text Box'), icon: 'shape', run: () => act('drawTextBox') },
            ])} />
            {references ? <Button icon="file" label={t('Quick Parts')} title={t("Quick Parts — words you keep to put in again, AutoText, a field: a page number, the date, the file's name")} onClick={(e) => menu.open(e, [...(blocks?.head() || []), ...references.quickParts(), ...(blocks?.tail() || [])])} /> : null}
            <Button icon="wand" label={t('WordArt')} title={t('WordArt — decorative words in a box of their own that floats on the page')} onClick={(e) => menu.open(e, WORDART.map(([label, spec]) => ({ label, icon: 'wand', run: () => act('wordArt', spec) })))} />
            <Button icon="textbox" label={t('Drop Cap')} title={t('Drop Cap — the first letter, framed to stand tall beside the words that follow it')} onClick={(e) => menu.open(e, [
              { label: t('None'), icon: !format.dropCap ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: null }) },
              { label: t('Dropped'), icon: format.dropCap?.kind === 'drop' && format.dropCap?.lines === 3 ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: { kind: 'drop', lines: 3 } }) },
              { label: t('Dropped, two lines'), icon: format.dropCap?.kind === 'drop' && format.dropCap?.lines === 2 ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: { kind: 'drop', lines: 2 } }) },
              { label: t('In margin'), icon: format.dropCap?.kind === 'margin' ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: { kind: 'margin', lines: 3 } }) },
            ])} />
            <Button icon="check" label={t('Signature Line')} title={t('Signature Line — a line for someone to sign in Word, with their name, title and e-mail under it')} onClick={() => act('signatureLine')} />
            <Button icon="clock" label={t('Date & Time')} onClick={() => openDialog('dateTime')} />
            <Button icon="file" label={t('Object')} title={t('Object — a Word, Excel or PowerPoint document embedded as an icon; a double-click opens it')} onClick={() => act('insertObject')} />
          </Group>
          <Group label={t('Symbols')}>
            <Button tall icon="formula" label={t('Equation')} title={t("Equation — type one in Word's linear format, or build it from structures and symbols (Alt+=)")} onClick={() => act('equation')} />
            <Button
              icon="chevronDown"
              label={t('Built-in')}
              title={t('Built-in equations — the quadratic formula, the area of a circle, the binomial theorem and more, put in at the caret')}
              onClick={(e) => menu.open(e, EQUATION_GALLERY.map((g) => ({ label: t(g.name), icon: 'formula', run: () => act('equation', { linear: g.linear }) })))}
            />
            <Button tall icon="star" label={t('Symbol')} onClick={() => openDialog('symbol')} />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label={t('Drawing Tools')}>
            <Button tall icon="mouse" label={t('Select')} pressed={!ink?.tool} title={t('Select — put the pen down and type again (Esc)')} onClick={() => act('inkTool', null)} />
            <Button tall icon="shape" label={t('Lasso')} pressed={ink?.tool === 'lasso'} title={t('Lasso Select — draw a loop round strokes to select them')} onClick={() => act('inkTool', 'lasso')} />
            <Button tall icon="close" label={t('Eraser')} pressed={ink?.tool === 'eraser'} title={t('Eraser — take away each stroke the pointer passes over')} onClick={() => act('inkTool', 'eraser')} />
            {(ink?.pens || []).map((p) => {
              const on = ink.tool === 'pen' && ink.penId === p.id;
              const label = { pen: t('Pen'), pencil: t('Pencil'), highlighter: t('Highlighter') }[p.tool];
              return (
                <Button key={p.id} tall icon="wand" label={label} pressed={on} className={`sl-pen sl-pen-${p.tool}`} data-pen={p.id} style={{ '--pen': p.color }}
                  title={on ? t('{label} — click again for its colour and thickness', { label }) : t('{label} — draw on the page', { label })}
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
          <Group label={t('Touch')}>
            <Button tall icon="wand" label={t('Draw with Touch')} pressed={ink?.touch !== false} title={t('Draw with Touch — a finger draws with the pen in hand; off, only a pen or the mouse does')} onClick={() => act('inkTouch')} />
          </Group>
          <Group label={t('Stencils')}>
            <Button tall icon="minus" label={t('Ruler')} pressed={Boolean(ink?.ruler)} title={t('Ruler — a straight edge on the page: a stroke begun along it follows it; drag it to move, the wheel to turn it')} onClick={() => act('inkRuler')} />
          </Group>
          <Group label={t('Convert')}>
            <Button tall icon="shape" label={t('Ink to Shape')} pressed={Boolean(ink?.toShape)} title={t('Ink to Shape — a rectangle, oval or triangle drawn becomes that shape')} onClick={() => act('inkToShape')} />
            <Soon tall icon="formula" label={t('Ink to Maths')} why={t('Turning handwriting into an equation needs handwriting recognition, which this suite does not have.')} />
          </Group>
          <Group label={t('Insert')}>
            <Button tall icon="shape" label={t('Shape')} title={t('The shapes the engine can draw')} onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', preset, widthPx: 200, heightPx: 120 }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Design ───────────────────────────────────────────────────────── */}
      {tab === 'design' ? (
        <>
          <Group label={t('Document Formatting')}>
            <Button
              tall
              icon="wand"
              label={t('Themes')}
              title={t('Themes — colours, fonts and effects for the whole document, in one; now {theme}', { theme: design?.name || 'Office Theme' })}
              onClick={(e) => menu.open(e, THEMES.map((th) => ({ label: t(th.name), icon: design?.builtIn === th.id ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { theme: th.id } }) })))}
            />
            <div className="wd-styles">
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).slice(0, 6).map((s) => (
                <button key={s.id} type="button" className={`wd-style${(format.paragraphStyle || 'Normal') === s.id ? ' on' : ''}`} style={STYLE_LOOK[s.id] || undefined} title={t('Apply {style} to this paragraph', { style: s.name || s.id })} onClick={() => para({ styleId: s.id })}>
                  {s.name || s.id}
                </button>
              ))}
            </div>
            <Button
              icon="contrast"
              label={t('Colours')}
              title={t("Colours — the theme's twelve colours, which headings and theme-coloured text follow; now {colours}", { colours: design?.colorName || 'Office' })}
              onClick={(e) => menu.open(e, PALETTES.map((p) => ({ label: t(p.name), icon: design?.colorName === p.name ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { colors: p.id } }) })))}
            />
            <Button
              icon="word"
              label={t('Fonts')}
              title={t("Fonts — the theme's heading and body faces, which the styles follow; now {fonts}", { fonts: design?.fonts ? `${design.fonts.major} / ${design.fonts.minor}` : 'Calibri Light / Calibri' })}
              onClick={(e) => menu.open(e, FONT_PAIRS.map((p) => ({ label: `${t(p.name)} — ${p.major} / ${p.minor}`, icon: design?.fontName === p.name ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { fonts: p.id } }) })))}
            />
            <Button icon="listNumber" label={t('Paragraph Spacing')} title={t('Spacing for the paragraphs you have selected')} onClick={(e) =>
              menu.open(e, [
                { label: t('No paragraph space'), run: () => para({ spaceBefore: 0, spaceAfter: 0, lineSpacing: 1 }) },
                { label: t('Compact'), run: () => para({ spaceBefore: 0, spaceAfter: 4, lineSpacing: 1 }) },
                { label: t('Tight'), run: () => para({ spaceBefore: 0, spaceAfter: 6, lineSpacing: 1.15 }) },
                { label: t('Open'), run: () => para({ spaceBefore: 0, spaceAfter: 10, lineSpacing: 1.15 }) },
                { label: t('Relaxed'), run: () => para({ spaceBefore: 0, spaceAfter: 6, lineSpacing: 1.5 }) },
                { label: t('Double'), run: () => para({ spaceBefore: 0, spaceAfter: 8, lineSpacing: 2 }) },
              ])
            } />
            <Button
              icon="wand"
              label={t('Effects')}
              title={t('Effects — the fills, lines and shadows that shapes taking their look from the theme are drawn with; now {effects}', { effects: design?.effectName || 'Office' })}
              onClick={(e) => menu.open(e, EFFECT_PRESETS.map((p) => ({ label: t(p.name), title: t(p.description), icon: design?.effects === p.id ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { effects: p.id } }) })))}
            />
            <Button icon="check" label={t('Set as Default')} title={t("Set as Default — new blank documents start in this document's theme and styles")} onClick={() => act('setDefaultDesign')} />
          </Group>
          <Group label={t('Page Background')}>
            <Button tall icon="shield" label={t('Watermark')} title={t('Watermark — faint words, or a picture, behind every page')} onClick={(e) => menu.open(e, [
              ...WATERMARKS.map((text) => ({ label: text, run: () => dispatch({ op: 'setWatermark', text }) })),
              { label: t('Custom watermark…'), icon: 'textbox', run: () => openDialog('watermark') },
              { label: t('Picture watermark…'), icon: 'picture', run: () => act('pictureWatermark') },
              { label: t('Remove watermark'), icon: 'close', run: () => dispatch({ op: 'setWatermark', text: null }) },
            ])} />
            <Button tall icon="contrast" label={t('Page Colour')} title={t('Page Colour — a colour behind every page')} onClick={(e) => menu.open(e, SHADES.map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => dispatch({ op: 'setPageColour', colour: value }) })))} />
            <Button tall icon="grid" label={t('Page Borders')} title={t('Page Borders — a frame around every page')} onClick={(e) => menu.open(e, PAGE_BORDERS.map(([label, borders]) => ({ label, icon: borders ? undefined : 'close', run: () => dispatch({ op: 'setPageBorders', borders }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Layout ───────────────────────────────────────────────────────── */}
      {tab === 'layout' ? (
        <>
          <Group label={t('Page Setup')}>
            <Button tall icon="crop" label={t('Margins')} onClick={(e) => menu.open(e, MARGINS.map(([v, l]) => ({ label: l, run: () => dispatch({ op: 'setPageSetup', spec: { margins: v } }) })))} />
            <Button tall icon="rotate" label={t('Orientation')} title={landscape ? t('Orientation — now landscape') : t('Orientation — now portrait')} onClick={(e) =>
              menu.open(e, [
                { label: t('Portrait'), icon: landscape ? undefined : 'check', run: () => dispatch({ op: 'setPageSetup', spec: { orientation: 'portrait' } }) },
                { label: t('Landscape'), icon: landscape ? 'check' : undefined, run: () => dispatch({ op: 'setPageSetup', spec: { orientation: 'landscape' } }) },
              ])
            } />
            <Button tall icon="file" label={t('Size')} onClick={(e) => menu.open(e, PAGE_SIZES.map(([v, l]) => ({ label: l, run: () => dispatch({ op: 'setPageSetup', spec: { size: v } }) })))} />
            <Button tall icon="grid" label={t('Columns')} title={colCount > 1 ? t('Columns — now {count}', { count: colCount }) : t('Columns — now one')} onClick={(e) => menu.open(e, [
              { label: t('One'), icon: isColumnPreset('one') ? 'check' : undefined, run: () => setColumns(null) },
              { label: t('Two'), icon: isColumnPreset('two') ? 'check' : undefined, run: () => setColumns({ count: 2, spaceTwips: 720, separator: Boolean(columns.separator) }) },
              { label: t('Three'), icon: isColumnPreset('three') ? 'check' : undefined, run: () => setColumns({ count: 3, spaceTwips: 720, separator: Boolean(columns.separator) }) },
              { label: t('Left'), icon: isColumnPreset('left') ? 'check' : undefined, run: () => setColumns(unevenSpec(true)) },
              { label: t('Right'), icon: isColumnPreset('right') ? 'check' : undefined, run: () => setColumns(unevenSpec(false)) },
              '-',
              // The section right to left: its columns run from the right, as an Arabic or Hebrew page's do.
              { label: t('Right to left'), icon: section?.rtl ? 'check' : undefined, run: () => dispatch({ op: 'setPageSetup', spec: { rtl: !section?.rtl } }) },
              {
                label: t('Line between'), icon: columns.separator ? 'check' : undefined, disabled: colCount <= 1,
                run: () => setColumns({
                  count: colCount, spaceTwips: Math.round((columns.spacePx || 36) * 15), separator: !columns.separator,
                  ...(columns.widths ? { widths: columns.widths.map((w) => Math.round(w * 15)) } : {}),
                }),
              },
            ])} />
            <Button icon="file" label={t('Breaks')} onClick={(e) =>
              menu.open(e, [
                { label: t('Page break'), icon: 'file', run: () => dispatch({ op: 'insertPageBreak' }) },
                { label: t('Section break — not built yet'), disabled: true },
              ])
            } />
            <Button icon="listNumber" label={t('Line Numbers')} title={t('Line Numbers — a number beside every line, down the left margin')} onClick={(e) => menu.open(e, [
              { label: t('None'), icon: 'close', run: () => dispatch({ op: 'setLineNumbers', spec: null }) },
              { label: t('Continuous'), run: () => dispatch({ op: 'setLineNumbers', spec: { countBy: 1, restart: 'continuous' } }) },
              { label: t('Restart each page'), run: () => dispatch({ op: 'setLineNumbers', spec: { countBy: 1, restart: 'newPage' } }) },
              { label: t('Every fifth line'), run: () => dispatch({ op: 'setLineNumbers', spec: { countBy: 5, restart: 'continuous' } }) },
            ])} />
            <Button icon="minus" label={t('Hyphenation')} title={model?.hyphenation?.auto ? t('Hyphenation — words broken at the ends of lines: now automatic') : t('Hyphenation — words broken at the ends of lines: now none')} onClick={(e) => menu.open(e, [
              { label: t('None'), icon: !model?.hyphenation?.auto ? 'check' : undefined, run: () => act('hyphenation', 'none') },
              { label: t('Automatic'), icon: model?.hyphenation?.auto ? 'check' : undefined, run: () => act('hyphenation', 'auto') },
              { label: t('Manual'), run: () => act('hyphenation', 'manual') },
              '-',
              { label: t('Hyphenation Options…'), run: () => act('hyphenation', 'options') },
              { label: t("Don't hyphenate this paragraph"), icon: format.noHyphens ? 'check' : undefined, run: () => act('hyphenation', 'paragraph') },
            ])} />
          </Group>
          <Group label={t('Paragraph')}>
            <div className="wd-fields">
              <label>{t('Indent')}</label>
              <Button icon="chevronLeft" title={t('Decrease left indent')} disabled={!format.indentLevel} onClick={() => para({ indentDelta: -1 })} />
              <span className="wd-field-value">{t('{value} cm', { value: (format.indentLevel || 0) * 1.27 })}</span>
              <Button icon="chevronRight" title={t('Increase left indent')} onClick={() => para({ indentDelta: 1 })} />
            </div>
            <div className="wd-fields">
              <label>{t('Spacing')}</label>
              <Input type="number" min="0" max="200" value={format.spaceBefore ?? 0} title={t('Before (pt)')} onChange={(e) => para({ spaceBefore: Number(e.target.value) })} style={{ width: 56 }} />
              <Input type="number" min="0" max="200" value={format.spaceAfter ?? 0} title={t('After (pt)')} onChange={(e) => para({ spaceAfter: Number(e.target.value) })} style={{ width: 56 }} />
            </div>
          </Group>
          {arrange}
        </>
      ) : null}

      {/* ── Shape Format (contextual) ─────────────────────────────────────── */}
      {tab === 'shapeFormat' ? (
        <>
          <Group label={t('Insert Shapes')}>
            <Button tall icon="textbox" label={t('Draw Text Box')} pressed={Boolean(view.drawBox)} title={t('Draw Text Box — drag on the page to draw one')} onClick={() => act('drawTextBox')} />
          </Group>
          <Group label={t('Shape Styles')}>
            <Button tall icon="wand" label={t('Shape Fill')} disabled={!shapeLike} title={shapeLike ? (look.fill ? t('Shape Fill — now {fill}', { fill: look.fill }) : t('Shape Fill — now no fill')) : t('Shape Fill — select a text box or a shape first')} onClick={(e) => menu.open(e, [
              { label: t('No Fill'), icon: 'close', run: () => act('boxFormat', { fill: null }) },
              '-',
              ...SHAPE_COLOURS.map(([hex, label]) => ({ label, icon: look.fill && look.fill.replace('#', '').toUpperCase() === hex ? 'check' : undefined, run: () => act('boxFormat', { fill: hex }) })),
            ])} />
            <Button tall icon="shape" label={t('Shape Outline')} disabled={!shapeLike} title={shapeLike ? (look.line ? t('Shape Outline — now {line}', { line: look.line }) : t('Shape Outline — now no outline')) : t('Shape Outline — select a text box or a shape first')} onClick={(e) => menu.open(e, [
              { label: t('No Outline'), icon: 'close', run: () => act('boxFormat', { line: null }) },
              '-',
              ...SHAPE_COLOURS.map(([hex, label]) => ({ label, icon: look.line && look.line.replace('#', '').toUpperCase() === hex ? 'check' : undefined, run: () => act('boxFormat', { line: { colour: hex, widthPx: look.lineWidthPx || 1 } }) })),
              { heading: true, label: t('Weight') },
              ...WEIGHTS.map(([label, px]) => ({ label, icon: look.line && Math.abs((look.lineWidthPx || 1) - px) < 0.2 ? 'check' : undefined, run: () => act('boxFormat', { line: { colour: (look.line || '#000000').replace('#', ''), widthPx: px } }) })),
            ])} />
          </Group>
          <Group label={t('Text')}>
            <Button tall icon="rotate" label={t('Text Direction')} disabled={drawing?.kind !== 'textbox'} title={t('Text Direction — across, or turned to read down or up')} onClick={(e) => menu.open(e, [
              { label: t('Horizontal'), icon: (look.vert || 'horz') === 'horz' ? 'check' : undefined, run: () => act('boxFormat', { vert: 'horz' }) },
              { label: t('Rotate all text 90°'), icon: look.vert === 'vert' ? 'check' : undefined, run: () => act('boxFormat', { vert: 'vert' }) },
              { label: t('Rotate all text 270°'), icon: look.vert === 'vert270' ? 'check' : undefined, run: () => act('boxFormat', { vert: 'vert270' }) },
            ])} />
            <Button tall icon="alignCenter" label={t('Align Text')} disabled={drawing?.kind !== 'textbox'} title={t('Align Text — the words at the top, middle or bottom of the box')} onClick={(e) => menu.open(e, [
              { label: t('Top'), icon: (look.vAnchor || 'top') === 'top' ? 'check' : undefined, run: () => act('boxFormat', { vAnchor: 'top' }) },
              { label: t('Middle'), icon: look.vAnchor === 'middle' ? 'check' : undefined, run: () => act('boxFormat', { vAnchor: 'middle' }) },
              { label: t('Bottom'), icon: look.vAnchor === 'bottom' ? 'check' : undefined, run: () => act('boxFormat', { vAnchor: 'bottom' }) },
            ])} />
            <Button tall icon="crop" label={t('Margins')} disabled={drawing?.kind !== 'textbox'} title={t("Margins — the room between the box's edge and its words")} onClick={(e) => menu.open(e, BOX_MARGINS.map(([label, insets]) => ({
              label, icon: look.insets && ['l', 't', 'r', 'b'].every((k) => Math.abs((look.insets[k] || 0) - insets[k]) < 0.6) ? 'check' : undefined, run: () => act('boxFormat', { insets }),
            })))} />
          </Group>
          {/* WordArt Styles → Text Effects → Transform, as Word's: a text box's words along a path. */}
          <Group label={t('Text Effects: Transform')}>
            {WARP_PRESETS.map((p) => (
              <Button
                key={p.id} tall icon="wand" label={t(p.label)} className="wd-warp" data-preset={p.id}
                disabled={drawing?.kind !== 'textbox'}
                pressed={drawing?.kind === 'textbox' && (look.warp || 'textNoShape') === p.id}
                title={drawing?.kind !== 'textbox' ? t('{label} — select a text box or WordArt first', { label: t(p.label) }) : p.id === 'textNoShape' ? t('No Transform — the words in straight lines') : p.label === 'Button' ? t('Transform — the words along a button: an arc, a line and an arc') : p.label === 'Circle' ? t('Transform — the words along the circle of the box') : t('Transform — the words along the arc of the box')}
                onClick={() => act('boxFormat', { warp: p.id === 'textNoShape' ? null : p.id })}
              />
            ))}
            {/* The warps, the rest of Office's gallery, from a menu. */}
            <Button tall disabled={drawing?.kind !== 'textbox'} icon="wand" label={t('More')} className="wd-warp-more" pressed={WARP_MORE.some((p) => p.id === look.warp)} title={WARP_MORE.some((p) => p.id === look.warp) ? t('Transform — now {warp}; the warps: the words stretched between two curves', { warp: t(warpLabel(look.warp)) }) : t('More Transforms — the warps: the words stretched between two curves, a wave, a slant, a chevron and the rest')} onClick={(e) => menu.open(e, WARP_MORE.map((p) => ({ label: t(p.label), icon: look.warp === p.id ? 'check' : 'wand', run: () => act('boxFormat', { warp: p.id }) })))} />
          </Group>
          {arrange}
          <Group label={t('Size')}>
            <Rows>
              {sizeBox(t('Height'), 'h')}
              {sizeBox(t('Width'), 'w')}
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Table Design (contextual) ─────────────────────────────────────── */}
      {tab === 'tableDesign' && table ? (
        <>
          <Group label={t('Table Style Options')}>
            <Rows>
              <Button icon={styleLook.firstRow ? 'check' : undefined} label={t('Header Row')} pressed={styleLook.firstRow} title={t("Header Row — the first row in its style's header look")} onClick={() => setLook({ firstRow: !styleLook.firstRow })} />
              <Button icon={styleLook.lastRow ? 'check' : undefined} label={t('Total Row')} pressed={styleLook.lastRow} title={t("Total Row — the last row in its style's total look")} onClick={() => setLook({ lastRow: !styleLook.lastRow })} />
              <Button icon={!styleLook.noHBand ? 'check' : undefined} label={t('Banded Rows')} pressed={!styleLook.noHBand} title={t('Banded Rows — every other row shaded, as the style bands them')} onClick={() => setLook({ noHBand: !styleLook.noHBand })} />
            </Rows>
            <Rows>
              <Button icon={styleLook.firstColumn ? 'check' : undefined} label={t('First Column')} pressed={styleLook.firstColumn} title={t("First Column — the first column in its style's look")} onClick={() => setLook({ firstColumn: !styleLook.firstColumn })} />
              <Button icon={styleLook.lastColumn ? 'check' : undefined} label={t('Last Column')} pressed={styleLook.lastColumn} title={t("Last Column — the last column in its style's look")} onClick={() => setLook({ lastColumn: !styleLook.lastColumn })} />
              <Button icon={!styleLook.noVBand ? 'check' : undefined} label={t('Banded Columns')} pressed={!styleLook.noVBand} title={t('Banded Columns — every other column shaded, as the style bands them')} onClick={() => setLook({ noVBand: !styleLook.noVBand })} />
            </Rows>
          </Group>
          <Group label={t('Table Styles')}>
            <Button tall icon="table" label={t('Table Styles')} title={tableStyleId ? t("Table Styles — Word's own, in this document's theme colours; now {style}", { style: t(TABLE_STYLES.find((st) => st.id === tableStyleId)?.name || tableStyleId) }) : t("Table Styles — Word's own, in this document's theme colours; now none")} onClick={(e) => menu.open(e, [
              { heading: true, label: t('Plain Tables') },
              ...TABLE_STYLES.filter((st) => st.family === 'plain').map((st) => ({ label: t(st.name), icon: st.id === tableStyleId ? 'check' : undefined, run: () => tableOp('style', { id: st.id }) })),
              { heading: true, label: t('Grid Tables') },
              ...TABLE_STYLES.filter((st) => st.family === 'grid4').map((st) => ({ label: t(st.name), icon: st.id === tableStyleId ? 'check' : undefined, run: () => tableOp('style', { id: st.id }) })),
              { heading: true, label: t('List Tables') },
              ...TABLE_STYLES.filter((st) => st.family === 'list4').map((st) => ({ label: t(st.name), icon: st.id === tableStyleId ? 'check' : undefined, run: () => tableOp('style', { id: st.id }) })),
              '-',
              { label: t('Clear'), icon: 'close', disabled: !tableStyleId, run: () => tableOp('style', { id: null }) },
            ])} />
          </Group>
          <Group label={t('Borders')}>
            <Rows>
              <Button icon="minus" label={t('Line: {style}', { style: PEN_STYLES.find(([v]) => v === pen.val)?.[1] || t('Single') })} title={t("Line Style — the pen's line for the borders it draws")} onClick={(e) => menu.open(e, PEN_STYLES.map(([val, label]) => ({ label, icon: pen.val === val ? 'check' : undefined, run: () => setPen({ ...pen, val }) })))} />
              <Button icon="sliders" label={t('Weight: {weight}', { weight: PEN_WEIGHTS.find(([sz]) => sz === pen.sz)?.[1] || t('½ pt') })} title={t('Pen Weight — how thick a line the pen draws')} onClick={(e) => menu.open(e, PEN_WEIGHTS.map(([sz, label]) => ({ label, icon: pen.sz === sz ? 'check' : undefined, run: () => setPen({ ...pen, sz }) })))} />
              <Button icon="wand" label={t('Pen Colour')} title={t('Pen Colour — now {colour}', { colour: SHAPE_COLOURS.find(([hex]) => hex === pen.color)?.[1] || t('Automatic') })} onClick={(e) => menu.open(e, [
                { label: t('Automatic'), icon: pen.color === 'auto' ? 'check' : undefined, run: () => setPen({ ...pen, color: 'auto' }) },
                '-',
                ...SHAPE_COLOURS.map(([hex, label]) => ({ label, icon: pen.color === hex ? 'check' : undefined, run: () => setPen({ ...pen, color: hex }) })),
              ])} />
            </Rows>
            <Button tall icon="grid" label={t('Borders')} title={t("Borders — lines round, between or through the selected cells, in the pen's line")} onClick={(e) => menu.open(e, [
              ['bottom', t('Bottom Border')], ['top', t('Top Border')], ['left', t('Left Border')], ['right', t('Right Border')], '-',
              ['none', t('No Border')], ['all', t('All Borders')], ['outside', t('Outside Borders')], ['inside', t('Inside Borders')], '-',
              ['insideH', t('Inside Horizontal Border')], ['insideV', t('Inside Vertical Border')],
            ].map((item) => (item === '-' ? '-' : { label: item[1], run: () => tableOp('borders', { kind: item[0], pen }) })))} />
          </Group>
          <Group label={t('Shading')}>
            <Button tall icon="wand" label={t('Shading')} title={t("Shading — the selected cells' background")} onClick={(e) => menu.open(e, [
              { label: t('No Colour'), icon: 'close', run: () => tableOp('shading', { fill: null }) },
              '-',
              ...SHAPE_COLOURS.map(([hex, label]) => ({ label, run: () => tableOp('shading', { fill: hex }) })),
            ])} />
          </Group>
        </>
      ) : null}

      {/* ── Table Layout (contextual) ─────────────────────────────────────── */}
      {tab === 'tableLayout' && table ? (
        <>
          <Group label={t('Draw')}>
            <Button tall icon="pen" label={t('Draw Table')} pressed={view.tableDraw === 'pen'} title={t('Draw Table — draw a line down a cell to split it into two columns there, or across it to split it into two rows; Escape puts the pen down')} onClick={() => act('tableDraw', 'pen')} />
            <Button tall icon="eraser" label={t('Eraser')} pressed={view.tableDraw === 'eraser'} title={t('Eraser — click a line between two cells to join them into one; Escape puts the eraser down')} onClick={() => act('tableDraw', 'eraser')} />
          </Group>
          <Group label={t('Table')}>
            <Button tall icon="settings" label={t('Properties')} title={t("Table Properties — the table's width, alignment and alt text, the caret's row and column")} onClick={() => openDialog('tableProperties')} />
            <Button tall icon="alignCenter" label={t('Align Table')} title={t('Align Table — the table at the left margin, centred or at the right')} onClick={(e) => menu.open(e, [
              { label: t('Left'), run: () => tableOp('align', { align: 'left' }) },
              { label: t('Centre'), run: () => tableOp('align', { align: 'center' }) },
              { label: t('Right'), run: () => tableOp('align', { align: 'right' }) },
            ])} />
            <Button tall icon="grid" label={t('View Gridlines')} pressed={!view.noTableGridlines} title={t("View Gridlines — a table's faint dashes where it has no lines of its own; on screen only, never printed")} onClick={() => act('toggleTableGridlines')} />
          </Group>
          <Group label={t('Rows & Columns')}>
            <Button tall icon="minus" label={t('Delete')} title={t('Delete — the row, the column or the whole table at the caret')} onClick={(e) => menu.open(e, [
              { label: t('Delete columns'), icon: 'minus', run: () => tableOp('deleteColumn') },
              { label: t('Delete rows'), icon: 'minus', run: () => tableOp('deleteRow') },
              { label: t('Delete table'), icon: 'trash', run: () => tableOp('deleteTable') },
            ])} />
            <Button tall icon="rowAbove" label={t('Insert Above')} title={t("Insert a row above the caret's")} onClick={() => tableOp('insertRowAbove')} />
            <Button tall icon="rowBelow" label={t('Insert Below')} title={t("Insert a row below the caret's")} onClick={() => tableOp('insertRowBelow')} />
            <Rows>
              <Button icon="colLeft" label={t('Insert Left')} title={t("Insert a column to the left of the caret's")} onClick={() => tableOp('insertColumnLeft')} />
              <Button icon="colRight" label={t('Insert Right')} title={t("Insert a column to the right of the caret's")} onClick={() => tableOp('insertColumnRight')} />
            </Rows>
          </Group>
          <Group label={t('Merge')}>
            <Button tall icon="mergeCells" label={t('Merge Cells')} title={acrossCells ? t('Merge the selected cells into one') : t('Merge Cells — select from one cell to another first')} disabled={!acrossCells} onClick={() => tableOp('mergeCells')} />
            <Button tall icon="splitCells" label={t('Split Cells')} title={t("Split Cells — a merged cell back into the cells it covers, or the caret's cell into columns or rows")} onClick={(e) => mergedCaret ? tableOp('splitCell') : menu.open(e, [
              { label: t('Into 2 columns'), run: () => tableOp('splitInto', { columns: 2 }) },
              { label: t('Into 3 columns'), run: () => tableOp('splitInto', { columns: 3 }) },
              { label: t('Into 2 rows'), run: () => tableOp('splitInto', { rows: 2 }) },
              { label: t('Into 2 columns and 2 rows'), run: () => tableOp('splitInto', { columns: 2, rows: 2 }) },
              { label: t('A merged cell back into its cells'), run: () => tableOp('splitCell') },
            ])} />
            <Button tall icon="splitCells" label={t('Split Table')} title={t("Split Table — the caret's row starts a table of its own, an empty paragraph between")} onClick={() => tableOp('splitTable')} />
          </Group>
          <Group label={t('Cell Size')}>
            <Rows>
              <div className="wd-fields" title={t("The caret's column width, in centimetres")}>
                <label>{t('Width')}</label>
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
                <span className="wd-field-value" style={{ minWidth: 0 }}>{t('cm')}</span>
              </div>
            </Rows>
            <Button tall icon="table" label={t('AutoFit')} title={t('AutoFit — the table to the window, to fixed column widths, or to its contents')} onClick={(e) => menu.open(e, [
              { label: t('AutoFit Contents'), run: () => tableOp('autoFit', { mode: 'contents' }) },
              { label: t('AutoFit Window'), run: () => tableOp('autoFit', { mode: 'window' }) },
              { label: t('Fixed Column Width'), run: () => tableOp('autoFit', { mode: 'fixed' }) },
            ])} />
            <Button tall icon="sliders" label={t('Distribute Columns')} title={t('Distribute Columns — every column the same width, the table as wide as before')} onClick={() => tableOp('distributeColumns')} />
          </Group>
          <Group label={t('Alignment')}>
            <Button tall icon="rotate" label={t('Text Direction')} title={t("Text Direction — the selected cells' words across, or turned to read down or up")} onClick={(e) => menu.open(e, [
              { label: t('Horizontal'), run: () => tableOp('textDirection', { dir: null }) },
              { label: t('Rotate all text 90°'), run: () => tableOp('textDirection', { dir: 'down' }) },
              { label: t('Rotate all text 270°'), run: () => tableOp('textDirection', { dir: 'up' }) },
            ])} />
            <Button tall icon="crop" label={t('Cell Margins')} title={t('Cell Margins — the room inside every cell of the table')} onClick={(e) => menu.open(e, [
              { label: t('Narrow'), run: () => tableOp('cellMargins', { margins: { top: 0, left: 57, bottom: 0, right: 57 } }) },
              { label: t('Normal'), run: () => tableOp('cellMargins', { margins: { top: 0, left: 108, bottom: 0, right: 108 } }) },
              { label: t('Wide'), run: () => tableOp('cellMargins', { margins: { top: 72, left: 216, bottom: 72, right: 216 } }) },
            ])} />
            <Button tall icon="alignCenter" label={t('Cell Alignment')} title={t('Cell Alignment — the words at the top, centre or bottom of the selected cells; now {align}', { align: { top: t('top'), center: t('centre'), bottom: t('bottom') }[cellAlign] || cellAlign })} onClick={(e) => menu.open(e, [
              { label: t('Top'), icon: cellAlign === 'top' ? 'check' : undefined, run: () => tableOp('cellVAlign', { v: 'top' }) },
              { label: t('Centre'), icon: cellAlign === 'center' ? 'check' : undefined, run: () => tableOp('cellVAlign', { v: 'center' }) },
              { label: t('Bottom'), icon: cellAlign === 'bottom' ? 'check' : undefined, run: () => tableOp('cellVAlign', { v: 'bottom' }) },
            ])} />
          </Group>
          <Group label={t('Direction')}>
            <Button tall icon="textRtl" label={t('Right to Left')} title={t("Right to left — the table's columns run from the right, as an Arabic or Hebrew table's do")} pressed={table.rtl} onClick={() => tableOp('direction', { rtl: !table.rtl })} />
          </Group>
          <Group label={t('Data')}>
            <Button tall icon="sort" label={t('Sort')} title={t("Sort — the table's rows by the caret's column, its header row kept at the top")} onClick={(e) => menu.open(e, [
              { label: t('Sort A to Z by this column'), run: () => tableOp('sort', { descending: false }) },
              { label: t('Sort Z to A by this column'), run: () => tableOp('sort', { descending: true }) },
              '-',
              { label: t('A to Z, the first row a header'), run: () => tableOp('sort', { descending: false, header: true }) },
              { label: t('Z to A, the first row a header'), run: () => tableOp('sort', { descending: true, header: true }) },
            ])} />
            <Button tall icon="file" label={t('Convert to Text')} title={t('Convert to Text — each row a paragraph, its cells between tabs or commas')} onClick={(e) => menu.open(e, [
              { label: t('Separated by tabs'), run: () => tableOp('toText', { separator: 'tab' }) },
              { label: t('Separated by commas'), run: () => tableOp('toText', { separator: 'comma' }) },
              { label: t('Each cell its own paragraph'), run: () => tableOp('toText', { separator: 'paragraph' }) },
            ])} />
            <Button tall icon="formula" label={t('Formula')} title={t("Formula — a field in the caret's cell worked out from the numbers above or to the left of it, or from cells like B2 and B2:B4, as Word's =SUM(ABOVE)")} onClick={(e) => menu.open(e, [
              { label: t('Sum above'), run: () => tableOp('formula', { formula: '=SUM(ABOVE)' }) },
              { label: t('Sum left'), run: () => tableOp('formula', { formula: '=SUM(LEFT)' }) },
              { label: t('Average above'), run: () => tableOp('formula', { formula: '=AVERAGE(ABOVE)', format: '#,##0.00' }) },
              { label: t('Count above'), run: () => tableOp('formula', { formula: '=COUNT(ABOVE)' }) },
              { label: t('Largest above'), run: () => tableOp('formula', { formula: '=MAX(ABOVE)' }) },
              { label: t('Smallest above'), run: () => tableOp('formula', { formula: '=MIN(ABOVE)' }) },
              { label: t('Your own formula…'), run: () => openDialog('tableFormula') },
              { label: t('Update all formulas'), run: () => dispatch({ op: 'updateTableFormulas' }) },
            ])} />
            <Button tall icon="refresh" label={t('Repeat Header Rows')} pressed={Boolean(caretCell?.rowHeader)} title={t('Repeat Header Rows — the rows from the top through this one drawn again at the top of every page the table runs onto')} onClick={() => tableOp('headerRows', { on: !caretCell?.rowHeader })} />
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
            <Group label={t('3D Model Views')}>
              {rows.map((row, i) => (
                <Rows key={i}>
                  {row.map(([key, label, view]) => <Button key={key} icon={at(view) ? 'check' : 'shape'} label={t(label)} className="wd-model3d-view" data-view={key} pressed={at(view)} title={t('{label} — the model turned to show it from there', { label: t(label) })} onClick={() => act('model3dView', view)} />)}
                </Rows>
              ))}
            </Group>
            <Group label={t('Turn')}>
              <Rows>
                <Button icon="chevronLeft" label={t('Turn Left')} className="wd-model3d-turn-left" title={t('Turn Left — fifteen degrees about its upright')} onClick={() => act('model3dView', { turn: { yaw: -15, pitch: 0 } })} />
                <Button icon="chevronRight" label={t('Turn Right')} title={t('Turn Right — fifteen degrees about its upright')} onClick={() => act('model3dView', { turn: { yaw: 15, pitch: 0 } })} />
              </Rows>
              <Rows>
                <Button icon="chevronUp" label={t('Tip Back')} title={t('Tip Back — fifteen degrees, the top towards you')} onClick={() => act('model3dView', { turn: { yaw: 0, pitch: 15 } })} />
                <Button icon="chevronDown" label={t('Tip Forward')} title={t('Tip Forward — fifteen degrees, the bottom towards you')} onClick={() => act('model3dView', { turn: { yaw: 0, pitch: -15 } })} />
              </Rows>
            </Group>
            <Group label={t('Adjust')}>
              <Button tall icon="undo" label={t('Reset 3D Model')} className="wd-model3d-reset" title={t('Reset 3D Model — back to the view it was put in at')} onClick={() => act('model3dView', 'reset')} />
            </Group>
          </>
        );
      })() : null}

      {/* ── Picture Format (contextual) ───────────────────────────────────── */}
      {tab === 'pictureFormat' ? (
        <>
          {arrange}
          <Group label={t('Size')}>
            <Rows>
              {sizeBox(t('Height'), 'h')}
              {sizeBox(t('Width'), 'w')}
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── References ───────────────────────────────────────────────────── */}
      {tab === 'references' ? (
        <>
          <Group label={t('Table of Contents')}>
            <Button tall icon="listBullet" label={t('Table of Contents')} title={t('Table of Contents — a live field built from the headings, with page numbers')} onClick={() => act('tableOfContents')} />
            <Button icon="plus" label={t('Add Text')} title={t('Make this paragraph a heading, so it appears in the table of contents')} onClick={(e) =>
              menu.open(e, [
                { label: t('Do not show in table of contents'), run: () => para({ styleId: null }) },
                { label: t('Level 1 (Heading 1)'), run: () => para({ styleId: 'Heading1' }) },
                { label: t('Level 2 (Heading 2)'), run: () => para({ styleId: 'Heading2' }) },
                { label: t('Level 3 (Heading 3)'), run: () => para({ styleId: 'Heading3' }) },
              ])
            } />
            <Button icon="refresh" label={t('Update Table')} disabled={!model?.tableOfContents} title={t('Update Table — rebuild the entries from the current headings and pages')} onClick={() => act('updateTableOfContents')} />
            <Button icon="close" label={t('Remove Table of Contents')} disabled={!model?.tableOfContents} title={t('Remove Table of Contents')} onClick={() => act('removeTableOfContents')} />
          </Group>
          <Group label={t('Footnotes')}>
            <Button tall icon="file" label={t('Insert Footnote')} title={t('A raised number at the caret, and its words under the body')} onClick={() => act('insertNote', 'footnote')} />
            <Button icon="file" label={t('Insert Endnote')} title={t('A raised number at the caret, and its words at the end of the document')} onClick={() => act('insertNote', 'endnote')} />
            <Button icon="chevronDown" label={t('Next Footnote')} title={t('Go to the next footnote reference')} onClick={() => act('nextNote', 1)} />
            <Button icon="eye" label={t('Show Notes')} title={t('Scroll to the footnotes under the body')} onClick={() => act('showNotes')} />
          </Group>
          {references ? <CitationsGroup refs={references} menu={menu} /> : null}
          <Group label={t('Captions')}>
            <Button tall icon="textbox" label={t('Insert Caption')} title={t('Insert Caption — a label, a running number kept live by Update Fields, and your own words')} onClick={() => openDialog('caption')} />
            {references ? captionsExtra(references, <Button icon="link" label={t('Cross-reference')} title={t('Cross-reference — a REF field to a bookmark, its words kept live by Update Fields')} onClick={() => openDialog('crossReference')} />) : <Button icon="link" label={t('Cross-reference')} title={t('Cross-reference — a REF field to a bookmark, its words kept live by Update Fields')} onClick={() => openDialog('crossReference')} />}
          </Group>
          <Group label={t('Fields')}>
            <Button tall icon="refresh" label={t('Update Fields')} title={t("Update Fields (F9) — refresh every cross-reference to its bookmark's current words")} onClick={() => commands['field.update']?.run?.()} />
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
          <Group label={t('Proofing')}>
            <Button tall icon="check" label={t('Spelling')} pressed={review?.pane === 'editor'} title={t('Spelling (F7) — check the whole document from the caret in the Editor pane: body, tables, text boxes, notes, headers and footers')} onClick={() => review?.startSpelling()} />
            <Button tall icon="word" label={t('Thesaurus')} pressed={review?.pane === 'thesaurus'} title={t('Thesaurus (Shift+F7) — words of like meaning for the word at the caret, one of them put in its place')} onClick={() => review?.openThesaurus()} />
            <Button tall icon="listNumber" label={t('Word Count')} onClick={() => openDialog('wordCount')} />
          </Group>
          <Group label={t('Speech')}>
            <Button tall icon="volume" label={t('Read Aloud')} pressed={Boolean(view.reading)} onClick={() => act('readAloud')} />
          </Group>
          <Group label={t('Accessibility')}>
            <Button tall icon="shield" label={t('Check Accessibility')} pressed={review?.pane === 'accessibility'} title={t('Check Accessibility — alt text, headings, tables, contrast, links and spacing, with a fix for each')} onClick={() => review?.openAccessibility()} />
            <Button icon="textbox" label={t('Alt Text')} disabled={!picked} title={picked ? t('Alt Text — describe the selected picture for people who cannot see it') : t('Alt Text — click a picture first, then describe it')} onClick={() => picked && review?.openAltText({ block: picked.block, image: picked.image })} />
          </Group>
          <Group label={t('Language')}>
            <Soon tall icon="globe" label={t('Translate')} why={t('Translation is a network service; this suite makes no requests it has not declared.')} />
            <Button tall icon="globe" label={t('Language')} title={t('Language — mark the selected words as a language, or as not to be checked, for Spelling here and proofing in Office')} onClick={() => act('language')} />
          </Group>
          <Group label={t('Comments')}>
            <Button tall icon="reply" label={t('New Comment')} onClick={() => openDialog('comment')} />
            <Button
              tall
              icon="trash"
              label={t('Delete')}
              disabled={!comments.length}
              title={comments.length ? t("Delete — the comment on the caret's paragraph, or every comment in the document") : t('Delete — this document has no comments')}
              onClick={(e) => menu.open(e, [
                { label: t('Delete'), icon: 'trash', disabled: !comments.some((c) => c.blockIndex === (model?.selection?.focus?.block ?? -1)), run: () => act('deleteComment') },
                { label: t('Delete All Comments in Document'), icon: 'trash', run: () => act('deleteComment', 'all') },
              ])}
            />
            <Button icon="chevronUp" label={t('Previous')} disabled={!comments.length} onClick={() => act('comment', -1)} />
            <Button icon="chevronDown" label={t('Next')} disabled={!comments.length} onClick={() => act('comment', 1)} />
            <Button icon="listBullet" label={comments.length ? t('Show ({count})', { count: comments.length }) : t('Show Comments')} disabled={!comments.length} onClick={() => openDialog('comments')} />
          </Group>
          <Group label={t('Tracking')}>
            <Button tall icon="eye" label={t('Track Changes')} pressed={Boolean(model?.trackRevisions)} disabled={Boolean(model?.protection?.lockedTracking)} title={model?.protection?.lockedTracking ? t('Track Changes — locked on: the document is protected for tracked changes') : t('Track Changes — record every insertion and deletion as w:ins / w:del while you edit')} onClick={() => act('toggleTrackChanges')} />
            <Button icon="eye" label={MARKUP_LABELS[view.markupMode] || t('Simple Markup')} title={t('Display for Review — how tracked changes are shown')} onClick={(e) => menu.open(e, [
              { label: t('All Markup'), run: () => act('markupMode', 'all') },
              { label: t('Simple Markup'), run: () => act('markupMode', 'simple') },
              { label: t('No Markup'), run: () => act('markupMode', 'final') },
              { label: t('Original'), run: () => act('markupMode', 'original') },
            ])} />
            <Button icon="list" label={t('Reviewing Pane')} title={t('Every tracked change in this document')} disabled={!model?.blocks?.some((b) => b.tracked)} onClick={() => openDialog('tracked')} />
          </Group>
          <Group label={t('Changes')}>
            <Button tall icon="check" label={t('Accept')} disabled={!trackedHere || Boolean(model?.protection?.lockedTracking)} title={t('Accept — keep this change')} onClick={(e) => menu.open(e, [
              { label: t('Accept This Change'), run: () => act('acceptChanges') },
              { label: t('Accept All Changes'), run: () => act('acceptChanges', 'all') },
            ])} />
            <Button tall icon="close" label={t('Reject')} disabled={!trackedHere || Boolean(model?.protection?.lockedTracking)} title={t('Reject — undo this change')} onClick={(e) => menu.open(e, [
              { label: t('Reject This Change'), run: () => act('rejectChanges') },
              { label: t('Reject All Changes'), run: () => act('rejectChanges', 'all') },
            ])} />
            <Button icon="chevronUp" label={t('Previous')} disabled={!anyTracked} title={t('Previous tracked change')} onClick={() => act('nextChange', -1)} />
            <Button icon="chevronDown" label={t('Next')} disabled={!anyTracked} title={t('Next tracked change')} onClick={() => act('nextChange', 1)} />
          </Group>
          <Group label={t('Compare')}>
            <Button tall icon="copy" label={t('Compare')} title={t('Compare — two versions of a document side by side, as a new document with what changed marked as revisions')} onClick={() => act('compare')} />
          </Group>
          <Group label={t('Protect')}>
            <Button tall icon="lock" label={t('Restrict Editing')} pressed={Boolean(view.restrict)} title={model?.protection?.enforced ? t('Restrict Editing — the document is protected; the pane says what you may do and stops protection') : t('Restrict Editing — limit formatting and editing, with exceptions, and enforce it with an optional password')} onClick={() => act('restrictPane')} />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
      {tab === 'immersive' && view.immersive ? (
        <>
          <Group label={t('Immersive Reader')}>
            <Button
              tall
              icon="alignJustify"
              label={t('Column Width')}
              title={t('Column Width — how long the lines are')}
              onClick={(e) => menu.open(e, IR_WIDTHS.map(([k, label]) => ({ label, icon: view.immersive.width === k ? 'check' : undefined, run: () => act('immersiveSet', { width: k }) })))}
            />
            <Button
              tall
              icon="contrast"
              label={t('Page Color')}
              title={t('Page Color — the colour behind the words')}
              onClick={(e) => menu.open(e, IR_COLOURS.map(([k, label]) => ({ label, icon: view.immersive.colour === k ? 'check' : undefined, run: () => act('immersiveSet', { colour: k }) })))}
            />
            <Button
              tall
              icon="eye"
              label={t('Line Focus')}
              title={t('Line Focus — one, three or five lines in view at a time; the arrow keys move the words through them')}
              onClick={(e) => menu.open(e, IR_FOCUS.map(([n, label]) => ({ label, icon: view.immersive.focus === n ? 'check' : undefined, run: () => act('immersiveSet', { focus: n }) })))}
            />
            <Button tall icon="sliders" label={t('Text Spacing')} pressed={Boolean(view.immersive.spacing)} title={t('Text Spacing — more room between letters, words and lines')} onClick={() => act('immersiveSet', { spacing: !view.immersive.spacing })} />
            <Button tall icon="scissors" label={t('Syllables')} pressed={Boolean(view.immersive.syllables)} title={t('Syllables — long words shown in their syllables, a dot between each')} onClick={() => act('immersiveSet', { syllables: !view.immersive.syllables })} />
            <Button tall icon="volume" label={t('Read Aloud')} pressed={Boolean(view.reading)} onClick={() => act('readAloud')} />
          </Group>
          <Group label={t('Close')}>
            <Button tall icon="close" label={t('Close Immersive Reader')} onClick={() => act('immersiveClose')} />
          </Group>
        </>
      ) : null}

      {tab === 'view' ? (
        <>
          <Group label={t('Views')}>
            <Button tall icon="eye" label={t('Read Mode')} pressed={view.mode === 'read'} onClick={() => act('mode', 'read')} />
            <Button tall icon="file" label={t('Print Layout')} pressed={!view.mode || view.mode === 'print'} onClick={() => act('mode', 'print')} />
            <Button tall icon="globe" label={t('Web Layout')} pressed={view.mode === 'web'} onClick={() => act('mode', 'web')} />
            <Button icon="listBullet" label={t('Outline')} pressed={view.mode === 'outline'} onClick={() => act('mode', 'outline')} />
            <Button icon="textbox" label={t('Draft')} pressed={view.mode === 'draft'} onClick={() => act('mode', 'draft')} />
          </Group>
          <Group label={t('Immersive')}>
            <Button tall icon="maximize" label={t('Focus')} pressed={Boolean(view.focus)} title={t('Just the page, full screen')} onClick={() => act('focus')} />
            <Button tall icon="eye" label={t('Immersive Reader')} pressed={Boolean(view.immersive)} title={t('Immersive Reader — the words in a column of their own, with line focus, wider spacing, syllables and a page colour that is easier to read')} onClick={() => act(view.immersive ? 'immersiveClose' : 'immersive')} />
          </Group>
          <Group label={t('Show')}>
            <Button icon="minus" label={t('Ruler')} pressed={Boolean(view.ruler)} onClick={() => act('toggleRuler')} />
            <Button icon="grid" label={t('Gridlines')} pressed={Boolean(view.gridlines)} title={t('Gridlines — a quarter-inch grid over the page, to line drawings up by eye; on screen only')} onClick={() => act('toggleGridlines')} />
            <Button icon="list" label={t('Navigation Pane')} pressed={Boolean(view.navigation)} title={t('Headings, to move around a long document')} onClick={() => act('toggleNavigation')} />
          </Group>
          <Group label={t('Zoom')}>
            <Button tall icon="find" label={t('Zoom')} onClick={(e) => menu.open(e, [50, 75, 100, 125, 150, 200].map((z) => ({ label: `${z}%`, run: () => act('zoom', z / 100) })))} />
            <Button icon="check" label="100%" onClick={() => act('zoom', 1)} />
            <Button icon="file" label={t('One Page')} onClick={() => act('zoom', 'page')} />
            <Button icon="copy" label={t('Multiple Pages')} onClick={() => act('zoom', 'pages')} />
            <Button icon="maximize" label={t('Page Width')} onClick={() => act('zoom', 'width')} />
          </Group>
          <Group label={t('Window')}>
            <Button tall icon="new" label={t('New Window')} title={t('This document again, in another window')} onClick={() => act('newWindow')} />
            <Button icon="grid" label={t('Arrange All')} title={t('Arrange All — every Documents window, one above another')} onClick={() => act('arrange', 'stack')} />
            <Button icon="minus" label={t('Split')} pressed={Boolean(view.split)} title={view.split ? t('Remove Split — the document in one pane again') : t('Split — the document in two panes, one over the other, each scrolled on its own')} onClick={() => act('toggleSplit')} />
            <Button icon="copy" label={t('Side by Side')} title={t('View Side by Side — this document and the one before it, half the screen each')} onClick={() => act('arrange', 'sideBySide')} />
          </Group>
          <Group label={t('Macros')}>
            <Soon tall icon="settings" label={t('Macros')} why={t('A macro engine is not built, and VBA in a file is preserved untouched rather than run.')} />
          </Group>
          <Group label={t('Properties')}>
            <Button tall icon="info" label={t('Properties')} onClick={() => openDialog('properties')} />
          </Group>
        </>
      ) : null}

      {/* ── Help ─────────────────────────────────────────────────────────── */}
      {tab === 'help' ? (
        <>
          <Group label={t('Help')}>
            <Button tall icon="info" label={t('Help')} onClick={() => act('open', 'help')} />
            <Button tall icon="reply" label={t('Feedback')} onClick={() => act('open', 'contact')} />
            <Button tall icon="globe" label={t("What's New")} onClick={() => act('open', 'releases')} />
            <Button tall icon="download" label={t('Check for Updates')} onClick={() => act('open', 'about')} />
          </Group>
          <Group label={t('Shortcuts')}>
            <Button tall icon="listBullet" label={t('Keyboard Shortcuts')} onClick={() => openDialog('shortcuts')} />
          </Group>
        </>
      ) : null}

      {/* ── PDF ──────────────────────────────────────────────────────────── */}
      {tab === 'pdf' ? (
        <>
          <Group label={t('Create')}>
            <Button tall icon="pdf" label={t('Create a PDF')} onClick={() => exportAs('pdf')} />
            {/* words-ok: a file format's name */}
            <Button tall icon="export" label="Markdown" onClick={() => exportAs('md')} />
            <Button tall icon="export" label={t('Plain Text')} onClick={() => exportAs('txt')} />
            <Button tall icon="export" label={t('Web Page')} onClick={() => exportAs('html')} />
          </Group>
          <Group label={t('Print')}>
            <Button tall icon="print" label={t('Print')} onClick={() => commands['file.print']?.run?.()} />
          </Group>
          <Group label={t('File')}>
            <Button tall icon="new" label={t('New')} onClick={() => shell.win.create({ app: 'word' })} />
            <Button tall icon="open" label={t('Open')} onClick={openFile} />
            <Button tall icon="save" label={t('Save')} onClick={() => save(false)} />
            <Button tall icon="save" label={t('Save As')} onClick={() => save(true)} />
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
