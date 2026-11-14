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
import { Ribbon, Group, Rows, Button, Separator, Select, Input, Gallery, t, msg } from '@rutba/office-ui';
import { EQUATION_GALLERY } from '@rutba/ooxml/math-linear';
import { THEMES, PALETTES, FONT_PAIRS, EFFECT_PRESETS } from '@rutba/office-formats/themes';
import { MailingsTab } from './mailings.js';
import { CitationsGroup } from './references.js';
import { IndexGroup } from './references-index.js';
import { PEN_COLOURS, PEN_WIDTHS } from '../../ink-geometry.js';
import { ToaGroup } from './references-toa.js';
import { captionsExtra } from './references-figures.js';
import { MODEL_VIEWS } from '@rutba/imaging/model3d';
import { untrackedOk } from './untracked.js';
import { DeveloperTab } from './forms.js';
import { ColourSplit } from '../colour/palette.js';

/* ── vocabularies ────────────────────────────────────────────────────────── */

const FONTS = ['Calibri', 'Calibri Light', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Tahoma', 'Garamond', 'Consolas', 'Courier New'];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
/** Word's three Arabic justifications, each stretching the letters with kashida a little more. */
const KASHIDA = [['lowKashida', msg('word.ribbon.justify-low')], ['mediumKashida', msg('word.ribbon.justify-medium')], ['highKashida', msg('word.ribbon.justify-high')]];

// Immersive Reader's choices, as Word names them.
const IR_WIDTHS = [['veryNarrow', t('word.ribbon.very-narrow')], ['narrow', t('common.narrow')], ['moderate', t('word.ribbon.moderate')], ['wide', t('common.wide')]];
const IR_COLOURS = [['none', t('common.none')], ['sepia', t('common.sepia')], ['inverse', t('word.ribbon.inverse')]];
const IR_FOCUS = [[0, t('common.none')], [1, t('word.ribbon.one-line')], [3, t('word.ribbon.three-lines')], [5, t('word.ribbon.five-lines')]];

const SPACING = [[1, '1.0'], [1.15, '1.15'], [1.5, '1.5'], [2, '2.0'], [2.5, '2.5'], [3, '3.0']];
/** Word's half-point rule, for the Borders menu. */
const LINE = { style: 'single', widthPx: 1, colour: null, spacePt: 1 };
const SHADES = [['#FFF2CC', t('word.ribbon.light-yellow')], ['#DEEBF7', t('common.light-blue')], ['#E2EFDA', t('common.light-green')], ['#FCE4D6', t('word.ribbon.light-orange')], ['#EDEDED', t('common.light-grey')], [null, t('word.ribbon.no-colour-2')]];
const BORDERS = [
  [t('common.bottom-border'), { bottom: LINE }], [t('common.top-border'), { top: LINE }], [t('common.left-border'), { left: LINE }], [t('common.right-border'), { right: LINE }],
  [t('common.outside-borders'), { top: LINE, bottom: LINE, left: LINE, right: LINE }], [t('common.no-border'), null],
];
/** Page Borders: Word's Box setting in four lines, measured 24 pt in from the page edge as Word measures it. */
const FRAME = (style, widthPx) => ({ style, widthPx, colour: null, spacePt: 24 });
const BOX = (line) => ({ offsetFrom: 'page', top: line, left: line, bottom: line, right: line });
const PAGE_BORDERS = [
  [t('word.ribbon.box-thin-line'), BOX(FRAME('single', 1))], [t('word.ribbon.box-thick-line'), BOX(FRAME('single', 3))],
  [t('word.ribbon.box-double-line'), BOX(FRAME('double', 3))], [t('word.ribbon.box-dashed-line'), BOX(FRAME('dashed', 1))],
  [t('word.ribbon.no-page-border'), null],
];
/** Watermark: Word's own stock words, then a custom one. */
const WATERMARKS = [t('word.ribbon.draft'), t('word.ribbon.confidential'), t('word.ribbon.sample'), t('word.ribbon.do-not-copy'), t('word.ribbon.urgent')];
const PAGE_SIZES = [['A4', t('word.ribbon.a4-21-29')], ['Letter', t('word.ribbon.letter-8-5')], ['Legal', t('word.ribbon.legal-8-5')]];
/** Review → Display for Review: the button's own label per mode. */
const MARKUP_LABELS = { all: t('word.ribbon.all-markup'), simple: t('word.ribbon.simple'), final: t('word.ribbon.no-markup'), original: t('word.ribbon.original') };
const MARGINS = [['normal', t('word.ribbon.normal-2-54')], ['narrow', t('word.ribbon.narrow-1-27')], ['wide', t('word.ribbon.wide-5-08')]];
const SHAPES = [
  ['rect', t('common.rectangle')], ['roundRect', t('common.rounded-rectangle')], ['ellipse', t('common.oval')], ['triangle', t('common.triangle')], ['diamond', t('common.diamond')],
  ['rightArrow', t('word.ribbon.arrow-right')], ['leftArrow', t('word.ribbon.arrow-left')], ['upArrow', t('word.ribbon.arrow-up')], ['downArrow', t('word.ribbon.arrow-down')],
  ['pentagon', t('common.pentagon')], ['hexagon', t('common.hexagon')], ['star5', t('word.ribbon.star-5-points')], ['plus', t('common.cross')], ['chevron', t('common.chevron')],
  ['parallelogram', t('common.parallelogram')], ['trapezoid', t('common.trapezoid')], ['line', t('common.line')],
];
const CHARTS = [['column', t('common.column')], ['bar', t('common.bar')], ['line', t('common.line')], ['pie', t('common.pie')]];
const CASES = [['sentence', t('word.ribbon.sentence-case')], ['lower', t('common.lowercase-2')], ['upper', t('common.uppercase')], ['title', t('word.ribbon.capitalize-each-word')], ['toggle', t('word.ribbon.t-oggle-c')]];
/** The "A" button's glow colours — Word's own accent palette, 4pt radius. */
const GLOWS = [['FFC000', t('word.ribbon.glow-gold')], ['4472C4', t('word.ribbon.glow-blue')], ['70AD47', t('word.ribbon.glow-green')], ['FF0000', t('word.ribbon.glow-red')]];

/** Shape Fill and Shape Outline: Word's standard colours, in the suite's own names. */
const SHAPE_COLOURS = [
  ['FFFFFF', t('common.white')], ['000000', t('common.black')], ['E7E6E6', t('common.light-grey')], ['767171', t('common.grey')], ['1F3864', t('common.dark-blue')],
  ['2B5FD9', t('common.blue')], ['DEEBF7', t('common.light-blue')], ['0D8F6F', t('common.teal')], ['0F9D58', t('common.green')], ['E2EFDA', t('common.light-green')],
  ['E0A800', t('common.gold')], ['FFF2CC', t('word.ribbon.light-yellow')], ['E08B2B', t('common.orange')], ['C00000', t('common.dark-red')], ['7B5CD6', t('common.purple')],
];
/** Shape Outline → Weight, in Word's points and the pixels the engine writes. */
const WEIGHTS = [[t('word.ribbon.1-2-pt'), 0.67], [t('word.ribbon.1-pt'), 1.33], [t('word.ribbon.11-2-pt'), 2], [t('word.ribbon.21-4-pt'), 3], [t('word.ribbon.3-pt'), 4], [t('word.ribbon.6-pt'), 8]];
/** A text box's Margins, Word's four: its insets in px, left/top/right/bottom. */
const BOX_MARGINS = [
  [t('common.none'), { l: 0, t: 0, r: 0, b: 0 }],
  [t('common.narrow'), { l: 4.8, t: 4.8, r: 4.8, b: 4.8 }],
  [t('common.normal'), { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }],
  [t('common.wide'), { l: 14.4, t: 14.4, r: 14.4, b: 14.4 }],
];
/** Home → Multilevel List's library, as Word's gallery shows each: its first three levels. */
const MULTILEVEL = [
  ['outline', '1.  1.1.  1.1.1.'],
  ['outlineParen', '1)  a)  i)'],
  ['outlineRoman', 'I.  A.  1.'],
  ['legal', t('word.ribbon.article-section-1')],
  ['bullet', '•  ○  ▪'],
];

/** Insert → WordArt: four styles of our own, made of the text effects the engine writes. */
const WORDART = [
  [t('word.ribbon.fill-blue-shadow'), { colour: '2B5FD9', effects: { shadow: true } }],
  [t('word.ribbon.outline-blue'), { colour: '2B5FD9', effects: { outline: true } }],
  [t('word.ribbon.fill-gold-glow'), { colour: 'C98A00', effects: { glow: { colour: 'FFC000', radiusPt: 4 } } }],
  [t('word.ribbon.fill-black-shadow'), { colour: '262626', effects: { shadow: true } }],
];
/** Arrange → Position's nine places with the words wrapped round, by `${vertical}-${horizontal}`. */
const POSITIONS = {
  'top-left': t('word.ribbon.top-left'), 'top-center': t('word.ribbon.top-centre'), 'top-right': t('word.ribbon.top-right'),
  'middle-left': t('word.ribbon.middle-left'), 'middle-center': t('word.ribbon.middle-centre'), 'middle-right': t('word.ribbon.middle-right'),
  'bottom-left': t('word.ribbon.bottom-left'), 'bottom-center': t('word.ribbon.bottom-centre'), 'bottom-right': t('word.ribbon.bottom-right'),
};
/** 1 cm in CSS px, for the Size boxes. */
const CM = 96 / 2.54;

/** Table Design's pen: Word's line styles, and its weights in eighths of a point. */
const PEN_STYLES = [['single', t('word.ribbon.single')], ['double', t('word.ribbon.double')], ['dotted', t('word.ribbon.dotted')], ['dashed', t('word.ribbon.dashed')], ['thick', t('word.ribbon.thick')]];
const PEN_WEIGHTS = [[2, t('word.ribbon.1-4-pt')], [4, t('word.ribbon.1-2-pt')], [6, t('word.ribbon.3-4-pt')], [8, t('word.ribbon.1-pt')], [12, t('word.ribbon.11-2-pt')], [18, t('word.ribbon.21-4-pt')], [24, t('word.ribbon.3-pt')], [36, t('word.ribbon.41-2-pt')], [48, t('word.ribbon.6-pt')]];


/** A style's name as Word shows it: its built-in names are kept in lower case in the file ("heading 1"), shown capitalised. */
const styleName = (s) => String(s.name || s.id).replace(/(^|\s)\p{Ll}/gu, (m) => m.toUpperCase());

export default function WordRibbon({
  tab, setTab, doc, model, dispatch, commands, shell, menu, save, openFile, exportAs, openDialog, modifyStyle = null, insertPicture, act, view = {}, picked = null, mailings = null, review = null, drawing = null, references = null, ink = null, blocks = null, table = null,
  forms = null, redaction = null,
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
  const needDrawing = t('word.ribbon.click-picture-shape');
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
          title={t('word.ribbon.label-cm', { label })}
          style={{ width: 64 }}
          onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }}
          onBlur={(e) => {
            const v = Number(e.target.value);
            if (!(v > 0) || Math.abs(v - cm) < 0.005) return;
            act('drawingSize', key === 'w' ? { widthPx: Math.round(v * CM) } : { heightPx: Math.round(v * CM) });
          }}
        />
        <span className="wd-field-value" style={{ minWidth: 0 }}>{t('word.ribbon.cm')}</span>
      </div>
    );
  };
  const arrange = (
    <Group label={t('common.arrange')}>
      <Button dropdown tall icon="grid" label={t('common.position')} disabled={!hasDrawing} title={hasDrawing ? t('word.ribbon.position-drawing-sits') : needDrawing} onClick={(e) => menu.open(e, [
        { label: t('word.ribbon.left-words-round'), icon: 'alignLeft', run: () => act('position', 'left') },
        { label: t('word.ribbon.centre-words-above'), icon: 'alignCenter', run: () => act('position', 'center') },
        { label: t('word.ribbon.right-words'), icon: 'alignRight', run: () => act('position', 'right') },
        { heading: true, label: t('word.ribbon.with-text') },
        ...['top', 'middle', 'bottom'].flatMap((v) => ['left', 'center', 'right'].map((h) => ({ label: POSITIONS[`${v}-${h}`], run: () => act('position', `${v}-${h}`) }))),
      ])} />
      <Button dropdown tall icon="picture" label={t('common.wrap-text')} disabled={!hasDrawing} title={hasDrawing ? t('word.ribbon.wrap-text-how') : needDrawing} onClick={(e) => menu.open(e, [
        { label: t('word.line-text'), icon: drawing && !drawing.anchored ? 'check' : undefined, run: () => act('wrap', 'inline') },
        { label: t('common.square'), icon: drawing?.anchored && drawing.wrap === 'square' ? 'check' : undefined, run: () => act('wrap', 'square') },
        { label: t('common.tight'), icon: drawing?.anchored && drawing.wrap === 'tight' ? 'check' : undefined, run: () => act('wrap', 'tight') },
        { label: t('word.ribbon.through'), icon: drawing?.anchored && drawing.wrap === 'through' ? 'check' : undefined, run: () => act('wrap', 'through') },
        { label: t('word.ribbon.top-and'), icon: drawing?.anchored && drawing.wrap === 'topAndBottom' ? 'check' : undefined, run: () => act('wrap', 'topAndBottom') },
        '-',
        { label: t('word.ribbon.behind-text'), icon: drawing?.anchored && drawing.wrap === 'none' && drawing.behind ? 'check' : undefined, run: () => act('wrap', 'behind') },
        { label: t('word.ribbon.front-text'), icon: drawing?.anchored && drawing.wrap === 'none' && !drawing.behind ? 'check' : undefined, run: () => act('wrap', 'front') },
      ])} />
      <Button dropdown tall icon="chevronUp" label={t('common.bring')} disabled={!floatingSel.length} title={floatingSel.length ? t('word.ribbon.bring-forward-one') : t('word.ribbon.bring-forward-select')} onClick={(e) => menu.open(e, [
        { label: t('common.bring'), icon: 'chevronUp', run: () => act('order', 'forward') },
        { label: t('common.bring-to'), run: () => act('order', 'front') },
        { label: t('word.ribbon.bring-front'), run: () => act('order', 'inFront') },
      ])} />
      <Button dropdown tall icon="chevronDown" label={t('common.send-backward')} disabled={!floatingSel.length} title={floatingSel.length ? t('word.ribbon.send-backward-one') : t('word.ribbon.send-backward-select')} onClick={(e) => menu.open(e, [
        { label: t('common.send-backward'), icon: 'chevronDown', run: () => act('order', 'backward') },
        { label: t('common.send-to-back'), run: () => act('order', 'back') },
        { label: t('word.ribbon.send-behind'), run: () => act('order', 'behind') },
      ])} />
      <Rows>
        <Button icon="list" label={t('common.selection-pane')} pressed={Boolean(view.selectionPane)} title={t('word.ribbon.selection-pane-every')} onClick={() => act('selectionPane')} />
        <>
          <Button dropdown icon="alignLeft" label={t('common.align')} disabled={!floatingSel.length} title={floatingSel.length ? (alignMode === 'selected' ? t('word.ribbon.align-selected') : alignMode === 'page' ? t('word.ribbon.align-page') : t('word.ribbon.align-margin')) : t('word.ribbon.align-select-floating')} onClick={(e) => menu.open(e, [
            { label: t('common.align-left'), icon: 'alignLeft', run: () => act('align', { edge: 'left', to: alignMode }) },
            { label: t('word.ribbon.align-centre'), icon: 'alignCenter', run: () => act('align', { edge: 'center', to: alignMode }) },
            { label: t('common.align-right'), icon: 'alignRight', run: () => act('align', { edge: 'right', to: alignMode }) },
            { label: t('common.align-top'), run: () => act('align', { edge: 'top', to: alignMode }) },
            { label: t('common.align-middle'), run: () => act('align', { edge: 'middle', to: alignMode }) },
            { label: t('common.align-bottom'), run: () => act('align', { edge: 'bottom', to: alignMode }) },
            '-',
            { label: t('common.distribute-horizontally'), disabled: floatingSel.length < 3, title: floatingSel.length < 3 ? t('word.ribbon.select-three-more') : undefined, run: () => act('distribute', { axis: 'horizontal' }) },
            { label: t('common.distribute'), disabled: floatingSel.length < 3, title: floatingSel.length < 3 ? t('word.ribbon.select-three-more') : undefined, run: () => act('distribute', { axis: 'vertical' }) },
            '-',
            { label: t('word.ribbon.align-to'), icon: alignMode === 'page' ? 'check' : undefined, run: () => setAlignTo('page') },
            { label: t('word.ribbon.align-to-margin'), icon: alignMode === 'margin' ? 'check' : undefined, run: () => setAlignTo('margin') },
            { label: t('common.align-selected'), icon: alignMode === 'selected' ? 'check' : undefined, disabled: floatingSel.length < 2, run: () => setAlignTo('selected') },
          ])} />
          <Button dropdown icon="grid" label={t('common.group')} disabled={floatingSel.length < 2 && !isGroup} title={floatingSel.length >= 2 ? t('word.ribbon.group-selected-drawings') : isGroup ? t('word.ribbon.group-ungroup') : t('word.ribbon.group-select-two')} onClick={(e) => menu.open(e, [
            { label: t('common.group'), icon: 'grid', disabled: floatingSel.length < 2, run: () => act('group') },
            { label: t('common.ungroup'), disabled: !isGroup, run: () => act('ungroup') },
          ])} />
        </>
        <Button dropdown icon="rotate" label={t('common.rotate')} disabled={!hasDrawing} title={hasDrawing ? t('word.ribbon.rotate-turn-flip') : needDrawing} onClick={(e) => menu.open(e, [
          { label: t('common.rotate-right-90'), icon: 'rotate', run: () => act('rotate', 90) },
          { label: t('common.rotate-left-90'), run: () => act('rotate', -90) },
          { label: t('common.flip'), icon: 'flip', run: () => act('rotate', 'flipV') },
          { label: t('common.flip-horizontal'), run: () => act('rotate', 'flipH') },
          { label: t('common.more-rotation'), run: () => act('rotationDialog') },
        ])} />
      </Rows>
    </Group>
  );

  const run = (delta) => dispatch({ op: 'setRunFormat', delta });
  const para = (delta) => dispatch({ op: 'setParagraphFormat', delta });
  // A style's right-click, as Word's: the style takes the selection's look, or
  // its Modify dialog opens. Either changes every paragraph with the style.
  const styleMenu = (s) => [
    { label: t('word.ribbon.update-style-match', { style: styleName(s) }), icon: 'refresh', run: () => dispatch({ op: 'updateStyleToMatch', styleId: s.id }) },
    ...(modifyStyle ? [{ label: t('common.modify'), icon: 'styles', run: () => modifyStyle({ id: s.id, name: styleName(s) }) }] : []),
  ];
  const toggleList = (kind) => para({ list: format.listType === kind ? null : kind });
  const size = Number(format.fontSize || 11);
  const nearer = (dir) => {
    const at = SIZES.findIndex((s) => s >= size);
    const next = dir > 0 ? SIZES[Math.min(SIZES.length - 1, (at < 0 ? SIZES.length - 1 : at) + 1)] : SIZES[Math.max(0, (at < 0 ? SIZES.length : at) - 1)];
    return next ?? size;
  };

  return (
    <Ribbon
      tabs={[
        { id: 'home', label: t('common.home') },
        { id: 'insert', label: t('common.insert') },
        { id: 'draw', label: t('common.draw') },
        { id: 'design', label: t('common.design') },
        { id: 'layout', label: t('common.layout') },
        { id: 'references', label: t('word.references') },
        { id: 'mailings', label: t('word.ribbon.mailings') },
        { id: 'review', label: t('common.review') },
        { id: 'view', label: t('common.view') },
        { id: 'developer', label: t('word.ribbon.developer') },
        { id: 'help', label: t('common.help') },
        { id: 'pdf', label: 'PDF' }, // words-ok: a file format's name
        // The contextual tab, as Word's: there while a drawing is selected.
        ...(formatTab === 'shapeFormat' ? [{ id: 'shapeFormat', label: t('common.shape-format'), contextual: true }] : []),
        ...(formatTab === 'pictureFormat' ? [{ id: 'pictureFormat', label: t('common.picture-format'), contextual: true }] : []),
        // 3D Model, as Word's: there while a 3D model is picked.
        ...(pickedModel ? [{ id: 'model3d', label: t('common.3d-model'), contextual: true }] : []),
        // Table Design and Table Layout, as Word's: there while the caret is in a table.
        ...(table ? [{ id: 'tableDesign', label: t('common.table-design'), contextual: true }, { id: 'tableLayout', label: t('word.ribbon.table-layout'), contextual: true }] : []),
        // View → Immersive Reader's own tab, there while the reader is open.
        ...(view.immersive ? [{ id: 'immersive', label: t('word.ribbon.immersive-reader'), contextual: true }] : []),
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title={t('common.save-ctrl-s')} onClick={() => save(false)} />
          <Button icon="undo" title={t('common.undo-ctrl-z')} disabled={!doc?.canUndo} onClick={() => commands['edit.undo'].run()} />
          <Button icon="redo" title={t('common.redo-ctrl-y')} disabled={!doc?.canRedo} onClick={() => commands['edit.redo'].run()} />
          <Button icon="print" title={t('common.print')} onClick={() => commands['file.print']?.run?.()} />
        </>
      }
    >
      {/* ── Home ─────────────────────────────────────────────────────────── */}
      {tab === 'home' ? (
        <>
          <Group label={t('common.clipboard')}>
            <Button tall icon="paste" label={t('common.paste')} title={t('common.paste-ctrl-v')} onClick={() => commands['edit.paste']?.run?.()} />
            <Rows>
              <>
                <Button icon="cut" label={t('common.cut')} title={t('word.ribbon.cut-ctrl-x')} onClick={() => commands['edit.cut']?.run?.()} />
                <Button icon="copy" label={t('common.copy')} title={t('word.ribbon.copy-ctrl-c')} onClick={() => commands['edit.copy']?.run?.()} />
              </>
              <Button icon="brush" label={t('common.format-painter')} pressed={Boolean(view.painting)} title={t('word.ribbon.format-painter-copy')} onClick={() => act('formatPainter')} />
            </Rows>
          </Group>

          <Group label={t('common.font')}>
            <Rows>
              <>
                <Select value={format.fontName || 'Calibri'} onChange={(e) => run({ fontName: e.target.value })} style={{ width: 132 }} title={t('common.font')}>
                  {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
                  {format.fontName && !FONTS.includes(format.fontName) ? <option value={format.fontName}>{format.fontName}</option> : null}
                </Select>
                <Select value={String(size)} onChange={(e) => run({ fontSize: Number(e.target.value) })} style={{ width: 58 }} title={t('common.size')}>
                  {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
                  {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
                </Select>
                <Button icon="fontGrow" title={t('word.ribbon.grow-font')} onClick={() => run({ fontSize: nearer(1) })} />
                <Button icon="fontShrink" title={t('word.ribbon.shrink-font')} onClick={() => run({ fontSize: nearer(-1) })} />
                <Button dropdown icon="changeCase" title={t('common.change-case')} onClick={(e) => menu.open(e, CASES.map(([mode, label]) => ({ label, run: () => act('changeCase', mode) })))} />
                <Button icon="clearFormat" title={t('word.ribbon.clear-all-formatting')} onClick={() => dispatch({ op: 'clearFormat' })} />
              </>
              <>
                <Button icon="bold" title={t('common.bold-ctrl-b')} pressed={format.bold} onClick={() => dispatch({ op: 'toggleFormat', tag: 'bold' })} />
                <Button icon="italic" title={t('common.italic-ctrl-i')} pressed={format.italic} onClick={() => dispatch({ op: 'toggleFormat', tag: 'italic' })} />
                <Button icon="underline" title={t('common.underline-ctrl-u')} pressed={format.underline} onClick={() => dispatch({ op: 'toggleFormat', tag: 'underline' })} />
                <Button icon="strike" title={t('common.strikethrough')} pressed={format.strike} onClick={() => dispatch({ op: 'toggleFormat', tag: 'strike' })} />
                <Button icon="subscript" title={t('common.subscript')} pressed={format.vertAlign === 'subscript'} onClick={() => run({ vertAlign: format.vertAlign === 'subscript' ? null : 'subscript' })} />
                <Button icon="superscript" title={t('common.superscript')} pressed={format.vertAlign === 'superscript'} onClick={() => run({ vertAlign: format.vertAlign === 'superscript' ? null : 'superscript' })} />
                <Separator />
                <Button dropdown icon="textEffects" title={t('word.ribbon.text-effects-outline')} onClick={(e) => menu.open(e, [
                  { label: t('common.outline'), icon: format.outline ? 'check' : undefined, run: () => run({ outline: !format.outline }) },
                  { label: t('common.shadow'), icon: format.shadow ? 'check' : undefined, run: () => run({ shadow: !format.shadow }) },
                  '-',
                  ...GLOWS.map(([colour, label]) => ({ label, icon: format.glow?.colour === colour ? 'check' : undefined, run: () => run({ glow: { colour, radiusPt: 4 } }) })),
                  { label: t('common.no-glow'), run: () => run({ glow: null }) },
                  '-',
                  { label: t('word.ribbon.clear'), run: () => run({ outline: false, shadow: false, glow: null }) },
                ])} />
                {/* Word's two split buttons: the face puts on the colour last chosen, the arrow opens the palette. */}
                <ColourSplit kind="word.highlight" icon="highlighter" name={t('common.text-highlight')} highlight first="yellow" current={format.highlight ?? null}
                  pressed={Boolean(format.highlight)} empty={{ label: t('common.no-highlight'), value: null }} onPick={(value) => run({ highlight: value })} />
                <ColourSplit kind="word.font" icon="fontColour" name={t('common.font-colour')} first="FF0000" theme={design?.colors} current={format.fontColour ?? null}
                  empty={{ label: t('common.automatic'), value: null, swatch: '#000000' }} onPick={(value) => run({ fontColour: value })} />
              </>
            </Rows>
          </Group>

          <Group label={t('common.paragraph')}>
            <Rows>
              <>
                <Button icon="listBullet" title={t('common.bullets')} pressed={format.listType === 'bullet'} onClick={() => toggleList('bullet')} />
                <Button icon="listNumber" title={t('common.numbering')} pressed={format.listType === 'number'} onClick={() => toggleList('number')} />
                <Button dropdown icon="listMulti" className="wd-multilevel" title={t('word.ribbon.multilevel-list-list')} pressed={format.listType === 'number' && (format.listLevel || 0) > 0} onClick={(e) => menu.open(e, [
                  { heading: true, label: t('word.ribbon.list-library') },
                  ...MULTILEVEL.map(([list, label]) => ({ label, icon: 'listNumber', run: () => para({ list }) })),
                  '-',
                  { label: t('common.none'), icon: 'close', run: () => para({ list: null }) },
                  { label: t('word.ribbon.define-new'), icon: 'listNumber', run: () => openDialog('defineList') },
                ])} />
                <Separator />
                <Button icon="indentLess" title={t('common.decrease-indent')} disabled={!format.indentLevel} onClick={() => para({ indentDelta: -1 })} />
                <Button icon="indentMore" title={t('common.increase-indent')} onClick={() => para({ indentDelta: 1 })} />
                <Button dropdown icon="sort" title={t('word.ribbon.sort-paragraphs-z')} onClick={(e) => menu.open(e, [
                  { label: t('common.a-to-z'), run: () => dispatch({ op: 'sortParagraphs', descending: false }) },
                  { label: t('common.z-to-a'), run: () => dispatch({ op: 'sortParagraphs', descending: true }) },
                ])} />
                <Button icon="pilcrow" title={t('word.ribbon.show-formatting-marks')} pressed={Boolean(view.marks)} onClick={() => act('toggleMarks')} />
                <Separator />
                <Button icon="textLtr" title={t('common.left-right')} pressed={!format.rtl} onClick={() => para({ rtl: false })} />
                <Button icon="textRtl" title={t('word.ribbon.right-left-text')} pressed={Boolean(format.rtl)} onClick={() => para({ rtl: true })} />
              </>
              <>
                <Button icon="alignLeft" title={t('word.ribbon.align-left-ctrl')} pressed={format.paragraphAlign === 'left'} onClick={() => para({ align: 'left' })} />
                <Button icon="alignCenter" title={t('word.ribbon.centre-ctrl-e')} pressed={format.paragraphAlign === 'center'} onClick={() => para({ align: 'center' })} />
                <Button icon="alignRight" title={t('word.ribbon.align-right-ctrl')} pressed={format.paragraphAlign === 'right'} onClick={() => para({ align: 'right' })} />
                <Button icon="alignJustify" title={t('word.ribbon.justify-ctrl-j')} pressed={format.paragraphAlign === 'justify' || format.paragraphAlign === 'both'} onClick={() => para({ align: 'both' })} />
                {format.rtl ? (
                  // Word's Arabic justifications: the line filled by stretching the letters with kashida.
                  <Button icon="chevronDown" title={t('word.ribbon.justify-kashida-stretch')} pressed={KASHIDA.some(([v]) => v === format.paragraphAlign)} onClick={(e) =>
                    menu.open(e, KASHIDA.map(([v, label]) => ({ label: t(label), icon: format.paragraphAlign === v ? 'check' : undefined, run: () => para({ align: v }) })))} />
                ) : null}
                <Separator />
                <Button dropdown icon="listNumber" title={format.lineSpacing ? t('word.ribbon.line-paragraph-spacing', { spacing: format.lineSpacing }) : t('word.ribbon.line')} onClick={(e) =>
                  menu.open(e, [
                    ...SPACING.map(([v, l]) => ({ label: l, icon: format.lineSpacing === v ? 'check' : undefined, run: () => para({ lineSpacing: v }) })),
                    '-',
                    { label: t('word.ribbon.add-space-before'), run: () => para({ spaceBefore: 12 }) },
                    { label: t('word.ribbon.remove-space-before'), run: () => para({ spaceBefore: 0 }) },
                    { label: t('word.ribbon.add-space'), run: () => para({ spaceAfter: 8 }) },
                    { label: t('word.ribbon.remove-space'), run: () => para({ spaceAfter: 0 }) },
                  ])
                } />
                <Button dropdown icon="bucketFill" title={t('word.ribbon.shading-colour-behind')} onClick={(e) => menu.open(e, SHADES.map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => para({ shading: value }) })))} />
                <Button dropdown icon="borders" title={t('word.ribbon.borders-lines-round')} onClick={(e) => menu.open(e, BORDERS.map(([label, value]) => ({ label, icon: value ? undefined : 'close', run: () => para({ borders: value }) })))} />
              </>
            </Rows>
          </Group>

          <Group label={t('common.styles')} icon="styles">
            {/* Word's Styles gallery: as many as there is room for, the rest a page away on its arrows, every one under More. */}
            <Gallery
              className="wd-styles"
              max={8}
              min={3}
              label={t('word.ribbon.styles')}
              moreTip={t('word.ribbon.more-styles-every')}
              onMore={(e) => menu.open(e, (styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).map((s) => ({
                label: styleName(s),
                icon: (format.paragraphStyle || 'Normal') === s.id ? 'check' : undefined,
                run: () => para({ styleId: s.id === 'Normal' && !styles.some((x) => x.id === 'Normal') ? null : s.id }),
              })))}
            >
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).map((s) => (
                <button
                  key={s.id}
                  type="button"
                  data-style={s.id}
                  className={`wd-style${(format.paragraphStyle || 'Normal') === s.id ? ' on' : ''}`}
                  title={t('word.ribbon.apply-style', { style: s.name || s.id })}
                  onClick={() => para({ styleId: s.id === 'Normal' && !styles.some((x) => x.id === 'Normal') ? null : s.id })}
                  onContextMenu={(e) => menu.open(e, styleMenu(s))}
                  style={STYLE_LOOK[s.id] || undefined}
                >
                  <span className="wd-style-name">{s.name || s.id}</span>
                </button>
              ))}
            </Gallery>
          </Group>

          {/* One above another, as Word stacks them: a column as wide as Replace, so the group keeps its names when the window narrows. */}
          <Group label={t('common.editing')}>
            <Rows>
              <Button icon="find" label={t('common.find')} title={t('word.ribbon.find-ctrl-f')} onClick={() => openDialog('find')} />
              <Button icon="replace" label={t('common.replace')} title={t('word.ribbon.replace-ctrl-h')} onClick={() => openDialog('find')} />
              <Button icon="pointer" label={t('common.select')} title={t('word.ribbon.select-all-ctrl')} onClick={() => dispatch({ op: 'selectAll' })} />
            </Rows>
          </Group>

          <Group label={t('common.voice')}>
            <Button tall twoLine icon="volume" label={t('word.ribbon.read-aloud-2')} pressed={Boolean(view.reading)} title={t('word.ribbon.read-document-aloud')} onClick={() => act('readAloud')} />
          </Group>
        </>
      ) : null}

      {/* ── Insert ───────────────────────────────────────────────────────── */}
      {tab === 'insert' ? (
        <>
          <Group label={t('common.pages')}>
            <Button tall icon="file" label={t('word.ribbon.cover-page')} title={t('word.ribbon.title-page-front')} onClick={() => act('coverPage')} />
            <Button tall icon="file" label={t('word.ribbon.blank-page')} title={t('word.ribbon.page-break-another')} onClick={() => act('blankPage')} />
            <Button tall icon="file" label={t('word.ribbon.page-break')} title={t('word.ribbon.page-break-ctrl')} onClick={() => dispatch({ op: 'insertPageBreak' })} />
          </Group>
          <Group label={t('common.tables')}>
            <Button tall icon="table" label={t('common.table')} onClick={() => openDialog('table')} />
            <Button dropdown icon="table" label={t('word.ribbon.text-to')} disabled={Boolean(table)} title={table ? t('word.ribbon.convert-text-table') : t('word.ribbon.convert-text-table-selected')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.split-at'), run: () => dispatch({ op: 'textToTable', separator: 'tab' }) },
              { label: t('word.ribbon.split-at-commas'), run: () => dispatch({ op: 'textToTable', separator: 'comma' }) },
            ])} />
          </Group>
          <Group label={t('common.illustrations')}>
            <Button tall icon="picture" label={t('common.pictures')} onClick={insertPicture} />
            <Button dropdown tall icon="shape" label={t('common.shapes')} onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', preset, widthPx: 200, heightPx: 120 }) })))} />
            <Button tall icon="star" label={t('common.icons')} title={t('word.ribbon.icons-one-suites')} onClick={() => act('icons')} />
            <Button tall icon="qrcode" label={t('common.qr-code')} className="wd-qr-insert" title={t('word.ribbon.qr-code-code')} onClick={() => act('qrcode')} />
            <Button tall icon="shape" label={t('common.3d-models')} className="wd-model3d-insert" title={t('word.ribbon.3d-models-model')} onClick={() => act('model3d')} />
            <Button tall icon="grid" label={t('common.smart-art')} title={t('word.ribbon.smart-art-list')} onClick={() => act('insertSmartArt')} />
            <Button dropdown tall icon="chart" label={t('common.chart')} title={t('word.ribbon.chart-drawn-table')} onClick={(e) => menu.open(e, CHARTS.map(([k, l]) => ({ label: l, icon: 'chart', run: () => dispatch({ op: 'insertChart', kind: k }) })))} />
            <Button tall icon="picture" label={t('common.screenshot')} title={t('word.ribbon.screenshot-picture')} onClick={() => act('screenshot')} />
          </Group>
          <Group label={t('common.links')}>
            <Button tall icon="link" label={t('common.link')} title={t('word.ribbon.link-ctrl-k')} pressed={Boolean(format.link)} onClick={() => openDialog('link')} />
            <Button icon="flag" label={t('common.bookmark')} title={t('word.ribbon.bookmark-name-place')} onClick={() => openDialog('bookmark')} />
            <Button icon="link" label={t('word.cross')} title={t('word.ribbon.cross-reference-ref')} onClick={() => openDialog('crossReference')} />
          </Group>
          <Group label={t('common.comments')}>
            <Button tall icon="comment" label={t('common.comment')} onClick={() => openDialog('comment')} />
          </Group>
          <Group label={t('common.header-and')}>
            <Button tall icon="header" label={t('common.header')} onClick={() => openDialog('header')} />
            <Button tall icon="footer" label={t('common.footer')} onClick={() => openDialog('footer')} />
            <Button tall icon="file" label={t('common.page-number')} onClick={() => openDialog('pageNumber')} />
          </Group>
          <Group label={t('common.text')}>
            <Button dropdown tall icon="textbox" label={t('common.text-box')} title={t('word.ribbon.text-box-box')} onClick={(e) => menu.open(e, [
              { heading: true, label: t('common.built-in') },
              { label: t('word.simple-text-box'), icon: 'textbox', run: () => act('textBox', 'simple') },
              { label: t('word.sidebar'), icon: 'textbox', run: () => act('textBox', 'sidebar') },
              { label: t('word.pull-quote'), icon: 'textbox', run: () => act('textBox', 'quote') },
              '-',
              { label: t('word.ribbon.draw-text'), icon: 'shape', run: () => act('drawTextBox') },
            ])} />
            {references ? <Button dropdown icon="file" label={t('word.quick-parts')} title={t('word.ribbon.quick-parts-words')} onClick={(e) => menu.open(e, [...(blocks?.head() || []), ...references.quickParts(), ...(blocks?.tail() || [])])} /> : null}
            <Button dropdown icon="wordArt" label={t('common.word-art')} title={t('word.ribbon.word-art-decorative')} onClick={(e) => menu.open(e, WORDART.map(([label, spec]) => ({ label, icon: 'wand', run: () => act('wordArt', spec) })))} />
            <Button dropdown icon="textbox" label={t('word.ribbon.drop-cap')} title={t('word.ribbon.drop-cap-first')} onClick={(e) => menu.open(e, [
              { label: t('common.none'), icon: !format.dropCap ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: null }) },
              { label: t('word.ribbon.dropped'), icon: format.dropCap?.kind === 'drop' && format.dropCap?.lines === 3 ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: { kind: 'drop', lines: 3 } }) },
              { label: t('word.ribbon.dropped-two'), icon: format.dropCap?.kind === 'drop' && format.dropCap?.lines === 2 ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: { kind: 'drop', lines: 2 } }) },
              { label: t('word.ribbon.in-margin'), icon: format.dropCap?.kind === 'margin' ? 'check' : undefined, run: () => dispatch({ op: 'setDropCap', spec: { kind: 'margin', lines: 3 } }) },
            ])} />
            <Button icon="check" label={t('word.ribbon.signature')} title={t('word.ribbon.signature-line-line')} onClick={() => act('signatureLine')} />
            <Button icon="clock" label={t('common.date-and')} onClick={() => openDialog('dateTime')} />
            <Button icon="file" label={t('common.object')} title={t('word.ribbon.object-word-excel')} onClick={() => act('insertObject')} />
          </Group>
          <Group label={t('common.symbols')}>
            <Button tall icon="formula" label={t('common.equation')} title={t('word.ribbon.equation-type-one')} onClick={() => act('equation')} />
            <Button dropdown
              icon="chevronDown"
              label={t('common.built-in')}
              title={t('word.ribbon.built-equations')}
              onClick={(e) => menu.open(e, EQUATION_GALLERY.map((g) => ({ label: t(g.name), icon: 'formula', run: () => act('equation', { linear: g.linear }) })))}
            />
            <Button tall icon="star" label={t('common.symbol')} onClick={() => openDialog('symbol')} />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label={t('common.drawing-tools')}>
            <Button tall icon="mouse" label={t('common.select')} pressed={!ink?.tool} title={t('word.ribbon.select-put-pen')} onClick={() => act('inkTool', null)} />
            <Button tall icon="shape" label={t('common.lasso')} pressed={ink?.tool === 'lasso'} title={t('word.ribbon.lasso-select-draw')} onClick={() => act('inkTool', 'lasso')} />
            <Button tall icon="eraser" label={t('common.eraser')} pressed={ink?.tool === 'eraser'} title={t('common.eraser-take-away')} onClick={() => act('inkTool', 'eraser')} />
            {(ink?.pens || []).map((p) => {
              const on = ink.tool === 'pen' && ink.penId === p.id;
              const label = { pen: t('common.pen'), pencil: t('common.pencil'), highlighter: t('common.highlighter') }[p.tool];
              return (
                <Button key={p.id} tall icon="wand" label={label} pressed={on} className={`sl-pen sl-pen-${p.tool}`} data-pen={p.id} style={{ '--pen': p.color }}
                  title={on ? t('common.label-click-again', { label }) : t('word.ribbon.label-draw-page', { label })}
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
          <Group label={t('common.touch')}>
            <Button tall icon="wand" label={t('common.draw-with')} pressed={ink?.touch !== false} title={t('common.draw-touch-finger')} onClick={() => act('inkTouch')} />
          </Group>
          <Group label={t('common.stencils')}>
            <Button tall icon="minus" label={t('common.ruler')} pressed={Boolean(ink?.ruler)} title={t('word.ribbon.ruler-straight-edge')} onClick={() => act('inkRuler')} />
          </Group>
          <Group label={t('common.convert')}>
            <Button tall icon="shape" label={t('common.ink-to-shape')} pressed={Boolean(ink?.toShape)} title={t('common.ink-shape-rectangle')} onClick={() => act('inkToShape')} />
          </Group>
          <Group label={t('common.insert')}>
            <Button dropdown tall icon="shape" label={t('common.shape')} title={t('word.ribbon.shapes-engine-draw')} onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', preset, widthPx: 200, heightPx: 120 }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Design ───────────────────────────────────────────────────────── */}
      {tab === 'design' ? (
        <>
          <Group label={t('word.ribbon.document')} icon="theme">
            <Button dropdown
              tall
              icon="wand"
              label={t('common.themes')}
              title={t('word.ribbon.themes-colours-fonts', { theme: design?.name || 'Office Theme' })}
              onClick={(e) => menu.open(e, THEMES.map((th) => ({ label: t(th.name), icon: design?.builtIn === th.id ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { theme: th.id } }) })))}
            />
            <Gallery
              className="wd-styles"
              max={6}
              min={3}
              label={t('word.ribbon.styles')}
              moreTip={t('word.ribbon.more-styles-every')}
              onMore={(e) => menu.open(e, (styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).map((s) => ({ label: styleName(s), icon: (format.paragraphStyle || 'Normal') === s.id ? 'check' : undefined, run: () => para({ styleId: s.id }) })))}
            >
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).map((s) => (
                <button key={s.id} type="button" data-style={s.id} className={`wd-style${(format.paragraphStyle || 'Normal') === s.id ? ' on' : ''}`} onContextMenu={(e) => menu.open(e, styleMenu(s))} style={STYLE_LOOK[s.id] || undefined} title={t('word.ribbon.apply-style-paragraph', { style: s.name || s.id })} onClick={() => para({ styleId: s.id })}>
                  <span className="wd-style-name">{s.name || s.id}</span>
                </button>
              ))}
            </Gallery>
            <Button dropdown
              icon="contrast"
              label={t('common.colours')}
              title={t('word.ribbon.colours-themes-twelve', { colours: design?.colorName || 'Office' })}
              onClick={(e) => menu.open(e, PALETTES.map((p) => ({ label: t(p.name), icon: design?.colorName === p.name ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { colors: p.id } }) })))}
            />
            <Button dropdown
              icon="word"
              label={t('common.fonts')}
              title={t('word.ribbon.fonts-themes-heading', { fonts: design?.fonts ? `${design.fonts.major} / ${design.fonts.minor}` : 'Calibri Light / Calibri' })}
              onClick={(e) => menu.open(e, FONT_PAIRS.map((p) => ({ label: `${t(p.name)} — ${p.major} / ${p.minor}`, icon: design?.fontName === p.name ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { fonts: p.id } }) })))}
            />
            <Button dropdown icon="listNumber" label={t('word.ribbon.paragraph')} title={t('word.ribbon.spacing-paragraphs')} onClick={(e) =>
              menu.open(e, [
                { label: t('word.ribbon.no-paragraph-space'), run: () => para({ spaceBefore: 0, spaceAfter: 0, lineSpacing: 1 }) },
                { label: t('common.compact'), run: () => para({ spaceBefore: 0, spaceAfter: 4, lineSpacing: 1 }) },
                { label: t('common.tight'), run: () => para({ spaceBefore: 0, spaceAfter: 6, lineSpacing: 1.15 }) },
                { label: t('common.open'), run: () => para({ spaceBefore: 0, spaceAfter: 10, lineSpacing: 1.15 }) },
                { label: t('word.ribbon.relaxed'), run: () => para({ spaceBefore: 0, spaceAfter: 6, lineSpacing: 1.5 }) },
                { label: t('word.ribbon.double'), run: () => para({ spaceBefore: 0, spaceAfter: 8, lineSpacing: 2 }) },
              ])
            } />
            <Button dropdown
              icon="wand"
              label={t('common.effects')}
              title={t('word.ribbon.effects-fills-lines', { effects: design?.effectName || 'Office' })}
              onClick={(e) => menu.open(e, EFFECT_PRESETS.map((p) => ({ label: t(p.name), title: t(p.description), icon: design?.effects === p.id ? 'check' : undefined, run: () => dispatch({ op: 'setDocTheme', spec: { effects: p.id } }) })))}
            />
            <Button icon="check" label={t('word.ribbon.set-as')} title={t('word.ribbon.set-default-new')} onClick={() => act('setDefaultDesign')} />
          </Group>
          <Group label={t('word.ribbon.page-background')}>
            <Button dropdown tall icon="shield" label={t('word.watermark')} title={t('word.ribbon.watermark-faint-words')} onClick={(e) => menu.open(e, [
              ...WATERMARKS.map((text) => ({ label: text, run: () => dispatch({ op: 'setWatermark', text }) })),
              { label: t('word.ribbon.custom-watermark'), icon: 'textbox', run: () => openDialog('watermark') },
              { label: t('word.ribbon.picture'), icon: 'picture', run: () => act('pictureWatermark') },
              { label: t('word.ribbon.remove'), icon: 'close', run: () => dispatch({ op: 'setWatermark', text: null }) },
            ])} />
            <Button dropdown tall icon="contrast" label={t('word.ribbon.page-colour')} title={t('word.ribbon.page-colour-colour')} onClick={(e) => menu.open(e, SHADES.map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => dispatch({ op: 'setPageColour', colour: value }) })))} />
            <Button dropdown tall icon="grid" label={t('word.ribbon.page-borders')} title={t('word.ribbon.page-borders-frame')} onClick={(e) => menu.open(e, PAGE_BORDERS.map(([label, borders]) => ({ label, icon: borders ? undefined : 'close', run: () => dispatch({ op: 'setPageBorders', borders }) })))} />
          </Group>
        </>
      ) : null}

      {/* ── Layout ───────────────────────────────────────────────────────── */}
      {tab === 'layout' ? (
        <>
          <Group label={t('common.page-setup')}>
            <Button dropdown tall icon="crop" label={t('common.margins')} onClick={(e) => menu.open(e, MARGINS.map(([v, l]) => ({ label: l, run: () => dispatch({ op: 'setPageSetup', spec: { margins: v } }) })))} />
            <Button dropdown tall icon="rotate" label={t('common.orientation')} title={landscape ? t('common.orientation-now-landscape') : t('common.orientation-now')} onClick={(e) =>
              menu.open(e, [
                { label: t('common.portrait'), icon: landscape ? undefined : 'check', run: () => dispatch({ op: 'setPageSetup', spec: { orientation: 'portrait' } }) },
                { label: t('common.landscape'), icon: landscape ? 'check' : undefined, run: () => dispatch({ op: 'setPageSetup', spec: { orientation: 'landscape' } }) },
              ])
            } />
            <Button dropdown tall icon="file" label={t('common.size')} onClick={(e) => menu.open(e, PAGE_SIZES.map(([v, l]) => ({ label: l, run: () => dispatch({ op: 'setPageSetup', spec: { size: v } }) })))} />
            <Button dropdown tall icon="grid" label={t('common.columns')} title={colCount > 1 ? t('word.ribbon.columns-now-count', { count: colCount }) : t('word.ribbon.columns-now-one')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.one'), icon: isColumnPreset('one') ? 'check' : undefined, run: () => setColumns(null) },
              { label: t('word.ribbon.two'), icon: isColumnPreset('two') ? 'check' : undefined, run: () => setColumns({ count: 2, spaceTwips: 720, separator: Boolean(columns.separator) }) },
              { label: t('word.ribbon.three'), icon: isColumnPreset('three') ? 'check' : undefined, run: () => setColumns({ count: 3, spaceTwips: 720, separator: Boolean(columns.separator) }) },
              { label: t('common.left'), icon: isColumnPreset('left') ? 'check' : undefined, run: () => setColumns(unevenSpec(true)) },
              { label: t('common.right'), icon: isColumnPreset('right') ? 'check' : undefined, run: () => setColumns(unevenSpec(false)) },
              '-',
              // The section right to left: its columns run from the right, as an Arabic or Hebrew page's do.
              { label: t('word.ribbon.right-to-left'), icon: section?.rtl ? 'check' : undefined, run: () => dispatch({ op: 'setPageSetup', spec: { rtl: !section?.rtl } }) },
              {
                label: t('word.ribbon.line-between'), icon: columns.separator ? 'check' : undefined, disabled: colCount <= 1,
                run: () => setColumns({
                  count: colCount, spaceTwips: Math.round((columns.spacePx || 36) * 15), separator: !columns.separator,
                  ...(columns.widths ? { widths: columns.widths.map((w) => Math.round(w * 15)) } : {}),
                }),
              },
            ])} />
            <Button dropdown icon="file" label={t('common.breaks')} onClick={(e) =>
              menu.open(e, [
                { label: t('word.page-break'), icon: 'file', run: () => dispatch({ op: 'insertPageBreak' }) },
                { label: t('word.ribbon.section-break-not'), disabled: true },
              ])
            } />
            <Button dropdown icon="listNumber" label={t('word.ribbon.line-numbers')} title={t('word.ribbon.line-numbers-number')} onClick={(e) => menu.open(e, [
              { label: t('common.none'), icon: 'close', run: () => dispatch({ op: 'setLineNumbers', spec: null }) },
              { label: t('common.continuous'), run: () => dispatch({ op: 'setLineNumbers', spec: { countBy: 1, restart: 'continuous' } }) },
              { label: t('word.ribbon.restart-each'), run: () => dispatch({ op: 'setLineNumbers', spec: { countBy: 1, restart: 'newPage' } }) },
              { label: t('word.ribbon.every-fifth'), run: () => dispatch({ op: 'setLineNumbers', spec: { countBy: 5, restart: 'continuous' } }) },
            ])} />
            <Button dropdown icon="minus" label={t('word.hyphenation')} title={model?.hyphenation?.auto ? t('word.ribbon.hyphenation-words-broken-ends') : t('word.ribbon.hyphenation-words-broken')} onClick={(e) => menu.open(e, [
              { label: t('common.none'), icon: !model?.hyphenation?.auto ? 'check' : undefined, run: () => act('hyphenation', 'none') },
              { label: t('common.automatic'), icon: model?.hyphenation?.auto ? 'check' : undefined, run: () => act('hyphenation', 'auto') },
              { label: t('common.manual'), run: () => act('hyphenation', 'manual') },
              '-',
              { label: t('word.ribbon.hyphenation'), run: () => act('hyphenation', 'options') },
              { label: t('word.ribbon.dont-hyphenate'), icon: format.noHyphens ? 'check' : undefined, run: () => act('hyphenation', 'paragraph') },
            ])} />
          </Group>
          <Group label={t('common.paragraph')}>
            <div className="wd-fields">
              <label>{t('word.ribbon.indent')}</label>
              <Button icon="chevronLeft" title={t('word.ribbon.decrease-left-indent')} disabled={!format.indentLevel} onClick={() => para({ indentDelta: -1 })} />
              <span className="wd-field-value">{t('word.value-cm', { value: (format.indentLevel || 0) * 1.27 })}</span>
              <Button icon="chevronRight" title={t('word.ribbon.increase-left-indent')} onClick={() => para({ indentDelta: 1 })} />
            </div>
            <div className="wd-fields">
              <label>{t('word.ribbon.spacing')}</label>
              <Input type="number" min="0" max="200" value={format.spaceBefore ?? 0} title={t('word.ribbon.before-pt')} onChange={(e) => para({ spaceBefore: Number(e.target.value) })} style={{ width: 56 }} />
              <Input type="number" min="0" max="200" value={format.spaceAfter ?? 0} title={t('word.ribbon.after-pt')} onChange={(e) => para({ spaceAfter: Number(e.target.value) })} style={{ width: 56 }} />
            </div>
          </Group>
          {arrange}
        </>
      ) : null}

      {/* ── Shape Format (contextual) ─────────────────────────────────────── */}
      {tab === 'shapeFormat' ? (
        <>
          <Group label={t('word.ribbon.insert-shapes')}>
            <Button tall icon="textbox" label={t('word.ribbon.draw-text')} pressed={Boolean(view.drawBox)} title={t('word.ribbon.draw-text-box')} onClick={() => act('drawTextBox')} />
          </Group>
          <Group label={t('word.ribbon.shape-styles')}>
            <Button dropdown tall icon="wand" label={t('common.shape-fill')} disabled={!shapeLike} title={shapeLike ? (look.fill ? t('word.ribbon.shape-fill-now', { fill: look.fill }) : t('word.ribbon.shape-fill')) : t('word.ribbon.shape-fill-select')} onClick={(e) => menu.open(e, [
              { label: t('common.no-fill'), icon: 'close', run: () => act('boxFormat', { fill: null }) },
              '-',
              ...SHAPE_COLOURS.map(([hex, label]) => ({ label, icon: look.fill && look.fill.replace('#', '').toUpperCase() === hex ? 'check' : undefined, run: () => act('boxFormat', { fill: hex }) })),
            ])} />
            <Button dropdown tall icon="shape" label={t('common.shape-outline')} disabled={!shapeLike} title={shapeLike ? (look.line ? t('word.ribbon.shape-outline-now', { line: look.line }) : t('word.ribbon.shape-outline')) : t('word.ribbon.shape-outline-select')} onClick={(e) => menu.open(e, [
              { label: t('common.no-outline'), icon: 'close', run: () => act('boxFormat', { line: null }) },
              '-',
              ...SHAPE_COLOURS.map(([hex, label]) => ({ label, icon: look.line && look.line.replace('#', '').toUpperCase() === hex ? 'check' : undefined, run: () => act('boxFormat', { line: { colour: hex, widthPx: look.lineWidthPx || 1 } }) })),
              { heading: true, label: t('common.weight') },
              ...WEIGHTS.map(([label, px]) => ({ label, icon: look.line && Math.abs((look.lineWidthPx || 1) - px) < 0.2 ? 'check' : undefined, run: () => act('boxFormat', { line: { colour: (look.line || '#000000').replace('#', ''), widthPx: px } }) })),
            ])} />
          </Group>
          <Group label={t('common.text')}>
            <Button dropdown tall icon="rotate" label={t('word.ribbon.text')} disabled={drawing?.kind !== 'textbox'} title={t('word.ribbon.text-direction-across')} onClick={(e) => menu.open(e, [
              { label: t('common.horizontal'), icon: (look.vert || 'horz') === 'horz' ? 'check' : undefined, run: () => act('boxFormat', { vert: 'horz' }) },
              { label: t('word.ribbon.rotate-all'), icon: look.vert === 'vert' ? 'check' : undefined, run: () => act('boxFormat', { vert: 'vert' }) },
              { label: t('word.ribbon.rotate-all-text'), icon: look.vert === 'vert270' ? 'check' : undefined, run: () => act('boxFormat', { vert: 'vert270' }) },
            ])} />
            <Button dropdown tall icon="alignCenter" label={t('word.ribbon.align-text')} disabled={drawing?.kind !== 'textbox'} title={t('word.ribbon.align-text-words')} onClick={(e) => menu.open(e, [
              { label: t('common.top'), icon: (look.vAnchor || 'top') === 'top' ? 'check' : undefined, run: () => act('boxFormat', { vAnchor: 'top' }) },
              { label: t('common.middle'), icon: look.vAnchor === 'middle' ? 'check' : undefined, run: () => act('boxFormat', { vAnchor: 'middle' }) },
              { label: t('common.bottom'), icon: look.vAnchor === 'bottom' ? 'check' : undefined, run: () => act('boxFormat', { vAnchor: 'bottom' }) },
            ])} />
            <Button dropdown tall icon="crop" label={t('common.margins')} disabled={drawing?.kind !== 'textbox'} title={t('word.ribbon.margins-room-between')} onClick={(e) => menu.open(e, BOX_MARGINS.map(([label, insets]) => ({
              label, icon: look.insets && ['l', 't', 'r', 'b'].every((k) => Math.abs((look.insets[k] || 0) - insets[k]) < 0.6) ? 'check' : undefined, run: () => act('boxFormat', { insets }),
            })))} />
          </Group>
          {/* WordArt Styles → Text Effects → Transform, as Word's: a text box's words along a path. */}
          <Group label={t('common.text-effects')}>
            {WARP_PRESETS.map((p) => (
              <Button
                key={p.id} tall icon="wand" label={t(p.label)} className="wd-warp" data-preset={p.id}
                disabled={drawing?.kind !== 'textbox'}
                pressed={drawing?.kind === 'textbox' && (look.warp || 'textNoShape') === p.id}
                title={drawing?.kind !== 'textbox' ? t('word.ribbon.label-select-text', { label: t(p.label) }) : p.id === 'textNoShape' ? t('common.no-transform-words') : p.label === 'Button' ? t('common.transform-words-along-button') : p.label === 'Circle' ? t('word.ribbon.transform-words-along-circle') : t('word.ribbon.transform-words-along')}
                onClick={() => act('boxFormat', { warp: p.id === 'textNoShape' ? null : p.id })}
              />
            ))}
            {/* The warps, the rest of Office's gallery, from a menu. */}
            <Button dropdown tall disabled={drawing?.kind !== 'textbox'} icon="wand" label={t('common.more')} className="wd-warp-more" pressed={WARP_MORE.some((p) => p.id === look.warp)} title={WARP_MORE.some((p) => p.id === look.warp) ? t('common.transform-now-warp', { warp: t(warpLabel(look.warp)) }) : t('common.more-transforms-warps')} onClick={(e) => menu.open(e, WARP_MORE.map((p) => ({ label: t(p.label), icon: look.warp === p.id ? 'check' : 'wand', run: () => act('boxFormat', { warp: p.id }) })))} />
          </Group>
          {arrange}
          <Group label={t('common.size')}>
            <Rows>
              {sizeBox(t('common.height'), 'h')}
              {sizeBox(t('common.width'), 'w')}
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Table Design (contextual) ─────────────────────────────────────── */}
      {tab === 'tableDesign' && table ? (
        <>
          <Group label={t('common.table-style')}>
            <Rows>
              <Button icon={styleLook.firstRow ? 'check' : undefined} label={t('common.header-row')} pressed={styleLook.firstRow} title={t('common.header-row-first')} onClick={() => setLook({ firstRow: !styleLook.firstRow })} />
              <Button icon={styleLook.lastRow ? 'check' : undefined} label={t('common.total-row')} pressed={styleLook.lastRow} title={t('common.total-row-last')} onClick={() => setLook({ lastRow: !styleLook.lastRow })} />
              <Button icon={!styleLook.noHBand ? 'check' : undefined} label={t('common.banded-rows')} pressed={!styleLook.noHBand} title={t('common.banded-rows-every')} onClick={() => setLook({ noHBand: !styleLook.noHBand })} />
            </Rows>
            <Rows>
              <Button icon={styleLook.firstColumn ? 'check' : undefined} label={t('common.first-column')} pressed={styleLook.firstColumn} title={t('common.first-column-first')} onClick={() => setLook({ firstColumn: !styleLook.firstColumn })} />
              <Button icon={styleLook.lastColumn ? 'check' : undefined} label={t('common.last-column')} pressed={styleLook.lastColumn} title={t('common.last-column-last')} onClick={() => setLook({ lastColumn: !styleLook.lastColumn })} />
              <Button icon={!styleLook.noVBand ? 'check' : undefined} label={t('common.banded')} pressed={!styleLook.noVBand} title={t('common.banded-columns-every')} onClick={() => setLook({ noVBand: !styleLook.noVBand })} />
            </Rows>
          </Group>
          <Group label={t('word.ribbon.table-styles')}>
            <Button dropdown tall icon="table" label={t('word.ribbon.table-styles')} title={tableStyleId ? t('word.ribbon.table-styles-words-own', { style: t(TABLE_STYLES.find((st) => st.id === tableStyleId)?.name || tableStyleId) }) : t('word.ribbon.table-styles-words')} onClick={(e) => menu.open(e, [
              { heading: true, label: t('word.ribbon.plain-tables') },
              ...TABLE_STYLES.filter((st) => st.family === 'plain').map((st) => ({ label: t(st.name), icon: st.id === tableStyleId ? 'check' : undefined, run: () => tableOp('style', { id: st.id }) })),
              { heading: true, label: t('word.ribbon.grid') },
              ...TABLE_STYLES.filter((st) => st.family === 'grid4').map((st) => ({ label: t(st.name), icon: st.id === tableStyleId ? 'check' : undefined, run: () => tableOp('style', { id: st.id }) })),
              { heading: true, label: t('word.ribbon.list') },
              ...TABLE_STYLES.filter((st) => st.family === 'list4').map((st) => ({ label: t(st.name), icon: st.id === tableStyleId ? 'check' : undefined, run: () => tableOp('style', { id: st.id }) })),
              '-',
              { label: t('common.clear'), icon: 'close', disabled: !tableStyleId, run: () => tableOp('style', { id: null }) },
            ])} />
          </Group>
          <Group label={t('common.borders')}>
            <Rows>
              <Button dropdown icon="minus" label={t('word.ribbon.line-style', { style: PEN_STYLES.find(([v]) => v === pen.val)?.[1] || t('word.ribbon.single') })} title={t('word.ribbon.line-style-pens')} onClick={(e) => menu.open(e, PEN_STYLES.map(([val, label]) => ({ label, icon: pen.val === val ? 'check' : undefined, run: () => setPen({ ...pen, val }) })))} />
              <Button dropdown icon="sliders" label={t('word.ribbon.weight-weight', { weight: PEN_WEIGHTS.find(([sz]) => sz === pen.sz)?.[1] || t('word.ribbon.1-2-pt') })} title={t('word.ribbon.pen-weight-how')} onClick={(e) => menu.open(e, PEN_WEIGHTS.map(([sz, label]) => ({ label, icon: pen.sz === sz ? 'check' : undefined, run: () => setPen({ ...pen, sz }) })))} />
              <Button dropdown icon="wand" label={t('word.ribbon.pen-colour')} title={t('word.ribbon.pen-colour-now', { colour: SHAPE_COLOURS.find(([hex]) => hex === pen.color)?.[1] || t('common.automatic') })} onClick={(e) => menu.open(e, [
                { label: t('common.automatic'), icon: pen.color === 'auto' ? 'check' : undefined, run: () => setPen({ ...pen, color: 'auto' }) },
                '-',
                ...SHAPE_COLOURS.map(([hex, label]) => ({ label, icon: pen.color === hex ? 'check' : undefined, run: () => setPen({ ...pen, color: hex }) })),
              ])} />
            </Rows>
            <Button dropdown tall icon="borders" label={t('common.borders')} title={t('word.ribbon.borders-lines-round-between')} onClick={(e) => menu.open(e, [
              ['bottom', t('word.ribbon.bottom-border')], ['top', t('word.ribbon.top-border')], ['left', t('word.ribbon.left-border')], ['right', t('word.ribbon.right-border')], '-',
              ['none', t('word.ribbon.no-border')], ['all', t('word.ribbon.all-borders')], ['outside', t('word.ribbon.outside-borders')], ['inside', t('word.ribbon.inside-borders')], '-',
              ['insideH', t('word.ribbon.inside-horizontal')], ['insideV', t('word.ribbon.inside-vertical')],
            ].map((item) => (item === '-' ? '-' : { label: item[1], run: () => tableOp('borders', { kind: item[0], pen }) })))} />
          </Group>
          <Group label={t('common.shading')}>
            <Button dropdown tall icon="bucketFill" label={t('common.shading')} title={t('word.ribbon.shading-selected-cells')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.no-colour'), icon: 'close', run: () => tableOp('shading', { fill: null }) },
              '-',
              ...SHAPE_COLOURS.map(([hex, label]) => ({ label, run: () => tableOp('shading', { fill: hex }) })),
            ])} />
          </Group>
        </>
      ) : null}

      {/* ── Table Layout (contextual) ─────────────────────────────────────── */}
      {tab === 'tableLayout' && table ? (
        <>
          <Group label={t('common.draw')}>
            <Button tall icon="pen" label={t('word.ribbon.draw-table')} pressed={view.tableDraw === 'pen'} title={t('word.ribbon.draw-table-draw')} onClick={() => act('tableDraw', 'pen')} />
            <Button tall icon="eraser" label={t('common.eraser')} pressed={view.tableDraw === 'eraser'} title={t('word.ribbon.eraser-click-line')} onClick={() => act('tableDraw', 'eraser')} />
          </Group>
          <Group label={t('common.table')}>
            <Button tall icon="settings" label={t('common.properties')} title={t('word.ribbon.table-properties-tables')} onClick={() => openDialog('tableProperties')} />
            <Button dropdown tall icon="alignCenter" label={t('word.ribbon.align-table')} title={t('word.ribbon.align-table-table')} onClick={(e) => menu.open(e, [
              { label: t('common.left'), run: () => tableOp('align', { align: 'left' }) },
              { label: t('common.centre'), run: () => tableOp('align', { align: 'center' }) },
              { label: t('common.right'), run: () => tableOp('align', { align: 'right' }) },
            ])} />
            <Button tall icon="grid" label={t('word.ribbon.view')} pressed={!view.noTableGridlines} title={t('word.ribbon.view-gridlines-tables')} onClick={() => act('toggleTableGridlines')} />
          </Group>
          <Group label={t('word.ribbon.rows-and')}>
            <Button dropdown tall icon="minus" label={t('common.delete')} title={t('word.ribbon.delete-row-column')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.delete-columns'), icon: 'minus', run: () => tableOp('deleteColumn') },
              { label: t('word.ribbon.delete'), icon: 'minus', run: () => tableOp('deleteRow') },
              { label: t('word.ribbon.delete-table'), icon: 'trash', run: () => tableOp('deleteTable') },
            ])} />
            <Button tall icon="rowAbove" label={t('word.ribbon.insert-above')} title={t('word.ribbon.insert-row-above')} onClick={() => tableOp('insertRowAbove')} />
            <Button tall icon="rowBelow" label={t('word.ribbon.insert-below')} title={t('word.ribbon.insert-row-below')} onClick={() => tableOp('insertRowBelow')} />
            <Rows>
              <Button icon="colLeft" label={t('word.ribbon.insert')} title={t('word.ribbon.insert-column-left')} onClick={() => tableOp('insertColumnLeft')} />
              <Button icon="colRight" label={t('word.ribbon.insert-right')} title={t('word.ribbon.insert-column-right')} onClick={() => tableOp('insertColumnRight')} />
            </Rows>
          </Group>
          <Group label={t('word.merge')}>
            <Button tall icon="mergeCells" label={t('word.ribbon.merge-cells')} title={acrossCells ? t('word.ribbon.merge-selected') : t('word.ribbon.merge-cells-select')} disabled={!acrossCells} onClick={() => tableOp('mergeCells')} />
            <Button tall icon="splitCells" label={t('word.ribbon.split')} title={t('word.ribbon.split-cells-merged')} onClick={(e) => mergedCaret ? tableOp('splitCell') : menu.open(e, [
              { label: t('word.ribbon.into-2-columns'), run: () => tableOp('splitInto', { columns: 2 }) },
              { label: t('word.ribbon.into-3'), run: () => tableOp('splitInto', { columns: 3 }) },
              { label: t('word.ribbon.into-2'), run: () => tableOp('splitInto', { rows: 2 }) },
              { label: t('word.ribbon.2-columns-2'), run: () => tableOp('splitInto', { columns: 2, rows: 2 }) },
              { label: t('word.ribbon.merged-cell-back'), run: () => tableOp('splitCell') },
            ])} />
            <Button tall icon="splitCells" label={t('word.ribbon.split-table')} title={t('word.ribbon.split-table-carets')} onClick={() => tableOp('splitTable')} />
          </Group>
          <Group label={t('word.ribbon.cell-size')}>
            <Rows>
              <div className="wd-fields" title={t('word.ribbon.carets-column-width')}>
                <label>{t('common.width')}</label>
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
                <span className="wd-field-value" style={{ minWidth: 0 }}>{t('word.ribbon.cm')}</span>
              </div>
            </Rows>
            <Button dropdown tall icon="table" label={t('word.ribbon.auto-fit')} title={t('word.ribbon.auto-fit-table')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.auto-fit-contents'), run: () => tableOp('autoFit', { mode: 'contents' }) },
              { label: t('word.ribbon.auto-fit-window'), run: () => tableOp('autoFit', { mode: 'window' }) },
              { label: t('word.ribbon.fixed'), run: () => tableOp('autoFit', { mode: 'fixed' }) },
            ])} />
            <Button tall icon="sliders" label={t('word.ribbon.distribute')} title={t('word.ribbon.distribute-columns-every')} onClick={() => tableOp('distributeColumns')} />
          </Group>
          <Group label={t('common.alignment')}>
            <Button dropdown tall icon="rotate" label={t('word.ribbon.text')} title={t('word.ribbon.text-direction-selected')} onClick={(e) => menu.open(e, [
              { label: t('common.horizontal'), run: () => tableOp('textDirection', { dir: null }) },
              { label: t('word.ribbon.rotate-all'), run: () => tableOp('textDirection', { dir: 'down' }) },
              { label: t('word.ribbon.rotate-all-text'), run: () => tableOp('textDirection', { dir: 'up' }) },
            ])} />
            <Button dropdown tall icon="crop" label={t('word.ribbon.cell')} title={t('word.ribbon.cell-margins-room')} onClick={(e) => menu.open(e, [
              { label: t('common.narrow'), run: () => tableOp('cellMargins', { margins: { top: 0, left: 57, bottom: 0, right: 57 } }) },
              { label: t('common.normal'), run: () => tableOp('cellMargins', { margins: { top: 0, left: 108, bottom: 0, right: 108 } }) },
              { label: t('common.wide'), run: () => tableOp('cellMargins', { margins: { top: 72, left: 216, bottom: 72, right: 216 } }) },
            ])} />
            <Button dropdown tall icon="alignCenter" label={t('word.ribbon.cell-alignment')} title={t('word.ribbon.cell-alignment-words', { align: { top: t('common.top-2'), center: t('word.ribbon.centre'), bottom: t('common.bottom-2') }[cellAlign] || cellAlign })} onClick={(e) => menu.open(e, [
              { label: t('common.top'), icon: cellAlign === 'top' ? 'check' : undefined, run: () => tableOp('cellVAlign', { v: 'top' }) },
              { label: t('common.centre'), icon: cellAlign === 'center' ? 'check' : undefined, run: () => tableOp('cellVAlign', { v: 'center' }) },
              { label: t('common.bottom'), icon: cellAlign === 'bottom' ? 'check' : undefined, run: () => tableOp('cellVAlign', { v: 'bottom' }) },
            ])} />
          </Group>
          <Group label={t('common.direction')}>
            <Button tall icon="textRtl" label={t('word.ribbon.right-to')} title={t('word.ribbon.right-left-tables')} pressed={table.rtl} onClick={() => tableOp('direction', { rtl: !table.rtl })} />
          </Group>
          <Group label={t('common.data')}>
            <Button dropdown tall icon="sort" label={t('common.sort')} title={t('word.ribbon.sort-tables-rows')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.sort-z-column'), run: () => tableOp('sort', { descending: false }) },
              { label: t('word.ribbon.sort-z-to-a-by-this-column'), run: () => tableOp('sort', { descending: true }) },
              '-',
              { label: t('word.ribbon.z-first-row'), run: () => tableOp('sort', { descending: false, header: true }) },
              { label: t('word.ribbon.z-first-row-header'), run: () => tableOp('sort', { descending: true, header: true }) },
            ])} />
            <Button dropdown tall icon="file" label={t('word.ribbon.convert-to')} title={t('word.ribbon.convert-text-each')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.separated-by-tabs'), run: () => tableOp('toText', { separator: 'tab' }) },
              { label: t('word.ribbon.separated'), run: () => tableOp('toText', { separator: 'comma' }) },
              { label: t('word.ribbon.each-cell-own'), run: () => tableOp('toText', { separator: 'paragraph' }) },
            ])} />
            <Button dropdown tall icon="formula" label={t('common.formula')} title={t('word.ribbon.formula-field-carets')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.sum-above'), run: () => tableOp('formula', { formula: '=SUM(ABOVE)' }) },
              { label: t('word.ribbon.sum-left'), run: () => tableOp('formula', { formula: '=SUM(LEFT)' }) },
              { label: t('word.ribbon.average'), run: () => tableOp('formula', { formula: '=AVERAGE(ABOVE)', format: '#,##0.00' }) },
              { label: t('word.ribbon.count'), run: () => tableOp('formula', { formula: '=COUNT(ABOVE)' }) },
              { label: t('word.ribbon.largest'), run: () => tableOp('formula', { formula: '=MAX(ABOVE)' }) },
              { label: t('word.ribbon.smallest'), run: () => tableOp('formula', { formula: '=MIN(ABOVE)' }) },
              { label: t('word.ribbon.your-own'), run: () => openDialog('tableFormula') },
              { label: t('word.ribbon.update-all'), run: () => dispatch({ op: 'updateTableFormulas' }) },
            ])} />
            <Button tall icon="refresh" label={t('word.ribbon.repeat')} pressed={Boolean(caretCell?.rowHeader)} title={t('word.ribbon.repeat-header-rows')} onClick={() => tableOp('headerRows', { on: !caretCell?.rowHeader })} />
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
            <Group label={t('common.3d-model-views')}>
              {rows.map((row, i) => (
                <Rows key={i}>
                  {row.map(([key, label, view]) => <Button key={key} icon={at(view) ? 'check' : 'shape'} label={t(label)} className="wd-model3d-view" data-view={key} pressed={at(view)} title={t('common.label-model-turned', { label: t(label) })} onClick={() => act('model3dView', view)} />)}
                </Rows>
              ))}
            </Group>
            <Group label={t('common.turn')}>
              <Rows>
                <Button icon="chevronLeft" label={t('common.turn-left')} className="wd-model3d-turn-left" title={t('common.turn-left-fifteen')} onClick={() => act('model3dView', { turn: { yaw: -15, pitch: 0 } })} />
                <Button icon="chevronRight" label={t('common.turn-right')} title={t('common.turn-right-fifteen')} onClick={() => act('model3dView', { turn: { yaw: 15, pitch: 0 } })} />
              </Rows>
              <Rows>
                <Button icon="chevronUp" label={t('common.tip-back')} title={t('common.tip-back-fifteen')} onClick={() => act('model3dView', { turn: { yaw: 0, pitch: 15 } })} />
                <Button icon="chevronDown" label={t('common.tip-forward')} title={t('common.tip-forward-fifteen')} onClick={() => act('model3dView', { turn: { yaw: 0, pitch: -15 } })} />
              </Rows>
            </Group>
            <Group label={t('common.adjust')}>
              <Button tall icon="undo" label={t('common.reset-3d')} className="wd-model3d-reset" title={t('common.reset-3d-model')} onClick={() => act('model3dView', 'reset')} />
            </Group>
          </>
        );
      })() : null}

      {/* ── Picture Format (contextual) ───────────────────────────────────── */}
      {tab === 'pictureFormat' ? (
        <>
          <Group label={t('common.edit')}>
            <Button
              tall
              icon="image"
              label={t('common.edit-in')}
              title={t('common.edit-image-crop')}
              disabled={picked?.image == null || Boolean(pickedModel) || Boolean(model?.blocks?.[picked?.block]?.images?.[picked?.image]?.object)}
              onClick={() => act('editInImage')}
              data-role="edit-in-image"
            />
          </Group>
          {arrange}
          <Group label={t('common.size')}>
            <Rows>
              {sizeBox(t('common.height'), 'h')}
              {sizeBox(t('common.width'), 'w')}
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── References ───────────────────────────────────────────────────── */}
      {tab === 'references' ? (
        <>
          <Group label={t('word.ribbon.table-of')}>
            <Button tall icon="listBullet" label={t('word.ribbon.table-of')} title={t('word.ribbon.table-contents-live')} onClick={() => act('tableOfContents')} />
            <Button dropdown icon="plus" label={t('common.add-text')} title={t('word.ribbon.make-paragraph-heading')} onClick={(e) =>
              menu.open(e, [
                { label: t('word.ribbon.do-not-show'), run: () => para({ styleId: null }) },
                { label: t('word.ribbon.level-1-heading'), run: () => para({ styleId: 'Heading1' }) },
                { label: t('word.ribbon.level-2'), run: () => para({ styleId: 'Heading2' }) },
                { label: t('word.ribbon.level-3'), run: () => para({ styleId: 'Heading3' }) },
              ])
            } />
            <Button icon="refresh" label={t('word.update-table')} disabled={!model?.tableOfContents} title={t('word.ribbon.update-table-rebuild')} onClick={() => act('updateTableOfContents')} />
            <Button icon="close" label={t('word.ribbon.remove-table')} disabled={!model?.tableOfContents} title={t('word.ribbon.remove-table')} onClick={() => act('removeTableOfContents')} />
          </Group>
          <Group label={t('word.ribbon.footnotes')}>
            <Button tall icon="file" label={t('word.ribbon.insert-footnote')} title={t('word.ribbon.raised-number-caret')} onClick={() => act('insertNote', 'footnote')} />
            <Button icon="file" label={t('word.ribbon.insert-endnote')} title={t('word.ribbon.raised-number-caret-words')} onClick={() => act('insertNote', 'endnote')} />
            <Button icon="chevronDown" label={t('word.ribbon.next')} title={t('word.ribbon.go-next-footnote')} onClick={() => act('nextNote', 1)} />
            <Button icon="eye" label={t('word.ribbon.show-notes')} title={t('word.ribbon.scroll-footnotes-under')} onClick={() => act('showNotes')} />
          </Group>
          {references ? <CitationsGroup refs={references} menu={menu} /> : null}
          <Group label={t('common.captions')}>
            <Button tall icon="textbox" label={t('word.ribbon.insert-caption')} title={t('word.ribbon.insert-caption-label')} onClick={() => openDialog('caption')} />
            {references ? captionsExtra(references, <Button icon="link" label={t('word.cross')} title={t('word.ribbon.cross-reference-ref')} onClick={() => openDialog('crossReference')} />) : <Button icon="link" label={t('word.cross')} title={t('word.ribbon.cross-reference-ref')} onClick={() => openDialog('crossReference')} />}
          </Group>
          <Group label={t('common.fields')}>
            <Button tall icon="refresh" label={t('word.update')} title={t('word.ribbon.update-fields-f9')} onClick={() => commands['field.update']?.run?.()} />
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
          <Group label={t('common.proofing')}>
            <Button tall icon="check" label={t('common.spelling')} pressed={review?.pane === 'editor'} title={t('word.ribbon.spelling-f7-check')} onClick={() => review?.startSpelling()} />
            <Button tall icon="word" label={t('common.thesaurus')} pressed={review?.pane === 'thesaurus'} title={t('word.ribbon.thesaurus-shift-f7')} onClick={() => review?.openThesaurus()} />
            <Button tall icon="listNumber" label={t('word.ribbon.word-count')} onClick={() => openDialog('wordCount')} />
          </Group>
          <Group label={t('word.ribbon.speech')}>
            <Button tall icon="volume" label={t('word.ribbon.read-aloud')} pressed={Boolean(view.reading)} onClick={() => act('readAloud')} />
          </Group>
          <Group label={t('common.accessibility')}>
            <Button tall icon="shield" label={t('common.check-accessibility')} pressed={review?.pane === 'accessibility'} title={t('word.ribbon.check-accessibility-alt')} onClick={() => review?.openAccessibility()} />
            <Button icon="textbox" label={t('common.alt-text')} disabled={!picked} title={picked ? t('word.ribbon.alt-text-describe') : t('word.ribbon.alt-text-click')} onClick={() => picked && review?.openAltText({ block: picked.block, image: picked.image })} />
          </Group>
          <Group label={t('common.language')}>
            <Button tall icon="globe" label={t('common.language')} title={t('word.ribbon.language-mark-selected')} onClick={() => act('language')} />
          </Group>
          <Group label={t('common.comments')}>
            <Button tall icon="comment" label={t('common.new-comment')} onClick={() => openDialog('comment')} />
            <Button dropdown
              tall
              icon="trash"
              label={t('common.delete')}
              disabled={!comments.length}
              title={comments.length ? t('word.ribbon.delete-comment-carets') : t('word.ribbon.delete-document-no')}
              onClick={(e) => menu.open(e, [
                { label: t('common.delete'), icon: 'trash', disabled: !comments.some((c) => c.blockIndex === (model?.selection?.focus?.block ?? -1)), run: () => act('deleteComment') },
                { label: t('word.ribbon.delete-all-comments'), icon: 'trash', run: () => act('deleteComment', 'all') },
              ])}
            />
            <Button icon="chevronUp" label={t('common.previous')} disabled={!comments.length} onClick={() => act('comment', -1)} />
            <Button icon="chevronDown" label={t('common.next')} disabled={!comments.length} onClick={() => act('comment', 1)} />
            <Button icon="listBullet" label={comments.length ? t('word.ribbon.show-count', { count: comments.length }) : t('common.show-comments')} disabled={!comments.length} onClick={() => openDialog('comments')} />
          </Group>
          <Group label={t('word.ribbon.tracking')}>
            <Button tall icon="eye" label={t('word.ribbon.track')} pressed={Boolean(model?.trackRevisions)} disabled={Boolean(model?.protection?.lockedTracking)} title={model?.protection?.lockedTracking ? t('word.ribbon.track-changes-locked') : t('word.ribbon.track-changes-record')} onClick={() => act('toggleTrackChanges')} />
            <Button dropdown icon="eye" label={MARKUP_LABELS[view.markupMode] || t('word.ribbon.simple')} title={t('word.ribbon.display-review-how')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.all-markup'), run: () => act('markupMode', 'all') },
              { label: t('word.ribbon.simple'), run: () => act('markupMode', 'simple') },
              { label: t('word.ribbon.no-markup'), run: () => act('markupMode', 'final') },
              { label: t('word.ribbon.original'), run: () => act('markupMode', 'original') },
            ])} />
            <Button icon="list" label={t('common.reviewing')} title={t('word.ribbon.every-tracked-change')} disabled={!model?.blocks?.some((b) => b.tracked)} onClick={() => openDialog('tracked')} />
          </Group>
          <Group label={t('common.changes')}>
            {/* This Change is the change selected (Next and Previous select one); All works wherever the caret is. */}
            <Button dropdown tall icon="check" label={t('common.accept')} disabled={!anyTracked || Boolean(model?.protection?.lockedTracking)} title={t('word.ribbon.accept-keep-change')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.accept-this-change'), disabled: !trackedHere, run: () => act('acceptChanges') },
              { label: t('word.ribbon.accept-all-changes'), run: () => act('acceptChanges', 'all') },
            ])} />
            <Button dropdown tall icon="close" label={t('common.reject')} disabled={!anyTracked || Boolean(model?.protection?.lockedTracking)} title={t('word.ribbon.reject-undo-change')} onClick={(e) => menu.open(e, [
              { label: t('word.ribbon.reject-this-change'), disabled: !trackedHere, run: () => act('rejectChanges') },
              { label: t('word.ribbon.reject-all-changes'), run: () => act('rejectChanges', 'all') },
            ])} />
            <Button icon="chevronUp" label={t('common.previous')} disabled={!anyTracked} title={t('word.ribbon.previous-tracked')} onClick={() => act('nextChange', -1)} />
            <Button icon="chevronDown" label={t('common.next')} disabled={!anyTracked} title={t('word.ribbon.next-tracked-change')} onClick={() => act('nextChange', 1)} />
          </Group>
          <Group label={t('common.compare')}>
            <Button tall icon="copy" label={t('common.compare')} title={t('word.ribbon.compare-two-versions')} onClick={() => act('compare')} />
          </Group>
          <Group label={t('common.protect')}>
            <Button tall icon="lock" label={t('word.restrict')} pressed={Boolean(view.restrict)} title={model?.protection?.enforced ? t('word.ribbon.restrict-editing') : t('word.ribbon.restrict-editing-limit')} onClick={() => act('restrictPane')} />
            <Button tall icon="shield" label={t('common.redact')} pressed={Boolean(redaction?.open)} title={t('common.redact-mark-words')} onClick={() => redaction?.toggle()} />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
      {tab === 'immersive' && view.immersive ? (
        <>
          <Group label={t('word.ribbon.immersive-reader')}>
            <Button dropdown
              tall
              icon="alignJustify"
              label={t('word.ribbon.column-width')}
              title={t('word.ribbon.column-width-how')}
              onClick={(e) => menu.open(e, IR_WIDTHS.map(([k, label]) => ({ label, icon: view.immersive.width === k ? 'check' : undefined, run: () => act('immersiveSet', { width: k }) })))}
            />
            <Button dropdown
              tall
              icon="contrast"
              label={t('word.ribbon.page-color')}
              title={t('word.ribbon.page-color-colour')}
              onClick={(e) => menu.open(e, IR_COLOURS.map(([k, label]) => ({ label, icon: view.immersive.colour === k ? 'check' : undefined, run: () => act('immersiveSet', { colour: k }) })))}
            />
            <Button dropdown
              tall
              icon="eye"
              label={t('word.ribbon.line-focus')}
              title={t('word.ribbon.line-focus-one')}
              onClick={(e) => menu.open(e, IR_FOCUS.map(([n, label]) => ({ label, icon: view.immersive.focus === n ? 'check' : undefined, run: () => act('immersiveSet', { focus: n }) })))}
            />
            <Button tall icon="sliders" label={t('word.ribbon.text-spacing')} pressed={Boolean(view.immersive.spacing)} title={t('word.ribbon.text-spacing-more')} onClick={() => act('immersiveSet', { spacing: !view.immersive.spacing })} />
            <Button tall icon="scissors" label={t('word.ribbon.syllables')} pressed={Boolean(view.immersive.syllables)} title={t('word.ribbon.syllables-long-words')} onClick={() => act('immersiveSet', { syllables: !view.immersive.syllables })} />
            <Button tall icon="volume" label={t('word.ribbon.read-aloud')} pressed={Boolean(view.reading)} onClick={() => act('readAloud')} />
          </Group>
          <Group label={t('common.close')}>
            <Button tall icon="close" label={t('word.ribbon.close')} onClick={() => act('immersiveClose')} />
          </Group>
        </>
      ) : null}

      {tab === 'view' ? (
        <>
          <Group label={t('common.views')}>
            <Button tall icon="eye" label={t('word.ribbon.read-mode')} pressed={view.mode === 'read'} onClick={() => act('mode', 'read')} />
            <Button tall icon="file" label={t('word.ribbon.print-layout')} pressed={!view.mode || view.mode === 'print'} onClick={() => act('mode', 'print')} />
            <Button tall icon="globe" label={t('word.ribbon.web-layout')} pressed={view.mode === 'web'} onClick={() => act('mode', 'web')} />
            <Button icon="listBullet" label={t('common.outline')} pressed={view.mode === 'outline'} onClick={() => act('mode', 'outline')} />
            <Button icon="textbox" label={t('word.ribbon.draft-2')} pressed={view.mode === 'draft'} onClick={() => act('mode', 'draft')} />
          </Group>
          <Group label={t('word.ribbon.immersive')}>
            <Button tall icon="maximize" label={t('word.ribbon.focus')} pressed={Boolean(view.focus)} title={t('word.ribbon.just-page-full')} onClick={() => act('focus')} />
            <Button tall icon="eye" label={t('word.ribbon.immersive-reader')} pressed={Boolean(view.immersive)} title={t('word.ribbon.immersive-reader-words')} onClick={() => act(view.immersive ? 'immersiveClose' : 'immersive')} />
          </Group>
          <Group label={t('word.ribbon.page')}>
            <Button tall icon="pagesDown" label={t('common.vertical')} pressed={view.movement !== 'side'} title={t('word.ribbon.vertical-scroll-down')} onClick={() => act('movement', 'vertical')} />
            <Button tall icon="pagesAcross" label={t('word.ribbon.side-to-side')} pressed={view.movement === 'side'} title={t('word.ribbon.side-side-pages')} onClick={() => act('movement', 'side')} />
          </Group>
          <Group label={t('common.show')}>
            <Button icon="minus" label={t('common.ruler')} pressed={Boolean(view.ruler)} onClick={() => act('toggleRuler')} />
            <Button icon="grid" label={t('common.gridlines')} pressed={Boolean(view.gridlines)} title={t('word.ribbon.gridlines-quarter-inch')} onClick={() => act('toggleGridlines')} />
            <Button icon="list" label={t('word.ribbon.navigation')} pressed={Boolean(view.navigation)} title={t('word.ribbon.headings-move-around')} onClick={() => act('toggleNavigation')} />
          </Group>
          <Group label={t('common.zoom')}>
            <Button dropdown tall icon="find" label={t('common.zoom')} onClick={(e) => menu.open(e, [50, 75, 100, 125, 150, 200].map((z) => ({ label: `${z}%`, run: () => act('zoom', z / 100) })))} />
            <Button icon="check" label="100%" onClick={() => act('zoom', 1)} />
            <Button icon="file" label={t('word.ribbon.one-page')} onClick={() => act('zoom', 'page')} />
            <Button icon="copy" label={t('word.ribbon.multiple')} onClick={() => act('zoom', 'pages')} />
            <Button icon="maximize" label={t('word.ribbon.page-width')} onClick={() => act('zoom', 'width')} />
          </Group>
          <Group label={t('common.window')}>
            <Button tall icon="new" label={t('common.new-window')} title={t('word.ribbon.document-again-another')} onClick={() => act('newWindow')} />
            <Button icon="grid" label={t('common.arrange-all')} title={t('word.ribbon.arrange-all-every')} onClick={() => act('arrange', 'stack')} />
            <Button icon="minus" label={t('common.split')} pressed={Boolean(view.split)} title={view.split ? t('word.ribbon.remove-split-document') : t('word.ribbon.split-document-two')} onClick={() => act('toggleSplit')} />
            <Button icon="copy" label={t('word.ribbon.side-by-side')} title={t('word.ribbon.view-side-side')} onClick={() => act('arrange', 'sideBySide')} />
          </Group>
          <Group label={t('common.properties')}>
            <Button tall icon="info" label={t('common.properties')} onClick={() => openDialog('properties')} />
          </Group>
        </>
      ) : null}

      {/* ── Developer: content controls and legacy form fields (forms.js) ── */}
      {tab === 'developer' && forms ? <DeveloperTab forms={forms} menu={menu} act={act} view={view} /> : null}

      {/* ── Help ─────────────────────────────────────────────────────────── */}
      {tab === 'help' ? (
        <>
          <Group label={t('common.help')}>
            <Button tall icon="info" label={t('common.help')} onClick={() => act('open', 'help')} />
            <Button tall icon="reply" label={t('common.feedback')} onClick={() => act('open', 'contact')} />
            <Button tall icon="globe" label={t('common.whats-new')} onClick={() => act('open', 'releases')} />
            <Button tall icon="download" label={t('word.ribbon.check-for')} onClick={() => act('open', 'about')} />
          </Group>
          <Group label={t('word.ribbon.shortcuts')}>
            <Button tall icon="listBullet" label={t('common.keyboard')} onClick={() => openDialog('shortcuts')} />
          </Group>
        </>
      ) : null}

      {/* ── PDF ──────────────────────────────────────────────────────────── */}
      {tab === 'pdf' ? (
        <>
          <Group label={t('common.create')}>
            <Button tall icon="pdf" label={t('word.ribbon.create-a-pdf')} onClick={() => exportAs('pdf')} />
            {/* words-ok: a file format's name */}
            <Button tall icon="export" label="Markdown" onClick={() => exportAs('md')} />
            <Button tall icon="export" label={t('word.ribbon.plain-text')} onClick={() => exportAs('txt')} />
            <Button tall icon="export" label={t('word.ribbon.web-page')} onClick={() => exportAs('html')} />
          </Group>
          <Group label={t('common.print')}>
            <Button tall icon="print" label={t('common.print')} onClick={() => commands['file.print']?.run?.()} />
          </Group>
          <Group label={t('common.file')}>
            <Button tall icon="new" label={t('common.new')} onClick={() => shell.win.create({ app: 'word' })} />
            <Button tall icon="open" label={t('common.open')} onClick={openFile} />
            <Button tall icon="save" label={t('common.save')} onClick={() => save(false)} />
            <Button tall icon="save" label={t('common.save-as')} onClick={() => save(true)} />
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
