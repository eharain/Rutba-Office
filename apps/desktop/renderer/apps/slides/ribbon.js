// The Presentation ribbon.
//
// Laid out the way PowerPoint lays it out — Home, Insert, Draw, Design,
// Transitions, Animations, Slide Show, Record, Review, View, Help, and a PDF
// tab where PowerPoint has Acrobat — with every group PowerPoint has in each.
// What the engine can do is wired; what it cannot is drawn where PowerPoint
// draws it, disabled, with a title that says exactly why. Transitions are
// authored here and played in the show, and so are animations — the
// entrance, emphasis and exit effects PowerPoint's gallery offers first.

import React from 'react';
import { Ribbon, Group, Rows, Button, Separator, Select, Icon, t, tn } from '@rutba/office-ui';
import { WARP_PRESETS, WARP_MORE, warpLabel } from '@rutba/drawing/warp';
import { TRANSITION_GALLERY, TRANSITION_OPTIONS, galleryKeyOf, optionOf, describeTransition } from './motion.js';
import { ANIMATION_GALLERY, EFFECT_MENU, ANIMATION_OPTIONS } from './animate.js';
import { MODEL_VIEWS } from '@rutba/imaging/model3d';
import { RibbonStrip } from './design.js';
import { wordArtMenu } from '../../wordart.js';
import { SOUNDS, soundFile } from './sounds.js';
import { PEN_COLOURS, PEN_WIDTHS } from './ink.js';

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
const COLOURS = [
  ['#000000', t('Black')], ['#404040', t('Dark grey')], ['#808080', t('Grey')], ['#ffffff', t('White')],
  ['#c00000', t('Dark red')], ['#ff0000', t('Red')], ['#ffc000', t('Orange')], ['#ffff00', t('Yellow')],
  ['#92d050', t('Light green')], ['#00b050', t('Green')], ['#00b0f0', t('Light blue')], ['#0070c0', t('Blue')],
  ['#002060', t('Dark blue')], ['#7030a0', t('Purple')],
];
const LAYOUTS = [
  ['title', t('Title Slide')], ['obj', t('Title and Content')], ['blank', t('Blank')],
];
/** The preset geometries the engine writes and the renderer draws, in PowerPoint's names. */
const SHAPES = [
  ['rect', t('Rectangle')], ['roundRect', t('Rectangle: Rounded Corners')], ['ellipse', t('Oval')],
  ['triangle', t('Isosceles Triangle')], ['rtTriangle', t('Right Triangle')], ['diamond', t('Diamond')],
  ['parallelogram', t('Parallelogram')], ['trapezoid', t('Trapezoid')], ['pentagon', t('Pentagon')],
  ['hexagon', t('Hexagon')], ['octagon', t('Octagon')], ['star5', t('Star: 5 Points')],
  ['rightArrow', t('Arrow: Right')], ['chevron', t('Chevron')], ['line', t('Line')],
];
/** Insert → Table: the sizes PowerPoint's own gallery offers first. */
const TABLE_SIZES = [
  [2, 2, '2 × 2'], [3, 3, '3 × 3'], [4, 3, '4 × 3'], [5, 4, '5 × 4'],
];
/** Insert → Chart: the kinds `chartPartXml` writes — the ones a chart on a slide can be. */
const CHART_TYPES = [
  ['column', t('Column')], ['bar', t('Bar')], ['line', t('Line')], ['pie', t('Pie')],
];

/** Design → Background Styles: the theme backgrounds PowerPoint's gallery offers first, then a few flat colours. */
const BACKGROUND_STYLES = [
  [t('Background 1'), { scheme: 'bg1' }],
  [t('Background 2'), { scheme: 'bg2' }],
  [t('Text 1'), { scheme: 'tx1' }],
  [t('Text 2'), { scheme: 'tx2' }],
  [t('Accent 1 gradient'), { gradient: { from: { scheme: 'accent1' }, to: { scheme: 'accent1', lumMod: 75 } } }],
  '-',
  [t('White'), { colour: 'FFFFFF' }],
  [t('Black'), { colour: '000000' }],
  [t('Light grey'), { colour: 'F2F2F2' }],
  [t('Dark blue'), { colour: '1F3864' }],
];

/** A control that is drawn where PowerPoint draws it, and says why it is not live. */
const Soon = ({ icon, label, tall, why }) => (
  <Button tall={tall} icon={icon} label={label} disabled title={t('{label} — not built yet. {why}', { label, why })} />
);
/** Shape Effects: the shadows PowerPoint offers first, and none. */
const SHADOW_MENU = [['br', t('Shadow: bottom right')], ['b', t('Shadow: below')], ['r', t('Shadow: right')], ['tl', t('Shadow: top left')], ['c', t('Shadow: all round')], ['none', t('No shadow')]];
/** Shape Effects → Glow: PowerPoint's own gallery radii. */
const GLOW_MENU = [[5, t('Glow: 5 pt')], [8, t('Glow: 8 pt')], [11, t('Glow: 11 pt')], [18, t('Glow: 18 pt')], [null, t('No glow')]];
/** Shape Effects → Soft Edges. */
const SOFTEDGE_MENU = [[1, t('Soft Edges: 1 pt')], [2.5, t('Soft Edges: 2.5 pt')], [5, t('Soft Edges: 5 pt')], [10, t('Soft Edges: 10 pt')], [null, t('No soft edges')]];
/** Shape Effects → Reflection: the three gallery presets. */
const REFLECTION_MENU = [['tight', t('Reflection: tight')], ['half', t('Reflection: half')], ['full', t('Reflection: full')], [null, t('No reflection')]];
const INK = 'Ink is a drawing part (ink ML) the engine does not write, and the stage has no pen surface yet.';
/** The gallery's small pictures, one per effect. */
const TRANSITION_ICONS = { none: 'trNone', cut: 'trCut', fade: 'trFade', push: 'trPush', wipe: 'trWipe', split: 'trSplit', pull: 'trUncover', cover: 'trCover', randomBar: 'trBars', shape: 'trShape', dissolve: 'trDissolve' };

/**
 * A number of seconds in the ribbon (Duration, After): typed or stepped,
 * and sent once it settles — on Enter, on leaving the box, or a moment
 * after the last keystroke — rather than once per keystroke.
 */
function SecondsField({ value, onCommit, disabled = false, min = 0, max = 3600, className = '' }) {
  const shown = (v) => (v == null ? '' : Number(v).toFixed(2));
  const [text, setText] = React.useState(shown(value));
  const timer = React.useRef(null);
  React.useEffect(() => setText(shown(value)), [value]);
  React.useEffect(() => () => clearTimeout(timer.current), []);
  const commit = (raw) => {
    clearTimeout(timer.current);
    const v = Number(raw);
    if (raw === '' || !Number.isFinite(v)) return setText(shown(value));
    const clamped = Math.round(Math.min(max, Math.max(min, v)) * 100) / 100;
    if (value == null || Math.abs(clamped - value) > 0.0005) onCommit(clamped);
    else setText(shown(value));
  };
  return (
    <input
      type="number"
      step="0.25"
      min={min}
      max={max}
      className={`rw-input sl-rb-seconds ${className}`}
      value={text}
      disabled={disabled}
      onChange={(e) => {
        const raw = e.target.value;
        setText(raw);
        clearTimeout(timer.current);
        timer.current = setTimeout(() => commit(raw), 700);
      }}
      onBlur={(e) => commit(e.target.value)}
      onKeyDown={(e) => { if (e.key === 'Enter') commit(e.currentTarget.value); }}
    />
  );
}

export default function SlidesRibbon({
  tab, setTab, model, doc, commands, shell, menu, save, openFile, exportAs,
  act, view = {}, index = 0, selected = null, selectedIds = [], format = {}, canPaste = false, painter = false, animation = null, animPainter = false, addSlide, insertPicture, presentWithNotes, setPresent, setNotesOpen,
  monitorName = t('Automatic'),
  ink = null,
  designStrip = null,
  masterView = null,
  masterPart = null,
  review = null,

}) {
  const count = model?.count || 0;
  /** How many comment threads the deck holds. */
  const commentCount = (model?.comments || []).length;
  const hasShape = Boolean(selected);
  // Align and Distribute's own toggle — "Align Selected Objects" (the
  // default) or "Align to Slide" — is how the ribbon remembers which one is
  // ticked; it is a preference about the next press, not part of the file.
  const [alignTo, setAlignTo] = React.useState('selection');
  const selectedShapeObj = (model?.slide?.shapes || []).find((s) => s.id === selected) || null;
  const isGroup = selectedShapeObj?.kind === 'group';
  // A shape with an outline to edit point by point: not a picture, a table or a group.
  const canEditPoints = Boolean(selectedShapeObj?.kind === 'shape' && selectedShapeObj.geometry && (selectedShapeObj.path || selectedShapeObj.preset !== 'custom'));
  const multiCount = selectedIds.length;
  const size = Number(format.size || 18);
  const nearer = (dir) => {
    const bigger = SIZES.filter((s) => (dir > 0 ? s > size : s < size));
    return dir > 0 ? bigger[0] ?? size : bigger[bigger.length - 1] ?? size;
  };
  const fmt = (delta) => act('format', delta);
  const needShape = hasShape ? undefined : t('Select a text box first — click it once');
  const hasSections = (model?.sections || []).length > 0;
  const inSection = (model?.sections || []).some((s) => s.slides.includes(index));
  // Transitions: this slide's own, the gallery button it presses (none for
  // an effect this does not write), and that button's Effect Options.
  const transition = model?.slide?.transition || null;
  const transitionKey = transition?.known === false ? null : galleryKeyOf(transition?.type);
  const transitionOptions = TRANSITION_OPTIONS[transitionKey] || null;
  const transitionPlays = Boolean(transition && (transition.known === false || transitionKey !== 'none'));
  // Animations: the slide's sequence, whether the selected shape has any, and
  // the current effect's Effect Options (only for one this writes).
  const animationList = model?.slide?.animations || [];
  const shapeAnimated = selected != null && animationList.some((a) => String(a.shapeId) === String(selected));
  const animationOptions = animation?.known ? ANIMATION_OPTIONS[animation.effect] || null : null;
  // Effect Options → Sequence: a text shape's entrance, emphasis or exit as one object or paragraph by paragraph.
  const animShape = animation ? (model?.slide?.shapes || []).find((s) => String(s.id) === String(animation.shapeId)) : null;
  const animWords = (animShape?.text?.paragraphs || []).filter((p) => String(p.plain || '').trim()).length;
  const sequenceItems = animation && animation.kind !== 'path' && (animation.known || animation.paragraph != null) && animWords > 1
    ? [
        { heading: true, label: t('Sequence') },
        { label: t('As One Object'), icon: animation.paragraph == null ? 'check' : undefined, run: () => act('animSequence', 'object') },
        { label: t('All at Once'), run: () => act('animSequence', 'together') },
        { label: t('By Paragraph'), icon: animation.paragraph != null ? 'check' : undefined, run: () => act('animSequence', 'paragraph') },
      ]
    : [];
  /** More Effects and Add Animation: every effect under its heading, each a verb on the selected shape. */
  const effectMenu = (verb) => EFFECT_MENU.flatMap(([kind, heading, effects]) => [
    { heading: true, label: heading },
    ...effects.map(([effect, label]) => ({
      label,
      icon: verb === 'animate' && animation?.kind === kind && animation?.effect === effect ? 'check' : 'star',
      run: () => act(verb, { kind, effect }),
    })),
  ]);

  // Slide Master view: the part on the stage, the master it belongs to.
  const masterItem = masterView?.items?.find((it) => it.part === masterPart) || null;
  const masterOfItem = masterItem ? masterView.items.find((it) => it.part === (masterItem.kind === 'master' ? masterItem.part : masterItem.master)) : null;
  const isLayout = masterItem?.kind === 'layout';
  /** Design → Background Styles, and the Slide Master tab's own: the gallery, Reset, and (on slides) Apply to all. */
  const backgroundMenu = (e) => {
    const own = model?.slide?.ownBackground ?? null;
    const same = (spec) => JSON.stringify(spec ?? null) === JSON.stringify(own);
    menu.open(e, [
      ...BACKGROUND_STYLES.map((row) => (row === '-' ? '-' : { label: row[0], icon: same(row[1]) ? 'check' : undefined, run: () => act('background', { spec: row[1] }) })),
      '-',
      { label: masterView ? (isLayout ? t("Reset to the master's") : t('Reset to the theme\'s')) : t("Reset to the layout's"), icon: same(null) ? 'check' : undefined, run: () => act('background', { spec: null }) },
      ...(masterView ? [] : ['-', { label: t('Apply to all slides'), run: () => act('background', { spec: own, all: true }) }]),
    ]);
  };
  const galleryAt = (kind) => (e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind, anchor: { left: r.left, bottom: r.bottom + 4 } }); };

  return (
    <Ribbon
      tabs={[
        ...(masterView ? [masterView.kind === 'notes' ? { id: 'notesMaster', label: t('Notes Master') } : masterView.kind === 'handout' ? { id: 'handoutMaster', label: t('Handout Master') } : { id: 'master', label: t('Slide Master') }] : []),
        { id: 'home', label: t('Home') },
        { id: 'insert', label: t('Insert') },
        { id: 'draw', label: t('Draw') },
        { id: 'design', label: t('Design') },
        { id: 'transitions', label: t('Transitions') },
        { id: 'animations', label: t('Animations') },
        { id: 'show', label: t('Slide Show') },
        { id: 'record', label: t('Record') },
        { id: 'review', label: t('Review') },
        { id: 'view', label: t('View') },
        { id: 'help', label: t('Help') },
        { id: 'pdf', label: 'PDF' }, // words-ok: a file format's name
        // Table Design, as PowerPoint's: there while a table is selected.
        ...(selectedShapeObj?.kind === 'table' ? [{ id: 'tableDesign', label: t('Table Design') }] : []),
        // Shape Format, as PowerPoint's: there while a shape with words is selected.
        ...(selectedShapeObj?.kind === 'shape' && selectedShapeObj?.text && !selectedShapeObj?.cameo ? [{ id: 'shapeFormat', label: t('Shape Format') }] : []),
        // Camera Format, as PowerPoint's: there while a cameo is selected.
        ...(selectedShapeObj?.cameo ? [{ id: 'cameraFormat', label: t('Camera Format') }] : []),
        // 3D Model, as PowerPoint's: there while a 3D model is selected.
        ...(selectedShapeObj?.model3d ? [{ id: 'model3d', label: t('3D Model') }] : []),
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title={t('Save (Ctrl+S)')} onClick={() => save(false)} />
          <Button icon="undo" title={t('Undo (Ctrl+Z)')} disabled={!doc?.canUndo} onClick={() => commands['edit.undo']?.run?.()} />
          <Button icon="play" title={t('Start from the beginning (F5)')} onClick={() => act('present', 'start')} />
        </>
      }
    >
      {/* ── Home ─────────────────────────────────────────────────────────── */}
      {tab === 'home' ? (
        <>
          <Group label={t('Clipboard')}>
            <Button tall icon="paste" label={t('Paste')} title={canPaste ? t('Paste the copied shape on this slide (Ctrl+V)') : t('Paste — copy a shape first, then Ctrl+V')} disabled={!canPaste} onClick={() => act('pasteShape')} />
            <Rows>
              <>
                <Button icon="cut" label={t('Cut')} title={needShape || t('Cut the shape (Ctrl+X)')} disabled={!hasShape} onClick={() => act('cutShape')} />
                <Button icon="copy" label={t('Copy')} title={needShape || t('Copy the shape (Ctrl+C)')} disabled={!hasShape} onClick={() => act('copyShape')} />
              </>
              <Button
                icon="wand"
                label={t('Format Painter')}
                pressed={painter}
                disabled={!hasShape && !painter}
                title={painter ? t('Format Painter — armed: click a shape to give it this look; Esc puts it down') : (needShape || t('Format Painter — this shape’s fill, outline and font onto the next shape you click'))}
                onClick={() => act('painter')}
              />
            </Rows>
          </Group>
          <Group label={t('Slides')}>
            <Button tall icon="plus" label={t('New Slide')} onClick={(e) => menu.open(e, LAYOUTS.map(([layout, label]) => ({ label, icon: 'slides', run: () => addSlide(layout) })))} />
            <Rows>
              <>
                <Button icon="grid" label={t('Layout')} title={t("Layout — put this slide on another of the deck's layouts")} onClick={(e) => menu.open(e, (model?.layouts || []).map((l) => ({ label: l.name || l.part, icon: l.part === model?.slide?.layout ? 'check' : undefined, run: () => act('applyLayout', l.part) })))} />
                <Button icon="undo" label={t('Reset')} title={t('Reset — the placeholders back where the layout puts them')} onClick={() => act('resetSlide')} />
                <Button
                  icon="list"
                  label={t('Section')}
                  title={t('Section — a section before this slide, its name, or one taken away')}
                  onClick={(e) => menu.open(e, [
                    { label: t('Add Section'), icon: 'plus', run: () => act('addSection') },
                    { label: t('Rename Section…'), icon: 'textbox', disabled: !inSection, run: () => act('renameSection') },
                    { label: t('Remove Section'), icon: 'trash', disabled: !inSection, run: () => act('removeSection') },
                    { label: t('Remove All Sections'), icon: 'close', disabled: !hasSections, run: () => act('removeAllSections') },
                  ])}
                />
              </>
              <>
                <Button icon="copy" label={t('Duplicate')} onClick={() => commands['slide.new']?.run?.()} />
                <Button icon="trash" label={t('Delete')} disabled={count < 2} onClick={() => commands['slide.delete']?.run?.()} />
              </>
            </Rows>
          </Group>
          <Group label={t('Font')}>
            <Rows>
              <>
                <Select value={format.font || ''} onChange={(e) => fmt({ font: e.target.value })} style={{ width: 118 }} title={needShape || t('Font')} disabled={!hasShape}>
                  <option value="">{format.font || (format.themeFont ? t('{font} (theme)', { font: format.themeFont }) : t('Theme font'))}</option>
                  {['Calibri', 'Calibri Light', 'Arial', 'Segoe UI', 'Georgia', 'Times New Roman', 'Verdana', 'Consolas'].filter((f) => f !== format.font).map((f) => <option key={f} value={f}>{f}</option>)}
                </Select>
                <Select value={String(size)} onChange={(e) => fmt({ size: Number(e.target.value) })} style={{ width: 58 }} title={needShape || t('Font size')} disabled={!hasShape}>
                  {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
                  {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
                </Select>
                <Button icon="chevronUp" title={needShape || t('Increase font size')} disabled={!hasShape} onClick={() => fmt({ size: nearer(1) })} />
                <Button icon="chevronDown" title={needShape || t('Decrease font size')} disabled={!hasShape} onClick={() => fmt({ size: nearer(-1) })} />
                <Button icon="undo" title={needShape || t('Clear all formatting — the words keep only their links')} disabled={!hasShape} onClick={() => act('clearFormat')} />
              </>
              <>
                <Button icon="bold" title={needShape || t('Bold (Ctrl+B)')} pressed={format.bold} disabled={!hasShape} onClick={() => fmt({ bold: 'toggle' })} />
                <Button icon="italic" title={needShape || t('Italic (Ctrl+I)')} pressed={format.italic} disabled={!hasShape} onClick={() => fmt({ italic: 'toggle' })} />
                <Button icon="underline" title={needShape || t('Underline (Ctrl+U)')} pressed={format.underline} disabled={!hasShape} onClick={() => fmt({ underline: 'toggle' })} />
                <Button icon="strike" title={needShape || t('Strikethrough')} pressed={format.strike} disabled={!hasShape} onClick={() => fmt({ strike: 'toggle' })} />
                <Button icon="textbox" label="AV" title={needShape || (format.spacing ? t('Character spacing — now {spacing} pt', { spacing: format.spacing }) : t('Character spacing — now normal'))} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('Very tight'), -1.5], [t('Tight'), -0.75], [t('Normal'), 0], [t('Loose'), 1.5], [t('Very loose'), 3]].map(([label, v]) => ({ label, icon: (format.spacing || 0) === v ? 'check' : undefined, run: () => fmt({ spacing: v }) })))} />{/* words-ok: AV is the control's picture of letter spacing, as PowerPoint draws it */}
                <Button icon="textbox" label="Aa" title={needShape || t('Change case')} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('Sentence case'), 'sentence'], [t('lowercase'), 'lower'], [t('UPPERCASE'), 'upper'], [t('Capitalise Each Word'), 'title']].map(([label, mode]) => ({ label, run: () => fmt({ case: mode }) })))} />{/* words-ok: Aa is the control's picture of letter case, as PowerPoint draws it */}
                <Separator />
                <Button icon="wand" title={needShape || t('Text highlight colour')} pressed={Boolean(format.highlight)} disabled={!hasShape} onClick={(e) => menu.open(e, [['#FFFF00', t('Yellow')], ['#00FF00', t('Bright green')], ['#00FFFF', t('Turquoise')], ['#FF00FF', t('Pink')], [null, t('No colour')]].map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => fmt({ highlight: value }) })))} />
                <Button icon="contrast" title={needShape || t('Font colour')} disabled={!hasShape} onClick={(e) => menu.open(e, COLOURS.map(([value, label]) => ({ label, run: () => fmt({ color: value }) })))} />
              </>
            </Rows>
          </Group>
          <Group label={t('Paragraph')}>
            <Rows>
              <>
                <Button icon="listBullet" title={needShape || t('Bullets')} pressed={format.bullet === 'char'} disabled={!hasShape} onClick={() => fmt({ bullet: format.bullet === 'char' ? 'none' : 'char' })} />
                <Button icon="listNumber" title={needShape || t('Numbering')} pressed={format.bullet === 'number'} disabled={!hasShape} onClick={() => fmt({ bullet: format.bullet === 'number' ? 'none' : 'number' })} />
                <Button icon="chevronLeft" title={needShape || t('Decrease list level')} disabled={!hasShape || !format.level} onClick={() => fmt({ level: -1 })} />
                <Button icon="chevronRight" title={needShape || t('Increase list level')} disabled={!hasShape} onClick={() => fmt({ level: 1 })} />
                <Button icon="list" title={needShape || (format.lineHeight ? t('Line spacing — now {value}', { value: String(format.lineHeight) }) : t('Line spacing — now as the layout has it'))} disabled={!hasShape} onClick={(e) => menu.open(e, [1, 1.15, 1.5, 2].map((v) => ({ label: String(v), icon: format.lineHeight === v ? 'check' : undefined, run: () => fmt({ lineHeight: v }) })))} />
                <Separator />
                <Button icon="grid" title={needShape || t('Columns — now {columns}', { columns: format.columns || 1 })} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('One column'), 1], [t('Two columns'), 2], [t('Three columns'), 3]].map(([label, n]) => ({ label, icon: (format.columns || 1) === n ? 'check' : undefined, run: () => act('body', { columns: n }) })))} />
                <Button icon="rotate" title={needShape || t('Text direction — now {direction}', { direction: ({ horz: t('horizontal'), vert270: t('rotated up'), vert: t('rotated down'), eaVert: t('stacked') })[format.vert] || t('horizontal') })} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('Horizontal'), 'horz'], [t('Rotate all text 90° (reads down)'), 'vert'], [t('Rotate all text 270° (reads up)'), 'vert270']].map(([label, v]) => ({ label, icon: (format.vert || 'horz') === v ? 'check' : undefined, run: () => act('body', { vert: v }) })))} />
                <Button icon="chevronUp" title={needShape || t('Align text — now {anchor}', { anchor: ({ top: t('top'), middle: t('middle'), bottom: t('bottom') })[format.anchor || 'top'] || format.anchor })} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('Top'), 'top'], [t('Middle'), 'middle'], [t('Bottom'), 'bottom']].map(([label, a]) => ({ label, icon: (format.anchor || 'top') === a ? 'check' : undefined, run: () => act('body', { anchor: a }) })))} />
              </>
              <>
                <Button icon="alignLeft" title={needShape || t('Align left')} pressed={format.align === 'left'} disabled={!hasShape} onClick={() => fmt({ align: 'left' })} />
                <Button icon="alignCenter" title={needShape || t('Centre')} pressed={format.align === 'center'} disabled={!hasShape} onClick={() => fmt({ align: 'center' })} />
                <Button icon="alignRight" title={needShape || t('Align right')} pressed={format.align === 'right'} disabled={!hasShape} onClick={() => fmt({ align: 'right' })} />
                <Button icon="alignJustify" title={needShape || t('Justify')} pressed={format.align === 'justify'} disabled={!hasShape} onClick={() => fmt({ align: 'justify' })} />
                <Button icon="textLtr" title={needShape || t('Left-to-right text direction')} pressed={hasShape && !format.rtl} disabled={!hasShape} onClick={() => fmt({ rtl: false })} />
                <Button icon="textRtl" title={needShape || t('Right-to-left text direction — the words read from the right, as Arabic and Hebrew do')} pressed={Boolean(format.rtl)} disabled={!hasShape} onClick={() => fmt({ rtl: true })} />
                <Separator />
                <Button icon="shape" label={t('SmartArt')} title={needShape || t('Convert to SmartArt — this box\'s lines as a diagram: a list, a process, a cycle or a hierarchy')} disabled={!hasShape} onClick={() => act('convertSmartArt')} />
              </>
            </Rows>
          </Group>
          <Group label={t('Drawing')}>
            <Button tall icon="shape" label={t('Shapes')} title={t("Shapes — a rectangle, an oval, an arrow, a star, in the theme's colours")} onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => act('addShape', preset) })))} />

            <Button tall icon="grid" label={t('Arrange')} onClick={(e) => menu.open(e, [
              { label: t('Bring to front'), icon: 'chevronUp', run: () => act('order', 'front') },
              { label: t('Bring forward'), run: () => act('order', 'forward') },
              { label: t('Send backward'), run: () => act('order', 'backward') },
              { label: t('Send to back'), icon: 'chevronDown', run: () => act('order', 'back') },
              '-',
              { label: t('Group'), icon: 'grid', disabled: multiCount < 2, title: multiCount < 2 ? t('Select two or more shapes to group') : undefined, run: () => act('group') },
              { label: t('Ungroup'), icon: 'grid', disabled: !isGroup, title: !isGroup ? t('Select a group to ungroup') : undefined, run: () => act('ungroup') },
              '-',
              { label: t('Align Left'), run: () => act('align', { edge: 'left', to: alignTo }) },
              { label: t('Align Center'), run: () => act('align', { edge: 'center', to: alignTo }) },
              { label: t('Align Right'), run: () => act('align', { edge: 'right', to: alignTo }) },
              { label: t('Align Top'), run: () => act('align', { edge: 'top', to: alignTo }) },
              { label: t('Align Middle'), run: () => act('align', { edge: 'middle', to: alignTo }) },
              { label: t('Align Bottom'), run: () => act('align', { edge: 'bottom', to: alignTo }) },
              { label: t('Distribute Horizontally'), disabled: multiCount < 3, title: multiCount < 3 ? t('Select three or more shapes to distribute') : undefined, run: () => act('distribute', { axis: 'horizontal', to: alignTo }) },
              { label: t('Distribute Vertically'), disabled: multiCount < 3, title: multiCount < 3 ? t('Select three or more shapes to distribute') : undefined, run: () => act('distribute', { axis: 'vertical', to: alignTo }) },
              '-',
              { label: t('Align to Slide'), icon: alignTo === 'slide' ? 'check' : undefined, run: () => setAlignTo('slide') },
              { label: t('Align Selected Objects'), icon: alignTo === 'selection' ? 'check' : undefined, run: () => setAlignTo('selection') },
              '-',
              { label: t('Rotate Right 90°'), run: () => act('rotateBy', 90) },
              { label: t('Rotate Left 90°'), run: () => act('rotateBy', -90) },
              { label: t('Flip Vertical'), run: () => act('flipShape', 'vertical') },
              { label: t('Flip Horizontal'), run: () => act('flipShape', 'horizontal') },
              '-',
              { label: t('Selection pane (layers)'), icon: 'list', run: () => act('pane', 'layers') },
              '-',
              { label: t('Delete shape'), icon: 'trash', run: () => act('deleteShape') },
              { label: t('Move up (nudge)'), run: () => act('nudge', { dy: -8 }) },
              { label: t('Move down (nudge)'), run: () => act('nudge', { dy: 8 }) },
              { label: t('Move left (nudge)'), run: () => act('nudge', { dx: -8 }) },
              { label: t('Move right (nudge)'), run: () => act('nudge', { dx: 8 }) },
            ])} disabled={!hasShape} title={needShape || t('Arrange the selected shape(s)')} />
            <Button tall icon="wand" label={t('Quick Styles')} disabled={!hasShape} title={needShape || t('The theme\'s own looks: filled in an accent, outlined in the same')} onClick={(e) => menu.open(e, [1, 2, 3, 4, 5, 6].map((n) => ({ label: t('Accent {n}', { n }), icon: 'shape', run: () => act('quickStyle', n) })))} />
            <Button icon="wand" label={t('Shape Fill')} disabled={!hasShape} title={needShape || t('The fill of the selected shape, in the Format pane')} onClick={() => act('formatPane')} />
            <Button icon="shape" label={t('Shape Outline')} disabled={!hasShape} title={needShape || t('The outline of the selected shape, in the Format pane')} onClick={() => act('formatPane')} />
            <Button icon="wand" label={t('Shape Effects')} disabled={!hasShape} title={needShape || t('Shape Effects — a shadow, a glow, soft edges or a reflection on the selected shape')} onClick={(e) => menu.open(e, [
              ...SHADOW_MENU.map(([key, label]) => ({ label, icon: key === 'none' ? 'close' : undefined, run: () => act('shapeShadow', key) })),
              '-',
              ...GLOW_MENU.map(([pt, label]) => ({ label, icon: pt === null ? 'close' : undefined, run: () => act('shapeEffects', { glow: pt === null ? null : { radius: pt, color: selectedShapeObj?.effects?.glow?.color } }) })),
              '-',
              ...SOFTEDGE_MENU.map(([pt, label]) => ({ label, icon: pt === null ? 'close' : undefined, run: () => act('shapeEffects', { softEdge: pt === null ? null : { radius: pt } }) })),
              '-',
              ...REFLECTION_MENU.map(([key, label]) => ({ label, icon: key === null ? 'close' : undefined, run: () => act('shapeEffects', { reflection: key }) })),
            ])} />
          </Group>
          <Group label={t('Edit Shape')}>
            <Button tall icon="shape" label={t('Edit Shape')} disabled={!canEditPoints} title={canEditPoints ? t('Edit Points — drag the selected shape\'s points, add one on its outline or delete one') : hasShape ? t('A picture, a chart, a table or a group has no points to edit — select a shape') : needShape} onClick={(e) => menu.open(e, [
              { label: t('Edit Points'), icon: 'shape', run: () => act('editPoints') },
            ])} />
          </Group>
          <Group label={t('Editing')}>
            <Button tall icon="find" label={t('Find')} title={t('Find — words on every slide, walked one hit at a time (Ctrl+F)')} onClick={() => act('find')} />
            <Button tall icon="find" label={t('Replace')} title={t('Replace — find and replace words across the deck (Ctrl+H)')} onClick={() => act('replace')} />
          </Group>
          <Group label={t('Voice')}>
            <Soon tall icon="volume" label={t('Dictate')} why={t('Dictation is an online speech service this suite does not call.')} />
          </Group>
          <Group label={t('Designer')}>
            <Button tall icon="wand" label={t('Design Ideas')} pressed={view?.pane === 'ideas'} title={t("Design Ideas — layouts this slide's title, words and pictures suit, worked out on this computer; a click applies one")} onClick={() => act('pane', 'ideas')} />
          </Group>
        </>
      ) : null}

      {/* ── Insert ───────────────────────────────────────────────────────── */}
      {/* ── Slide Master (only while the master and its layouts are on the stage) ── */}
      {tab === 'master' && masterView ? (
        <>
          <Group label={t('Edit Master')}>
            <Button tall icon="slides" label={t('Insert Slide Master')} title={t('Insert Slide Master — a second master, a copy of this one with its layouts and theme, to restyle on its own')} onClick={() => act('insertMaster')} />
            <Button tall icon="plus" label={t('Insert Layout')} title={t('Insert Layout — a new layout on this master, with a title and the footers')} onClick={() => act('insertLayout')} />
            <Rows>
              <Button icon="trash" label={t('Delete')} disabled={!isLayout || masterItem?.used > 0} title={!isLayout ? t('Delete — pick a layout; the master stays') : masterItem?.used ? tn(masterItem.used, 'Delete — a slide uses this layout; put it on another first', 'Delete — {count} slides use this layout; put them on another first') : t('Delete — this layout, which no slide uses')} onClick={() => act('deleteLayout')} />
              <Button icon="textbox" label={t('Rename')} title={isLayout ? t("Rename — this layout's name") : t("Rename — this master's name")} onClick={() => act('renameLayout')} />
              <Button icon="lock" label={t('Preserve')} pressed={Boolean(masterOfItem?.preserve)} title={t('Preserve — keep this master in the file even when no slide uses it')} onClick={() => act('preserve')} />
            </Rows>
          </Group>
          <Group label={t('Master Layout')}>
            <Button tall icon="grid" label={t('Insert Placeholder')} disabled={!isLayout} title={isLayout ? t('Insert Placeholder — a content, text or picture placeholder on this layout') : t('Insert Placeholder — pick a layout; the master holds the placeholders every layout draws from')} onClick={(e) => menu.open(e, [['content', t('Content')], ['text', t('Text')], ['picture', t('Picture')]].map(([kind, label]) => ({ label, icon: kind === 'picture' ? 'picture' : 'textbox', run: () => act('insertPlaceholder', kind) })))} />
            <Rows>
              <label className="sl-rb-field sl-rb-check" data-tip={t("Title — this layout's title placeholder, or none")}><input type="checkbox" className="sl-master-title" checked={Boolean(masterItem?.hasTitle)} onChange={(e) => act('masterPlaceholders', { title: e.target.checked })} /> {t('Title')}</label>
              <label className="sl-rb-field sl-rb-check" data-tip={t('Footers — the date, footer and slide number placeholders, or none')}><input type="checkbox" className="sl-master-footers" checked={Boolean(masterItem?.hasFooters)} onChange={(e) => act('masterPlaceholders', { footers: e.target.checked })} /> {t('Footers')}</label>
            </Rows>
          </Group>
          <Group label={t('Edit Theme')}>
            <Button tall icon="wand" label={t('Themes')} title={t('Themes — every theme, drawn on the first slide')} onClick={galleryAt('themes')} />
            <Rows>
              <Button icon="contrast" label={t('Colours')} title={t("Colours — the theme's twelve colours")} onClick={galleryAt('colours')} />
              <Button icon="textbox" label={t('Fonts')} title={t("Fonts — the theme's heading and body faces")} onClick={galleryAt('fonts')} />
              <Button icon="wand" label={t('Effects')} title={t("Effects — the theme's format scheme")} onClick={galleryAt('effects')} />
            </Rows>
          </Group>
          <Group label={t('Background')}>
            <Button tall icon="picture" label={t('Background Styles')} title={isLayout ? t("Background Styles — this layout's background") : t("Background Styles — the master's background")} onClick={backgroundMenu} />
            <Rows>
              <label className="sl-rb-field sl-rb-check" data-tip={isLayout ? t("Hide Background Graphics — this layout without the master's own shapes") : t('Hide Background Graphics — pick a layout; the master always draws its own')}><input type="checkbox" className="sl-master-hidebg" disabled={!isLayout} checked={Boolean(masterItem?.hidesBackgroundGraphics)} onChange={(e) => act('hideBackgroundGraphics', e.target.checked)} /> {t('Hide Background Graphics')}</label>
            </Rows>
          </Group>
          <Group label={t('Close')}>
            <Button tall icon="close" label={t('Close Master View')} title={t('Close Master View — back to the slides')} onClick={() => act('closeMaster')} />
          </Group>
        </>
      ) : null}

      {(tab === 'notesMaster' || tab === 'handoutMaster') && (masterView?.kind === 'notes' || masterView?.kind === 'handout') ? (
        <>
          <Group label={t('Page Setup')}>
            <Button tall icon="file" label={masterView.kind === 'notes' ? t('Notes Page Orientation') : t('Handout Orientation')} title={t('Orientation — the printed page portrait or landscape; the notes page and the handout share one size')} onClick={(e) => menu.open(e, [
              { label: t('Portrait'), icon: masterView.portrait ? 'check' : undefined, run: () => act('notesOrientation', 'portrait') },
              { label: t('Landscape'), icon: masterView.portrait ? undefined : 'check', run: () => act('notesOrientation', 'landscape') },
            ])} />
            {masterView.kind === 'handout' ? (
              <Button tall icon="grid" label={t('Slides Per Page')} title={t('Slides Per Page — how many slides each handout page shows')} onClick={(e) => menu.open(e, [1, 2, 3, 4, 6, 9].map((n) => ({ label: tn(n, '{count} Slide', '{count} Slides'), icon: (view.handoutPer || 6) === n ? 'check' : undefined, run: () => act('handoutPer', n) })))} />
            ) : null}
          </Group>
          <Group label={t('Placeholders')}>
            <Rows>
              {(masterView.kind === 'notes'
                ? [['hdr', t('Header')], ['sldImg', t('Slide Image')], ['ftr', t('Footer')], ['dt', t('Date')], ['body', t('Body')], ['sldNum', t('Page Number')]]
                : [['hdr', t('Header')], ['ftr', t('Footer')], ['dt', t('Date')], ['sldNum', t('Page Number')]]
              ).map(([type, label]) => (
                <label key={type} className="sl-rb-field sl-rb-check" data-tip={t('{label} — this placeholder on the master, or none', { label })}><input type="checkbox" className={`sl-nm-ph sl-nm-${type}`} checked={Boolean(masterView.placeholders?.[type])} onChange={(e) => act('masterPlaceholder', { type, on: e.target.checked })} /> {label}</label>
              ))}
            </Rows>
          </Group>
          <Group label={t('Close')}>
            <Button tall icon="close" label={t('Close Master View')} title={t('Close Master View — back to the slides')} onClick={() => act('closeMaster')} />
          </Group>
        </>
      ) : null}

      {tab === 'insert' ? (
        <>
          <Group label={t('Slides')}>
            <Button tall icon="plus" label={t('New Slide')} onClick={(e) => menu.open(e, LAYOUTS.map(([layout, label]) => ({ label, icon: 'slides', run: () => addSlide(layout) })))} />
          </Group>
          <Group label={t('Tables')}>
            <Button tall icon="table" label={t('Table')} title={t("Table — a grid of cells, in PowerPoint's own default style")} onClick={(e) => menu.open(e, TABLE_SIZES.map(([rows, cols, label]) => ({ label, icon: 'table', run: () => act('addTable', { rows, cols }) })))} />
          </Group>
          <Group label={t('Images')}>
            <Button tall icon="picture" label={t('Pictures')} title={t('Pictures — a picture from this device, onto this slide')} onClick={insertPicture} />

            <Button tall icon="picture" label={t('Screenshot')} title={t('Screenshot — a picture of another open window, or of a whole screen, put on this slide')} onClick={() => act('screenshot')} />
            <Button tall icon="picture" label={t('Photo Album')} title={t('Photo Album — a new presentation of your pictures, one, two or four to a slide, captioned if you like')} onClick={() => act('photoAlbum')} />
          </Group>
          <Group label={t('Camera')}>
            <Button tall icon="video" label={t('Cameo')} className="sl-cameo-insert" title={t('Cameo — your camera, live, in a shape on the slide; the show fills it, and Camera Format → Preview shows it here')} onClick={() => act('cameo')} />
          </Group>
          <Group label={t('Illustrations')}>
            <Button tall icon="shape" label={t('Shapes')} title={t("Shapes — a rectangle, an oval, an arrow, a star, in the theme's colours")} onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => act('addShape', preset) })))} />

            <Button tall icon="star" label={t('Icons')} title={t("Icons — one of the suite's own icons, in the colour you choose, as a picture on this slide")} onClick={() => act('icons')} />
            <Button tall icon="shape" label={t('3D Models')} className="sl-model3d-insert" title={t('3D Models — a model from a .glb or .gltf file on this computer, drawn on the slide; turn it from the 3D Model tab or by its handle')} onClick={() => act('model3d')} />
            <Button tall icon="shape" label={t('SmartArt')} title={t('SmartArt — a list, a process, a cycle or a hierarchy, drawn from lines you type, as a group of shapes')} onClick={() => act('insertSmartArt')} />
            <Button tall icon="chart" label={t('Chart')} title={t('Chart — a sample chart, drawn from the writer Documents and Worksheets already use; double-click it to edit its data')} onClick={(e) => menu.open(e, CHART_TYPES.map(([type, label]) => ({ label, icon: 'chart', run: () => act('addChart', { type }) })))} />
          </Group>
          <Group label={t('Forms')}>
            <Soon tall icon="check" label={t('Forms')} why={t('Forms is a Microsoft cloud service.')} />
          </Group>
          <Group label={t('Links')}>
            <Button tall icon="zoomIn" label={t('Zoom')} title={t('Zoom — a picture of a slide or a section that takes the show there with a click, and back here after')} onClick={(e) => menu.open(e, [{ label: t('Summary Zoom…'), icon: 'grid', run: () => act('insertZoom', 'summary') }, { label: t('Section Zoom…'), icon: 'list', disabled: !(model?.sections || []).length, run: () => act('insertZoom', 'section') }, { label: t('Slide Zoom…'), icon: 'slides', run: () => act('insertZoom', 'slide') }])} />
            <Button tall icon="link" label={t('Link')} disabled={!hasShape} title={needShape || t('Link — a web address on the selected shape’s words (Ctrl+K); Ctrl+click follows it')} onClick={() => act('link')} />
            <Button tall icon="play" label={t('Action')} title={t('Action — what a click on the selected shape does in the show: another slide, the end of the show, or a web address')} onClick={() => act('action')} />
          </Group>
          <Group label={t('Comments')}>
            <Button tall icon="reply" label={t('Comment')} title={t('Comment — a new comment on the selected shape, or on this slide')} onClick={() => act('newComment')} />
          </Group>
          <Group label={t('Text')}>
            <Button tall icon="textbox" label={t('Text Box')} onClick={() => commands['slide.textbox']?.run?.()} />
            <Button tall icon="file" label={t('Header & Footer')} title={t("Header & Footer — the date, the slide number and the footer's words, on this slide or all of them")} onClick={() => act('footer')} />
            <Button tall icon="wand" label={t('WordArt')} title={t('WordArt — big words in a style of their own: a fill, an outline round the letters, a shadow or a glow')} onClick={(e) => menu.open(e, wordArtMenu((style) => act('wordArt', style)))} />
            <Button icon="clock" label={t('Date & Time')} title={t('Date & Time — a date along the bottom, kept current or fixed')} onClick={() => act('footer', 'date')} />
            <Button icon="list" label={t('Slide Number')} title={t("Slide Number — the slide's number along the bottom, following it when slides move")} onClick={() => act('footer', 'number')} />
            <Button icon="file" label={t('Object')} title={t('Object — a Word, Excel or PowerPoint document embedded on the slide as an icon; a double-click opens it')} onClick={() => act('insertObject')} />
          </Group>
          <Group label={t('Symbols')}>
            <Button tall icon="formula" label={t('Equation')} title={t("Equation — Office Math on this slide, typed in Word's linear format with a live preview; double-click one to edit it")} onClick={() => act('equation')} />
            <Button tall icon="star" label={t('Symbol')} title={t('Symbol — a character from the symbol sets, put in where the caret is in a text box')} onClick={() => act('symbol')} />
          </Group>
          <Group label={t('Media')}>
            <Button tall icon="video" label={t('Video')} title={t('Video — a video from this computer on the slide; a click on it plays it in the show')} onClick={(e) => menu.open(e, [{ label: t('This Device…'), icon: 'video', run: () => act('insertMedia', 'video') }])} />
            <Button tall icon="volume" label={t('Audio')} title={t('Audio — a sound from this computer on the slide, as a speaker; a click on it plays it in the show')} onClick={(e) => menu.open(e, [{ label: t('Audio on My PC…'), icon: 'volume', run: () => act('insertMedia', 'audio') }])} />
            <Button tall icon="video" label={t('Screen Recording')} title={t('Screen Recording — record a window or a screen until Stop, and put the recording on this slide as a video')} onClick={() => act('screenRecording')} />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label={t('Drawing Tools')}>
            <Button tall icon="mouse" label={t('Select')} pressed={!ink?.tool} title={t("Select — put the pen down and work with the slide's shapes again (Esc)")} onClick={() => act('inkTool', null)} />
            <Button tall icon="wand" label={t('Lasso')} pressed={ink?.tool === 'lasso'} title={t('Lasso Select — draw a loop round strokes to select them')} onClick={() => act('inkTool', 'lasso')} />
            <Button tall icon="close" label={t('Eraser')} pressed={ink?.tool === 'eraser'} title={t('Eraser — take away each stroke the pointer passes over')} onClick={() => act('inkTool', 'eraser')} />
            {(ink?.pens || []).map((p) => {
              const on = ink.tool === 'pen' && ink.penId === p.id;
              const label = { pen: t('Pen'), pencil: t('Pencil'), highlighter: t('Highlighter') }[p.tool];
              return (
                <Button key={p.id} tall icon="wand" label={label} pressed={on} className={`sl-pen sl-pen-${p.tool}`} data-pen={p.id}
                  style={{ '--pen': p.color }}
                  title={on ? t('{label} — click again for its colour and thickness', { label }) : t('{label} — draw on the slide', { label })}
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
            <Button tall icon="plus" label={t('Add')} title={t('Add Pen — another pen, pencil or highlighter in the gallery')} onClick={(e) => menu.open(e, [['pen', t('Pen')], ['pencil', t('Pencil')], ['highlighter', t('Highlighter')]].map(([tool, label]) => ({ label, icon: 'plus', run: () => act('inkAdd', tool) })))} />
          </Group>
          <Group label={t('Touch')}>
            <Button tall icon="wand" label={t('Draw with Touch')} pressed={ink?.touch !== false} title={t('Draw with Touch — a finger draws with the pen in hand; off, only a pen or the mouse does')} onClick={() => act('inkTouch')} />
          </Group>
          <Group label={t('Stencils')}>
            <Button tall icon="minus" label={t('Ruler')} pressed={Boolean(ink?.ruler)} title={t('Ruler — a straight edge on the slide: a stroke begun along it follows it; drag it to move, the wheel to turn it')} onClick={() => act('inkRuler')} />
          </Group>
          <Group label={t('Convert')}>
            <Button tall icon="shape" label={t('Ink to Shape')} pressed={Boolean(ink?.toShape)} title={t('Ink to Shape — a rectangle, oval or triangle drawn becomes that shape')} onClick={() => act('inkToShape')} />
            <Soon tall icon="formula" label={t('Ink to Math')} why={t('Turning handwriting into an equation needs handwriting recognition, which this suite does not have.')} />
          </Group>
          <Group label={t('Replay')}>
            <Button tall icon="play" label={t('Ink Replay')} title={t("Ink Replay — the slide's strokes drawn again, in the order they were made")} onClick={() => act('inkReplay')} />
          </Group>
          <Group label={t('Help')}>
            <Button tall icon="info" label={t('Ink Help')} onClick={() => act('help')} />
          </Group>
        </>
      ) : null}

      {/* ── Design ───────────────────────────────────────────────────────── */}
      {tab === 'design' ? (
        <>
          <Group label={t('Themes')}>
            <RibbonStrip
              kind="themes"
              items={designStrip?.themes}
              count={4}
              label={t('Drawing this slide…')}
              onPick={(it) => act('applyTheme', it)}
              onMore={(e) => { const r = e.currentTarget.closest('.sl-rs').getBoundingClientRect(); act('designGallery', { kind: 'themes', anchor: { left: r.left, bottom: r.bottom + 4 } }); }}
            />
          </Group>
          <Group label={t('Variants')}>
            <RibbonStrip
              kind="variants"
              items={designStrip?.variants}
              count={4}
              label={t('Drawing this slide…')}
              onPick={(it) => act('applyVariant', it.id)}
              onMore={(e) => { const r = e.currentTarget.closest('.sl-rs').getBoundingClientRect(); act('designGallery', { kind: 'variants', anchor: { left: r.left, bottom: r.bottom + 4 } }); }}
            />
            <Rows>
              <Button icon="contrast" label={t('Colours')} title={t("Colours — the theme's twelve colours: a palette, or your own")} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind: 'colours', anchor: { left: r.left, bottom: r.bottom + 4 } }); }} />
              <Button icon="textbox" label={t('Fonts')} title={t("Fonts — the theme's heading and body faces: a pair, or your own")} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind: 'fonts', anchor: { left: r.left, bottom: r.bottom + 4 } }); }} />
              <Button icon="wand" label={t('Effects')} title={t('Effects — how theme-styled shapes are filled, outlined and lifted')} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind: 'effects', anchor: { left: r.left, bottom: r.bottom + 4 } }); }} />
            </Rows>
          </Group>
          <Group label={t('Customise')}>
            <Button
              tall
              icon="picture"
              label={t('Background Styles')}
              title={t('Background Styles — a solid or gradient background for this slide, or every slide')}
              onClick={backgroundMenu}
            />
          </Group>
          <Group label={t('Slides')}>
            <Button tall icon="grid" label={t('Layouts')} pressed={view.pane === 'designs'} title={t("The deck's layouts in a pane — put this slide on one, or start a new slide from it")} onClick={() => act('pane', 'designs')} />
            <Button tall icon="slides" label={t('Layout')} title={t('A new slide with this layout, after the current one')} onClick={(e) => menu.open(e, LAYOUTS.map(([layout, label]) => ({ label, icon: 'slides', run: () => addSlide(layout) })))} />
            <Button
              tall
              icon="grid"
              label={t('Slide Size')}
              title={model?.size ? t('Slides are {width} by {height} pixels — {ratio}. Changing it rescales every slide; not built yet.', { width: Math.round(model.size.width), height: Math.round(model.size.height), ratio: Math.abs(model.size.width / model.size.height - 16 / 9) < 0.02 ? t('16:9 widescreen') : Math.abs(model.size.width / model.size.height - 4 / 3) < 0.02 ? t('4:3 standard') : t('a custom ratio') }) : t('The deck decides its own size')}
              disabled
            />
          </Group>
          <Group label={t('Edit')}>
            <Button tall icon="wand" label={t('Design Ideas')} pressed={view?.pane === 'ideas'} title={t("Design Ideas — layouts this slide's title, words and pictures suit, worked out on this computer; a click applies one")} onClick={() => act('pane', 'ideas')} />
          </Group>
        </>
      ) : null}

      {/* ── Transitions ──────────────────────────────────────────────────── */}
      {tab === 'transitions' ? (
        <>
          <Group label={t('Preview')}>
            <Button
              tall
              icon="play"
              label={t('Preview')}
              disabled={!transitionPlays}
              title={transitionPlays ? t("Preview — play this slide's transition on the stage ({transition})", { transition: describeTransition(transition) }) : t('Preview — this slide has no transition to play; pick one from the gallery')}
              onClick={() => act('preview', 'transition')}
            />
          </Group>
          <Group label={t('Transition to This Slide')}>
            {TRANSITION_GALLERY.map(([key, label, blurb]) => (
              <Button
                key={key}
                tall
                icon={TRANSITION_ICONS[key]}
                label={label}
                className="sl-tr-pick"
                data-transition={key}
                pressed={transitionKey === key}
                title={t('{label} — {blurb}', { label, blurb })}
                onClick={() => act('transition', key === 'none' ? { type: 'none' } : { type: key === 'shape' ? 'circle' : key })}
              />
            ))}
            <Button
              tall
              icon="settings"
              label={t('Effect Options')}
              disabled={!transitionOptions}
              title={transitionOptions ? t('Effect Options — which way the {name} goes', { name: TRANSITION_GALLERY.find(([k]) => k === transitionKey)?.[1] || t('transition') }) : t('Effect Options — pick a transition with a direction first (Push, Wipe, Split, Cover…)')}
              onClick={(e) => menu.open(e, transitionOptions.map(([value, label]) => ({
                label,
                icon: optionOf(transition) === value ? 'check' : undefined,
                run: () => act('transition', transitionKey === 'shape' ? { type: value.split(':')[0], direction: value.split(':')[1] ?? null } : { direction: value }),
              })))}
            />
          </Group>
          <Group label={t('Timing')}>
            <Rows>
              <Button icon="volume" label={t('Sound')} title={t("Sound — what plays as this slide comes in: one of the suite's sounds, a WAV of your own, or one that stops the last")} onClick={(e) => {
                const s = model?.slide?.transitionSound || null;
                const tick = (on) => (on ? 'check' : undefined);
                menu.open(e, [
                  { label: t('[No Sound]'), icon: tick(!s), run: () => act('transitionSound', { none: true }) },
                  { label: t('[Stop Previous Sound]'), icon: tick(Boolean(s?.stop)), run: () => act('transitionSound', { stop: true }) },
                  '-',
                  ...SOUNDS.map((name) => ({ label: name, icon: tick(s?.name === soundFile(name)) || 'volume', run: () => act('transitionSound', { builtin: name }) })),
                  '-',
                  { label: t('Other Sound…'), icon: 'folder', run: () => act('transitionSound', { other: true }) },
                  '-',
                  { label: t('Loop Until Next Sound'), icon: tick(Boolean(s?.loop)), disabled: !s || s.stop, run: () => act('transitionSound', { loop: !s?.loop }) },
                ]);
              }} />
              <div className="sl-rb-field" data-tip={t('Duration — how long the transition takes, in seconds')}>
                <Icon name="clock" size={15} />
                <span>{t('Duration')}</span>
                <SecondsField className="sl-tr-duration" value={transition && transitionKey !== 'none' ? transition.duration : null} disabled={!transition || transitionKey === 'none'} min={0.01} max={59} onCommit={(v) => act('transition', { duration: v })} />
              </div>
              <Button icon="check" label={t('Apply To All')} disabled={count < 2} title={count < 2 ? t('Apply To All — the deck has one slide') : t("Apply To All — this slide's transition, timing and advance settings on every slide")} onClick={() => act('transitionAll')} />
            </Rows>
            <Separator />
            <Rows>
              <div className="sl-rb-caption">{t('Advance Slide')}</div>
              <label className="sl-rb-field" data-tip={t('On Mouse Click — a click moves the show on from this slide')}>
                <input type="checkbox" className="sl-tr-onclick" checked={transition ? transition.advanceOnClick !== false : true} onChange={(e) => act('transition', { advanceOnClick: e.target.checked })} />
                <span>{t('On Mouse Click')}</span>
              </label>
              <div className="sl-rb-field" data-tip={t('After — the show moves on by itself this many seconds after the slide is shown')}>
                <label className="sl-rb-check">
                  <input type="checkbox" className="sl-tr-after-on" checked={transition?.advanceAfter != null} onChange={(e) => act('transition', { advanceAfter: e.target.checked ? (transition?.advanceAfter ?? 0) : null })} />
                  <span>{t('After:')}</span>
                </label>
                <SecondsField className="sl-tr-after" value={transition?.advanceAfter ?? 0} min={0} max={3600} onCommit={(v) => act('transition', { advanceAfter: v })} />
              </div>
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Animations ───────────────────────────────────────────────────── */}
      {tab === 'animations' ? (
        <>
          <Group label={t('Preview')}>
            <Button
              tall
              icon="play"
              label={t('Preview')}
              disabled={!animationList.length}
              title={animationList.length ? tn(animationList.length, "Preview — play this slide's animation on the stage", "Preview — play this slide's {count} animations on the stage") : t('Preview — this slide has no animations yet; select a shape and pick an effect')}
              onClick={() => act('preview', 'animation')}
            />
          </Group>
          <Group label={t('Animation')}>
            <Button
              tall
              icon="star"
              label={t('None')}
              className="sl-an-pick sl-an-none"
              data-effect="none"
              pressed={hasShape && !shapeAnimated}
              disabled={!hasShape}
              title={hasShape ? t('None — take every animation off the selected shape') : t('None — select a shape first; its animations come off')}
              onClick={() => act('animate', { kind: 'entr', effect: 'none' })}
            />
            {ANIMATION_GALLERY.map(([effect, label, blurb]) => (
              <Button
                key={effect}
                tall
                icon="star"
                label={label}
                className="sl-an-pick sl-an-entr"
                data-effect={effect}
                pressed={animation?.kind === 'entr' && animation?.effect === effect}
                disabled={!hasShape}
                title={hasShape ? (animation ? t('{label} — {blurb}; it replaces the selected effect', { label, blurb }) : t('{label} — {blurb}', { label, blurb })) : t('{label} — select a shape first, then pick its entrance', { label })}
                onClick={() => act('animate', { kind: 'entr', effect })}
              />
            ))}
            <Button
              tall
              icon="redo"
              label={t('Motion Paths')}
              className="sl-an-paths"
              pressed={animation?.kind === 'path'}
              disabled={!hasShape}
              title={hasShape ? t('Motion Paths — the shape moves along a line, an arc, a turn, a circle, a square or a loop from where it stands') : t('Motion Paths — select a shape first')}
              onClick={(e) => menu.open(e, EFFECT_MENU.find(([kind]) => kind === 'path')[2].map(([effect, label]) => ({ label, icon: 'redo', run: () => act('animate', { kind: 'path', effect }) })))}
            />
            <Button
              tall
              icon="more"
              label={t('More Effects')}
              disabled={!hasShape}
              title={hasShape ? t('More Effects — every entrance, emphasis and exit effect') : t('More Effects — select a shape first')}
              onClick={(e) => menu.open(e, effectMenu('animate'))}
            />
            <Button
              tall
              icon="settings"
              label={t('Effect Options')}
              disabled={!animationOptions && !sequenceItems.length}
              title={animationOptions || sequenceItems.length ? (animationOptions && sequenceItems.length ? t('Effect Options — which way the {name} goes, and the words as one object or paragraph by paragraph', { name: animation?.name || t('effect') }) : animationOptions ? t('Effect Options — which way the {name} goes', { name: animation?.name || t('effect') }) : t('Effect Options — the words as one object or paragraph by paragraph')) : t('Effect Options — pick an effect with a direction (Fly In, Float In, Split, Wipe, Spin), or one on words of more than one paragraph')}
              onClick={(e) => menu.open(e, [
                ...(animationOptions || []).map(([value, label]) => ({ label, icon: animation?.direction === value ? 'check' : undefined, run: () => act('animPatch', { direction: value }) })),
                ...(animationOptions && sequenceItems.length ? ['-'] : []),
                ...sequenceItems,
              ])}
            />
          </Group>
          <Group label={t('Advanced Animation')}>
            <Button tall icon="plus" label={t('Add Animation')} disabled={!hasShape} title={hasShape ? t('Add Animation — another effect on the selected shape, after its others') : t('Add Animation — select a shape first')} onClick={(e) => menu.open(e, effectMenu('addAnimation'))} />
            <Rows>
              <Button icon="list" label={t('Animation Pane')} pressed={view.pane === 'animations'} title={t("Animation Pane — the slide's effects in order, to pick, reorder, retime or remove")} onClick={() => act('pane', 'animations')} />
              <Button
                icon="play"
                label={t('Trigger')}
                disabled={!animation}
                title={animation ? (animation.triggerShape ? t('Trigger — start this effect on a click on a shape in the show, not in the slide\'s own order') : t('Trigger — start this effect on a click on a shape in the show')) : t('Trigger — pick an effect in the Animation Pane first')}
                onClick={(e) => menu.open(e, [
                  { label: t('On Click Sequence'), icon: !animation?.triggerShape ? 'check' : undefined, run: () => act('animPatch', { triggerShape: null }) },
                  '-',
                  ...(model?.slide?.shapes || []).filter((s) => s.groupId == null).map((s) => ({
                    label: t('On Click of {name}', { name: s.name || t('Shape {id}', { id: s.id }) }),
                    icon: String(animation?.triggerShape) === String(s.id) ? 'check' : undefined,
                    run: () => act('animPatch', { triggerShape: s.id }),
                  })),
                ])}
              />
              <Button
                icon="wand"
                label={t('Animation Painter')}
                pressed={animPainter}
                disabled={!animPainter && !shapeAnimated}
                title={animPainter ? t('Animation Painter — armed: click a shape to give it these effects; Esc puts it down') : shapeAnimated ? t("Animation Painter — this shape's effects onto the next shape you click") : t('Animation Painter — select an animated shape first')}
                onClick={() => act('animPainter')}
              />
            </Rows>
          </Group>
          <Group label={t('Timing')}>
            <Rows>
              <div className="sl-rb-field" data-tip={t('Start — when the effect plays: on a click, with the one before it, or after it')}>
                <Icon name="play" size={14} />
                <span className="sl-rb-label">{t('Start:')}</span>
                <Select className="rw-select sl-an-start" value={animation?.trigger || 'onClick'} disabled={!animation} style={{ width: 128, height: 24 }} onChange={(e) => act('animPatch', { trigger: e.target.value })}>
                  <option value="onClick">{t('On Click')}</option>
                  <option value="withPrevious">{t('With Previous')}</option>
                  <option value="afterPrevious">{t('After Previous')}</option>
                </Select>
              </div>
              <div className="sl-rb-field" data-tip={t('Duration — how long the effect takes, in seconds')}>
                <Icon name="clock" size={14} />
                <span className="sl-rb-label">{t('Duration:')}</span>
                <SecondsField className="sl-an-duration" value={animation ? animation.duration : null} disabled={!animation || !animation.duration} min={0.01} max={59} onCommit={(v) => act('animPatch', { duration: v })} />
              </div>
              <div className="sl-rb-field" data-tip={t('Delay — how long after its start the effect waits, in seconds')}>
                <Icon name="clock" size={14} />
                <span className="sl-rb-label">{t('Delay:')}</span>
                <SecondsField className="sl-an-delay" value={animation ? animation.delay : null} disabled={!animation} min={0} max={3600} onCommit={(v) => act('animPatch', { delay: v })} />
              </div>
            </Rows>
            <Separator />
            <Rows>
              <div className="sl-rb-caption">{t('Reorder Animation')}</div>
              <Button icon="chevronUp" label={t('Move Earlier')} disabled={!animation || animation.index <= 0} title={animation ? t('Move Earlier — this effect plays before the one above it') : t('Move Earlier — pick an effect first')} onClick={() => act('animMove', 'earlier')} />
              <Button icon="chevronDown" label={t('Move Later')} disabled={!animation || animation.index >= animationList.length - 1} title={animation ? t('Move Later — this effect plays after the one below it') : t('Move Later — pick an effect first')} onClick={() => act('animMove', 'later')} />
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Slide Show ───────────────────────────────────────────────────── */}
      {tab === 'show' ? (
        <>
          <Group label={t('Start Slide Show')}>
            <Button tall icon="play" label={t('From Beginning')} title={t('Start the show at the first slide (F5)')} onClick={() => act('present', 'start')} />
            <Button tall icon="play" label={t('From Current Slide')} title={t('Start the show at this slide')} onClick={() => act('present', 'here')} />
            <Button tall icon="list" label={t('Custom Slide Show')} title={t("Custom Slide Show — play one of the deck's named shows, or make one: some of its slides, in an order of their own")} onClick={(e) => act('customShowMenu', e)} />
          </Group>
          <Group label={t('Rehearse')}>
            <Soon tall icon="volume" label={t('Rehearse with Coach')} why={t('Presenter Coach is a Microsoft cloud service.')} />
          </Group>
          <Group label={t('Set Up')}>
            <Button tall icon="settings" label={t('Set Up Slide Show')} title={t('Set Up Slide Show — who the show is for, looping, which slides and how it moves on')} onClick={() => act('setupShow')} />
            <Button
              tall
              icon="eye"
              label={t('Hide Slide')}
              pressed={Boolean(model?.slide?.hidden)}
              title={model?.slide?.hidden ? t('Hide Slide — this slide is hidden from the show; press again to show it') : t('Hide Slide — leave this slide out of the show')}
              onClick={() => act('hideSlide')}
            />
            <Button tall icon="clock" label={t('Rehearse Timings')} title={t('Rehearse Timings — run the show from the start with a clock, and keep how long each slide was on screen as its timing')} onClick={() => act('rehearse')} />
            <Button tall icon="video" label={t('Record')} title={t("Record — the show with the microphone on: each slide's narration and how long it was up")} onClick={(e) => menu.open(e, [{ label: t('From Current Slide…'), icon: 'play', run: () => act('recordShow', 'here') }, { label: t('From Beginning…'), icon: 'play', run: () => act('recordShow', 'start') }])} />
            <Soon icon="check" label={t('Keep Slides Updated')} why={t('Live co-authoring is not built.')} />
            <Button
              icon="check"
              label={t('Play Narrations')}
              pressed={model?.showSettings?.narration !== false}
              title={t('Play Narrations — whether the show plays the narration recorded with it; kept in the file for PowerPoint, as this suite records none')}
              onClick={() => act('showFlag', { narration: model?.showSettings?.narration === false })}
            />
            <Button
              icon="check"
              label={t('Use Timings')}
              pressed={model?.showSettings?.useTimings !== false}
              title={t('Use Timings — slides with a timing move on by themselves; off, the show moves only when you move it')}
              onClick={() => act('showFlag', { useTimings: model?.showSettings?.useTimings === false })}
            />
            <Button
              icon="check"
              label={t('Show Media Controls')}
              pressed={model?.showSettings?.mediaControls !== false}
              title={t('Show Media Controls — a video in the show gets its play bar once it has started')}
              onClick={() => act('showFlag', { mediaControls: model?.showSettings?.mediaControls === false })}
            />
          </Group>
          <Group label={t('Monitors')}>
            <Button icon="grid" label={t('Monitor: {name}', { name: monitorName })} title={t("Monitor — the screen the slide show plays on; Automatic puts it on another screen than Presenter View's")} onClick={(e) => act('monitorMenu', e)} />
            <Button icon="check" label={t('Use Presenter View')} title={t('Opens a second window with your notes, the next slide and a clock — put it on the other screen')} onClick={presentWithNotes} />
          </Group>
          <Group label={t('Captions & Subtitles')}>
            <Soon icon="check" label={t('Always Use Subtitles')} why={t('Live captions are an online speech service this suite does not call.')} />
            <Soon icon="settings" label={t('Subtitle Settings')} why={t('Comes with captions.')} />
          </Group>
        </>
      ) : null}

      {/* ── Record ───────────────────────────────────────────────────────── */}
      {tab === 'record' ? (
        <>
          <Group label={t('Preview')}>
            <Button tall icon="play" label={t('Preview')} title={t("Preview — this slide's narration, heard")} onClick={() => act('previewNarration')} />
          </Group>
          <Group label={t('Camera')}>
            <Button tall icon="video" label={t('Cameo')} title={t('Cameo — your camera, live, in a shape on this slide')} onClick={() => act('cameo')} />
          </Group>
          <Group label={t('Record')}>
            <Button tall icon="video" label={t('From Beginning')} title={t('Record From Beginning — the show from its first slide with the microphone on')} onClick={() => act('recordShow', 'start')} />
            <Button tall icon="video" label={t('From Current Slide')} title={t('Record From Current Slide — the show from this slide with the microphone on')} onClick={() => act('recordShow', 'here')} />
            <Button tall icon="video" label={t('Screen Recording')} title={t('Screen Recording — record a window or a screen until Stop, and put the recording on this slide as a video')} onClick={() => act('screenRecording')} />
            <Button tall icon="volume" label={t('Audio')} title={t('Record Audio — record a sound and put it on this slide')} onClick={() => act('recordAudio')} />
          </Group>
          <Group label={t('Edit')}>
            <Button tall icon="close" label={t('Clear Recording')} title={t('Clear — narration or timings, on this slide or every slide')} onClick={(e) => menu.open(e, [{ label: t('Clear Narration on Current Slide'), run: () => act('clearRecording', 'narrationHere') }, { label: t('Clear Narration on All Slides'), run: () => act('clearRecording', 'narrationAll') }, '-', { label: t('Clear Timings on Current Slide'), run: () => act('clearRecording', 'timingsHere') }, { label: t('Clear Timings on All Slides'), run: () => act('clearRecording', 'timingsAll') }])} />
            <Button tall icon="undo" label={t('Reset to Cameo')} className="sl-reset-cameo" title={t("Reset to Cameo — the camera's recording taken off this slide, its live cameo back")} onClick={() => act('resetCameo')} />
          </Group>
          <Group label={t('Export')}>
            <Button tall icon="export" label={t('Save as Show')} title={t('Save as Show — a copy as a .ppsx, which PowerPoint opens straight into the show')} onClick={() => act('saveAsShow')} />
            <Button tall icon="video" label={t('Export to Video')} title={t('Export to Video — the deck played into a video file, with its timings and narrations')} onClick={() => act('exportVideo')} />
          </Group>
          <Group label={t('Help')}>
            <Button tall icon="info" label={t('Learn More')} onClick={() => act('help')} />
          </Group>
        </>
      ) : null}

      {/* ── Review ───────────────────────────────────────────────────────── */}
      {tab === 'review' ? (
        <>
          <Group label={t('Proofing')}>
            <Button tall icon="check" label={t('Spelling')} pressed={review?.pane === 'editor'} title={t("Spelling (F7) — check every slide's words and notes, from this slide")} onClick={() => review?.startSpelling()} />
            <Button tall icon="find" label={t('Thesaurus')} pressed={review?.pane === 'thesaurus'} title={t("Thesaurus (Shift+F7) — words of like meaning for the word where the caret was in a shape's text")} onClick={() => review?.openThesaurus()} />
          </Group>
          <Group label={t('Accessibility')}>
            <Button tall icon="shield" label={t('Check Accessibility')} pressed={review?.pane === 'accessibility'} title={t('Check Accessibility — alt text, slide titles, reading order, contrast and table headers, with a fix for each')} onClick={() => review?.openAccessibility()} />
            <Button icon="textbox" label={t('Alt Text')} disabled={selected == null} title={selected == null ? t('Alt Text — select a picture or shape first, then describe it') : t('Alt Text — describe the selected object for people who cannot see it')} onClick={() => selected != null && review?.openAltText({ slide: index, shape: selected })} />
          </Group>
          <Group label={t('Language')}>
            <Soon tall icon="globe" label={t('Translate')} why={t('Translation is an online service this suite does not call.')} />
            <Button tall icon="globe" label={t('Language')} title={t("Language — mark the selected text box's words as a language, or as not to be checked, for Spelling here and proofing in PowerPoint")} onClick={() => act('language')} />
          </Group>
          <Group label={t('Activity')}>
            <Soon icon="check" label={t('Mark All as Read')} why={t('A deck keeps its comments and whether each is resolved, but not who has read them, so there is nothing in the file to mark.')} />
            <Button icon="eye" label={t('Show Changes')} className="sl-show-changes" pressed={view.pane === 'changes'} title={view.changes ? tn(view.changes, 'Show Changes — {count} slide different since this deck was last open here', 'Show Changes — {count} slides different since this deck was last open here') : t('Show Changes — what is different about this deck since it was last open on this computer')} onClick={() => act('pane', 'changes')} />
          </Group>
          <Group label={t('Comments')}>
            <Button tall icon="reply" label={t('New Comment')} title={t('New Comment — on the selected shape, or on this slide; it is signed with your name')} onClick={() => act('newComment')} />
            <Rows>
              <Button icon="trash" label={t('Delete')} disabled={!commentCount} title={commentCount ? t('Delete — this comment, every comment on this slide, or every one in the presentation') : t('Delete — there are no comments in this presentation')} onClick={(e) => menu.open(e, [
                { label: t('Delete Comment'), icon: 'trash', run: () => act('deleteComment') },
                { label: t('Delete All Comments on This Slide'), icon: 'trash', run: () => act('deleteComments', 'slide') },
                { label: t('Delete All Comments in This Presentation'), icon: 'trash', run: () => act('deleteComments', 'all') },
              ])} />
              <Button icon="chevronLeft" label={t('Previous')} disabled={!commentCount} title={commentCount ? t('Previous — the comment before this one, across the slides') : t('Previous — there are no comments in this presentation')} onClick={() => act('commentStep', -1)} />
              <Button icon="chevronRight" label={t('Next')} disabled={!commentCount} title={commentCount ? t('Next — the comment after this one, across the slides') : t('Next — there are no comments in this presentation')} onClick={() => act('commentStep', 1)} />
            </Rows>
            <Button tall icon="eye" label={t('Show Comments')} pressed={view.pane === 'comments'} title={t('Show Comments — the Comments pane beside the slide')} onClick={() => act('pane', 'comments')} />
          </Group>
          <Group label={t('Notes')}>
            <Button tall icon="word" label={t('Speaker Notes')} title={t("Speaker Notes — this slide's, shown in Presenter View")} onClick={() => setNotesOpen(true)} />
          </Group>
          <Group label={t('Ink')}>
            <Button icon="eye" label={t('Hide Ink')} pressed={Boolean(ink?.hide)} title={t("Hide Ink — the slide's strokes out of sight while you work; they stay in the deck")} onClick={() => act('hideInk')} />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
      {tab === 'view' ? (
        <>
          <Group label={t('Presentation Views')}>
            <Button tall icon="slides" label={t('Normal')} pressed={(view.mode || 'normal') === 'normal'} onClick={() => act('mode', 'normal')} />
            <Button tall icon="list" label={t('Outline View')} pressed={view.mode === 'outline'} title={t("Every slide's words, as an outline")} onClick={() => act('mode', 'outline')} />
            <Button tall icon="grid" label={t('Slide Sorter')} pressed={view.mode === 'sorter'} title={t('All the slides at once — click one to open it')} onClick={() => act('mode', 'sorter')} />
            <Button tall icon="word" label={t('Notes Page')} pressed={view.mode === 'notes'} title={t('The slide with its notes under it')} onClick={() => act('mode', 'notes')} />
            <Button tall icon="play" label={t('Reading View')} title={t('The show in this window, not full screen')} onClick={() => act('present', 'reading')} />
          </Group>
          <Group label={t('Master Views')}>
            <Button tall icon="slides" label={t('Slide Master')} pressed={masterView?.kind === 'slide'} title={t('Slide Master — edit the master and its layouts; every slide on them follows')} onClick={() => act(masterView?.kind === 'slide' ? 'closeMaster' : 'masterView')} />
            <Button tall icon="file" label={t('Handout Master')} pressed={masterView?.kind === 'handout'} title={t("Handout Master — the printed handout's header, date, footer and page number, and how many slides a page shows")} onClick={() => act(masterView?.kind === 'handout' ? 'closeMaster' : 'handoutMasterView')} />
            <Button tall icon="word" label={t('Notes Master')} pressed={masterView?.kind === 'notes'} title={t("Notes Master — the printed notes page: where the slide and the notes go, the notes' text styles, the header and footer")} onClick={() => act(masterView?.kind === 'notes' ? 'closeMaster' : 'notesMasterView')} />
          </Group>
          <Group label={t('Show')}>
            <Button icon="minus" label={t('Ruler')} pressed={Boolean(view.ruler)} onClick={() => act('toggle', 'ruler')} />
            <Button icon="grid" label={t('Gridlines')} pressed={Boolean(view.gridlines)} onClick={() => act('toggle', 'gridlines')} />
            <Button icon="plus" label={t('Guides')} pressed={Boolean(view.guides)} onClick={() => act('toggle', 'guides')} />
            <Button icon="word" label={t('Notes')} pressed={view.notes !== false} title={t('The notes strip under the slide')} onClick={() => act('toggle', 'notes')} />
            <Button icon="list" label={t('Layers')} pressed={view.pane === 'layers'} title={t("The slide's shapes as layers — select, reorder, hide, rename")} onClick={() => act('pane', 'layers')} />
            <Button icon="grid" label={t('Designs')} pressed={view.pane === 'designs'} title={t("The deck's layouts — put this slide on one, or start a new slide from it")} onClick={() => act('pane', 'designs')} />
          </Group>
          <Group label={t('Direction')}>
            <Button icon="rotate" label={t('View Direction')} pressed={Boolean(view.rtl)} title={view.rtl ? t("View Direction — right to left: the slides on the right, the strip and the panes mirrored; click for left to right") : t("View Direction — left to right; click for right to left, the slides strip on the right as a right-to-left language reads")} onClick={() => act('toggle', 'rtl')} />
          </Group>
          <Group label={t('Zoom')}>
            <Button tall icon="zoomIn" label={t('Zoom')} onClick={(e) => menu.open(e, [50, 75, 100, 150, 200].map((z) => ({ label: `${z}%`, run: () => act('zoom', z / 100) })))} />
            <Button tall icon="maximize" label={t('Fit to Window')} pressed={!view.zoom} onClick={() => act('zoom', null)} />
          </Group>
          <Group label={t('Colour/Greyscale')}>
            <Button icon="contrast" label={t('Colour')} pressed={!view.tone || view.tone === 'colour'} onClick={() => act('tone', 'colour')} />
            <Button icon="contrast" label={t('Greyscale')} pressed={view.tone === 'grey'} onClick={() => act('tone', 'grey')} />
            <Button icon="contrast" label={t('Black and White')} pressed={view.tone === 'mono'} onClick={() => act('tone', 'mono')} />
          </Group>
          <Group label={t('Window')}>
            <Button tall icon="new" label={t('New Window')} title={t('This deck in a second window')} onClick={() => act('newWindow')} />
            <Button icon="grid" label={t('Arrange All')} title={t('Arrange All — every Presentations window, side by side')} onClick={() => act('arrange', 'columns')} />
            <Button icon="grid" label={t('Cascade')} title={t('Cascade — every Presentations window, each a step down and across from the last')} onClick={() => act('arrange', 'cascade')} />
            <Button icon="minus" label={t('Move Split')} title={t("Move Split — the arrow keys move the slide pane's edge and the top of the notes; Enter or Esc when done")} onClick={() => act('moveSplit')} />
            <Button icon="list" label={t('Switch Windows')} title={t('Switch Windows — the Presentations windows open now')} onClick={(e) => act('switchWindows', e)} />
            <Button icon="maximize" label={t('Full Screen')} onClick={() => shell.win.fullscreen({})} />
          </Group>
          <Group label={t('Macros')}>
            <Soon tall icon="settings" label={t('Macros')} why={t('VBA is preserved in the file and never run: a deck that runs code it received in an email is how ransomware starts.')} />
          </Group>
        </>
      ) : null}

      {/* ── Help ─────────────────────────────────────────────────────────── */}
      {tab === 'help' ? (
        <>
          <Group label={t('Help & Support')}>
            <Button tall icon="info" label={t('Help')} title={t('The Presentation guide on office.rutba.io')} onClick={() => act('help')} />
            <Button tall icon="send" label={t('Feedback')} title={t('Tell us what is wrong or missing')} onClick={() => act('feedback')} />
            <Button tall icon="info" label={t('Show Training')} onClick={() => act('help')} />
            <Button tall icon="star" label={t("What's New")} title={t('The release notes')} onClick={() => act('releases')} />
            <Button icon="list" label={t('Keyboard Shortcuts')} onClick={() => act('shortcuts')} />
          </Group>
          <Group label={t('File')}>
            <Button tall icon="new" label={t('New')} onClick={() => shell.win.create({ app: 'slides' })} />
            <Button tall icon="open" label={t('Open')} onClick={openFile} />
            <Button tall icon="save" label={t('Save')} onClick={() => save(false)} />
          </Group>
        </>
      ) : null}

      {/* ── PDF ──────────────────────────────────────────────────────────── */}
      {tab === 'pdf' ? (
        <>
          <Group label={t('Export')}>
            <Button tall icon="pdf" label={t('Export as PDF')} title={t('Every slide as a page, drawn exactly as the stage draws it')} onClick={() => exportAs('pdf')} />
          </Group>
        </>
      ) : null}

      {/* ── 3D Model (contextual) ─────────────────────────────────────────── */}
      {tab === 'model3d' && selectedShapeObj?.model3d ? (() => {
        const v = selectedShapeObj.model3d.view || {};
        const at = (view) => Math.abs((v.yaw || 0) - view.yaw) < 0.5 && Math.abs((v.pitch || 0) - view.pitch) < 0.5;
        const rows = [MODEL_VIEWS.slice(0, 4), MODEL_VIEWS.slice(4, 7), MODEL_VIEWS.slice(7)];
        return (
          <>
            <Group label={t('3D Model Views')}>
              {rows.map((row, i) => (
                <Rows key={i}>
                  {row.map(([key, label, view]) => <Button key={key} icon={at(view) ? 'check' : 'shape'} label={t(label)} className="sl-model3d-view" data-view={key} pressed={at(view)} title={t('{label} — the model turned to show it from there', { label: t(label) })} onClick={() => act('model3dView', view)} />)}
                </Rows>
              ))}
            </Group>
            <Group label={t('Turn')}>
              <Rows>
                <Button icon="chevronLeft" label={t('Turn Left')} className="sl-model3d-turn-left" title={t('Turn Left — fifteen degrees about its upright')} onClick={() => act('model3dView', { turn: { yaw: -15, pitch: 0 } })} />
                <Button icon="chevronRight" label={t('Turn Right')} title={t('Turn Right — fifteen degrees about its upright')} onClick={() => act('model3dView', { turn: { yaw: 15, pitch: 0 } })} />
              </Rows>
              <Rows>
                <Button icon="chevronUp" label={t('Tip Back')} title={t('Tip Back — fifteen degrees, the top towards you')} onClick={() => act('model3dView', { turn: { yaw: 0, pitch: 15 } })} />
                <Button icon="chevronDown" label={t('Tip Forward')} title={t('Tip Forward — fifteen degrees, the bottom towards you')} onClick={() => act('model3dView', { turn: { yaw: 0, pitch: -15 } })} />
              </Rows>
            </Group>
            <Group label={t('Adjust')}>
              <Button tall icon="undo" label={t('Reset 3D Model')} className="sl-model3d-reset" title={t('Reset 3D Model — back to the view it was put in at')} onClick={() => act('model3dView', 'reset')} />
            </Group>
          </>
        );
      })() : null}

      {/* ── Camera Format (contextual) ─────────────────────────────────────── */}
      {/* ── Shape Format (contextual): WordArt's Transform ──────────────── */}
      {tab === 'shapeFormat' && selectedShapeObj?.kind === 'shape' && selectedShapeObj?.text ? (
        <>
          <Group label={t('Text Effects: Transform')}>
            {WARP_PRESETS.map((p) => (
              <Button key={p.id} tall icon="wand" label={t(p.label)} className="sl-warp" data-preset={p.id} pressed={(selectedShapeObj.text.warp?.preset || 'textNoShape') === p.id} title={p.id === 'textNoShape' ? t('No Transform — the words in straight lines') : p.label === 'Button' ? t('Transform — the words along a button: an arc, a line and an arc') : p.label === 'Circle' ? t('Transform — the words along the circle of the shape') : t('Transform — the words along the arc of the shape')} onClick={() => act('textWarp', p.id)} />
            ))}
            {/* The warps, the rest of Office's gallery, from a menu. */}
            <Button tall icon="wand" label={t('More')} className="sl-warp-more" pressed={WARP_MORE.some((p) => p.id === selectedShapeObj.text.warp?.preset)} title={WARP_MORE.some((p) => p.id === selectedShapeObj.text.warp?.preset) ? t('Transform — now {name}; the warps: the words stretched between two curves', { name: t(warpLabel(selectedShapeObj.text.warp?.preset)) }) : t('More Transforms — the warps: the words stretched between two curves, a wave, a slant, a chevron and the rest')} onClick={(e) => menu.open(e, WARP_MORE.map((p) => ({ label: t(p.label), icon: selectedShapeObj.text.warp?.preset === p.id ? 'check' : 'wand', run: () => act('textWarp', p.id) })))} />
          </Group>
        </>
      ) : null}

      {tab === 'cameraFormat' && selectedShapeObj?.cameo ? (
        <>
          <Group label={t('Camera')}>
            <Button tall icon="eye" label={t('Preview')} className="sl-cameo-preview" pressed={Boolean(view.cameoPreview)} title={view.cameoPreview ? t('Preview — the camera is on here; press to turn it off') : t('Preview — the camera, live, in the cameo here on the slide')} onClick={() => act('cameoPreview')} />
          </Group>
          <Group label={t('Camera Styles')}>
            {[['rect', t('Rectangle')], ['ellipse', t('Oval')], ['roundRect', t('Rounded Rectangle')]].map(([preset, label]) => (
              <Button key={preset} tall icon="shape" label={label} className="sl-cameo-shape" data-preset={preset} pressed={selectedShapeObj.preset === preset} title={t('Camera Shape — the cameo as a {shape}', { shape: label.toLowerCase() })} onClick={() => act('cameoShape', preset)} />
            ))}
          </Group>
        </>
      ) : null}

      {/* ── Table Design (contextual) ─────────────────────────────────────── */}
      {tab === 'tableDesign' && selectedShapeObj?.kind === 'table' ? (() => {
        const flags = selectedShapeObj.table?.flags || {};
        const toggle = (name, label, why) => <Button icon={flags[name] ? 'check' : undefined} label={label} pressed={Boolean(flags[name])} title={why} onClick={() => act('tableLook', { [name]: !flags[name] })} />;
        return (
          <>
            <Group label={t('Table Style Options')}>
              <Rows>
                {toggle('firstRow', t('Header Row'), t('Header Row — the first row in its style\'s header look'))}
                {toggle('lastRow', t('Total Row'), t('Total Row — the last row in its style\'s total look'))}
                {toggle('bandRow', t('Banded Rows'), t('Banded Rows — every other row shaded, as the style bands them'))}
              </Rows>
              <Rows>
                {toggle('firstCol', t('First Column'), t('First Column — the first column in its style\'s look'))}
                {toggle('lastCol', t('Last Column'), t('Last Column — the last column in its style\'s look'))}
                {toggle('bandCol', t('Banded Columns'), t('Banded Columns — every other column shaded, as the style bands them'))}
              </Rows>
            </Group>
            <Group label={t('Shading')}>
              <Button tall icon="wand" label={t('Shading')} title={t('Shading — the cell last typed in, or every cell of the table')} onClick={(e) => menu.open(e, [
                { label: t('No Fill'), icon: 'close', run: () => act('tableShading', { fill: null }) },
                '-',
                ...COLOURS.map(([hex, label]) => ({ label, run: () => act('tableShading', { fill: hex }) })),
              ])} />
            </Group>
          </>
        );
      })() : null}
    </Ribbon>
  );
}
