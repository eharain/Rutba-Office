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
import { Ribbon, Group, Rows, Button, Separator, Select, Icon } from '@rutba/office-ui';
import { TRANSITION_GALLERY, TRANSITION_OPTIONS, galleryKeyOf, optionOf, describeTransition } from './motion.js';
import { ANIMATION_GALLERY, EFFECT_MENU, ANIMATION_OPTIONS } from './animate.js';
import { RibbonStrip } from './design.js';

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
const COLOURS = [
  ['#000000', 'Black'], ['#404040', 'Dark grey'], ['#808080', 'Grey'], ['#ffffff', 'White'],
  ['#c00000', 'Dark red'], ['#ff0000', 'Red'], ['#ffc000', 'Orange'], ['#ffff00', 'Yellow'],
  ['#92d050', 'Light green'], ['#00b050', 'Green'], ['#00b0f0', 'Light blue'], ['#0070c0', 'Blue'],
  ['#002060', 'Dark blue'], ['#7030a0', 'Purple'],
];
const LAYOUTS = [
  ['title', 'Title Slide'], ['obj', 'Title and Content'], ['blank', 'Blank'],
];
/** The preset geometries the engine writes and the renderer draws, in PowerPoint's names. */
const SHAPES = [
  ['rect', 'Rectangle'], ['roundRect', 'Rectangle: Rounded Corners'], ['ellipse', 'Oval'],
  ['triangle', 'Isosceles Triangle'], ['rtTriangle', 'Right Triangle'], ['diamond', 'Diamond'],
  ['parallelogram', 'Parallelogram'], ['trapezoid', 'Trapezoid'], ['pentagon', 'Pentagon'],
  ['hexagon', 'Hexagon'], ['octagon', 'Octagon'], ['star5', 'Star: 5 Points'],
  ['rightArrow', 'Arrow: Right'], ['chevron', 'Chevron'], ['line', 'Line'],
];
/** Insert → Table: the sizes PowerPoint's own gallery offers first. */
const TABLE_SIZES = [
  [2, 2, '2 × 2'], [3, 3, '3 × 3'], [4, 3, '4 × 3'], [5, 4, '5 × 4'],
];
/** Insert → Chart: the kinds `chartPartXml` writes — the ones a chart on a slide can be. */
const CHART_TYPES = [
  ['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['pie', 'Pie'],
];

/** Design → Background Styles: the theme backgrounds PowerPoint's gallery offers first, then a few flat colours. */
const BACKGROUND_STYLES = [
  ['Background 1', { scheme: 'bg1' }],
  ['Background 2', { scheme: 'bg2' }],
  ['Text 1', { scheme: 'tx1' }],
  ['Text 2', { scheme: 'tx2' }],
  ['Accent 1 gradient', { gradient: { from: { scheme: 'accent1' }, to: { scheme: 'accent1', lumMod: 75 } } }],
  '-',
  ['White', { colour: 'FFFFFF' }],
  ['Black', { colour: '000000' }],
  ['Light grey', { colour: 'F2F2F2' }],
  ['Dark blue', { colour: '1F3864' }],
];

/** A control that is drawn where PowerPoint draws it, and says why it is not live. */
const Soon = ({ icon, label, tall, why }) => (
  <Button tall={tall} icon={icon} label={label} disabled title={`${label} — not built yet. ${why}`} />
);
/** Shape Effects: the shadows PowerPoint offers first, and none. */
const SHADOW_MENU = [['br', 'Shadow: bottom right'], ['b', 'Shadow: below'], ['r', 'Shadow: right'], ['tl', 'Shadow: top left'], ['c', 'Shadow: all round'], ['none', 'No shadow']];
/** Shape Effects → Glow: PowerPoint's own gallery radii. */
const GLOW_MENU = [[5, 'Glow: 5 pt'], [8, 'Glow: 8 pt'], [11, 'Glow: 11 pt'], [18, 'Glow: 18 pt'], [null, 'No glow']];
/** Shape Effects → Soft Edges. */
const SOFTEDGE_MENU = [[1, 'Soft Edges: 1 pt'], [2.5, 'Soft Edges: 2.5 pt'], [5, 'Soft Edges: 5 pt'], [10, 'Soft Edges: 10 pt'], [null, 'No soft edges']];
/** Shape Effects → Reflection: the three gallery presets. */
const REFLECTION_MENU = [['tight', 'Reflection: tight'], ['half', 'Reflection: half'], ['full', 'Reflection: full'], [null, 'No reflection']];
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
  designStrip = null,

}) {
  const count = model?.count || 0;
  const hasShape = Boolean(selected);
  // Align and Distribute's own toggle — "Align Selected Objects" (the
  // default) or "Align to Slide" — is how the ribbon remembers which one is
  // ticked; it is a preference about the next press, not part of the file.
  const [alignTo, setAlignTo] = React.useState('selection');
  const selectedShapeObj = (model?.slide?.shapes || []).find((s) => s.id === selected) || null;
  const isGroup = selectedShapeObj?.kind === 'group';
  const multiCount = selectedIds.length;
  const size = Number(format.size || 18);
  const nearer = (dir) => {
    const bigger = SIZES.filter((s) => (dir > 0 ? s > size : s < size));
    return dir > 0 ? bigger[0] ?? size : bigger[bigger.length - 1] ?? size;
  };
  const fmt = (delta) => act('format', delta);
  const needShape = hasShape ? undefined : 'Select a text box first — click it once';
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
  /** More Effects and Add Animation: every effect under its heading, each a verb on the selected shape. */
  const effectMenu = (verb) => EFFECT_MENU.flatMap(([kind, heading, effects]) => [
    { heading: true, label: heading },
    ...effects.map(([effect, label]) => ({
      label,
      icon: verb === 'animate' && animation?.kind === kind && animation?.effect === effect ? 'check' : 'star',
      run: () => act(verb, { kind, effect }),
    })),
  ]);

  return (
    <Ribbon
      tabs={[
        { id: 'home', label: 'Home' },
        { id: 'insert', label: 'Insert' },
        { id: 'draw', label: 'Draw' },
        { id: 'design', label: 'Design' },
        { id: 'transitions', label: 'Transitions' },
        { id: 'animations', label: 'Animations' },
        { id: 'show', label: 'Slide Show' },
        { id: 'record', label: 'Record' },
        { id: 'review', label: 'Review' },
        { id: 'view', label: 'View' },
        { id: 'help', label: 'Help' },
        { id: 'pdf', label: 'PDF' },
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title="Save (Ctrl+S)" onClick={() => save(false)} />
          <Button icon="undo" title="Undo (Ctrl+Z)" disabled={!doc?.canUndo} onClick={() => commands['edit.undo']?.run?.()} />
          <Button icon="play" title="Start from the beginning (F5)" onClick={() => act('present', 'start')} />
        </>
      }
    >
      {/* ── Home ─────────────────────────────────────────────────────────── */}
      {tab === 'home' ? (
        <>
          <Group label="Clipboard">
            <Button tall icon="paste" label="Paste" title={canPaste ? 'Paste the copied shape on this slide (Ctrl+V)' : 'Paste — copy a shape first, then Ctrl+V'} disabled={!canPaste} onClick={() => act('pasteShape')} />
            <Rows>
              <>
                <Button icon="cut" label="Cut" title={needShape || 'Cut the shape (Ctrl+X)'} disabled={!hasShape} onClick={() => act('cutShape')} />
                <Button icon="copy" label="Copy" title={needShape || 'Copy the shape (Ctrl+C)'} disabled={!hasShape} onClick={() => act('copyShape')} />
              </>
              <Button
                icon="wand"
                label="Format Painter"
                pressed={painter}
                disabled={!hasShape && !painter}
                title={painter ? 'Format Painter — armed: click a shape to give it this look; Esc puts it down' : (needShape || 'Format Painter — this shape’s fill, outline and font onto the next shape you click')}
                onClick={() => act('painter')}
              />
            </Rows>
          </Group>
          <Group label="Slides">
            <Button tall icon="plus" label="New Slide" onClick={(e) => menu.open(e, LAYOUTS.map(([layout, label]) => ({ label, icon: 'slides', run: () => addSlide(layout) })))} />
            <Rows>
              <>
                <Button icon="grid" label="Layout" title="Layout — put this slide on another of the deck's layouts" onClick={(e) => menu.open(e, (model?.layouts || []).map((l) => ({ label: l.name || l.part, icon: l.part === model?.slide?.layout ? 'check' : undefined, run: () => act('applyLayout', l.part) })))} />
                <Button icon="undo" label="Reset" title="Reset — the placeholders back where the layout puts them" onClick={() => act('resetSlide')} />
                <Button
                  icon="list"
                  label="Section"
                  title="Section — a section before this slide, its name, or one taken away"
                  onClick={(e) => menu.open(e, [
                    { label: 'Add Section', icon: 'plus', run: () => act('addSection') },
                    { label: 'Rename Section…', icon: 'textbox', disabled: !inSection, run: () => act('renameSection') },
                    { label: 'Remove Section', icon: 'trash', disabled: !inSection, run: () => act('removeSection') },
                    { label: 'Remove All Sections', icon: 'close', disabled: !hasSections, run: () => act('removeAllSections') },
                  ])}
                />
              </>
              <>
                <Button icon="copy" label="Duplicate" onClick={() => commands['slide.new']?.run?.()} />
                <Button icon="trash" label="Delete" disabled={count < 2} onClick={() => commands['slide.delete']?.run?.()} />
              </>
            </Rows>
          </Group>
          <Group label="Font">
            <Rows>
              <>
                <Select value={format.font || ''} onChange={(e) => fmt({ font: e.target.value })} style={{ width: 118 }} title={needShape || 'Font'} disabled={!hasShape}>
                  <option value="">{format.font || (format.themeFont ? `${format.themeFont} (theme)` : 'Theme font')}</option>
                  {['Calibri', 'Calibri Light', 'Arial', 'Segoe UI', 'Georgia', 'Times New Roman', 'Verdana', 'Consolas'].filter((f) => f !== format.font).map((f) => <option key={f} value={f}>{f}</option>)}
                </Select>
                <Select value={String(size)} onChange={(e) => fmt({ size: Number(e.target.value) })} style={{ width: 58 }} title={needShape || 'Font size'} disabled={!hasShape}>
                  {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
                  {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
                </Select>
                <Button icon="chevronUp" title={needShape || 'Increase font size'} disabled={!hasShape} onClick={() => fmt({ size: nearer(1) })} />
                <Button icon="chevronDown" title={needShape || 'Decrease font size'} disabled={!hasShape} onClick={() => fmt({ size: nearer(-1) })} />
                <Button icon="undo" title={needShape || 'Clear all formatting — the words keep only their links'} disabled={!hasShape} onClick={() => act('clearFormat')} />
              </>
              <>
                <Button icon="bold" title={needShape || 'Bold (Ctrl+B)'} pressed={format.bold} disabled={!hasShape} onClick={() => fmt({ bold: 'toggle' })} />
                <Button icon="italic" title={needShape || 'Italic (Ctrl+I)'} pressed={format.italic} disabled={!hasShape} onClick={() => fmt({ italic: 'toggle' })} />
                <Button icon="underline" title={needShape || 'Underline (Ctrl+U)'} pressed={format.underline} disabled={!hasShape} onClick={() => fmt({ underline: 'toggle' })} />
                <Button icon="strike" title={needShape || 'Strikethrough'} pressed={format.strike} disabled={!hasShape} onClick={() => fmt({ strike: 'toggle' })} />
                <Button icon="textbox" label="AV" title={needShape || `Character spacing — now ${format.spacing ? `${format.spacing} pt` : 'normal'}`} disabled={!hasShape} onClick={(e) => menu.open(e, [['Very tight', -1.5], ['Tight', -0.75], ['Normal', 0], ['Loose', 1.5], ['Very loose', 3]].map(([label, v]) => ({ label, icon: (format.spacing || 0) === v ? 'check' : undefined, run: () => fmt({ spacing: v }) })))} />
                <Button icon="textbox" label="Aa" title={needShape || 'Change case'} disabled={!hasShape} onClick={(e) => menu.open(e, [['Sentence case', 'sentence'], ['lowercase', 'lower'], ['UPPERCASE', 'upper'], ['Capitalise Each Word', 'title']].map(([label, mode]) => ({ label, run: () => fmt({ case: mode }) })))} />
                <Separator />
                <Button icon="wand" title={needShape || 'Text highlight colour'} pressed={Boolean(format.highlight)} disabled={!hasShape} onClick={(e) => menu.open(e, [['#FFFF00', 'Yellow'], ['#00FF00', 'Bright green'], ['#00FFFF', 'Turquoise'], ['#FF00FF', 'Pink'], [null, 'No colour']].map(([value, label]) => ({ label, icon: value ? undefined : 'close', run: () => fmt({ highlight: value }) })))} />
                <Button icon="contrast" title={needShape || 'Font colour'} disabled={!hasShape} onClick={(e) => menu.open(e, COLOURS.map(([value, label]) => ({ label, run: () => fmt({ color: value }) })))} />
              </>
            </Rows>
          </Group>
          <Group label="Paragraph">
            <Rows>
              <>
                <Button icon="listBullet" title={needShape || 'Bullets'} pressed={format.bullet === 'char'} disabled={!hasShape} onClick={() => fmt({ bullet: format.bullet === 'char' ? 'none' : 'char' })} />
                <Button icon="listNumber" title={needShape || 'Numbering'} pressed={format.bullet === 'number'} disabled={!hasShape} onClick={() => fmt({ bullet: format.bullet === 'number' ? 'none' : 'number' })} />
                <Button icon="chevronLeft" title={needShape || 'Decrease list level'} disabled={!hasShape || !format.level} onClick={() => fmt({ level: -1 })} />
                <Button icon="chevronRight" title={needShape || 'Increase list level'} disabled={!hasShape} onClick={() => fmt({ level: 1 })} />
                <Button icon="list" title={needShape || `Line spacing — now ${format.lineHeight ? String(format.lineHeight) : 'as the layout has it'}`} disabled={!hasShape} onClick={(e) => menu.open(e, [1, 1.15, 1.5, 2].map((v) => ({ label: String(v), icon: format.lineHeight === v ? 'check' : undefined, run: () => fmt({ lineHeight: v }) })))} />
                <Separator />
                <Button icon="grid" title={needShape || `Columns — now ${format.columns || 1}`} disabled={!hasShape} onClick={(e) => menu.open(e, [['One column', 1], ['Two columns', 2], ['Three columns', 3]].map(([label, n]) => ({ label, icon: (format.columns || 1) === n ? 'check' : undefined, run: () => act('body', { columns: n }) })))} />
                <Button icon="rotate" title={needShape || `Text direction — now ${({ horz: 'horizontal', vert270: 'rotated up', vert: 'rotated down', eaVert: 'stacked' })[format.vert] || 'horizontal'}`} disabled={!hasShape} onClick={(e) => menu.open(e, [['Horizontal', 'horz'], ['Rotate all text 90° (reads down)', 'vert'], ['Rotate all text 270° (reads up)', 'vert270']].map(([label, v]) => ({ label, icon: (format.vert || 'horz') === v ? 'check' : undefined, run: () => act('body', { vert: v }) })))} />
                <Button icon="chevronUp" title={needShape || `Align text — now ${format.anchor || 'top'}`} disabled={!hasShape} onClick={(e) => menu.open(e, [['Top', 'top'], ['Middle', 'middle'], ['Bottom', 'bottom']].map(([label, a]) => ({ label, icon: (format.anchor || 'top') === a ? 'check' : undefined, run: () => act('body', { anchor: a }) })))} />
              </>
              <>
                <Button icon="alignLeft" title={needShape || 'Align left'} pressed={format.align === 'left'} disabled={!hasShape} onClick={() => fmt({ align: 'left' })} />
                <Button icon="alignCenter" title={needShape || 'Centre'} pressed={format.align === 'center'} disabled={!hasShape} onClick={() => fmt({ align: 'center' })} />
                <Button icon="alignRight" title={needShape || 'Align right'} pressed={format.align === 'right'} disabled={!hasShape} onClick={() => fmt({ align: 'right' })} />
                <Button icon="alignJustify" title={needShape || 'Justify'} pressed={format.align === 'justify'} disabled={!hasShape} onClick={() => fmt({ align: 'justify' })} />
                <Separator />
                <Soon icon="shape" label="SmartArt" why="SmartArt is a diagram part the engine does not write." />
              </>
            </Rows>
          </Group>
          <Group label="Drawing">
            <Button tall icon="shape" label="Shapes" title="Shapes — a rectangle, an oval, an arrow, a star, in the theme's colours" onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => act('addShape', preset) })))} />

            <Button tall icon="grid" label="Arrange" onClick={(e) => menu.open(e, [
              { label: 'Bring to front', icon: 'chevronUp', run: () => act('order', 'front') },
              { label: 'Bring forward', run: () => act('order', 'forward') },
              { label: 'Send backward', run: () => act('order', 'backward') },
              { label: 'Send to back', icon: 'chevronDown', run: () => act('order', 'back') },
              '-',
              { label: 'Group', icon: 'grid', disabled: multiCount < 2, title: multiCount < 2 ? 'Select two or more shapes to group' : undefined, run: () => act('group') },
              { label: 'Ungroup', icon: 'grid', disabled: !isGroup, title: !isGroup ? 'Select a group to ungroup' : undefined, run: () => act('ungroup') },
              '-',
              { label: 'Align Left', run: () => act('align', { edge: 'left', to: alignTo }) },
              { label: 'Align Center', run: () => act('align', { edge: 'center', to: alignTo }) },
              { label: 'Align Right', run: () => act('align', { edge: 'right', to: alignTo }) },
              { label: 'Align Top', run: () => act('align', { edge: 'top', to: alignTo }) },
              { label: 'Align Middle', run: () => act('align', { edge: 'middle', to: alignTo }) },
              { label: 'Align Bottom', run: () => act('align', { edge: 'bottom', to: alignTo }) },
              { label: 'Distribute Horizontally', disabled: multiCount < 3, title: multiCount < 3 ? 'Select three or more shapes to distribute' : undefined, run: () => act('distribute', { axis: 'horizontal', to: alignTo }) },
              { label: 'Distribute Vertically', disabled: multiCount < 3, title: multiCount < 3 ? 'Select three or more shapes to distribute' : undefined, run: () => act('distribute', { axis: 'vertical', to: alignTo }) },
              '-',
              { label: 'Align to Slide', icon: alignTo === 'slide' ? 'check' : undefined, run: () => setAlignTo('slide') },
              { label: 'Align Selected Objects', icon: alignTo === 'selection' ? 'check' : undefined, run: () => setAlignTo('selection') },
              '-',
              { label: 'Rotate Right 90°', run: () => act('rotateBy', 90) },
              { label: 'Rotate Left 90°', run: () => act('rotateBy', -90) },
              { label: 'Flip Vertical', run: () => act('flipShape', 'vertical') },
              { label: 'Flip Horizontal', run: () => act('flipShape', 'horizontal') },
              '-',
              { label: 'Selection pane (layers)', icon: 'list', run: () => act('pane', 'layers') },
              '-',
              { label: 'Delete shape', icon: 'trash', run: () => act('deleteShape') },
              { label: 'Move up (nudge)', run: () => act('nudge', { dy: -8 }) },
              { label: 'Move down (nudge)', run: () => act('nudge', { dy: 8 }) },
              { label: 'Move left (nudge)', run: () => act('nudge', { dx: -8 }) },
              { label: 'Move right (nudge)', run: () => act('nudge', { dx: 8 }) },
            ])} disabled={!hasShape} title={needShape || 'Arrange the selected shape(s)'} />
            <Button tall icon="wand" label="Quick Styles" disabled={!hasShape} title={needShape || 'The theme\'s own looks: filled in an accent, outlined in the same'} onClick={(e) => menu.open(e, [1, 2, 3, 4, 5, 6].map((n) => ({ label: `Accent ${n}`, icon: 'shape', run: () => act('quickStyle', n) })))} />
            <Button icon="wand" label="Shape Fill" disabled={!hasShape} title={needShape || 'The fill of the selected shape, in the Format pane'} onClick={() => act('formatPane')} />
            <Button icon="shape" label="Shape Outline" disabled={!hasShape} title={needShape || 'The outline of the selected shape, in the Format pane'} onClick={() => act('formatPane')} />
            <Button icon="wand" label="Shape Effects" disabled={!hasShape} title={needShape || 'Shape Effects — a shadow, a glow, soft edges or a reflection on the selected shape'} onClick={(e) => menu.open(e, [
              ...SHADOW_MENU.map(([key, label]) => ({ label, icon: key === 'none' ? 'close' : undefined, run: () => act('shapeShadow', key) })),
              '-',
              ...GLOW_MENU.map(([pt, label]) => ({ label, icon: pt === null ? 'close' : undefined, run: () => act('shapeEffects', { glow: pt === null ? null : { radius: pt, color: selectedShapeObj?.effects?.glow?.color } }) })),
              '-',
              ...SOFTEDGE_MENU.map(([pt, label]) => ({ label, icon: pt === null ? 'close' : undefined, run: () => act('shapeEffects', { softEdge: pt === null ? null : { radius: pt } }) })),
              '-',
              ...REFLECTION_MENU.map(([key, label]) => ({ label, icon: key === null ? 'close' : undefined, run: () => act('shapeEffects', { reflection: key }) })),
            ])} />
          </Group>
          <Group label="Editing">
            <Button tall icon="find" label="Find" title="Find — words on every slide, walked one hit at a time (Ctrl+F)" onClick={() => act('find')} />
            <Button tall icon="find" label="Replace" title="Replace — find and replace words across the deck (Ctrl+H)" onClick={() => act('replace')} />
          </Group>
          <Group label="Voice">
            <Soon tall icon="volume" label="Dictate" why="Dictation is an online speech service this suite does not call." />
          </Group>
          <Group label="Designer">
            <Soon tall icon="wand" label="Design Suggestions" why="Design Ideas is a Microsoft cloud service." />
          </Group>
        </>
      ) : null}

      {/* ── Insert ───────────────────────────────────────────────────────── */}
      {tab === 'insert' ? (
        <>
          <Group label="Slides">
            <Button tall icon="plus" label="New Slide" onClick={(e) => menu.open(e, LAYOUTS.map(([layout, label]) => ({ label, icon: 'slides', run: () => addSlide(layout) })))} />
          </Group>
          <Group label="Tables">
            <Button tall icon="table" label="Table" title="Table — a grid of cells, in PowerPoint's own default style" onClick={(e) => menu.open(e, TABLE_SIZES.map(([rows, cols, label]) => ({ label, icon: 'table', run: () => act('addTable', { rows, cols }) })))} />
          </Group>
          <Group label="Images">
            <Button tall icon="picture" label="Pictures" title="Pictures — a picture from this device, onto this slide" onClick={insertPicture} />

            <Soon tall icon="picture" label="Screenshot" why="Comes with pictures." />
            <Soon tall icon="picture" label="Photo Album" why="Comes with pictures." />
          </Group>
          <Group label="Camera">
            <Soon tall icon="video" label="Cameo" why="A live camera feed on a slide is a PowerPoint-only feature." />
          </Group>
          <Group label="Illustrations">
            <Button tall icon="shape" label="Shapes" title="Shapes — a rectangle, an oval, an arrow, a star, in the theme's colours" onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => act('addShape', preset) })))} />

            <Soon tall icon="star" label="Icons" why="Icons are an online library." />
            <Soon tall icon="shape" label="3D Models" why="3D models are an online library." />
            <Soon tall icon="shape" label="SmartArt" why="SmartArt is a diagram part the engine does not write." />
            <Button tall icon="chart" label="Chart" title="Chart — a sample chart, drawn from the writer Word and Worksheets already use; double-click it to edit its data" onClick={(e) => menu.open(e, CHART_TYPES.map(([type, label]) => ({ label, icon: 'chart', run: () => act('addChart', { type }) })))} />
          </Group>
          <Group label="Forms">
            <Soon tall icon="check" label="Forms" why="Forms is a Microsoft cloud service." />
          </Group>
          <Group label="Links">
            <Soon tall icon="zoomIn" label="Zoom" why="Slide zooms are a PowerPoint-only feature." />
            <Button tall icon="link" label="Link" disabled={!hasShape} title={needShape || 'Link — a web address on the selected shape’s words (Ctrl+K); Ctrl+click follows it'} onClick={() => act('link')} />
            <Soon tall icon="play" label="Action" why="Comes with links." />
          </Group>
          <Group label="Comments">
            <Soon tall icon="reply" label="Comment" why="Slide comments are a comments part the engine does not write yet." />
          </Group>
          <Group label="Text">
            <Button tall icon="textbox" label="Text Box" onClick={() => commands['slide.textbox']?.run?.()} />
            <Button tall icon="file" label="Header & Footer" title="Header & Footer — the date, the slide number and the footer's words, on this slide or all of them" onClick={() => act('footer')} />
            <Soon tall icon="wand" label="WordArt" why="WordArt is text effects the deck writer does not write." />
            <Button icon="clock" label="Date & Time" title="Date & Time — a date along the bottom, kept current or fixed" onClick={() => act('footer', 'date')} />
            <Button icon="list" label="Slide Number" title="Slide Number — the slide's number along the bottom, following it when slides move" onClick={() => act('footer', 'number')} />
            <Soon icon="file" label="Object" why="Embedded objects are not built." />
          </Group>
          <Group label="Symbols">
            <Soon tall icon="formula" label="Equation" why="Equations are OMML the engine does not write." />
            <Soon tall icon="plus" label="Symbol" why="Comes with text selection inside a box." />
          </Group>
          <Group label="Media">
            <Soon tall icon="video" label="Video" why="Video on a slide is a media part plus a relationship the deck writer does not write yet." />
            <Soon tall icon="volume" label="Audio" why="Comes with media." />
            <Soon tall icon="video" label="Screen Recording" why="Comes with media." />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label="Drawing Tools">
            <Soon tall icon="wand" label="Select" why={INK} />
            <Soon tall icon="wand" label="Lasso" why={INK} />
            <Soon tall icon="close" label="Eraser" why={INK} />
            <Soon tall icon="wand" label="Pen" why={INK} />
            <Soon tall icon="wand" label="Pencil" why={INK} />
            <Soon tall icon="wand" label="Highlighter" why={INK} />
            <Soon tall icon="plus" label="Add" why={INK} />
          </Group>
          <Group label="Touch">
            <Soon tall icon="wand" label="Draw with Touch" why={INK} />
          </Group>
          <Group label="Stencils">
            <Soon tall icon="minus" label="Ruler" why="The drawing ruler comes with ink." />
          </Group>
          <Group label="Convert">
            <Soon tall icon="shape" label="Ink to Shape" why={INK} />
            <Soon tall icon="formula" label="Ink to Math" why={INK} />
          </Group>
          <Group label="Replay">
            <Soon tall icon="play" label="Ink Replay" why={INK} />
          </Group>
          <Group label="Help">
            <Button tall icon="info" label="Ink Help" onClick={() => act('help')} />
          </Group>
        </>
      ) : null}

      {/* ── Design ───────────────────────────────────────────────────────── */}
      {tab === 'design' ? (
        <>
          <Group label="Themes">
            <RibbonStrip
              kind="themes"
              items={designStrip?.themes}
              count={4}
              label="Drawing this slide…"
              onPick={(it) => act('applyTheme', it)}
              onMore={(e) => { const r = e.currentTarget.closest('.sl-rs').getBoundingClientRect(); act('designGallery', { kind: 'themes', anchor: { left: r.left, bottom: r.bottom + 4 } }); }}
            />
          </Group>
          <Group label="Variants">
            <RibbonStrip
              kind="variants"
              items={designStrip?.variants}
              count={4}
              label="Drawing this slide…"
              onPick={(it) => act('applyVariant', it.id)}
              onMore={(e) => { const r = e.currentTarget.closest('.sl-rs').getBoundingClientRect(); act('designGallery', { kind: 'variants', anchor: { left: r.left, bottom: r.bottom + 4 } }); }}
            />
            <Rows>
              <Button icon="contrast" label="Colours" title="Colours — the theme's twelve colours: a palette, or your own" onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind: 'colours', anchor: { left: r.left, bottom: r.bottom + 4 } }); }} />
              <Button icon="textbox" label="Fonts" title="Fonts — the theme's heading and body faces: a pair, or your own" onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind: 'fonts', anchor: { left: r.left, bottom: r.bottom + 4 } }); }} />
              <Button icon="wand" label="Effects" title="Effects — how theme-styled shapes are filled, outlined and lifted" onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind: 'effects', anchor: { left: r.left, bottom: r.bottom + 4 } }); }} />
            </Rows>
          </Group>
          <Group label="Customise">
            <Button
              tall
              icon="picture"
              label="Background Styles"
              title="Background Styles — a solid or gradient background for this slide, or every slide"
              onClick={(e) => {
                const own = model?.slide?.ownBackground ?? null;
                const same = (spec) => JSON.stringify(spec ?? null) === JSON.stringify(own);
                menu.open(e, [
                  ...BACKGROUND_STYLES.map((row) => (row === '-' ? '-' : { label: row[0], icon: same(row[1]) ? 'check' : undefined, run: () => act('background', { spec: row[1] }) })),
                  '-',
                  { label: "Reset to the layout's", icon: same(null) ? 'check' : undefined, run: () => act('background', { spec: null }) },
                  '-',
                  { label: 'Apply to all slides', run: () => act('background', { spec: own, all: true }) },
                ]);
              }}
            />
          </Group>
          <Group label="Slides">
            <Button tall icon="grid" label="Layouts" pressed={view.pane === 'designs'} title="The deck's layouts in a pane — put this slide on one, or start a new slide from it" onClick={() => act('pane', 'designs')} />
            <Button tall icon="slides" label="Layout" title="A new slide with this layout, after the current one" onClick={(e) => menu.open(e, LAYOUTS.map(([layout, label]) => ({ label, icon: 'slides', run: () => addSlide(layout) })))} />
            <Button
              tall
              icon="grid"
              label="Slide Size"
              title={model?.size ? `Slides are ${Math.round(model.size.width)} by ${Math.round(model.size.height)} pixels — ${Math.abs(model.size.width / model.size.height - 16 / 9) < 0.02 ? '16:9 widescreen' : Math.abs(model.size.width / model.size.height - 4 / 3) < 0.02 ? '4:3 standard' : 'a custom ratio'}. Changing it rescales every slide; not built yet.` : 'The deck decides its own size'}
              disabled
            />
          </Group>
          <Group label="Edit">
            <Soon tall icon="wand" label="Design Suggestions" why="Design Ideas is a Microsoft cloud service." />
          </Group>
        </>
      ) : null}

      {/* ── Transitions ──────────────────────────────────────────────────── */}
      {tab === 'transitions' ? (
        <>
          <Group label="Preview">
            <Button
              tall
              icon="play"
              label="Preview"
              disabled={!transitionPlays}
              title={transitionPlays ? `Preview — play this slide's transition on the stage (${describeTransition(transition)})` : 'Preview — this slide has no transition to play; pick one from the gallery'}
              onClick={() => act('preview', 'transition')}
            />
          </Group>
          <Group label="Transition to This Slide">
            {TRANSITION_GALLERY.map(([key, label, blurb]) => (
              <Button
                key={key}
                tall
                icon={TRANSITION_ICONS[key]}
                label={label}
                className="sl-tr-pick"
                data-transition={key}
                pressed={transitionKey === key}
                title={`${label} — ${blurb}`}
                onClick={() => act('transition', key === 'none' ? { type: 'none' } : { type: key === 'shape' ? 'circle' : key })}
              />
            ))}
            <Button
              tall
              icon="settings"
              label="Effect Options"
              disabled={!transitionOptions}
              title={transitionOptions ? `Effect Options — which way the ${TRANSITION_GALLERY.find(([k]) => k === transitionKey)?.[1] || 'transition'} goes` : 'Effect Options — pick a transition with a direction first (Push, Wipe, Split, Cover…)'}
              onClick={(e) => menu.open(e, transitionOptions.map(([value, label]) => ({
                label,
                icon: optionOf(transition) === value ? 'check' : undefined,
                run: () => act('transition', transitionKey === 'shape' ? { type: value.split(':')[0], direction: value.split(':')[1] ?? null } : { direction: value }),
              })))}
            />
          </Group>
          <Group label="Timing">
            <Rows>
              <Soon icon="volume" label="Sound" why="A transition sound is a media part the deck writer does not embed; one already in the file is kept." />
              <div className="sl-rb-field" data-tip="Duration — how long the transition takes, in seconds">
                <Icon name="clock" size={15} />
                <span>Duration</span>
                <SecondsField className="sl-tr-duration" value={transition && transitionKey !== 'none' ? transition.duration : null} disabled={!transition || transitionKey === 'none'} min={0.01} max={59} onCommit={(v) => act('transition', { duration: v })} />
              </div>
              <Button icon="check" label="Apply To All" disabled={count < 2} title={count < 2 ? 'Apply To All — the deck has one slide' : "Apply To All — this slide's transition, timing and advance settings on every slide"} onClick={() => act('transitionAll')} />
            </Rows>
            <Separator />
            <Rows>
              <div className="sl-rb-caption">Advance Slide</div>
              <label className="sl-rb-field" data-tip="On Mouse Click — a click moves the show on from this slide">
                <input type="checkbox" className="sl-tr-onclick" checked={transition ? transition.advanceOnClick !== false : true} onChange={(e) => act('transition', { advanceOnClick: e.target.checked })} />
                <span>On Mouse Click</span>
              </label>
              <div className="sl-rb-field" data-tip="After — the show moves on by itself this many seconds after the slide is shown">
                <label className="sl-rb-check">
                  <input type="checkbox" className="sl-tr-after-on" checked={transition?.advanceAfter != null} onChange={(e) => act('transition', { advanceAfter: e.target.checked ? (transition?.advanceAfter ?? 0) : null })} />
                  <span>After:</span>
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
          <Group label="Preview">
            <Button
              tall
              icon="play"
              label="Preview"
              disabled={!animationList.length}
              title={animationList.length ? `Preview — play this slide's ${animationList.length === 1 ? 'animation' : `${animationList.length} animations`} on the stage` : 'Preview — this slide has no animations yet; select a shape and pick an effect'}
              onClick={() => act('preview', 'animation')}
            />
          </Group>
          <Group label="Animation">
            <Button
              tall
              icon="star"
              label="None"
              className="sl-an-pick sl-an-none"
              data-effect="none"
              pressed={hasShape && !shapeAnimated}
              disabled={!hasShape}
              title={hasShape ? 'None — take every animation off the selected shape' : 'None — select a shape first; its animations come off'}
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
                title={hasShape ? `${label} — ${blurb}${animation ? '; it replaces the selected effect' : ''}` : `${label} — select a shape first, then pick its entrance`}
                onClick={() => act('animate', { kind: 'entr', effect })}
              />
            ))}
            <Button
              tall
              icon="more"
              label="More Effects"
              disabled={!hasShape}
              title={hasShape ? 'More Effects — every entrance, emphasis and exit effect' : 'More Effects — select a shape first'}
              onClick={(e) => menu.open(e, effectMenu('animate'))}
            />
            <Button
              tall
              icon="settings"
              label="Effect Options"
              disabled={!animationOptions}
              title={animationOptions ? `Effect Options — which way the ${animation?.name || 'effect'} goes` : 'Effect Options — pick an effect with a direction first (Fly In, Float In, Split, Wipe, Spin)'}
              onClick={(e) => menu.open(e, animationOptions.map(([value, label]) => ({ label, icon: animation?.direction === value ? 'check' : undefined, run: () => act('animPatch', { direction: value }) })))}
            />
          </Group>
          <Group label="Advanced Animation">
            <Button tall icon="plus" label="Add Animation" disabled={!hasShape} title={hasShape ? 'Add Animation — another effect on the selected shape, after its others' : 'Add Animation — select a shape first'} onClick={(e) => menu.open(e, effectMenu('addAnimation'))} />
            <Rows>
              <Button icon="list" label="Animation Pane" pressed={view.pane === 'animations'} title="Animation Pane — the slide's effects in order, to pick, reorder, retime or remove" onClick={() => act('pane', 'animations')} />
              <Soon icon="play" label="Trigger" why="A trigger — starting an effect on a click on another shape — is an interactive sequence this does not write; one already in the file is kept and plays as the file has it." />
              <Button
                icon="wand"
                label="Animation Painter"
                pressed={animPainter}
                disabled={!animPainter && !shapeAnimated}
                title={animPainter ? 'Animation Painter — armed: click a shape to give it these effects; Esc puts it down' : shapeAnimated ? "Animation Painter — this shape's effects onto the next shape you click" : 'Animation Painter — select an animated shape first'}
                onClick={() => act('animPainter')}
              />
            </Rows>
          </Group>
          <Group label="Timing">
            <Rows>
              <div className="sl-rb-field" data-tip="Start — when the effect plays: on a click, with the one before it, or after it">
                <Icon name="play" size={14} />
                <span className="sl-rb-label">Start:</span>
                <Select className="rw-select sl-an-start" value={animation?.trigger || 'onClick'} disabled={!animation} style={{ width: 128, height: 24 }} onChange={(e) => act('animPatch', { trigger: e.target.value })}>
                  <option value="onClick">On Click</option>
                  <option value="withPrevious">With Previous</option>
                  <option value="afterPrevious">After Previous</option>
                </Select>
              </div>
              <div className="sl-rb-field" data-tip="Duration — how long the effect takes, in seconds">
                <Icon name="clock" size={14} />
                <span className="sl-rb-label">Duration:</span>
                <SecondsField className="sl-an-duration" value={animation ? animation.duration : null} disabled={!animation || !animation.duration} min={0.01} max={59} onCommit={(v) => act('animPatch', { duration: v })} />
              </div>
              <div className="sl-rb-field" data-tip="Delay — how long after its start the effect waits, in seconds">
                <Icon name="clock" size={14} />
                <span className="sl-rb-label">Delay:</span>
                <SecondsField className="sl-an-delay" value={animation ? animation.delay : null} disabled={!animation} min={0} max={3600} onCommit={(v) => act('animPatch', { delay: v })} />
              </div>
            </Rows>
            <Separator />
            <Rows>
              <div className="sl-rb-caption">Reorder Animation</div>
              <Button icon="chevronUp" label="Move Earlier" disabled={!animation || animation.index <= 0} title={animation ? 'Move Earlier — this effect plays before the one above it' : 'Move Earlier — pick an effect first'} onClick={() => act('animMove', 'earlier')} />
              <Button icon="chevronDown" label="Move Later" disabled={!animation || animation.index >= animationList.length - 1} title={animation ? 'Move Later — this effect plays after the one below it' : 'Move Later — pick an effect first'} onClick={() => act('animMove', 'later')} />
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Slide Show ───────────────────────────────────────────────────── */}
      {tab === 'show' ? (
        <>
          <Group label="Start Slide Show">
            <Button tall icon="play" label="From Beginning" title="Start the show at the first slide (F5)" onClick={() => act('present', 'start')} />
            <Button tall icon="play" label="From Current Slide" title="Start the show at this slide" onClick={() => act('present', 'here')} />
            <Soon tall icon="list" label="Custom Slide Show" why="Named subsets of a deck are a presentation-part list the engine does not write yet." />
          </Group>
          <Group label="Rehearse">
            <Soon tall icon="volume" label="Rehearse with Coach" why="Presenter Coach is a Microsoft cloud service." />
          </Group>
          <Group label="Set Up">
            <Soon tall icon="settings" label="Set Up Slide Show" why="Show settings (kiosk, loop, pen colour) are not built." />
            <Button
              tall
              icon="eye"
              label="Hide Slide"
              pressed={Boolean(model?.slide?.hidden)}
              title={model?.slide?.hidden ? 'Hide Slide — this slide is hidden from the show; press again to show it' : 'Hide Slide — leave this slide out of the show'}
              onClick={() => act('hideSlide')}
            />
            <Soon tall icon="clock" label="Rehearse Timings" why="Comes with recorded timings." />
            <Soon tall icon="video" label="Record" why="Recording narration is not built." />
            <Soon icon="check" label="Keep Slides Updated" why="Live co-authoring is not built." />
            <Soon icon="check" label="Play Narrations" why="Comes with recording." />
            <Soon icon="check" label="Use Timings" why="Comes with recorded timings." />
            <Soon icon="check" label="Show Media Controls" why="Comes with media." />
          </Group>
          <Group label="Monitors">
            <Soon icon="grid" label="Monitor: Automatic" why="Choosing the display comes with a second-window layout the shell does not manage yet; Presenter View opens on whichever screen you drag it to." />
            <Button icon="check" label="Use Presenter View" title="Opens a second window with your notes, the next slide and a clock — put it on the other screen" onClick={presentWithNotes} />
          </Group>
          <Group label="Captions & Subtitles">
            <Soon icon="check" label="Always Use Subtitles" why="Live captions are an online speech service this suite does not call." />
            <Soon icon="settings" label="Subtitle Settings" why="Comes with captions." />
          </Group>
        </>
      ) : null}

      {/* ── Record ───────────────────────────────────────────────────────── */}
      {tab === 'record' ? (
        <>
          <Group label="Preview">
            <Soon tall icon="play" label="Preview" why="Comes with recording." />
          </Group>
          <Group label="Camera">
            <Soon tall icon="video" label="Cameo" why="A live camera feed on a slide is a PowerPoint-only feature." />
          </Group>
          <Group label="Record">
            <Soon tall icon="video" label="From Beginning" why="Recording a narrated show (audio plus timings into the file) is not built." />
            <Soon tall icon="video" label="From Current Slide" why="Comes with recording." />
            <Soon tall icon="video" label="Screen Recording" why="Comes with media." />
            <Soon tall icon="volume" label="Audio" why="Comes with media." />
          </Group>
          <Group label="Edit">
            <Soon tall icon="close" label="Clear Recording" why="Comes with recording." />
            <Soon tall icon="undo" label="Reset to Cameo" why="Comes with recording." />
          </Group>
          <Group label="Export">
            <Soon tall icon="export" label="Save as Show" why="A .ppsx is a .pptx with a different content type; on the list." />
            <Soon tall icon="video" label="Export to Video" why="Rendering a deck to video (frames plus encoding) is not built." />
          </Group>
          <Group label="Help">
            <Button tall icon="info" label="Learn More" onClick={() => act('help')} />
          </Group>
        </>
      ) : null}

      {/* ── Review ───────────────────────────────────────────────────────── */}
      {tab === 'review' ? (
        <>
          <Group label="Proofing">
            <Soon tall icon="check" label="Spelling" why="The system spell checker underlines as you type in a box; a pass over the deck is not built." />
            <Soon tall icon="find" label="Thesaurus" why="A thesaurus is a dictionary this suite does not ship." />
          </Group>
          <Group label="Accessibility">
            <Soon tall icon="check" label="Check Accessibility" why="An accessibility pass is on the list." />
          </Group>
          <Group label="Language">
            <Soon tall icon="globe" label="Translate" why="Translation is an online service this suite does not call." />
            <Soon tall icon="globe" label="Language" why="Proofing language is a run property the deck writer does not write yet." />
          </Group>
          <Group label="Activity">
            <Soon icon="check" label="Mark All as Read" why="Comes with comments." />
            <Soon icon="eye" label="Show Changes" why="Change tracking on a deck is not built." />
          </Group>
          <Group label="Comments">
            <Soon tall icon="reply" label="New Comment" why="Slide comments are a comments part the engine does not write yet." />
            <Soon icon="close" label="Delete" why="Comes with comments." />
            <Soon icon="chevronLeft" label="Previous" why="Comes with comments." />
            <Soon icon="chevronRight" label="Next" why="Comes with comments." />
            <Soon icon="eye" label="Show Comments" why="Comes with comments." />
          </Group>
          <Group label="Notes">
            <Button tall icon="word" label="Speaker Notes" title="Speaker Notes — this slide's, shown in Presenter View" onClick={() => setNotesOpen(true)} />
          </Group>
          <Group label="Ink">
            <Soon icon="eye" label="Hide Ink" why={INK} />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
      {tab === 'view' ? (
        <>
          <Group label="Presentation Views">
            <Button tall icon="slides" label="Normal" pressed={(view.mode || 'normal') === 'normal'} onClick={() => act('mode', 'normal')} />
            <Button tall icon="list" label="Outline View" pressed={view.mode === 'outline'} title="Every slide's words, as an outline" onClick={() => act('mode', 'outline')} />
            <Button tall icon="grid" label="Slide Sorter" pressed={view.mode === 'sorter'} title="All the slides at once — click one to open it" onClick={() => act('mode', 'sorter')} />
            <Button tall icon="word" label="Notes Page" pressed={view.mode === 'notes'} title="The slide with its notes under it" onClick={() => act('mode', 'notes')} />
            <Button tall icon="play" label="Reading View" title="The show in this window, not full screen" onClick={() => act('present', 'reading')} />
          </Group>
          <Group label="Master Views">
            <Soon tall icon="slides" label="Slide Master" why="Master editing rewrites layouts every slide inherits from; the engine reads masters and does not yet write one." />
            <Soon tall icon="file" label="Handout Master" why="Comes with master editing." />
            <Soon tall icon="word" label="Notes Master" why="Comes with master editing." />
          </Group>
          <Group label="Show">
            <Button icon="minus" label="Ruler" pressed={Boolean(view.ruler)} onClick={() => act('toggle', 'ruler')} />
            <Button icon="grid" label="Gridlines" pressed={Boolean(view.gridlines)} onClick={() => act('toggle', 'gridlines')} />
            <Button icon="plus" label="Guides" pressed={Boolean(view.guides)} onClick={() => act('toggle', 'guides')} />
            <Button icon="word" label="Notes" pressed={view.notes !== false} title="The notes strip under the slide" onClick={() => act('toggle', 'notes')} />
            <Button icon="list" label="Layers" pressed={view.pane === 'layers'} title="The slide's shapes as layers — select, reorder, hide, rename" onClick={() => act('pane', 'layers')} />
            <Button icon="grid" label="Designs" pressed={view.pane === 'designs'} title="The deck's layouts — put this slide on one, or start a new slide from it" onClick={() => act('pane', 'designs')} />
          </Group>
          <Group label="Direction">
            <Soon icon="rotate" label="View Direction" why="Right-to-left layout is not built." />
          </Group>
          <Group label="Zoom">
            <Button tall icon="zoomIn" label="Zoom" onClick={(e) => menu.open(e, [50, 75, 100, 150, 200].map((z) => ({ label: `${z}%`, run: () => act('zoom', z / 100) })))} />
            <Button tall icon="maximize" label="Fit to Window" pressed={!view.zoom} onClick={() => act('zoom', null)} />
          </Group>
          <Group label="Colour/Greyscale">
            <Button icon="contrast" label="Colour" pressed={!view.tone || view.tone === 'colour'} onClick={() => act('tone', 'colour')} />
            <Button icon="contrast" label="Greyscale" pressed={view.tone === 'grey'} onClick={() => act('tone', 'grey')} />
            <Button icon="contrast" label="Black and White" pressed={view.tone === 'mono'} onClick={() => act('tone', 'mono')} />
          </Group>
          <Group label="Window">
            <Button tall icon="new" label="New Window" title="This deck in a second window" onClick={() => act('newWindow')} />
            <Soon icon="grid" label="Arrange All" why="Window tiling is the operating system's; not built." />
            <Soon icon="grid" label="Cascade" why="Comes with window tiling." />
            <Soon icon="minus" label="Move Split" why="The slide pane is dragged at its edge; a split command is not built." />
            <Soon icon="list" label="Switch Windows" why="Switching windows is the operating system's; not built." />
            <Button icon="maximize" label="Full Screen" onClick={() => shell.win.fullscreen({})} />
          </Group>
          <Group label="Macros">
            <Soon tall icon="settings" label="Macros" why="VBA is preserved in the file and never run: a deck that runs code it received in an email is how ransomware starts." />
          </Group>
        </>
      ) : null}

      {/* ── Help ─────────────────────────────────────────────────────────── */}
      {tab === 'help' ? (
        <>
          <Group label="Help & Support">
            <Button tall icon="info" label="Help" title="The Presentation guide on office.rutba.io" onClick={() => act('help')} />
            <Button tall icon="send" label="Feedback" title="Tell us what is wrong or missing" onClick={() => act('feedback')} />
            <Button tall icon="info" label="Show Training" onClick={() => act('help')} />
            <Button tall icon="star" label="What's New" title="The release notes" onClick={() => act('releases')} />
            <Button icon="list" label="Keyboard Shortcuts" onClick={() => act('shortcuts')} />
          </Group>
          <Group label="File">
            <Button tall icon="new" label="New" onClick={() => shell.win.create({ app: 'slides' })} />
            <Button tall icon="open" label="Open" onClick={openFile} />
            <Button tall icon="save" label="Save" onClick={() => save(false)} />
          </Group>
        </>
      ) : null}

      {/* ── PDF ──────────────────────────────────────────────────────────── */}
      {tab === 'pdf' ? (
        <>
          <Group label="Export">
            <Button tall icon="pdf" label="Export as PDF" title="Every slide as a page, drawn exactly as the stage draws it" onClick={() => exportAs('pdf')} />
          </Group>
        </>
      ) : null}
    </Ribbon>
  );
}
