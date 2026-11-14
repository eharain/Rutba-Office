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
import { Ribbon, Group, Rows, Button, Separator, Select, Icon, t, tn, numberTag } from '@rutba/office-ui';
import { WARP_PRESETS, WARP_MORE, warpLabel } from '@rutba/drawing/warp';
import { TRANSITION_GALLERY, TRANSITION_OPTIONS, galleryKeyOf, optionOf, describeTransition } from './motion.js';
import { ANIMATION_GALLERY, EFFECT_MENU, ANIMATION_OPTIONS } from './animate.js';
import { MODEL_VIEWS } from '@rutba/imaging/model3d';
import { RibbonStrip } from '../../design.js';
import { wordArtMenu } from '../../wordart.js';
import { SOUNDS, soundFile } from './sounds.js';
import { PEN_COLOURS, PEN_WIDTHS } from '../../ink.js';
import { ColourSplit } from '../colour/palette.js';
import { HIGHLIGHT, sameColour } from '../colour/grid.js';

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
/** A colour's little square in a menu, as the colour menus show it. */
const swatch = (c) => React.createElement('span', { style: { display: 'inline-block', width: 16, height: 16, borderRadius: 3, background: c, border: '1px solid rgba(0,0,0,.25)' } });
const COLOURS = [
  ['#000000', t('common.black')], ['#404040', t('slides.ribbon.dark-grey')], ['#808080', t('common.grey')], ['#ffffff', t('common.white')],
  ['#c00000', t('common.dark-red')], ['#ff0000', t('common.red')], ['#ffc000', t('common.orange')], ['#ffff00', t('common.yellow')],
  ['#92d050', t('common.light-green')], ['#00b050', t('common.green')], ['#00b0f0', t('common.light-blue')], ['#0070c0', t('common.blue')],
  ['#002060', t('common.dark-blue')], ['#7030a0', t('common.purple')],
];
const LAYOUTS = [
  ['title', t('slides.ribbon.title-slide')], ['obj', t('slides.ribbon.title-and-content')], ['blank', t('common.blank')],
];
/** The preset geometries the engine writes and the renderer draws, in PowerPoint's names. */
const SHAPES = [
  ['rect', t('common.rectangle')], ['roundRect', t('slides.ribbon.rectangle-rounded')], ['ellipse', t('common.oval')],
  ['triangle', t('slides.ribbon.isosceles-triangle')], ['rtTriangle', t('slides.ribbon.right-triangle')], ['diamond', t('common.diamond')],
  ['parallelogram', t('common.parallelogram')], ['trapezoid', t('common.trapezoid')], ['pentagon', t('common.pentagon')],
  ['hexagon', t('common.hexagon')], ['octagon', t('slides.ribbon.octagon')], ['star5', t('slides.ribbon.star-5-points')],
  ['rightArrow', t('slides.ribbon.arrow-right')], ['chevron', t('common.chevron')], ['line', t('common.line')],
];
/** Insert → Table: the sizes PowerPoint's own gallery offers first. */
const TABLE_SIZES = [
  [2, 2, '2 × 2'], [3, 3, '3 × 3'], [4, 3, '4 × 3'], [5, 4, '5 × 4'],
];
/** Insert → Chart: the kinds `chartPartXml` writes — the ones a chart on a slide can be. */
const CHART_TYPES = [
  ['column', t('common.column')], ['bar', t('common.bar')], ['line', t('common.line')], ['pie', t('common.pie')],
];

/** Design → Background Styles: the theme backgrounds PowerPoint's gallery offers first, then a few flat colours. */
const BACKGROUND_STYLES = [
  [t('common.background-1'), { scheme: 'bg1' }],
  [t('common.background-2'), { scheme: 'bg2' }],
  [t('common.text-1'), { scheme: 'tx1' }],
  [t('common.text-2'), { scheme: 'tx2' }],
  [t('slides.ribbon.accent-1-gradient'), { gradient: { from: { scheme: 'accent1' }, to: { scheme: 'accent1', lumMod: 75 } } }],
  '-',
  [t('common.white'), { colour: 'FFFFFF' }],
  [t('common.black'), { colour: '000000' }],
  [t('common.light-grey'), { colour: 'F2F2F2' }],
  [t('common.dark-blue'), { colour: '1F3864' }],
];

/** EMU as the centimetres Slide Size shows, two places at most, in the language's digits. */
export const cmOf = (emu) => (Math.round(Number(emu) / 3600) / 100).toLocaleString(numberTag(), { maximumFractionDigits: 2 });
/** A slide's shape in words: widescreen, standard or another. */
export function ratioOf(size) {
  if (!size?.width || !size?.height) return '';
  const r = size.width / size.height;
  return Math.abs(r - 16 / 9) < 0.02 ? t('slides.ribbon.16-9-widescreen') : Math.abs(r - 4 / 3) < 0.02 ? t('slides.ribbon.4-3-standard') : t('slides.ribbon.a-custom-ratio');
}
/** Shape Effects: the shadows PowerPoint offers first, and none. */
const SHADOW_MENU = [['br', t('slides.ribbon.shadow-bottom-right')], ['b', t('slides.ribbon.shadow-below')], ['r', t('slides.ribbon.shadow-right')], ['tl', t('slides.ribbon.shadow-top-left')], ['c', t('slides.ribbon.shadow-all-round')], ['none', t('slides.ribbon.no-shadow')]];
/** Shape Effects → Glow: PowerPoint's own gallery radii. */
const GLOW_MENU = [[5, t('slides.ribbon.glow-5-pt')], [8, t('slides.ribbon.glow-8-pt')], [11, t('slides.ribbon.glow-11-pt')], [18, t('slides.ribbon.glow-18-pt')], [null, t('common.no-glow')]];
/** Shape Effects → Soft Edges. */
const SOFTEDGE_MENU = [[1, t('slides.ribbon.soft-edges-1')], [2.5, t('slides.ribbon.soft-edges-2')], [5, t('slides.ribbon.soft-edges-5')], [10, t('slides.ribbon.soft-edges-10')], [null, t('slides.ribbon.no-soft-edges')]];
/** Shape Effects → Reflection: the three gallery presets. */
const REFLECTION_MENU = [['tight', t('slides.ribbon.reflection-tight')], ['half', t('slides.ribbon.reflection-half')], ['full', t('slides.ribbon.reflection-full')], [null, t('slides.ribbon.no-reflection')]];
/** The gallery's small pictures, one per effect. */
const TRANSITION_ICONS = { none: 'trNone', cut: 'trCut', fade: 'trFade', push: 'trPush', wipe: 'trWipe', split: 'trSplit', pull: 'trUncover', cover: 'trCover', randomBar: 'trBars', shape: 'trShape', dissolve: 'trDissolve' };

/**
 * A number of seconds in the ribbon (Duration, After): typed or stepped,
 * and sent once it settles — on Enter, on leaving the box, or a moment
 * after the last keystroke — rather than once per keystroke.
 */
function SecondsField({ value, onCommit, disabled = false, min = 0, max = 3600, className = '', label }) {
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
      aria-label={label}
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
  monitorName = t('common.automatic'),
  ink = null,
  redaction = null,
  designStrip = null,
  masterView = null,
  masterPart = null,
  review = null,

}) {
  const count = model?.count || 0;
  /** How many comment threads the deck holds. */
  const commentCount = (model?.comments || []).length;
  const hasShape = Boolean(selected);
  // New Slide: the deck's own layouts, by their names, as PowerPoint lists
  // them; the three plain ones for a deck that names none.
  const newSlideMenu = (e) => {
    const layouts = (model?.layouts || []).filter((l) => l.part);
    menu.open(e, layouts.length
      ? layouts.map((l) => ({ label: l.name || l.part, icon: 'slides', run: () => act('newSlideFrom', l) }))
      : LAYOUTS.map(([layout, label]) => ({ label, icon: 'slides', run: () => addSlide(layout) })));
  };
  // Align and Distribute's own toggle — "Align Selected Objects" (the
  // default) or "Align to Slide" — is how the ribbon remembers which one is
  // ticked; it is a preference about the next press, not part of the file.
  const [alignTo, setAlignTo] = React.useState('selection');
  const selectedShapeObj = (model?.slide?.shapes || []).find((s) => s.id === selected) || null;
  // A picture of its own (not a video, a model, an object or a cameo): Picture Format and Edit in Image.
  const pictureSelected = selectedShapeObj?.kind === 'picture' && !selectedShapeObj.media && !selectedShapeObj.model3d && !selectedShapeObj.object && !selectedShapeObj.cameo;
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
  const needShape = hasShape ? undefined : t('slides.ribbon.select-text-box');
  // A control that waits for a text box keeps its own name first, and says what it waits for after it.
  const need = (tip) => (hasShape ? tip : `${tip} ${String.fromCharCode(0x2014)} ${needShape}`);
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
        { heading: true, label: t('slides.ribbon.sequence') },
        { label: t('slides.ribbon.as-one'), icon: animation.paragraph == null ? 'check' : undefined, run: () => act('animSequence', 'object') },
        { label: t('slides.ribbon.all-at-once'), run: () => act('animSequence', 'together') },
        { label: t('slides.ribbon.by-paragraph'), icon: animation.paragraph != null ? 'check' : undefined, run: () => act('animSequence', 'paragraph') },
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
      { label: masterView ? (isLayout ? t('slides.ribbon.reset-masters') : t('slides.ribbon.reset-themes')) : t('slides.ribbon.reset'), icon: same(null) ? 'check' : undefined, run: () => act('background', { spec: null }) },
      ...(masterView ? [] : ['-', { label: t('slides.ribbon.apply-all'), run: () => act('background', { spec: own, all: true }) }]),
    ]);
  };
  const galleryAt = (kind) => (e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind, anchor: { left: r.left, bottom: r.bottom + 4 } }); };

  return (
    <Ribbon
      tabs={[
        ...(masterView ? [masterView.kind === 'notes' ? { id: 'notesMaster', label: t('slides.notes') } : masterView.kind === 'handout' ? { id: 'handoutMaster', label: t('slides.handout') } : { id: 'master', label: t('slides.slide-master') }] : []),
        { id: 'home', label: t('common.home') },
        { id: 'insert', label: t('common.insert') },
        { id: 'draw', label: t('common.draw') },
        { id: 'design', label: t('common.design') },
        { id: 'transitions', label: t('common.transitions') },
        { id: 'animations', label: t('slides.ribbon.animations') },
        { id: 'show', label: t('slides.ribbon.slide-show') },
        { id: 'record', label: t('common.record') },
        { id: 'review', label: t('common.review') },
        { id: 'view', label: t('common.view') },
        { id: 'help', label: t('common.help') },
        { id: 'pdf', label: 'PDF' }, // words-ok: a file format's name
        // Table Design, as PowerPoint's: there while a table is selected.
        ...(selectedShapeObj?.kind === 'table' ? [{ id: 'tableDesign', label: t('common.table-design'), contextual: true }] : []),
        // Shape Format, as PowerPoint's: there while a shape with words is selected.
        ...(selectedShapeObj?.kind === 'shape' && selectedShapeObj?.text && !selectedShapeObj?.cameo ? [{ id: 'shapeFormat', label: t('common.shape-format'), contextual: true }] : []),
        // Camera Format, as PowerPoint's: there while a cameo is selected.
        ...(selectedShapeObj?.cameo ? [{ id: 'cameraFormat', label: t('slides.ribbon.camera'), contextual: true }] : []),
        // 3D Model, as PowerPoint's: there while a 3D model is selected.
        ...(selectedShapeObj?.model3d ? [{ id: 'model3d', label: t('common.3d-model'), contextual: true }] : []),
        // Picture Format, as PowerPoint's: there while a picture is selected.
        ...(pictureSelected ? [{ id: 'pictureFormat', label: t('common.picture-format'), contextual: true }] : []),
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title={t('common.save-ctrl-s')} onClick={() => save(false)} />
          <Button icon="undo" title={t('common.undo-ctrl-z')} disabled={!doc?.canUndo} onClick={() => commands['edit.undo']?.run?.()} />
          <Button icon="play" title={t('slides.ribbon.start-beginning-f5')} onClick={() => act('present', 'start')} />
        </>
      }
    >
      {/* ── Home ─────────────────────────────────────────────────────────── */}
      {tab === 'home' ? (
        <>
          <Group label={t('common.clipboard')}>
            <Button tall icon="paste" label={t('common.paste')} title={canPaste ? t('slides.ribbon.paste-copied-shape') : t('slides.ribbon.paste-copy-shape')} disabled={!canPaste} onClick={() => act('pasteShape')} />
            <Rows>
              <>
                <Button icon="cut" label={t('common.cut')} title={need(t('slides.ribbon.cut-shape'))} disabled={!hasShape} onClick={() => act('cutShape')} />
                <Button icon="copy" label={t('common.copy')} title={need(t('slides.ribbon.copy-shape'))} disabled={!hasShape} onClick={() => act('copyShape')} />
              </>
              <Button
                icon="brush"
                label={t('common.format-painter')}
                pressed={painter}
                disabled={!hasShape && !painter}
                title={painter ? t('slides.ribbon.format-painter-armed') : (needShape || t('slides.ribbon.format-painter-shapes'))}
                onClick={() => act('painter')}
              />
            </Rows>
          </Group>
          <Group label={t('common.slides')}>
            <Button dropdown tall icon="plus" label={t('slides.ribbon.new-slide')} onClick={(e) => newSlideMenu(e)} />
            <Rows>
              <>
                <Button dropdown icon="grid" label={t('common.layout')} title={t('slides.ribbon.layout-put-slide')} onClick={(e) => menu.open(e, (model?.layouts || []).map((l) => ({ label: l.name || l.part, icon: l.part === model?.slide?.layout ? 'check' : undefined, run: () => act('applyLayout', l.part) })))} />
                <Button icon="undo" label={t('common.reset')} title={t('slides.ribbon.reset-placeholders-back')} onClick={() => act('resetSlide')} />
                <Button dropdown
                  icon="list"
                  label={t('slides.ribbon.section')}
                  title={t('slides.ribbon.section-section-before')}
                  onClick={(e) => menu.open(e, [
                    { label: t('slides.ribbon.add-section'), icon: 'plus', run: () => act('addSection') },
                    { label: t('slides.ribbon.rename-section'), icon: 'textbox', disabled: !inSection, run: () => act('renameSection') },
                    { label: t('slides.ribbon.remove'), icon: 'trash', disabled: !inSection, run: () => act('removeSection') },
                    { label: t('slides.ribbon.remove-all'), icon: 'close', disabled: !hasSections, run: () => act('removeAllSections') },
                  ])}
                />
              </>
              <>
                <Button icon="copy" label={t('common.duplicate')} onClick={() => commands['slide.new']?.run?.()} />
                <Button icon="trash" label={t('common.delete')} disabled={count < 2} onClick={() => commands['slide.delete']?.run?.()} />
              </>
            </Rows>
          </Group>
          <Group label={t('common.font')}>
            <Rows>
              <>
                <Select value={format.font || ''} onChange={(e) => fmt({ font: e.target.value })} style={{ width: 118 }} title={need(t('common.font'))} disabled={!hasShape}>
                  <option value="">{format.font || (format.themeFont ? t('slides.ribbon.font-theme', { font: format.themeFont }) : t('slides.ribbon.theme-font'))}</option>
                  {['Calibri', 'Calibri Light', 'Arial', 'Segoe UI', 'Georgia', 'Times New Roman', 'Verdana', 'Consolas'].filter((f) => f !== format.font).map((f) => <option key={f} value={f}>{f}</option>)}
                </Select>
                <Select value={String(size)} onChange={(e) => fmt({ size: Number(e.target.value) })} style={{ width: 58 }} title={need(t('common.font-size'))} disabled={!hasShape}>
                  {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
                  {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
                </Select>
                <Button icon="chevronUp" title={need(t('common.increase'))} disabled={!hasShape} onClick={() => fmt({ size: nearer(1) })} />
                <Button icon="chevronDown" title={need(t('common.decrease'))} disabled={!hasShape} onClick={() => fmt({ size: nearer(-1) })} />
                <Button icon="undo" title={needShape || t('slides.ribbon.clear-all-formatting')} disabled={!hasShape} onClick={() => act('clearFormat')} />
              </>
              <>
                <Button icon="bold" title={need(t('common.bold-ctrl-b'))} pressed={format.bold} disabled={!hasShape} onClick={() => fmt({ bold: 'toggle' })} />
                <Button icon="italic" title={need(t('common.italic-ctrl-i'))} pressed={format.italic} disabled={!hasShape} onClick={() => fmt({ italic: 'toggle' })} />
                <Button icon="underline" title={need(t('common.underline-ctrl-u'))} pressed={format.underline} disabled={!hasShape} onClick={() => fmt({ underline: 'toggle' })} />
                <Button icon="strike" title={need(t('common.strikethrough'))} pressed={format.strike} disabled={!hasShape} onClick={() => fmt({ strike: 'toggle' })} />
                <Button dropdown icon="textbox" label="AV" title={need((format.spacing ? t('slides.ribbon.character-spacing-now', { spacing: format.spacing }) : t('slides.ribbon.character')))} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('slides.ribbon.very-tight'), -1.5], [t('common.tight'), -0.75], [t('common.normal'), 0], [t('slides.ribbon.loose'), 1.5], [t('slides.ribbon.very-loose'), 3]].map(([label, v]) => ({ label, icon: (format.spacing || 0) === v ? 'check' : undefined, run: () => fmt({ spacing: v }) })))} />{/* words-ok: AV is the control's picture of letter spacing, as PowerPoint draws it */}
                <Button dropdown icon="textbox" label="Aa" title={need(t('common.change-case'))} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('slides.ribbon.sentence-case'), 'sentence'], [t('common.lowercase-2'), 'lower'], [t('common.uppercase'), 'upper'], [t('common.capitalise-each-word'), 'title']].map(([label, mode]) => ({ label, run: () => fmt({ case: mode }) })))} />{/* words-ok: Aa is the control's picture of letter case, as PowerPoint draws it */}
                <Separator />
                {/* PowerPoint's two split buttons: the face puts on the colour last chosen, the arrow opens the palette. */}
                <ColourSplit kind="slides.highlight" icon="highlighter" name={t('common.text-highlight')} highlight first="yellow" disabled={!hasShape} explain={need}
                  current={HIGHLIGHT.find(([, hex]) => sameColour(hex, format.highlight))?.[0] ?? null} pressed={Boolean(format.highlight)}
                  empty={{ label: t('common.no-highlight'), value: null }} onPick={(name) => fmt({ highlight: name ? `#${HIGHLIGHT.find(([k]) => k === name)[1]}` : null })} />
                <ColourSplit kind="slides.font" icon="fontColour" name={t('common.font-colour')} first="FF0000" disabled={!hasShape} explain={need} theme={model?.slide?.theme?.colors} current={format.color ?? null}
                  empty={{ label: t('common.automatic'), value: null, swatch: '#000000' }} onPick={(value) => fmt({ color: value ? `#${value.toLowerCase()}` : null })} />
              </>
            </Rows>
          </Group>
          <Group label={t('common.paragraph')}>
            <Rows>
              <>
                <Button icon="listBullet" title={need(t('common.bullets'))} pressed={format.bullet === 'char'} disabled={!hasShape} onClick={() => fmt({ bullet: format.bullet === 'char' ? 'none' : 'char' })} />
                <Button icon="listNumber" title={need(t('common.numbering'))} pressed={format.bullet === 'number'} disabled={!hasShape} onClick={() => fmt({ bullet: format.bullet === 'number' ? 'none' : 'number' })} />
                <Button icon="chevronLeft" title={need(t('slides.ribbon.decrease'))} disabled={!hasShape || !format.level} onClick={() => fmt({ level: -1 })} />
                <Button icon="chevronRight" title={need(t('slides.ribbon.increase'))} disabled={!hasShape} onClick={() => fmt({ level: 1 })} />
                <Button dropdown icon="list" title={need((format.lineHeight ? t('slides.ribbon.line-spacing-now', { value: String(format.lineHeight) }) : t('slides.ribbon.line-spacing-now-layout')))} disabled={!hasShape} onClick={(e) => menu.open(e, [1, 1.15, 1.5, 2].map((v) => ({ label: String(v), icon: format.lineHeight === v ? 'check' : undefined, run: () => fmt({ lineHeight: v }) })))} />
                <Separator />
                <Button dropdown icon="grid" title={need(t('slides.ribbon.columns-now', { columns: format.columns || 1 }))} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('slides.ribbon.one-column'), 1], [t('slides.ribbon.two-columns'), 2], [t('slides.ribbon.three-columns'), 3]].map(([label, n]) => ({ label, icon: (format.columns || 1) === n ? 'check' : undefined, run: () => act('body', { columns: n }) })))} />
                <Button dropdown icon="rotate" title={need(t('slides.ribbon.text-direction-now', { direction: ({ horz: t('slides.ribbon.horizontal'), vert270: t('slides.ribbon.rotated-up'), vert: t('slides.ribbon.rotated-down'), eaVert: t('slides.ribbon.stacked') })[format.vert] || t('slides.ribbon.horizontal') }))} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('common.horizontal'), 'horz'], [t('slides.ribbon.rotate-all-text-90'), 'vert'], [t('slides.ribbon.rotate-all-text'), 'vert270']].map(([label, v]) => ({ label, icon: (format.vert || 'horz') === v ? 'check' : undefined, run: () => act('body', { vert: v }) })))} />
                <Button dropdown icon="chevronUp" title={need(t('slides.ribbon.align-text-now', { anchor: ({ top: t('common.top-2'), middle: t('slides.ribbon.middle'), bottom: t('common.bottom-2') })[format.anchor || 'top'] || format.anchor }))} disabled={!hasShape} onClick={(e) => menu.open(e, [[t('common.top'), 'top'], [t('common.middle'), 'middle'], [t('common.bottom'), 'bottom']].map(([label, a]) => ({ label, icon: (format.anchor || 'top') === a ? 'check' : undefined, run: () => act('body', { anchor: a }) })))} />
              </>
              <>
                <Button icon="alignLeft" title={need(t('common.align-left-2'))} pressed={format.align === 'left'} disabled={!hasShape} onClick={() => fmt({ align: 'left' })} />
                <Button icon="alignCenter" title={need(t('common.centre'))} pressed={format.align === 'center'} disabled={!hasShape} onClick={() => fmt({ align: 'center' })} />
                <Button icon="alignRight" title={need(t('common.align-right-2'))} pressed={format.align === 'right'} disabled={!hasShape} onClick={() => fmt({ align: 'right' })} />
                <Button icon="alignJustify" title={need(t('common.justify'))} pressed={format.align === 'justify'} disabled={!hasShape} onClick={() => fmt({ align: 'justify' })} />
                <Button icon="textLtr" title={need(t('common.left-right'))} pressed={hasShape && !format.rtl} disabled={!hasShape} onClick={() => fmt({ rtl: false })} />
                <Button icon="textRtl" title={needShape || t('slides.ribbon.right-left-text')} pressed={Boolean(format.rtl)} disabled={!hasShape} onClick={() => fmt({ rtl: true })} />
                <Separator />
                <Button icon="shape" label={t('common.smart-art')} title={needShape || t('slides.ribbon.convert-smart-art')} disabled={!hasShape} onClick={() => act('convertSmartArt')} />
              </>
            </Rows>
          </Group>
          <Group label={t('common.drawing')}>
            <Button dropdown tall icon="shape" label={t('common.shapes')} title={t('slides.ribbon.shapes-rectangle-oval')} onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => act('addShape', preset) })))} />

            <Button dropdown tall icon="grid" label={t('common.arrange')} onClick={(e) => menu.open(e, [
              { label: t('slides.bring-to-front'), icon: 'chevronUp', run: () => act('order', 'front') },
              { label: t('common.bring-forward'), run: () => act('order', 'forward') },
              { label: t('common.send-backward-2'), run: () => act('order', 'backward') },
              { label: t('slides.send-to-back'), icon: 'chevronDown', run: () => act('order', 'back') },
              '-',
              { label: t('common.group'), icon: 'grid', disabled: multiCount < 2, title: multiCount < 2 ? t('slides.select-two-more') : undefined, run: () => act('group') },
              { label: t('common.ungroup'), icon: 'grid', disabled: !isGroup, title: !isGroup ? t('common.select-group') : undefined, run: () => act('ungroup') },
              '-',
              { label: t('common.align-left'), run: () => act('align', { edge: 'left', to: alignTo }) },
              { label: t('common.align-center'), run: () => act('align', { edge: 'center', to: alignTo }) },
              { label: t('common.align-right'), run: () => act('align', { edge: 'right', to: alignTo }) },
              { label: t('common.align-top'), run: () => act('align', { edge: 'top', to: alignTo }) },
              { label: t('common.align-middle'), run: () => act('align', { edge: 'middle', to: alignTo }) },
              { label: t('common.align-bottom'), run: () => act('align', { edge: 'bottom', to: alignTo }) },
              { label: t('common.distribute-horizontally'), disabled: multiCount < 3, title: multiCount < 3 ? t('slides.ribbon.select-three-more') : undefined, run: () => act('distribute', { axis: 'horizontal', to: alignTo }) },
              { label: t('common.distribute'), disabled: multiCount < 3, title: multiCount < 3 ? t('slides.ribbon.select-three-more') : undefined, run: () => act('distribute', { axis: 'vertical', to: alignTo }) },
              '-',
              { label: t('slides.ribbon.align-to'), icon: alignTo === 'slide' ? 'check' : undefined, run: () => setAlignTo('slide') },
              { label: t('common.align-selected'), icon: alignTo === 'selection' ? 'check' : undefined, run: () => setAlignTo('selection') },
              '-',
              { label: t('common.rotate-right-90'), run: () => act('rotateBy', 90) },
              { label: t('common.rotate-left-90'), run: () => act('rotateBy', -90) },
              { label: t('common.flip'), run: () => act('flipShape', 'vertical') },
              { label: t('common.flip-horizontal'), run: () => act('flipShape', 'horizontal') },
              { label: t('common.more-rotation'), run: () => act('rotationDialog') },
              '-',
              { label: t('slides.ribbon.selection'), icon: 'list', run: () => act('pane', 'layers') },
              '-',
              { label: t('slides.delete-shape'), icon: 'trash', run: () => act('deleteShape') },
              { label: t('slides.ribbon.move-up'), run: () => act('nudge', { dy: -8 }) },
              { label: t('slides.ribbon.move-down'), run: () => act('nudge', { dy: 8 }) },
              { label: t('slides.ribbon.move-left'), run: () => act('nudge', { dx: -8 }) },
              { label: t('slides.ribbon.move-right'), run: () => act('nudge', { dx: 8 }) },
            ])} disabled={!hasShape} title={need(t('slides.ribbon.arrange-selected-shape'))} />
            <Button dropdown tall icon="wand" label={t('slides.ribbon.quick-styles')} disabled={!hasShape} title={need(t('slides.ribbon.themes-own-looks'))} onClick={(e) => menu.open(e, [1, 2, 3, 4, 5, 6].map((n) => ({ label: t('common.accent-n', { n }), icon: 'shape', run: () => act('quickStyle', n) })))} />
            <Button dropdown icon="wand" label={t('common.shape-fill')} disabled={!hasShape} title={need(t('slides.ribbon.shape-fill-colour'))} onClick={(e) => menu.open(e, [
              ...COLOURS.map(([value, label]) => ({ label, preview: swatch(value), run: () => act('shapeFill', { color: value, alpha: 1 }) })),
              '-',
              { label: t('common.no-fill'), icon: 'close', run: () => act('shapeFill', 'none') },
              { label: t('slides.ribbon.more-fill'), run: () => act('formatPane') },
            ])} />
            <Button dropdown icon="shape" label={t('common.shape-outline')} disabled={!hasShape} title={need(t('slides.ribbon.shape-outline-colour'))} onClick={(e) => menu.open(e, [
              ...COLOURS.map(([value, label]) => ({ label, preview: swatch(value), run: () => act('shapeLine', { color: value }) })),
              '-',
              { label: t('common.no-outline'), icon: 'close', run: () => act('shapeLine', 'none') },
              { label: t('slides.ribbon.more-outline'), run: () => act('formatPane') },
            ])} />
            <Button dropdown icon="wand" label={t('slides.ribbon.shape')} disabled={!hasShape} title={need(t('slides.ribbon.shape-effects-shadow'))} onClick={(e) => menu.open(e, [
              ...SHADOW_MENU.map(([key, label]) => ({ label, icon: key === 'none' ? 'close' : undefined, run: () => act('shapeShadow', key) })),
              '-',
              ...GLOW_MENU.map(([pt, label]) => ({ label, icon: pt === null ? 'close' : undefined, run: () => act('shapeEffects', { glow: pt === null ? null : { radius: pt, color: selectedShapeObj?.effects?.glow?.color } }) })),
              '-',
              ...SOFTEDGE_MENU.map(([pt, label]) => ({ label, icon: pt === null ? 'close' : undefined, run: () => act('shapeEffects', { softEdge: pt === null ? null : { radius: pt } }) })),
              '-',
              ...REFLECTION_MENU.map(([key, label]) => ({ label, icon: key === null ? 'close' : undefined, run: () => act('shapeEffects', { reflection: key }) })),
            ])} />
          </Group>
          <Group label={t('slides.ribbon.edit-shape')}>
            <Button dropdown tall icon="shape" label={t('slides.ribbon.edit-shape')} disabled={!canEditPoints} title={canEditPoints ? t('slides.ribbon.edit-points-drag') : hasShape ? t('slides.ribbon.picture-chart-table') : needShape} onClick={(e) => menu.open(e, [
              { label: t('slides.edit-points'), icon: 'shape', run: () => act('editPoints') },
            ])} />
          </Group>
          <Group label={t('common.editing')}>
            <Button tall icon="find" label={t('common.find')} title={t('slides.ribbon.find-words-every')} onClick={() => act('find')} />
            <Button tall icon="replace" label={t('common.replace')} title={t('slides.ribbon.replace-find-replace')} onClick={() => act('replace')} />
          </Group>
          <Group label={t('slides.ribbon.designer')}>
            <Button tall icon="wand" label={t('slides.design-ideas')} pressed={view?.pane === 'ideas'} title={t('slides.ribbon.design-ideas-layouts')} onClick={() => act('pane', 'ideas')} />
          </Group>
        </>
      ) : null}

      {/* ── Insert ───────────────────────────────────────────────────────── */}
      {/* ── Slide Master (only while the master and its layouts are on the stage) ── */}
      {tab === 'master' && masterView ? (
        <>
          <Group label={t('slides.ribbon.edit')}>
            <Button tall icon="slides" label={t('slides.ribbon.insert-slide')} title={t('slides.ribbon.insert-slide-master')} onClick={() => act('insertMaster')} />
            <Button tall icon="plus" label={t('slides.insert')} title={t('slides.ribbon.insert-layout-new')} onClick={() => act('insertLayout')} />
            <Rows>
              <Button icon="trash" label={t('common.delete')} disabled={!isLayout || masterItem?.used > 0} title={!isLayout ? t('slides.ribbon.delete-pick-layout') : masterItem?.used ? tn(masterItem.used, 'slides.ribbon.delete-slide-uses') : t('slides.ribbon.delete-layout-no')} onClick={() => act('deleteLayout')} />
              <Button icon="textbox" label={t('common.rename')} title={isLayout ? t('slides.ribbon.rename-layouts') : t('slides.ribbon.rename-masters')} onClick={() => act('renameLayout')} />
              <Button icon="lock" label={t('slides.ribbon.preserve')} pressed={Boolean(masterOfItem?.preserve)} title={t('slides.ribbon.preserve-keep-master')} onClick={() => act('preserve')} />
            </Rows>
          </Group>
          <Group label={t('slides.ribbon.master')}>
            <Button dropdown tall icon="grid" label={t('slides.ribbon.insert')} disabled={!isLayout} title={isLayout ? t('slides.ribbon.insert-placeholder') : t('slides.ribbon.insert-placeholder-pick')} onClick={(e) => menu.open(e, [['content', t('slides.ribbon.content')], ['text', t('common.text')], ['picture', t('common.picture')]].map(([kind, label]) => ({ label, icon: kind === 'picture' ? 'picture' : 'textbox', run: () => act('insertPlaceholder', kind) })))} />
            <Rows>
              <label className="sl-rb-field sl-rb-check" data-tip={t('slides.ribbon.title-layouts-title')}><input type="checkbox" className="sl-master-title" checked={Boolean(masterItem?.hasTitle)} onChange={(e) => act('masterPlaceholders', { title: e.target.checked })} /> {t('common.title')}</label>
              <label className="sl-rb-field sl-rb-check" data-tip={t('slides.ribbon.footers-date-footer')}><input type="checkbox" className="sl-master-footers" checked={Boolean(masterItem?.hasFooters)} onChange={(e) => act('masterPlaceholders', { footers: e.target.checked })} /> {t('common.footers')}</label>
            </Rows>
          </Group>
          <Group label={t('slides.ribbon.edit-theme')}>
            <Button tall icon="wand" label={t('common.themes')} title={t('slides.ribbon.themes-every-theme')} onClick={galleryAt('themes')} />
            <Rows>
              <Button icon="contrast" label={t('common.colours')} title={t('slides.ribbon.colours-themes-twelve')} onClick={galleryAt('colours')} />
              <Button icon="textbox" label={t('common.fonts')} title={t('slides.ribbon.fonts-themes-heading')} onClick={galleryAt('fonts')} />
              <Button icon="wand" label={t('common.effects')} title={t('slides.ribbon.effects-themes-format')} onClick={galleryAt('effects')} />
            </Rows>
          </Group>
          <Group label={t('common.background')}>
            <Button tall icon="picture" label={t('slides.ribbon.background')} title={isLayout ? t('slides.ribbon.background-styles-layouts') : t('slides.ribbon.background-styles')} onClick={backgroundMenu} />
            <Rows>
              <label className="sl-rb-field sl-rb-check" data-tip={isLayout ? t('slides.ribbon.hide-background-graphics') : t('slides.ribbon.hide-background-graphics-pick')}><input type="checkbox" className="sl-master-hidebg" disabled={!isLayout} checked={Boolean(masterItem?.hidesBackgroundGraphics)} onChange={(e) => act('hideBackgroundGraphics', e.target.checked)} /> {t('slides.ribbon.hide-background')}</label>
            </Rows>
          </Group>
          <Group label={t('common.close')}>
            <Button tall icon="close" label={t('slides.ribbon.close-master')} title={t('slides.ribbon.close-master-view')} onClick={() => act('closeMaster')} />
          </Group>
        </>
      ) : null}

      {(tab === 'notesMaster' || tab === 'handoutMaster') && (masterView?.kind === 'notes' || masterView?.kind === 'handout') ? (
        <>
          <Group label={t('common.page-setup')}>
            <Button dropdown tall icon="file" label={masterView.kind === 'notes' ? t('slides.ribbon.notes-page-orientation') : t('slides.ribbon.handout')} title={t('slides.ribbon.orientation-printed-page')} onClick={(e) => menu.open(e, [
              { label: t('common.portrait'), icon: masterView.portrait ? 'check' : undefined, run: () => act('notesOrientation', 'portrait') },
              { label: t('common.landscape'), icon: masterView.portrait ? undefined : 'check', run: () => act('notesOrientation', 'landscape') },
            ])} />
            {masterView.kind === 'handout' ? (
              <Button dropdown tall icon="grid" label={t('slides.ribbon.slides-per')} title={t('slides.ribbon.slides-per-page')} onClick={(e) => menu.open(e, [1, 2, 3, 4, 6, 9].map((n) => ({ label: tn(n, 'slides.ribbon.count-slide'), icon: (view.handoutPer || 6) === n ? 'check' : undefined, run: () => act('handoutPer', n) })))} />
            ) : null}
          </Group>
          <Group label={t('common.placeholders')}>
            <Rows>
              {(masterView.kind === 'notes'
                ? [['hdr', t('common.header')], ['sldImg', t('slides.ribbon.slide-image')], ['ftr', t('common.footer')], ['dt', t('common.date')], ['body', t('common.body')], ['sldNum', t('common.page-number')]]
                : [['hdr', t('common.header')], ['ftr', t('common.footer')], ['dt', t('common.date')], ['sldNum', t('common.page-number')]]
              ).map(([type, label]) => (
                <label key={type} className="sl-rb-field sl-rb-check" data-tip={t('slides.ribbon.label-placeholder-master', { label })}><input type="checkbox" className={`sl-nm-ph sl-nm-${type}`} checked={Boolean(masterView.placeholders?.[type])} onChange={(e) => act('masterPlaceholder', { type, on: e.target.checked })} /> {label}</label>
              ))}
            </Rows>
          </Group>
          <Group label={t('common.close')}>
            <Button tall icon="close" label={t('slides.ribbon.close-master')} title={t('slides.ribbon.close-master-view')} onClick={() => act('closeMaster')} />
          </Group>
        </>
      ) : null}

      {tab === 'insert' ? (
        <>
          <Group label={t('common.slides')}>
            <Button dropdown tall icon="plus" label={t('slides.ribbon.new-slide')} onClick={(e) => newSlideMenu(e)} />
          </Group>
          <Group label={t('common.tables')}>
            <Button dropdown tall icon="table" label={t('common.table')} title={t('slides.ribbon.table-grid-cells')} onClick={(e) => menu.open(e, TABLE_SIZES.map(([rows, cols, label]) => ({ label, icon: 'table', run: () => act('addTable', { rows, cols }) })))} />
          </Group>
          <Group label={t('common.images')}>
            <Button tall icon="picture" label={t('common.pictures')} title={t('slides.ribbon.pictures-picture-device')} onClick={insertPicture} />

            <Button tall icon="picture" label={t('common.screenshot')} title={t('slides.ribbon.screenshot-picture')} onClick={() => act('screenshot')} />
            <Button tall icon="picture" label={t('slides.photo-album')} title={t('slides.ribbon.photo-album-new')} onClick={() => act('photoAlbum')} />
          </Group>
          <Group label={t('common.camera')}>
            <Button tall icon="video" label={t('slides.ribbon.cameo')} className="sl-cameo-insert" title={t('slides.ribbon.cameo-camera-live-shape')} onClick={() => act('cameo')} />
          </Group>
          <Group label={t('common.illustrations')}>
            <Button dropdown tall icon="shape" label={t('common.shapes')} title={t('slides.ribbon.shapes-rectangle-oval')} onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => act('addShape', preset) })))} />

            <Button tall icon="qrcode" label={t('common.qr-code')} className="sl-qr-insert" title={t('slides.ribbon.qr-code-code')} onClick={() => act('qrcode')} />
            <Button tall icon="star" label={t('common.icons')} title={t('slides.ribbon.icons-one-suites')} onClick={() => act('icons')} />
            <Button tall icon="shape" label={t('common.3d-models')} className="sl-model3d-insert" title={t('slides.ribbon.3d-models-model')} onClick={() => act('model3d')} />
            <Button tall icon="shape" label={t('common.smart-art')} title={t('slides.ribbon.smart-art-list')} onClick={() => act('insertSmartArt')} />
            <Button dropdown tall icon="chart" label={t('common.chart')} title={t('slides.ribbon.chart-sample-chart')} onClick={(e) => menu.open(e, CHART_TYPES.map(([type, label]) => ({ label, icon: 'chart', run: () => act('addChart', { type }) })))} />
          </Group>
          <Group label={t('common.links')}>
            <Button dropdown tall icon="zoomIn" label={t('common.zoom')} title={t('slides.ribbon.zoom-picture-slide')} onClick={(e) => menu.open(e, [{ label: t('slides.ribbon.summary-zoom'), icon: 'grid', run: () => act('insertZoom', 'summary') }, { label: t('slides.ribbon.section-zoom'), icon: 'list', disabled: !(model?.sections || []).length, run: () => act('insertZoom', 'section') }, { label: t('slides.ribbon.slide-zoom'), icon: 'slides', run: () => act('insertZoom', 'slide') }])} />
            <Button tall icon="link" label={t('common.link')} disabled={!hasShape} title={needShape || t('slides.ribbon.link-web-address')} onClick={() => act('link')} />
            <Button tall icon="play" label={t('slides.ribbon.action')} title={t('slides.ribbon.action-click-selected')} onClick={() => act('action')} />
          </Group>
          <Group label={t('common.comments')}>
            <Button tall icon="comment" label={t('common.comment')} title={t('slides.ribbon.comment-new-comment')} onClick={() => act('newComment')} />
          </Group>
          <Group label={t('common.text')}>
            <Button tall icon="textbox" label={t('common.text-box')} onClick={() => commands['slide.textbox']?.run?.()} />
            <Button tall icon="header" label={t('common.header-and')} title={t('slides.ribbon.header-footer-date')} onClick={() => act('footer')} />
            <Button dropdown tall icon="wordArt" label={t('common.word-art')} title={t('common.word-art-big')} onClick={(e) => menu.open(e, wordArtMenu((style) => act('wordArt', style)))} />
            <Button icon="clock" label={t('common.date-and')} title={t('slides.ribbon.date-time-date')} onClick={() => act('footer', 'date')} />
            <Button icon="list" label={t('slides.ribbon.slide-number')} title={t('slides.ribbon.slide-number-slides')} onClick={() => act('footer', 'number')} />
            <Button icon="file" label={t('common.object')} title={t('slides.ribbon.object-word-excel')} onClick={() => act('insertObject')} />
          </Group>
          <Group label={t('common.symbols')}>
            <Button tall icon="formula" label={t('common.equation')} title={t('slides.ribbon.equation-office-math')} onClick={() => act('equation')} />
            <Button tall icon="star" label={t('common.symbol')} title={t('slides.ribbon.symbol-character-symbol')} onClick={() => act('symbol')} />
          </Group>
          <Group label={t('common.media')}>
            <Button dropdown tall icon="video" label={t('common.video')} title={t('slides.ribbon.video-video-computer')} onClick={(e) => menu.open(e, [{ label: t('slides.ribbon.this-device'), icon: 'video', run: () => act('insertMedia', 'video') }])} />
            <Button dropdown tall icon="volume" label={t('common.audio')} title={t('slides.ribbon.audio-sound-computer')} onClick={(e) => menu.open(e, [{ label: t('slides.ribbon.audio-pc'), icon: 'volume', run: () => act('insertMedia', 'audio') }])} />
            <Button tall icon="video" label={t('common.screen-recording')} title={t('slides.ribbon.screen-recording-record')} onClick={() => act('screenRecording')} />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label={t('common.drawing-tools')}>
            <Button tall icon="mouse" label={t('common.select')} pressed={!ink?.tool} title={t('slides.ribbon.select-put-pen')} onClick={() => act('inkTool', null)} />
            <Button tall icon="lasso" label={t('common.lasso')} pressed={ink?.tool === 'lasso'} title={t('common.lasso-select-draw')} onClick={() => act('inkTool', 'lasso')} />
            <Button tall icon="eraser" label={t('common.eraser')} pressed={ink?.tool === 'eraser'} title={t('common.eraser-take-away')} onClick={() => act('inkTool', 'eraser')} />
            {(ink?.pens || []).map((p) => {
              const on = ink.tool === 'pen' && ink.penId === p.id;
              const label = { pen: t('common.pen'), pencil: t('common.pencil'), highlighter: t('common.highlighter') }[p.tool];
              return (
                <Button key={p.id} tall icon="wand" label={label} pressed={on} className={`sl-pen sl-pen-${p.tool}`} data-pen={p.id}
                  style={{ '--pen': p.color }}
                  title={on ? t('common.label-click-again', { label }) : t('slides.ribbon.label-draw-slide', { label })}
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
            <Button dropdown tall icon="plus" label={t('common.add')} title={t('slides.ribbon.add-pen-another')} onClick={(e) => menu.open(e, [['pen', t('common.pen')], ['pencil', t('common.pencil')], ['highlighter', t('common.highlighter')]].map(([tool, label]) => ({ label, icon: 'plus', run: () => act('inkAdd', tool) })))} />
          </Group>
          <Group label={t('common.touch')}>
            <Button tall icon="wand" label={t('common.draw-with')} pressed={ink?.touch !== false} title={t('common.draw-touch-finger')} onClick={() => act('inkTouch')} />
          </Group>
          <Group label={t('common.stencils')}>
            <Button tall icon="minus" label={t('common.ruler')} pressed={Boolean(ink?.ruler)} title={t('slides.ribbon.ruler-straight-edge')} onClick={() => act('inkRuler')} />
          </Group>
          <Group label={t('common.convert')}>
            <Button tall icon="shape" label={t('common.ink-to-shape')} pressed={Boolean(ink?.toShape)} title={t('common.ink-shape-rectangle')} onClick={() => act('inkToShape')} />
          </Group>
          <Group label={t('slides.ribbon.replay')}>
            <Button tall icon="play" label={t('slides.ribbon.ink-replay')} title={t('slides.ribbon.ink-replay-slides')} onClick={() => act('inkReplay')} />
          </Group>
          <Group label={t('common.help')}>
            <Button tall icon="info" label={t('slides.ribbon.ink-help')} onClick={() => act('help')} />
          </Group>
        </>
      ) : null}

      {/* ── Design ───────────────────────────────────────────────────────── */}
      {tab === 'design' ? (
        <>
          <Group label={t('common.themes')} icon="theme">
            <RibbonStrip
              kind="themes"
              items={designStrip?.themes}
              count={4}
              label={t('slides.ribbon.drawing-this-slide')}
              onPick={(it) => act('applyTheme', it)}
              onMore={(e) => { const r = e.currentTarget.closest('.sl-rs').getBoundingClientRect(); act('designGallery', { kind: 'themes', anchor: { left: r.left, bottom: r.bottom + 4 } }); }}
            />
          </Group>
          <Group label={t('common.variants')} icon="variants">
            <RibbonStrip
              kind="variants"
              items={designStrip?.variants}
              count={4}
              label={t('slides.ribbon.drawing-this-slide')}
              onPick={(it) => act('applyVariant', it.id)}
              onMore={(e) => { const r = e.currentTarget.closest('.sl-rs').getBoundingClientRect(); act('designGallery', { kind: 'variants', anchor: { left: r.left, bottom: r.bottom + 4 } }); }}
            />
            <Rows>
              <Button icon="contrast" label={t('common.colours')} title={t('slides.ribbon.colours-themes-twelve-colours')} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind: 'colours', anchor: { left: r.left, bottom: r.bottom + 4 } }); }} />
              <Button icon="textbox" label={t('common.fonts')} title={t('slides.ribbon.fonts-themes-heading-body')} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind: 'fonts', anchor: { left: r.left, bottom: r.bottom + 4 } }); }} />
              <Button icon="wand" label={t('common.effects')} title={t('slides.ribbon.effects-how-theme')} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); act('designGallery', { kind: 'effects', anchor: { left: r.left, bottom: r.bottom + 4 } }); }} />
            </Rows>
          </Group>
          <Group label={t('slides.ribbon.customise')}>
            <Button
              tall
              icon="picture"
              label={t('slides.ribbon.background')}
              title={t('slides.ribbon.background-styles-solid')}
              onClick={backgroundMenu}
            />
          </Group>
          <Group label={t('common.slides')}>
            <Button tall icon="grid" label={t('slides.ribbon.layouts')} pressed={view.pane === 'designs'} title={t('slides.ribbon.decks-layouts-pane')} onClick={() => act('pane', 'designs')} />
            <Button dropdown tall icon="slides" label={t('slides.ribbon.new-slide')} title={t('slides.ribbon.new-slide-one')} onClick={(e) => newSlideMenu(e)} />
            <Button
              dropdown
              tall
              icon="grid"
              label={t('slides.slide-size')}
              className="sl-slide-size"
              title={model?.size ? t('slides.ribbon.slide-size-slides', { width: cmOf(model.size.cx), height: cmOf(model.size.cy), ratio: ratioOf(model.size) }) : t('slides.ribbon.slide-size-standard')}
              onClick={(e) => menu.open(e, [
                { label: t('slides.standard-4-3'), icon: ratioOf(model?.size) === t('slides.ribbon.4-3-standard') ? 'check' : 'slides', run: () => act('slideSize', { cx: 9144000, cy: 6858000 }) },
                { label: t('slides.ribbon.widescreen-16-9'), icon: ratioOf(model?.size) === t('slides.ribbon.16-9-widescreen') ? 'check' : 'slides', run: () => act('slideSize', { cx: 12192000, cy: 6858000 }) },
                '-',
                { label: t('slides.ribbon.custom-slide-size'), icon: 'settings', run: () => act('slideSize', null) },
              ])}
            />
          </Group>
          <Group label={t('common.edit')}>
            <Button tall icon="wand" label={t('slides.design-ideas')} pressed={view?.pane === 'ideas'} title={t('slides.ribbon.design-ideas-layouts')} onClick={() => act('pane', 'ideas')} />
          </Group>
        </>
      ) : null}

      {/* ── Transitions ──────────────────────────────────────────────────── */}
      {tab === 'transitions' ? (
        <>
          <Group label={t('common.preview')}>
            <Button
              tall
              icon="play"
              label={t('common.preview')}
              disabled={!transitionPlays}
              title={transitionPlays ? t('slides.ribbon.preview-play-slides-transition', { transition: describeTransition(transition) }) : t('slides.ribbon.preview-slide-no')}
              onClick={() => act('preview', 'transition')}
            />
          </Group>
          <Group label={t('slides.ribbon.transition-slide')}>
            {TRANSITION_GALLERY.map(([key, label, blurb]) => (
              <Button
                key={key}
                tall
                icon={TRANSITION_ICONS[key]}
                label={label}
                className="sl-tr-pick"
                data-transition={key}
                pressed={transitionKey === key}
                title={t('slides.ribbon.label-blurb', { label, blurb })}
                onClick={() => act('transition', key === 'none' ? { type: 'none' } : { type: key === 'shape' ? 'circle' : key })}
              />
            ))}
            <Button dropdown
              tall
              icon="settings"
              label={t('slides.ribbon.effect-options')}
              disabled={!transitionOptions}
              title={transitionOptions ? t('slides.ribbon.effect-options-way', { name: TRANSITION_GALLERY.find(([k]) => k === transitionKey)?.[1] || t('slides.ribbon.transition') }) : t('slides.ribbon.effect-options-pick')}
              onClick={(e) => menu.open(e, transitionOptions.map(([value, label]) => ({
                label,
                icon: optionOf(transition) === value ? 'check' : undefined,
                run: () => act('transition', transitionKey === 'shape' ? { type: value.split(':')[0], direction: value.split(':')[1] ?? null } : { direction: value }),
              })))}
            />
          </Group>
          <Group label={t('common.timing')}>
            <Rows>
              <Button icon="volume" label={t('common.sound')} title={t('slides.ribbon.sound-plays-slide')} onClick={(e) => {
                const s = model?.slide?.transitionSound || null;
                const tick = (on) => (on ? 'check' : undefined);
                menu.open(e, [
                  { label: t('slides.ribbon.no-sound'), icon: tick(!s), run: () => act('transitionSound', { none: true }) },
                  { label: t('slides.ribbon.stop-previous-sound'), icon: tick(Boolean(s?.stop)), run: () => act('transitionSound', { stop: true }) },
                  '-',
                  ...SOUNDS.map((name) => ({ label: name, icon: tick(s?.name === soundFile(name)) || 'volume', run: () => act('transitionSound', { builtin: name }) })),
                  '-',
                  { label: t('slides.ribbon.other-sound'), icon: 'folder', run: () => act('transitionSound', { other: true }) },
                  '-',
                  { label: t('slides.ribbon.loop-until'), icon: tick(Boolean(s?.loop)), disabled: !s || s.stop, run: () => act('transitionSound', { loop: !s?.loop }) },
                ]);
              }} />
              <div className="sl-rb-field" data-tip={t('slides.ribbon.duration-how-long-transition')}>
                <Icon name="clock" size={15} />
                <span>{t('common.duration')}</span>
                <SecondsField className="sl-tr-duration" label={t('common.duration')} value={transition && transitionKey !== 'none' ? transition.duration : null} disabled={!transition || transitionKey === 'none'} min={0.01} max={59} onCommit={(v) => act('transition', { duration: v })} />
              </div>
              <Button icon="check" label={t('slides.ribbon.apply-to-all')} disabled={count < 2} title={count < 2 ? t('slides.ribbon.apply-all-deck') : t('slides.ribbon.apply-all-slides')} onClick={() => act('transitionAll')} />
            </Rows>
            <Separator />
            <Rows>
              <div className="sl-rb-caption">{t('slides.ribbon.advance-slide')}</div>
              <label className="sl-rb-field" data-tip={t('slides.ribbon.mouse-click-click')}>
                <input type="checkbox" className="sl-tr-onclick" checked={transition ? transition.advanceOnClick !== false : true} onChange={(e) => act('transition', { advanceOnClick: e.target.checked })} />
                <span>{t('slides.ribbon.on-mouse-click')}</span>
              </label>
              <div className="sl-rb-field" data-tip={t('slides.ribbon.after-show-moves')}>
                <label className="sl-rb-check">
                  <input type="checkbox" className="sl-tr-after-on" checked={transition?.advanceAfter != null} onChange={(e) => act('transition', { advanceAfter: e.target.checked ? (transition?.advanceAfter ?? 0) : null })} />
                  <span>{t('slides.ribbon.after')}</span>
                </label>
                <SecondsField className="sl-tr-after" label={t('slides.ribbon.after')} value={transition?.advanceAfter ?? 0} min={0} max={3600} onCommit={(v) => act('transition', { advanceAfter: v })} />
              </div>
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Animations ───────────────────────────────────────────────────── */}
      {tab === 'animations' ? (
        <>
          <Group label={t('common.preview')}>
            <Button
              tall
              icon="play"
              label={t('common.preview')}
              disabled={!animationList.length}
              title={animationList.length ? tn(animationList.length, 'slides.ribbon.preview-play-slides') : t('slides.ribbon.preview-slide-no-animations')}
              onClick={() => act('preview', 'animation')}
            />
          </Group>
          <Group label={t('common.animation')}>
            <Button
              tall
              icon="star"
              label={t('common.none')}
              className="sl-an-pick sl-an-none"
              data-effect="none"
              pressed={hasShape && !shapeAnimated}
              disabled={!hasShape}
              title={hasShape ? t('slides.ribbon.none-take-every') : t('slides.ribbon.none-select-shape')}
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
                title={hasShape ? (animation ? t('slides.ribbon.label-blurb-replaces', { label, blurb }) : t('slides.ribbon.label-blurb', { label, blurb })) : t('slides.ribbon.label-select-shape', { label })}
                onClick={() => act('animate', { kind: 'entr', effect })}
              />
            ))}
            <Button dropdown
              tall
              icon="redo"
              label={t('slides.motion-paths')}
              className="sl-an-paths"
              pressed={animation?.kind === 'path'}
              disabled={!hasShape}
              title={hasShape ? t('slides.ribbon.motion-paths-shape') : t('slides.ribbon.motion-paths-select')}
              onClick={(e) => menu.open(e, EFFECT_MENU.find(([kind]) => kind === 'path')[2].map(([effect, label]) => ({ label, icon: 'redo', run: () => act('animate', { kind: 'path', effect }) })))}
            />
            <Button dropdown
              tall
              icon="more"
              label={t('slides.ribbon.more-effects')}
              disabled={!hasShape}
              title={hasShape ? t('slides.ribbon.more-effects-every') : t('slides.ribbon.more-effects-select')}
              onClick={(e) => menu.open(e, effectMenu('animate'))}
            />
            <Button dropdown
              tall
              icon="settings"
              label={t('slides.ribbon.effect-options')}
              disabled={!animationOptions && !sequenceItems.length}
              title={animationOptions || sequenceItems.length ? (animationOptions && sequenceItems.length ? t('slides.ribbon.effect-options-way-name', { name: animation?.name || t('slides.ribbon.effect') }) : animationOptions ? t('slides.ribbon.effect-options-way', { name: animation?.name || t('slides.ribbon.effect') }) : t('slides.ribbon.effect-options-words')) : t('slides.ribbon.effect-options-pick-effect')}
              onClick={(e) => menu.open(e, [
                ...(animationOptions || []).map(([value, label]) => ({ label, icon: animation?.direction === value ? 'check' : undefined, run: () => act('animPatch', { direction: value }) })),
                ...(animationOptions && sequenceItems.length ? ['-'] : []),
                ...sequenceItems,
              ])}
            />
          </Group>
          <Group label={t('slides.ribbon.advanced')}>
            <Button dropdown tall icon="plus" label={t('slides.ribbon.add')} disabled={!hasShape} title={hasShape ? t('slides.ribbon.add-animation-another') : t('slides.ribbon.add-animation')} onClick={(e) => menu.open(e, effectMenu('addAnimation'))} />
            <Rows>
              <Button icon="list" label={t('slides.animation')} pressed={view.pane === 'animations'} title={t('slides.ribbon.animation-pane-slides')} onClick={() => act('pane', 'animations')} />
              <Button dropdown
                icon="play"
                label={t('slides.ribbon.trigger')}
                disabled={!animation}
                title={animation ? (animation.triggerShape ? t('slides.ribbon.trigger-start-effect-click') : t('slides.ribbon.trigger-start-effect')) : t('slides.ribbon.trigger-pick-effect')}
                onClick={(e) => menu.open(e, [
                  { label: t('slides.ribbon.on-click-sequence'), icon: !animation?.triggerShape ? 'check' : undefined, run: () => act('animPatch', { triggerShape: null }) },
                  '-',
                  ...(model?.slide?.shapes || []).filter((s) => s.groupId == null).map((s) => ({
                    label: t('slides.ribbon.click-name', { name: s.name || t('slides.shape-id', { id: s.id }) }),
                    icon: String(animation?.triggerShape) === String(s.id) ? 'check' : undefined,
                    run: () => act('animPatch', { triggerShape: s.id }),
                  })),
                ])}
              />
              <Button
                icon="wand"
                label={t('slides.ribbon.animation')}
                pressed={animPainter}
                disabled={!animPainter && !shapeAnimated}
                title={animPainter ? t('slides.ribbon.animation-painter-armed') : shapeAnimated ? t('slides.ribbon.animation-painter-shapes') : t('slides.ribbon.animation-painter-select')}
                onClick={() => act('animPainter')}
              />
            </Rows>
          </Group>
          <Group label={t('common.timing')}>
            <Rows>
              <div className="sl-rb-field" data-tip={t('slides.ribbon.start-effect-plays')}>
                <Icon name="play" size={14} />
                <span className="sl-rb-label">{t('slides.ribbon.start')}</span>
                <Select className="rw-select sl-an-start" aria-label={t('slides.ribbon.start')} value={animation?.trigger || 'onClick'} disabled={!animation} style={{ width: 128, height: 24 }} onChange={(e) => act('animPatch', { trigger: e.target.value })}>
                  <option value="onClick">{t('slides.ribbon.on-click')}</option>
                  <option value="withPrevious">{t('slides.ribbon.with-previous')}</option>
                  <option value="afterPrevious">{t('slides.ribbon.after-previous')}</option>
                </Select>
              </div>
              <div className="sl-rb-field" data-tip={t('slides.ribbon.duration-how-long')}>
                <Icon name="clock" size={14} />
                <span className="sl-rb-label">{t('slides.ribbon.duration')}</span>
                <SecondsField className="sl-an-duration" label={t('slides.ribbon.duration')} value={animation ? animation.duration : null} disabled={!animation || !animation.duration} min={0.01} max={59} onCommit={(v) => act('animPatch', { duration: v })} />
              </div>
              <div className="sl-rb-field" data-tip={t('slides.ribbon.delay-how-long')}>
                <Icon name="clock" size={14} />
                <span className="sl-rb-label">{t('slides.ribbon.delay')}</span>
                <SecondsField className="sl-an-delay" label={t('slides.ribbon.delay')} value={animation ? animation.delay : null} disabled={!animation} min={0} max={3600} onCommit={(v) => act('animPatch', { delay: v })} />
              </div>
            </Rows>
            <Separator />
            <Rows>
              <div className="sl-rb-caption">{t('slides.ribbon.reorder-animation')}</div>
              <Button icon="chevronUp" label={t('common.move-earlier')} disabled={!animation || animation.index <= 0} title={animation ? t('slides.ribbon.move-earlier-effect') : t('slides.ribbon.move-earlier-pick')} onClick={() => act('animMove', 'earlier')} />
              <Button icon="chevronDown" label={t('common.move-later')} disabled={!animation || animation.index >= animationList.length - 1} title={animation ? t('slides.ribbon.move-later-effect') : t('slides.ribbon.move-later-pick')} onClick={() => act('animMove', 'later')} />
            </Rows>
          </Group>
        </>
      ) : null}

      {/* ── Slide Show ───────────────────────────────────────────────────── */}
      {tab === 'show' ? (
        <>
          <Group label={t('slides.ribbon.start-slide')}>
            <Button tall icon="play" label={t('slides.ribbon.from')} title={t('slides.ribbon.start-show-first')} onClick={() => act('present', 'start')} />
            <Button tall icon="play" label={t('slides.ribbon.from-current')} title={t('slides.ribbon.start-show-slide')} onClick={() => act('present', 'here')} />
            <Button tall icon="list" label={t('slides.ribbon.custom-slide')} title={t('slides.ribbon.custom-slide-show')} onClick={(e) => act('customShowMenu', e)} />
          </Group>
          <Group label={t('slides.ribbon.set-up')}>
            <Button tall icon="settings" label={t('slides.ribbon.set-slide')} title={t('slides.ribbon.set-slide-show')} onClick={() => act('setupShow')} />
            <Button
              tall
              icon="eye"
              label={t('slides.ribbon.hide-slide')}
              pressed={Boolean(model?.slide?.hidden)}
              title={model?.slide?.hidden ? t('slides.ribbon.hide-slide-slide') : t('slides.ribbon.hide-slide-leave')}
              onClick={() => act('hideSlide')}
            />
            <Button tall icon="clock" label={t('slides.ribbon.rehearse')} title={t('slides.ribbon.rehearse-timings-run')} onClick={() => act('rehearse')} />
            <Button dropdown tall icon="video" label={t('common.record')} title={t('slides.ribbon.record-show-microphone')} onClick={(e) => menu.open(e, [{ label: t('slides.ribbon.from-current-slide'), icon: 'play', run: () => act('recordShow', 'here') }, { label: t('slides.ribbon.from-beginning'), icon: 'play', run: () => act('recordShow', 'start') }])} />
            <Button
              icon="check"
              label={t('slides.ribbon.play')}
              pressed={model?.showSettings?.narration !== false}
              title={t('slides.ribbon.play-narrations-whether')}
              onClick={() => act('showFlag', { narration: model?.showSettings?.narration === false })}
            />
            <Button
              icon="check"
              label={t('slides.ribbon.use-timings')}
              pressed={model?.showSettings?.useTimings !== false}
              title={t('slides.ribbon.use-timings-slides')}
              onClick={() => act('showFlag', { useTimings: model?.showSettings?.useTimings === false })}
            />
            <Button
              icon="check"
              label={t('slides.ribbon.show-media')}
              pressed={model?.showSettings?.mediaControls !== false}
              title={t('slides.ribbon.show-media-controls')}
              onClick={() => act('showFlag', { mediaControls: model?.showSettings?.mediaControls === false })}
            />
          </Group>
          <Group label={t('slides.ribbon.monitors')}>
            <Button icon="grid" label={t('slides.ribbon.monitor-name', { name: monitorName })} title={t('slides.ribbon.monitor-screen-slide')} onClick={(e) => act('monitorMenu', e)} />
            <Button icon="check" label={t('slides.ribbon.use')} title={t('slides.ribbon.use-presenter-view')} onClick={presentWithNotes} />
          </Group>
        </>
      ) : null}

      {/* ── Record ───────────────────────────────────────────────────────── */}
      {tab === 'record' ? (
        <>
          <Group label={t('common.preview')}>
            <Button tall icon="play" label={t('common.preview')} title={t('slides.ribbon.preview-slides-narration')} onClick={() => act('previewNarration')} />
          </Group>
          <Group label={t('common.camera')}>
            <Button tall icon="video" label={t('slides.ribbon.cameo')} title={t('slides.ribbon.cameo-camera-live')} onClick={() => act('cameo')} />
          </Group>
          <Group label={t('common.record')}>
            <Button tall icon="video" label={t('slides.ribbon.from')} title={t('slides.ribbon.record-beginning-show')} onClick={() => act('recordShow', 'start')} />
            <Button tall icon="video" label={t('slides.ribbon.from-current')} title={t('slides.ribbon.record-current-slide')} onClick={() => act('recordShow', 'here')} />
            <Button tall icon="video" label={t('common.screen-recording')} title={t('slides.ribbon.screen-recording-record')} onClick={() => act('screenRecording')} />
            <Button tall icon="volume" label={t('common.audio')} title={t('slides.ribbon.record-audio-record')} onClick={() => act('recordAudio')} />
          </Group>
          <Group label={t('common.edit')}>
            <Button dropdown tall icon="close" label={t('slides.ribbon.clear')} title={t('slides.ribbon.clear-narration-timings')} onClick={(e) => menu.open(e, [{ label: t('slides.ribbon.clear-narration'), run: () => act('clearRecording', 'narrationHere') }, { label: t('slides.ribbon.clear-narration-all'), run: () => act('clearRecording', 'narrationAll') }, '-', { label: t('slides.ribbon.clear-timings'), run: () => act('clearRecording', 'timingsHere') }, { label: t('slides.ribbon.clear-timings-all'), run: () => act('clearRecording', 'timingsAll') }])} />
            <Button tall icon="undo" label={t('slides.ribbon.reset-to')} className="sl-reset-cameo" title={t('slides.ribbon.reset-cameo-cameras')} onClick={() => act('resetCameo')} />
          </Group>
          <Group label={t('common.export')}>
            <Button tall icon="export" label={t('slides.save-as-show')} title={t('slides.ribbon.save-show-copy')} onClick={() => act('saveAsShow')} />
            <Button tall icon="video" label={t('slides.export-to-video')} title={t('slides.ribbon.export-video-deck')} onClick={() => act('exportVideo')} />
          </Group>
          <Group label={t('common.help')}>
            <Button tall icon="info" label={t('slides.ribbon.learn-more')} onClick={() => act('help')} />
          </Group>
        </>
      ) : null}

      {/* ── Review ───────────────────────────────────────────────────────── */}
      {tab === 'review' ? (
        <>
          <Group label={t('common.proofing')}>
            <Button tall icon="check" label={t('common.spelling')} pressed={review?.pane === 'editor'} title={t('slides.ribbon.spelling-f7-check')} onClick={() => review?.startSpelling()} />
            <Button tall icon="find" label={t('common.thesaurus')} pressed={review?.pane === 'thesaurus'} title={t('slides.ribbon.thesaurus-shift-f7')} onClick={() => review?.openThesaurus()} />
          </Group>
          <Group label={t('common.accessibility')}>
            <Button tall icon="shield" label={t('common.check-accessibility')} pressed={review?.pane === 'accessibility'} title={t('slides.ribbon.check-accessibility-alt')} onClick={() => review?.openAccessibility()} />
            <Button icon="textbox" label={t('common.alt-text')} disabled={selected == null} title={selected == null ? t('slides.ribbon.alt-text-select') : t('slides.ribbon.alt-text-describe')} onClick={() => selected != null && review?.openAltText({ slide: index, shape: selected })} />
          </Group>
          <Group label={t('word.ribbon.speech')}>
            <Button tall dropdown icon="volume" label={t('word.ribbon.read-aloud')} pressed={Boolean(view.readingAloud)} title={t('slides.ribbon.read-aloud-words')} onClick={(e) => (view.readingAloud ? act('readAloud', 'stop') : menu.open(e, [
              { label: t('slides.ribbon.this-slide'), icon: 'slides', run: () => act('readAloud', 'slide') },
              { label: t('slides.ribbon.its-notes'), icon: 'word', run: () => act('readAloud', 'notes') },
              { label: t('slides.ribbon.slide-notes'), icon: 'volume', run: () => act('readAloud', 'both') },
            ]))} />
          </Group>
          <Group label={t('common.language')}>
            <Button tall icon="globe" label={t('common.language')} title={t('slides.ribbon.language-mark-selected')} onClick={() => act('language')} />
          </Group>
          <Group label={t('slides.ribbon.activity')}>
            <Button icon="check" label={t('slides.ribbon.mark-all')} className="sl-mark-read" disabled={!view.unread} title={view.unread ? tn(view.unread, 'slides.ribbon.mark-all-read-count') : t('slides.ribbon.mark-all-read')} onClick={() => act('markAllRead')} />
            <Button icon="eye" label={t('slides.ribbon.show-changes')} className="sl-show-changes" pressed={view.pane === 'changes'} title={view.changes ? tn(view.changes, 'slides.ribbon.show-changes-count') : t('slides.ribbon.show-changes-different')} onClick={() => act('pane', 'changes')} />
          </Group>
          <Group label={t('common.comments')}>
            <Button tall icon="comment" label={t('common.new-comment')} title={t('slides.ribbon.new-comment-selected')} onClick={() => act('newComment')} />
            <Rows>
              <Button dropdown icon="trash" label={t('common.delete')} disabled={!commentCount} title={commentCount ? t('slides.ribbon.delete-comment-every') : t('slides.ribbon.delete-no-comments')} onClick={(e) => menu.open(e, [
                { label: t('slides.ribbon.delete-comment'), icon: 'trash', run: () => act('deleteComment') },
                { label: t('slides.ribbon.delete-all-comments'), icon: 'trash', run: () => act('deleteComments', 'slide') },
                { label: t('slides.ribbon.delete-all-comments-presentation'), icon: 'trash', run: () => act('deleteComments', 'all') },
              ])} />
              <Button icon="chevronLeft" label={t('common.previous')} disabled={!commentCount} title={commentCount ? t('slides.ribbon.previous-comment-before') : t('slides.ribbon.previous-no-comments')} onClick={() => act('commentStep', -1)} />
              <Button icon="chevronRight" label={t('common.next')} disabled={!commentCount} title={commentCount ? t('slides.ribbon.next-comment-after') : t('slides.ribbon.next-no-comments')} onClick={() => act('commentStep', 1)} />
            </Rows>
            <Button tall icon="eye" label={t('common.show-comments')} pressed={view.pane === 'comments'} title={t('slides.ribbon.show-comments-comments')} onClick={() => act('pane', 'comments')} />
          </Group>
          <Group label={t('common.compare')}>
            <Button tall icon="swap" label={t('common.compare')} title={t('slides.ribbon.compare-merge-reviewed')} onClick={() => act('compare')} />
            <Button tall dropdown icon="check" label={t('common.accept')} disabled={!model?.review} title={model?.review ? t('slides.ribbon.accept-change-slides') : t('slides.ribbon.accept-compare')} onClick={(e) => menu.open(e, [
              { label: t('slides.ribbon.accept-change'), icon: 'check', run: () => act('reviewDecide', { scope: 'change', accept: true }) },
              { label: t('slides.ribbon.accept-all-changes'), run: () => act('reviewDecide', { scope: 'slide', accept: true }) },
              { label: t('slides.ribbon.accept-all-changes-presentation'), run: () => act('reviewDecide', { scope: 'all', accept: true }) },
            ])} />
            <Button tall dropdown icon="close" label={t('common.reject')} disabled={!model?.review} title={model?.review ? t('slides.ribbon.reject-change-slides') : t('slides.ribbon.reject-compare')} onClick={(e) => menu.open(e, [
              { label: t('slides.ribbon.reject-change'), icon: 'close', run: () => act('reviewDecide', { scope: 'change', accept: false }) },
              { label: t('slides.ribbon.reject-all-changes'), run: () => act('reviewDecide', { scope: 'slide', accept: false }) },
              { label: t('slides.ribbon.reject-all-changes-presentation'), run: () => act('reviewDecide', { scope: 'all', accept: false }) },
            ])} />
            <Rows>
              <Button icon="chevronLeft" label={t('slides.ribbon.previous')} disabled={!model?.review} title={t('slides.ribbon.previous-change-slide')} onClick={() => act('reviewStep', -1)} />
              <Button icon="chevronRight" label={t('slides.ribbon.next-change')} disabled={!model?.review} title={t('slides.ribbon.next-change-slide')} onClick={() => act('reviewStep', 1)} />
              <Button icon="list" label={t('common.reviewing')} pressed={view.pane === 'revisions'} title={t('slides.ribbon.reviewing-pane-revisions')} onClick={() => act('pane', 'revisions')} />
            </Rows>
            <Button tall icon="stop" label={t('slides.end-review')} disabled={!model?.review} title={t('slides.ribbon.end-review-close')} onClick={() => act('reviewEnd')} />
          </Group>
          <Group label={t('common.protect')}>
            <Button tall icon="shield" label={t('common.redact')} pressed={Boolean(redaction?.open)} title={t('common.redact-mark-words')} onClick={() => redaction?.toggle()} />
          </Group>
          <Group label={t('common.notes')}>
            <Button tall icon="word" label={t('slides.ribbon.speaker')} title={t('slides.ribbon.speaker-notes-slides')} onClick={() => setNotesOpen(true)} />
          </Group>
          <Group label={t('common.ink')}>
            <Button icon="eye" label={t('common.hide-ink')} pressed={Boolean(ink?.hide)} title={t('slides.ribbon.hide-ink-slides')} onClick={() => act('hideInk')} />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
      {tab === 'view' ? (
        <>
          <Group label={t('slides.ribbon.presentation')}>
            <Button tall icon="slides" label={t('common.normal')} pressed={(view.mode || 'normal') === 'normal'} onClick={() => act('mode', 'normal')} />
            <Button tall icon="list" label={t('slides.ribbon.outline-view')} pressed={view.mode === 'outline'} title={t('slides.ribbon.every-slides-words')} onClick={() => act('mode', 'outline')} />
            <Button tall icon="grid" label={t('slides.ribbon.slide-sorter')} pressed={view.mode === 'sorter'} title={t('slides.ribbon.all-slides-once')} onClick={() => act('mode', 'sorter')} />
            <Button tall icon="word" label={t('slides.ribbon.notes-page')} pressed={view.mode === 'notes'} title={t('slides.ribbon.slide-notes-under')} onClick={() => act('mode', 'notes')} />
            <Button tall icon="play" label={t('slides.ribbon.reading-view')} title={t('slides.ribbon.show-window-not')} onClick={() => act('present', 'reading')} />
          </Group>
          <Group label={t('slides.ribbon.master-views')}>
            <Button tall icon="slides" label={t('slides.slide-master')} pressed={masterView?.kind === 'slide'} title={t('slides.ribbon.slide-master-edit')} onClick={() => act(masterView?.kind === 'slide' ? 'closeMaster' : 'masterView')} />
            <Button tall icon="file" label={t('slides.handout')} pressed={masterView?.kind === 'handout'} title={t('slides.ribbon.handout-master-printed')} onClick={() => act(masterView?.kind === 'handout' ? 'closeMaster' : 'handoutMasterView')} />
            <Button tall icon="word" label={t('slides.notes')} pressed={masterView?.kind === 'notes'} title={t('slides.ribbon.notes-master-printed')} onClick={() => act(masterView?.kind === 'notes' ? 'closeMaster' : 'notesMasterView')} />
          </Group>
          <Group label={t('common.show')}>
            <Button icon="minus" label={t('common.ruler')} pressed={Boolean(view.ruler)} onClick={() => act('toggle', 'ruler')} />
            <Button icon="grid" label={t('common.gridlines')} pressed={Boolean(view.gridlines)} onClick={() => act('toggle', 'gridlines')} />
            <Button icon="plus" label={t('slides.ribbon.guides')} pressed={Boolean(view.guides)} onClick={() => act('toggle', 'guides')} />
            <Button icon="word" label={t('common.notes')} pressed={view.notes !== false} title={t('slides.ribbon.notes-strip-under')} onClick={() => act('toggle', 'notes')} />
            <Button icon="list" label={t('common.layers')} pressed={view.pane === 'layers'} title={t('slides.ribbon.slides-shapes-layers')} onClick={() => act('pane', 'layers')} />
            <Button icon="grid" label={t('slides.designs')} pressed={view.pane === 'designs'} title={t('slides.ribbon.decks-layouts-put')} onClick={() => act('pane', 'designs')} />
          </Group>
          <Group label={t('common.direction')}>
            <Button icon="rotate" label={t('slides.ribbon.view')} pressed={Boolean(view.rtl)} title={view.rtl ? t('slides.ribbon.view-direction-right') : t('slides.ribbon.view-direction-left')} onClick={() => act('toggle', 'rtl')} />
          </Group>
          <Group label={t('common.zoom')}>
            <Button dropdown tall icon="zoomIn" label={t('common.zoom')} onClick={(e) => menu.open(e, [50, 75, 100, 150, 200].map((z) => ({ label: `${z}%`, run: () => act('zoom', z / 100) })))} />
            <Button tall icon="maximize" label={t('common.fit-to')} pressed={!view.zoom} onClick={() => act('zoom', null)} />
          </Group>
          <Group label={t('slides.ribbon.colour')}>
            <Button icon="contrast" label={t('common.colour')} pressed={!view.tone || view.tone === 'colour'} onClick={() => act('tone', 'colour')} />
            <Button icon="contrast" label={t('common.greyscale')} pressed={view.tone === 'grey'} onClick={() => act('tone', 'grey')} />
            <Button icon="contrast" label={t('slides.ribbon.black-and')} pressed={view.tone === 'mono'} onClick={() => act('tone', 'mono')} />
          </Group>
          <Group label={t('common.window')}>
            <Button tall icon="new" label={t('common.new-window')} title={t('slides.ribbon.deck-second-window')} onClick={() => act('newWindow')} />
            <Button icon="grid" label={t('common.arrange-all')} title={t('slides.ribbon.arrange-all-every')} onClick={() => act('arrange', 'columns')} />
            <Button icon="grid" label={t('slides.ribbon.cascade')} title={t('slides.ribbon.cascade-every')} onClick={() => act('arrange', 'cascade')} />
            <Button icon="minus" label={t('slides.ribbon.move-split')} title={t('slides.ribbon.move-split-arrow')} onClick={() => act('moveSplit')} />
            <Button icon="list" label={t('slides.ribbon.switch')} title={t('slides.ribbon.switch-windows')} onClick={(e) => act('switchWindows', e)} />
            <Button icon="maximize" label={t('common.full-screen')} onClick={() => shell.win.fullscreen({})} />
          </Group>
        </>
      ) : null}

      {/* ── Help ─────────────────────────────────────────────────────────── */}
      {tab === 'help' ? (
        <>
          <Group label={t('common.help-and')}>
            <Button tall icon="info" label={t('common.help')} title={t('slides.ribbon.presentation-guide')} onClick={() => act('help')} />
            <Button tall icon="send" label={t('common.feedback')} title={t('common.tell-wrong-missing')} onClick={() => act('feedback')} />
            <Button tall icon="info" label={t('slides.ribbon.show')} onClick={() => act('help')} />
            <Button tall icon="star" label={t('common.whats-new')} title={t('slides.ribbon.the-release-notes')} onClick={() => act('releases')} />
            <Button icon="list" label={t('common.keyboard')} onClick={() => act('shortcuts')} />
          </Group>
          <Group label={t('common.file')}>
            <Button tall icon="new" label={t('common.new')} onClick={() => shell.win.create({ app: 'slides' })} />
            <Button tall icon="open" label={t('common.open')} onClick={openFile} />
            <Button tall icon="save" label={t('common.save')} onClick={() => save(false)} />
          </Group>
        </>
      ) : null}

      {/* ── PDF ──────────────────────────────────────────────────────────── */}
      {tab === 'pdf' ? (
        <>
          <Group label={t('common.export')}>
            <Button tall icon="pdf" label={t('slides.ribbon.export-as')} title={t('slides.ribbon.every-slide-page')} onClick={() => exportAs('pdf')} />
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
            <Group label={t('common.3d-model-views')}>
              {rows.map((row, i) => (
                <Rows key={i}>
                  {row.map(([key, label, view]) => <Button key={key} icon={at(view) ? 'check' : 'shape'} label={t(label)} className="sl-model3d-view" data-view={key} pressed={at(view)} title={t('common.label-model-turned', { label: t(label) })} onClick={() => act('model3dView', view)} />)}
                </Rows>
              ))}
            </Group>
            <Group label={t('common.turn')}>
              <Rows>
                <Button icon="chevronLeft" label={t('common.turn-left')} className="sl-model3d-turn-left" title={t('common.turn-left-fifteen')} onClick={() => act('model3dView', { turn: { yaw: -15, pitch: 0 } })} />
                <Button icon="chevronRight" label={t('common.turn-right')} title={t('common.turn-right-fifteen')} onClick={() => act('model3dView', { turn: { yaw: 15, pitch: 0 } })} />
              </Rows>
              <Rows>
                <Button icon="chevronUp" label={t('common.tip-back')} title={t('common.tip-back-fifteen')} onClick={() => act('model3dView', { turn: { yaw: 0, pitch: 15 } })} />
                <Button icon="chevronDown" label={t('common.tip-forward')} title={t('common.tip-forward-fifteen')} onClick={() => act('model3dView', { turn: { yaw: 0, pitch: -15 } })} />
              </Rows>
            </Group>
            <Group label={t('common.adjust')}>
              <Button tall icon="undo" label={t('common.reset-3d')} className="sl-model3d-reset" title={t('common.reset-3d-model')} onClick={() => act('model3dView', 'reset')} />
            </Group>
          </>
        );
      })() : null}

      {/* ── Camera Format (contextual) ─────────────────────────────────────── */}
      {/* ── Shape Format (contextual): WordArt's Transform ──────────────── */}
      {tab === 'shapeFormat' && selectedShapeObj?.kind === 'shape' && selectedShapeObj?.text ? (
        <>
          <Group label={t('common.text-effects')}>
            {WARP_PRESETS.map((p) => (
              <Button key={p.id} tall icon="wand" label={t(p.label)} className="sl-warp" data-preset={p.id} pressed={(selectedShapeObj.text.warp?.preset || 'textNoShape') === p.id} title={p.id === 'textNoShape' ? t('common.no-transform-words') : p.label === 'Button' ? t('common.transform-words-along-button') : p.label === 'Circle' ? t('common.transform-words-along-circle') : t('common.transform-words-along')} onClick={() => act('textWarp', p.id)} />
            ))}
            {/* The warps, the rest of Office's gallery, from a menu. */}
            <Button dropdown tall icon="wand" label={t('common.more')} className="sl-warp-more" pressed={WARP_MORE.some((p) => p.id === selectedShapeObj.text.warp?.preset)} title={WARP_MORE.some((p) => p.id === selectedShapeObj.text.warp?.preset) ? t('slides.ribbon.transform-now-name', { name: t(warpLabel(selectedShapeObj.text.warp?.preset)) }) : t('common.more-transforms-warps')} onClick={(e) => menu.open(e, WARP_MORE.map((p) => ({ label: t(p.label), icon: selectedShapeObj.text.warp?.preset === p.id ? 'check' : 'wand', run: () => act('textWarp', p.id) })))} />
          </Group>
        </>
      ) : null}

      {tab === 'pictureFormat' && pictureSelected ? (
        <>
          <Group label={t('common.edit')}>
            <Button tall icon="image" label={t('common.edit-in')} title={t('common.edit-image-crop')} onClick={() => act('editInImage', { shape: selectedShapeObj.id })} data-role="edit-in-image" />
          </Group>
          <Group label={t('common.arrange')}>
            <Rows>
              <Button icon="layerUp" label={t('common.bring')} onClick={() => act('order', 'forward')} />
              <Button icon="layerDown" label={t('common.send-backward')} onClick={() => act('order', 'backward')} />
            </Rows>
            <Rows>
              <Button icon="layerTop" label={t('common.bring-to')} onClick={() => act('order', 'front')} />
              <Button icon="layerBottom" label={t('common.send-to-back')} onClick={() => act('order', 'back')} />
            </Rows>
          </Group>
        </>
      ) : null}

      {tab === 'cameraFormat' && selectedShapeObj?.cameo ? (
        <>
          <Group label={t('common.camera')}>
            <Button tall icon="eye" label={t('common.preview')} className="sl-cameo-preview" pressed={Boolean(view.cameoPreview)} title={view.cameoPreview ? t('slides.ribbon.preview-camera-press') : t('slides.ribbon.preview-camera-live')} onClick={() => act('cameoPreview')} />
          </Group>
          <Group label={t('slides.ribbon.camera-styles')}>
            {[['rect', t('common.rectangle')], ['ellipse', t('common.oval')], ['roundRect', t('slides.ribbon.rounded-rectangle')]].map(([preset, label]) => (
              <Button key={preset} tall icon="shape" label={label} className="sl-cameo-shape" data-preset={preset} pressed={selectedShapeObj.preset === preset} title={t('slides.ribbon.camera-shape-cameo', { shape: label.toLowerCase() })} onClick={() => act('cameoShape', preset)} />
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
            <Group label={t('common.table-style')}>
              <Rows>
                {toggle('firstRow', t('common.header-row'), t('common.header-row-first'))}
                {toggle('lastRow', t('common.total-row'), t('common.total-row-last'))}
                {toggle('bandRow', t('common.banded-rows'), t('common.banded-rows-every'))}
              </Rows>
              <Rows>
                {toggle('firstCol', t('common.first-column'), t('common.first-column-first'))}
                {toggle('lastCol', t('common.last-column'), t('common.last-column-last'))}
                {toggle('bandCol', t('common.banded'), t('common.banded-columns-every'))}
              </Rows>
            </Group>
            <Group label={t('common.shading')}>
              <Button dropdown tall icon="bucketFill" label={t('common.shading')} title={t('slides.ribbon.shading-cell-last')} onClick={(e) => menu.open(e, [
                { label: t('common.no-fill'), icon: 'close', run: () => act('tableShading', { fill: null }) },
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
