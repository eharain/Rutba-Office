// The Presentation ribbon.
//
// Laid out the way PowerPoint lays it out — Home, Insert, Draw, Design,
// Transitions, Animations, Slide Show, Record, Review, View, Help, and a PDF
// tab where PowerPoint has Acrobat — with every group PowerPoint has in each.
// What the engine can do is wired; what it cannot is drawn where PowerPoint
// draws it, disabled, with a title that says exactly why. Transitions and
// animations are the honest case: the file's own survive a round trip, and
// nothing here pretends to author one.

import React from 'react';
import { Ribbon, Group, Rows, Button, Separator, Select } from '@rutba/office-ui';

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
const TRANSITIONS = ['None', 'Morph', 'Fade', 'Push', 'Wipe', 'Split', 'Reveal', 'Cut', 'Random Bars', 'Shape', 'Uncover'];

const ANIMATIONS = ['None', 'Appear', 'Fade', 'Fly In', 'Float In', 'Split'];

/** A control that is drawn where PowerPoint draws it, and says why it is not live. */
const Soon = ({ icon, label, tall, why }) => (
  <Button tall={tall} icon={icon} label={label} disabled title={`${label} — not built yet. ${why}`} />
);
const INK = 'Ink is a drawing part (ink ML) the engine does not write, and the stage has no pen surface yet.';
const TRANSITION_WHY = 'A transition is preserved in the file when the deck has one; authoring one (writing p:transition) is not built.';
const ANIMATION_WHY = 'Animations are preserved in the file when the deck has them; authoring one (writing p:timing) is not built.';

export default function SlidesRibbon({
  tab, setTab, model, doc, commands, shell, menu, save, openFile, exportAs,
  act, view = {}, index = 0, selected = null, format = {}, addSlide, insertPicture, presentWithNotes, setPresent, setNotesOpen,

}) {
  const count = model?.count || 0;
  const hasShape = Boolean(selected);
  const size = Number(format.size || 18);
  const nearer = (dir) => {
    const bigger = SIZES.filter((s) => (dir > 0 ? s > size : s < size));
    return dir > 0 ? bigger[0] ?? size : bigger[bigger.length - 1] ?? size;
  };
  const fmt = (delta) => act('format', delta);
  const needShape = hasShape ? undefined : 'Select a text box first — click it once';

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
            <Soon tall icon="paste" label="Paste" why="A shape clipboard (copying shapes between slides) is not built; text inside a box pastes as text." />
            <Rows>
              <>
                <Soon icon="cut" label="Cut" why="Comes with the shape clipboard." />
                <Soon icon="copy" label="Copy" why="Comes with the shape clipboard." />
              </>
              <Soon icon="wand" label="Format Painter" why="Comes with shape formatting." />
            </Rows>
          </Group>
          <Group label="Slides">
            <Button tall icon="plus" label="New Slide" onClick={(e) => menu.open(e, LAYOUTS.map(([layout, label]) => ({ label, icon: 'slides', run: () => addSlide(layout) })))} />
            <Rows>
              <>
                <Soon icon="grid" label="Layout" why="Changing an existing slide's layout re-parents its placeholders; the engine inserts with a layout and does not yet swap one." />
                <Soon icon="undo" label="Reset" why="Comes with layouts." />
                <Soon icon="list" label="Section" why="Sections are a presentation-part list the engine does not write yet." />
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
                  <option value="">{format.font || 'Theme font'}</option>
                  {['Calibri', 'Calibri Light', 'Arial', 'Segoe UI', 'Georgia', 'Times New Roman', 'Verdana', 'Consolas'].filter((f) => f !== format.font).map((f) => <option key={f} value={f}>{f}</option>)}
                </Select>
                <Select value={String(size)} onChange={(e) => fmt({ size: Number(e.target.value) })} style={{ width: 58 }} title={needShape || 'Font size'} disabled={!hasShape}>
                  {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
                  {SIZES.includes(size) ? null : <option value={String(size)}>{size}</option>}
                </Select>
                <Button icon="chevronUp" title={needShape || 'Increase font size'} disabled={!hasShape} onClick={() => fmt({ size: nearer(1) })} />
                <Button icon="chevronDown" title={needShape || 'Decrease font size'} disabled={!hasShape} onClick={() => fmt({ size: nearer(-1) })} />
                <Soon icon="undo" label="" why="Clear all formatting comes with shape formatting." />
              </>
              <>
                <Button icon="bold" title={needShape || 'Bold (Ctrl+B)'} pressed={format.bold} disabled={!hasShape} onClick={() => fmt({ bold: 'toggle' })} />
                <Button icon="italic" title={needShape || 'Italic (Ctrl+I)'} pressed={format.italic} disabled={!hasShape} onClick={() => fmt({ italic: 'toggle' })} />
                <Button icon="underline" title={needShape || 'Underline (Ctrl+U)'} pressed={format.underline} disabled={!hasShape} onClick={() => fmt({ underline: 'toggle' })} />
                <Soon icon="strike" label="" why="Strikethrough is a run property the deck writer does not write yet." />
                <Soon icon="textbox" label="AV" why="Character spacing is a run property the deck writer does not write yet." />
                <Soon icon="textbox" label="Aa" why="Change case comes with text selection inside a box." />
                <Separator />
                <Soon icon="wand" label="" why="Text highlight is a run property the deck writer does not write yet." />
                <Button icon="contrast" title={needShape || 'Font colour'} disabled={!hasShape} onClick={(e) => menu.open(e, COLOURS.map(([value, label]) => ({ label, run: () => fmt({ color: value }) })))} />
              </>
            </Rows>
          </Group>
          <Group label="Paragraph">
            <Rows>
              <>
                <Soon icon="listBullet" label="" why="Bullets are a paragraph property (a:buChar) the deck writer does not write yet." />
                <Soon icon="listNumber" label="" why="Numbering comes with bullets." />
                <Soon icon="chevronLeft" label="" why="Indent levels come with bullets." />
                <Soon icon="chevronRight" label="" why="Indent levels come with bullets." />
                <Soon icon="list" label="" why="Line spacing is a paragraph property the deck writer does not write yet." />
                <Separator />
                <Soon icon="grid" label="" why="Columns inside a text box are a body property the deck writer does not write yet." />
                <Soon icon="rotate" label="" why="Text direction comes with body properties." />
                <Soon icon="chevronUp" label="" why="Vertical alignment inside the box comes with body properties." />
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
              { label: 'Delete shape', icon: 'trash', run: () => act('deleteShape') },
              { label: 'Move up (nudge)', run: () => act('nudge', { dy: -8 }) },
              { label: 'Move down (nudge)', run: () => act('nudge', { dy: 8 }) },
              { label: 'Move left (nudge)', run: () => act('nudge', { dx: -8 }) },
              { label: 'Move right (nudge)', run: () => act('nudge', { dx: 8 }) },
            ])} disabled={!hasShape} title={needShape || 'Arrange the selected shape'} />
            <Soon tall icon="wand" label="Quick Styles" why="Shape styles are fill and outline presets; comes with shape formatting." />
            <Soon icon="wand" label="Shape Fill" why="A shape's fill is spPr the deck writer does not write yet." />
            <Soon icon="shape" label="Shape Outline" why="Comes with shape formatting." />
            <Soon icon="wand" label="Shape Effects" why="Comes with shape formatting." />
          </Group>
          <Group label="Editing">
            <Soon tall icon="find" label="Editing" why="Find, Replace and Select across a deck are on the list." />
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
            <Soon tall icon="table" label="Table" why="A table on a slide is a graphic frame the engine reads and does not yet write." />
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
            <Soon tall icon="chart" label="Chart" why="A chart on a slide is a chart part plus a graphic frame; the Word and Worksheets chart writer will reach here." />
          </Group>
          <Group label="Forms">
            <Soon tall icon="check" label="Forms" why="Forms is a Microsoft cloud service." />
          </Group>
          <Group label="Links">
            <Soon tall icon="zoomIn" label="Zoom" why="Slide zooms are a PowerPoint-only feature." />
            <Soon tall icon="link" label="Link" why="A hyperlink on a shape needs a relationship the deck writer does not write yet." />
            <Soon tall icon="play" label="Action" why="Comes with links." />
          </Group>
          <Group label="Comments">
            <Soon tall icon="reply" label="Comment" why="Slide comments are a comments part the engine does not write yet." />
          </Group>
          <Group label="Text">
            <Button tall icon="textbox" label="Text Box" onClick={() => commands['slide.textbox']?.run?.()} />
            <Soon tall icon="file" label="Header & Footer" why="Headers and footers are placeholders on the master the engine does not write yet." />
            <Soon tall icon="wand" label="WordArt" why="WordArt is text effects the deck writer does not write." />
            <Soon icon="clock" label="Date & Time" why="A date field is a placeholder on the master." />
            <Soon icon="list" label="Slide Number" why="A slide-number field is a placeholder on the master." />
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
            <Soon tall icon="wand" label="Themes" why="A theme is the theme part; the engine reads the deck's own and does not yet swap it for another." />
          </Group>
          <Group label="Customise">
            <Soon tall icon="wand" label="Variants" why="Comes with themes." />
            <Soon icon="contrast" label="Colours" why="Comes with themes." />
            <Soon icon="textbox" label="Fonts" why="Comes with themes." />
            <Soon icon="wand" label="Effects" why="Comes with themes." />
            <Soon icon="picture" label="Background Styles" why="A slide background is cSld/bg the deck writer does not write yet." />
          </Group>
          <Group label="Slides">
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
            <Soon tall icon="play" label="Preview" why={TRANSITION_WHY} />
          </Group>
          <Group label="Transition to This Slide">
            {TRANSITIONS.map((t) => <Soon key={t} tall icon="slides" label={t} why={TRANSITION_WHY} />)}
            <Soon icon="settings" label="Effect Options" why={TRANSITION_WHY} />
          </Group>
          <Group label="Timing">
            <Soon icon="volume" label="Sound" why={TRANSITION_WHY} />
            <Soon icon="clock" label="Duration" why={TRANSITION_WHY} />
            <Soon icon="check" label="Apply To All" why={TRANSITION_WHY} />
            <Soon icon="chevronRight" label="Advance Slide" why={TRANSITION_WHY} />
          </Group>
        </>
      ) : null}

      {/* ── Animations ───────────────────────────────────────────────────── */}
      {tab === 'animations' ? (
        <>
          <Group label="Preview">
            <Soon tall icon="play" label="Preview" why={ANIMATION_WHY} />
          </Group>
          <Group label="Animation">
            {ANIMATIONS.map((a) => <Soon key={a} tall icon="star" label={a} why={ANIMATION_WHY} />)}
            <Soon icon="settings" label="Effect Options" why={ANIMATION_WHY} />
          </Group>
          <Group label="Advanced Animation">
            <Soon tall icon="plus" label="Add Animation" why={ANIMATION_WHY} />
            <Soon icon="list" label="Animation Pane" why={ANIMATION_WHY} />
            <Soon icon="play" label="Trigger" why={ANIMATION_WHY} />
            <Soon icon="wand" label="Animation Painter" why={ANIMATION_WHY} />
          </Group>
          <Group label="Timing">
            <Soon icon="play" label="Start" why={ANIMATION_WHY} />
            <Soon icon="clock" label="Duration" why={ANIMATION_WHY} />
            <Soon icon="clock" label="Delay" why={ANIMATION_WHY} />
            <Soon icon="chevronUp" label="Move Earlier" why={ANIMATION_WHY} />
            <Soon icon="chevronDown" label="Move Later" why={ANIMATION_WHY} />
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
            <Soon tall icon="eye" label="Hide Slide" why="Hiding a slide is a show attribute the engine does not write yet." />
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
