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
import { Ribbon, Group, Button, Separator, Select, Input } from '@rutba/office-ui';

/* ── vocabularies ────────────────────────────────────────────────────────── */

const FONTS = ['Calibri', 'Calibri Light', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Tahoma', 'Garamond', 'Consolas', 'Courier New'];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

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
const PAGE_SIZES = [['A4', 'A4 — 21 × 29.7 cm'], ['Letter', 'Letter — 8.5 × 11 in'], ['Legal', 'Legal — 8.5 × 14 in']];
const MARGINS = [['normal', 'Normal — 2.54 cm all round'], ['narrow', 'Narrow — 1.27 cm all round'], ['wide', 'Wide — 5.08 cm at the sides']];
const SHAPES = [
  ['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Oval'], ['triangle', 'Triangle'], ['diamond', 'Diamond'],
  ['rightArrow', 'Arrow: right'], ['leftArrow', 'Arrow: left'], ['upArrow', 'Arrow: up'], ['downArrow', 'Arrow: down'],
  ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'], ['star5', 'Star: 5 points'], ['plus', 'Cross'], ['chevron', 'Chevron'],
  ['parallelogram', 'Parallelogram'], ['trapezoid', 'Trapezoid'], ['line', 'Line'],
];
const CHARTS = [['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['pie', 'Pie']];
const CASES = [['sentence', 'Sentence case.'], ['lower', 'lowercase'], ['upper', 'UPPERCASE'], ['title', 'Capitalize Each Word'], ['toggle', 'tOGGLE cASE']];

/** A control that is drawn where Word draws it, and says why it is not live. */
const Soon = ({ icon, label, tall, why }) => (
  <Button tall={tall} icon={icon} label={label} disabled title={`${label} — not built yet. ${why}`} />
);

export default function WordRibbon({
  tab, setTab, doc, model, dispatch, commands, shell, menu, save, openFile, exportAs, openDialog, insertPicture, act, view = {},
}) {
  const format = model?.format || {};
  const styles = Array.isArray(model?.styles) ? model.styles : [];
  const section = model?.section || null;
  const comments = model?.comments || [];
  const landscape = section?.orientation === 'landscape';

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
            <Button icon="cut" label="Cut" title="Cut (Ctrl+X)" onClick={() => commands['edit.cut']?.run?.()} />
            <Button icon="copy" label="Copy" title="Copy (Ctrl+C)" onClick={() => commands['edit.copy']?.run?.()} />
            <Button icon="wand" label="Format Painter" pressed={Boolean(view.painting)} title="Format Painter — copy the formatting here, then select the words to paint" onClick={() => act('formatPainter')} />
          </Group>

          <Group label="Font">
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
            <Separator />
            <Button icon="bold" title="Bold (Ctrl+B)" pressed={format.bold} onClick={() => dispatch({ op: 'toggleFormat', tag: 'bold' })} />
            <Button icon="italic" title="Italic (Ctrl+I)" pressed={format.italic} onClick={() => dispatch({ op: 'toggleFormat', tag: 'italic' })} />
            <Button icon="underline" title="Underline (Ctrl+U)" pressed={format.underline} onClick={() => dispatch({ op: 'toggleFormat', tag: 'underline' })} />
            <Button icon="strike" title="Strikethrough" pressed={format.strike} onClick={() => dispatch({ op: 'toggleFormat', tag: 'strike' })} />
            <Button icon="chevronDown" label="x₂" title="Subscript" pressed={format.vertAlign === 'subscript'} onClick={() => run({ vertAlign: format.vertAlign === 'subscript' ? null : 'subscript' })} />
            <Button icon="chevronUp" label="x²" title="Superscript" pressed={format.vertAlign === 'superscript'} onClick={() => run({ vertAlign: format.vertAlign === 'superscript' ? null : 'superscript' })} />
            <Separator />
            <Soon icon="wand" label="A" why="Text effects (outline, shadow, glow) are DrawingML the engine does not write." />
            <Button icon="wand" title="Text highlight colour" pressed={Boolean(format.highlight)} onClick={(e) => colourMenu(e, 'highlight', HIGHLIGHTS)} />
            <Button icon="contrast" title="Font colour" onClick={(e) => colourMenu(e, 'fontColour', TEXT_COLOURS)} />
          </Group>

          <Group label="Paragraph">
            <Button icon="listBullet" title="Bullets" pressed={format.listType === 'bullet'} onClick={() => toggleList('bullet')} />
            <Button icon="listNumber" title="Numbering" pressed={format.listType === 'number'} onClick={() => toggleList('number')} />
            <Soon icon="listNumber" label="" why="Multilevel lists need a numbering definition the engine does not write." />
            <Separator />
            <Button icon="chevronLeft" title="Decrease indent" disabled={!format.indentLevel} onClick={() => para({ indentDelta: -1 })} />
            <Button icon="chevronRight" title="Increase indent" onClick={() => para({ indentDelta: 1 })} />
            <Soon icon="sort" label="" why="Sort paragraphs — coming with the table sorter." />
            <Button icon="formula" title="Show formatting marks (¶)" pressed={Boolean(view.marks)} onClick={() => act('toggleMarks')} />
            <Separator />
            <Button icon="alignLeft" title="Align left (Ctrl+L)" pressed={format.paragraphAlign === 'left'} onClick={() => para({ align: 'left' })} />
            <Button icon="alignCenter" title="Centre (Ctrl+E)" pressed={format.paragraphAlign === 'center'} onClick={() => para({ align: 'center' })} />
            <Button icon="alignRight" title="Align right (Ctrl+R)" pressed={format.paragraphAlign === 'right'} onClick={() => para({ align: 'right' })} />
            <Button icon="alignJustify" title="Justify (Ctrl+J)" pressed={format.paragraphAlign === 'both'} onClick={() => para({ align: 'both' })} />
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
            <Soon icon="wand" label="" why="Paragraph shading needs w:shd on the paragraph, which the engine does not write yet." />
            <Soon icon="grid" label="" why="Paragraph borders need w:pBdr, which the engine does not write yet." />
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
            <Button icon="find" label="Find" title="Find (Ctrl+F)" onClick={() => openDialog('find')} />
            <Button icon="refresh" label="Replace" title="Replace (Ctrl+H)" onClick={() => openDialog('find')} />
            <Button icon="check" label="Select" title="Select all (Ctrl+A)" onClick={() => dispatch({ op: 'selectAll' })} />
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
          </Group>
          <Group label="Illustrations">
            <Button tall icon="picture" label="Pictures" onClick={insertPicture} />
            <Button tall icon="shape" label="Shapes" onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', preset, widthPx: 200, heightPx: 120 }) })))} />
            <Soon tall icon="star" label="Icons" why="An icon library is a download; this suite makes no requests it has not declared." />
            <Soon tall icon="shape" label="3D Models" why="3D models need a renderer the suite does not have." />
            <Soon tall icon="grid" label="SmartArt" why="SmartArt is a diagram grammar the engine does not write yet." />
            <Button tall icon="chart" label="Chart" title="A chart drawn from the table the caret is in — put the caret in a table first" onClick={(e) => menu.open(e, CHARTS.map(([k, l]) => ({ label: l, icon: 'chart', run: () => dispatch({ op: 'insertChart', kind: k }) })))} />
            <Soon tall icon="picture" label="Screenshot" why="Screenshots need a capture permission this build does not ask for." />
          </Group>
          <Group label="Media">
            <Soon tall icon="video" label="Online Videos" why="Embedding a video is a network request; this suite makes none it has not declared." />
          </Group>
          <Group label="Links">
            <Button tall icon="link" label="Link" title="Link (Ctrl+K)" pressed={Boolean(format.link)} onClick={() => openDialog('link')} />
            <Soon icon="flag" label="Bookmark" why="Bookmarks need w:bookmarkStart, which the engine does not write yet." />
            <Soon icon="link" label="Cross-reference" why="Cross-references need bookmarks and fields, neither written yet." />
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
            <Soon tall icon="textbox" label="Text Box" why="A text box is a floating shape with text; the engine draws shapes but does not lay out text in them yet." />
            <Soon icon="file" label="Quick Parts" why="Building blocks need the glossary part, which the engine does not write." />
            <Soon icon="wand" label="WordArt" why="WordArt is DrawingML text the engine does not write." />
            <Soon icon="textbox" label="Drop Cap" why="A drop cap is a framed paragraph the layout does not do yet." />
            <Soon icon="check" label="Signature Line" why="Signature lines are content controls plus a signature part; not written yet." />
            <Button icon="clock" label="Date & Time" onClick={() => openDialog('dateTime')} />
            <Soon icon="file" label="Object" why="Embedded objects (OLE) need an embeddings part the engine does not write." />
          </Group>
          <Group label="Symbols">
            <Soon tall icon="formula" label="Equation" why="Equations are OMML, which the engine reads past but does not write." />
            <Button tall icon="star" label="Symbol" onClick={() => openDialog('symbol')} />
          </Group>
        </>
      ) : null}

      {/* ── Draw ─────────────────────────────────────────────────────────── */}
      {tab === 'draw' ? (
        <>
          <Group label="Drawing Tools">
            <Soon icon="find" label="Select" why="Inking needs a pen layer the page does not have yet." />
            <Soon icon="shape" label="Lasso" why="Inking needs a pen layer the page does not have yet." />
            <Soon icon="wand" label="Eraser" why="Inking needs a pen layer the page does not have yet." />
            <Soon icon="wand" label="Pen" why="Inking needs a pen layer the page does not have yet." />
            <Soon icon="wand" label="Highlighter" why="Inking needs a pen layer the page does not have yet." />
          </Group>
          <Group label="Touch">
            <Soon tall icon="wand" label="Draw with Touch" why="Inking needs a pen layer the page does not have yet." />
          </Group>
          <Group label="Stencils">
            <Soon tall icon="minus" label="Ruler" why="Ink stencils come with inking." />
          </Group>
          <Group label="Convert">
            <Soon tall icon="shape" label="Ink to Shape" why="There is no ink to convert yet." />
            <Soon tall icon="formula" label="Ink to Maths" why="There is no ink to convert yet." />
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
            <Soon tall icon="wand" label="Themes" why="Themes rewrite the theme part; the engine reads it and does not write it yet." />
            <div className="wd-styles">
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).slice(0, 6).map((s) => (
                <button key={s.id} type="button" className={`wd-style${(format.paragraphStyle || 'Normal') === s.id ? ' on' : ''}`} style={STYLE_LOOK[s.id] || undefined} title={`Apply ${s.name || s.id} to this paragraph`} onClick={() => para({ styleId: s.id })}>
                  {s.name || s.id}
                </button>
              ))}
            </div>
            <Soon icon="contrast" label="Colours" why="Theme colours live in the theme part, not written yet." />
            <Soon icon="word" label="Fonts" why="Theme fonts live in the theme part, not written yet." />
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
            <Soon icon="wand" label="Effects" why="Theme effects live in the theme part, not written yet." />
            <Soon icon="check" label="Set as Default" why="A default template is a file the suite does not keep yet." />
          </Group>
          <Group label="Page Background">
            <Soon tall icon="shield" label="Watermark" why="A watermark is a header shape; headers hold text only so far." />
            <Soon tall icon="contrast" label="Page Colour" why="Page colour is w:background, not written yet." />
            <Soon tall icon="grid" label="Page Borders" why="Page borders are w:pgBorders, not written yet." />
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
            <Soon tall icon="grid" label="Columns" why="Columns are w:cols on the section; the layout does not flow text into columns yet." />
            <Button icon="file" label="Breaks" onClick={(e) =>
              menu.open(e, [
                { label: 'Page break', icon: 'file', run: () => dispatch({ op: 'insertPageBreak' }) },
                { label: 'Section break — not built yet', disabled: true },
              ])
            } />
            <Soon icon="listNumber" label="Line Numbers" why="Line numbering is w:lnNumType on the section, not written yet." />
            <Soon icon="minus" label="Hyphenation" why="Hyphenation is a layout setting the paginator does not honour yet." />
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
          <Group label="Arrange">
            <Soon tall icon="grid" label="Position" why="Floating layout — text wrapping round a box — is not in the paginator yet. Pictures sit in their own paragraph." />
            <Soon tall icon="grid" label="Wrap Text" why="Floating layout is not in the paginator yet." />
            <Soon tall icon="chevronUp" label="Bring Forward" why="Z-order needs floating layout." />
            <Soon tall icon="chevronDown" label="Send Backward" why="Z-order needs floating layout." />
            <Soon icon="list" label="Selection Pane" why="Comes with floating layout." />
            <Soon icon="alignLeft" label="Align" why="Comes with floating layout." />
            <Soon icon="grid" label="Group" why="Comes with floating layout." />
            <Soon icon="rotate" label="Rotate" why="Comes with floating layout." />
          </Group>
        </>
      ) : null}

      {/* ── References ───────────────────────────────────────────────────── */}
      {tab === 'references' ? (
        <>
          <Group label="Table of Contents">
            <Button tall icon="listBullet" label="Table of Contents" title="Built from the headings in this document, as text you can edit" onClick={() => act('tableOfContents')} />
            <Button icon="plus" label="Add Text" title="Make this paragraph a heading, so it appears in the table of contents" onClick={(e) =>
              menu.open(e, [
                { label: 'Do not show in table of contents', run: () => para({ styleId: null }) },
                { label: 'Level 1 (Heading 1)', run: () => para({ styleId: 'Heading1' }) },
                { label: 'Level 2 (Heading 2)', run: () => para({ styleId: 'Heading2' }) },
                { label: 'Level 3 (Heading 3)', run: () => para({ styleId: 'Heading3' }) },
              ])
            } />
            <Soon icon="refresh" label="Update Table" why="A live TOC is a field; the one inserted here is text. Insert it again to refresh." />
          </Group>
          <Group label="Footnotes">
            <Soon tall icon="file" label="Insert Footnote" why="Footnotes need the footnotes part and a reference run; on the list, and the corpus says they matter." />
            <Soon icon="file" label="Insert Endnote" why="Endnotes come with footnotes." />
            <Button icon="chevronDown" label="Next Footnote" title="Go to the next footnote reference" onClick={() => act('nextNote', 1)} />
            <Button icon="eye" label="Show Notes" title="Scroll to the footnotes under the body" onClick={() => act('showNotes')} />
          </Group>
          <Group label="Citations & Bibliography">
            <Soon tall icon="reply" label="Insert Citation" why="Citations need the bibliography part, not written yet." />
            <Soon icon="list" label="Manage Sources" why="Comes with citations." />
            <Soon icon="listNumber" label="Style" why="Comes with citations." />
            <Soon icon="listBullet" label="Bibliography" why="Comes with citations." />
          </Group>
          <Group label="Captions">
            <Soon tall icon="textbox" label="Insert Caption" why="A caption is a SEQ field; fields are not written yet." />
            <Soon icon="listBullet" label="Table of Figures" why="Comes with captions." />
            <Soon icon="link" label="Cross-reference" why="Needs bookmarks and fields." />
          </Group>
          <Group label="Index">
            <Soon tall icon="flag" label="Mark Entry" why="Index entries are XE fields, not written yet." />
            <Soon icon="listBullet" label="Insert Index" why="Comes with index entries." />
          </Group>
          <Group label="Table of Authorities">
            <Soon tall icon="flag" label="Mark Citation" why="TA fields, not written yet." />
            <Soon icon="listBullet" label="Insert Table of Authorities" why="Comes with citations." />
          </Group>
        </>
      ) : null}

      {/* ── Mailings ─────────────────────────────────────────────────────── */}
      {tab === 'mailings' ? (
        <>
          <Group label="Create">
            <Soon tall icon="mail" label="Envelopes" why="Envelope printing is a page-setup preset plus a label layout; not built yet." />
            <Soon tall icon="grid" label="Labels" why="Label sheets are a table layout to a stock size; not built yet." />
          </Group>
          <Group label="Start Mail Merge">
            <Soon tall icon="mail" label="Start Mail Merge" why="Mail merge is on the list: a data source, merge fields, and a run through the Mail app. None of it is built yet." />
            <Soon tall icon="reply" label="Select Recipients" why="Comes with mail merge." />
            <Soon tall icon="list" label="Edit Recipient List" why="Comes with mail merge." />
          </Group>
          <Group label="Write & Insert Fields">
            <Soon tall icon="wand" label="Highlight Merge Fields" why="Comes with mail merge." />
            <Soon tall icon="textbox" label="Address Block" why="Comes with mail merge." />
            <Soon tall icon="textbox" label="Greeting Line" why="Comes with mail merge." />
            <Soon tall icon="plus" label="Insert Merge Field" why="Comes with mail merge." />
          </Group>
          <Group label="Preview Results">
            <Soon tall icon="eye" label="Preview Results" why="Comes with mail merge." />
          </Group>
          <Group label="Finish">
            <Soon tall icon="send" label="Finish & Merge" why="Comes with mail merge." />
          </Group>
        </>
      ) : null}

      {/* ── Review ───────────────────────────────────────────────────────── */}
      {tab === 'review' ? (
        <>
          <Group label="Proofing">
            <Button tall icon="check" label="Spelling" pressed={view.spell !== false} title="Underline misspellings as you type (the system's dictionary)" onClick={() => act('toggleSpell')} />
            <Soon tall icon="word" label="Thesaurus" why="A thesaurus is a data file the suite does not ship yet." />
            <Button tall icon="listNumber" label="Word Count" onClick={() => openDialog('wordCount')} />
          </Group>
          <Group label="Speech">
            <Button tall icon="volume" label="Read Aloud" pressed={Boolean(view.reading)} onClick={() => act('readAloud')} />
          </Group>
          <Group label="Accessibility">
            <Soon tall icon="shield" label="Check Accessibility" why="An accessibility check — alt text, heading order, contrast — is on the list." />
          </Group>
          <Group label="Language">
            <Soon tall icon="globe" label="Translate" why="Translation is a network service; this suite makes no requests it has not declared." />
            <Soon tall icon="globe" label="Language" why="Proofing language is w:lang on runs, not written yet." />
          </Group>
          <Group label="Comments">
            <Button tall icon="reply" label="New Comment" onClick={() => openDialog('comment')} />
            <Soon tall icon="trash" label="Delete" why="Deleting a comment means removing its range marks; the engine writes comments and does not remove them yet." />
            <Button icon="chevronUp" label="Previous" disabled={!comments.length} onClick={() => act('comment', -1)} />
            <Button icon="chevronDown" label="Next" disabled={!comments.length} onClick={() => act('comment', 1)} />
            <Button icon="listBullet" label={comments.length ? `Show (${comments.length})` : 'Show Comments'} disabled={!comments.length} onClick={() => openDialog('comments')} />
          </Group>
          <Group label="Tracking">
            <Soon tall icon="eye" label="Track Changes" why="Tracked changes are read and shown; recording them (w:ins / w:del on every edit) is not built yet." />
            <Button icon="list" label="Reviewing Pane" title="Every tracked change in this document" disabled={!model?.blocks?.some((b) => b.tracked)} onClick={() => openDialog('tracked')} />
          </Group>
          <Group label="Changes">
            <Soon tall icon="check" label="Accept" why="Comes with recording tracked changes." />
            <Soon tall icon="close" label="Reject" why="Comes with recording tracked changes." />
          </Group>
          <Group label="Compare">
            <Soon tall icon="copy" label="Compare" why="Comparing two documents is a diff over blocks; on the list." />
          </Group>
          <Group label="Protect">
            <Soon tall icon="lock" label="Restrict Editing" why="Document protection is w:documentProtection; not written yet." />
          </Group>
        </>
      ) : null}

      {/* ── View ─────────────────────────────────────────────────────────── */}
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
            <Soon tall icon="eye" label="Immersive Reader" why="Line focus and syllable spacing are on the list." />
          </Group>
          <Group label="Show">
            <Button icon="minus" label="Ruler" pressed={Boolean(view.ruler)} onClick={() => act('toggleRuler')} />
            <Soon icon="grid" label="Gridlines" why="Drawing gridlines come with floating layout." />
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
            <Soon icon="grid" label="Arrange All" why="Window tiling is on the list." />
            <Soon icon="minus" label="Split" why="A split view is on the list." />
            <Soon icon="copy" label="Side by Side" why="Comes with window tiling." />
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
