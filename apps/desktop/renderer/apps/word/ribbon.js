// The Word ribbon.
//
// Home, Insert, Layout, References, Review, View — where a person arriving from
// Word, Writer or Docs expects to find things.
//
// Every control here speaks the engine's own vocabulary, and the first version
// of this file did not: it sent `font` where the engine listens for `fontName`,
// `indent` where it wants `indentDelta`, `style` where it wants `styleId`, and
// `a4` where the paper table says `A4`. Each of those was a silent no-op — a
// button that pressed and did nothing — and nothing caught it, because the
// checks drove the engine directly and never clicked a button. They do now.
//
// The state the ribbon shows comes from `model.format`, which is what the
// engine reports at the caret: `fontName`, `fontSize`, `fontColour`,
// `highlight`, `paragraphAlign`, `listType`, `paragraphStyle`, `lineSpacing`.

import React from 'react';
import { Ribbon, Group, Button, Separator, Select } from '@rutba/office-ui';

const FONTS = ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Consolas'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

/** Hex without the hash: that is how OOXML stores w:color and how the engine reads it back. */
const TEXT_COLOURS = [
  [null, 'Automatic'],
  ['444444', 'Dark grey'], ['767171', 'Grey'],
  ['C00000', 'Dark red'], ['E03131', 'Red'], ['E08B2B', 'Orange'],
  ['0F9D58', 'Green'], ['0D8F6F', 'Teal'], ['2B5FD9', 'Blue'], ['7B5CD6', 'Purple'],
];

/** Word's highlighter has sixteen named colours, not a palette. These are the ones people use. */
const HIGHLIGHTS = [
  [null, 'No highlight'], ['yellow', 'Yellow'], ['green', 'Bright green'], ['cyan', 'Turquoise'],
  ['magenta', 'Pink'], ['lightGray', 'Grey'], ['red', 'Red'], ['blue', 'Blue'],
];

const SPACING = [[1, 'Single'], [1.15, '1.15'], [1.5, '1.5'], [2, 'Double']];

/** The engine's paper table, by its own names. */
const PAGE_SIZES = [['A4', 'A4 — 210 × 297 mm'], ['Letter', 'Letter — 8.5 × 11 in'], ['Legal', 'Legal — 8.5 × 14 in']];

/** The engine's margin presets, by their own names. */
const MARGINS = [['normal', 'Normal — 2.54 cm'], ['narrow', 'Narrow — 1.27 cm'], ['wide', 'Wide — 5.08 cm at the sides']];

/** Every preset the document engine can draw. Anything else is refused by name. */
const SHAPES = [
  ['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Ellipse'], ['triangle', 'Triangle'],
  ['diamond', 'Diamond'], ['rightArrow', 'Arrow right'], ['leftArrow', 'Arrow left'], ['upArrow', 'Arrow up'],
  ['downArrow', 'Arrow down'], ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'], ['star5', 'Star'],
  ['plus', 'Plus'], ['chevron', 'Chevron'], ['parallelogram', 'Parallelogram'], ['trapezoid', 'Trapezoid'], ['line', 'Line'],
];

export default function WordRibbon({ tab, setTab, doc, model, dispatch, commands, shell, menu, save, openFile, exportAs, openDialog, insertPicture }) {
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

  return (
    <Ribbon
      tabs={[
        { id: 'home', label: 'Home' },
        { id: 'insert', label: 'Insert' },
        { id: 'layout', label: 'Layout' },
        { id: 'references', label: 'References' },
        { id: 'review', label: 'Review' },
        { id: 'view', label: 'View' },
      ]}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <Button icon="save" title="Save" onClick={() => save(false)} />
          <Button icon="undo" title="Undo" disabled={!doc?.canUndo} onClick={() => commands['edit.undo'].run()} />
          <Button icon="redo" title="Redo" disabled={!doc?.canRedo} onClick={() => commands['edit.redo'].run()} />
          <Button icon="print" title="Print" onClick={() => commands['file.print']?.run?.()} />
        </>
      }
    >
      {tab === 'home' ? (
        <>
          <Group label="Clipboard">
            <Button tall icon="paste" label="Paste" onClick={() => commands['edit.paste']?.run?.()} />
            <Button icon="copy" label="Copy" onClick={() => commands['edit.copy']?.run?.()} />
            <Button icon="cut" label="Cut" onClick={() => commands['edit.cut']?.run?.()} />
          </Group>

          <Group label="Font">
            <Select value={format.fontName || 'Calibri'} onChange={(e) => run({ fontName: e.target.value })} style={{ width: 130 }} title="Font">
              {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
              {format.fontName && !FONTS.includes(format.fontName) ? <option value={format.fontName}>{format.fontName}</option> : null}
            </Select>
            <Select value={String(format.fontSize || 11)} onChange={(e) => run({ fontSize: Number(e.target.value) })} style={{ width: 56 }} title="Size">
              {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
              {format.fontSize && !SIZES.includes(format.fontSize) ? <option value={String(format.fontSize)}>{format.fontSize}</option> : null}
            </Select>
            <Separator />
            <Button icon="bold" title="Bold (Ctrl+B)" pressed={format.bold} onClick={() => dispatch({ op: 'toggleFormat', tag: 'bold' })} />
            <Button icon="italic" title="Italic (Ctrl+I)" pressed={format.italic} onClick={() => dispatch({ op: 'toggleFormat', tag: 'italic' })} />
            <Button icon="underline" title="Underline (Ctrl+U)" pressed={format.underline} onClick={() => dispatch({ op: 'toggleFormat', tag: 'underline' })} />
            <Button icon="strike" title="Strikethrough" pressed={format.strike} onClick={() => dispatch({ op: 'toggleFormat', tag: 'strike' })} />
            <Separator />
            <Button icon="contrast" title="Text colour" onClick={(e) => colourMenu(e, 'fontColour', TEXT_COLOURS)} />
            <Button icon="wand" title="Highlight" pressed={Boolean(format.highlight)} onClick={(e) => colourMenu(e, 'highlight', HIGHLIGHTS)} />
            <Button icon="undo" title="Clear formatting" onClick={() => dispatch({ op: 'clearFormat' })} />
          </Group>

          <Group label="Paragraph">
            <Button icon="listBullet" title="Bulleted list" pressed={format.listType === 'bullet'} onClick={() => toggleList('bullet')} />
            <Button icon="listNumber" title="Numbered list" pressed={format.listType === 'number'} onClick={() => toggleList('number')} />
            <Separator />
            <Button icon="alignLeft" title="Align left" pressed={format.paragraphAlign === 'left'} onClick={() => para({ align: 'left' })} />
            <Button icon="alignCenter" title="Centre" pressed={format.paragraphAlign === 'center'} onClick={() => para({ align: 'center' })} />
            <Button icon="alignRight" title="Align right" pressed={format.paragraphAlign === 'right'} onClick={() => para({ align: 'right' })} />
            <Button icon="alignJustify" title="Justify" pressed={format.paragraphAlign === 'both'} onClick={() => para({ align: 'both' })} />
            <Separator />
            <Button
              icon="listNumber"
              title={`Line spacing${format.lineSpacing ? ` — ${format.lineSpacing}` : ''}`}
              onClick={(e) => menu.open(e, SPACING.map(([v, l]) => ({ label: l, icon: format.lineSpacing === v ? 'check' : undefined, run: () => para({ lineSpacing: v }) })))}
            />
            <Button icon="chevronRight" title="Increase indent" onClick={() => para({ indentDelta: 1 })} />
            <Button icon="chevronLeft" title="Decrease indent" disabled={!format.indentLevel} onClick={() => para({ indentDelta: -1 })} />
          </Group>

          <Group label="Styles">
            <Select
              value={format.paragraphStyle || 'Normal'}
              onChange={(e) => para({ styleId: e.target.value === 'Normal' && !styles.some((s) => s.id === 'Normal') ? null : e.target.value })}
              style={{ width: 150 }}
              title="Paragraph style"
            >
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).map((s) => (
                <option key={s.id} value={s.id}>{s.name || s.id}</option>
              ))}
            </Select>
          </Group>

          <Group label="Editing">
            <Button icon="find" label="Find" onClick={() => openDialog('find')} />
          </Group>
        </>
      ) : tab === 'insert' ? (
        <>
          <Group label="Pages">
            <Button tall icon="file" label="Page break" onClick={() => dispatch({ op: 'insertPageBreak' })} />
          </Group>
          <Group label="Tables">
            <Button tall icon="table" label="Table" onClick={() => openDialog('table')} />
          </Group>
          <Group label="Illustrations">
            <Button tall icon="picture" label="Picture" onClick={insertPicture} />
            <Button
              tall
              icon="shape"
              label="Shape"
              onClick={(e) => menu.open(e, SHAPES.map(([preset, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', preset, widthPx: 200, heightPx: 120 }) })))}
            />
            <Button
              tall
              icon="chart"
              label="Chart"
              title="A chart is drawn from a table: put the caret in a table first"
              onClick={(e) => menu.open(e, ['column', 'bar', 'line', 'pie'].map((k) => ({ label: k[0].toUpperCase() + k.slice(1), icon: 'chart', run: () => dispatch({ op: 'insertChart', kind: k }) })))}
            />
          </Group>
          <Group label="Links">
            <Button tall icon="link" label="Link" pressed={Boolean(format.link)} onClick={() => openDialog('link')} />
          </Group>
          <Group label="Comments">
            <Button tall icon="reply" label="Comment" onClick={() => openDialog('comment')} />
          </Group>
          <Group label="Header and footer">
            <Button icon="chevronUp" label="Header" onClick={() => openDialog('header')} />
            <Button icon="chevronDown" label="Footer" onClick={() => openDialog('footer')} />
          </Group>
        </>
      ) : tab === 'layout' ? (
        <>
          <Group label="Page setup">
            <Button tall icon="file" label="Size" onClick={(e) => menu.open(e, PAGE_SIZES.map(([v, l]) => ({ label: l, run: () => dispatch({ op: 'setPageSetup', spec: { size: v } }) })))} />
            <Button tall icon="rotate" label={landscape ? 'Landscape' : 'Portrait'} title="Switch orientation" onClick={() => dispatch({ op: 'setPageSetup', spec: { orientation: landscape ? 'portrait' : 'landscape' } })} />
            <Button tall icon="crop" label="Margins" onClick={(e) => menu.open(e, MARGINS.map(([v, l]) => ({ label: l, run: () => dispatch({ op: 'setPageSetup', spec: { margins: v } }) })))} />
          </Group>
          <Group label="Paragraph">
            <Button icon="chevronRight" label="Indent" onClick={() => para({ indentDelta: 1 })} />
            <Button icon="chevronLeft" label="Outdent" disabled={!format.indentLevel} onClick={() => para({ indentDelta: -1 })} />
            <Button icon="listNumber" label="Spacing" onClick={(e) => menu.open(e, SPACING.map(([v, l]) => ({ label: l, run: () => para({ lineSpacing: v }) })))} />
          </Group>
          <Group label="Export">
            <Button tall icon="pdf" label="PDF" onClick={() => exportAs('pdf')} />
            <Button icon="export" label="Markdown" onClick={() => exportAs('md')} />
            <Button icon="export" label="Text" onClick={() => exportAs('txt')} />
          </Group>
        </>
      ) : tab === 'references' ? (
        <>
          <Group label="Headers and footers">
            <Button tall icon="chevronUp" label="Edit header" onClick={() => openDialog('header')} />
            <Button tall icon="chevronDown" label="Edit footer" onClick={() => openDialog('footer')} />
          </Group>
          <Group label="Fields">
            <Button icon="file" label="Page number" title="Put a page-number field in the footer" onClick={() => openDialog('pageNumber')} />
          </Group>
          <Group label="Links">
            <Button tall icon="link" label="Hyperlink" onClick={() => openDialog('link')} />
          </Group>
        </>
      ) : tab === 'review' ? (
        <>
          <Group label="Comments">
            <Button tall icon="reply" label="New comment" onClick={() => openDialog('comment')} />
            <Button tall icon="listBullet" label={comments.length ? `${comments.length} comment${comments.length === 1 ? '' : 's'}` : 'No comments'} disabled={!comments.length} onClick={() => openDialog('comments')} />
          </Group>
          <Group label="Proofing">
            <Button tall icon="check" label="Word count" onClick={() => openDialog('wordCount')} />
          </Group>
          <Group label="Editing">
            <Button tall icon="find" label="Find and replace" onClick={() => openDialog('find')} />
          </Group>
        </>
      ) : (
        <>
          <Group label="Window">
            <Button tall icon="maximize" label="Full screen" onClick={() => shell.win.fullscreen({})} />
          </Group>
          <Group label="Zoom">
            <Button icon="zoomOut" label="Out" onClick={() => shell.win.zoom({ delta: -0.1 })} />
            <Button icon="zoomIn" label="In" onClick={() => shell.win.zoom({ delta: 0.1 })} />
            <Button icon="check" label="100%" onClick={() => shell.win.zoom({ reset: true })} />
          </Group>
          <Group label="File">
            <Button tall icon="new" label="New" onClick={() => shell.win.create({ app: 'word' })} />
            <Button tall icon="open" label="Open" onClick={openFile} />
            <Button tall icon="save" label="Save" onClick={() => save(false)} />
            <Button icon="print" label="Print" onClick={() => commands['file.print']?.run?.()} />
          </Group>
        </>
      )}
    </Ribbon>
  );
}
