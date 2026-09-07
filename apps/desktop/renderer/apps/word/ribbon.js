// The Word ribbon.
//
// Home, Insert, Layout, References, Review, View — where a person arriving from
// Word, Writer or Docs expects to find things.
//
// As with the spreadsheet, most of what is new here is not new work: links,
// comments, headers and footers, page setup, table editing and the style
// catalogue were all in the engine with nothing calling them.

import React from 'react';
import { Ribbon, Group, Button, Separator, Select } from '@rutba/office-ui';

const FONTS = ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Consolas'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

const TEXT_COLOURS = [
  ['000000', 'Automatic'], ['444444', 'Dark grey'], ['767171', 'Grey'],
  ['C00000', 'Dark red'], ['E03131', 'Red'], ['E08B2B', 'Orange'],
  ['0F9D58', 'Green'], ['0D8F6F', 'Teal'], ['2B5FD9', 'Blue'], ['7B5CD6', 'Purple'],
];

const HIGHLIGHTS = [
  [null, 'No highlight'], ['FFFF00', 'Yellow'], ['00FF00', 'Bright green'], ['00FFFF', 'Turquoise'],
  ['FF00FF', 'Pink'], ['C0C0C0', 'Grey'], ['FFC7CE', 'Light red'],
];

const SPACING = [
  [1, 'Single'], [1.15, '1.15'], [1.5, '1.5'], [2, 'Double'],
];

const PAGE_SIZES = [
  ['a4', 'A4 — 210 × 297 mm'],
  ['letter', 'Letter — 8.5 × 11 in'],
  ['legal', 'Legal — 8.5 × 14 in'],
  ['a3', 'A3 — 297 × 420 mm'],
  ['a5', 'A5 — 148 × 210 mm'],
];

const MARGINS = [
  ['normal', 'Normal — 2.54 cm', 1440],
  ['narrow', 'Narrow — 1.27 cm', 720],
  ['moderate', 'Moderate — 2.54 / 1.91 cm', 1080],
  ['wide', 'Wide — 5.08 cm', 2880],
];

export default function WordRibbon({
  tab, setTab, doc, model, dispatch, commands, shell, menu, save, openFile, exportAs, openDialog,
}) {
  const format = model?.format || {};
  const styles = model?.styles || [];
  const section = model?.section || null;
  const comments = model?.comments || [];
  const landscape = section?.orientation === 'landscape' || (section?.width && section?.height && section.width > section.height);

  const run = (delta) => dispatch({ op: 'setRunFormat', delta });
  const para = (delta) => dispatch({ op: 'setParagraphFormat', delta });

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
            <Select value={format.font || 'Calibri'} onChange={(e) => run({ font: e.target.value })} style={{ width: 130 }} title="Font">
              {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
              {FONTS.includes(format.font) || !format.font ? null : <option value={format.font}>{format.font}</option>}
            </Select>
            <Select value={String(format.size || 11)} onChange={(e) => run({ size: Number(e.target.value) })} style={{ width: 56 }} title="Size">
              {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
            </Select>
            <Separator />
            <Button icon="bold" title="Bold" pressed={format.bold} onClick={() => dispatch({ op: 'toggleFormat', tag: 'bold' })} />
            <Button icon="italic" title="Italic" pressed={format.italic} onClick={() => dispatch({ op: 'toggleFormat', tag: 'italic' })} />
            <Button icon="underline" title="Underline" pressed={format.underline} onClick={() => dispatch({ op: 'toggleFormat', tag: 'underline' })} />
            <Button icon="strike" title="Strikethrough" pressed={format.strike} onClick={() => dispatch({ op: 'toggleFormat', tag: 'strike' })} />
            <Separator />
            <Button icon="contrast" title="Text colour" onClick={(e) => colourMenu(e, 'colour', TEXT_COLOURS)} />
            <Button icon="wand" title="Highlight" onClick={(e) => colourMenu(e, 'highlight', HIGHLIGHTS)} />
            <Button icon="undo" title="Clear formatting" onClick={() => dispatch({ op: 'clearFormat' })} />
          </Group>

          <Group label="Paragraph">
            <Button icon="listBullet" title="Bulleted list" onClick={() => para({ list: 'bullet' })} />
            <Button icon="listNumber" title="Numbered list" onClick={() => para({ list: 'number' })} />
            <Separator />
            <Button icon="alignLeft" title="Align left" pressed={format.align === 'left'} onClick={() => para({ align: 'left' })} />
            <Button icon="alignCenter" title="Centre" pressed={format.align === 'center'} onClick={() => para({ align: 'center' })} />
            <Button icon="alignRight" title="Align right" pressed={format.align === 'right'} onClick={() => para({ align: 'right' })} />
            <Button icon="alignJustify" title="Justify" pressed={format.align === 'both'} onClick={() => para({ align: 'both' })} />
            <Separator />
            <Button
              icon="listNumber"
              title="Line spacing"
              onClick={(e) => menu.open(e, SPACING.map(([v, l]) => ({ label: l, run: () => para({ lineSpacing: v }) })))}
            />
            <Button icon="chevronRight" title="Increase indent" onClick={() => para({ indent: (format.indent || 0) + 720 })} />
            <Button icon="chevronLeft" title="Decrease indent" onClick={() => para({ indent: Math.max(0, (format.indent || 0) - 720) })} />
          </Group>

          <Group label="Styles">
            <Select
              value={format.style || 'Normal'}
              onChange={(e) => para({ style: e.target.value })}
              style={{ width: 150 }}
              title="Paragraph style"
            >
              {(styles.length ? styles : [{ id: 'Normal', name: 'Normal' }]).map((s) => (
                <option key={s.id || s.name} value={s.id || s.name}>{s.name || s.id}</option>
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
            <Button tall icon="picture" label="Picture" onClick={() => commands['insert.image']?.run?.()} />
            <Button tall icon="shape" label="Shape" onClick={(e) =>
              menu.open(e, ['rect', 'ellipse', 'roundRect', 'triangle', 'arrow', 'star'].map((g) => ({
                label: g[0].toUpperCase() + g.slice(1),
                icon: 'shape',
                run: () => dispatch({ op: 'insertShape', geometry: g, text: '' }),
              })))
            } />
            <Button tall icon="chart" label="Chart" onClick={(e) =>
              menu.open(e, ['column', 'bar', 'line', 'pie'].map((k) => ({
                label: k[0].toUpperCase() + k.slice(1),
                icon: 'chart',
                run: () => dispatch({ op: 'insertChart', kind: k }),
              })))
            } />
          </Group>
          <Group label="Links">
            <Button tall icon="link" label="Link" onClick={() => openDialog('link')} />
          </Group>
          <Group label="Header and footer">
            <Button icon="chevronUp" label="Header" onClick={() => openDialog('header')} />
            <Button icon="chevronDown" label="Footer" onClick={() => openDialog('footer')} />
          </Group>
        </>
      ) : tab === 'layout' ? (
        <>
          <Group label="Page setup">
            <Button
              tall
              icon="file"
              label="Size"
              onClick={(e) => menu.open(e, PAGE_SIZES.map(([v, l]) => ({ label: l, run: () => dispatch({ op: 'setPageSetup', spec: { size: v } }) })))}
            />
            <Button
              tall
              icon="rotate"
              label={landscape ? 'Landscape' : 'Portrait'}
              onClick={() => dispatch({ op: 'setPageSetup', spec: { orientation: landscape ? 'portrait' : 'landscape' } })}
            />
            <Button
              tall
              icon="crop"
              label="Margins"
              onClick={(e) =>
                menu.open(e, MARGINS.map(([v, l, twips]) => ({
                  label: l,
                  run: () => dispatch({ op: 'setPageSetup', spec: { margins: { top: twips, bottom: twips, left: twips, right: twips } } }),
                })))
              }
            />
          </Group>
          <Group label="Paragraph">
            <Button icon="chevronRight" label="Indent" onClick={() => para({ indent: (format.indent || 0) + 720 })} />
            <Button icon="chevronLeft" label="Outdent" onClick={() => para({ indent: Math.max(0, (format.indent || 0) - 720) })} />
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
            <Button
              icon="file"
              label="Page number"
              title="Insert a page-number field into the footer"
              onClick={() => openDialog('pageNumber')}
            />
          </Group>
          <Group label="Links">
            <Button tall icon="link" label="Hyperlink" onClick={() => openDialog('link')} />
          </Group>
        </>
      ) : tab === 'review' ? (
        <>
          <Group label="Comments">
            <Button tall icon="reply" label="New comment" onClick={() => openDialog('comment')} />
            <Button
              tall
              icon="listBullet"
              label={comments.length ? `${comments.length} comment${comments.length === 1 ? '' : 's'}` : 'No comments'}
              disabled={!comments.length}
              onClick={() => openDialog('comments')}
            />
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
