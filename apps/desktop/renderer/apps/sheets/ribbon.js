// The Worksheets ribbon.
//
// Laid out the way every spreadsheet has been laid out since 2007, because a
// person arriving from Excel, Calc or Sheets should not have to learn where
// anything is: Home, Insert, Formulas, Data, Review, View.
//
// Every button here drives an engine capability that already existed. The Home
// tab in particular was the largest gap in the suite — a spreadsheet whose
// cells could not be made bold, given a currency format or a fill colour, when
// the engine had done all three for months.

import React from 'react';
import { Ribbon, Group, Button, Separator, Icon, Select } from '@rutba/office-ui';
import { NUMBER_FORMATS } from './dialogs.js';

/** The palette a toolbar offers before it offers a colour picker. */
const SWATCHES = [
  ['#000000', 'Black'], ['#444444', 'Dark grey'], ['#888888', 'Grey'], ['#ffffff', 'White'],
  ['#c00000', 'Dark red'], ['#e03131', 'Red'], ['#e08b2b', 'Orange'], ['#e0a800', 'Amber'],
  ['#0f9d58', 'Green'], ['#0d8f6f', 'Teal'], ['#2b5fd9', 'Blue'], ['#7b5cd6', 'Purple'],
];

const FILLS = [
  [null, 'No fill'], ['FFF3BF', 'Light amber'], ['D3F9D8', 'Light green'], ['D0EBFF', 'Light blue'],
  ['FFE3E3', 'Light red'], ['E9ECEF', 'Light grey'], ['FFD8A8', 'Light orange'], ['E5DBFF', 'Light purple'],
];

const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Consolas'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];

/**
 * A border delta names edges. The engine keeps the edges it is not told about,
 * writes the ones it is, and deletes the ones set to null — so "all" is four
 * thin black edges and "none" is four nulls.
 */
const THIN = { style: 'thin', colour: '#000000' };
const BORDERS = [
  ['All borders', { top: THIN, bottom: THIN, left: THIN, right: THIN }],
  ['Outline', { top: THIN, bottom: THIN, left: THIN, right: THIN }],
  ['Top', { top: THIN }],
  ['Bottom', { bottom: THIN }],
  ['Left', { left: THIN }],
  ['Right', { right: THIN }],
  ['Thick bottom', { bottom: { style: 'medium', colour: '#000000' } }],
  ['No border', { top: null, bottom: null, left: null, right: null }],
];

/** What the engine can draw, by the names it draws them under. */
const CHARTS = [
  ['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie'], ['doughnut', 'Doughnut'],
];
const SHAPES = [
  ['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Ellipse'], ['line', 'Line'],
  ['triangle', 'Triangle'], ['diamond', 'Diamond'], ['rightArrow', 'Arrow right'], ['leftArrow', 'Arrow left'],
  ['upArrow', 'Arrow up'], ['downArrow', 'Arrow down'], ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'],
  ['star5', 'Star'], ['plus', 'Plus'], ['chevron', 'Chevron'], ['parallelogram', 'Parallelogram'], ['trapezoid', 'Trapezoid'],
];

export default function SheetsRibbon({
  tab, setTab, model, dispatch, commands, shell, menu, save, openFile, exportAs, doc, selectionRef, openDialog,
}) {
  const format = model?.format || {};
  const frozen = model?.frozen || { rows: 0, cols: 0 };
  const protectedSheet = Boolean(model?.protection?.sheet);

  const setFormat = (delta) => dispatch({ op: 'setFormat', delta });

  const swatchMenu = (event, key, list) =>
    menu.open(
      event,
      list.map(([value, label]) => ({
        label,
        icon: value ? undefined : 'close',
        run: () => setFormat({ [key]: value }),
      }))
    );

  return (
    <Ribbon
      tabs={[
        { id: 'home', label: 'Home' },
        { id: 'insert', label: 'Insert' },
        { id: 'formulas', label: 'Formulas' },
        { id: 'data', label: 'Data' },
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
          <Button icon="sum" title="AutoSum" onClick={() => commands['sheet.autoSum'].run()} />
        </>
      }
    >
      {tab === 'home' ? (
        <>
          <Group label="Clipboard">
            <Button tall icon="paste" label="Paste" onClick={async () => dispatch({ op: 'paste', text: await shell.clipboard.readText() })} />
            <Button icon="copy" label="Copy" onClick={() => commands['edit.copy'].run()} />
            <Button icon="cut" label="Cut" onClick={() => commands['edit.cut']?.run?.()} />
            <Button icon="wand" label="Format" title="Copy formatting to another selection" onClick={() => dispatch({ op: 'formatBrush' })} />
          </Group>

          <Group label="Font">
            <Select
              value={format.fontName || 'Calibri'}
              onChange={(e) => setFormat({ fontName: e.target.value })}
              style={{ width: 118 }}
              title="Font"
            >
              {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
            </Select>
            <Select
              value={String(format.fontSize || 11)}
              onChange={(e) => setFormat({ fontSize: Number(e.target.value) })}
              style={{ width: 56 }}
              title="Size"
            >
              {SIZES.map((s) => <option key={s} value={String(s)}>{s}</option>)}
            </Select>
            <Separator />
            <Button icon="bold" title="Bold" pressed={format.bold} onClick={() => setFormat({ bold: 'toggle' })} />
            <Button icon="italic" title="Italic" pressed={format.italic} onClick={() => setFormat({ italic: 'toggle' })} />
            <Button icon="underline" title="Underline" pressed={format.underline} onClick={() => setFormat({ underline: 'toggle' })} />
            <Button icon="strike" title="Strikethrough" pressed={format.strike} onClick={() => setFormat({ strike: 'toggle' })} />
            <Separator />
            <Button icon="contrast" title="Text colour" onClick={(e) => swatchMenu(e, 'fontColour', SWATCHES)} />
            <Button icon="wand" title="Fill colour" onClick={(e) => swatchMenu(e, 'fill', FILLS.map(([c, l]) => [c ? `#${c}` : null, l]))} />
            <Button icon="grid" title="Borders" onClick={(e) => menu.open(e, BORDERS.map(([label, edges]) => ({ label, run: () => setFormat({ border: edges }) })))} />
          </Group>

          <Group label="Alignment">
            <Button icon="alignLeft" title="Align left" pressed={format.align === 'left'} onClick={() => setFormat({ align: 'left' })} />
            <Button icon="alignCenter" title="Centre" pressed={format.align === 'center'} onClick={() => setFormat({ align: 'center' })} />
            <Button icon="alignRight" title="Align right" pressed={format.align === 'right'} onClick={() => setFormat({ align: 'right' })} />
            <Separator />
            <Button icon="chevronUp" title="Align top" pressed={format.valign === 'top'} onClick={() => setFormat({ valign: 'top' })} />
            <Button icon="minus" title="Align middle" pressed={format.valign === 'center'} onClick={() => setFormat({ valign: 'center' })} />
            <Button icon="chevronDown" title="Align bottom" pressed={format.valign === 'bottom'} onClick={() => setFormat({ valign: 'bottom' })} />
            <Separator />
            <Button icon="listBullet" label="Wrap" pressed={format.wrap} onClick={() => setFormat({ wrap: !format.wrap })} />
            <Button icon="table" label="Merge" onClick={() => dispatch({ op: 'merge' })} />
            <Button icon="minus" label="Unmerge" onClick={() => dispatch({ op: 'unmerge' })} />
          </Group>

          <Group label="Number">
            <Select
              value={format.numberFormat || 'General'}
              onChange={(e) => setFormat({ numberFormat: e.target.value })}
              style={{ width: 148 }}
              title="Number format"
            >
              {NUMBER_FORMATS.map((f) => <option key={f.label} value={f.code}>{f.label}</option>)}
              {NUMBER_FORMATS.some((f) => f.code === (format.numberFormat || 'General')) ? null : (
                <option value={format.numberFormat}>{format.numberFormat}</option>
              )}
            </Select>
            <Separator />
            {/* The mark is the button. An icon beside a currency symbol says nothing. */}
            <Button title="Currency" onClick={() => setFormat({ numberFormat: '"£"#,##0.00' })} label="£" />
            <Button title="Percent" onClick={() => setFormat({ numberFormat: '0%' })} label="%" />
            <Button title="Thousands separator" onClick={() => setFormat({ numberFormat: '#,##0.00' })} label="," />
          </Group>

          <Group label="Styles">
            <Button tall icon="wand" label="Conditional" onClick={() => openDialog('conditional')} />
          </Group>

          <Group label="Cells">
            <Button icon="plus" label="Row" onClick={() => commands['sheet.insertRow'].run()} />
            <Button icon="minus" label="Row" onClick={() => commands['sheet.deleteRow'].run()} />
            <Button icon="plus" label="Column" onClick={() => commands['sheet.insertCol'].run()} />
            <Button icon="minus" label="Column" onClick={() => commands['sheet.deleteCol'].run()} />
          </Group>

          <Group label="Editing">
            <Button icon="sum" label="AutoSum" onClick={() => commands['sheet.autoSum'].run()} />
            <Button icon="sort" label="Sort" onClick={() => commands['sheet.sortAsc'].run()} />
            <Button icon="find" label="Find" onClick={() => openDialog('find')} />
            <Button icon="close" label="Clear" onClick={() => dispatch({ op: 'clear' })} />
          </Group>
        </>
      ) : tab === 'insert' ? (
        <>
          <Group label="Tables">
            <Button tall icon="table" label="Pivot table" onClick={() => openDialog('pivot')} />
            <Button tall icon="filter" label="Filter" pressed={model?.filtered} onClick={() => dispatch({ op: 'autoFilter' })} />
          </Group>
          <Group label="Charts">
            {CHARTS.map(([kind, label]) => (
              <Button key={kind} tall icon="chart" label={label} title={`${label} chart from the data around the selection`} onClick={() => dispatch({ op: 'insertChart', kind })} />
            ))}
          </Group>
          <Group label="Illustrations">
            <Button
              tall
              icon="shape"
              label="Shape"
              onClick={(e) => menu.open(e, SHAPES.map(([geometry, label]) => ({ label, icon: 'shape', run: () => dispatch({ op: 'insertShape', geometry, text: '' }) })))}
            />
          </Group>
          <Group label="Names">
            <Button tall icon="find" label="Define name" onClick={() => openDialog('names')} />
          </Group>
        </>
      ) : tab === 'formulas' ? (
        <>
          <Group label="Function library">
            <Button tall icon="sum" label="AutoSum" onClick={(e) =>
              menu.open(e, ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((fn) => ({
                label: fn[0] + fn.slice(1).toLowerCase(),
                icon: 'sum',
                run: () => dispatch({ op: 'autoSum', fn }),
              })))
            } />
          </Group>
          <Group label="Defined names">
            <Button tall icon="find" label="Name manager" onClick={() => openDialog('names')} />
          </Group>
          <Group label="What-if analysis">
            <Button tall icon="wand" label="Goal seek" onClick={() => openDialog('goalSeek')} />
            <Button tall icon="table" label="Data table" onClick={() => openDialog('dataTable')} />
          </Group>
          <Group label="Calculation">
            <Button tall icon="refresh" label="Recalculate" title="The engine recalculates on every edit; this forces a full pass." onClick={() => dispatch({ op: 'select', row: 0, col: 0 })} />
          </Group>
        </>
      ) : tab === 'data' ? (
        <>
          <Group label="Sort and filter">
            <Button tall icon="sort" label="A → Z" onClick={() => commands['sheet.sortAsc'].run()} />
            <Button tall icon="sort" label="Z → A" onClick={() => commands['sheet.sortDesc'].run()} />
            <Button tall icon="filter" label="Filter" pressed={model?.filtered} onClick={() => dispatch({ op: 'autoFilter' })} />
          </Group>
          <Group label="Data tools">
            <Button tall icon="check" label="Validation" onClick={() => openDialog('validation')} />
            <Button tall icon="wand" label="Goal seek" onClick={() => openDialog('goalSeek')} />
            <Button tall icon="table" label="What-if" onClick={() => openDialog('dataTable')} />
          </Group>
          <Group label="Outline">
            <Button icon="freeze" label="Freeze" pressed={frozen.rows > 0 || frozen.cols > 0} onClick={() => openDialog('freeze')} />
          </Group>
          <Group label="Export">
            <Button tall icon="export" label="CSV" onClick={() => exportAs('csv')} />
            <Button tall icon="export" label="TSV" onClick={() => exportAs('tsv')} />
            <Button tall icon="pdf" label="PDF" onClick={() => exportAs('pdf')} />
          </Group>
        </>
      ) : tab === 'review' ? (
        <>
          <Group label="Protect">
            <Button
              tall
              icon="lock"
              label={protectedSheet ? 'Unprotect' : 'Protect sheet'}
              pressed={protectedSheet}
              onClick={() => dispatch({ op: protectedSheet ? 'unprotect' : 'protect' })}
            />
            <Button
              icon="lock"
              label={format.locked === false ? 'Unlocked' : 'Locked'}
              title="Whether these cells are locked when the sheet is protected"
              pressed={format.locked !== false}
              onClick={() => dispatch({ op: 'setFormat', delta: { locked: format.locked === false } })}
            />
          </Group>
          <Group label="Rules">
            <Button tall icon="wand" label="Conditional" onClick={() => openDialog('conditional')} />
            <Button tall icon="check" label="Validation" onClick={() => openDialog('validation')} />
          </Group>
          <Group label="Find">
            <Button tall icon="find" label="Find and replace" onClick={() => openDialog('find')} />
          </Group>
        </>
      ) : (
        <>
          <Group label="Window">
            <Button tall icon="freeze" label="Freeze panes" pressed={frozen.rows > 0 || frozen.cols > 0} onClick={() => openDialog('freeze')} />
            <Button icon="maximize" label="Full screen" onClick={() => shell.win.fullscreen({})} />
          </Group>
          <Group label="Zoom">
            <Button icon="zoomOut" label="Out" onClick={() => shell.win.zoom({ delta: -0.1 })} />
            <Button icon="zoomIn" label="In" onClick={() => shell.win.zoom({ delta: 0.1 })} />
            <Button icon="check" label="100%" onClick={() => shell.win.zoom({ reset: true })} />
          </Group>
          <Group label="File">
            <Button tall icon="new" label="New" onClick={() => shell.win.create({ app: 'sheets' })} />
            <Button tall icon="open" label="Open" onClick={openFile} />
            <Button tall icon="save" label="Save" onClick={() => save(false)} />
          </Group>
        </>
      )}
    </Ribbon>
  );
}
