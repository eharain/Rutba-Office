'use strict';

/**
 * A page of paper you can pour content onto.
 *
 * `document.js` will draw text at a coordinate; it has no idea what a
 * paragraph is. This is the layer between: a cursor that walks down the
 * page, blocks that know their own height, and a page break that happens
 * when the next block will not fit rather than when somebody remembered to
 * ask.
 *
 * == The three things that make it a document rather than a drawing ======
 *
 * 1. **Nothing is measured twice.** A block computes its height with the
 *    same wrapper that later draws it, so "will this fit?" and "how much did
 *    it take?" cannot disagree. Height estimated one way and drawn another
 *    is how the last line of a table lands on top of the footer.
 *
 * 2. **A table header repeats.** A three-page list of invoice lines whose
 *    columns are only labelled on page one is a table on page one and a grid
 *    of numbers after that.
 *
 * 3. **Footers are written last.** "Page 3 of 7" cannot be drawn while
 *    laying out page 3, because nothing yet knows there will be seven. So
 *    every page keeps a late layer, and `end()` fills it in - which is also
 *    the only honest way to promise a reader that a document is complete and
 *    they have all of it.
 */

const { PdfDocument } = require('./document');
const { lineHeight } = require('./metrics');

const DEFAULT_MARGINS = { top: 54, right: 46, bottom: 58, left: 46 };

/** Ink. Grey rather than black for rules: a page of full-strength lines
 * reads as a form to fill in, not a document to keep. */
const INK = {
  text: '#1a1a1a',
  quiet: '#6b6b6b',
  rule: '#c8c8c8',
  faint: '#ececec',
  head: '#111111',
};

class Sheet {
  constructor({
    size = 'A4', landscape = false, margins = {}, title = '', author = '', subject = '',
    created = null, footer = '', font = 'Helvetica', fontSize = 9.5,
  } = {}) {
    this.doc = new PdfDocument({ size, landscape, title, author, subject, created });
    this.margins = { ...DEFAULT_MARGINS, ...margins };
    this.font = font;
    this.fontSize = fontSize;
    this.footerText = footer;
    this.page = null;
    this.y = 0;
    // Redrawn at the top of every page after the first - a table's column
    // headings, set by `table()` while it is running.
    this.continuation = null;
    this.finished = false;
    this.newPage();
  }

  get left() { return this.margins.left; }
  get right() { return this.doc.size[0] - this.margins.right; }
  get contentWidth() { return this.right - this.left; }
  get bottom() { return this.doc.size[1] - this.margins.bottom; }

  newPage() {
    this.page = this.doc.addPage();
    this.y = this.margins.top;
    if (this.continuation) this.continuation();
    return this.page;
  }

  /** Make room for `height`, taking a new page if there is not any. */
  space(height) {
    if (this.y + height <= this.bottom) return false;
    this.newPage();
    return true;
  }

  move(dy) { this.y += dy; return this.y; }

  /**
   * Break a string into lines that fit `width`.
   *
   * Greedy, which is what every word processor does by default, and exact,
   * because the widths are the viewer's own (see metrics.js). A single word
   * too long for the column - a URL, a part number, an email address - is
   * broken mid-word rather than allowed to run into the next column: an
   * ugly break is a smaller lie than a value that overlaps its neighbour.
   */
  wrap(value, width, { font = this.font, size = this.fontSize } = {}) {
    const out = [];
    const paragraphs = String(value === null || value === undefined ? '' : value).split(/\r?\n/);
    for (const para of paragraphs) {
      const words = para.split(/[ \t]+/).filter((w) => w.length);
      if (!words.length) { out.push(''); continue; }
      let line = '';
      for (const word of words) {
        const candidate = line ? `${line} ${word}` : word;
        if (this.doc.widthOf(candidate, { font, size }) <= width) { line = candidate; continue; }
        if (line) { out.push(line); line = ''; }
        if (this.doc.widthOf(word, { font, size }) <= width) { line = word; continue; }
        let chunk = '';
        for (const ch of word) {
          if (this.doc.widthOf(chunk + ch, { font, size }) > width && chunk) {
            out.push(chunk);
            chunk = ch;
          } else chunk += ch;
        }
        line = chunk;
      }
      if (line) out.push(line);
    }
    return out;
  }

  /** Draw one already-fitting line, honouring alignment within a box. */
  _drawLine(text, x, width, baseline, { align = 'left', font, size, colour }) {
    const opts = { font, size, colour };
    if (align === 'right') return this.page.textRight(text, x + width, baseline, opts);
    if (align === 'center') return this.page.textCentre(text, x + width / 2, baseline, opts);
    return this.page.text(text, x, baseline, opts);
  }

  /**
   * Flowing text. Breaks across pages line by line, because a paragraph that
   * refuses to split leaves half a page blank whenever it is long.
   */
  paragraph(value, {
    font = this.font, size = this.fontSize, colour = INK.text, align = 'left',
    x = null, width = null, gap = 4, leading = null,
  } = {}) {
    const boxX = x === null ? this.left : x;
    const boxWidth = width === null ? this.right - boxX : width;
    const step = leading || lineHeight(size);
    for (const line of this.wrap(value, boxWidth, { font, size })) {
      this.space(step);
      this._drawLine(line, boxX, boxWidth, this.y + size * 0.85, { align, font, size, colour });
      this.y += step;
    }
    this.y += gap;
    return this.y;
  }

  heading(value, { size = 15, font = 'Helvetica-Bold', colour = INK.head, gap = 8, ...rest } = {}) {
    return this.paragraph(value, { size, font, colour, gap, ...rest });
  }

  spacer(height = 8) { this.y += height; return this.y; }

  /**
   * An image at the cursor, from `doc.addImage`'s ref. The caller gives the
   * DRAWN size in points (aspect is its concern - it knows the pixels); the
   * block reserves its height like any paragraph and breaks the page the
   * same way.
   */
  image(ref, width, height, { gap = 6, x = null } = {}) {
    this.space(height + gap);
    this.page.image(ref, x === null ? this.left : x, this.y, width, height);
    this.y += height + gap;
    return this.y;
  }

  rule({ gap = 6, colour = INK.rule, width = 0.5, x = null, to = null } = {}) {
    this.space(gap * 2);
    this.y += gap;
    this.page.line(x === null ? this.left : x, this.y, to === null ? this.right : to, this.y,
      { colour, width });
    this.y += gap;
    return this.y;
  }

  /**
   * Blocks side by side - two addresses, or a party and a stack of dates.
   *
   * Each column is laid out independently and the cursor ends below the
   * TALLEST, so a five-line address next to a two-line one does not have the
   * next block overlapping it.
   */
  columns(blocks, { gap = 18, top = null } = {}) {
    const startY = top === null ? this.y : top;
    const flexTotal = blocks.reduce((n, b) => n + (b.width ? 0 : (b.flex || 1)), 0);
    const fixed = blocks.reduce((n, b) => n + (b.width || 0), 0);
    const free = this.contentWidth - fixed - gap * (blocks.length - 1);
    let x = this.left;
    let lowest = startY;
    for (const block of blocks) {
      const width = block.width || (flexTotal ? (free * (block.flex || 1)) / flexTotal : free);
      this.y = startY;
      const items = block.lines || [];
      for (const item of items) {
        const line = typeof item === 'string' ? { text: item } : item;
        if (line.gap) { this.y += line.gap; continue; }
        const font = line.font || block.font || this.font;
        const size = line.size || block.size || this.fontSize;
        const colour = line.colour || block.colour || INK.text;
        const align = line.align || block.align || 'left';
        for (const wrapped of this.wrap(line.text, width, { font, size })) {
          this._drawLine(wrapped, x, width, this.y + size * 0.85, { align, font, size, colour });
          this.y += lineHeight(size);
        }
      }
      lowest = Math.max(lowest, this.y);
      x += width + gap;
    }
    this.y = lowest;
    return this.y;
  }

  /**
   * Label/value pairs stacked in one column. What the top-right of an
   * invoice is made of.
   */
  keyValues(rows, {
    x = null, width = null, labelWidth = null, size = this.fontSize, gap = 2, align = 'right',
  } = {}) {
    const boxX = x === null ? this.left : x;
    const boxWidth = width === null ? this.right - boxX : width;
    const labels = labelWidth === null ? boxWidth * 0.45 : labelWidth;
    for (const [label, value] of rows) {
      if (value === null || value === undefined || value === '') continue;
      const lines = this.wrap(String(value), boxWidth - labels - 6, { size });
      // The label wraps inside ITS column too. A long one - a party's name
      // with their address, a document's full file name - used to run on
      // into the value and print over it; now it takes the lines it needs
      // and the row grows to hold both sides.
      const labelLines = this.wrap(String(label), labels - 6, { size });
      const rowLines = Math.max(1, lines.length, labelLines.length);
      const height = rowLines * lineHeight(size);
      this.space(height);
      const baseline = this.y + size * 0.85;
      labelLines.forEach((line, i) => {
        this.page.text(line, boxX, baseline + i * lineHeight(size), { size, colour: INK.quiet });
      });
      lines.forEach((line, i) => {
        this._drawLine(line, boxX + labels, boxWidth - labels, baseline + i * lineHeight(size),
          { align, size, font: this.font, colour: INK.text });
      });
      this.y += height + gap;
    }
    return this.y;
  }

  /**
   * A table, with the two properties that separate one from a grid: columns
   * sized once for the whole table, and headings that come back after a page
   * break.
   *
   * `columns` is `[{ key, label, width|flex, align, font, size }]`; `rows` is
   * an array of objects keyed by `key`, or of arrays in column order.
   */
  table({
    columns, rows, size = this.fontSize, headSize = null, padding = 4,
    headFont = 'Helvetica-Bold', zebra = false, gap = 8, rowRule = true,
  }) {
    if (!columns || !columns.length) throw new Error('A table needs columns.');
    const headerSize = headSize || size;
    const flexTotal = columns.reduce((n, c) => n + (c.width ? 0 : (c.flex || 1)), 0);
    const fixed = columns.reduce((n, c) => n + (c.width || 0), 0);
    const free = Math.max(0, this.contentWidth - fixed);
    const widths = columns.map((c) => (c.width || (flexTotal ? (free * (c.flex || 1)) / flexTotal : 0)));

    const headHeight = lineHeight(headerSize) + padding * 2;
    const drawHead = () => {
      let x = this.left;
      const baseline = this.y + padding + headerSize * 0.85;
      columns.forEach((column, i) => {
        this._drawLine(String(column.label ?? ''), x, widths[i], baseline,
          { align: column.align || 'left', font: headFont, size: headerSize, colour: INK.head });
        x += widths[i];
      });
      this.y += headHeight;
      this.page.line(this.left, this.y, this.right, this.y, { colour: INK.rule, width: 0.9 });
    };

    this.space(headHeight + lineHeight(size) + padding * 2);
    drawHead();
    // From here until the table ends, a new page starts with the headings.
    this.continuation = drawHead;

    const cellsOf = (row) => columns.map((column, i) =>
      (Array.isArray(row) ? row[i] : row[column.key]));

    let index = 0;
    for (const row of rows || []) {
      const values = cellsOf(row).map((v) => (v === null || v === undefined ? '' : String(v)));
      const wrapped = values.map((value, i) => this.wrap(value, widths[i] - 6,
        { font: columns[i].font || this.font, size: columns[i].size || size }));
      const lines = Math.max(1, ...wrapped.map((w) => w.length));
      const step = lineHeight(size);

      /**
       * Rows are drawn a LINE at a time, across all columns, rather than a
       * column at a time. It costs nothing for the ordinary row - one pass
       * of the loop - and it means a cell too tall for the remaining page
       * continues onto the next one instead of being drawn off the paper.
       * Column-at-a-time cannot do that: by the time the first column has
       * run out of page, the others have not started.
       */
      let done = 0;
      while (done < lines) {
        const room = Math.floor((this.bottom - this.y - padding * 2) / step);
        if (room < 1) {
          const wasAt = this.y;
          this.newPage();
          if (this.y >= wasAt) {
            throw new Error('A table row does not fit even on an empty page; the margins leave no body.');
          }
          continue;
        }
        const take = Math.min(room, lines - done);
        const height = take * step + padding * 2;
        if (zebra && index % 2 === 1) {
          this.page.rect(this.left, this.y, this.contentWidth, height, { fill: INK.faint });
        }
        let x = this.left;
        wrapped.forEach((cellLines, i) => {
          const column = columns[i];
          const cellSize = column.size || size;
          for (let l = done; l < done + take; l++) {
            if (cellLines[l] === undefined) break;
            this._drawLine(cellLines[l], x + 3, widths[i] - 6,
              this.y + padding + cellSize * 0.85 + (l - done) * step,
              { align: column.align || 'left', font: column.font || this.font, size: cellSize, colour: column.colour || INK.text });
          }
          x += widths[i];
        });
        this.y += height;
        done += take;
        if (done < lines) this.newPage();
      }
      if (rowRule) {
        this.page.line(this.left, this.y, this.right, this.y, { colour: INK.faint, width: 0.5 });
      }
      index += 1;
    }

    this.continuation = null;
    this.y += gap;
    return this.y;
  }

  /**
   * The totals block: right-aligned pairs under the table, the last one
   * emphasised because it is the number the reader came for.
   */
  totals(rows, { width = 220, size = this.fontSize, padding = 3 } = {}) {
    const x = this.right - width;
    const list = rows.filter(([, value]) => value !== null && value !== undefined && value !== '');
    for (let i = 0; i < list.length; i++) {
      const [label, value, options] = list[i];
      const strong = options?.strong ?? (i === list.length - 1);
      const rowSize = strong ? size + 1 : size;
      const height = lineHeight(rowSize) + padding;
      this.space(height + 4);
      if (strong) {
        this.page.line(x, this.y, this.right, this.y, { colour: INK.rule, width: 0.7 });
        this.y += 3;
      }
      const baseline = this.y + rowSize * 0.85;
      this.page.text(String(label), x, baseline,
        { size: rowSize, font: strong ? 'Helvetica-Bold' : this.font, colour: strong ? INK.head : INK.quiet });
      this.page.textRight(String(value), this.right, baseline,
        { size: rowSize, font: strong ? 'Helvetica-Bold' : this.font, colour: INK.text });
      this.y += height;
    }
    return this.y;
  }

  /**
   * Finish, and only now write the footers.
   *
   * "Page 2 of 5" is the one thing on a document that cannot be written
   * while it is being laid out, and it is also the thing that tells a reader
   * whether the post lost a sheet. So it is worth the late layer.
   */
  end() {
    // Idempotent: `toBuffer()` calls this, and a caller who already had
    // would otherwise get every footer drawn twice, a shade bolder each time.
    if (this.finished) return this.doc;
    this.finished = true;
    const total = this.doc.pages.length;
    this.doc.pages.forEach((page, i) => {
      const y = page.height - this.margins.bottom + 22;
      page.line(this.margins.left, y - 12, page.width - this.margins.right, y - 12,
        { colour: INK.faint, width: 0.5, late: true });
      if (this.footerText) {
        page.text(this.footerText, this.margins.left, y, { size: 7.5, colour: INK.quiet, late: true });
      }
      page.textRight(`Page ${i + 1} of ${total}`, page.width - this.margins.right, y,
        { size: 7.5, colour: INK.quiet, late: true });
    });
    return this.doc;
  }

  toBuffer() { return this.end().toBuffer(); }
  toBase64() { return this.end().toBase64(); }
}

module.exports = { Sheet, INK, DEFAULT_MARGINS };
