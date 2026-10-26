'use strict';

/**
 * An incremental update: objects added, redefined or freed after the last
 * byte of a PDF, with cross-references for them alone and a trailer chained
 * to the file's own by /Prev. The original bytes stay the exact start of the
 * file, as every reader that writes comments and form values keeps them.
 *
 * textlayer.js wrote the first of these by hand; annotations, form values
 * and signatures as stamps all write through this. The cross-references are
 * written in the form the file's last section has (a table after a table, a
 * stream after a stream), so a reader that met the original meets the update
 * in the same form.
 *
 * Values are the reader's own (Dict, Name, Ref, PdfString, arrays, numbers);
 * a Stream is written with its bytes as they are and a /Length for them.
 *
 * A file protected with a password, read with it, is updated under the same
 * protection: each object written has its strings and stream encrypted with
 * the file's key, as the rest of the file is, and the trailer keeps its
 * Encrypt entry and its ID. A protected file read without its password
 * cannot be written into.
 */

const zlib = require('zlib');
const { openPdf, PdfFile, Name, Ref, Dict, Stream } = require('./reader');
const { valueText } = require('./assemble');

class UpdateError extends Error {
  /** code: 'encrypted' | 'damaged' | 'page' | 'missing' */
  constructor(message, code) {
    super(message);
    this.name = 'PdfUpdateError';
    this.code = code;
  }
}

class PdfUpdate {
  /**
   * @param {Buffer|Uint8Array|PdfFile} source the whole file, or a PdfFile read from it
   * @param {{ what?: string }} [options] what the update writes, for the sentences it refuses with
   */
  constructor(source, { what = 'This change' } = {}) {
    const file = source instanceof PdfFile ? source : openPdf(source);
    this.file = file;
    this.buf = file.buf;
    if (file.locked) throw new UpdateError(`${what} cannot be written into a PDF protected with a password until it is opened with its password.`, 'encrypted');
    if (file.repaired) throw new UpdateError(`This PDF's cross-references are damaged, so ${what.charAt(0).toLowerCase()}${what.slice(1)} cannot be added to it as it is. Save a copy first, then work on the copy.`, 'damaged');
    this.size = Number(file.trailer.get('Size')) || 0;
    this.next = Math.max(this.size, ...[...file.entries.keys()].map((n) => n + 1), 1);
    /** num -> { gen, value } | { gen, free: true } */
    this.objects = new Map();
  }

  /** A new object number. */
  alloc() { return this.next++; }

  /** A new object holding `value`; its reference. */
  add(value) {
    const num = this.alloc();
    this.objects.set(num, { gen: 0, value });
    return new Ref(num, 0);
  }

  /** An object redefined: the same number, a new value. */
  set(ref, value) {
    const num = ref instanceof Ref ? ref.num : Number(ref);
    const gen = ref instanceof Ref ? ref.gen : (this.file.entries.get(num)?.gen || 0);
    this.objects.set(num, { gen, value });
    return new Ref(num, gen);
  }

  /** An object no longer in the file: free from this update on, its next generation one more. */
  free(ref) {
    const num = ref instanceof Ref ? ref.num : Number(ref);
    if (!(num > 0)) return;
    const gen = (ref instanceof Ref ? ref.gen : (this.file.entries.get(num)?.gen || 0)) + 1;
    this.objects.set(num, { gen: Math.min(gen, 65535), free: true });
  }

  /** The value an object has in the file as this update leaves it. */
  get(ref) {
    const num = ref instanceof Ref ? ref.num : Number(ref);
    const mine = this.objects.get(num);
    if (mine) return mine.free ? null : mine.value;
    return this.file.getObject(num);
  }

  resolve(value) {
    let v = value;
    for (let i = 0; i < 32 && v instanceof Ref; i++) v = this.get(v);
    return v instanceof Ref ? null : v;
  }

  get empty() { return this.objects.size === 0; }

  /** The file with the update after it. */
  toBuffer() {
    const { file, buf } = this;
    if (this.empty) return buf;
    const base = file.base;
    const startXref = file.startXref();
    const classic = buf.toString('latin1', startXref, startXref + 4) === 'xref';
    const parts = [buf];
    let length = buf.length;
    const push = (data) => {
      const b = Buffer.isBuffer(data) ? data : Buffer.from(data, 'latin1');
      parts.push(b);
      length += b.length;
    };
    if (buf[buf.length - 1] !== 0x0a && buf[buf.length - 1] !== 0x0d) push('\n');
    const offsets = new Map(); // num -> { offset, gen, free }
    for (const [num, o] of [...this.objects].sort((a, b) => a[0] - b[0])) {
      if (o.free) { offsets.set(num, { offset: 0, gen: o.gen, free: true }); continue; }
      offsets.set(num, { offset: length - base, gen: o.gen });
      push(`${num} ${o.gen} obj\n`);
      const v = file.security ? file.security.encryptObject(o.value, num, o.gen, (x) => this.resolve(x)) : o.value;
      if (v instanceof Stream) {
        const dict = new Dict(v.dict.entries());
        dict.set('Length', v.data.length);
        push(`${valueText(dict)}\nstream\n`);
        push(v.data);
        push('\nendstream');
      } else {
        push(valueText(v));
      }
      push('\nendobj\n');
    }
    const trailer = new Dict();
    for (const key of ['Root', 'Info', 'ID', 'Encrypt']) if (file.trailer.has(key)) trailer.set(key, file.trailer.get(key));
    trailer.set('Prev', startXref - base);

    const runs = (nums) => {
      const out = [];
      for (const n of nums.sort((a, b) => a - b)) {
        const last = out[out.length - 1];
        if (last && last.first + last.count === n) last.count++;
        else out.push({ first: n, count: 1 });
      }
      return out;
    };

    if (classic) {
      trailer.set('Size', Math.max(this.size, this.next));
      const xrefAt = length - base;
      let xref = 'xref\n';
      for (const run of runs([...offsets.keys()])) {
        xref += `${run.first} ${run.count}\n`;
        for (let n = run.first; n < run.first + run.count; n++) {
          const e = offsets.get(n);
          xref += `${String(e.offset).padStart(10, '0')} ${String(e.gen).padStart(5, '0')} ${e.free ? 'f' : 'n'} \n`;
        }
      }
      push(xref);
      push(`trailer\n${valueText(trailer)}\nstartxref\n${xrefAt}\n%%EOF\n`);
    } else {
      const xrefNum = this.next++;
      const xrefAt = length - base;
      offsets.set(xrefNum, { offset: xrefAt, gen: 0 });
      const list = runs([...offsets.keys()]);
      const rows = [];
      for (const run of list) {
        for (let n = run.first; n < run.first + run.count; n++) {
          const e = offsets.get(n);
          const row = Buffer.alloc(7);
          row[0] = e.free ? 0 : 1;
          row.writeUInt32BE(e.free ? 0 : e.offset, 1);
          row.writeUInt16BE(e.gen, 5);
          rows.push(row);
        }
      }
      const data = zlib.deflateSync(Buffer.concat(rows));
      trailer.set('Type', new Name('XRef'));
      trailer.set('Size', Math.max(this.size, this.next));
      trailer.set('W', [1, 4, 2]);
      trailer.set('Index', list.flatMap((r) => [r.first, r.count]));
      trailer.set('Filter', new Name('FlateDecode'));
      trailer.set('Length', data.length);
      push(`${xrefNum} 0 obj\n${valueText(trailer)}\nstream\n`);
      push(data);
      push(`\nendstream\nendobj\nstartxref\n${xrefAt}\n%%EOF\n`);
    }
    return Buffer.concat(parts, length);
  }
}

/** A Dict's entries copied into a new Dict, so a redefinition never changes the reader's cached value. */
function copyDict(dict) {
  return new Dict(dict instanceof Dict ? dict.entries() : []);
}

/** A content stream (or an appearance) deflated, with the dictionary entries given. */
function flateStream(entries, text) {
  const data = zlib.deflateSync(Buffer.isBuffer(text) ? text : Buffer.from(text, 'latin1'));
  const dict = new Dict(entries);
  dict.set('Filter', new Name('FlateDecode'));
  return new Stream(dict, data);
}

module.exports = { PdfUpdate, UpdateError, copyDict, flateStream };
