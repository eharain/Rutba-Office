/**
 * The OOXML package: parts, content types and relationships.
 *
 * This is the layer that makes "from scratch" survivable. An .xlsx or .docx is
 * an Open Packaging Convention container — a zip of parts, a content-type map,
 * and relationship graphs. Almost every feature we are not going to implement
 * for years (charts, pivot tables, slicers, VML, embedded OLE, macros, custom
 * XML, digital signature parts) lives in a part we can carry untouched.
 *
 * The rule this file enforces:
 *
 *     WE ONLY REWRITE PARTS WE DELIBERATELY EDIT.
 *     EVERY OTHER PART IS RETURNED BYTE-IDENTICAL.
 *
 * That inverts the usual difficulty curve. A conventional implementation must
 * understand a feature to avoid destroying it, so fidelity grows only as fast as
 * the feature list. Here, fidelity is high on day one and degrades only in the
 * specific places we choose to touch — which are also the places we test.
 *
 * What this does NOT give us: correct RENDERING or RECALCULATION of features we
 * do not model. A chart we preserve is a chart we cannot draw. That is the
 * honest boundary, and it is why the editor roadmap and the fidelity roadmap are
 * different roadmaps.
 */
import { readZip, writeZip, ZipEntry, inflateEntries } from './zip.js';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

export class OoxmlError extends Error {
  constructor(message) {
    super(message);
    this.name = 'OoxmlError';
  }
}

const esc = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * Normalise browser-shaped input (Uint8Array, ArrayBuffer) to a Buffer view,
 * zero-copy where the host has Buffer. Anything else passes through untouched
 * so `read`'s own guard produces its usual error, and a host with no Buffer at
 * all passes through too — the remaining Buffer dependence is inventoried, not
 * hidden.
 */
function toBuffer(bytes) {
  if (typeof Buffer === 'undefined' || Buffer.isBuffer(bytes)) return bytes;
  if (bytes instanceof ArrayBuffer) return Buffer.from(bytes);
  if (ArrayBuffer.isView(bytes)) return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return bytes;
}

/** Attribute reader that does not care about attribute order or quoting style. */
export function attrs(tag) {
  const out = {};
  const re = /([\w:.-]+)\s*=\s*"([^"]*)"|([\w:.-]+)\s*=\s*'([^']*)'/g;
  let m;
  while ((m = re.exec(tag))) out[m[1] ?? m[3]] = m[2] ?? m[4];
  return out;
}

export class OoxmlPackage {
  constructor({ entries, comment }) {
    /** @type {ZipEntry[]} in original order */
    this.entries = entries;
    this.comment = comment;
    this.byName = new Map(entries.map((e) => [e.name, e]));
    this._contentTypes = null;
  }

  static read(buf) {
    if (!Buffer.isBuffer(buf) || buf.length < 22 || buf.readUInt16LE(0) !== 0x4b50) {
      // The first real corpus contained a zero-byte .docx; 'missing zip
      // signature' sent someone hunting for a parser bug that was not there.
      throw new OoxmlError(buf.length === 0
        ? 'empty file — zero bytes, nothing to open'
        : 'not an OOXML package: missing zip signature');
    }
    const pkg = new OoxmlPackage(readZip(buf));
    if (!pkg.byName.has('[Content_Types].xml')) {
      throw new OoxmlError('not an OOXML package: [Content_Types].xml missing');
    }
    return pkg;
  }

  /**
   * `read`, awaitable — the environment-agnostic door.
   *
   * `zlib.inflateRawSync` is the one Node-only call on the whole read path, and
   * its web-standard twin, `DecompressionStream('deflate-raw')`, is async by
   * nature. So the async lives HERE, once, at the door: the parse is the same
   * code `read` runs (readAsync calls it), and then every deflated part is
   * inflated up front, so everything downstream — `part`, `text`, Workbook,
   * Document — stays synchronous and identical. Pre-inflating populates the
   * lazy cache only; `modified` stays false, untouched parts still round-trip
   * from their original compressed bytes, and `readAsync(b)` then `.write()`
   * equals `read(b).write()` byte for byte.
   *
   * Accepts what a browser hands you (Uint8Array from a fetch or a file input,
   * ArrayBuffer) as well as a Buffer.
   *
   * @param {Buffer|Uint8Array|ArrayBuffer} bytes
   * @param {{inflate?: 'stream'}} [options] internal — `'stream'` forces the
   *   DecompressionStream inflater even where zlib exists; tests use it to
   *   exercise the browser path on Node.
   * @returns {Promise<OoxmlPackage>}
   */
  static async readAsync(bytes, options = {}) {
    const pkg = OoxmlPackage.read(toBuffer(bytes));
    await inflateEntries(pkg.entries, options);
    return pkg;
  }

  /**
   * A package assembled from nothing — parts in, package out.
   *
   * `read` is the door for a file that already exists, and everything about
   * this package is built around preserving such a file: entries keep their
   * original compressed bytes until something replaces them, and `modified`
   * is what the fidelity report is computed from. None of that has an answer
   * for a file that has never existed, which is why creating one needed its
   * own door rather than a flag on that one.
   *
   * It exists because the code that CREATES .xlsx and .docx files — templates,
   * the workbook a .csv decodes into, the seeded demo documents — had been
   * carrying its own zip writer, its own CRC table and its own XML escaper in
   * a package named for data bindings. Two zip layers in one repo is one
   * layer too many: this one handles ZIP64, which a real workbook can need,
   * and the other one did not.
   *
   * Content types are the CALLER's to declare. Every builder already emits a
   * complete `[Content_Types].xml`, and inferring one here would quietly
   * disagree with what they wrote — so the part is required rather than
   * generated, and its absence is refused with the same message `read` uses.
   *
   * @param {Array<{name: string, data: Buffer|string}>} parts
   */
  static fromParts(parts) {
    const entries = (parts ?? []).map(({ name, data }) => {
      const entry = new ZipEntry({
        name,
        method: 8,
        crc: 0,
        compressedSize: 0,
        uncompressedSize: 0,
        compressed: Buffer.alloc(0),
        flags: 0,
        // A fixed timestamp, so building the same document twice produces the
        // same bytes. A build that embedded the clock would make every
        // content-addressed store treat identical files as different ones.
        dosTime: 0,
        dosDate: 0x2821,
        externalAttrs: 0,
        comment: Buffer.alloc(0),
      });
      entry.data = Buffer.isBuffer(data) ? data : Buffer.from(String(data), 'utf8');
      return entry;
    });

    const pkg = new OoxmlPackage({ entries, comment: Buffer.alloc(0) });
    if (!pkg.byName.has('[Content_Types].xml')) {
      throw new OoxmlError('not an OOXML package: [Content_Types].xml missing');
    }
    return pkg;
  }

  write() {
    return writeZip(this.entries, this.comment);
  }

  has(name) { return this.byName.has(name); }
  part(name) { return this.byName.get(name) ?? null; }
  partNames() { return this.entries.map((e) => e.name); }

  /** Decompressed bytes of a part. Reading never marks it modified. */
  read(name) {
    const entry = this.byName.get(name);
    if (!entry) throw new OoxmlError('no such part: ' + name);
    return entry.data;
  }

  text(name) { return this.read(name).toString('utf8'); }

  /** Replace a part. This is the ONLY way a part stops being byte-identical. */
  write_(name, data) {
    const entry = this.byName.get(name);
    if (!entry) throw new OoxmlError('no such part: ' + name);
    entry.data = Buffer.isBuffer(data) ? data : Buffer.from(String(data), 'utf8');
    return entry;
  }

  /** Add a new part, registering its content type and appending in order. */
  addPart(name, data, contentType) {
    if (this.byName.has(name)) throw new OoxmlError('part already exists: ' + name);
    const entry = new ZipEntry({
      name, method: 8, crc: 0, compressedSize: 0, uncompressedSize: 0,
      compressed: Buffer.alloc(0), flags: 0, dosTime: 0, dosDate: 0x2821,
      externalAttrs: 0, comment: Buffer.alloc(0),
    });
    entry.data = Buffer.isBuffer(data) ? data : Buffer.from(String(data), 'utf8');
    this.entries.push(entry);
    this.byName.set(name, entry);
    if (contentType) this.setOverride(name, contentType);
    return entry;
  }

  /**
   * Remove a part, its content-type override and its `.rels`.
   *
   * The counterpart to `addPart`, and it exists for ONE reason: undoing an
   * edit that created parts. Without it, undoing "add a pivot table" leaves
   * the pivot's parts in the package with a relationship still pointing at
   * them — a ghost feature that reappears the moment anything re-reads the
   * file. Removing a part some other part still references is the caller's
   * problem to avoid; this is deliberately a low-level operation.
   */
  removePart(name) {
    if (!this.byName.has(name)) return false;
    const relsPath = OoxmlPackage.relsPathFor(name);
    this.entries = this.entries.filter((e) => e.name !== name && e.name !== relsPath);
    this.byName.delete(name);
    this.byName.delete(relsPath);
    const partName = name.startsWith('/') ? name : '/' + name;
    const xml = this.text('[Content_Types].xml');
    const without = xml.replace(
      new RegExp('<Override[^>]*PartName="' + partName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'),
      '',
    );
    if (without !== xml) this.write_('[Content_Types].xml', without);
    this._contentTypes = null;
    return true;
  }

  /** Which parts have been rewritten. The fidelity report is built from this. */
  modifiedParts() {
    return this.entries.filter((e) => e.modified).map((e) => e.name);
  }
  untouchedParts() {
    return this.entries.filter((e) => !e.modified).map((e) => e.name);
  }

  // ---- content types -------------------------------------------------------

  contentTypes() {
    if (this._contentTypes) return this._contentTypes;
    const xml = this.text('[Content_Types].xml');
    const defaults = new Map();
    const overrides = new Map();
    for (const m of xml.matchAll(/<Default\b([^>]*)\/?>/g)) {
      const a = attrs(m[1]);
      if (a.Extension) defaults.set(a.Extension.toLowerCase(), a.ContentType);
    }
    for (const m of xml.matchAll(/<Override\b([^>]*)\/?>/g)) {
      const a = attrs(m[1]);
      if (a.PartName) overrides.set(a.PartName, a.ContentType);
    }
    this._contentTypes = { defaults, overrides };
    return this._contentTypes;
  }

  contentTypeOf(name) {
    const { defaults, overrides } = this.contentTypes();
    const partName = name.startsWith('/') ? name : '/' + name;
    if (overrides.has(partName)) return overrides.get(partName);
    const ext = name.includes('.') ? name.slice(name.lastIndexOf('.') + 1).toLowerCase() : '';
    return defaults.get(ext) ?? null;
  }

  setOverride(name, contentType) {
    const partName = name.startsWith('/') ? name : '/' + name;
    const { overrides } = this.contentTypes();
    if (overrides.get(partName) === contentType) return;
    overrides.set(partName, contentType);
    let xml = this.text('[Content_Types].xml');
    if (new RegExp('<Override[^>]*PartName="' + partName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/?>').test(xml)) {
      xml = xml.replace(
        new RegExp('<Override[^>]*PartName="' + partName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'),
        '<Override PartName="' + esc(partName) + '" ContentType="' + esc(contentType) + '"/>',
      );
    } else {
      xml = xml.replace(/<\/Types>\s*$/, '<Override PartName="' + esc(partName) + '" ContentType="' + esc(contentType) + '"/></Types>');
    }
    this.write_('[Content_Types].xml', xml);
  }

  /**
   * Make sure an extension has a `<Default>` content type.
   *
   * A media part is named by its extension, not by an override: PowerPoint
   * and Word write `<Default Extension="png" .../>` once and never list the
   * pictures. An override per picture is legal and Office reads it, but a
   * package whose content types are not the shape Office writes is a package
   * a stricter reader may refuse — and a diff against the original should be
   * one line, not one per picture.
   */
  ensureDefault(extension, contentType) {
    const ext = String(extension).replace(/^\./, '').toLowerCase();
    const { defaults } = this.contentTypes();
    if (defaults.has(ext)) return defaults.get(ext);
    defaults.set(ext, contentType);
    const xml = this.text('[Content_Types].xml');
    const entry = '<Default Extension="' + esc(ext) + '" ContentType="' + esc(contentType) + '"/>';
    const at = xml.indexOf('>', xml.indexOf('<Types'));
    if (at < 0) throw new OoxmlError('[Content_Types].xml has no <Types> element');
    this.write_('[Content_Types].xml', xml.slice(0, at + 1) + entry + xml.slice(at + 1));
    return contentType;
  }

  // ---- relationships -------------------------------------------------------

  /** `xl/workbook.xml` -> `xl/_rels/workbook.xml.rels` */
  static relsPathFor(partName) {
    const i = partName.lastIndexOf('/');
    const dir = i < 0 ? '' : partName.slice(0, i + 1);
    const base = i < 0 ? partName : partName.slice(i + 1);
    return dir + '_rels/' + base + '.rels';
  }

  /**
   * Add one relationship to a part's `.rels`, creating the rels part when it
   * does not exist yet. Returns the minted rId. One implementation for every
   * part-creating edit — pivots, drawings, charts — because the id-minting
   * and the create-or-append split are exactly the kind of code that drifts
   * when copied.
   */
  addRelationshipTo(fromPart, type, target) {
    const relsPath = OoxmlPackage.relsPathFor(fromPart);
    const existing = this.has(relsPath) ? this.text(relsPath) : null;
    const used = new Set((existing ?? '').match(/Id="rId(\d+)"/g)?.map((s) => Number(/\d+/.exec(s)[0])) ?? []);
    let n = 1;
    while (used.has(n)) n += 1;
    const id = 'rId' + n;
    const entry = '<Relationship Id="' + id + '" Type="' + type + '" Target="' + esc(target) + '"/>';
    if (existing) {
      this.write_(relsPath, existing.replace('</Relationships>', entry + '</Relationships>'));
    } else {
      this.addPart(relsPath,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + entry + '</Relationships>',
        'application/vnd.openxmlformats-package.relationships+xml');
    }
    return id;
  }

  /** The next free number for a numbered part family, e.g. `drawing3.xml`. */
  nextPartNumber(dir, prefix) {
    let n = 1;
    while (this.has(dir + prefix + n + '.xml')) n += 1;
    return n;
  }

  /** @returns {Array<{Id: string, Type: string, Target: string, TargetMode?: string}>} */
  rels(partName) {
    const relsPath = OoxmlPackage.relsPathFor(partName);
    if (!this.has(relsPath)) return [];
    const xml = this.text(relsPath);
    const out = [];
    for (const m of xml.matchAll(/<Relationship\b([^>]*)\/?>/g)) out.push(attrs(m[1]));
    return out;
  }

  /** Resolve a relationship Target against the part that declares it. */
  static resolveTarget(fromPart, target) {
    if (/^[a-z]+:\/\//i.test(target)) return target; // external
    if (target.startsWith('/')) return target.slice(1);
    const i = fromPart.lastIndexOf('/');
    const dir = i < 0 ? '' : fromPart.slice(0, i + 1);
    const joined = dir + target;
    const segments = [];
    for (const seg of joined.split('/')) {
      if (seg === '.' || seg === '') continue;
      if (seg === '..') segments.pop();
      else segments.push(seg);
    }
    return segments.join('/');
  }

  /** Root document part, via the package-level relationships. */
  mainDocument() {
    for (const rel of this.rels('')) {
      if (String(rel.Type).endsWith('/officeDocument')) {
        return OoxmlPackage.resolveTarget('', rel.Target);
      }
    }
    for (const candidate of ['xl/workbook.xml', 'word/document.xml', 'ppt/presentation.xml']) {
      if (this.has(candidate)) return candidate;
    }
    throw new OoxmlError('no officeDocument relationship and no known main part');
  }

  kind() {
    const main = this.mainDocument();
    if (main.startsWith('xl/')) return 'sheet';
    if (main.startsWith('word/')) return 'document';
    if (main.startsWith('ppt/')) return 'presentation';
    return 'unknown';
  }
}

/** Package-level `_rels/.rels` lives at a fixed path; make relsPathFor agree. */
const originalRelsPathFor = OoxmlPackage.relsPathFor;
OoxmlPackage.relsPathFor = (partName) => (partName === '' ? '_rels/.rels' : originalRelsPathFor(partName));

export { esc };
