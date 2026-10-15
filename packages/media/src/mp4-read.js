// An MP4 or QuickTime file read down to its samples.
//
// What the export's decoders and the tests need: the top-level boxes in
// order, each track's codec as WebCodecs names it with its configuration
// (avcC, hvcC, vpcC, av1C, the AAC config or the Opus header), its size,
// rate, edit list and every sample (where it is, how big, its decode and
// presentation times, whether it is a sync sample). Plain files read their
// sample tables; a fragmented file (a recorder's) is read fragment by
// fragment.

const dec = (b, at, len) => new TextDecoder('latin1').decode(b.subarray(at, at + len));
const u16 = (b, at) => (b[at] << 8) | b[at + 1];
const u32 = (b, at) => ((b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3]) >>> 0;
const i32 = (b, at) => (b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3];
const u64 = (b, at) => u32(b, at) * 2 ** 32 + u32(b, at + 4);
const hex2 = (n) => n.toString(16).padStart(2, '0');

/** The boxes between `start` and `end`: [{ type, at, size, body, end }]. */
export function boxesIn(b, start = 0, end = b.length) {
  const out = [];
  let at = start;
  while (at + 8 <= end) {
    let size = u32(b, at);
    const type = dec(b, at + 4, 4);
    let head = 8;
    if (size === 1) {
      if (at + 16 > end) break;
      size = u64(b, at + 8);
      head = 16;
    } else if (size === 0) {
      size = end - at;
    }
    if (size < head) break;
    out.push({ type, at, size, body: at + head, end: Math.min(at + size, end) });
    at += size;
  }
  return out;
}

const child = (b, parent, type) => boxesIn(b, parent.body, parent.end).find((x) => x.type === type) || null;
const children = (b, parent, type) => boxesIn(b, parent.body, parent.end).filter((x) => x.type === type);
const path = (b, parent, ...types) => {
  let at = parent;
  for (const t of types) {
    at = at && child(b, at, t);
    if (!at) return null;
  }
  return at;
};

/** The codec string WebCodecs takes, from a sample entry and its configuration. */
function codecString(entry, config) {
  if (entry === 'avc1' || entry === 'avc3') {
    if (!config || config.length < 4) return 'avc1.42001f';
    return `avc1.${hex2(config[1])}${hex2(config[2])}${hex2(config[3])}`;
  }
  if (entry === 'hvc1' || entry === 'hev1') {
    if (!config || config.length < 13) return `${entry}.1.6.L93.B0`;
    const space = ['', 'A', 'B', 'C'][config[1] >> 6];
    const tier = (config[1] >> 5) & 1 ? 'H' : 'L';
    const profile = config[1] & 0x1f;
    let compat = u32(config, 2);
    let rev = 0;
    for (let i = 0; i < 32; i++) { rev = (rev << 1) | (compat & 1); compat >>>= 1; }
    const level = config[12];
    const cons = [...config.subarray(6, 12)];
    while (cons.length && cons[cons.length - 1] === 0) cons.pop();
    return [entry, `${space}${profile}`, (rev >>> 0).toString(16), `${tier}${level}`, ...cons.map((c) => c.toString(16).toUpperCase())].join('.');
  }
  if (entry === 'vp09') {
    if (!config || config.length < 8) return 'vp09.00.10.08';
    // vpcC is a full box: version and flags first.
    return `vp09.${String(config[4]).padStart(2, '0')}.${String(config[5]).padStart(2, '0')}.${String(config[6] >> 4).padStart(2, '0')}`;
  }
  if (entry === 'av01') {
    if (!config || config.length < 4) return 'av01.0.04M.08';
    const profile = config[1] >> 5;
    const level = config[1] & 0x1f;
    const tier = config[2] >> 7 ? 'H' : 'M';
    const high = (config[2] >> 6) & 1;
    const twelve = (config[2] >> 5) & 1;
    const depth = high ? (twelve ? 12 : 10) : 8;
    return `av01.${profile}.${String(level).padStart(2, '0')}${tier}.${String(depth).padStart(2, '0')}`;
  }
  if (entry === 'mp4a') return 'mp4a.40.2';
  if (entry === 'Opus') return 'opus';
  if (entry === 'vp08') return 'vp8';
  return null;
}

/** The AAC AudioSpecificConfig inside an esds box. */
function ascOf(b, esds) {
  let at = esds.body + 4;
  const end = esds.end;
  const readLen = () => {
    let n = 0;
    for (let i = 0; i < 4 && at < end; i++) {
      const c = b[at++];
      n = (n << 7) | (c & 0x7f);
      if (!(c & 0x80)) break;
    }
    return n;
  };
  while (at < end) {
    const tag = b[at++];
    const len = readLen();
    if (tag === 0x03) { at += 3; continue; } // ES descriptor: ES_ID and flags, then its children
    if (tag === 0x04) { at += 13; continue; } // decoder config: then the specific info
    if (tag === 0x05) return b.slice(at, at + len);
    at += len;
  }
  return null;
}

/** dOps (big-endian) as the OpusHead a decoder takes (little-endian). */
function opusHead(b, dops) {
  const at = dops.body;
  const head = new Uint8Array(19);
  head.set(new TextEncoder().encode('OpusHead'));
  head[8] = 1;
  head[9] = b[at + 1];
  const dv = new DataView(head.buffer);
  dv.setUint16(10, u16(b, at + 2), true);
  dv.setUint32(12, u32(b, at + 4), true);
  dv.setInt16(16, (b[at + 8] << 8 | b[at + 9]) << 16 >> 16, true);
  head[18] = b[at + 10];
  return head;
}

function sampleEntry(b, stsd) {
  const at = stsd.body + 8; // version, flags, entry count
  const size = u32(b, at);
  const type = dec(b, at + 4, 4);
  const entry = { type, at, end: Math.min(at + size, stsd.end) };
  const visual = ['avc1', 'avc3', 'hvc1', 'hev1', 'vp09', 'vp08', 'av01'].includes(type);
  const audio = ['mp4a', 'Opus'].includes(type);
  const out = { entry: type, codec: null, description: null, width: null, height: null, channels: null, sampleRate: null, colour: null };
  if (visual) {
    out.width = u16(b, at + 32);
    out.height = u16(b, at + 34);
    const inner = boxesIn(b, at + 86, entry.end);
    const cfg = inner.find((x) => ['avcC', 'hvcC', 'vpcC', 'av1C'].includes(x.type));
    const config = cfg ? b.slice(cfg.body, cfg.end) : null;
    out.codec = codecString(type, config);
    if (cfg && cfg.type !== 'vpcC') out.description = config;
    out.configBox = cfg ? cfg.type : null;
    const colr = inner.find((x) => x.type === 'colr' && dec(b, x.body, 4) === 'nclx');
    if (colr) out.colour = { primaries: u16(b, colr.body + 4), transfer: u16(b, colr.body + 6), matrix: u16(b, colr.body + 8), fullRange: Boolean(b[colr.body + 10] & 0x80) };
  } else if (audio) {
    out.channels = u16(b, at + 24);
    out.sampleRate = u32(b, at + 32) / 65536;
    const inner = boxesIn(b, at + 36, entry.end);
    const esds = inner.find((x) => x.type === 'esds');
    const dops = inner.find((x) => x.type === 'dOps');
    out.codec = codecString(type);
    if (esds) {
      out.description = ascOf(b, esds);
      if (out.description && out.description.length >= 2) out.codec = `mp4a.40.${out.description[0] >> 3}`;
    }
    if (dops) {
      out.description = opusHead(b, dops);
      out.sampleRate = 48000;
    }
    out.configBox = esds ? 'esds' : dops ? 'dOps' : null;
  } else {
    out.codec = null;
  }
  return out;
}

function table(b, box, width = 1) {
  if (!box) return null;
  const n = u32(b, box.body + 4);
  const rows = [];
  for (let i = 0; i < n; i++) {
    const at = box.body + 8 + i * width * 4;
    const row = [];
    for (let k = 0; k < width; k++) row.push(u32(b, at + k * 4));
    rows.push(row);
  }
  return rows;
}

function samplesOf(b, stbl) {
  const stts = table(b, child(b, stbl, 'stts'), 2) || [];
  const cttsBox = child(b, stbl, 'ctts');
  const ctts = cttsBox ? table(b, cttsBox, 2).map(([n, o]) => [n, b[cttsBox.body] === 1 ? (o | 0) : o]) : null;
  const stss = child(b, stbl, 'stss') ? table(b, child(b, stbl, 'stss'), 1).map((r) => r[0]) : null;
  const stszBox = child(b, stbl, 'stsz');
  const stsc = table(b, child(b, stbl, 'stsc'), 3) || [];
  const stcoBox = child(b, stbl, 'stco') || child(b, stbl, 'co64');
  const wide = stcoBox?.type === 'co64';
  const offsets = [];
  if (stcoBox) {
    const n = u32(b, stcoBox.body + 4);
    for (let i = 0; i < n; i++) offsets.push(wide ? u64(b, stcoBox.body + 8 + i * 8) : u32(b, stcoBox.body + 8 + i * 4));
  }
  const fixed = stszBox ? u32(b, stszBox.body + 4) : 0;
  const count = stszBox ? u32(b, stszBox.body + 8) : 0;
  const sizes = [];
  for (let i = 0; i < count; i++) sizes.push(fixed || u32(b, stszBox.body + 12 + i * 4));
  const keys = stss ? new Set(stss) : null;

  const samples = [];
  let dts = 0;
  let sttsRow = 0;
  let sttsLeft = stts[0]?.[0] ?? 0;
  let cttsRow = 0;
  let cttsLeft = ctts?.[0]?.[0] ?? 0;
  let index = 0;
  for (let c = 0; c < offsets.length && index < count; c++) {
    // stsc: the row that covers chunk c+1.
    let row = 0;
    while (row + 1 < stsc.length && stsc[row + 1][0] <= c + 1) row++;
    const per = stsc[row]?.[1] ?? 0;
    let at = offsets[c];
    for (let k = 0; k < per && index < count; k++) {
      while (sttsLeft === 0 && sttsRow + 1 < stts.length) sttsLeft = stts[++sttsRow][0];
      const duration = stts[sttsRow]?.[1] ?? 0;
      sttsLeft--;
      let off = 0;
      if (ctts) {
        while (cttsLeft === 0 && cttsRow + 1 < ctts.length) cttsLeft = ctts[++cttsRow][0];
        off = ctts[cttsRow]?.[1] ?? 0;
        cttsLeft--;
      }
      samples.push({ offset: at, size: sizes[index], dts, pts: dts + off, duration, key: keys ? keys.has(index + 1) : true });
      at += sizes[index];
      dts += duration;
      index++;
    }
  }
  return { samples, tables: { stts, ctts, stss, stsc, chunkOffsets: offsets, co64: wide, sizes } };
}

/**
 * Read an MP4. `b` is the whole file.
 *
 * @returns {{ brand: string, boxes: Array<{ type, at, size }>, timescale: number, duration: number,
 *   fragmented: boolean, tracks: Array<object> }|null}
 */
export function readMp4(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const top = boxesIn(b);
  if (!top.length || top[0].type !== 'ftyp') return null;
  const brand = dec(b, top[0].body, 4);
  const moov = top.find((x) => x.type === 'moov');
  if (!moov) return { brand, boxes: top.map(({ type, at, size }) => ({ type, at, size })), tracks: [], duration: 0, timescale: 0, fragmented: false };
  const mvhd = child(b, moov, 'mvhd');
  const v1 = mvhd && b[mvhd.body] === 1;
  const timescale = mvhd ? (v1 ? u32(b, mvhd.body + 20) : u32(b, mvhd.body + 12)) : 1000;
  const movieDur = mvhd ? (v1 ? u64(b, mvhd.body + 24) : u32(b, mvhd.body + 16)) : 0;
  const trex = new Map();
  const mvex = child(b, moov, 'mvex');
  if (mvex) {
    for (const x of children(b, mvex, 'trex')) trex.set(u32(b, x.body + 4), { duration: u32(b, x.body + 12), size: u32(b, x.body + 16), flags: u32(b, x.body + 20) });
  }

  const tracks = [];
  for (const trak of children(b, moov, 'trak')) {
    const tkhd = child(b, trak, 'tkhd');
    const tv1 = b[tkhd.body] === 1;
    const id = tv1 ? u32(b, tkhd.body + 20) : u32(b, tkhd.body + 12);
    const mdhd = path(b, trak, 'mdia', 'mdhd');
    const mv1 = b[mdhd.body] === 1;
    const ts = mv1 ? u32(b, mdhd.body + 20) : u32(b, mdhd.body + 12);
    const mediaDur = mv1 ? u64(b, mdhd.body + 24) : u32(b, mdhd.body + 16);
    const hdlr = path(b, trak, 'mdia', 'hdlr');
    const handler = hdlr ? dec(b, hdlr.body + 8, 4) : '';
    const kind = handler === 'vide' ? 'video' : handler === 'soun' ? 'audio' : 'other';
    const stbl = path(b, trak, 'mdia', 'minf', 'stbl');
    const entry = stbl && child(b, stbl, 'stsd') ? sampleEntry(b, child(b, stbl, 'stsd')) : {};
    const { samples, tables } = stbl ? samplesOf(b, stbl) : { samples: [], tables: null };
    const elst = path(b, trak, 'edts', 'elst');
    const edits = [];
    if (elst) {
      const ev1 = b[elst.body] === 1;
      const n = u32(b, elst.body + 4);
      for (let i = 0; i < n; i++) {
        const at = elst.body + 8 + i * (ev1 ? 20 : 12);
        edits.push(ev1 ? { duration: u64(b, at), mediaTime: Number(BigInt.asIntN(64, (BigInt(u32(b, at + 8)) << 32n) | BigInt(u32(b, at + 12)))) } : { duration: u32(b, at), mediaTime: i32(b, at + 4) });
      }
    }
    tracks.push({ id, kind, timescale: ts, movieTimescale: timescale, mediaDuration: mediaDur, duration: mediaDur / ts, edits, samples, tables, ...entry });
  }

  // Fragments: each moof's traf adds samples to its track.
  let fragmented = false;
  for (const moof of top.filter((x) => x.type === 'moof')) {
    fragmented = true;
    for (const traf of children(b, moof, 'traf')) {
      const tfhd = child(b, traf, 'tfhd');
      const flags = (b[tfhd.body + 1] << 16) | (b[tfhd.body + 2] << 8) | b[tfhd.body + 3];
      const id = u32(b, tfhd.body + 4);
      const track = tracks.find((t) => t.id === id);
      if (!track) continue;
      const defaults = trex.get(id) || { duration: 0, size: 0, flags: 0 };
      let p = tfhd.body + 8;
      let base = moof.at;
      if (flags & 0x01) { base = u64(b, p); p += 8; }
      if (flags & 0x02) p += 4;
      let defDur = defaults.duration;
      let defSize = defaults.size;
      let defFlags = defaults.flags;
      if (flags & 0x08) { defDur = u32(b, p); p += 4; }
      if (flags & 0x10) { defSize = u32(b, p); p += 4; }
      if (flags & 0x20) { defFlags = u32(b, p); p += 4; }
      const tfdt = child(b, traf, 'tfdt');
      let dts = tfdt ? (b[tfdt.body] === 1 ? u64(b, tfdt.body + 4) : u32(b, tfdt.body + 4)) : (track.samples.length ? track.samples[track.samples.length - 1].dts + track.samples[track.samples.length - 1].duration : 0);
      let dataEnd = null;
      for (const trun of children(b, traf, 'trun')) {
        const tf = (b[trun.body + 1] << 16) | (b[trun.body + 2] << 8) | b[trun.body + 3];
        const signed = b[trun.body] === 1;
        const count = u32(b, trun.body + 4);
        let q = trun.body + 8;
        let offset = dataEnd ?? base;
        if (tf & 0x01) { offset = base + i32(b, q); q += 4; }
        let firstFlags = null;
        if (tf & 0x04) { firstFlags = u32(b, q); q += 4; }
        for (let i = 0; i < count; i++) {
          let duration = defDur;
          let size = defSize;
          let sflags = i === 0 && firstFlags !== null ? firstFlags : defFlags;
          let off = 0;
          if (tf & 0x100) { duration = u32(b, q); q += 4; }
          if (tf & 0x200) { size = u32(b, q); q += 4; }
          if (tf & 0x400) { sflags = u32(b, q); q += 4; }
          if (tf & 0x800) { off = signed ? i32(b, q) : u32(b, q); q += 4; }
          // sample_is_non_sync_sample is bit 16 of the flags.
          track.samples.push({ offset, size, dts, pts: dts + off, duration, key: !((sflags >> 16) & 1) });
          offset += size;
          dts += duration;
        }
        dataEnd = offset;
      }
    }
  }
  if (fragmented) {
    for (const t of tracks) {
      const last = t.samples[t.samples.length - 1];
      if (last) t.duration = Math.max(t.duration, (last.dts + last.duration - (t.samples[0]?.dts || 0)) / t.timescale);
    }
  }

  return {
    brand,
    boxes: top.map(({ type, at, size }) => ({ type, at, size })),
    timescale,
    duration: fragmented ? Math.max(movieDur / timescale, ...tracks.map((t) => t.duration)) : movieDur / timescale,
    fragmented,
    tracks,
  };
}

/** Where a track's samples show, in seconds, after its edit list: [{ pts, ... }] with `time`. */
export function presentationTimes(track) {
  const ts = track.timescale || 1;
  let shift = 0;
  let empty = 0;
  for (const e of track.edits || []) {
    if (e.mediaTime === -1) empty += e.duration;
    else { shift = e.mediaTime; break; }
  }
  // An empty edit's duration is in the movie's timescale.
  const lead = track.movieTimescale ? empty / track.movieTimescale : 0;
  return track.samples.map((s) => ({ ...s, time: lead + (s.pts - shift) / ts }));
}
