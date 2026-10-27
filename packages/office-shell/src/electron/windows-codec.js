// Pictures Chromium cannot draw, through the codecs Windows has: HEIC and
// HEIF (the HEIF and HEVC extensions from the Store), camera RAW (the Raw
// Image extension) and JPEG XR (built in), decoded by the Windows Imaging
// Component at the screen's size (2048 pixels on the long side, not the
// 256 to 512 pixel shell thumbnail), and handed back as PNG. No HEIC decoder
// is bundled (the HEVC patents): where Windows has no codec the answer is
// null, and the picture says so.
//
// The route is the one the system already offers every program: WPF's
// BitmapDecoder in a short PowerShell (no native module, nothing installed),
// the file's path passed in the environment rather than in the command, the
// PNG written to standard output. It scales while it decodes where the codec
// can, and turns the picture upright by its orientation.
//
// A camera RAW or DNG with no codec still has a picture: every camera puts a
// JPEG of what it saw inside the file (usually full size), which is read
// here directly from the TIFF directories (and a Fujifilm RAF's header) and
// drawn at the size asked for.

import { spawn } from 'node:child_process';
import fsp from 'node:fs/promises';
import path from 'node:path';

/** The kinds of file this is for. */
export const CODEC_EXTS = ['.heic', '.heif', '.hif', '.avci', '.jxr', '.wdp', '.hdp', '.dng', '.cr2', '.nef', '.nrw', '.arw', '.srw', '.orf', '.rw2', '.raf', '.pef', '.3fr', '.erf', '.kdc', '.mrw', '.dcr', '.x3f', '.iiq', '.rwl'];
const RAW_EXTS = new Set(['.dng', '.cr2', '.nef', '.nrw', '.arw', '.srw', '.orf', '.rw2', '.raf', '.pef', '.3fr', '.erf', '.kdc', '.mrw', '.dcr', '.x3f', '.iiq', '.rwl']);
const HEIF_EXTS = new Set(['.heic', '.heif', '.hif', '.avci']);

/** The longest side asked of Windows: the screen's size. */
export const MAX_SIDE = 2048;
/** Past this a decode is given up: a codec that hangs never holds a window. */
const TIMEOUT_MS = 30000;
/** The PNG read back from the decoder at most (2048 × 2048 RGBA, uncompressed, is 16 MB). */
const MAX_OUT = 40 * 1024 * 1024;
/** A RAW's preview is read from at most this much of the file's start, for its directories. */
const HEAD = 256 * 1024;
const MAX_PREVIEW = 64 * 1024 * 1024;

// ── the Windows Imaging Component, through PowerShell ───────────────────────

// Exit codes: 0 a PNG written; 3 no codec for this file; anything else a failure.
const SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationCore
$file = $env:RUTBA_CODEC_FILE
$max = [double]$env:RUTBA_CODEC_MAX
$turn = $env:RUTBA_CODEC_TURN -eq '1'
$stream = [IO.File]::OpenRead($file)
try {
  try {
    $decoder = [Windows.Media.Imaging.BitmapDecoder]::Create($stream, [Windows.Media.Imaging.BitmapCreateOptions]::IgnoreColorProfile, [Windows.Media.Imaging.BitmapCacheOption]::None)
  } catch [System.NotSupportedException] { exit 3 } catch [System.IO.FileFormatException] { exit 3 }
  $frame = $decoder.Frames[0]
  $o = 1
  if ($turn) { try { $q = $frame.Metadata.GetQuery('System.Photo.Orientation'); if ($q) { $o = [int]$q } } catch { $o = 1 } }
  $k = [Math]::Min(1.0, $max / [Math]::Max($frame.PixelWidth, $frame.PixelHeight))
  $group = New-Object Windows.Media.TransformGroup
  if ($k -lt 1.0) { $group.Children.Add((New-Object Windows.Media.ScaleTransform($k, $k))) }
  switch ($o) {
    2 { $group.Children.Add((New-Object Windows.Media.ScaleTransform(-1, 1))) }
    3 { $group.Children.Add((New-Object Windows.Media.RotateTransform(180))) }
    4 { $group.Children.Add((New-Object Windows.Media.ScaleTransform(1, -1))) }
    5 { $group.Children.Add((New-Object Windows.Media.RotateTransform(90))); $group.Children.Add((New-Object Windows.Media.ScaleTransform(-1, 1))) }
    6 { $group.Children.Add((New-Object Windows.Media.RotateTransform(90))) }
    7 { $group.Children.Add((New-Object Windows.Media.RotateTransform(270))); $group.Children.Add((New-Object Windows.Media.ScaleTransform(-1, 1))) }
    8 { $group.Children.Add((New-Object Windows.Media.RotateTransform(270))) }
  }
  $source = $frame
  if ($group.Children.Count -gt 0) { $source = New-Object Windows.Media.Imaging.TransformedBitmap($frame, $group) }
  $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
  $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($source))
  # The encoder wants a stream it can seek: memory first, then standard output.
  $memory = New-Object IO.MemoryStream
  $encoder.Save($memory)
  $out = [Console]::OpenStandardOutput()
  $memory.WriteTo($out)
  $out.Flush()
} finally { $stream.Close() }
exit 0
`;

/**
 * The picture decoded by Windows at `max` pixels on its long side at most,
 * upright, as PNG bytes; null where Windows has no codec for it (or is not
 * Windows), or the decode failed or took too long.
 */
export function decodeThroughWindows(file, max = MAX_SIDE, { timeoutMs = TIMEOUT_MS, turn = true } = {}) {
  if (process.platform !== 'win32') return Promise.resolve(null);
  return new Promise((resolve) => {
    let child;
    try {
      child = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', SCRIPT], {
        windowsHide: true,
        stdio: ['ignore', 'pipe', 'ignore'],
        env: { ...process.env, RUTBA_CODEC_FILE: file, RUTBA_CODEC_MAX: String(Math.max(16, Math.min(8192, Math.round(max)))), RUTBA_CODEC_TURN: turn ? '1' : '0' },
      });
    } catch {
      resolve(null);
      return;
    }
    const parts = [];
    let size = 0;
    let done = false;
    const finish = (value) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      resolve(value);
    };
    const timer = setTimeout(() => { try { child.kill(); } catch { /* gone */ } finish(null); }, timeoutMs);
    child.stdout.on('data', (chunk) => {
      size += chunk.length;
      if (size > MAX_OUT) { try { child.kill(); } catch { /* gone */ } finish(null); return; }
      parts.push(chunk);
    });
    child.on('error', () => finish(null));
    child.on('close', (code) => {
      const bytes = Buffer.concat(parts);
      const png = bytes.length > 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47;
      finish(code === 0 && png ? bytes : null);
    });
  });
}

// ── a camera's own JPEG inside its RAW ──────────────────────────────────────

const u16 = (b, o, le) => (le ? b[o] | (b[o + 1] << 8) : (b[o] << 8) | b[o + 1]);
const u32 = (b, o, le) => (le ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16)) + b[o + 3] * 0x1000000 : b[o] * 0x1000000 + ((b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]));

/** A JPEG's size and kind from its frame header: { width, height, lossless } or null. */
function jpegFrame(b, start, end) {
  if (b[start] !== 0xff || b[start + 1] !== 0xd8) return null;
  let at = start + 2;
  for (let n = 0; n < 512 && at + 9 < end; n++) {
    if (b[at] !== 0xff) return null;
    const marker = b[at + 1];
    if (marker === 0xff) { at += 1; continue; }
    const length = (b[at + 2] << 8) | b[at + 3];
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { height: (b[at + 5] << 8) | b[at + 6], width: (b[at + 7] << 8) | b[at + 8], lossless: marker === 0xc3 || marker === 0xc7 || marker === 0xcb || marker === 0xcf };
    }
    if (marker === 0xda || length < 2) return null;
    at += 2 + length;
  }
  return null;
}

/**
 * The JPEG pictures a TIFF-based RAW (DNG, CR2, NEF, ARW, ORF, RW2, PEF…)
 * points to from its directories and their SubIFDs, and the orientation of
 * its first: `{ previews: [{ offset, length }], orientation }`. Reads only
 * `head` (the file's start); a preview's own bytes are fetched later.
 */
export function rawPreviews(head, size = head.length) {
  const b = head;
  const out = { previews: [], orientation: 1 };
  if (b.length < 16) return out;
  // Fujifilm: a header of its own, the JPEG's place at 84 and its length at 88.
  if (String.fromCharCode(...b.subarray(0, 15)) === 'FUJIFILMCCD-RAW') {
    const offset = u32(b, 84, false);
    const length = u32(b, 88, false);
    if (offset > 0 && length > 0 && offset + length <= size) out.previews.push({ offset, length });
    return out;
  }
  const le = b[0] === 0x49 && b[1] === 0x49;
  if (!le && !(b[0] === 0x4d && b[1] === 0x4d)) return out;
  // 42 for TIFF; Olympus writes 0x4f52 or 0x5352, Panasonic 0x55.
  if (![42, 0x4f52, 0x5352, 0x55].includes(u16(b, 2, le))) return out;
  const seen = new Set();
  const queue = [{ at: u32(b, 4, le), depth: 0 }];
  let first = true;
  for (let steps = 0; queue.length && steps < 64; steps++) {
    const { at, depth } = queue.shift();
    if (!at || seen.has(at) || at + 2 > b.length) continue;
    seen.add(at);
    const count = u16(b, at, le);
    if (count > 1000 || at + 2 + count * 12 > b.length) continue;
    const tags = new Map();
    for (let i = 0; i < count; i++) {
      const e = at + 2 + i * 12;
      const tag = u16(b, e, le);
      const type = u16(b, e + 2, le);
      const n = u32(b, e + 4, le);
      const value = type === 3 ? u16(b, e + 8, le) : u32(b, e + 8, le);
      tags.set(tag, { type, n, value, at: e + 8 });
    }
    if (first && tags.has(0x0112)) out.orientation = tags.get(0x0112).value || 1;
    first = false;
    // A JPEG by its interchange format, or as one strip whose compression says JPEG.
    const jif = tags.get(0x0201);
    const jifLength = tags.get(0x0202);
    if (jif && jifLength && jifLength.value > 0) out.previews.push({ offset: jif.value, length: jifLength.value });
    const compression = tags.get(0x0103)?.value;
    const strips = tags.get(0x0111);
    const counts = tags.get(0x0117);
    if ((compression === 6 || compression === 7) && strips && counts && strips.n === 1 && counts.n === 1) out.previews.push({ offset: strips.value, length: counts.value });
    // The directories below: SubIFDs (a list, or one inline) and the next.
    const sub = tags.get(0x014a);
    if (sub && depth < 3) {
      if (sub.n === 1) queue.push({ at: sub.value, depth: depth + 1 });
      else if (sub.value + sub.n * 4 <= b.length) for (let k = 0; k < Math.min(sub.n, 16); k++) queue.push({ at: u32(b, sub.value + k * 4, le), depth: depth + 1 });
    }
    const next = at + 2 + count * 12;
    if (next + 4 <= b.length) queue.push({ at: u32(b, next, le), depth });
  }
  out.previews = out.previews.filter((p) => p.offset > 0 && p.length > 0 && p.offset + p.length <= size && p.length <= MAX_PREVIEW);
  return out;
}

/** The orientation turned into a bitmap: BGRA `data` of w × h, answered upright. */
export function turnBitmap(data, width, height, orientation) {
  if (!(orientation >= 2 && orientation <= 8)) return { data, width, height };
  const swap = orientation >= 5;
  const W = swap ? height : width;
  const H = swap ? width : height;
  const out = Buffer.alloc(W * H * 4);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      let nx; let ny;
      switch (orientation) {
        case 2: nx = width - 1 - x; ny = y; break;
        case 3: nx = width - 1 - x; ny = height - 1 - y; break;
        case 4: nx = x; ny = height - 1 - y; break;
        case 5: nx = y; ny = x; break;
        case 6: nx = height - 1 - y; ny = x; break;
        case 7: nx = height - 1 - y; ny = width - 1 - x; break;
        default: nx = y; ny = width - 1 - x; break; // 8
      }
      const s = (y * width + x) * 4;
      const d = (ny * W + nx) * 4;
      out[d] = data[s]; out[d + 1] = data[s + 1]; out[d + 2] = data[s + 2]; out[d + 3] = data[s + 3];
    }
  }
  return { data: out, width: W, height: H };
}

async function readRange(file, offset, length) {
  const handle = await fsp.open(file, 'r');
  try {
    const buf = Buffer.alloc(length);
    const { bytesRead } = await handle.read(buf, 0, length, offset);
    return buf.subarray(0, bytesRead);
  } finally {
    await handle.close();
  }
}

/**
 * The largest picture JPEG a RAW carries (never its lossless raw data), at
 * `max` on its long side at most, upright by the RAW's orientation, as
 * `{ png, width, height, stored }` (the preview's own size), or null.
 */
export async function rawPreview(file, max = MAX_SIDE) {
  let stat;
  try { stat = await fsp.stat(file); } catch { return null; }
  const head = await readRange(file, 0, Math.min(HEAD, stat.size)).catch(() => null);
  if (!head) return null;
  const { previews, orientation } = rawPreviews(head, stat.size);
  let best = null;
  for (const p of previews) {
    // Its frame header tells a picture from the lossless raw data, and its size.
    const start = await readRange(file, p.offset, Math.min(p.length, 64 * 1024)).catch(() => null);
    const frame = start && jpegFrame(start, 0, start.length);
    if (!frame || frame.lossless || !frame.width) continue;
    if (!best || frame.width * frame.height > best.frame.width * best.frame.height) best = { ...p, frame };
  }
  if (!best) return null;
  const jpeg = await readRange(file, best.offset, best.length).catch(() => null);
  if (!jpeg) return null;
  const { nativeImage } = await import('electron');
  let image = nativeImage.createFromBuffer(jpeg);
  if (image.isEmpty()) return null;
  const size = image.getSize();
  const k = Math.min(1, max / Math.max(size.width, size.height));
  if (k < 1) image = image.resize({ width: Math.max(1, Math.round(size.width * k)), height: Math.max(1, Math.round(size.height * k)), quality: 'best' });
  let { width, height } = image.getSize();
  if (orientation !== 1) {
    const turned = turnBitmap(image.toBitmap(), width, height, orientation);
    image = nativeImage.createFromBitmap(turned.data, { width: turned.width, height: turned.height });
    ({ width, height } = turned);
  }
  return { png: image.toPNG(), width, height, stored: { width: size.width, height: size.height } };
}

/**
 * The shell's picture of a file at `max` on its long side (Windows answers
 * up to 2560), as PNG bytes; null where it has none, or only one smaller
 * than a screen's worth (a provider that hands back its small embedded
 * thumbnail), which the Imaging Component then decodes properly.
 */
export async function shellPicture(file, max = MAX_SIDE) {
  if (process.platform !== 'win32') return null;
  try {
    const { nativeImage } = await import('electron');
    const image = await nativeImage.createThumbnailFromPath(path.resolve(file), { width: max, height: max });
    if (!image || image.isEmpty()) return null;
    const { width, height } = image.getSize();
    if (Math.max(width, height) < Math.min(max, 1024)) return null;
    return image.toPNG();
  } catch {
    return null;
  }
}

/**
 * A picture Chromium cannot draw, as PNG bytes at `max` on its long side at
 * most, upright; null where neither Windows nor the file itself has one.
 * HEIC, HEIF and JPEG XR go to Windows' codecs; a camera RAW goes to its
 * codec first (the Raw Image extension develops the sensor's data), then to
 * the JPEG its camera put inside it.
 */
export async function decodeWithSystem(file, max = MAX_SIDE) {
  const ext = path.extname(String(file)).toLowerCase();
  if (!CODEC_EXTS.includes(ext)) return null;
  // The shell's own picture first: the same codecs behind it, no process
  // started, upright, and kept in Windows' thumbnail cache for the next look.
  const fromShell = await shellPicture(file, max);
  if (fromShell) return fromShell;
  // A HEIF's decoder turns it by the file's own transform boxes: its EXIF orientation is not applied again.
  const fromWindows = await decodeThroughWindows(file, max, { turn: !HEIF_EXTS.has(ext) });
  if (fromWindows) return fromWindows;
  if (RAW_EXTS.has(ext)) return (await rawPreview(file, max))?.png || null;
  return null;
}
