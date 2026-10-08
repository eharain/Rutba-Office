// 3D models: a glTF 2.0 file read into triangles, ready to draw.
//
// A `.glb` is a small header and two chunks, the JSON that describes the
// scene and the binary buffer its numbers live in; a `.gltf` is the JSON
// alone, its buffers in data: addresses or in files beside it (handed in by
// the caller, who reads them). The default scene's nodes are walked with
// their transforms, each mesh's triangles put into the model's own space —
// positions, normals, texture coordinates and vertex colours — with the
// material's base colour, its texture if it has one, and whether it is cut
// out or double-sided. What this does not read is said in a sentence: a mesh
// compressed with Draco or meshopt, which needs a decoder this suite does
// not ship.

const COMPONENTS = { 5120: Int8Array, 5121: Uint8Array, 5122: Int16Array, 5123: Uint16Array, 5125: Uint32Array, 5126: Float32Array };
const WIDTH = { SCALAR: 1, VEC2: 2, VEC3: 3, VEC4: 4, MAT2: 4, MAT3: 9, MAT4: 16 };
const NORMALISE = { 5120: 127, 5121: 255, 5122: 32767, 5123: 65535 };

/** Whether these bytes are a binary glTF. */
export const isGlb = (bytes) => bytes.length >= 12 && bytes[0] === 0x67 && bytes[1] === 0x6c && bytes[2] === 0x54 && bytes[3] === 0x46;

/** A `.glb` split into its JSON and its binary chunk. */
export function parseGlb(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (!isGlb(u8)) throw new Error('This is not a binary glTF (.glb) file');
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const version = dv.getUint32(4, true);
  if (version !== 2) throw new Error(`This is a glTF ${version} file; only glTF 2.0 is read`);
  const length = Math.min(dv.getUint32(8, true), u8.length);
  let at = 12;
  let json = null;
  let bin = null;
  while (at + 8 <= length) {
    const size = dv.getUint32(at, true);
    const type = dv.getUint32(at + 4, true);
    const body = u8.subarray(at + 8, at + 8 + size);
    if (type === 0x4e4f534a) json = JSON.parse(new TextDecoder().decode(body));
    else if (type === 0x004e4942 && !bin) bin = body;
    at += 8 + size + ((4 - (size % 4)) % 4);
  }
  if (!json) throw new Error('The .glb file has no scene description');
  return { json, bin };
}

/** The bytes of a data: address, or null for an address that is not one. */
function dataUri(uri) {
  const m = /^data:[^;,]*(;base64)?,(.*)$/s.exec(String(uri || ''));
  if (!m) return null;
  if (m[1]) {
    const bin = typeof atob === 'function' ? atob(m[2]) : Buffer.from(m[2], 'base64').toString('binary');
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  return new TextEncoder().encode(decodeURIComponent(m[2]));
}

/**
 * The file's buffers: the .glb's own chunk, data: addresses, and the files a
 * `.gltf` names, read with `readFile(uri)` (bytes, or null when it cannot be).
 */
function buffersOf(json, bin, readFile) {
  return (json.buffers || []).map((b, i) => {
    if (b.uri == null) {
      if (i === 0 && bin) return bin;
      throw new Error('The model\'s buffer is missing');
    }
    const inline = dataUri(b.uri);
    if (inline) return inline;
    const bytes = readFile ? readFile(decodeURIComponent(b.uri)) : null;
    if (!bytes) throw new Error(`The model needs the file "${b.uri}" beside it`);
    return bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  });
}

function viewBytes(json, buffers, index) {
  const v = json.bufferViews?.[index];
  if (!v) throw new Error('The model points at a buffer view it does not have');
  const buf = buffers[v.buffer];
  return { bytes: buf.subarray(v.byteOffset || 0, (v.byteOffset || 0) + v.byteLength), stride: v.byteStride || 0 };
}

/** An accessor's values as a flat Float32Array (or Uint32Array for indices), normalised as asked. */
function readAccessor(json, buffers, index, { asIndex = false } = {}) {
  const a = json.accessors?.[index];
  if (!a) throw new Error('The model points at an accessor it does not have');
  const width = WIDTH[a.type] || 1;
  const Type = COMPONENTS[a.componentType];
  if (!Type) throw new Error(`The model stores numbers in a form (${a.componentType}) this does not read`);
  const out = asIndex ? new Uint32Array(a.count * width) : new Float32Array(a.count * width);
  if (a.bufferView != null) {
    const { bytes, stride } = viewBytes(json, buffers, a.bufferView);
    const size = Type.BYTES_PER_ELEMENT;
    const step = stride || size * width;
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const get = {
      5120: (o) => dv.getInt8(o), 5121: (o) => dv.getUint8(o), 5122: (o) => dv.getInt16(o, true),
      5123: (o) => dv.getUint16(o, true), 5125: (o) => dv.getUint32(o, true), 5126: (o) => dv.getFloat32(o, true),
    }[a.componentType];
    const base = a.byteOffset || 0;
    const scale = a.normalized && NORMALISE[a.componentType] ? 1 / NORMALISE[a.componentType] : 1;
    for (let i = 0; i < a.count; i++) {
      for (let k = 0; k < width; k++) {
        const v = get(base + i * step + k * size);
        out[i * width + k] = scale === 1 ? v : Math.max(-1, v * scale);
      }
    }
  }
  if (a.sparse) {
    const idx = a.sparse.indices;
    const val = a.sparse.values;
    const ib = viewBytes(json, buffers, idx.bufferView);
    const vb = viewBytes(json, buffers, val.bufferView);
    const IT = COMPONENTS[idx.componentType];
    const ids = new IT(ib.bytes.slice((idx.byteOffset || 0), (idx.byteOffset || 0) + a.sparse.count * IT.BYTES_PER_ELEMENT).buffer);
    const vals = new Type(vb.bytes.slice((val.byteOffset || 0), (val.byteOffset || 0) + a.sparse.count * width * Type.BYTES_PER_ELEMENT).buffer);
    for (let i = 0; i < a.sparse.count; i++) for (let k = 0; k < width; k++) out[ids[i] * width + k] = vals[i * width + k];
  }
  return { values: out, width, count: a.count, min: a.min, max: a.max };
}

/* ── matrices, column-major as glTF writes them ─────────────────────────── */

const identity = () => [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1];
function multiply(a, b) {
  const o = new Array(16);
  for (let c = 0; c < 4; c++) for (let r = 0; r < 4; r++) {
    o[c * 4 + r] = a[r] * b[c * 4] + a[4 + r] * b[c * 4 + 1] + a[8 + r] * b[c * 4 + 2] + a[12 + r] * b[c * 4 + 3];
  }
  return o;
}
function compose(t = [0, 0, 0], q = [0, 0, 0, 1], s = [1, 1, 1]) {
  const [x, y, z, w] = q;
  const xx = x * x; const yy = y * y; const zz = z * z;
  const xy = x * y; const xz = x * z; const yz = y * z;
  const wx = w * x; const wy = w * y; const wz = w * z;
  return [
    (1 - 2 * (yy + zz)) * s[0], 2 * (xy + wz) * s[0], 2 * (xz - wy) * s[0], 0,
    2 * (xy - wz) * s[1], (1 - 2 * (xx + zz)) * s[1], 2 * (yz + wx) * s[1], 0,
    2 * (xz + wy) * s[2], 2 * (yz - wx) * s[2], (1 - 2 * (xx + yy)) * s[2], 0,
    t[0], t[1], t[2], 1,
  ];
}
/** The upper 3×3 of a matrix's inverse transpose, for normals. */
function normalMatrix(m) {
  const a = m[0], b = m[4], c = m[8], d = m[1], e = m[5], f = m[9], g = m[2], h = m[6], i = m[10];
  const A = e * i - f * h, B = -(d * i - f * g), C = d * h - e * g;
  const det = a * A + b * B + c * C || 1;
  return [
    A / det, B / det, C / det,
    -(b * i - c * h) / det, (a * i - c * g) / det, -(a * h - b * g) / det,
    (b * f - c * e) / det, -(a * f - c * d) / det, (a * e - b * d) / det,
  ];
}

/**
 * A glTF read into what a renderer draws: `{ meshes: [{ positions, normals,
 * uvs, colors, indices, material }], images: [{ mimeType, bytes }], bounds:
 * { min, max, centre, radius }, triangles }`. `readFile(uri)` gives the
 * bytes of a file a `.gltf` names beside it.
 */
export function readModel(bytes, { readFile = null } = {}) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let json;
  let bin = null;
  if (isGlb(u8)) ({ json, bin } = parseGlb(u8));
  else {
    try { json = JSON.parse(new TextDecoder().decode(u8)); } catch { throw new Error('This is not a glTF file: it is neither a .glb nor glTF\'s JSON'); }
  }
  if (!json.asset || String(json.asset.version || '').split('.')[0] !== '2') throw new Error('Only glTF 2.0 models are read');
  const required = json.extensionsRequired || [];
  if (required.includes('KHR_draco_mesh_compression')) throw new Error('The model is compressed with Draco, which this suite does not decode');
  if (required.some((x) => /meshopt/i.test(x))) throw new Error('The model is compressed with meshopt, which this suite does not decode');
  const buffers = buffersOf(json, bin, readFile);

  const images = (json.images || []).map((img) => {
    if (img.bufferView != null) return { mimeType: img.mimeType || 'image/png', bytes: viewBytes(json, buffers, img.bufferView).bytes };
    const inline = dataUri(img.uri);
    if (inline) return { mimeType: /^data:([^;,]+)/.exec(img.uri)?.[1] || 'image/png', bytes: inline };
    const file = readFile && img.uri ? readFile(decodeURIComponent(img.uri)) : null;
    return file ? { mimeType: /\.jpe?g$/i.test(img.uri) ? 'image/jpeg' : 'image/png', bytes: new Uint8Array(file) } : null;
  });
  const materialOf = (index) => {
    const m = json.materials?.[index] || {};
    const pbr = m.pbrMetallicRoughness || {};
    const texIndex = pbr.baseColorTexture?.index;
    const tex = texIndex != null ? json.textures?.[texIndex] : null;
    const sampler = tex?.sampler != null ? json.samplers?.[tex.sampler] : null;
    return {
      color: pbr.baseColorFactor || [1, 1, 1, 1],
      image: tex?.source ?? null,
      texCoord: pbr.baseColorTexture?.texCoord || 0,
      wrap: [sampler?.wrapS ?? 10497, sampler?.wrapT ?? 10497],
      emissive: m.emissiveFactor || [0, 0, 0],
      metallic: pbr.metallicFactor ?? 1,
      roughness: pbr.roughnessFactor ?? 1,
      alphaMode: m.alphaMode || 'OPAQUE',
      alphaCutoff: m.alphaCutoff ?? 0.5,
      doubleSided: Boolean(m.doubleSided),
      unlit: Boolean(m.extensions?.KHR_materials_unlit),
    };
  };

  const meshes = [];
  const min = [Infinity, Infinity, Infinity];
  const max = [-Infinity, -Infinity, -Infinity];
  let triangles = 0;
  const visit = (nodeIndex, parent, depth) => {
    if (depth > 64) return;
    const node = json.nodes?.[nodeIndex];
    if (!node) return;
    const local = node.matrix ? node.matrix.slice() : compose(node.translation, node.rotation, node.scale);
    const world = multiply(parent, local);
    if (node.mesh != null) {
      const nm = normalMatrix(world);
      for (const prim of json.meshes?.[node.mesh]?.primitives || []) {
        const mode = prim.mode ?? 4;
        if (mode !== 4 && mode !== 5 && mode !== 6) continue; // points and lines are not drawn
        if (prim.attributes?.POSITION == null) continue;
        const pos = readAccessor(json, buffers, prim.attributes.POSITION);
        const count = pos.count;
        const positions = new Float32Array(count * 3);
        for (let i = 0; i < count; i++) {
          const x = pos.values[i * 3], y = pos.values[i * 3 + 1], z = pos.values[i * 3 + 2];
          const wx = world[0] * x + world[4] * y + world[8] * z + world[12];
          const wy = world[1] * x + world[5] * y + world[9] * z + world[13];
          const wz = world[2] * x + world[6] * y + world[10] * z + world[14];
          positions[i * 3] = wx; positions[i * 3 + 1] = wy; positions[i * 3 + 2] = wz;
          if (wx < min[0]) min[0] = wx; if (wy < min[1]) min[1] = wy; if (wz < min[2]) min[2] = wz;
          if (wx > max[0]) max[0] = wx; if (wy > max[1]) max[1] = wy; if (wz > max[2]) max[2] = wz;
        }
        let normals = null;
        if (prim.attributes.NORMAL != null) {
          const n = readAccessor(json, buffers, prim.attributes.NORMAL).values;
          normals = new Float32Array(count * 3);
          for (let i = 0; i < count; i++) {
            const x = n[i * 3], y = n[i * 3 + 1], z = n[i * 3 + 2];
            let nx = nm[0] * x + nm[1] * y + nm[2] * z;
            let ny = nm[3] * x + nm[4] * y + nm[5] * z;
            let nz = nm[6] * x + nm[7] * y + nm[8] * z;
            const len = Math.hypot(nx, ny, nz) || 1;
            nx /= len; ny /= len; nz /= len;
            normals[i * 3] = nx; normals[i * 3 + 1] = ny; normals[i * 3 + 2] = nz;
          }
        }
        const material = materialOf(prim.material);
        const uvKey = `TEXCOORD_${material.texCoord}`;
        const uvs = prim.attributes[uvKey] != null ? readAccessor(json, buffers, prim.attributes[uvKey]).values : null;
        let colors = null;
        if (prim.attributes.COLOR_0 != null) {
          const c = readAccessor(json, buffers, prim.attributes.COLOR_0);
          colors = new Float32Array(count * 4);
          for (let i = 0; i < count; i++) for (let k = 0; k < 4; k++) colors[i * 4 + k] = k < c.width ? c.values[i * c.width + k] : 1;
        }
        let indices = prim.indices != null ? readAccessor(json, buffers, prim.indices, { asIndex: true }).values : Uint32Array.from({ length: count }, (_, i) => i);
        if (mode === 5) {
          const tri = [];
          for (let i = 0; i + 2 < indices.length; i++) tri.push(...(i % 2 ? [indices[i + 1], indices[i], indices[i + 2]] : [indices[i], indices[i + 1], indices[i + 2]]));
          indices = Uint32Array.from(tri);
        } else if (mode === 6) {
          const tri = [];
          for (let i = 1; i + 1 < indices.length; i++) tri.push(indices[0], indices[i], indices[i + 1]);
          indices = Uint32Array.from(tri);
        }
        // A mirrored node (negative determinant) turns its triangles inside out.
        const det = world[0] * (world[5] * world[10] - world[9] * world[6]) - world[4] * (world[1] * world[10] - world[9] * world[2]) + world[8] * (world[1] * world[6] - world[5] * world[2]);
        if (det < 0) for (let i = 0; i + 2 < indices.length; i += 3) { const t = indices[i + 1]; indices[i + 1] = indices[i + 2]; indices[i + 2] = t; }
        triangles += Math.floor(indices.length / 3);
        meshes.push({ positions, normals, uvs, colors, indices, material });
      }
    }
    for (const child of node.children || []) visit(child, world, depth + 1);
  };
  const scene = json.scenes?.[json.scene ?? 0];
  const roots = scene ? scene.nodes || [] : (json.nodes || []).map((_, i) => i).filter((i) => !(json.nodes || []).some((n) => (n.children || []).includes(i)));
  for (const r of roots) visit(r, identity(), 0);
  if (!meshes.length) throw new Error('The model has nothing to draw: no triangles in its scene');
  const centre = [(min[0] + max[0]) / 2, (min[1] + max[1]) / 2, (min[2] + max[2]) / 2];
  let radius = 0;
  for (const m of meshes) for (let i = 0; i < m.positions.length; i += 3) radius = Math.max(radius, Math.hypot(m.positions[i] - centre[0], m.positions[i + 1] - centre[1], m.positions[i + 2] - centre[2]));
  return { meshes, images, bounds: { min, max, centre, radius: radius || 1 }, triangles, name: json.asset?.generator || null };
}

/**
 * A model as one self-contained `.glb`: a `.glb` as it is, a `.gltf` with its
 * buffers and pictures — in data: addresses or in files beside it, read with
 * `readFile(uri)` — gathered into the binary chunk, so a document can keep
 * the model whole in one part.
 */
export function toGlb(bytes, { readFile = null } = {}) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (isGlb(u8)) return u8;
  let json;
  try { json = JSON.parse(new TextDecoder().decode(u8)); } catch { throw new Error('This is not a glTF file: it is neither a .glb nor glTF\'s JSON'); }
  const buffers = buffersOf(json, null, readFile);
  const parts = [];
  let length = 0;
  const add = (data) => {
    const offset = length;
    parts.push(data);
    length += data.length;
    const pad = (4 - (length % 4)) % 4;
    if (pad) { parts.push(new Uint8Array(pad)); length += pad; }
    return offset;
  };
  // Every buffer laid end to end in the one buffer, its views moved along with it.
  const starts = buffers.map((b) => add(b));
  const out = JSON.parse(JSON.stringify(json));
  for (const v of out.bufferViews || []) { v.byteOffset = (v.byteOffset || 0) + starts[v.buffer]; v.buffer = 0; }
  // A picture given by address becomes a view of its own.
  for (const img of out.images || []) {
    if (img.bufferView != null || !img.uri) continue;
    const inline = dataUri(img.uri);
    const data = inline || (readFile ? readFile(decodeURIComponent(img.uri)) : null);
    if (!data) throw new Error(`The model needs the picture "${img.uri}" beside it`);
    const bytesOf = data instanceof Uint8Array ? data : new Uint8Array(data);
    const mime = inline ? /^data:([^;,]+)/.exec(img.uri)?.[1] : /\.jpe?g$/i.test(img.uri) ? 'image/jpeg' : 'image/png';
    out.bufferViews = out.bufferViews || [];
    out.bufferViews.push({ buffer: 0, byteOffset: add(bytesOf), byteLength: bytesOf.length });
    img.bufferView = out.bufferViews.length - 1;
    img.mimeType = mime || 'image/png';
    delete img.uri;
  }
  out.buffers = [{ byteLength: length }];
  const bin = new Uint8Array(length);
  let at = 0;
  for (const p of parts) { bin.set(p, at); at += p.length; }
  const text = new TextEncoder().encode(JSON.stringify(out));
  const jsonLen = text.length + ((4 - (text.length % 4)) % 4);
  const total = 12 + 8 + jsonLen + 8 + bin.length;
  const glb = new Uint8Array(total);
  const dv = new DataView(glb.buffer);
  glb.set([0x67, 0x6c, 0x54, 0x46], 0);
  dv.setUint32(4, 2, true);
  dv.setUint32(8, total, true);
  dv.setUint32(12, jsonLen, true);
  dv.setUint32(16, 0x4e4f534a, true);
  glb.set(text, 20);
  glb.fill(0x20, 20 + text.length, 20 + jsonLen);
  dv.setUint32(20 + jsonLen, bin.length, true);
  dv.setUint32(24 + jsonLen, 0x004e4942, true);
  glb.set(bin, 28 + jsonLen);
  return glb;
}

/** The files a `.gltf` names beside it — its buffers and pictures not given inline — for the caller to read first. */
export function gltfFiles(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (isGlb(u8)) return [];
  let json;
  try { json = JSON.parse(new TextDecoder().decode(u8)); } catch { return []; }
  return [...(json.buffers || []), ...(json.images || [])].map((x) => x.uri).filter((u) => u && !/^data:/.test(u)).map((u) => decodeURIComponent(u));
}
