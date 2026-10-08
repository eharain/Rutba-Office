// A cube as a binary glTF, made here rather than kept as a file: six faces
// of four corners each, the front face red and the rest grey, or — with
// `texture` — the front face carrying a red and white checked picture.
import zlib from 'node:zlib';

/** A tiny RGBA PNG: `pixels` is rows of [r, g, b, a]. */
export function pngOf(pixels) {
  const height = pixels.length;
  const width = pixels[0].length;
  const table = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf) => { let c = 0xffffffff; for (const b of buf) c = table[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, body) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(body.length);
    const typed = Buffer.concat([Buffer.from(type), body]);
    const sum = Buffer.alloc(4); sum.writeUInt32BE(crc(typed));
    return Buffer.concat([len, typed, sum]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0); ihdr.writeUInt32BE(height, 4); ihdr[8] = 8; ihdr[9] = 6;
  const raw = Buffer.concat(pixels.map((row) => Buffer.from([0, ...row.flat()])));
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

export function cubeGlb({ texture = false } = {}) {
  const faces = [
    [[0, 0, 1], [[-1, -1, 1], [1, -1, 1], [1, 1, 1], [-1, 1, 1]]],
    [[0, 0, -1], [[1, -1, -1], [-1, -1, -1], [-1, 1, -1], [1, 1, -1]]],
    [[1, 0, 0], [[1, -1, 1], [1, -1, -1], [1, 1, -1], [1, 1, 1]]],
    [[-1, 0, 0], [[-1, -1, -1], [-1, -1, 1], [-1, 1, 1], [-1, 1, -1]]],
    [[0, 1, 0], [[-1, 1, 1], [1, 1, 1], [1, 1, -1], [-1, 1, -1]]],
    [[0, -1, 0], [[-1, -1, -1], [1, -1, -1], [1, -1, 1], [-1, -1, 1]]],
  ];
  const pos = []; const nor = []; const uv = []; const front = []; const rest = [];
  faces.forEach(([n, corners], f) => {
    const base = pos.length / 3;
    corners.forEach((c, k) => { pos.push(...c); nor.push(...n); uv.push(...[[0, 1], [1, 1], [1, 0], [0, 0]][k]); });
    (f === 0 ? front : rest).push(base, base + 1, base + 2, base, base + 2, base + 3);
  });
  const pad = (b) => Buffer.concat([b, Buffer.alloc((4 - (b.length % 4)) % 4)]);
  const blobs = [
    Buffer.from(new Float32Array(pos).buffer),
    Buffer.from(new Float32Array(nor).buffer),
    Buffer.from(new Float32Array(uv).buffer),
    Buffer.from(new Uint16Array(front).buffer),
    Buffer.from(new Uint16Array(rest).buffer),
  ];
  if (texture) {
    const R = [220, 30, 30, 255];
    const W = [255, 255, 255, 255];
    blobs.push(pngOf([[R, W, R, W], [W, R, W, R], [R, W, R, W], [W, R, W, R]]));
  }
  let off = 0;
  const views = blobs.map((b) => { const v = { buffer: 0, byteOffset: off, byteLength: b.length }; off += pad(b).length; return v; });
  const bin = Buffer.concat(blobs.map(pad));
  const json = {
    asset: { version: '2.0', generator: 'Rutba Office tests' },
    scene: 0,
    scenes: [{ nodes: [0] }],
    nodes: [{ mesh: 0, name: 'Cube' }],
    meshes: [{ primitives: [
      { attributes: { POSITION: 0, NORMAL: 1, TEXCOORD_0: 2 }, indices: 3, material: 0 },
      { attributes: { POSITION: 0, NORMAL: 1, TEXCOORD_0: 2 }, indices: 4, material: 1 },
    ] }],
    materials: [
      texture
        ? { pbrMetallicRoughness: { baseColorTexture: { index: 0 }, baseColorFactor: [1, 1, 1, 1] } }
        : { pbrMetallicRoughness: { baseColorFactor: [0.8, 0.05, 0.05, 1] } },
      { pbrMetallicRoughness: { baseColorFactor: [0.6, 0.6, 0.6, 1] } },
    ],
    ...(texture ? { textures: [{ source: 0, sampler: 0 }], samplers: [{ magFilter: 9728, minFilter: 9728 }], images: [{ bufferView: 5, mimeType: 'image/png' }] } : {}),
    accessors: [
      { bufferView: 0, componentType: 5126, count: 24, type: 'VEC3', min: [-1, -1, -1], max: [1, 1, 1] },
      { bufferView: 1, componentType: 5126, count: 24, type: 'VEC3' },
      { bufferView: 2, componentType: 5126, count: 24, type: 'VEC2' },
      { bufferView: 3, componentType: 5123, count: 6, type: 'SCALAR' },
      { bufferView: 4, componentType: 5123, count: 30, type: 'SCALAR' },
    ],
    bufferViews: views,
    buffers: [{ byteLength: bin.length }],
  };
  const text = Buffer.from(JSON.stringify(json));
  const jsonChunk = Buffer.concat([text, Buffer.alloc((4 - (text.length % 4)) % 4, 0x20)]);
  const header = Buffer.alloc(12);
  header.write('glTF', 0);
  header.writeUInt32LE(2, 4);
  header.writeUInt32LE(12 + 8 + jsonChunk.length + 8 + bin.length, 8);
  const c1 = Buffer.alloc(8); c1.writeUInt32LE(jsonChunk.length, 0); c1.writeUInt32LE(0x4e4f534a, 4);
  const c2 = Buffer.alloc(8); c2.writeUInt32LE(bin.length, 0); c2.writeUInt32LE(0x004e4942, 4);
  return Buffer.concat([header, c1, jsonChunk, c2, bin]);
}
