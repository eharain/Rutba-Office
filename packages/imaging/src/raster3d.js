// 3D models, drawn: a model read by gltf.js turned to a view and painted
// into pixels, with no graphics card — a depth buffer and the triangles
// filled one by one, so the picture is the same on every computer and in a
// check run with no window on screen.
//
// The camera stands in front of the model (glTF's +Z faces the viewer, +Y
// is up) and looks at its centre with a modest perspective; the view turns
// the model about its upright axis (`yaw`), then tips it towards or away
// from the viewer (`pitch`), then about the line of sight (`roll`), in
// degrees. A key light from the upper left and a softer fill light the
// faces, over an ambient share so no side goes black. Colours are worked
// in linear light and written as sRGB; a base-colour texture is sampled
// with its own coordinates, smoothly, wrapping as its sampler says. Each
// pixel is the average of a 2×2 grid of samples, so edges are smooth.
//
// The picture is fitted to the model as turned: `fit` picks the width and
// height for a given longest side from the model's outline at that view,
// and `renderModel` fills the frame it is given, the model centred in it.

const toLinear = (c) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const toSrgb = (c) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);
const LINEAR = new Float32Array(256).map((_, i) => toLinear(i / 255));

/** The rotation of a view as a 3×3 matrix, row-major: yaw about Y, then pitch about X, then roll about Z. */
export function viewMatrix({ yaw = 0, pitch = 0, roll = 0 } = {}) {
  const r = Math.PI / 180;
  const cy = Math.cos(yaw * r), sy = Math.sin(yaw * r);
  const cp = Math.cos(pitch * r), sp = Math.sin(pitch * r);
  const cr = Math.cos(roll * r), sr = Math.sin(roll * r);
  const Y = [cy, 0, sy, 0, 1, 0, -sy, 0, cy];
  const X = [1, 0, 0, 0, cp, -sp, 0, sp, cp];
  const Z = [cr, -sr, 0, sr, cr, 0, 0, 0, 1];
  const mul = (a, b) => [0, 1, 2].flatMap((i) => [0, 1, 2].map((j) => a[i * 3] * b[j] + a[i * 3 + 1] * b[3 + j] + a[i * 3 + 2] * b[6 + j]));
  return mul(Z, mul(X, Y));
}

const FOV = 30;

/** Each vertex of the model as the camera sees it: x and y on the picture plane (units of the model), and depth. */
function project(model, view) {
  const R = viewMatrix(view);
  const { centre, radius } = model.bounds;
  const distance = radius / Math.tan((FOV / 2) * Math.PI / 180) * 1.05;
  const focal = distance;
  return model.meshes.map((m) => {
    const n = m.positions.length / 3;
    const sx = new Float32Array(n), sy = new Float32Array(n), sz = new Float32Array(n);
    for (let i = 0; i < n; i++) {
      const x = m.positions[i * 3] - centre[0], y = m.positions[i * 3 + 1] - centre[1], z = m.positions[i * 3 + 2] - centre[2];
      const rx = R[0] * x + R[1] * y + R[2] * z;
      const ry = R[3] * x + R[4] * y + R[5] * z;
      const rz = R[6] * x + R[7] * y + R[8] * z;
      const depth = distance - rz; // distance from the camera along its axis
      const k = focal / Math.max(1e-6, depth);
      sx[i] = rx * k; sy[i] = ry * k; sz[i] = depth;
    }
    return { sx, sy, sz, R };
  });
}

/** The outline of the model at a view, on the picture plane: { minX, maxX, minY, maxY }. */
export function outline(model, view = {}) {
  let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
  for (const p of project(model, view)) {
    for (let i = 0; i < p.sx.length; i++) {
      if (p.sx[i] < minX) minX = p.sx[i]; if (p.sx[i] > maxX) maxX = p.sx[i];
      if (p.sy[i] < minY) minY = p.sy[i]; if (p.sy[i] > maxY) maxY = p.sy[i];
    }
  }
  return { minX, maxX, minY, maxY };
}

/** A picture's size for the model at a view, its longest side `longest` pixels. */
export function fit(model, view = {}, longest = 800) {
  const o = outline(model, view);
  const w = Math.max(1e-6, o.maxX - o.minX);
  const h = Math.max(1e-6, o.maxY - o.minY);
  return w >= h ? { width: longest, height: Math.max(1, Math.round((longest * h) / w)) } : { width: Math.max(1, Math.round((longest * w) / h)), height: longest };
}

function sampler(texture, wrap) {
  if (!texture) return null;
  const { width, height, data } = texture;
  const wrapOne = (v, mode, n) => {
    if (mode === 33071) return Math.min(n - 1, Math.max(0, v)); // clamp
    if (mode === 33648) { const p = ((v % (2 * n)) + 2 * n) % (2 * n); return p < n ? p : 2 * n - 1 - p; } // mirrored
    return ((v % n) + n) % n;
  };
  return (u, v, out) => {
    const x = u * width - 0.5, y = v * height - 0.5;
    const x0 = Math.floor(x), y0 = Math.floor(y);
    const fx = x - x0, fy = y - y0;
    const xa = wrapOne(x0, wrap[0], width), xb = wrapOne(x0 + 1, wrap[0], width);
    const ya = wrapOne(y0, wrap[1], height), yb = wrapOne(y0 + 1, wrap[1], height);
    for (let k = 0; k < 4; k++) {
      const p00 = data[(ya * width + xa) * 4 + k], p10 = data[(ya * width + xb) * 4 + k];
      const p01 = data[(yb * width + xa) * 4 + k], p11 = data[(yb * width + xb) * 4 + k];
      const top = p00 + (p10 - p00) * fx, bottom = p01 + (p11 - p01) * fx;
      const c = top + (bottom - top) * fy;
      out[k] = k < 3 ? LINEAR[Math.max(0, Math.min(255, Math.round(c)))] : c / 255;
    }
    return out;
  };
}

/**
 * The model drawn at `view` into a `width` × `height` RGBA picture,
 * transparent round it. `textures` maps an image index to its decoded
 * pixels `{ width, height, data }`; a texture not given is left out and the
 * base colour drawn alone.
 */
export function renderModel(model, { width, height, view = {}, textures = new Map(), samples = 2, margin = 0.02 } = {}) {
  const W = Math.max(1, Math.round(width)) * samples;
  const H = Math.max(1, Math.round(height)) * samples;
  const depth = new Float32Array(W * H).fill(Infinity);
  const colour = new Float32Array(W * H * 4);
  const projected = project(model, view);
  const o = outline(model, view);
  const spanX = Math.max(1e-6, o.maxX - o.minX), spanY = Math.max(1e-6, o.maxY - o.minY);
  const scale = Math.min((W * (1 - 2 * margin)) / spanX, (H * (1 - 2 * margin)) / spanY);
  const cx = (o.minX + o.maxX) / 2, cy = (o.minY + o.maxY) / 2;
  const toPx = (x) => W / 2 + (x - cx) * scale;
  const toPy = (y) => H / 2 - (y - cy) * scale;

  // Lights, in the camera's frame: towards the light from the surface.
  const norm = (v) => { const l = Math.hypot(...v) || 1; return v.map((c) => c / l); };
  const key = norm([-0.45, 0.6, 0.65]);
  const fill = norm([0.6, -0.1, 0.8]);
  const half = norm([key[0], key[1], key[2] + 1]);
  const AMBIENT = 0.32, KEY = 0.62, FILL = 0.22;
  const tex = [0, 0, 0, 1];

  model.meshes.forEach((m, mi) => {
    const { sx, sy, sz, R } = projected[mi];
    const mat = m.material;
    const sample = m.uvs && mat.image != null ? sampler(textures.get(mat.image), mat.wrap) : null;
    const base = mat.color;
    const idx = m.indices;
    for (let t = 0; t + 2 < idx.length; t += 3) {
      const a = idx[t], b = idx[t + 1], c = idx[t + 2];
      const ax = toPx(sx[a]), ay = toPy(sy[a]), bx = toPx(sx[b]), by = toPy(sy[b]), qx = toPx(sx[c]), qy = toPy(sy[c]);
      const area = (bx - ax) * (qy - ay) - (by - ay) * (qx - ax);
      if (!area || !Number.isFinite(area)) continue;
      // Triangles wound clockwise on screen face away; a double-sided material draws them too.
      const back = area > 0;
      if (back && !mat.doubleSided && m.normals) continue;
      // The face's own normal, for a mesh that has none, in the camera's frame.
      let flat = null;
      if (!m.normals) {
        const p = m.positions;
        const ux = p[b * 3] - p[a * 3], uy = p[b * 3 + 1] - p[a * 3 + 1], uz = p[b * 3 + 2] - p[a * 3 + 2];
        const vx = p[c * 3] - p[a * 3], vy = p[c * 3 + 1] - p[a * 3 + 1], vz = p[c * 3 + 2] - p[a * 3 + 2];
        const n = [uy * vz - uz * vy, uz * vx - ux * vz, ux * vy - uy * vx];
        flat = norm([R[0] * n[0] + R[1] * n[1] + R[2] * n[2], R[3] * n[0] + R[4] * n[1] + R[5] * n[2], R[6] * n[0] + R[7] * n[1] + R[8] * n[2]]);
        if (flat[2] < 0) flat = flat.map((v) => -v);
      }
      const minX = Math.max(0, Math.floor(Math.min(ax, bx, qx))), maxX = Math.min(W - 1, Math.ceil(Math.max(ax, bx, qx)));
      const minY = Math.max(0, Math.floor(Math.min(ay, by, qy))), maxY = Math.min(H - 1, Math.ceil(Math.max(ay, by, qy)));
      if (minX > maxX || minY > maxY) continue;
      // Perspective-correct interpolation: attributes over depth, then divided back.
      const iza = 1 / sz[a], izb = 1 / sz[b], izc = 1 / sz[c];
      for (let y = minY; y <= maxY; y++) {
        const py = y + 0.5;
        for (let x = minX; x <= maxX; x++) {
          const px = x + 0.5;
          let w0 = (bx - px) * (qy - py) - (by - py) * (qx - px);
          let w1 = (qx - px) * (ay - py) - (qy - py) * (ax - px);
          let w2 = (ax - px) * (by - py) - (ay - py) * (bx - px);
          if (area < 0 ? w0 > 0 || w1 > 0 || w2 > 0 : w0 < 0 || w1 < 0 || w2 < 0) continue;
          w0 /= area; w1 /= area; w2 /= area;
          const iz = w0 * iza + w1 * izb + w2 * izc;
          const z = 1 / iz;
          const at = y * W + x;
          if (z >= depth[at]) continue;
          const pa = (w0 * iza) / iz, pb = (w1 * izb) / iz, pc = (w2 * izc) / iz;
          // The surface's colour: base colour, vertex colour, texture.
          let r = base[0], g = base[1], bl = base[2], al = base[3] ?? 1;
          if (m.colors) {
            r *= pa * m.colors[a * 4] + pb * m.colors[b * 4] + pc * m.colors[c * 4];
            g *= pa * m.colors[a * 4 + 1] + pb * m.colors[b * 4 + 1] + pc * m.colors[c * 4 + 1];
            bl *= pa * m.colors[a * 4 + 2] + pb * m.colors[b * 4 + 2] + pc * m.colors[c * 4 + 2];
            al *= pa * m.colors[a * 4 + 3] + pb * m.colors[b * 4 + 3] + pc * m.colors[c * 4 + 3];
          }
          if (sample) {
            const u = pa * m.uvs[a * 2] + pb * m.uvs[b * 2] + pc * m.uvs[c * 2];
            const v = pa * m.uvs[a * 2 + 1] + pb * m.uvs[b * 2 + 1] + pc * m.uvs[c * 2 + 1];
            sample(u, v, tex);
            r *= tex[0]; g *= tex[1]; bl *= tex[2]; al *= tex[3];
          }
          if (mat.alphaMode === 'MASK' && al < mat.alphaCutoff) continue;
          // The light on it.
          let shade = 1;
          if (!mat.unlit) {
            let nx, ny, nz;
            if (m.normals) {
              const N = m.normals;
              const wx = pa * N[a * 3] + pb * N[b * 3] + pc * N[c * 3];
              const wy = pa * N[a * 3 + 1] + pb * N[b * 3 + 1] + pc * N[c * 3 + 1];
              const wz = pa * N[a * 3 + 2] + pb * N[b * 3 + 2] + pc * N[c * 3 + 2];
              nx = R[0] * wx + R[1] * wy + R[2] * wz;
              ny = R[3] * wx + R[4] * wy + R[5] * wz;
              nz = R[6] * wx + R[7] * wy + R[8] * wz;
              const l = Math.hypot(nx, ny, nz) || 1;
              nx /= l; ny /= l; nz /= l;
              if (back) { nx = -nx; ny = -ny; nz = -nz; }
            } else {
              [nx, ny, nz] = flat;
            }
            const k = Math.max(0, nx * key[0] + ny * key[1] + nz * key[2]);
            const f = Math.max(0, nx * fill[0] + ny * fill[1] + nz * fill[2]);
            shade = AMBIENT + KEY * k + FILL * f;
            // A little shine on a smooth surface, towards the key light.
            const spec = (1 - (mat.roughness ?? 1)) * 0.35 * Math.max(0, nx * half[0] + ny * half[1] + nz * half[2]) ** 24;
            r = r * shade + spec; g = g * shade + spec; bl = bl * shade + spec;
          }
          r += mat.emissive[0]; g += mat.emissive[1]; bl += mat.emissive[2];
          depth[at] = z;
          colour[at * 4] = r; colour[at * 4 + 1] = g; colour[at * 4 + 2] = bl;
          colour[at * 4 + 3] = mat.alphaMode === 'BLEND' ? Math.max(0.05, Math.min(1, al)) : 1;
        }
      }
    }
  });

  // Down to the picture's own size, each pixel the average of its samples, as sRGB.
  const w = Math.round(width), h = Math.round(height);
  const out = new Uint8ClampedArray(w * h * 4);
  const n = samples * samples;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      let r = 0, g = 0, b = 0, a = 0;
      for (let j = 0; j < samples; j++) {
        for (let i = 0; i < samples; i++) {
          const at = ((y * samples + j) * W + (x * samples + i)) * 4;
          const al = colour[at + 3];
          r += colour[at] * al; g += colour[at + 1] * al; b += colour[at + 2] * al; a += al;
        }
      }
      const o4 = (y * w + x) * 4;
      if (a > 0) {
        out[o4] = Math.round(toSrgb(Math.min(1, r / a)) * 255);
        out[o4 + 1] = Math.round(toSrgb(Math.min(1, g / a)) * 255);
        out[o4 + 2] = Math.round(toSrgb(Math.min(1, b / a)) * 255);
        out[o4 + 3] = Math.round((a / n) * 255);
      }
    }
  }
  return { width: w, height: h, data: out };
}

/** PowerPoint's 3D Model Views, as turns of the model: [name, label, view]. */
export const MODEL_VIEWS = [
  ['front', 'Front', { yaw: 0, pitch: 0 }],
  ['back', 'Back', { yaw: 180, pitch: 0 }],
  ['left', 'Left', { yaw: 90, pitch: 0 }],
  ['right', 'Right', { yaw: -90, pitch: 0 }],
  ['top', 'Top', { yaw: 0, pitch: 90 }],
  ['bottom', 'Bottom', { yaw: 0, pitch: -90 }],
  ['aboveFrontLeft', 'Above Front Left', { yaw: 35, pitch: 25 }],
  ['aboveFrontRight', 'Above Front Right', { yaw: -35, pitch: 25 }],
  ['belowFrontLeft', 'Below Front Left', { yaw: 35, pitch: -25 }],
  ['belowFrontRight', 'Below Front Right', { yaw: -35, pitch: -25 }],
];
