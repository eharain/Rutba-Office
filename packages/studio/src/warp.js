/**
 * A picture taken through four corners (the Image app's advanced editor,
 * docs/STUDIO.md 7.2, I14): free transform with corner handles (scale, skew,
 * distort) and perspective correction are both a projective map from the
 * picture's rectangle to a four-cornered shape on the canvas.
 *
 * A layer that has been transformed carries `corners`: its picture's four
 * corners, top left, top right, bottom right, bottom left, as fractions of the
 * canvas, like every other position in a studio document. The painter maps
 * the picture onto them with `homography` and `warpPixels`; nothing is
 * resampled into the project, so a transform can be changed or taken off
 * again with the picture as it was.
 *
 * Perspective correction is the same map the other way round: four points
 * set on something that should be a rectangle (a sign, a page, a window seen
 * at an angle) go to the rectangle `rectifyTarget` works out for them, and
 * the layer's corners go through that map too.
 *
 * Pure arithmetic: no DOM, no canvas. Runs in a page, a worker and Node.
 */

const EPS = 1e-12;

/**
 * The 3 × 3 projective map (row major, the last entry 1) taking the four
 * points `src` to the four points `dst`: [[x, y], …] each, in order.
 * Answers null for four points that do not make a shape (three in a line).
 */
export function homography(src, dst) {
    if (!Array.isArray(src) || !Array.isArray(dst) || src.length !== 4 || dst.length !== 4) return null;
    // Eight equations in h0..h7:
    //   u = (h0 x + h1 y + h2) / (h6 x + h7 y + 1)
    //   v = (h3 x + h4 y + h5) / (h6 x + h7 y + 1)
    const A = [];
    const b = [];
    for (let i = 0; i < 4; i++) {
        const [x, y] = src[i];
        const [u, v] = dst[i];
        A.push([x, y, 1, 0, 0, 0, -u * x, -u * y]);
        b.push(u);
        A.push([0, 0, 0, x, y, 1, -v * x, -v * y]);
        b.push(v);
    }
    const h = solve(A, b);
    if (!h) return null;
    return [h[0], h[1], h[2], h[3], h[4], h[5], h[6], h[7], 1];
}

/** Gaussian elimination with partial pivoting; null when the system has no single answer. */
export function solve(A, b) {
    const n = b.length;
    const M = A.map((row, i) => [...row, b[i]]);
    for (let c = 0; c < n; c++) {
        let pivot = c;
        for (let r = c + 1; r < n; r++) if (Math.abs(M[r][c]) > Math.abs(M[pivot][c])) pivot = r;
        if (Math.abs(M[pivot][c]) < EPS) return null;
        if (pivot !== c) [M[pivot], M[c]] = [M[c], M[pivot]];
        for (let r = 0; r < n; r++) {
            if (r === c) continue;
            const f = M[r][c] / M[c][c];
            if (!f) continue;
            for (let k = c; k <= n; k++) M[r][k] -= f * M[c][k];
        }
    }
    return M.map((row, i) => row[n] / row[i]);
}

/** A point through a projective map: [x', y']. */
export function applyHomography(H, x, y) {
    const w = H[6] * x + H[7] * y + H[8];
    const d = Math.abs(w) < EPS ? EPS : w;
    return [(H[0] * x + H[1] * y + H[2]) / d, (H[3] * x + H[4] * y + H[5]) / d];
}

/** The map that undoes `H`, scaled so its last entry is 1; null when H has none. */
export function invertHomography(H) {
    const [a, b, c, d, e, f, g, h, i] = H;
    const A = e * i - f * h;
    const B = -(d * i - f * g);
    const C = d * h - e * g;
    const det = a * A + b * B + c * C;
    if (Math.abs(det) < EPS) return null;
    const inv = [
        A, -(b * i - c * h), b * f - c * e,
        B, a * i - c * g, -(a * f - c * d),
        C, -(a * h - b * g), a * e - b * d,
    ].map((v) => v / det);
    const k = Math.abs(inv[8]) < EPS ? 1 : inv[8];
    return inv.map((v) => v / k);
}

/** The four corners of a `width` × `height` picture, in the order a quad keeps them. */
export const rectCorners = (width, height, x = 0, y = 0) => [[x, y], [x + width, y], [x + width, y + height], [x, y + height]];

/** The smallest axis-aligned box holding four points: { x, y, w, h }. */
export function quadBounds(quad) {
    const xs = quad.map((p) => p[0]);
    const ys = quad.map((p) => p[1]);
    const x = Math.min(...xs);
    const y = Math.min(...ys);
    return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
}

/** Whether four points, in order, make a shape with no crossing edges and no corner turned in. */
export function isConvexQuad(quad) {
    let sign = 0;
    for (let i = 0; i < 4; i++) {
        const [ax, ay] = quad[i];
        const [bx, by] = quad[(i + 1) % 4];
        const [cx, cy] = quad[(i + 2) % 4];
        const cross = (bx - ax) * (cy - by) - (by - ay) * (cx - bx);
        if (Math.abs(cross) < 1e-9) return false;
        const s = Math.sign(cross);
        if (sign && s !== sign) return false;
        sign = s;
    }
    return true;
}

/**
 * The rectangle four points set on a skewed rectangle are straightened to:
 * as wide as its top and bottom edges on average, as tall as its sides on
 * average, centred where the four points are. Answers its corners in the
 * same order (top left first).
 */
export function rectifyTarget(quad) {
    const len = (a, b) => Math.hypot(b[0] - a[0], b[1] - a[1]);
    const w = (len(quad[0], quad[1]) + len(quad[3], quad[2])) / 2;
    const h = (len(quad[0], quad[3]) + len(quad[1], quad[2])) / 2;
    const cx = quad.reduce((s, p) => s + p[0], 0) / 4;
    const cy = quad.reduce((s, p) => s + p[1], 0) / 4;
    return rectCorners(w, h, cx - w / 2, cy - h / 2);
}

/**
 * Straighten: the map taking the four points set on the picture to their
 * rectangle, and every point of `corners` (the layer's own four) through it.
 * Answers { corners, rect, H } or null when the points make no shape.
 */
export function straighten(points, corners) {
    if (!isConvexQuad(points)) return null;
    const rect = rectifyTarget(points);
    const H = homography(points, rect);
    if (!H) return null;
    return { corners: corners.map(([x, y]) => applyHomography(H, x, y)), rect, H };
}

/**
 * A layer's four corners in canvas pixels from its box `{ x, y, w, h }` and a
 * turn of `rot` degrees about the box's centre.
 */
export function boxCorners(b, rot = 0) {
    const pts = rectCorners(b.w, b.h, b.x, b.y);
    if (!rot) return pts;
    const cx = b.x + b.w / 2;
    const cy = b.y + b.h / 2;
    const a = (rot * Math.PI) / 180;
    const cos = Math.cos(a);
    const sin = Math.sin(a);
    return pts.map(([x, y]) => [cx + (x - cx) * cos - (y - cy) * sin, cy + (x - cx) * sin + (y - cy) * cos]);
}

/**
 * Free transform: the four corners after the handle at `index` (0 top left,
 * 1 top right, 2 bottom right, 3 bottom left) is dragged from where it was in
 * `start` to `point`.
 *
 * - `scale`: the box the corners make grows or shrinks from the opposite
 *   corner, every corner with it; `keep` holds the shape.
 * - `skew`: the corner and its neighbour on the edge the drag runs along
 *   slide together, the other edge stays put (a slant, not a stretch).
 * - `distort`: the corner goes where it is dragged, alone.
 * - `perspective`: the corner and its mirror on the same edge move in
 *   opposite directions along that edge, which is the "seen at an angle" look.
 */
export function dragCorner(start, index, point, mode = 'scale', { keep = false } = {}) {
    const q = start.map((p) => [...p]);
    const [sx, sy] = start[index];
    let dx = point[0] - sx;
    let dy = point[1] - sy;
    if (mode === 'distort') {
        q[index] = [point[0], point[1]];
        return q;
    }
    if (mode === 'skew' || mode === 'perspective') {
        // Along the top or bottom edge for a mostly sideways drag, the left or right side otherwise.
        const across = Math.abs(dx) >= Math.abs(dy);
        const partner = across ? [1, 0, 3, 2][index] : [3, 2, 1, 0][index];
        if (across) dy = 0; else dx = 0;
        q[index] = [sx + dx, sy + dy];
        const [px, py] = start[partner];
        q[partner] = mode === 'skew' ? [px + dx, py + dy] : [px - dx, py - dy];
        return q;
    }
    // Scale from the opposite corner, in the quad's own frame (it may be turned).
    const far = start[(index + 2) % 4];
    const ux = [start[1][0] - start[0][0], start[1][1] - start[0][1]];
    const uy = [start[3][0] - start[0][0], start[3][1] - start[0][1]];
    const lx = Math.hypot(ux[0], ux[1]) || 1;
    const ly = Math.hypot(uy[0], uy[1]) || 1;
    const ex = [ux[0] / lx, ux[1] / lx];
    const ey = [uy[0] / ly, uy[1] / ly];
    const was = [sx - far[0], sy - far[1]];
    const now = [point[0] - far[0], point[1] - far[1]];
    const dot = (a, b) => a[0] * b[0] + a[1] * b[1];
    let kx = Math.abs(dot(was, ex)) < 1e-9 ? 1 : dot(now, ex) / dot(was, ex);
    let ky = Math.abs(dot(was, ey)) < 1e-9 ? 1 : dot(now, ey) / dot(was, ey);
    if (keep) {
        const k = Math.abs(kx) > Math.abs(ky) ? kx : ky;
        kx = k;
        ky = k;
    }
    const min = 0.02;
    if (Math.abs(kx) < min) kx = kx < 0 ? -min : min;
    if (Math.abs(ky) < min) ky = ky < 0 ? -min : min;
    return start.map(([x, y]) => {
        const v = [x - far[0], y - far[1]];
        const a = dot(v, ex) * kx;
        const b = dot(v, ey) * ky;
        return [far[0] + ex[0] * a + ey[0] * b, far[1] + ex[1] * a + ey[1] * b];
    });
}

/**
 * Warp a picture onto four corners. `src` is { data, width, height } (RGBA);
 * `corners` the picture's four corners in the output's pixels; `out` is
 * { width, height } of the surface written. Bilinear, in premultiplied alpha
 * so a clear edge does not darken; what falls outside the picture is clear.
 * Answers a new Uint8ClampedArray of out.width × out.height × 4.
 */
export function warpPixels(src, corners, out) {
    const { width: sw, height: sh, data } = src;
    const W = Math.max(1, Math.round(out.width));
    const H = Math.max(1, Math.round(out.height));
    const result = new Uint8ClampedArray(W * H * 4);
    const toSrc = homography(corners, rectCorners(sw, sh));
    if (!toSrc) return result;
    const b = quadBounds(corners);
    const x0 = Math.max(0, Math.floor(b.x));
    const y0 = Math.max(0, Math.floor(b.y));
    const x1 = Math.min(W, Math.ceil(b.x + b.w));
    const y1 = Math.min(H, Math.ceil(b.y + b.h));
    const px = (x, y, c) => {
        if (x < 0 || y < 0 || x >= sw || y >= sh) return 0;
        return data[(y * sw + x) * 4 + c];
    };
    for (let y = y0; y < y1; y++) {
        for (let x = x0; x < x1; x++) {
            const [u, v] = applyHomography(toSrc, x + 0.5, y + 0.5);
            const fx = u - 0.5;
            const fy = v - 0.5;
            if (fx < -1 || fy < -1 || fx > sw || fy > sh) continue;
            const ix = Math.floor(fx);
            const iy = Math.floor(fy);
            const ax = fx - ix;
            const ay = fy - iy;
            const w00 = (1 - ax) * (1 - ay);
            const w10 = ax * (1 - ay);
            const w01 = (1 - ax) * ay;
            const w11 = ax * ay;
            const a00 = px(ix, iy, 3);
            const a10 = px(ix + 1, iy, 3);
            const a01 = px(ix, iy + 1, 3);
            const a11 = px(ix + 1, iy + 1, 3);
            const alpha = a00 * w00 + a10 * w10 + a01 * w01 + a11 * w11;
            if (alpha <= 0) continue;
            const o = (y * W + x) * 4;
            for (let c = 0; c < 3; c++) {
                const pre = px(ix, iy, c) * a00 * w00 + px(ix + 1, iy, c) * a10 * w10 + px(ix, iy + 1, c) * a01 * w01 + px(ix + 1, iy + 1, c) * a11 * w11;
                result[o + c] = pre / alpha;
            }
            result[o + 3] = alpha;
        }
    }
    return result;
}
