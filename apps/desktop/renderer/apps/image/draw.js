// Drawing a plan onto a canvas: the picture through the plan's matrix and
// filter, then the marks on top. The preview and the export both come here,
// the preview at the window's scale and the export at full size, so what is
// saved is what was seen.

/** The picture alone, at `scale` of the result's size; the canvas is sized to it. */
export function drawBase(canvas, image, plan, scale = 1, { readBack = false } = {}) {
  const width = Math.max(1, Math.round(plan.target.width * scale));
  const height = Math.max(1, Math.round(plan.target.height * scale));
  if (canvas.width !== width) canvas.width = width;
  if (canvas.height !== height) canvas.height = height;
  // A canvas read back on every frame keeps its pixels where reading is cheap.
  const ctx = canvas.getContext('2d', readBack ? { willReadFrequently: true } : undefined);
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, width, height);
  ctx.save();
  ctx.filter = plan.filter || 'none';
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  const m = plan.matrix;
  ctx.setTransform(m[0] * scale, m[1] * scale, m[2] * scale, m[3] * scale, m[4] * scale, m[5] * scale);
  ctx.drawImage(image, 0, 0, plan.source.width, plan.source.height);
  ctx.restore();
  return ctx;
}

/** The arrow's head: two short strokes back from the point. */
function arrowHead(ctx, from, to, size) {
  const angle = Math.atan2(to.y - from.y, to.x - from.x);
  const spread = Math.PI / 7;
  ctx.beginPath();
  ctx.moveTo(to.x - size * Math.cos(angle - spread), to.y - size * Math.sin(angle - spread));
  ctx.lineTo(to.x, to.y);
  ctx.lineTo(to.x - size * Math.cos(angle + spread), to.y - size * Math.sin(angle + spread));
  ctx.closePath();
  ctx.fill();
}

/** A pen stroke through its points, smoothed through the midpoints. */
function penPath(ctx, points) {
  if (!points.length) return;
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  if (points.length === 1) {
    ctx.lineTo(points[0].x + 0.01, points[0].y);
  } else {
    for (let i = 1; i < points.length - 1; i++) {
      const mx = (points[i].x + points[i + 1].x) / 2;
      const my = (points[i].y + points[i + 1].y) / 2;
      ctx.quadraticCurveTo(points[i].x, points[i].y, mx, my);
    }
    const last = points[points.length - 1];
    ctx.lineTo(last.x, last.y);
  }
  ctx.stroke();
}

export const MARK_FONT = '"Segoe UI", system-ui, sans-serif';

/** The marks, in the source's pixels, drawn through the plan at `scale`. */
export function drawMarks(ctx, plan, marks, scale = 1) {
  const m = plan.matrix;
  for (const o of marks || []) {
    ctx.save();
    ctx.setTransform(m[0] * scale, m[1] * scale, m[2] * scale, m[3] * scale, m[4] * scale, m[5] * scale);
    ctx.strokeStyle = o.colour || '#e0393e';
    ctx.fillStyle = o.colour || '#e0393e';
    ctx.lineWidth = Math.max(0.1, o.width || 1);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    if (o.shape === 'rect' || o.shape === 'ellipse') {
      const f = o.frame || [1, 0, 0, 1];
      ctx.translate(o.cx, o.cy);
      ctx.transform(f[0], f[1], f[2], f[3], 0, 0);
      if (o.shape === 'rect') ctx.strokeRect(-o.w / 2, -o.h / 2, o.w, o.h);
      else {
        ctx.beginPath();
        ctx.ellipse(0, 0, Math.max(0.1, o.w / 2), Math.max(0.1, o.h / 2), 0, 0, Math.PI * 2);
        ctx.stroke();
      }
    } else if (o.shape === 'arrow' && o.points?.length === 2) {
      const [a, b] = o.points;
      const head = o.head || o.width * 4;
      const len = Math.hypot(b.x - a.x, b.y - a.y) || 1;
      // The shaft stops short of the point, so the head's tip is sharp.
      const back = Math.min(len, head * 0.6);
      ctx.beginPath();
      ctx.moveTo(a.x, a.y);
      ctx.lineTo(b.x - ((b.x - a.x) / len) * back, b.y - ((b.y - a.y) / len) * back);
      ctx.stroke();
      arrowHead(ctx, a, b, head);
    } else if (o.shape === 'pen') {
      penPath(ctx, o.points || []);
    } else if (o.shape === 'text' && o.text) {
      const f = o.frame || [1, 0, 0, 1];
      ctx.translate(o.x, o.y);
      ctx.transform(f[0], f[1], f[2], f[3], 0, 0);
      ctx.font = `600 ${o.size}px ${MARK_FONT}`;
      ctx.textBaseline = 'top';
      String(o.text).split('\n').forEach((line, i) => ctx.fillText(line, 0, i * o.size * 1.25));
    }
    ctx.restore();
  }
}
