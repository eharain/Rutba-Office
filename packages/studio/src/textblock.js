/**
 * Text blocks — a paragraph that wraps inside a box, with lists.
 *
 * WHY THIS IS ONE LAYER TYPE AND NOT TWO. The plan had "text lists" as an easy
 * item and "rich text blocks" as a harder one behind it. They are the same
 * thing: a list is a paragraph whose lines carry a marker and a hanging indent.
 * Building the block and giving it a list MODE is less code than building a
 * list, and it means the two can never drift on wrapping, alignment or leading
 * — which is exactly where two implementations of text always drift.
 *
 * WHY NOT THE EXISTING `text` LAYER. That one is a single-line label with a
 * pill behind it: it does not wrap, and it is anchored at a point rather than
 * bounded by a box. Both are right for what it does (a price chip, a handle)
 * and neither is what a paragraph needs. Widening it would have meant one
 * painter with two layout models inside it.
 *
 * THREE THINGS HERE ARE DELIBERATE AND WORTH THE WORDS:
 *
 *  - TRACKING IS DRAWN PER CHARACTER, not set through `ctx.letterSpacing`.
 *    That property is recent and unevenly supported, so a host that lacks it
 *    would silently lay the same text out differently — and this renderer's
 *    whole claim is that two hosts produce the same frames. Per-character
 *    drawing is deterministic everywhere. It only happens when tracking is
 *    non-zero; the common case still draws a line in one call.
 *
 *  - AUTO-FIT SHRINKS, NEVER GROWS. A block that grew to fill its box would
 *    change size as somebody typed, which is the opposite of what a layout is
 *    for. It searches downward from the size that was asked for and stops at
 *    the first that fits, so the requested size is an upper bound and a block
 *    that already fits is untouched.
 *
 *  - THE CHECK MARKER IS A PATH, NOT A GLYPH. `✓` is not in every font stack
 *    and a fallback would draw something else on one of the two hosts — the
 *    same reason icons.js is path data rather than an icon font. Bullets,
 *    dashes and digits are safe in any font and stay as text.
 *
 * Everything is fractional (fx/fy/fw of the frame, sizeFrac of the width) like
 * every other layer, so one block is correct at 9:16 and at 1:1.
 */

import { registerLayerType } from './index.js';

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const num = (v, fallback) => (Number.isFinite(Number(v)) ? Number(v) : fallback);

const FAMILY = '"Segoe UI", system-ui, Roboto, "Helvetica Neue", Arial, sans-serif';

/** How a block's lines are marked. */
export const LIST_KINDS = ['none', 'bullet', 'dash', 'number', 'check'];

export const LIST_LABELS = {
    none: 'Paragraph', bullet: 'Bulleted', dash: 'Dashed', number: 'Numbered', check: 'Checklist',
};

const fontOf = (layer, sizePx) =>
    `${layer.italic ? 'italic ' : ''}${layer.weight || 500} ${Math.round(sizePx)}px ${layer.family || FAMILY}`;

// ── layout ───────────────────────────────────────────────────────────────────

/**
 * Wrap one paragraph to `maxWidth`, returning whole strings.
 *
 * Deliberately NOT `layoutLines` from index.js: that one returns slices of a
 * shared source string because the caption reveals itself character by
 * character and needs to index back into the original. Nothing here types on,
 * and strings are what the marker logic and the per-character tracking want.
 */
function wrap(ctx, text, maxWidth) {
    const out = [];
    for (const word of String(text).split(/\s+/).filter(Boolean)) {
        if (!out.length) { out.push(word); continue; }
        const merged = `${out[out.length - 1]} ${word}`;
        if (ctx.measureText(merged).width <= maxWidth) out[out.length - 1] = merged;
        else out.push(word);
    }
    // A word wider than the line — a URL, a long product code — is broken by
    // character rather than allowed to run out of the box.
    const fitted = [];
    for (const line of out) {
        if (ctx.measureText(line).width <= maxWidth) { fitted.push(line); continue; }
        let cur = '';
        for (const ch of line) {
            if (cur && ctx.measureText(cur + ch).width > maxWidth) { fitted.push(cur); cur = ch; }
            else cur += ch;
        }
        if (cur) fitted.push(cur);
    }
    return fitted.length ? fitted : [''];
}

const markerFor = (kind, index) => {
    if (kind === 'bullet') return '•';
    if (kind === 'dash') return '–';
    if (kind === 'number') return `${index + 1}.`;
    return null; // 'check' is drawn, 'none' has none
};

/**
 * Lay the block out at a given size.
 *
 * Returns `{ lines, height, sizePx, lead, indent }` where each line carries the
 * marker it should be drawn with — only the FIRST wrapped line of an item does,
 * so a wrapped bullet hangs under its own text rather than restarting.
 */
function layoutAt(ctx, layer, boxW, sizePx) {
    const lead = sizePx * clamp(num(layer.lineHeight, 1.35), 0.8, 3);
    const isList = layer.list && layer.list !== 'none';
    const indent = isList ? sizePx * clamp(num(layer.indent, 1.35), 0.4, 4) : 0;
    const textW = Math.max(4, boxW - indent);

    ctx.font = fontOf(layer, sizePx);

    const lines = [];
    const paragraphs = String(layer.text || '').replace(/\r\n/g, '\n').split('\n');
    let item = 0;
    for (const para of paragraphs) {
        if (!para.trim()) {
            // A blank line is a blank line, not a bulleted nothing.
            lines.push({ text: '', marker: null, first: false });
            continue;
        }
        const wrapped = wrap(ctx, para, textW);
        wrapped.forEach((text, i) => {
            lines.push({
                text,
                marker: i === 0 && isList ? markerFor(layer.list, item) : null,
                // 'check' draws rather than writes, so the painter needs to know
                // this is an item's first line even when there is no glyph.
                first: i === 0 && isList,
            });
        });
        item += 1;
    }

    return { lines, height: lines.length * lead, sizePx, lead, indent, textW };
}

/**
 * The laid-out block, memoised on the layer.
 *
 * A static block would otherwise re-wrap thirty times a second for a video that
 * never changes it — the same memoisation `wordart.js` keeps, for the same
 * reason and with the same shape.
 */
function layoutFor(ctx, plan, layer) {
    const boxW = plan.W * (layer.fw == null ? 0.7 : layer.fw);
    const boxH = layer.fh ? plan.H * layer.fh : null;
    const sig = JSON.stringify([layer.text, layer.sizePx, layer.weight, layer.italic, layer.family,
        layer.lineHeight, layer.list, layer.indent, layer.tracking, layer.autoFit, boxW, boxH]);
    if (layer._tb && layer._tb.sig === sig) return layer._tb.value;

    let laid = layoutAt(ctx, layer, boxW, layer.sizePx);

    /**
     * AUTO-FIT SHRINKS ONLY. Stepping down 6% at a time rather than binary
     * searching: the search would need the layout recomputed per probe anyway,
     * and a linear walk from the requested size means the FIRST size that fits
     * is the largest one that fits — a binary search can land a step below it.
     */
    if (layer.autoFit && boxH && laid.height > boxH) {
        let size = layer.sizePx;
        for (let i = 0; i < 24 && size > 8; i++) {
            size = Math.max(8, size * 0.94);
            const next = layoutAt(ctx, layer, boxW, size);
            if (next.height <= boxH) { laid = next; break; }
            laid = next;
        }
    }

    const value = { ...laid, boxW, boxH };
    layer._tb = { sig, value };
    return value;
}

// ── painting ─────────────────────────────────────────────────────────────────

/** A fill spec → a fillStyle, in the block's own box. Shared shape with shapes.js. */
function fillStyle(ctx, fill, w, h, fallback) {
    if (!fill || !fill.kind || fill.kind === 'solid') return (fill && fill.color) || fallback;
    const stops = Array.isArray(fill.stops) && fill.stops.length ? fill.stops : [[0, '#fff'], [1, '#888']];
    let g;
    if (fill.kind === 'radial') {
        const r = Math.max(w, h) / 2;
        g = ctx.createRadialGradient(0, 0, r * 0.05, 0, 0, r);
    } else {
        const a = (num(fill.angle, 90) * Math.PI) / 180;
        const half = (Math.abs(Math.cos(a)) * w + Math.abs(Math.sin(a)) * h) / 2;
        g = ctx.createLinearGradient(-Math.cos(a) * half, -Math.sin(a) * half, Math.cos(a) * half, Math.sin(a) * half);
    }
    for (const [at, color] of stops) g.addColorStop(clamp(num(at, 0), 0, 1), color);
    return g;
}

/** Draw one line, honouring tracking. See the note at the top on why. */
function drawLine(ctx, text, x, y, tracking, align, width) {
    if (!text) return;
    if (!tracking) { ctx.fillText(text, x, y); return; }

    const chars = [...text];
    const total = chars.reduce((sum, ch) => sum + ctx.measureText(ch).width, 0) + tracking * (chars.length - 1);
    let cx = x;
    if (align === 'center') cx = x - total / 2;
    else if (align === 'right') cx = x - total;
    // Per-character drawing has to place its own glyphs, so the alignment the
    // context would have applied is undone and redone here.
    const saved = ctx.textAlign;
    ctx.textAlign = 'left';
    for (const ch of chars) {
        ctx.fillText(ch, cx, y);
        cx += ctx.measureText(ch).width + tracking;
    }
    ctx.textAlign = saved;
    void width;
}

/** The checklist box and tick, as paths — see the note at the top. */
function drawCheck(ctx, x, y, size, color, done) {
    const s = size * 0.78;
    const top = y + (size - s) / 2;
    ctx.save();
    ctx.strokeStyle = color;
    ctx.lineWidth = Math.max(1, size * 0.08);
    ctx.lineJoin = 'round';
    ctx.beginPath();
    if (typeof ctx.roundRect === 'function') ctx.roundRect(x, top, s, s, s * 0.18);
    else ctx.rect(x, top, s, s);
    ctx.stroke();
    if (done) {
        ctx.beginPath();
        ctx.lineCap = 'round';
        ctx.lineWidth = Math.max(1, size * 0.12);
        ctx.moveTo(x + s * 0.22, top + s * 0.52);
        ctx.lineTo(x + s * 0.43, top + s * 0.73);
        ctx.lineTo(x + s * 0.79, top + s * 0.27);
        ctx.stroke();
    }
    ctx.restore();
}

function paintTextBlock(ctx, plan, layer) {
    const laid = layoutFor(ctx, plan, layer);
    const { lines, lead, indent, sizePx, boxW, boxH } = laid;
    if (!lines.length) return;

    const h = boxH || laid.height;
    ctx.save();
    ctx.translate(layer.x, layer.y);

    // The block is anchored at its CENTRE like every other layer, so the box
    // runs half its width either side — which is what makes rotation about the
    // centre correct with no per-layer correction.
    const left = -boxW / 2;
    const top = -h / 2;

    if (layer.panel) {
        const pad = sizePx * clamp(num(layer.panel.padding, 0.6), 0, 3);
        ctx.fillStyle = layer.panel.color || plan.theme.scrim;
        const r = clamp(num(layer.panel.radius, 0.3), 0, 2) * sizePx;
        ctx.beginPath();
        if (r > 0 && typeof ctx.roundRect === 'function') ctx.roundRect(left - pad, top - pad, boxW + pad * 2, h + pad * 2, r);
        else ctx.rect(left - pad, top - pad, boxW + pad * 2, h + pad * 2);
        ctx.fill();
    }

    ctx.font = fontOf(layer, sizePx);
    ctx.textBaseline = 'top';
    ctx.fillStyle = fillStyle(ctx, layer.fill, boxW, h, plan.theme.text);

    const align = layer.align || 'left';
    ctx.textAlign = align;
    const tracking = num(layer.tracking, 0) * sizePx;

    // Vertical placement inside the box. `middle` and `bottom` only differ from
    // `top` when the box is taller than the text, which is the case auto-fit
    // leaves behind.
    let y = top;
    if (boxH && laid.height < boxH) {
        if (layer.vAlign === 'middle') y = top + (boxH - laid.height) / 2;
        else if (layer.vAlign === 'bottom') y = top + (boxH - laid.height);
    }

    const textLeft = left + indent;
    const anchorX = align === 'center' ? textLeft + laid.textW / 2
        : align === 'right' ? textLeft + laid.textW
            : textLeft;

    const markerColor = layer.markerColor || null;
    lines.forEach((line) => {
        if (line.first) {
            if (layer.list === 'check') {
                drawCheck(ctx, left, y, sizePx, markerColor || ctx.fillStyle, false);
            } else if (line.marker) {
                const savedAlign = ctx.textAlign;
                const savedFill = ctx.fillStyle;
                ctx.textAlign = 'left';
                if (markerColor) ctx.fillStyle = markerColor;
                ctx.fillText(line.marker, left, y);
                ctx.fillStyle = savedFill;
                ctx.textAlign = savedAlign;
            }
        }
        drawLine(ctx, line.text, anchorX, y, tracking, align, laid.textW);
        y += lead;
    });

    ctx.restore();
}

/** Centred fractional box — the convention `fractionalBounds` assumes. */
function textBlockBounds(ctx, plan, layer) {
    const laid = layoutFor(ctx, plan, layer);
    const h = laid.boxH || laid.height;
    return { x: layer.x - laid.boxW / 2, y: layer.y - h / 2, w: laid.boxW, h };
}

// ── the stored form ──────────────────────────────────────────────────────────

function compileTextBlock(patch, cx) {
    const { W, H } = cx;
    const resolved = cx.tokens(String(patch.text == null ? '' : patch.text));
    // A block whose only content was a token with nothing behind it does not
    // draw — the same rule word art and the text layer both follow.
    if (!resolved.text && !patch.keepEmpty) return null;

    const sizeFrac = num(patch.sizeFrac, 0.035);
    const env = cx.envelope(patch.anim);

    return {
        name: patch.name || 'Text',
        text: resolved.text,
        visible: patch.visible !== false && !resolved.missing,
        missingToken: resolved.missing,
        timing: patch.timing || null,
        enter: patch.enter || env.enter,
        exit: patch.exit || env.exit,
        z: patch.z == null ? cx.nextZ() : patch.z,
        x: W * (patch.fx == null ? 0.5 : patch.fx),
        y: H * (patch.fy == null ? 0.5 : patch.fy),
        fx: patch.fx == null ? 0.5 : patch.fx,
        fy: patch.fy == null ? 0.5 : patch.fy,
        fw: patch.fw == null ? 0.7 : patch.fw,
        // No `fh` means "as tall as the words" — the block grows downward from
        // its centre rather than being clipped, and auto-fit has nothing to fit
        // against, which is the right behaviour for a free paragraph.
        fh: patch.fh == null ? null : num(patch.fh, 0.3),
        sizeFrac,
        sizePx: Math.max(8, Math.round(W * sizeFrac)),
        family: patch.family || FAMILY,
        weight: patch.weight || 500,
        italic: !!patch.italic,
        lineHeight: num(patch.lineHeight, 1.35),
        tracking: num(patch.tracking, 0),
        align: patch.align || 'left',
        vAlign: patch.vAlign || 'top',
        list: LIST_KINDS.includes(patch.list) ? patch.list : 'none',
        indent: num(patch.indent, 1.35),
        markerColor: patch.markerColor || null,
        autoFit: patch.autoFit !== false,
        fill: patch.fill === null ? null : (patch.fill || { kind: 'solid', color: cx.color(patch.color, cx.theme.text) }),
        panel: patch.panel || null,
        rot: patch.rot || 0,
        opacity: patch.opacity == null ? 1 : patch.opacity,
        blend: patch.blend || null,
        mask: patch.mask || null,
        anim: patch.anim || 'none',
        keys: cx.keys(patch.keys),
    };
}

export const TEXTBLOCK = registerLayerType('textblock', {
    paint: paintTextBlock,
    compile: compileTextBlock,
    bounds: textBlockBounds,
});
