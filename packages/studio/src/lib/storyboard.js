/**
 * A first draft of a video, composed from what the post already knows.
 *
 * This is not generation: every piece it emits is an ordinary layer patch or
 * an ordinary option, so the operator edits the result with the same controls
 * they would have used to build it by hand. What it saves is the blank page —
 * the twenty seconds of deciding where the price chip goes and when the QR
 * should appear, on every post, forever.
 *
 * Pure on purpose: it takes numbers and a context and returns patches, so
 * pages/dev/timeline-fixture.js can hold its rules without a post, a product
 * or a session.
 *
 * Two rules it will not break:
 *   1. Everything it adds is prefixed `sb-`, and drafting again REPLACES that
 *      set rather than stacking a second copy on top of the first.
 *   2. Nothing lands where a platform's own UI sits. The bottom band and the
 *      right rail are covered on TikTok and Reels, and a price chip under the
 *      caption is a price nobody reads.
 */

export const SB_PREFIX = 'sb-';

// The bands D3's guides draw, as fractions. Anything placed must keep its
// ANCHOR inside these — a centred chip still needs room for half its width.
export const SAFE = { top: 0.05, bottom: 0.82, left: 0.05, right: 0.88 };

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

/**
 * @param {object} input
 * @param {number} input.duration        the video's length in seconds
 * @param {object} [input.context]       {price, was, discount, product, url}
 * @param {string[]} [input.captionSegments] the caption already split
 * @param {boolean} [input.hasTitle]
 * @param {object} [input.options]       the CURRENT options; only empty ones fill
 * @returns {{patches: object[], options: object, notes: string[]}}
 */
export function draftStoryboard(input = {}) {
    const duration = Number(input.duration) > 0 ? Number(input.duration) : 10;
    const ctx = input.context || {};
    const segs = Array.isArray(input.captionSegments) ? input.captionSegments : [];
    const opts = input.options || {};
    const patches = [];
    const notes = [];
    const t = (a, b) => ({ start: +clamp(a, 0, duration - 0.4).toFixed(3), end: +clamp(b, 0.4, duration).toFixed(3) });

    // The hook: a discount is the reason to keep watching, so it lands early
    // and leaves before the viewer has decided anything else.
    if (ctx.discount) {
        patches.push({
            id: `${SB_PREFIX}discount`, type: 'text', name: 'Discount hook',
            text: '{discount} OFF', pill: 'accent', color: '#141118', weight: 800,
            sizeFrac: 0.052, fx: 0.5, fy: SAFE.top + 0.03, align: 'center',
            timing: t(0.3, Math.min(duration, 3.2)),
            enter: { kind: 'zoom', seconds: 0.35 }, exit: { kind: 'fade', seconds: 0.4 },
        });
        notes.push('a discount hook in the first three seconds');
    }

    // The price stays up once the product has been seen, not before it: a
    // number over the first frame is a number without a subject.
    if (ctx.price) {
        patches.push({
            id: `${SB_PREFIX}price`, type: 'text', name: 'Price chip',
            text: '{price}', pill: 'accent', color: '#141118', weight: 800,
            sizeFrac: 0.055, fx: 0.5, fy: ctx.discount ? SAFE.top + 0.03 : SAFE.top + 0.02, align: 'center',
            timing: t(ctx.discount ? 3.4 : 1.2, duration),
            enter: { kind: 'fade', seconds: 0.4 }, exit: { kind: 'none', seconds: 0 },
        });
        notes.push(ctx.discount ? 'the price taking over when the hook leaves' : 'a price chip after the opening');
    }

    // The QR is the call to action, so it belongs in the last stretch — when
    // someone is still watching, they are deciding.
    if (ctx.url) {
        patches.push({
            id: `${SB_PREFIX}qr`, type: 'qr', name: 'Scan to buy',
            data: '{url}', fx: clamp(SAFE.right - 0.24, SAFE.left, SAFE.right - 0.2), fy: 0.5, fw: 0.22,
            timing: t(Math.max(1, duration * 0.62), duration),
            enter: { kind: 'fade', seconds: 0.5 }, exit: { kind: 'none', seconds: 0 },
        });
        notes.push('a QR in the last third');
    }

    // Caption lines: the same shape splitCaption writes, so the operator's
    // "Back to one caption" still undoes it.
    if (segs.length > 1) {
        const lead = 0.2;
        const t0 = 0.8;
        const t1 = Math.max(t0 + 1, duration - 0.6);
        const weights = segs.map((s) => Math.max(1, s.length));
        const total = weights.reduce((a, b) => a + b, 0);
        let cur = t0;
        segs.forEach((s, i) => {
            const len = ((t1 - t0) * weights[i]) / total;
            patches.push({
                id: `caption-line-${i + 1}`, type: 'caption', name: `Line ${i + 1}`,
                text: s, leadIn: lead, timing: t(cur, cur + len),
            });
            cur += len;
        });
        patches.push({ id: 'caption', visible: false });
        notes.push(`the caption split into ${segs.length} timed lines`);
    }

    // Options are only FILLED, never overwritten: an operator who set an outro
    // already made this decision.
    const options = {};
    if (!opts.outroSeconds && ctx.url) {
        options.outroSeconds = 2.5;
        options.outroText = opts.outroText || opts.footer || 'Shop now';
        notes.push('an end card');
    }
    if (input.hasTitle && !opts.showTitle) {
        options.showTitle = true;
        notes.push('the title card on');
    }

    return { patches, options, notes };
}

/** Strip a previous draft, so drafting again replaces rather than stacks. */
export function withoutDraft(patches) {
    return (patches || []).filter((p) => p && !String(p.id).startsWith(SB_PREFIX));
}
