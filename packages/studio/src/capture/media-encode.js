/**
 * What this browser can actually encode, and what to call the result.
 *
 * Every in-browser capture in the suite - recording a take, trimming a library
 * video - ends at MediaRecorder, and MediaRecorder's supported types differ by
 * browser, by version and by machine. One list per job, most-wanted first, asked
 * of the browser rather than assumed.
 */

// Recording a live source. vp8 ahead of vp9 on purpose: encoding runs in real
// time next to a preview, and a dropped frame costs more than the few percent
// vp9 would save.
const LISTS = {
    audio: ["audio/webm;codecs=opus", "audio/webm", "audio/mp4", "audio/ogg;codecs=opus"],
    video: ["video/webm;codecs=vp8,opus", "video/webm;codecs=vp9,opus", "video/webm", "video/mp4"],
    // Re-encoding a file that already exists. mp4 would be the nicer output -
    // every platform uploader takes h264 without argument - but MediaRecorder's
    // mp4 path cannot be relied on: measured 2026-08-13 (Chrome 151), every mp4
    // export of a real clip came back either zero bytes or with a header
    // claiming 0.06s of a 1s trim, while every webm was exact. It stays in the
    // list as a last resort for a browser with no webm encoder at all.
    transcode: [
        "video/webm;codecs=vp8,opus",
        "video/webm;codecs=vp9,opus",
        "video/webm",
        "video/mp4;codecs=avc1.42E01E,mp4a.40.2",
        "video/mp4",
    ],
};

const EXT = {
    "audio/wav": ".wav", "audio/webm": ".webm", "audio/ogg": ".ogg", "audio/mp4": ".m4a", "audio/mpeg": ".mp3",
    "video/webm": ".webm", "video/mp4": ".mp4", "video/x-matroska": ".mkv",
};

/** A mime type stripped of its codec list - the part lookups key on. */
export const baseMime = (m) => String(m || "").split(";")[0].trim().toLowerCase();

/** The best supported type for `kind`, or "" to let the browser choose. */
export function pickRecordMime(kind = "video") {
    if (typeof MediaRecorder === "undefined" || !MediaRecorder.isTypeSupported) return "";
    for (const t of (LISTS[kind] || LISTS.video)) {
        try { if (MediaRecorder.isTypeSupported(t)) return t; } catch { /* keep looking */ }
    }
    return "";
}

/**
 * A type this machine can ACTUALLY encode - asked by encoding something.
 *
 * `isTypeSupported` answers for the muxer, not the encoder, and the two
 * disagree: a browser with no H.264/AAC encoder still says yes to
 * `video/mp4;codecs=avc1…`, accepts it in the constructor, and then writes a
 * ZERO-BYTE file (or one whose header claims a fraction of its real length).
 * Nothing throws. Verified 2026-08-13 in headless Chrome 151, where every mp4
 * export came out empty or 0.17s long while every webm was correct.
 *
 * So each candidate gets a 300 ms trial against a stream shaped like the real
 * one - audio included when the real recording will have audio, since it is the
 * AAC half that fails first - and the first that yields bytes wins. The answer
 * is cached per shape: this costs a fraction of a second, once.
 */
const verified = new Map();

export async function verifiedRecordMime(kind = "transcode", withAudio = false) {
    if (typeof MediaRecorder === "undefined" || typeof document === "undefined") return "";
    const key = `${kind}:${withAudio ? "av" : "v"}`;
    if (verified.has(key)) return verified.get(key);

    const candidates = (LISTS[kind] || LISTS.video).filter((t) => {
        try { return MediaRecorder.isTypeSupported(t); } catch { return false; }
    });
    let chosen = "";
    let probe = null;
    try {
        probe = probeStream(withAudio);
        for (const t of candidates) {
            // eslint-disable-next-line no-await-in-loop
            if (await encodesSomething(probe.stream, t)) { chosen = t; break; }
        }
    } catch {
        // A machine that cannot even build the probe still deserves an answer;
        // the empty-output guard downstream is the backstop.
        chosen = candidates[0] || "";
    } finally {
        probe?.stop();
    }
    verified.set(key, chosen);
    return chosen;
}

/** A tiny live stream shaped like a real recording, for the trial above. */
function probeStream(withAudio) {
    const canvas = document.createElement("canvas");
    canvas.width = 64;
    canvas.height = 64;
    const ctx = canvas.getContext("2d");
    let on = false;
    const paint = () => {
        on = !on;
        ctx.fillStyle = on ? "#fff" : "#000";
        ctx.fillRect(0, 0, 64, 64);
    };
    paint();
    const stream = canvas.captureStream(15);
    const timer = setInterval(paint, 40); // motion, so an encoder has work to do
    let ac = null;
    if (withAudio) {
        const Ctor = window.AudioContext || window.webkitAudioContext;
        ac = new Ctor();
        const osc = ac.createOscillator();
        const gain = ac.createGain();
        gain.gain.value = 0.001;
        const dest = ac.createMediaStreamDestination();
        osc.connect(gain);
        gain.connect(dest);
        osc.start();
        for (const t of dest.stream.getAudioTracks()) stream.addTrack(t);
    }
    return {
        stream,
        stop() {
            clearInterval(timer);
            for (const t of stream.getTracks()) { try { t.stop(); } catch { /* gone */ } }
            try { ac?.close(); } catch { /* already closed */ }
        },
    };
}

function encodesSomething(stream, mimeType, ms = 300) {
    return new Promise((resolve) => {
        let rec;
        try { rec = new MediaRecorder(stream, { mimeType }); } catch { resolve(false); return; }
        let bytes = 0;
        let done = false;
        const finish = (ok) => { if (!done) { done = true; resolve(ok); } };
        rec.ondataavailable = (e) => { bytes += e.data?.size || 0; };
        rec.onerror = () => finish(false);
        rec.onstop = () => finish(bytes > 0);
        try { rec.start(); } catch { finish(false); return; }
        setTimeout(() => { try { rec.stop(); } catch { finish(false); } }, ms);
        setTimeout(() => finish(false), ms + 4000); // a recorder that never stops is a no
    });
}

/** The file extension a mime type should be saved under. */
export function extForMime(mime) {
    const base = baseMime(mime);
    return EXT[base] || (base.startsWith("video/") ? ".webm" : ".wav");
}

/** Why in-browser video encoding is unavailable here, or null when it works. */
export function encodeBlocker() {
    if (typeof document === "undefined") return null; // SSR - decided in the browser
    if (typeof MediaRecorder === "undefined") return "This browser has no MediaRecorder, so it cannot encode video. Use Chrome or Edge.";
    if (typeof document.createElement("canvas").captureStream !== "function") {
        return "This browser cannot capture a canvas as video. Use Chrome or Edge.";
    }
    if (!pickRecordMime("transcode")) return "This browser exposes no video codec MediaRecorder can use. Use Chrome or Edge.";
    return null;
}

/**
 * Bytes for a media url: straight from the host when it allows cross-origin
 * reads, and through the app's own transport when it does not.
 *
 * Direct is tried FIRST and matters more than it looks. The app transport is a
 * same-origin proxy with a size cap and a fetch timeout - sized for a post
 * image, not for a library video - while the media file server answers with
 * `Access-Control-Allow-Origin: *` and honours ranges. Falling back keeps the
 * foreign-URL case working, since those are exactly the hosts that send no CORS
 * headers and can only be reached through the proxy's own allowlist.
 */
export async function fetchMediaBytes(url, fetchMedia) {
    try {
        const r = await fetch(url, { mode: "cors", credentials: "omit" });
        if (r.ok) return await r.blob();
    } catch { /* no CORS here - the transport below is the answer */ }
    if (!fetchMedia) throw new Error("That file could not be read from this page (no cross-origin access, and no media transport given).");
    return fetchMedia(url);
}

const even = (n) => Math.max(2, Math.round(n / 2) * 2); // h264 refuses odd dimensions

/**
 * The output frame for a re-framed video, and the source rectangle that fills
 * it - the biggest centred box of the wanted shape (a cover crop, never a
 * squeeze).
 *
 * `aspect` null keeps the source shape; `shortEdge` 0 keeps the source size.
 * The short edge is what is asked for rather than the width, because "720" means
 * 1280×720 on a landscape clip and 720×1280 on a vertical one - and it is never
 * allowed to exceed what the crop actually has, since upscaling invents nothing
 * but file size.
 */
export function frameFit(sw, sh, aspect = null, shortEdge = 0) {
    const want = aspect || (sw / sh);
    let rw = sw;
    let rh = sh;
    if (sw / sh > want) rw = sh * want; else rh = sw / want;
    const sourceShort = Math.min(rw, rh);
    const short = shortEdge ? Math.min(shortEdge, sourceShort) : sourceShort;
    return {
        sx: (sw - rw) / 2,
        sy: (sh - rh) / 2,
        sw: rw,
        sh: rh,
        w: even(want >= 1 ? short * want : short),
        h: even(want >= 1 ? short : short / want),
    };
}

/** A file-name stem that is safe on every platform, and never empty. */
export const safeStem = (s, fallback) => (
    String(s || "").trim().replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ").slice(0, 80) || fallback
);

const two = (n) => String(n).padStart(2, "0");

/** yyyymmdd-hhmm, for naming a take after the moment it was made. */
export function stampNow() {
    const d = new Date();
    return `${d.getFullYear()}${two(d.getMonth() + 1)}${two(d.getDate())}-${two(d.getHours())}${two(d.getMinutes())}`;
}

/** m:ss, the only clock format any of these dialogs needs. */
export const fmtClock = (s) => {
    const n = Math.max(0, Number(s) || 0);
    return `${Math.floor(n / 60)}:${two(Math.floor(n % 60))}`;
};

/** m:ss.d - the same clock with the tenth that trimming actually turns on. */
export const fmtPrecise = (s) => {
    const n = Math.max(0, Number(s) || 0);
    return `${Math.floor(n / 60)}:${two(Math.floor(n % 60))}.${Math.floor((n % 1) * 10)}`;
};

export const fmtBytes = (b) => (b > 1024 * 1024 ? `${(b / 1024 / 1024).toFixed(1)} MB` : `${Math.max(1, Math.round(b / 1024))} KB`);
