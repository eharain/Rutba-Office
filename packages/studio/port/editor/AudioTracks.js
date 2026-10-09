import { useEffect, useState } from "react";
import { audioTracks } from "../../lib/studio-api";
import { Choice } from "./controls";

/**
 * The audio library, fetched once per page load and shared by everything that
 * needs to name a track.
 *
 * TWO PANELS ASK FOR THIS LIST — the music bed and every sound layer — and a
 * video with three sound layers would otherwise issue four identical requests
 * for a list that is org-scoped, small, and changes about as often as someone
 * uploads a track. The cache is a module-level PROMISE rather than a resolved
 * array, so four panels mounting in the same tick share one request instead of
 * racing to start four.
 *
 * IT IS NOT INVALIDATED, and that is a real limitation rather than an
 * oversight: a track uploaded on the /audio page in another tab will not appear
 * here until this one is reloaded. The alternative — a subscription, or a
 * refetch on every focus — is machinery for a case (upload a track, then
 * immediately reach for it in an editor already open) that ends with someone
 * pressing reload anyway. `reloadAudioTracks` is the escape hatch, and the
 * picker offers it as a button rather than hiding it.
 */
let cache = null;

function fetchTracks() {
    cache ??= audioTracks.list()
        .then((rows) => (Array.isArray(rows) ? rows : []))
        // A failed fetch must not be cached as a permanent empty library: clear
        // the cache so the next mount — or the refresh button — tries again.
        .catch((err) => { cache = null; throw err; });
    return cache;
}

/** Drop the cached list so the next read goes to the API. */
export function reloadAudioTracks() {
    cache = null;
}

export function useAudioTracks() {
    const [state, setState] = useState({ tracks: [], loading: true, error: null });

    useEffect(() => {
        let live = true;
        setState((s) => ({ ...s, loading: true }));
        fetchTracks()
            .then((tracks) => { if (live) setState({ tracks, loading: false, error: null }); })
            .catch(() => { if (live) setState({ tracks: [], loading: false, error: "The audio library did not load." }); });
        return () => { live = false; };
    }, []);

    return state;
}

/** Seconds as m:ss — a track list is scanned, not read. */
export function clock(seconds) {
    const s = Math.max(0, Math.round(Number(seconds) || 0));
    return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
}

/**
 * Pick a track by id.
 *
 * `allowNone` exists because the two callers mean different things by "no
 * track": a sound layer with no track is unfinished and should say so, while
 * the music bed reaches "no track" through its MODE (`audioMode: 'none'`) and
 * so never needs the option.
 */
export function TrackChoice({ value, onChange, allowNone = true, tracks, loading, error, onRefresh }) {
    if (loading) return <span className="text-muted small">Loading tracks…</span>;
    if (error) {
        return (
            <div className="d-flex align-items-center gap-2">
                <span className="text-warning small">{error}</span>
                <button type="button" className="btn btn-sm btn-outline-secondary" onClick={onRefresh}>Retry</button>
            </div>
        );
    }
    if (!tracks.length) {
        return (
            <span className="text-muted small">
                No tracks in the library yet — add one on the Audio page.
            </span>
        );
    }

    const options = [
        ...(allowNone ? [{ value: "", label: "— none —" }] : []),
        ...tracks.map((t) => ({
            value: String(t.id),
            label: `${t.title || "Untitled"}${t.duration ? ` · ${clock(t.duration)}` : ""}`,
        })),
    ];

    return (
        <div className="d-flex align-items-center gap-1">
            <Choice
                value={value == null ? "" : String(value)}
                // "" is the sentinel the select can carry; null is what the
                // document stores. Translating here keeps every caller from
                // having to remember which one it is holding.
                onChange={(v) => onChange(v === "" ? null : v)}
                options={options}
            />
            {onRefresh && (
                <button type="button" className="btn btn-sm btn-link px-1" title="Reload the library"
                    onClick={onRefresh}><i className="fa-solid fa-rotate" /></button>
            )}
        </div>
    );
}
