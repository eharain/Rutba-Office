import { Row, Group, Slider, Segmented, Toggle } from "./controls";
import { useAudioTracks, TrackChoice, reloadAudioTracks } from "./AudioTracks";

/**
 * The lane's id, and therefore the selection this panel answers to.
 *
 * Exported from HERE rather than from the timeline because the panel owns what
 * the thing IS; the timeline only draws a lane for it. It cannot collide with a
 * layer either: `nextId` mints ids as `<type>-<n>`, so nothing generated will
 * ever be the literal string below.
 */
export const MUSIC_BED_ID = "music-bed";

const MODES = [
    { value: "none", label: "None" },
    { value: "pick", label: "Choose" },
    { value: "random", label: "Random" },
];

/**
 * The music bed.
 *
 * IT IS NOT A LAYER, and the timeline already says so by drawing it as a
 * read-only lane. The bed lives in `doc.options` — `audioMode`, `audioTrackId`
 * and the volume/fade scalars — because the recipe format predates layers and
 * because `audioMode: 'random'` has to survive as an INSTRUCTION rather than a
 * resolved id: a batch of twenty renders is varied only if each one picks its
 * own track at render time. Storing a chosen id instead would turn "random"
 * into "the same random track, twenty times".
 *
 * This panel is what that lane was missing. It was selectable from the day the
 * timeline learned to draw it, and selecting it did nothing — the inspector
 * renders per PATCH, and the bed has no patch, so it fell through to the empty
 * state. An affordance that responds to a click by showing "select something"
 * is worse than no affordance.
 *
 * THE MODE IS THE PRIMARY CONTROL and everything else is conditional on it,
 * because the three modes want genuinely different things: `none` wants
 * nothing, `pick` wants a track, and `random` wants a policy. Showing a track
 * picker under `random` would invite exactly the misunderstanding the mode
 * exists to prevent.
 */
export default function MusicBedPanel({ doc, setOption, duration }) {
    const options = doc.options || {};
    const mode = options.audioMode || "none";
    const { tracks, loading, error } = useAudioTracks();

    // Continuous controls coalesce into one undo step per drag, the same way a
    // canvas gesture does. A mode change is discrete and gets its own step.
    const cont = (field) => ({ label: field, group: `bed:${field}` });

    return (
        <div className="studio-inspector">
            <header className="studio-inspector-head">
                <div>
                    <div className="studio-inspector-title">
                        <i className="fa-solid fa-music me-2 opacity-75" />
                        Music bed
                    </div>
                    <div className="studio-inspector-id">the whole video</div>
                </div>
            </header>

            <Group title="Track">
                <Row label="Source">
                    <Segmented value={mode} options={MODES}
                        onChange={(audioMode) => setOption({ audioMode }, { label: "music" })} />
                </Row>

                {mode === "pick" && (
                    <Row label="Track">
                        <TrackChoice
                            value={options.audioTrackId}
                            allowNone={false}
                            tracks={tracks} loading={loading} error={error}
                            onRefresh={() => { reloadAudioTracks(); window.location.reload(); }}
                            onChange={(audioTrackId) => setOption({ audioTrackId }, { label: "track" })}
                        />
                    </Row>
                )}

                {mode === "random" && (
                    <p className="studio-note small text-muted mb-2">
                        A fresh track is drawn for every render, from this org&apos;s library plus
                        the shared one. Two renders of the same creative will not sound the same —
                        which is the point when a batch goes out at once.
                    </p>
                )}
            </Group>

            {mode !== "none" && (
                <>
                    <Group title="Level">
                        <Row label="Volume">
                            <Slider value={options.audioVolume ?? 0.7} min={0} max={1} step={0.01}
                                onChange={(audioVolume) => setOption({ audioVolume }, cont("volume"))} />
                        </Row>
                        <Row label="Fade in" hint="seconds">
                            {/* Capped at a third of the video by the renderer
                                anyway (see normalizeAudioClips) — the cap here
                                just stops the slider promising more. */}
                            <Slider value={options.audioFadeIn ?? 1.2} min={0} max={Math.max(1, (duration || 30) / 3)} step={0.1}
                                onChange={(audioFadeIn) => setOption({ audioFadeIn }, cont("fade in"))} />
                        </Row>
                        <Row label="Fade out" hint="seconds">
                            <Slider value={options.audioFadeOut ?? 1.6} min={0} max={Math.max(1, (duration || 30) / 3)} step={0.1}
                                onChange={(audioFadeOut) => setOption({ audioFadeOut }, cont("fade out"))} />
                        </Row>
                    </Group>

                    <Group title="Start">
                        <Row label="Random start"
                            hint="begin somewhere into the track rather than at its intro">
                            <Toggle value={options.audioRandomStart !== false}
                                onChange={(audioRandomStart) => setOption({ audioRandomStart }, { label: "start" })} />
                        </Row>
                        <p className="studio-note small text-muted mb-0">
                            Most library music opens with a build that a fifteen-second video is
                            over before. Starting part-way in lands on the part with the tune.
                        </p>
                    </Group>
                </>
            )}
        </div>
    );
}
