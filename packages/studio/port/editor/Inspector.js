import { useState } from "react";
import { WORDART_PRESETS, WARPS, GEOMETRIES, ADJUST_PRESETS, ICONS, ICON_GROUPS, ICON_LABELS, LIST_KINDS, LIST_LABELS, CHART_KINDS, CHART_LABELS, CHART_PALETTE } from "../../lib/renderer";
import { useAudioTracks, TrackChoice, reloadAudioTracks } from "./AudioTracks";
import {
    Row, Group, TextField, Slider, Choice, Segmented, Toggle, ColorField, FillField, BLEND_MODES,
} from "./controls";
import { dataToText, textToData } from "../../lib/chart-data";
import { slotOf, setSlot } from "../../lib/template-doc";

/**
 * The inspector: everything about the selected layer.
 *
 * ONE PANEL PER LAYER TYPE, plus a common panel every type gets. The split
 * follows the renderer's own: opacity, rotation, blend and mask are applied by
 * the frame wrapper for every layer type, so they belong to the common panel;
 * a warp belongs only to word art, so it belongs to that panel. Where the
 * editor's structure matches the renderer's, a new tool needs a new panel and
 * changes nothing else.
 *
 * `set` carries the undo label. Panels pass a group for anything continuous (a
 * slider drag is one undo step) and no group for a discrete choice (picking a
 * preset is its own step, even if you pick three in a row).
 */
export default function Inspector({ patch, layer, set, remove, duplicate, reorder }) {
    if (!patch) {
        return (
            <div className="studio-inspector studio-inspector-empty">
                <i className="fa-regular fa-hand-pointer fa-2x mb-2 opacity-50" />
                <p className="small mb-0">Select something on the canvas, or add a layer.</p>
            </div>
        );
    }

    return (
        <div className="studio-inspector">
            <header className="studio-inspector-head">
                <div>
                    <div className="studio-inspector-title">{patch.name || labelFor(patch.type)}</div>
                    <div className="studio-inspector-id">{patch.id}</div>
                </div>
                <div className="btn-group btn-group-sm">
                    <button type="button" className="btn btn-outline-secondary" title="Move down"
                        onClick={() => reorder(-1)}><i className="fa-solid fa-arrow-down" /></button>
                    <button type="button" className="btn btn-outline-secondary" title="Move up"
                        onClick={() => reorder(1)}><i className="fa-solid fa-arrow-up" /></button>
                    <button type="button" className="btn btn-outline-secondary" title="Duplicate"
                        onClick={duplicate}><i className="fa-regular fa-clone" /></button>
                    <button type="button" className="btn btn-outline-danger" title="Delete"
                        onClick={remove}><i className="fa-regular fa-trash-can" /></button>
                </div>
            </header>

            {patch.type === "wordart" && <WordArtPanel patch={patch} set={set} />}
            {patch.type === "shape" && <ShapePanel patch={patch} set={set} />}
            {patch.type === "adjust" && <AdjustPanel patch={patch} set={set} />}
            {patch.type === "text" && <PlainTextPanel patch={patch} set={set} />}
            {patch.type === "sound" && <SoundPanel patch={patch} set={set} />}
            {patch.type === "icon" && <IconPanel patch={patch} set={set} />}
            {patch.type === "video" && <VideoClipPanel patch={patch} layer={layer} set={set} />}
            {patch.type === "textblock" && <TextBlockPanel patch={patch} set={set} />}
            {patch.type === "chart" && <ChartPanel patch={patch} set={set} />}

            {/* A sound never paints, so opacity, rotation, blend and masking —
                everything the common panel offers — would be four controls with
                nothing to act on. It is the one layer type that skips it. */}
            {patch.type !== "sound" && <CommonPanel patch={patch} layer={layer} set={set} />}
        </div>
    );
}

const labelFor = (type) => ({
    wordart: "Word art", shape: "Shape", adjust: "Adjustment", text: "Text",
    image: "Image", qr: "QR code", sound: "Sound", icon: "Icon", video: "Clip", textblock: "Text", chart: "Chart",
}[type] || type);

// ── charts ───────────────────────────────────────────────────────────────────

/**
 * A chart layer.
 *
 * The controls that exist are the ones a creative needs; the ones that do not
 * are deliberate. There is no tick count, no axis title and no minimum: a
 * creative is read at a glance from a phone, and a chart with eleven labelled
 * ticks on it is a screenshot of a spreadsheet.
 */
function ChartPanel({ patch, set }) {
    const kind = patch.chart || "bar";
    const isRound = kind === "pie" || kind === "donut";

    return (
        <>
            <Group title="Chart">
                <Row label="Kind">
                    <Choice value={kind}
                        options={CHART_KINDS.map((k) => ({ value: k, label: CHART_LABELS[k] }))}
                        onChange={(chart) => set({ chart }, { label: "chart kind" })} />
                </Row>
                <Row label="Data" wide hint="one “label, value” per line">
                    <TextField value={dataToText(patch)} rows={5}
                        placeholder={"Mon, 12\nTue, 19\nWed, 9"}
                        onChange={(text) => {
                            const { labels, values } = textToData(text);
                            // Written as the shorthand and the series form is
                            // cleared: two sources of the same numbers is the
                            // one shape readData cannot disambiguate.
                            set({ labels, values, series: null }, { label: "data", group: `data:${patch.id}` });
                        }} />
                </Row>
            </Group>

            <Group title="Look">
                <Row label="Text size" hint="× the frame width">
                    <Slider value={patch.sizeFrac ?? 0.022} min={0.008} max={0.06} step={0.002}
                        onChange={(sizeFrac) => set({ sizeFrac }, { label: "text size", group: `cs:${patch.id}` })} />
                </Row>
                <Row label="Show labels">
                    <Toggle value={patch.showLabels !== false}
                        onChange={(showLabels) => set({ showLabels }, { label: "labels" })} />
                </Row>
                <Row label="Show numbers">
                    <Toggle value={kind === "line" || kind === "area" ? patch.showValues === true : patch.showValues !== false}
                        onChange={(showValues) => set({ showValues }, { label: "numbers" })} />
                </Row>
                {!isRound && (
                    <Row label="Gridlines">
                        <Toggle value={patch.showGrid !== false}
                            onChange={(showGrid) => set({ showGrid }, { label: "gridlines" })} />
                    </Row>
                )}
                {isRound && (
                    <Row label="Legend">
                        <Toggle value={patch.showLegend !== false}
                            onChange={(showLegend) => set({ showLegend }, { label: "legend" })} />
                    </Row>
                )}
            </Group>

            <Group title={isRound ? "Slices" : "Bars"}>
                {kind === "donut" && (
                    <>
                        <Row label="Hole">
                            <Slider value={patch.innerRatio ?? 0.58} min={0.1} max={0.9} step={0.02}
                                onChange={(innerRatio) => set({ innerRatio }, { label: "hole", group: `ir:${patch.id}` })} />
                        </Row>
                        <Row label="In the middle" wide>
                            <TextField value={patch.centerText || ""} placeholder="58%"
                                onChange={(centerText) => set({ centerText }, { label: "centre text", group: `ct:${patch.id}` })} />
                        </Row>
                    </>
                )}
                {(kind === "bar" || kind === "column") && (
                    <>
                        <Row label="Spacing">
                            <Slider value={patch.gap ?? 0.28} min={0} max={0.8} step={0.02}
                                onChange={(gap) => set({ gap }, { label: "spacing", group: `gap:${patch.id}` })} />
                        </Row>
                        <Row label="Rounded ends">
                            <Slider value={patch.radius ?? 0.18} min={0} max={0.5} step={0.02}
                                onChange={(radius) => set({ radius }, { label: "rounding", group: `rad:${patch.id}` })} />
                        </Row>
                    </>
                )}
                {(kind === "line" || kind === "area") && (
                    <>
                        <Row label="Line weight">
                            <Slider value={patch.lineWidth ?? 0.006} min={0.001} max={0.02} step={0.001}
                                onChange={(lineWidth) => set({ lineWidth }, { label: "line weight", group: `lw:${patch.id}` })} />
                        </Row>
                        <Row label="Points">
                            <Toggle value={patch.showPoints !== false}
                                onChange={(showPoints) => set({ showPoints }, { label: "points" })} />
                        </Row>
                    </>
                )}
                <Row label="Before / after" hint="a currency, a unit">
                    <div className="d-flex gap-1">
                        <TextField value={patch.prefix || ""} placeholder="£"
                            onChange={(prefix) => set({ prefix }, { label: "prefix", group: `pre:${patch.id}` })} />
                        <TextField value={patch.suffix || ""} placeholder="%"
                            onChange={(suffix) => set({ suffix }, { label: "suffix", group: `suf:${patch.id}` })} />
                    </div>
                </Row>
            </Group>

            <Group title="Colours">
                <p className="studio-note small text-muted mb-1">
                    Categories take these in order. A brand kit replaces them.
                </p>
                <div className="studio-icon-grid">
                    {(patch.palette || CHART_PALETTE).map((c, i) => (
                        <input
                            key={`${c}-${i}`} type="color" className="form-control form-control-color p-0 border-0"
                            style={{ height: 26 }}
                            value={c}
                            onChange={(e) => {
                                const next = [...(patch.palette || CHART_PALETTE)];
                                next[i] = e.target.value;
                                set({ palette: next }, { label: "colour", group: `pal:${patch.id}:${i}` });
                            }}
                        />
                    ))}
                </div>
                {patch.palette && (
                    <button type="button" className="btn btn-sm btn-link px-0"
                        onClick={() => set({ palette: null }, { label: "reset colours" })}>
                        Back to the defaults
                    </button>
                )}
            </Group>
        </>
    );
}

// ── text blocks ──────────────────────────────────────────────────────────────

const ALIGNS = [
    { value: "left", label: "Left" }, { value: "center", label: "Centre" }, { value: "right", label: "Right" },
];
const VALIGNS = [
    { value: "top", label: "Top" }, { value: "middle", label: "Middle" }, { value: "bottom", label: "Bottom" },
];

/**
 * A paragraph, or a list — the same layer either way.
 *
 * THE LIST MODE IS A CONTROL RATHER THAN A LAYER TYPE, because a list is a
 * paragraph whose lines carry a marker and a hanging indent. That is also why
 * turning a paragraph into a checklist keeps every word, every alignment and
 * every colour: nothing is rebuilt, one field changes.
 *
 * `Fit to the box` is off-by-default-looking but on: text that overflows its
 * box shrinks to fit. It only ever SHRINKS — a block that grew as somebody
 * typed would be the opposite of a layout — so the size slider is an upper
 * bound rather than a target, and that is what the hint says.
 */
function TextBlockPanel({ patch, set }) {
    const hasBox = patch.fh != null;

    return (
        <>
            <Group title="Text">
                <Row label="Words" wide>
                    <TextField value={patch.text} rows={4}
                        placeholder={"One line per item\nAnother line"}
                        onChange={(text) => set({ text }, { label: "typing", group: `text:${patch.id}` })} />
                </Row>
                <Row label="Kind">
                    <Choice
                        value={patch.list || "none"}
                        options={LIST_KINDS.map((k) => ({ value: k, label: LIST_LABELS[k] }))}
                        onChange={(list) => set({ list }, { label: "list" })}
                    />
                </Row>
                {patch.list && patch.list !== "none" && (
                    <>
                        <Row label="Indent" hint="× the text size">
                            <Slider value={patch.indent ?? 1.35} min={0.4} max={4} step={0.05}
                                onChange={(indent) => set({ indent }, { label: "indent", group: `indent:${patch.id}` })} />
                        </Row>
                        <Row label="Marker colour">
                            <ColorField allowNone value={patch.markerColor || null}
                                onChange={(markerColor) => set({ markerColor }, { label: "marker colour" })} />
                        </Row>
                    </>
                )}
            </Group>

            <Group title="Type">
                <Row label="Size" hint="× the frame width">
                    <Slider value={patch.sizeFrac ?? 0.035} min={0.012} max={0.12} step={0.002}
                        onChange={(sizeFrac) => set({ sizeFrac }, { label: "size", group: `size:${patch.id}` })} />
                </Row>
                <Row label="Weight">
                    <Choice value={String(patch.weight ?? 500)}
                        options={[["300", "Light"], ["400", "Regular"], ["500", "Medium"], ["700", "Bold"], ["900", "Black"]]
                            .map(([value, label]) => ({ value, label }))}
                        onChange={(w) => set({ weight: Number(w) }, { label: "weight" })} />
                </Row>
                <Row label="Italic">
                    <Toggle value={Boolean(patch.italic)} onChange={(italic) => set({ italic }, { label: "italic" })} />
                </Row>
                <Row label="Line spacing">
                    <Slider value={patch.lineHeight ?? 1.35} min={0.9} max={2.4} step={0.05}
                        onChange={(lineHeight) => set({ lineHeight }, { label: "line spacing", group: `lh:${patch.id}` })} />
                </Row>
                <Row label="Letter spacing" hint="× the text size">
                    <Slider value={patch.tracking ?? 0} min={-0.05} max={0.4} step={0.005}
                        onChange={(tracking) => set({ tracking }, { label: "letter spacing", group: `tr:${patch.id}` })} />
                </Row>
                <Row label="Colour">
                    <FillField fill={patch.fill} onChange={(fill) => set({ fill }, { label: "colour" })} />
                </Row>
            </Group>

            <Group title="Box">
                <Row label="Width" hint="× the frame width">
                    <Slider value={patch.fw ?? 0.7} min={0.1} max={1} step={0.01}
                        onChange={(fw) => set({ fw }, { label: "width", group: `fw:${patch.id}` })} />
                </Row>
                <Row label="Align">
                    <Segmented value={patch.align || "left"} options={ALIGNS}
                        onChange={(align) => set({ align }, { label: "align" })} />
                </Row>
                <Row label="Fixed height" hint="off lets the box grow with the words">
                    <Toggle value={hasBox}
                        onChange={(on) => set({ fh: on ? 0.25 : null }, { label: "height" })} />
                </Row>
                {hasBox && (
                    <>
                        <Row label="Height" hint="× the frame height">
                            <Slider value={patch.fh ?? 0.25} min={0.05} max={1} step={0.01}
                                onChange={(fh) => set({ fh }, { label: "height", group: `fh:${patch.id}` })} />
                        </Row>
                        <Row label="Sits">
                            <Segmented value={patch.vAlign || "top"} options={VALIGNS}
                                onChange={(vAlign) => set({ vAlign }, { label: "vertical align" })} />
                        </Row>
                        <Row label="Fit to the box" hint="shrinks to fit — never grows">
                            <Toggle value={patch.autoFit !== false}
                                onChange={(autoFit) => set({ autoFit }, { label: "fit" })} />
                        </Row>
                    </>
                )}
            </Group>
        </>
    );
}

// ── video clips ──────────────────────────────────────────────────────────────

/**
 * A video clip — the PICTURE half of it.
 *
 * A clip is two layers in this renderer, deliberately: the picture and a sound
 * layer on the same url. That is what lets you trim what you see and what you
 * hear independently, and it is why there is no volume control here — the sound
 * has its own lane and its own panel, and duplicating the control would give
 * two places to set one thing.
 *
 * `offset` is the one control that is not obvious and is always needed: where
 * INSIDE the source clip this window starts. Trimming the lane moves the window
 * along the video's timeline; this moves the video under the window. Recording
 * three takes into one file and keeping the third is exactly this control.
 *
 * WHETHER IT FILLS THE FRAME is decided by whether it has a width. A clip with
 * no `fw` covers the frame like a backdrop; giving it one makes it an inset you
 * can drag, which is the picture-in-picture case. Offered as a switch because
 * "why did my video suddenly go full screen when I dragged it" is otherwise a
 * mystery with no visible cause.
 */
function VideoClipPanel({ patch, layer, set }) {
    const inset = patch.fw != null;
    const duration = layer?.timing ? layer.timing.end - layer.timing.start : null;

    return (
        <Group title="Clip">
            <Row label="Starts at" hint="seconds into the recording">
                <Slider value={patch.offset ?? 0} min={0} max={Math.max(10, (patch.offset ?? 0) + 30)} step={0.1}
                    onChange={(offset) => set({ offset }, { label: "clip start", group: `offset:${patch.id}` })} />
            </Row>
            <Row label="Fills the frame" hint="off makes it an inset you can drag">
                <Toggle
                    value={!inset}
                    onChange={(full) => set(
                        // Dropping fw/fh is what "cover the frame" MEANS to the
                        // painter — there is no separate flag, so the switch has
                        // to clear them rather than set something.
                        full ? { fw: null, fh: null } : { fw: 0.42, fh: null, fx: 0.72, fy: 0.74 },
                        { label: "clip framing" },
                    )}
                />
            </Row>
            {duration != null && (
                <p className="studio-note small text-muted mb-0">
                    {duration.toFixed(1)}s on the timeline. Drag its lane to move or trim it.
                </p>
            )}
        </Group>
    );
}

// ── icons ────────────────────────────────────────────────────────────────────

/**
 * An icon layer: which glyph, and how it is painted.
 *
 * THE GLYPH PICKER IS A GRID, NOT A DROPDOWN. Fifty-eight names in a `<select>`
 * is a list you scroll while reading words, when the thing you are choosing is
 * a picture — you would be matching "droplet" against a mental image instead of
 * just seeing it. The grid is the same tiles the insert palette uses, filtered
 * by the same search, because swapping one icon for another is the commonest
 * edit an icon layer gets and it should not mean deleting and re-inserting.
 */
function IconPanel({ patch, set }) {
    const [query, setQuery] = useState("");
    const q = query.trim().toLowerCase();

    const groups = ICON_GROUPS
        .map((g) => ({
            ...g,
            names: g.names.filter((name) => !q
                || `${ICON_LABELS[name] || name} ${g.label}`.toLowerCase().includes(q)),
        }))
        .filter((g) => g.names.length);

    return (
        <>
            <Group title="Glyph" right={<span className="text-muted">{ICON_LABELS[patch.name] || patch.name}</span>}>
                <Row label="Find" wide>
                    <TextField value={query} placeholder="cart, star, clock…" onChange={setQuery} />
                </Row>
                <div className="studio-icon-picker">
                    {groups.length === 0 && <p className="text-muted small mb-0">Nothing matches.</p>}
                    {groups.map((g) => (
                        <div key={g.label}>
                            <div className="studio-insert-group">{g.label}</div>
                            <div className="studio-icon-grid">
                                {g.names.map((name) => (
                                    <button
                                        key={name}
                                        type="button"
                                        title={ICON_LABELS[name] || name}
                                        className={`studio-icon-cell${patch.name === name ? " is-selected" : ""}`}
                                        onClick={() => set({ name }, { label: "icon" })}
                                    >
                                        <svg width={20} height={20} viewBox="0 0 24 24" aria-hidden="true">
                                            <path d={ICONS[name]} fill="currentColor" fillRule="evenodd" />
                                        </svg>
                                    </button>
                                ))}
                            </div>
                        </div>
                    ))}
                </div>
            </Group>

            <Group title="Paint">
                <Row label="Fill">
                    <FillField fill={patch.fill} onChange={(fill) => set({ fill }, { label: "fill" })} />
                </Row>
                <Row label="Size" hint="× the frame width">
                    <Slider value={patch.fw ?? 0.16} min={0.02} max={0.9} step={0.01}
                        onChange={(fw) => set({ fw }, { label: "size", group: `size:${patch.id}` })} />
                </Row>
                <Row label="Outline">
                    <Slider value={patch.stroke?.width ?? 0} min={0} max={0.02} step={0.001}
                        onChange={(width) => set({ stroke: width > 0 ? { ...(patch.stroke || {}), width } : null },
                            { label: "outline", group: `stroke:${patch.id}` })} />
                </Row>
                {patch.stroke?.width > 0 && (
                    <Row label="Outline colour">
                        <ColorField value={patch.stroke?.color || "#000000"}
                            onChange={(color) => set({ stroke: { ...(patch.stroke || {}), color } }, { label: "outline colour" })} />
                    </Row>
                )}
            </Group>
        </>
    );
}

// ── sound ────────────────────────────────────────────────────────────────────

const MIXES = [
    { value: "mix", label: "Mix" },
    { value: "duck", label: "Duck" },
    { value: "solo", label: "Solo" },
];

/**
 * A sound layer: a voice-over, a sting, a clip's own audio.
 *
 * IT IS A LAYER LIKE ANY OTHER as far as the document is concerned — it has an
 * id, an order, a visibility and a lane on the timeline — which is why its
 * TIMING is not here. Where a sound starts and stops is a trim, and a trim is
 * something you do by dragging its lane; a pair of number boxes duplicating
 * that would be a second way to say the same thing, and the two would disagree
 * the first time one of them rounded.
 *
 * `mix` IS THE CONTROL WORTH KNOWING ABOUT. The renderer has carried
 * mix/duck/solo since the lift — `duckEnvelopes` builds a real gain envelope
 * per clip, ramping 0.35s outside the window so the dip is settled before
 * anyone speaks — and until now there was no way to ask for it. Ducking is what
 * separates a voice-over from a voice competing with a song.
 */
function SoundPanel({ patch, set }) {
    const { tracks, loading, error } = useAudioTracks();
    const cont = (field) => ({ label: field, group: `sound:${patch.id}:${field}` });

    return (
        <>
            <Group title="Sound">
                <Row label="Name">
                    <TextField value={patch.name || ""} placeholder="Voice-over"
                        onChange={(name) => set({ name }, { label: "name", group: `name:${patch.id}` })} />
                </Row>
                <Row label="Track">
                    <TrackChoice
                        value={patch.trackId}
                        tracks={tracks} loading={loading} error={error}
                        onRefresh={() => { reloadAudioTracks(); window.location.reload(); }}
                        onChange={(trackId) => set({ trackId }, { label: "track" })}
                    />
                </Row>
                <Row label="Or a URL" wide hint="anything the media proxy can reach">
                    <TextField value={patch.url || ""} placeholder="https://…"
                        onChange={(url) => set({ url: url || null }, { label: "url", group: `url:${patch.id}` })} />
                </Row>
                {!patch.trackId && !patch.url && (
                    <p className="studio-note small text-warning mb-0">
                        This layer has no sound yet — pick a track or paste a URL, or it renders silent.
                    </p>
                )}
            </Group>

            <Group title="Level">
                <Row label="Volume">
                    <Slider value={patch.volume ?? 1} min={0} max={1} step={0.01}
                        onChange={(volume) => set({ volume }, cont("volume"))} />
                </Row>
                <Row label="Against other sound">
                    <Segmented value={patch.mix || "mix"} options={MIXES}
                        onChange={(mix) => set({ mix }, { label: "mix" })} />
                </Row>
                <p className="studio-note small text-muted mb-0">
                    {patch.mix === "duck"
                        ? "Everything else drops to a fifth while this plays, and comes back after — the voice-over dip."
                        : patch.mix === "solo"
                            ? "Everything else is silenced for as long as this plays."
                            : "Plays on top of whatever else is going on, at its own level."}
                </p>
            </Group>

            <Group title="Shape">
                <Row label="Fade in" hint="seconds">
                    <Slider value={patch.enter?.seconds ?? 0} min={0} max={5} step={0.1}
                        onChange={(seconds) => set({ enter: { kind: "fade", seconds } }, cont("fade in"))} />
                </Row>
                <Row label="Fade out" hint="seconds">
                    <Slider value={patch.exit?.seconds ?? 0} min={0} max={5} step={0.1}
                        onChange={(seconds) => set({ exit: { kind: "fade", seconds } }, cont("fade out"))} />
                </Row>
                <Row label="Start into track" hint="seconds — skip an intro">
                    <Slider value={patch.offset ?? 0} min={0} max={120} step={0.5}
                        onChange={(offset) => set({ offset }, cont("offset"))} />
                </Row>
                <Row label="Loop" hint="repeat if the lane outlasts the track">
                    <Toggle value={Boolean(patch.loop)}
                        onChange={(loop) => set({ loop }, { label: "loop" })} />
                </Row>
            </Group>
        </>
    );
}

// ── word art ─────────────────────────────────────────────────────────────────

function WordArtPanel({ patch, set }) {
    const style = patch.style || {};
    const setStyle = (change, opts) => set({ style: change }, opts);
    const strokes = Array.isArray(style.strokes) ? style.strokes : [];
    const shadow = style.shadow || null;

    return (
        <>
            <Group title="Text">
                <Row label="Words" wide>
                    <TextField value={patch.text} rows={2}
                        placeholder="Your headline"
                        onChange={(text) => set({ text }, { label: "typing", group: `text:${patch.id}` })} />
                </Row>
                <Row label="Preset">
                    <Choice
                        value={patch.preset || ""}
                        onChange={(preset) => set({
                            preset: preset || null,
                            // A preset REPLACES the style rather than merging
                            // into it. Merging is what produces a "gold" that is
                            // still wearing the neon glow you picked before it,
                            // and no amount of clicking gets rid of it.
                            style: {},
                        }, { label: "preset" })}
                        options={[{ value: "", label: "Custom" },
                            ...Object.entries(WORDART_PRESETS).map(([value, p]) => ({ value, label: p.label }))]}
                    />
                </Row>
                <Row label="Size" hint="of frame width">
                    <Slider value={patch.sizeFrac ?? 0.09} min={0.02} max={0.4} step={0.005}
                        onChange={(sizeFrac) => set({ sizeFrac }, { label: "size", group: `size:${patch.id}` })} />
                </Row>
                <Row label="Weight">
                    <Choice value={String(style.weight ?? 800)}
                        onChange={(v) => setStyle({ weight: Number(v) }, { label: "weight" })}
                        options={["300", "400", "600", "700", "800", "900"]} />
                </Row>
                <Row label="Case">
                    <Toggle value={style.uppercase} label="UPPERCASE"
                        onChange={(uppercase) => setStyle({ uppercase }, { label: "case" })} />
                </Row>
                <Row label="Align">
                    <Segmented value={style.align || "center"}
                        onChange={(align) => setStyle({ align }, { label: "align" })}
                        options={[{ value: "left", label: "Left" }, { value: "center", label: "Centre" }, { value: "right", label: "Right" }]} />
                </Row>
                <Row label="Tracking" hint="of size">
                    <Slider value={style.tracking ?? 0} min={-0.2} max={0.6} step={0.01}
                        onChange={(tracking) => setStyle({ tracking }, { label: "tracking", group: `tr:${patch.id}` })} />
                </Row>
                <Row label="Leading" hint="line height">
                    <Slider value={style.leading ?? 1.12} min={0.7} max={2.5} step={0.01}
                        onChange={(leading) => setStyle({ leading }, { label: "leading", group: `ld:${patch.id}` })} />
                </Row>
            </Group>

            <Group title="Shape">
                <Row label="Warp">
                    <Choice value={style.warp || "none"}
                        onChange={(warp) => setStyle({ warp }, { label: "warp" })}
                        options={WARPS.map((w) => ({ value: w, label: w[0].toUpperCase() + w.slice(1) }))} />
                </Row>
                {style.warp && style.warp !== "none" && (
                    <Row label={style.warp === "bulge" ? "Amount" : style.warp === "wave" ? "Amplitude" : "Curve"}>
                        <Slider value={style.curve ?? 0.5} min={-1} max={1} step={0.01}
                            onChange={(curve) => setStyle({ curve }, { label: "curve", group: `cv:${patch.id}` })} />
                    </Row>
                )}
                {style.warp === "wave" && (
                    <Row label="Cycles">
                        <Slider value={style.frequency ?? 1.5} min={0.2} max={6} step={0.1}
                            onChange={(frequency) => setStyle({ frequency }, { label: "cycles", group: `fq:${patch.id}` })} />
                    </Row>
                )}
            </Group>

            <Group title="Fill">
                <FillField fill={style.fill} onChange={(fill) => setStyle({ fill }, { label: "fill" })} />
            </Group>

            <Group
                title="Outlines"
                right={(
                    <button type="button" className="btn btn-sm btn-link p-0"
                        onClick={() => setStyle({ strokes: [...strokes, { width: 0.05, color: "#000000" }] }, { label: "outline" })}>
                        <i className="fa-solid fa-plus" /> Add
                    </button>
                )}
            >
                {strokes.length === 0 && <p className="text-muted small mb-0">No outline.</p>}
                {strokes.map((s, i) => (
                    <div className="studio-substack" key={i}>
                        <Row label={`Width ${i + 1}`} hint="of size">
                            <Slider value={s.width ?? 0.05} min={0.005} max={0.3} step={0.005}
                                onChange={(width) => setStyle({
                                    strokes: strokes.map((x, j) => (j === i ? { ...x, width } : x)),
                                }, { label: "outline", group: `sw:${patch.id}:${i}` })} />
                        </Row>
                        <Row label="Colour">
                            <div className="d-flex gap-1">
                                <ColorField value={s.color}
                                    onChange={(color) => setStyle({
                                        strokes: strokes.map((x, j) => (j === i ? { ...x, color } : x)),
                                    }, { label: "outline" })} />
                                <button type="button" className="btn btn-sm btn-outline-danger"
                                    onClick={() => setStyle({ strokes: strokes.filter((_, j) => j !== i) }, { label: "outline" })}>
                                    <i className="fa-regular fa-trash-can" />
                                </button>
                            </div>
                        </Row>
                    </div>
                ))}
            </Group>

            <Group title="Shadow">
                <Row label="Kind">
                    <Choice value={shadow?.kind || ""}
                        onChange={(kind) => setStyle({
                            shadow: kind ? { ...(shadow || {}), kind } : null,
                        }, { label: "shadow" })}
                        options={[
                            { value: "", label: "None" },
                            { value: "drop", label: "Drop" },
                            { value: "glow", label: "Glow" },
                            { value: "long", label: "Long" },
                            { value: "3d", label: "3D extrude" },
                        ]} />
                </Row>
                {shadow && (
                    <>
                        <Row label="Colour">
                            <ColorField value={shadow.color}
                                onChange={(color) => setStyle({ shadow: { ...shadow, color } }, { label: "shadow" })} />
                        </Row>
                        {(shadow.kind === "drop" || shadow.kind === "glow") && (
                            <>
                                <Row label="Blur" hint="of size">
                                    <Slider value={shadow.blur ?? 0.05} min={0} max={0.4} step={0.005}
                                        onChange={(blur) => setStyle({ shadow: { ...shadow, blur } }, { label: "shadow", group: `sb:${patch.id}` })} />
                                </Row>
                                <Row label="Offset X">
                                    <Slider value={shadow.dx ?? 0} min={-0.3} max={0.3} step={0.005}
                                        onChange={(dx) => setStyle({ shadow: { ...shadow, dx } }, { label: "shadow", group: `sx:${patch.id}` })} />
                                </Row>
                                <Row label="Offset Y">
                                    <Slider value={shadow.dy ?? 0} min={-0.3} max={0.3} step={0.005}
                                        onChange={(dy) => setStyle({ shadow: { ...shadow, dy } }, { label: "shadow", group: `sy:${patch.id}` })} />
                                </Row>
                            </>
                        )}
                        {(shadow.kind === "3d" || shadow.kind === "long") && (
                            <>
                                <Row label="Depth" hint="of size">
                                    <Slider value={shadow.depth ?? 0.08} min={0.01} max={0.5} step={0.005}
                                        onChange={(depth) => setStyle({ shadow: { ...shadow, depth } }, { label: "shadow", group: `sd:${patch.id}` })} />
                                </Row>
                                <Row label="Direction" hint="degrees">
                                    <Slider value={shadow.angle ?? 45} min={0} max={360} step={1}
                                        onChange={(angle) => setStyle({ shadow: { ...shadow, angle } }, { label: "shadow", group: `sa:${patch.id}` })} />
                                </Row>
                            </>
                        )}
                    </>
                )}
            </Group>
        </>
    );
}

// ── shapes ───────────────────────────────────────────────────────────────────

/** Which extra dials a geometry actually has. Showing all of them for every shape
 *  would offer a "notch" on an ellipse, which does nothing and reads as broken. */
const SHAPE_DIALS = {
    rect: ["radius"],
    ellipse: [],
    triangle: [],
    polygon: ["points"],
    star: ["points", "innerRatio"],
    burst: ["points", "innerRatio"],
    arrow: ["thickness", "headLength"],
    bubble: ["radius", "tailAt", "tailWidth", "tailHeight"],
    line: ["thickness"],
    banner: ["notch"],
    blob: ["points", "innerRatio"],
};

const DIAL_META = {
    radius: { label: "Corner radius", min: 0, max: 0.5, step: 0.005 },
    points: { label: "Points", min: 3, max: 24, step: 1 },
    innerRatio: { label: "Inner radius", min: 0.05, max: 0.98, step: 0.01 },
    thickness: { label: "Thickness", min: 0.01, max: 1, step: 0.01 },
    headLength: { label: "Head length", min: 0.05, max: 0.9, step: 0.01 },
    notch: { label: "Notch", min: 0, max: 0.45, step: 0.01 },
    tailAt: { label: "Tail position", min: 0, max: 1, step: 0.01 },
    tailWidth: { label: "Tail width", min: 0.02, max: 0.6, step: 0.01 },
    tailHeight: { label: "Tail height", min: 0.02, max: 0.9, step: 0.01 },
};

function ShapePanel({ patch, set }) {
    const dials = SHAPE_DIALS[patch.geometry] || [];
    const stroke = patch.stroke || null;
    const shadow = patch.shadow || null;

    return (
        <>
            <Group title="Shape">
                <Row label="Geometry">
                    <Choice value={patch.geometry || "rect"}
                        onChange={(geometry) => set({ geometry }, { label: "geometry" })}
                        options={GEOMETRIES.map((g) => ({ value: g, label: g[0].toUpperCase() + g.slice(1) }))} />
                </Row>
                <Row label="Width" hint="of frame">
                    <Slider value={patch.fw ?? 0.3} min={0.02} max={1.5} step={0.005}
                        onChange={(fw) => set({ fw }, { label: "size", group: `fw:${patch.id}` })} />
                </Row>
                <Row label="Height" hint="of frame">
                    <Slider value={patch.fh ?? 0.2} min={0.01} max={1.5} step={0.005}
                        onChange={(fh) => set({ fh }, { label: "size", group: `fh:${patch.id}` })} />
                </Row>
                {dials.map((key) => {
                    const m = DIAL_META[key];
                    return (
                        <Row key={key} label={m.label}>
                            <Slider value={patch[key] ?? m.min} min={m.min} max={m.max} step={m.step}
                                onChange={(v) => set({ [key]: v }, { label: m.label, group: `${key}:${patch.id}` })} />
                        </Row>
                    );
                })}
            </Group>

            <Group title="Fill" right={(
                <Toggle value={patch.fill !== null} label="on"
                    onChange={(on) => set({ fill: on ? { kind: "solid", color: "#ffc107" } : null }, { label: "fill" })} />
            )}>
                {patch.fill === null
                    ? <p className="text-muted small mb-0">Outline only.</p>
                    : <FillField fill={patch.fill} onChange={(fill) => set({ fill }, { label: "fill" })} />}
            </Group>

            <Group title="Outline" right={(
                <Toggle value={Boolean(stroke)} label="on"
                    onChange={(on) => set({ stroke: on ? { width: 0.006, color: "#000000" } : null }, { label: "outline" })} />
            )}>
                {!stroke ? <p className="text-muted small mb-0">No outline.</p> : (
                    <>
                        <Row label="Width" hint="of frame width">
                            <Slider value={stroke.width ?? 0.006} min={0.001} max={0.06} step={0.001}
                                onChange={(width) => set({ stroke: { ...stroke, width } }, { label: "outline", group: `ow:${patch.id}` })} />
                        </Row>
                        <Row label="Colour">
                            <ColorField value={stroke.color}
                                onChange={(color) => set({ stroke: { ...stroke, color } }, { label: "outline" })} />
                        </Row>
                    </>
                )}
            </Group>

            <Group title="Shadow" right={(
                <Toggle value={Boolean(shadow)} label="on"
                    onChange={(on) => set({ shadow: on ? { blur: 0.02, dy: 0.01, color: "rgba(0,0,0,0.35)" } : null }, { label: "shadow" })} />
            )}>
                {!shadow ? <p className="text-muted small mb-0">No shadow.</p> : (
                    <>
                        <Row label="Blur"><Slider value={shadow.blur ?? 0.02} min={0} max={0.15} step={0.002}
                            onChange={(blur) => set({ shadow: { ...shadow, blur } }, { label: "shadow", group: `hb:${patch.id}` })} /></Row>
                        <Row label="Offset Y"><Slider value={shadow.dy ?? 0.01} min={-0.1} max={0.1} step={0.002}
                            onChange={(dy) => set({ shadow: { ...shadow, dy } }, { label: "shadow", group: `hy:${patch.id}` })} /></Row>
                        <Row label="Colour"><ColorField value={shadow.color}
                            onChange={(color) => set({ shadow: { ...shadow, color } }, { label: "shadow" })} /></Row>
                    </>
                )}
            </Group>
        </>
    );
}

// ── adjustments ──────────────────────────────────────────────────────────────

const TONE_DIALS = [
    { key: "exposure", label: "Exposure", hint: "stops", min: -3, max: 3, step: 0.05, zero: 0 },
    { key: "brightness", label: "Brightness", min: -0.5, max: 0.5, step: 0.01, zero: 0 },
    { key: "contrast", label: "Contrast", min: -0.9, max: 0.9, step: 0.01, zero: 0 },
];
const COLOUR_DIALS = [
    { key: "temperature", label: "Temperature", hint: "cool ↔ warm", min: -1, max: 1, step: 0.01, zero: 0 },
    { key: "tint", label: "Tint", hint: "green ↔ magenta", min: -1, max: 1, step: 0.01, zero: 0 },
    { key: "vibrance", label: "Vibrance", hint: "protects skin", min: -1, max: 2, step: 0.01, zero: 0 },
    { key: "saturation", label: "Saturation", min: -1, max: 3, step: 0.01, zero: 0 },
];
const EFFECT_DIALS = [
    { key: "sharpen", label: "Sharpen", min: 0, max: 1, step: 0.01, zero: 0 },
    { key: "vignette", label: "Vignette", min: -1, max: 1, step: 0.01, zero: 0 },
];
const LEVEL_DIALS = [
    { key: "inBlack", label: "Black point", min: 0, max: 254, step: 1, zero: 0 },
    { key: "inWhite", label: "White point", min: 1, max: 255, step: 1, zero: 255 },
    { key: "gamma", label: "Gamma", min: 0.1, max: 3, step: 0.01, zero: 1 },
    { key: "outBlack", label: "Output black", min: 0, max: 255, step: 1, zero: 0 },
    { key: "outWhite", label: "Output white", min: 0, max: 255, step: 1, zero: 255 },
];

function AdjustPanel({ patch, set }) {
    const adjust = patch.adjust || {};
    const levels = adjust.levels || {};
    const setAdjust = (change, opts) => set({ adjust: change }, opts);

    const dial = (d) => (
        <Row key={d.key} label={d.label} hint={d.hint}>
            <Slider value={adjust[d.key] ?? d.zero} min={d.min} max={d.max} step={d.step}
                onChange={(v) => setAdjust({ [d.key]: v }, { label: d.label, group: `${d.key}:${patch.id}` })} />
        </Row>
    );

    return (
        <>
            <Group title="Adjustment" right={(
                <button type="button" className="btn btn-sm btn-link p-0"
                    onClick={() => set({ preset: null, adjust: {} }, { label: "reset" })}>Reset</button>
            )}>
                <Row label="Preset">
                    <Choice value={patch.preset || ""}
                        onChange={(preset) => set({ preset: preset || null, adjust: {} }, { label: "preset" })}
                        options={[{ value: "", label: "Custom" },
                            ...Object.entries(ADJUST_PRESETS).map(([value, p]) => ({ value, label: p.label }))]} />
                </Row>
                <p className="text-muted small mb-0">
                    Changes every layer beneath this one. Drag it up or down the stack to choose how much.
                </p>
            </Group>

            <Group title="Tone">{TONE_DIALS.map(dial)}</Group>
            <Group title="Colour">{COLOUR_DIALS.map(dial)}</Group>
            <Group title="Effects">
                {EFFECT_DIALS.map(dial)}
                <Row label="Posterize" hint="0 = off">
                    <Slider value={adjust.posterize ?? 0} min={0} max={16} step={1}
                        onChange={(v) => setAdjust({ posterize: v || undefined }, { label: "posterize", group: `po:${patch.id}` })} />
                </Row>
                <Row label="Invert">
                    <Toggle value={adjust.invert}
                        onChange={(invert) => setAdjust({ invert: invert || undefined }, { label: "invert" })} />
                </Row>
            </Group>

            <Group title="Levels">
                {LEVEL_DIALS.map((d) => (
                    <Row key={d.key} label={d.label}>
                        <Slider value={levels[d.key] ?? d.zero} min={d.min} max={d.max} step={d.step}
                            onChange={(v) => setAdjust({ levels: { ...levels, [d.key]: v } }, { label: d.label, group: `lv${d.key}:${patch.id}` })} />
                    </Row>
                ))}
            </Group>
        </>
    );
}

// ── the plain text layer, and everything common ──────────────────────────────

function PlainTextPanel({ patch, set }) {
    return (
        <Group title="Text">
            <Row label="Words" wide>
                <TextField value={patch.text} rows={2}
                    onChange={(text) => set({ text }, { label: "typing", group: `text:${patch.id}` })} />
            </Row>
            <Row label="Size" hint="of frame width">
                <Slider value={patch.sizeFrac ?? 0.035} min={0.015} max={0.2} step={0.005}
                    onChange={(sizeFrac) => set({ sizeFrac }, { label: "size", group: `size:${patch.id}` })} />
            </Row>
            <Row label="Colour">
                <ColorField value={patch.color} onChange={(color) => set({ color }, { label: "colour" })} />
            </Row>
            <p className="text-muted small mb-0">
                A plain text layer is the renderer&apos;s original — one line, one fill. For gradients,
                outlines or a curve, add word art instead.
            </p>
        </Group>
    );
}

function CommonPanel({ patch, layer, set }) {
    const mask = patch.mask || null;
    return (
        <>
            <Group title="Layer">
                <Row label="Name">
                    <TextField value={patch.name} placeholder={labelFor(patch.type)}
                        onChange={(name) => set({ name }, { label: "rename", group: `nm:${patch.id}` })} />
                </Row>
                <Row label="Visible">
                    <Toggle value={patch.visible !== false}
                        onChange={(v) => set({ visible: v }, { label: v ? "show" : "hide" })} />
                </Row>
                {/* The template-authoring mark. A slot is the author saying
                    "change THIS one" — the annotation lives on the patch, where
                    the compilers never copy it, so the renderer stays deaf to
                    it and the editor keeps it. `set` deep-merges named
                    sub-objects only, so writing the whole slot object (or
                    undefined) works through the ordinary channel... except
                    that updateLayer's shallow merge KEEPS an existing key when
                    the change omits it, so clearing must write `slot: null`
                    rather than deleting. slotOf treats null as absent. */}
                <Row label="Template slot" hint="offered for editing when this is saved as a template">
                    <Toggle value={Boolean(slotOf(patch))}
                        onChange={(on) => set(
                            { slot: on ? (setSlot(patch, true).slot) : null },
                            { label: on ? "mark slot" : "unmark slot" },
                        )} />
                </Row>
                {slotOf(patch) && (
                    <Row label="Slot label" hint="what the person filling it is told">
                        <TextField value={slotOf(patch).label || ""}
                            placeholder={patch.name || patch.type}
                            onChange={(label) => set({ slot: { ...slotOf(patch), label } },
                                { label: "slot label", group: `slot:${patch.id}` })} />
                    </Row>
                )}
                <Row label="Opacity">
                    <Slider value={patch.opacity ?? 1} min={0} max={1} step={0.01}
                        onChange={(opacity) => set({ opacity }, { label: "opacity", group: `op:${patch.id}` })} />
                </Row>
                {patch.type !== "adjust" && (
                    <Row label="Rotation" hint="degrees">
                        <Slider value={patch.rot ?? 0} min={0} max={360} step={1}
                            onChange={(rot) => set({ rot }, { label: "rotate", group: `rt:${patch.id}` })} />
                    </Row>
                )}
                <Row label="Blend">
                    <Choice value={patch.blend || ""}
                        onChange={(blend) => set({ blend: blend || null }, { label: "blend" })}
                        options={BLEND_MODES} />
                </Row>
                {layer?.missingToken && (
                    <p className="text-warning small mb-0">
                        <i className="fa-solid fa-triangle-exclamation me-1" />
                        A <code>{"{token}"}</code> in this layer has no value, so it will not draw.
                    </p>
                )}
            </Group>

            <Group title="Mask" right={(
                <Toggle value={Boolean(mask)} label="on"
                    onChange={(on) => set({ mask: on ? { shape: "ellipse", fx: 0.5, fy: 0.5, fw: 0.6, fh: 0.6 } : null }, { label: "mask" })} />
            )}>
                {!mask ? <p className="text-muted small mb-0">No mask.</p> : (
                    <>
                        <Row label="Shape">
                            <Segmented value={mask.shape || "rect"}
                                onChange={(shape) => set({ mask: { ...mask, shape } }, { label: "mask" })}
                                options={[{ value: "rect", label: "Box" }, { value: "ellipse", label: "Ellipse" }]} />
                        </Row>
                        <Row label="Centre X"><Slider value={mask.fx ?? 0.5} min={0} max={1} step={0.005}
                            onChange={(fx) => set({ mask: { ...mask, fx } }, { label: "mask", group: `mx:${patch.id}` })} /></Row>
                        <Row label="Centre Y"><Slider value={mask.fy ?? 0.5} min={0} max={1} step={0.005}
                            onChange={(fy) => set({ mask: { ...mask, fy } }, { label: "mask", group: `my:${patch.id}` })} /></Row>
                        <Row label="Width"><Slider value={mask.fw ?? 0.6} min={0.02} max={1.5} step={0.005}
                            onChange={(fw) => set({ mask: { ...mask, fw } }, { label: "mask", group: `mw:${patch.id}` })} /></Row>
                        <Row label="Height"><Slider value={mask.fh ?? 0.6} min={0.02} max={1.5} step={0.005}
                            onChange={(fh) => set({ mask: { ...mask, fh } }, { label: "mask", group: `mh:${patch.id}` })} /></Row>
                        <Row label="Invert" hint="punch a hole">
                            <Toggle value={mask.invert}
                                onChange={(invert) => set({ mask: { ...mask, invert } }, { label: "mask" })} />
                        </Row>
                        {patch.type === "adjust" && (
                            <Row label="Feather" hint="soft edge">
                                <Slider value={mask.feather ?? 0} min={0} max={1} step={0.01}
                                    onChange={(feather) => set({ mask: { ...mask, feather } }, { label: "mask", group: `mf:${patch.id}` })} />
                            </Row>
                        )}
                    </>
                )}
            </Group>
        </>
    );
}
