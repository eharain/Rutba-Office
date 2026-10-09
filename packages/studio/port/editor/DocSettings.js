import { Row, Group, TextField, Slider, Choice, Segmented, Toggle } from "./controls";

const TRANSITIONS = [
    { value: "fade", label: "Fade" },
    { value: "cut", label: "Cut" },
    { value: "slide", label: "Slide" },
    { value: "push", label: "Push" },
    { value: "zoom", label: "Zoom" },
];

const FITS = [
    { value: "blur", label: "Whole picture" },
    { value: "cover", label: "Fill the frame" },
];

const LOGO_POSITIONS = [
    { value: "top-left", label: "Top left" },
    { value: "top-right", label: "Top right" },
    { value: "bottom-left", label: "Bottom left" },
    { value: "bottom-right", label: "Bottom right" },
];

const QUALITIES = [
    { value: "high", label: "High" },
    { value: "medium", label: "Medium" },
    { value: "low", label: "Low" },
];

/**
 * The document's own settings — everything that belongs to the creative rather
 * than to any one layer in it.
 *
 * WHERE IT LIVES IS THE DESIGN. It occupies the inspector when NOTHING is
 * selected, which is the slot that until now held "Select something on the
 * canvas, or add a layer." — a full-height panel spending itself on an
 * instruction. The inspector inspects whatever is selected; with nothing
 * selected, the thing in hand is the document, so this is not a new place to
 * look but the honest answer to the question the panel was already asking.
 * Clicking empty canvas is also already how you deselect, so the gesture that
 * gets here is one people have.
 *
 * WHY IT IS WORTH BUILDING AT ALL: every control below was a live renderer
 * option with no way to reach it. `DEFAULTS` in editor-core carries twenty-odd
 * of them — transitions, Ken Burns, the caption pace, the end card, the brand
 * mark's corner — all read on every render, all frozen at whatever the default
 * happened to be. A template could set them, because a template is JSON; a
 * person could not. That asymmetry is what this closes.
 *
 * KIND-AWARE, AND NOT DECORATIVELY SO. `buildImagePlan` hard-forces
 * `showProgress: false`, `edgeFadeSeconds: 0` and `kenBurns: false` for a
 * still, because each only reads at a moving playhead. Offering those three on
 * an image would be three controls that demonstrably do nothing — so they are
 * shown for video only, and the rule is the renderer's rather than this
 * panel's opinion of it.
 */
export default function DocSettings({ doc, setOption, plan }) {
    const o = doc.options || {};
    const isVideo = doc.kind === "video";
    const duration = plan?.duration || 0;

    // Continuous controls coalesce per drag; discrete ones stand alone.
    const cont = (field) => ({ label: field, group: `opt:${field}` });
    const once = (field) => ({ label: field });

    return (
        <div className="studio-inspector">
            <header className="studio-inspector-head">
                <div>
                    <div className="studio-inspector-title">
                        <i className="fa-solid fa-sliders me-2 opacity-75" />
                        {isVideo ? "Video settings" : "Image settings"}
                    </div>
                    <div className="studio-inspector-id">
                        {isVideo && duration ? `${duration.toFixed(1)}s · nothing selected` : "nothing selected"}
                    </div>
                </div>
            </header>

            <p className="studio-note small text-muted px-2 pt-2 mb-2">
                These apply to the whole creative. Select a layer to edit it on its own.
            </p>

            <Group title="Pictures">
                <Row label="Framing" hint="how a picture meets a frame it does not fit">
                    <Segmented value={o.fit || "blur"} options={FITS}
                        onChange={(fit) => setOption({ fit }, once("framing"))} />
                </Row>
                {isVideo && (
                    <>
                        <Row label="Transition">
                            <Choice value={o.transition || "fade"} options={TRANSITIONS}
                                onChange={(transition) => setOption({ transition }, once("transition"))} />
                        </Row>
                        <Row label="Ken Burns" hint="a slow drift across each picture">
                            <Toggle value={o.kenBurns !== false}
                                onChange={(kenBurns) => setOption({ kenBurns }, once("ken burns"))} />
                        </Row>
                    </>
                )}
            </Group>

            <Group title="Words">
                <Row label="Position">
                    <Segmented value={o.textPosition || "bottom"}
                        options={[{ value: "bottom", label: "Bottom" }, { value: "middle", label: "Middle" }]}
                        onChange={(textPosition) => setOption({ textPosition }, once("text position"))} />
                </Row>
                <Row label="Caption" hint="a panel behind the words, or the words alone">
                    <Segmented value={o.captionStyle || "box"}
                        options={[{ value: "box", label: "Panel" }, { value: "bare", label: "Bare" }]}
                        onChange={(captionStyle) => setOption({ captionStyle }, once("caption style"))} />
                </Row>
                <Row label="Text size" hint="× the frame's own scale">
                    <Slider value={o.fontScale ?? 1} min={0.6} max={1.8} step={0.05}
                        onChange={(fontScale) => setOption({ fontScale }, cont("text size"))} />
                </Row>
                <Row label="Show title">
                    <Toggle value={o.showTitle !== false}
                        onChange={(showTitle) => setOption({ showTitle }, once("title"))} />
                </Row>
                <Row label="Footer" wide>
                    <TextField value={o.footer || ""} placeholder="A handle, a site, a disclaimer"
                        onChange={(footer) => setOption({ footer }, { label: "footer", group: "opt:footer" })} />
                </Row>
                {isVideo && (
                    <Row label="Typing speed" hint="characters a second">
                        {/* The caption types itself on; this is how fast. It also
                            sets the video's LENGTH, because the duration is
                            derived from how long the words take to arrive. */}
                        <Slider value={o.charsPerSecond ?? 16} min={6} max={40} step={1}
                            onChange={(charsPerSecond) => setOption({ charsPerSecond }, cont("typing speed"))} />
                    </Row>
                )}
            </Group>

            {isVideo && (
                <Group title="Timing">
                    <Row label="Lead in" hint="seconds before the words start">
                        <Slider value={o.leadInSeconds ?? 0.8} min={0} max={5} step={0.1}
                            onChange={(leadInSeconds) => setOption({ leadInSeconds }, cont("lead in"))} />
                    </Row>
                    <Row label="Tail" hint="seconds held after the last word">
                        <Slider value={o.tailSeconds ?? 1.8} min={0} max={8} step={0.1}
                            onChange={(tailSeconds) => setOption({ tailSeconds }, cont("tail"))} />
                    </Row>
                    <Row label="Title held" hint="seconds">
                        <Slider value={o.titleSeconds ?? 3.2} min={0} max={10} step={0.1}
                            onChange={(titleSeconds) => setOption({ titleSeconds }, cont("title held"))} />
                    </Row>
                    <Row label="Open / close dip" hint="seconds of fade at each end">
                        <Slider value={o.edgeFadeSeconds ?? 0.45} min={0} max={2} step={0.05}
                            onChange={(edgeFadeSeconds) => setOption({ edgeFadeSeconds }, cont("edge fade"))} />
                    </Row>
                    <Row label="Longest" hint="seconds — a hard cap on the whole video">
                        {/* Not a preference: every platform has a limit, and a
                            render that overruns one is discovered at upload. */}
                        <Slider value={o.maxSeconds ?? 60} min={5} max={180} step={5}
                            onChange={(maxSeconds) => setOption({ maxSeconds }, cont("longest"))} />
                    </Row>
                    <Row label="Progress bar">
                        <Toggle value={o.showProgress !== false}
                            onChange={(showProgress) => setOption({ showProgress }, once("progress bar"))} />
                    </Row>
                </Group>
            )}

            {isVideo && (
                <Group title="End card">
                    <Row label="Held for" hint="seconds · 0 turns it off">
                        <Slider value={o.outroSeconds ?? 0} min={0} max={8} step={0.5}
                            onChange={(outroSeconds) => setOption({ outroSeconds }, cont("end card"))} />
                    </Row>
                    {(o.outroSeconds ?? 0) > 0 && (
                        <Row label="Says" wide>
                            <TextField value={o.outroText || ""} rows={2}
                                placeholder={o.footer || "Falls back to the footer"}
                                onChange={(outroText) => setOption({ outroText }, { label: "end card text", group: "opt:outro" })} />
                        </Row>
                    )}
                </Group>
            )}

            <Group title="Brand mark">
                <Row label="Show logo">
                    <Toggle value={o.showLogo !== false}
                        onChange={(showLogo) => setOption({ showLogo }, once("logo"))} />
                </Row>
                {o.showLogo !== false && (
                    <>
                        <Row label="Corner">
                            <Choice value={o.logoPosition || "top-right"} options={LOGO_POSITIONS}
                                onChange={(logoPosition) => setOption({ logoPosition }, once("logo corner"))} />
                        </Row>
                        <Row label="Size" hint="× the frame width">
                            <Slider value={o.logoScale ?? 0.16} min={0.05} max={0.4} step={0.01}
                                onChange={(logoScale) => setOption({ logoScale }, cont("logo size"))} />
                        </Row>
                        <Row label="Opacity">
                            <Slider value={o.logoOpacity ?? 0.92} min={0.1} max={1} step={0.01}
                                onChange={(logoOpacity) => setOption({ logoOpacity }, cont("logo opacity"))} />
                        </Row>
                    </>
                )}
            </Group>

            {isVideo && (
                <Group title="Output">
                    <Row label="Quality" hint="higher costs bitrate, not render time">
                        {/* Worth saying: the render takes as long as the video
                            is long whatever this says, because MediaRecorder
                            records in real time. This only moves the file size. */}
                        <Segmented value={o.quality || "high"} options={QUALITIES}
                            onChange={(quality) => setOption({ quality }, once("quality"))} />
                    </Row>
                </Group>
            )}
        </div>
    );
}
