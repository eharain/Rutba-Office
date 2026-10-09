import { createContext, useContext, useMemo } from "react";

/**
 * The inspector's control vocabulary.
 *
 * Every panel in this editor is a list of labelled controls over one patch, and
 * writing that out per panel produced the same eight lines of Bootstrap markup
 * forty times over in the app this was ported from. These are the eight lines,
 * once.
 *
 * ONE CONVENTION THROUGHOUT: a control reports its value and nothing else. It
 * does not know what a layer is, which patch it belongs to, or that undo
 * exists. The panel binds it to a field; the panel's `set` is what carries the
 * undo label and grouping. Controls that knew about the document would have to
 * be re-tested every time the document changed shape.
 */

export function Row({ label, hint, children, wide }) {
    return (
        <div className={`studio-row${wide ? " studio-row-wide" : ""}`}>
            <label className="studio-row-label">
                {label}
                {hint ? <span className="studio-row-hint">{hint}</span> : null}
            </label>
            <div className="studio-row-control">{children}</div>
        </div>
    );
}

export function Group({ title, children, right }) {
    return (
        <section className="studio-group">
            <header className="studio-group-head">
                <span>{title}</span>
                {right}
            </header>
            <div className="studio-group-body">{children}</div>
        </section>
    );
}

export function TextField({ value, onChange, placeholder, rows = 1 }) {
    if (rows > 1) {
        return (
            <textarea
                className="form-control form-control-sm"
                rows={rows}
                value={value ?? ""}
                placeholder={placeholder}
                onChange={(e) => onChange(e.target.value)}
            />
        );
    }
    return (
        <input
            type="text"
            className="form-control form-control-sm"
            value={value ?? ""}
            placeholder={placeholder}
            onChange={(e) => onChange(e.target.value)}
        />
    );
}

/**
 * A slider with its number beside it.
 *
 * The number is not decoration. A slider alone cannot be set to a value someone
 * was told to use, cannot be read back precisely, and cannot be nudged by one
 * step on a touch screen — and every one of those comes up when two people are
 * matching a brand's look.
 */
export function Slider({ value, onChange, min = 0, max = 1, step = 0.01, format }) {
    const v = Number.isFinite(Number(value)) ? Number(value) : min;
    return (
        <div className="studio-slider">
            <input
                type="range" className="form-range"
                min={min} max={max} step={step} value={v}
                onChange={(e) => onChange(Number(e.target.value))}
            />
            <input
                type="number" className="form-control form-control-sm studio-slider-num"
                min={min} max={max} step={step} value={v}
                onChange={(e) => {
                    const n = Number(e.target.value);
                    if (Number.isFinite(n)) onChange(Math.max(min, Math.min(max, n)));
                }}
            />
            {format ? <span className="studio-slider-unit">{format}</span> : null}
        </div>
    );
}

export function Choice({ value, onChange, options }) {
    return (
        <select
            className="form-select form-select-sm"
            value={value ?? ""}
            onChange={(e) => onChange(e.target.value)}
        >
            {options.map((o) => {
                const key = typeof o === "string" ? o : o.value;
                const label = typeof o === "string" ? o : o.label;
                return <option key={key} value={key}>{label}</option>;
            })}
        </select>
    );
}

/** Segmented buttons — for a choice of three or four where seeing them all matters. */
export function Segmented({ value, onChange, options }) {
    return (
        <div className="btn-group btn-group-sm w-100" role="group">
            {options.map((o) => {
                const key = typeof o === "string" ? o : o.value;
                const label = typeof o === "string" ? o : o.label;
                return (
                    <button
                        key={key} type="button"
                        className={`btn ${value === key ? "btn-info" : "btn-outline-secondary"}`}
                        onClick={() => onChange(key)}
                    >
                        {label}
                    </button>
                );
            })}
        </div>
    );
}

export function Toggle({ value, onChange, label }) {
    return (
        <div className="form-check form-switch mb-0">
            <input
                className="form-check-input" type="checkbox"
                checked={Boolean(value)}
                onChange={(e) => onChange(e.target.checked)}
            />
            {label ? <span className="form-check-label ms-1 small">{label}</span> : null}
        </div>
    );
}

/**
 * The brand palette, if one is applied to the document being edited.
 *
 * A CONTEXT RATHER THAN A PROP, because the alternative is threading a palette
 * through the inspector, five panels, and every FillField and ColorField inside
 * them — for something none of them has an opinion about. The swatches are
 * ambient by nature: "the colours this creative is allowed to use" is a fact
 * about the document, not an argument to a slider.
 */
const BrandPaletteContext = createContext([]);

export function BrandPaletteProvider({ colors, children }) {
    // Memoised on the joined value rather than the array identity: the doc is
    // rebuilt on every keystroke, and a fresh array each time would re-render
    // every control in the inspector for a palette that had not changed.
    const value = useMemo(() => (Array.isArray(colors) ? colors : []), [(colors || []).join("|")]);
    return <BrandPaletteContext.Provider value={value}>{children}</BrandPaletteContext.Provider>;
}

/** The brand swatch row. Renders nothing at all when no kit is applied. */
function Swatches({ onPick }) {
    const colors = useContext(BrandPaletteContext);
    if (!colors.length) return null;
    return (
        <div className="studio-swatches">
            {colors.map((c, i) => (
                <button
                    key={`${c}-${i}`}
                    type="button"
                    className="studio-swatch"
                    style={{ background: c }}
                    title={c}
                    aria-label={`Brand colour ${c}`}
                    onClick={() => onPick(c)}
                />
            ))}
        </div>
    );
}

/**
 * A colour, with a text field beside the swatch.
 *
 * `<input type="color">` cannot express `rgba(0,0,0,0.45)` — it has no alpha —
 * and a shadow without alpha is a black box rather than a shadow. So the text
 * field is the authoritative control and the picker is the convenient one.
 *
 * The brand row sits under both. It is the whole point of a brand kit being
 * "applied" rather than merely stored: the palette is one click away at every
 * point a colour is chosen, so staying on-brand is the path of least effort
 * rather than a thing to remember.
 */
export function ColorField({ value, onChange, allowNone }) {
    const hex = /^#[0-9a-f]{6}$/i.test(String(value || "")) ? value : "#ffffff";
    return (
        <div>
            <div className="studio-color">
                <input
                    type="color" className="form-control form-control-color form-control-sm"
                    value={hex} onChange={(e) => onChange(e.target.value)}
                />
                <input
                    type="text" className="form-control form-control-sm"
                    value={value ?? ""} placeholder={allowNone ? "none" : "#ffffff"}
                    onChange={(e) => onChange(e.target.value || (allowNone ? null : "#ffffff"))}
                />
            </div>
            <Swatches onPick={onChange} />
        </div>
    );
}

/**
 * A gradient's stops.
 *
 * Stops are `[position, colour]` pairs and are kept SORTED on every edit rather
 * than at paint time, because `addColorStop` throws on an out-of-order offset in
 * some engines and silently misbehaves in others — and a gradient editor whose
 * whole point is dragging a stop past another one will produce that input on
 * its first use.
 */
export function StopsField({ stops, onChange }) {
    const list = Array.isArray(stops) && stops.length ? stops : [[0, "#ffffff"], [1, "#888888"]];
    const write = (next) => onChange([...next].sort((a, b) => a[0] - b[0]));
    return (
        <div className="studio-stops">
            {list.map(([at, color], i) => (
                <div className="studio-stop" key={i}>
                    <input
                        type="number" className="form-control form-control-sm" min={0} max={1} step={0.01}
                        value={at}
                        onChange={(e) => write(list.map((s, j) => (j === i ? [Number(e.target.value), s[1]] : s)))}
                    />
                    <input
                        type="color" className="form-control form-control-color form-control-sm"
                        value={/^#[0-9a-f]{6}$/i.test(color) ? color : "#ffffff"}
                        onChange={(e) => write(list.map((s, j) => (j === i ? [s[0], e.target.value] : s)))}
                    />
                    <button
                        type="button" className="btn btn-sm btn-outline-secondary"
                        disabled={list.length <= 2}
                        title={list.length <= 2 ? "A gradient needs at least two stops" : "Remove this stop"}
                        onClick={() => write(list.filter((_, j) => j !== i))}
                    >
                        <i className="fa-solid fa-minus" />
                    </button>
                </div>
            ))}
            <button
                type="button" className="btn btn-sm btn-outline-info w-100"
                onClick={() => write([...list, [0.5, "#ffffff"]])}
            >
                <i className="fa-solid fa-plus me-1" /> Add stop
            </button>
        </div>
    );
}

/**
 * The fill editor, shared by word art and shapes.
 *
 * Both types store the same `{ kind, color | stops, angle }` shape and paint it
 * with near-identical code; giving them one control is what keeps a gradient
 * behaving the same way on a headline and on the badge behind it.
 */
export function FillField({ fill, onChange }) {
    const f = fill || { kind: "solid", color: "#ffffff" };
    return (
        <>
            <Row label="Fill">
                <Segmented
                    value={f.kind || "solid"}
                    onChange={(kind) => onChange({ ...f, kind })}
                    options={[
                        { value: "solid", label: "Solid" },
                        { value: "linear", label: "Linear" },
                        { value: "radial", label: "Radial" },
                    ]}
                />
            </Row>
            {(!f.kind || f.kind === "solid") ? (
                <Row label="Colour">
                    <ColorField value={f.color} onChange={(color) => onChange({ ...f, color })} />
                </Row>
            ) : (
                <>
                    {f.kind === "linear" && (
                        <Row label="Angle" hint="degrees">
                            <Slider value={f.angle ?? 90} min={0} max={360} step={1}
                                onChange={(angle) => onChange({ ...f, angle })} />
                        </Row>
                    )}
                    <Row label="Stops" wide>
                        <StopsField stops={f.stops} onChange={(stops) => onChange({ ...f, stops })} />
                    </Row>
                </>
            )}
        </>
    );
}

/**
 * Blend and mask — the two things the frame wrapper applies to every layer type.
 *
 * They live in one shared control for the same reason they live in the wrapper:
 * they behave identically for a photo, a headline and a shape, so an editor that
 * offered them per type would be describing an implementation that does not
 * exist.
 */
export const BLEND_MODES = [
    { value: "", label: "Normal" },
    { value: "multiply", label: "Multiply" },
    { value: "screen", label: "Screen" },
    { value: "overlay", label: "Overlay" },
    { value: "darken", label: "Darken" },
    { value: "lighten", label: "Lighten" },
    { value: "color-dodge", label: "Colour dodge" },
    { value: "color-burn", label: "Colour burn" },
    { value: "hard-light", label: "Hard light" },
    { value: "soft-light", label: "Soft light" },
    { value: "difference", label: "Difference" },
    { value: "hue", label: "Hue" },
    { value: "saturation", label: "Saturation" },
    { value: "color", label: "Colour" },
    { value: "luminosity", label: "Luminosity" },
];
