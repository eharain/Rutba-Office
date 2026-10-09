import { useCallback, useEffect, useMemo, useState } from "react";
import { paintShapeTile } from "../../lib/renderer";
import { insertCatalogFor, searchCatalog } from "../../lib/insert-catalog";

/**
 * The insert palette — everything that can be put on a creative, findable.
 *
 * IT REPLACED A ROW OF BUTTONS, and the reason is arithmetic. Six things fit in
 * a row across a 260px rail; there are now thirty geometries alone, with icons
 * and charts behind them. The failure mode of the row was not that it looked
 * cramped, it was that a shape nobody can find is a shape nobody uses — the
 * feature was built and then hidden.
 *
 * IT LIVES IN THE RAIL, NOT IN A MODAL. A modal is the right shape for picking
 * ONE thing and leaving (MediaPicker is a modal for exactly that reason). But
 * building a creative means inserting several things in a row, and a dialog
 * that must be reopened between each one turns a minute of work into a minute
 * of dismissing. Tabbing the rail costs no screen space and keeps the canvas
 * the same size either way.
 *
 * A SHAPE TILE IS THE SHAPE. Painted by the same code that paints the layer, so
 * the palette cannot drift from the renderer and a geometry added in shapes.js
 * turns up here with no work at all. Everything else — text, effects, sound —
 * gets a named icon, because those are concepts and there is nothing to draw.
 */
export default function InsertPanel({ kind, ask, onAdd }) {
    // Media first, matching the section order: the commonest first act on a
    // blank creative is putting a picture on it, not drawing a trapezoid. (This
    // opened on Shapes for a while — the default predated the Media section and
    // nobody moved it, which is how defaults go stale.)
    const [openKey, setOpenKey] = useState("media");
    const [query, setQuery] = useState("");

    // Opened from elsewhere on one section (the timeline's add row): that
    // section, with any search cleared so it is what shows.
    useEffect(() => {
        if (!ask?.section) return;
        setOpenKey(ask.section);
        setQuery("");
    }, [ask]);

    const sections = useMemo(() => insertCatalogFor(kind), [kind]);
    const results = useMemo(() => searchCatalog(kind, query), [kind, query]);

    return (
        <div className="studio-insert">
            <div className="studio-insert-search">
                <i className="fa-solid fa-magnifying-glass" />
                <input
                    className="form-control form-control-sm"
                    placeholder="Search shapes, text, effects"
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                />
                {query && (
                    <button type="button" className="btn btn-sm btn-link px-1" title="Clear"
                        onClick={() => setQuery("")}><i className="fa-solid fa-xmark" /></button>
                )}
            </div>

            <div className="studio-insert-body">
                {results ? (
                    results.length === 0 ? (
                        <p className="text-muted small px-2 py-3 mb-0">
                            Nothing matches “{query}”.
                        </p>
                    ) : (
                        <Tiles items={results} onAdd={onAdd} showSection />
                    )
                ) : (
                    sections.map((section) => (
                        <section key={section.key} className="studio-insert-section">
                            {/* Accordion rather than tabs across the top: the
                                rail is 260px and four tab labels do not fit
                                across it without truncating to initials. */}
                            <button
                                type="button"
                                className={`studio-insert-head${openKey === section.key ? " is-open" : ""}`}
                                onClick={() => setOpenKey(openKey === section.key ? null : section.key)}
                            >
                                <i className={`fa-solid ${section.icon} me-2`} />
                                {section.label}
                                <i className={`fa-solid fa-chevron-${openKey === section.key ? "up" : "down"} ms-auto`} />
                            </button>

                            {openKey === section.key && section.groups.map((g, i) => (
                                <div key={g.label || i}>
                                    {g.label && <div className="studio-insert-group">{g.label}</div>}
                                    <Tiles items={g.items} onAdd={onAdd} />
                                </div>
                            ))}
                        </section>
                    ))
                )}
            </div>
        </div>
    );
}

function Tiles({ items, onAdd, showSection }) {
    return (
        <div className="studio-insert-grid">
            {items.map((it) => (
                <button
                    key={it.id}
                    type="button"
                    className="studio-insert-tile"
                    title={showSection && it.section ? `${it.label} · ${it.section}` : (it.hint ? `${it.label} — ${it.hint}` : it.label)}
                    // The whole ITEM, not its patch. Some items insert a layer
                    // and some open a picker first — the panel does not need to
                    // know which, and the host already has to.
                    onClick={() => onAdd(it)}
                >
                    {it.tile?.type === "shape" ? <ShapeTile geometry={it.tile.geometry} patch={it.patch} />
                        : it.tile?.type === "path" ? <PathTile d={it.tile.d} />
                            : <i className={`fa-solid ${it.tile?.icon || "fa-plus"} studio-insert-glyph`} />}
                    <span className="studio-insert-label">{it.label}</span>
                </button>
            ))}
        </div>
    );
}

/**
 * An icon tile — the renderer's own path data, as inline SVG.
 *
 * SVG rather than a canvas here, where shapes get a canvas: an icon IS a path,
 * so there is nothing to run and nothing to translate, and SVG stays crisp at
 * any zoom without a device-pixel-ratio dance. It is still the same `d` the
 * renderer paints, so the tile cannot drift from the glyph.
 *
 * `evenodd` matches the painter exactly — without it the ring fills solid and
 * the clock loses its hands, and the tile would advertise the wrong shape.
 */
function PathTile({ d }) {
    return (
        <svg className="studio-insert-shape" width={34} height={34} viewBox="0 0 24 24" aria-hidden="true">
            <path d={d} fill="#cbd5e1" fillRule="evenodd" />
        </svg>
    );
}

/**
 * One geometry, painted at tile size by the real painter.
 *
 * A CALLBACK REF rather than an effect, because the only thing that has to
 * happen is "when this canvas exists, draw on it" — and a callback ref runs
 * again when its identity changes, which `useCallback` ties to the geometry.
 * An effect would need the canvas in a ref and a dependency on it anyway.
 *
 * Drawn at twice the CSS size so it is not soft on a retina screen. It uses the
 * layer's OWN parameters where the tile would otherwise mislead — a `pie` tile
 * drawn at the geometry default is a full circle, and a tile that does not show
 * what the click inserts is worse than no tile.
 */
function ShapeTile({ geometry, patch }) {
    const SIZE = 34;
    const draw = useCallback((node) => {
        if (!node) return;
        const ctx = node.getContext("2d");
        if (!ctx) return;
        const { width, height } = node;
        ctx.clearRect(0, 0, width, height);
        ctx.save();
        ctx.translate(width / 2, height / 2);
        ctx.fillStyle = "#cbd5e1";
        try {
            const params = { ...(patch || {}) };
            // Placement is not silhouette — the tile says what the shape looks
            // like, not where it lands.
            delete params.type; delete params.fx; delete params.fy;
            delete params.fw; delete params.fh;

            /**
             * The tile keeps the layer's PROPORTIONS, because for half of these
             * the proportion is the shape: a square arrow is a chevron, a
             * square rule is a box, and a tile that squares everything makes
             * five geometries look like the same one. Measured against the
             * default vertical frame and through `compileShape`'s own rule that
             * an absent `fh` means square in PIXELS.
             */
            const FW = 1080;
            const FH = 1920;
            const aw = (patch?.fw ?? 0.36) * FW;
            const ah = patch?.fh != null ? patch.fh * FH : aw;
            const box = Math.min(width, height) * 0.78;
            const ratio = ah > 0 ? aw / ah : 1;
            const w = ratio >= 1 ? box : box * ratio;
            const h = ratio >= 1 ? box / ratio : box;

            paintShapeTile(ctx, geometry, w, h, params);
        } catch { /* a tile that will not draw is an empty tile, not a crash */ }
        ctx.restore();
    }, [geometry, patch]);

    return (
        <canvas
            ref={draw}
            width={SIZE * 2}
            height={SIZE * 2}
            style={{ width: SIZE, height: SIZE }}
            className="studio-insert-shape"
        />
    );
}
