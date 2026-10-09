
const ICONS = {
    wordart: "fa-wand-magic-sparkles",
    textblock: "fa-align-left",
    chart: "fa-chart-simple",
    icon: "fa-icons",
    shape: "fa-shapes",
    adjust: "fa-sliders",
    text: "fa-font",
    image: "fa-image",
    qr: "fa-qrcode",
    photo: "fa-image",
    video: "fa-film",
    sound: "fa-music",
};

/**
 * The layer stack, and the row of things you can add to it.
 *
 * LISTED TOP-DOWN, which is the reverse of the patch list. The patch list is in
 * paint order — later entries land on top — and every graphics tool ever made
 * shows the topmost layer first. Reversing here rather than storing it reversed
 * keeps the document in the order the renderer wants it, which is the order
 * that has to be right.
 *
 * ONLY PATCHED LAYERS APPEAR. A compiled plan also contains the caption, the
 * photo lanes and the chrome, which are configured from the options panel and
 * cannot be reordered or deleted — putting them here would offer a delete
 * button that quietly does nothing.
 */
export default function LayerStack({ doc, selectedId, onSelect, onToggle, onShowSettings, onOpenInsert }) {
    const rows = [...(doc.patches || [])].reverse();

    return (
        <div className="studio-layers">
            {/* Deselecting by clicking empty canvas is the gesture people
                already have, but it is not a DISCOVERABLE one — nothing on
                screen says the settings are behind it. This row does. */}
            <button
                type="button"
                className={`studio-doc-row${selectedId == null ? " is-selected" : ""}`}
                onClick={onShowSettings}
            >
                <i className="fa-solid fa-sliders studio-layer-icon" />
                <span className="studio-layer-name">
                    {doc.kind === "video" ? "Video settings" : "Image settings"}
                </span>
            </button>

            {/* The add row moved to the Insert tab when it outgrew a row — see
                InsertPanel. What is left here is the way back to it, because a
                layer list with no visible way to add a layer is a dead end. */}
            <div className="studio-add-row">
                <button type="button" className="btn btn-sm btn-outline-info w-100" onClick={onOpenInsert}>
                    <i className="fa-solid fa-plus me-1" />
                    Add something
                </button>
            </div>

            <div className="studio-layer-list">
                {rows.length === 0 && (
                    <p className="text-muted small px-2 py-3 mb-0">
                        Nothing added yet — open <strong>Insert</strong> and pick something.
                    </p>
                )}
                {rows.map((p) => (
                    <div
                        key={p.id}
                        className={`studio-layer${p.id === selectedId ? " is-selected" : ""}${p.visible === false ? " is-hidden" : ""}`}
                        onClick={() => onSelect(p.id)}
                        role="button"
                        tabIndex={0}
                        onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); onSelect(p.id); } }}
                    >
                        <i className={`fa-solid ${ICONS[p.type] || "fa-layer-group"} studio-layer-icon`} />
                        <span className="studio-layer-name">
                            {p.name || p.text || labelFor(p)}
                        </span>
                        {/* The template-slot mark, visible where layers are
                            scanned. An author deciding which layers to offer
                            needs the decision at a glance, not one inspector
                            visit per layer. */}
                        {p.slot && (
                            <i className="fa-solid fa-square-pen studio-layer-icon"
                                title={`Template slot: ${p.slot.label || ""}`} />
                        )}
                        <button
                            type="button"
                            className="btn btn-sm btn-link studio-layer-eye"
                            title={p.visible === false ? "Show" : "Hide"}
                            onClick={(e) => { e.stopPropagation(); onToggle(p.id, p.visible === false); }}
                        >
                            <i className={`fa-regular ${p.visible === false ? "fa-eye-slash" : "fa-eye"}`} />
                        </button>
                    </div>
                ))}
            </div>
        </div>
    );
}

function labelFor(patch) {
    if (patch.type === "shape") return patch.geometry || "shape";
    if (patch.type === "adjust") return patch.preset || "adjustment";
    return patch.type;
}
