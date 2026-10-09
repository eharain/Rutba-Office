import Stage from "./Stage";
import Inspector from "./Inspector";
import LeftRail from "./LeftRail";
import DocSettings from "./DocSettings";
import MusicBedPanel, { MUSIC_BED_ID } from "./MusicBedPanel";
import { BrandPaletteProvider } from "./controls";
import { addLayer, updateLayer, removeLayer, reorderLayer, duplicateLayer, setOption, nextId } from "../../lib/doc";

/**
 * The three-column workspace: layers, canvas, inspector.
 *
 * EXTRACTED WHEN DECKS ARRIVED, and the reason is the interesting part. A deck
 * is a list of creatives, so its editor has to edit ONE OF THEM at a time with
 * the same tools — the same layer stack, the same canvas, the same inspector.
 * The alternative was a second editor for slides, and a second editor is a
 * second drag implementation, a second undo policy and a second set of bugs
 * within a month.
 *
 * WHAT IS DELIBERATELY NOT HERE: the toolbar and saving. Those are exactly what
 * differs between a creative (one document, one Save) and a deck (many slides,
 * one Save for the whole deck, plus slide navigation) — so they stay with each
 * host and this stays the part that is genuinely identical.
 *
 * Everything it needs comes from one `creative`, the object `useCreative`
 * returns. Passing the hook's result rather than fifteen props is what keeps a
 * new capability in the hook from having to be threaded through here.
 */
export default function EditorPanes({ creative }) {
    const {
        doc, setDoc, selectedId, setSelectedId, selectedPatch, selectedLayer,
        frame, selection, attachCanvas, loading, loadError, plan,
        onPointerDown, onPointerMove, onPointerUp,
    } = creative;

    /**
     * WHAT THE THIRD COLUMN SHOWS, decided here rather than inside Inspector.
     *
     * Inspector's own header says it is "everything about the selected layer",
     * and it should stay that. But the column has three jobs, because there are
     * three things that can be selected:
     *
     *   a patch          → Inspector, as before
     *   the music bed    → MusicBedPanel — a real selection with no patch
     *                      behind it, which is exactly why it used to fall
     *                      through to the empty state and do nothing
     *   nothing          → DocSettings, because with nothing selected the thing
     *                      in hand IS the document
     *
     * Putting that decision in the component that already owns selection keeps
     * Inspector from having to know about documents, and keeps a fourth case
     * from turning into a fourth branch inside a panel about layers.
     */
    const setDocOption = (change, opts) => setDoc((d) => setOption(d, change), opts);

    const thirdColumn = selectedId === MUSIC_BED_ID
        ? <MusicBedPanel doc={doc} setOption={setDocOption} duration={plan?.duration} />
        : selectedPatch
            ? (
                <Inspector
                    patch={selectedPatch}
                    layer={selectedLayer}
                    set={(change, opts) => setDoc((d) => updateLayer(d, selectedId, change), opts)}
                    remove={() => { setDoc((d) => removeLayer(d, selectedId), { label: "delete" }); setSelectedId(null); }}
                    duplicate={() => setDoc((d) => duplicateLayer(d, selectedId), { label: "duplicate" })}
                    reorder={(delta) => setDoc((d) => reorderLayer(d, selectedId, delta), { label: "reorder" })}
                />
            )
            : <DocSettings doc={doc} setOption={setDocOption} plan={plan} />;

    return (
        <BrandPaletteProvider colors={doc.brand?.colors}>
        <div className="studio-editor">
            <LeftRail
                doc={doc}
                insertAsk={creative.insertAsk}
                selectedId={selectedId}
                onSelect={setSelectedId}
                onShowSettings={() => setSelectedId(null)}
                /* The new layer arrives SELECTED, so its inspector is already
                   open on the right. The id is minted here rather than read
                   back out of the updater because a state updater may run
                   later than the line after it, and "select what I just added"
                   must not depend on when React chose to flush. addLayer
                   honours an id it is given, so both sides agree by
                   construction rather than by timing. */
                onAdd={(patch) => {
                    const id = patch.id || nextId(doc, patch.type || "layer");
                    setDoc((d) => addLayer(d, { ...patch, id }), { label: `add ${patch.type}` });
                    setSelectedId(id);
                }}
                onToggle={(id, visible) => setDoc((d) => updateLayer(d, id, { visible }), { label: visible ? "show" : "hide" })}
                /* For insert items that transform the document rather than
                   append one layer — a picture joins doc.images, a clip
                   becomes two layers. See insert-catalog. */
                onApply={(fn, label) => setDoc(fn, { label })}
            />

            <Stage
                attachCanvas={attachCanvas}
                frame={frame}
                selection={selection}
                loading={loading}
                loadError={loadError}
                onPointerDown={onPointerDown}
                onPointerMove={onPointerMove}
                onPointerUp={onPointerUp}
            />

            {thirdColumn}
        </div>
        </BrandPaletteProvider>
    );
}
