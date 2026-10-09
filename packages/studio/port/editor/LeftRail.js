import { useEffect, useState } from "react";
import LayerStack from "./LayerStack";
import InsertPanel from "./InsertPanel";
import MediaPicker from "../MediaPicker";

/**
 * The left column: what is on the creative, and what can be put on it.
 *
 * TWO TABS OVER ONE COLUMN, rather than a second column or a dialog. The editor
 * is a fixed three-column grid bounded to the viewport — the canvas gets
 * whatever the two rails leave, and a fourth region would come straight out of
 * the thing everybody is actually looking at. Tabbing costs nothing and the
 * canvas is the same size whichever tab is showing.
 *
 * INSERTING SWITCHES BACK TO LAYERS, and that is deliberate rather than tidy.
 * The thing you just added is selected, its inspector is open on the right, and
 * the next move is almost always to adjust it — so leaving the palette up would
 * mean a manual switch after every single insert. Adding several in a row is
 * still one click each: the Insert tab is one click away and it remembers which
 * section was open.
 */
export default function LeftRail({ doc, insertAsk, selectedId, onSelect, onAdd, onApply, onToggle, onShowSettings }) {
    /**
     * A BLANK creative opens on Insert; one with content opens on Layers.
     *
     * The usability pass found the blank case opening onto an empty layer list
     * whose only content was a sentence telling you to go to the other tab —
     * a door that opens onto a sign pointing at the next door. Computed once at
     * mount (the editor remounts per project), so it decides where you START
     * and never yanks the tab away mid-session.
     */
    const [tab, setTab] = useState(() => (
        (doc.patches?.length || 0) + (doc.images?.length || 0) > 0 ? "layers" : "insert"
    ));
    /**
     * The picker lives HERE rather than in the host, because the rail is what
     * owns Insert and both editors mount the rail. Putting it in CreativeEditor
     * would have meant the deck editor either duplicating it or going without.
     *
     * It holds the ITEM, not just the kind: the item carries the document
     * transform to run once something is chosen, so the picker's `onPick` does
     * not have to know what a clip becomes.
     */
    const [picking, setPicking] = useState(null);

    // Somewhere else asked for the Insert tab (the timeline's add row does):
    // open it, on the section asked for — see useCreative's askInsert.
    useEffect(() => {
        if (insertAsk) setTab("insert");
    }, [insertAsk]);

    return (
        <div className="studio-rail">
            <div className="studio-rail-tabs" role="tablist">
                <button
                    type="button" role="tab" aria-selected={tab === "layers"}
                    className={`studio-rail-tab${tab === "layers" ? " is-active" : ""}`}
                    onClick={() => setTab("layers")}
                >
                    <i className="fa-solid fa-layer-group me-1" />
                    Layers
                    {doc.patches?.length ? <span className="studio-rail-count">{doc.patches.length}</span> : null}
                </button>
                <button
                    type="button" role="tab" aria-selected={tab === "insert"}
                    className={`studio-rail-tab${tab === "insert" ? " is-active" : ""}`}
                    onClick={() => setTab("insert")}
                >
                    <i className="fa-solid fa-plus me-1" />
                    Insert
                </button>
            </div>

            {tab === "layers" ? (
                <LayerStack
                    doc={doc}
                    selectedId={selectedId}
                    onSelect={onSelect}
                    onToggle={onToggle}
                    onShowSettings={onShowSettings}
                    onOpenInsert={() => setTab("insert")}
                />
            ) : (
                <InsertPanel
                    kind={doc.kind}
                    ask={insertAsk}
                    onAdd={(item) => {
                        // Two shapes of item — see insert-catalog. One inserts a
                        // layer; the other has to ask what to insert first, and
                        // must NOT switch back to Layers yet or the picker would
                        // open behind a panel that just changed under it.
                        if (item.pick) { setPicking(item); return; }
                        onAdd(item.patch);
                        setTab("layers");
                    }}
                />
            )}

            <MediaPicker
                open={Boolean(picking)}
                kind={picking?.pick || "image"}
                onClose={() => setPicking(null)}
                onPick={(urls) => {
                    const item = picking;
                    setPicking(null);
                    if (!item || !urls.length) return;
                    onApply?.((d) => item.apply(d, urls), `add ${item.label.toLowerCase()}`);
                    setTab("layers");
                }}
            />
        </div>
    );
}
