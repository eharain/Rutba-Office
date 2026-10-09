import { useMemo } from "react";
import VideoTimeline from "@rutba/ui/components/VideoTimeline";
import { updateLayer, removeLayer, duplicateLayer, moveLayerToZ } from "../../lib/doc";
import { timelineInsertsFor } from "../../lib/insert-catalog";
import { MUSIC_BED_ID } from "./MusicBedPanel";

/**
 * The video timeline — lanes, trimming, reordering and the playhead.
 *
 * THE COMPONENT IS THE SHARED ONE. `@rutba/ui/components/VideoTimeline` was
 * moved onto the consumer shelf out of the social app in August precisely so a
 * second editor over the same renderer could drive it, and its own header names
 * this app as that second editor. Re-implementing it here would have meant two
 * timelines over one layer envelope, drifting from the first trim bug fixed in
 * only one of them.
 *
 * This file is the seam: the timeline speaks LAYER PATCHES, this app speaks
 * documents, and everything below translates between the two.
 *
 * THE ONE TRANSLATION THAT IS NOT MECHANICAL is reordering. The timeline
 * reports a lane drag as an absolute `z`, because that is what it compared
 * against; this app keeps order as the patch list's order, because the image
 * editor's layer stack moves rows up and down and two competing notions of
 * order is one too many. `moveLayerToZ` converts, and the z is not stored — see
 * the note on it in lib/doc.js.
 */
export default function Timeline({ creative }) {
    const { plan, doc, setDoc, selectedId, setSelectedId, time, setTime, askInsert } = creative;

    // Which lanes are patch-backed, and therefore deletable. The compiled
    // chrome — the caption, the progress bar, the edge fade — has no patch to
    // remove, so offering a delete on it would be a button that does nothing.
    const appendedIds = useMemo(
        () => new Set((doc.patches || []).map((p) => p.id)),
        [doc.patches],
    );

    /**
     * The music bed as a lane.
     *
     * It is not a layer: it lives in `options.audioMode/audioTrackId/...` and is
     * handed to `renderVideo` as its own argument, because the recipe format
     * predates layers and `audioMode: 'random'` — what makes an unattended batch
     * varied rather than twenty copies of one clip — depends on that shape.
     * Showing it as a lane makes it selectable and visible without changing the
     * format; `extraLanes` is display-only and never reaches renderVideo.
     *
     * Selecting it opens `MusicBedPanel`, which is where the bed is actually
     * edited. That panel is why the lane is worth drawing at all — for a while
     * this lane was selectable and selecting it showed "select something",
     * which is the one outcome worse than not drawing it.
     */
    const extraLanes = useMemo(() => {
        const mode = doc.options?.audioMode;
        if (!mode || mode === "none") return [];
        return [{
            id: MUSIC_BED_ID,
            type: "sound",
            name: mode === "random" ? "Music bed (random)" : "Music bed",
            readOnly: true,
        }];
    }, [doc.options]);

    // The add row opens the Insert tab on the section it names, rather than
    // carrying a shorter menu of its own — see timelineInsertsFor.
    const addRow = (
        <div className="btn-group btn-group-sm">
            {timelineInsertsFor(doc.kind).map((door) => (
                <button
                    key={door.section}
                    type="button"
                    className="btn btn-outline-secondary"
                    title={`Add a ${door.label.toLowerCase()} — opens the Insert tab`}
                    onClick={() => askInsert(door.section)}
                >
                    <i className={`fa-solid ${door.icon} me-1`} />
                    {door.label}
                </button>
            ))}
        </div>
    );

    return (
        <VideoTimeline
            plan={plan}
            previewTime={time}
            selectedLayerId={selectedId}
            appendedIds={appendedIds}
            extraLanes={extraLanes}
            addRow={addRow}
            onSelect={setSelectedId}
            onScrub={setTime}
            onPatch={({ id, ...change }) => {
                if (!id) return;
                // A pure reorder. Translated rather than stored — see above.
                if ("z" in change && Object.keys(change).length === 1) {
                    setDoc((d) => moveLayerToZ(d, plan, id, change.z), { label: "reorder" });
                    return;
                }
                // A trim or a move is continuous, so it coalesces into one undo
                // step per gesture the same way a canvas drag does.
                const group = change.timing ? `timing:${id}` : null;
                setDoc((d) => updateLayer(d, id, change), { label: "timing", group });
            }}
            onRemove={(id) => {
                setDoc((d) => removeLayer(d, id), { label: "delete" });
                if (id === selectedId) setSelectedId(null);
            }}
            onDuplicate={(layer) => setDoc((d) => duplicateLayer(d, layer.id), { label: "duplicate" })}
        />
    );
}
