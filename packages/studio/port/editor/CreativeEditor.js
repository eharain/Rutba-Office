import { useCallback, useEffect, useRef, useState } from "react";
import { useRouter } from "next/router";
import { useAuth } from "@rutba/ui/context/AuthContext";
import { setMediaAuth, ASPECTS, THEMES, IMAGE_SIZES } from "../../lib/renderer";
import { projects, refusalMessage } from "../../lib/studio-api";
// The layer mutations moved to EditorPanes with the workspace they belong to;
// what is left here is the document as a whole.
import { normaliseDoc, emptyDoc, suggestedTitle } from "../../lib/doc";
import { conversionLosses, deckFromCreative } from "../../lib/deck";
import { useCreative } from "./useCreative";
import EditorPanes from "./EditorPanes";
import ExportBar from "./ExportBar";
import VideoExportBar from "./VideoExportBar";
import MediaPicker from "../MediaPicker";
import ShareBar from "../ShareBar";
import BrandBar from "./BrandBar";
import RecordBar from "./RecordBar";
import Timeline from "./Timeline";
import MoreMenu from "./MoreMenu";
import FlashNote from "./FlashNote";
import { useConfirm } from "../ConfirmDialog";
import { SaveTemplateDialog } from "../TemplateDialogs";

/**
 * The editor screen, for both kinds of creative.
 *
 * The image editor and the video studio are the same screen. What differs is
 * the `kind` this is opened with, which decides the frame control (a still size
 * versus a video aspect) and whether there is a clock — and nothing else,
 * because in this renderer an image genuinely is a one-frame video. Splitting
 * them into two pages would have duplicated the toolbar, the persistence, the
 * layer stack and the inspector to express a difference that is two controls.
 *
 * SAVING IS EXPLICIT, and that is a decision rather than an omission. An
 * autosave here would write a row on every slider tick, and each write is a
 * plan-hash change that invalidates the render dedupe key — so a session of
 * fiddling would queue a fresh render for every intermediate state. The dirty
 * marker and a warning on leaving carry the same safety at none of that cost.
 */
export default function CreativeEditor({ kind, projectId }) {
    const router = useRouter();
    const { jwt } = useAuth();
    const [saving, setSaving] = useState(false);
    const [status, setStatus] = useState(null);
    const [error, setError] = useState(null);
    const [dirty, setDirty] = useState(false);
    const [loadedId, setLoadedId] = useState(projectId || null);
    const [title, setTitle] = useState("");
    const [pickerOpen, setPickerOpen] = useState(false);
    const [initialDoc, setInitialDoc] = useState(() => emptyDoc(kind));
    const [opening, setOpening] = useState(Boolean(projectId));

    // The media proxy checks foreign track URLs against the audio library using
    // the CALLER's credentials, so it needs the session before anything loads.
    useEffect(() => { setMediaAuth({ jwt }); }, [jwt]);

    // Open an existing project. Keyed on the id alone: re-running this on any
    // other change would throw away unsaved edits.
    useEffect(() => {
        if (!projectId) return;
        let cancelled = false;
        setOpening(true);
        projects.get(projectId)
            .then((row) => {
                if (cancelled || !row) return;
                setInitialDoc(normaliseDoc(row.doc, row.kind || kind));
                setTitle(row.title || "");
                setLoadedId(row.id);
            })
            .catch((err) => { if (!cancelled) setError(refusalMessage(err)); })
            .finally(() => { if (!cancelled) setOpening(false); });
        return () => { cancelled = true; };
    }, [projectId, kind]);

    return opening
        ? <div className="p-5 text-center text-muted"><span className="spinner-border spinner-border-sm me-2" />Opening…</div>
        : (
            <EditorBody
                key={loadedId || "new"}
                kind={kind}
                initialDoc={initialDoc}
                loadedId={loadedId}
                setLoadedId={setLoadedId}
                title={title} setTitle={setTitle}
                saving={saving} setSaving={setSaving}
                status={status} setStatus={setStatus}
                error={error} setError={setError}
                dirty={dirty} setDirty={setDirty}
                pickerOpen={pickerOpen} setPickerOpen={setPickerOpen}
                router={router}
            />
        );
}

/**
 * Split out so the whole editor REMOUNTS when a different project is opened.
 *
 * useCreative seeds its state from `initialDoc` once, which is the right shape
 * for an editor — it must not be yanked out from under someone mid-edit by a
 * parent re-render. The consequence is that opening a second project has to be
 * a remount, and `key` is what makes that explicit instead of an effect that
 * tries to detect it.
 */
function EditorBody({
    kind, initialDoc, loadedId, setLoadedId, title, setTitle,
    saving, setSaving, status, setStatus, error, setError, dirty, setDirty,
    pickerOpen, setPickerOpen, router,
}) {
    const creative = useCreative({
        doc: initialDoc,
        kind,
        onDirty: useCallback(() => setDirty(true), [setDirty]),
    });
    // Only what the TOOLBAR and persistence need. Everything the workspace uses
    // is passed through as `creative` — see EditorPanes.
    const { doc, setDoc, frame, images, undo, redo, history, plan } = creative;

    const statusTimer = useRef(null);
    const flash = useCallback((message, isError) => {
        if (isError) { setError(message); setStatus(null); } else { setStatus(message); setError(null); }
        clearTimeout(statusTimer.current);
        statusTimer.current = setTimeout(() => { setStatus(null); setError(null); }, 4000);
    }, [setError, setStatus]);
    useEffect(() => () => clearTimeout(statusTimer.current), []);

    // Leaving with unsaved work. The browser decides the wording; all a page can
    // do is say that there is something to lose.
    useEffect(() => {
        if (!dirty) return undefined;
        const warn = (e) => { e.preventDefault(); e.returnValue = ""; };
        window.addEventListener("beforeunload", warn);
        return () => window.removeEventListener("beforeunload", warn);
    }, [dirty]);

    async function save() {
        setSaving(true);
        try {
            const body = { kind: doc.kind, doc, title: title || suggestedTitle(doc) };
            const row = loadedId ? await projects.update(loadedId, body) : await projects.create(body);
            if (!loadedId && row?.id) {
                setLoadedId(row.id);
                // Replace rather than push: the blank-editor URL is not a place
                // anyone wants the back button to return them to.
                router.replace(`/projects/${row.id}`, undefined, { shallow: true });
            }
            setTitle(row?.title || body.title);
            setDirty(false);
            flash("Saved.");
        } catch (err) {
            flash(refusalMessage(err), true);
        } finally {
            setSaving(false);
        }
    }

    const mutate = (fn, label) => setDoc((d) => fn(d), { label });
    const [saveTplOpen, setSaveTplOpen] = useState(false);
    const [ask, confirmDialog] = useConfirm();

    return (
        <>
            {confirmDialog}
            <div className="studio-toolbar">
                <input
                    className="form-control form-control-sm studio-title-input"
                    placeholder={suggestedTitle(doc)}
                    value={title}
                    onChange={(e) => { setTitle(e.target.value); setDirty(true); }}
                />

                {doc.kind === "video" ? (
                    <select className="form-select form-select-sm w-auto"
                        value={doc.options?.aspect || "vertical"}
                        onChange={(e) => mutate((d) => ({ ...d, options: { ...d.options, aspect: e.target.value } }), "aspect")}>
                        {Object.entries(ASPECTS).map(([k, a]) => <option key={k} value={k}>{a.label}</option>)}
                    </select>
                ) : (
                    <select className="form-select form-select-sm w-auto"
                        value={typeof doc.size === "string" ? doc.size : "square"}
                        onChange={(e) => mutate((d) => ({ ...d, size: e.target.value }), "size")}>
                        {Object.entries(IMAGE_SIZES).map(([k, s]) => <option key={k} value={k}>{s.label}</option>)}
                    </select>
                )}

                <select className="form-select form-select-sm w-auto"
                    value={doc.options?.theme || "dark"}
                    onChange={(e) => mutate((d) => ({ ...d, options: { ...d.options, theme: e.target.value } }), "theme")}>
                    {Object.entries(THEMES).map(([k, t]) => <option key={k} value={k}>{t.label}</option>)}
                </select>

                <BrandBar doc={doc} setDoc={setDoc} onError={(m) => flash(m, true)} />

                {/* Capture lives in the toolbar rather than in the Insert
                    palette: everything in that panel appears instantly on the
                    canvas, and a recording is a thing you go away and MAKE. */}
                <RecordBar
                    creative={creative}
                    onDone={(m) => flash(m)}
                    onError={(m) => flash(m, true)}
                />


                <button type="button" className="btn btn-sm btn-outline-secondary"
                    onClick={() => setPickerOpen(true)}>
                    <i className="fa-regular fa-image me-1" />
                    Pictures{doc.images?.length ? ` (${doc.images.length})` : ""}
                </button>

                <div className="btn-group btn-group-sm">
                    <button type="button" className="btn btn-outline-secondary" title={history.undoLabel ? `Undo ${history.undoLabel}` : "Undo"}
                        disabled={!history.canUndo} onClick={undo}><i className="fa-solid fa-rotate-left" /></button>
                    <button type="button" className="btn btn-outline-secondary" title={history.redoLabel ? `Redo ${history.redoLabel}` : "Redo"}
                        disabled={!history.canRedo} onClick={redo}><i className="fa-solid fa-rotate-right" /></button>
                </div>

                <div className="studio-toolbar-spacer" />

                <span className="text-muted small">{frame.label} · {frame.width}×{frame.height}</span>
                {dirty && <span className="badge text-bg-warning">Unsaved</span>}
                <FlashNote status={status} error={error} />

                <ShareBar projectId={loadedId} kind={doc.kind} />

                {/* Crossing between the kinds lives in the overflow, not the
                    bar. The usability pass counted the toolbar at a dozen-plus
                    controls and wrapping on a laptop — and convert-to-still is
                    a once-per-project act sharing a row with Save, which is a
                    twenty-times-an-hour one. Frequency decides placement. */}
                <MoreMenu>
                    <button type="button" className="dropdown-item"
                        onClick={async () => {
                            const target = doc.kind === "video" ? "image" : "video";
                            const losses = conversionLosses(doc, target);
                            if (losses.length && !(await ask({ title: "Make this a still?", lines: losses, confirmLabel: "Make it a still" }))) return;
                            setDoc((d) => ({ ...d, kind: target }), { label: "convert" });
                        }}>
                        <i className={`fa-solid ${doc.kind === "video" ? "fa-image" : "fa-film"} me-2`} />
                        {doc.kind === "video" ? "Make this a still" : "Make this a video"}
                    </button>
                    <button type="button" className="dropdown-item"
                        onClick={async () => {
                            try {
                                const deckDoc = deckFromCreative(doc, title);
                                const row = await projects.create({ kind: "deck", doc: deckDoc, title: deckDoc.title || "Deck" });
                                // Navigating away with unsaved edits would lose
                                // them without the beforeunload warning firing —
                                // router.push is not a page unload. So the jump
                                // only happens when there is nothing to lose.
                                if (!dirty && row?.id) router.push(`/projects/${row.id}`);
                                else flash("Deck created — it is in your projects. (Staying here: this creative has unsaved changes.)");
                            } catch (err) {
                                flash(refusalMessage(err), true);
                            }
                        }}>
                        <i className="fa-solid fa-rectangle-list me-2" />
                        Make a deck from this
                    </button>
                    <button type="button" className="dropdown-item" onClick={() => setSaveTplOpen(true)}>
                        <i className="fa-solid fa-layer-group me-2" />
                        Save as a template…
                    </button>
                </MoreMenu>

                <button type="button" className="btn btn-sm btn-primary" disabled={saving} onClick={save}>
                    {saving ? <span className="spinner-border spinner-border-sm me-1" /> : <i className="fa-regular fa-floppy-disk me-1" />}
                    Save
                </button>
            </div>

            <EditorPanes creative={creative} />
            {/* The timeline replaced a bare scrub slider. It carries the
                playhead too, so there is one place the time is set rather than
                a slider and a ruler that could disagree. */}
            {doc.kind === "video" && plan && <Timeline creative={creative} />}

            <div className="py-2">
                {doc.kind === "image" ? (
                    <ExportBar
                        doc={doc}
                        images={images}
                        // So a saved still is recorded against its project the
                        // same way a video is, and both kinds can be offered
                        // wherever something asks which creatives have a file.
                        projectId={loadedId}
                        onDone={(m) => flash(m)}
                        onError={(m) => flash(m, true)}
                    />
                ) : (
                    <VideoExportBar
                        creative={creative}
                        projectTitle={title}
                        // So a browser render can be recorded against the
                        // project, not only dropped into the library. Null on an
                        // unsaved creative, and the bar says so rather than
                        // silently skipping it.
                        projectId={loadedId}
                        onDone={(m) => flash(m)}
                        onError={(m) => flash(m, true)}
                    />
                )}
            </div>

            <MediaPicker
                open={pickerOpen}
                onClose={() => setPickerOpen(false)}
                onPick={(urls) => setDoc((d) => ({ ...d, images: [...(d.images || []), ...urls] }), { label: "add pictures" })}
            />

            <SaveTemplateDialog
                open={saveTplOpen}
                // The doc AS EDITED, unsaved changes included: saving a
                // template of what is on screen, not of what was last saved,
                // is what anyone pressing the button means.
                doc={{ ...doc, title: title || doc.title }}
                onClose={() => setSaveTplOpen(false)}
                onSaved={(row) => flash(`Template saved: ${row?.name || "done"}.`)}
                onError={(m) => flash(m, true)}
            />
        </>
    );
}
