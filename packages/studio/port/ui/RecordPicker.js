import { useCallback, useEffect, useId, useRef, useState } from "react";

/**
 * Pick a record by its name, and send its document id.
 *
 * The estate keys everything by document id, and for a long time several
 * screens simply asked for one: "Holder document id", with a text box. That is
 * a fair description of the API and an impossible thing to ask a person, who
 * knows the warehouse as "Bay 3" and the technician as "Priya" and has never
 * seen either id in their life. This is the component that closes that gap.
 *
 * What it deliberately does NOT do:
 *
 *  - **It does not load the list.** `search` is called with what has been
 *    typed, so a register with forty thousand rows is searched in the database
 *    rather than dragged into the browser to be filtered there.
 *  - **It does not refuse a raw id.** Pasting one still works and is still
 *    submitted. Somebody debugging, or following a link out of another app,
 *    has the id and not the name, and a picker that would not accept it would
 *    be a step backwards from the text box it replaces.
 *  - **It does not hide what it chose.** The id is shown under the box once
 *    something is picked, because the person filling this in is often the
 *    person who will later be asked which record it was.
 *
 * Props:
 *   value       the document id currently chosen ("" when nothing is)
 *   onChange    called with the new document id, or "" when cleared
 *   search      async (term) => [{ value, label, hint? }] - the caller's
 *               endpoint, so this component knows nothing about any module
 *   label       the field label
 *   placeholder what to type (e.g. "Search people by name")
 *   required    marks the label; does not enforce - the service does that
 *   disabled    ...
 *   minChars    how much to type before searching (default 2; 0 searches on
 *               focus, which suits short lists like branches)
 */
export default function RecordPicker({
    value,
    onChange,
    search,
    label,
    placeholder = "Search…",
    required = false,
    disabled = false,
    minChars = 2,
}) {
    const [term, setTerm] = useState("");
    const [rows, setRows] = useState([]);
    const [open, setOpen] = useState(false);
    const [busy, setBusy] = useState(false);
    const [failed, setFailed] = useState(null);
    const [chosen, setChosen] = useState(null);
    const boxId = useId();
    const wrap = useRef(null);
    // Every search increments this; a reply whose ticket is not the current one
    // is a slower earlier request finishing late, and using it would put stale
    // rows under a newer term.
    const ticket = useRef(0);

    const run = useCallback(async (t) => {
        const mine = (ticket.current += 1);
        if (t.length < minChars) { setRows([]); setBusy(false); return; }
        setBusy(true);
        setFailed(null);
        try {
            const found = await search(t);
            if (ticket.current !== mine) return;
            setRows(Array.isArray(found) ? found.slice(0, 25) : []);
        } catch (err) {
            if (ticket.current !== mine) return;
            setRows([]);
            // A lookup that cannot run is worth saying out loud: without this
            // the box just never suggests anything and reads as "no matches".
            setFailed(err?.message || "that search could not run");
        }
        if (ticket.current === mine) setBusy(false);
    }, [search, minChars]);

    useEffect(() => {
        if (!open) return undefined;
        const id = setTimeout(() => run(term), 250);
        return () => clearTimeout(id);
    }, [term, open, run]);

    // Clicking anywhere else closes the list. Without this the menu survives a
    // click on the button beside it and covers whatever happens next.
    useEffect(() => {
        if (!open) return undefined;
        function away(e) { if (wrap.current && !wrap.current.contains(e.target)) setOpen(false); }
        document.addEventListener("mousedown", away);
        return () => document.removeEventListener("mousedown", away);
    }, [open]);

    function choose(row) {
        setChosen(row);
        setTerm(row.label);
        setOpen(false);
        onChange(row.value);
    }

    function clear() {
        setChosen(null);
        setTerm("");
        setRows([]);
        onChange("");
    }

    return (
        <div ref={wrap} className="position-relative">
            {label && (
                <label className="form-label small mb-0" htmlFor={boxId}>
                    {label}{required && <span className="text-danger ms-1">*</span>}
                </label>
            )}
            <div className="input-group input-group-sm">
                <input
                    id={boxId}
                    className="form-control form-control-sm"
                    placeholder={placeholder}
                    value={term}
                    disabled={disabled}
                    autoComplete="off"
                    onFocus={() => setOpen(true)}
                    onChange={(e) => {
                        const t = e.target.value;
                        setTerm(t);
                        setChosen(null);
                        setOpen(true);
                        // A pasted id is a valid answer, so what is typed is
                        // the value until something is picked from the list.
                        onChange(t);
                    }}
                />
                {(term || value) && !disabled && (
                    <button className="btn btn-outline-secondary" type="button" onClick={clear} title="Clear">
                        <i className="fa fa-times" aria-hidden="true" />
                        <span className="visually-hidden">Clear</span>
                    </button>
                )}
            </div>

            {chosen && (
                <div className="form-text small text-truncate">
                    {chosen.hint ? `${chosen.hint} · ` : ""}<code>{chosen.value}</code>
                </div>
            )}
            {!chosen && value && (
                <div className="form-text small text-truncate">
                    using this as an id: <code>{value}</code>
                </div>
            )}

            {open && (busy || failed || rows.length > 0 || term.length >= minChars) && (
                <div
                    className="list-group position-absolute w-100 shadow-sm"
                    style={{ zIndex: 1000, maxHeight: "16rem", overflowY: "auto" }}
                >
                    {busy && <div className="list-group-item small text-muted">Searching…</div>}
                    {!busy && failed && <div className="list-group-item small text-danger">{failed}</div>}
                    {!busy && !failed && rows.length === 0 && term.length >= minChars && (
                        <div className="list-group-item small text-muted">
                            Nothing matched. What is typed is still sent as an id.
                        </div>
                    )}
                    {!busy && rows.map((r) => (
                        <button
                            key={r.value}
                            type="button"
                            className="list-group-item list-group-item-action small py-1"
                            onClick={() => choose(r)}
                        >
                            <span>{r.label}</span>
                            {r.hint && <span className="text-muted ms-2">{r.hint}</span>}
                        </button>
                    ))}
                </div>
            )}
        </div>
    );
}
