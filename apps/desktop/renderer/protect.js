// Password-protected documents, the window's side: the password asked for
// when a protected file opens, File → Info with its Protect card, and the
// Encrypt with Password dialog. Shared by Rutba Word, Worksheets and
// Presentation, which differ only in their words ("document", "workbook",
// "presentation") and in what else their Protect menu offers.
//
// A password typed here goes straight to the document service in the main
// process and is not kept by the window: the service holds it in memory for
// the open document, and nothing writes it anywhere.

import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Dialog, Button, Icon, Spinner } from '@rutba/office-ui';

/** What each app calls the thing it edits. */
export const NOUN = { word: 'document', sheets: 'workbook', slides: 'presentation' };
const PROTECT_LABEL = { word: 'Protect Document', sheets: 'Protect Workbook', slides: 'Protect Presentation' };

/* ── a password box with a show/hide toggle ──────────────────────────────── */

export function SecretInput({ value, onChange, className = '', invalid = false, autoFocus = false, onKeyDown, inputRef, placeholder, label }) {
  const [shown, setShown] = useState(false);
  const own = useRef(null);
  const ref = inputRef || own;
  const [caps, setCaps] = useState(false);
  const watchCaps = (e) => {
    if (typeof e.getModifierState === 'function') setCaps(e.getModifierState('CapsLock'));
  };
  return (
    <div className="pw-secret-wrap">
      <div className={`pw-secret${invalid ? ' invalid' : ''}`}>
        <input
          ref={ref}
          className={`rw-input pw-secret-input ${className}`}
          type={shown ? 'text' : 'password'}
          value={value}
          spellCheck={false}
          autoComplete="off"
          aria-label={label}
          aria-invalid={invalid || undefined}
          placeholder={placeholder}
          autoFocus={autoFocus}
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={(e) => {
            watchCaps(e);
            onKeyDown?.(e);
          }}
          onKeyUp={watchCaps}
        />
        <button
          type="button"
          className="pw-reveal"
          data-tip={shown ? 'Hide password' : 'Show password'}
          aria-label={shown ? 'Hide password' : 'Show password'}
          aria-pressed={shown}
          tabIndex={-1}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            setShown((s) => !s);
            ref.current?.focus();
          }}
        >
          <Icon name={shown ? 'eyeOff' : 'eye'} size={15} />
        </button>
      </div>
      {caps ? <div className="pw-caps">Caps Lock is on</div> : null}
    </div>
  );
}

/* ── opening a protected file ───────────────────────────────────────────── */

/**
 * The Password dialog a protected file opens with: the file's name, a
 * masked box with a show/hide toggle, Cancel and OK, and — after a wrong
 * try — the box emptied, marked, and a sentence saying so. While the
 * password is being checked (100,000 rounds of hashing take a moment) OK
 * says so and nothing can be pressed twice.
 */
export function PasswordPrompt({ name, wrong = false, attempt = 0, checking = false, onCancel, onSubmit }) {
  const [password, setPassword] = useState('');
  const input = useRef(null);
  useEffect(() => {
    if (!attempt) return;
    setPassword('');
    input.current?.focus();
  }, [attempt]);
  const submit = () => {
    if (checking || !password) return;
    onSubmit(password);
  };
  return (
    <Dialog
      title="Password"
      width={440}
      onClose={checking ? undefined : onCancel}
      actions={
        <>
          <Button label="Cancel" className="pw-open-cancel" disabled={checking} onClick={onCancel} />
          <Button primary className="pw-open-ok" disabled={checking || !password} onClick={submit} label={checking ? 'Opening…' : 'OK'}>
            {checking ? <Spinner style={{ width: 13, height: 13, borderWidth: 2 }} /> : null}
          </Button>
        </>
      }
    >
      <style>{PROTECT_CSS}</style>
      <div className="pw-prompt">
        <div className="pw-prompt-head">
          <span className="pw-badge"><Icon name="lock" size={18} /></span>
          <div className="pw-prompt-words">
            <div className="pw-file" data-tip={name}>{name}</div>
            <div className="pw-sub">is protected. Enter the password to open it.</div>
          </div>
        </div>
        <label className="pw-label" htmlFor="pw-open-password">Password</label>
        <SecretInput
          inputRef={input}
          className="pw-open-password"
          label="Password"
          value={password}
          invalid={wrong && !password}
          autoFocus
          onChange={setPassword}
          onKeyDown={(e) => {
            if (e.key === 'Enter') submit();
          }}
        />
        {wrong ? (
          <div className="pw-error" role="alert">
            <Icon name="info" size={14} />
            <span>That password is not right — passwords are case-sensitive.</span>
          </div>
        ) : null}
      </div>
    </Dialog>
  );
}

/**
 * The window's password gate: `ask` shows the Password dialog and resolves
 * with what was typed (or null for Cancel); the dialog stays up, marked as
 * checking, until the next `ask` or `done`.
 */
export function usePasswordGate() {
  const [state, setState] = useState(null);
  const resolver = useRef(null);
  const ask = useCallback(
    ({ name, wrong }) =>
      new Promise((resolve) => {
        resolver.current = resolve;
        setState((s) => ({ name, wrong: Boolean(wrong), checking: false, attempt: (s?.attempt || 0) + (wrong ? 1 : 0) }));
      }),
    []
  );
  const done = useCallback(() => {
    resolver.current = null;
    setState(null);
  }, []);
  const node = state ? (
    <PasswordPrompt
      {...state}
      onCancel={() => {
        const r = resolver.current;
        resolver.current = null;
        setState(null);
        r?.(null);
      }}
      onSubmit={(password) => {
        setState((s) => (s ? { ...s, checking: true } : s));
        const r = resolver.current;
        resolver.current = null;
        r?.(password);
      }}
    />
  ) : null;
  return { ask, done, node };
}

/**
 * Open through the gate: `call(password)` is the service's open (or
 * recover), first without a password. A protected file answers `locked`,
 * and the gate asks until the password opens it or the person cancels —
 * which throws an error whose `locked` names the file.
 */
export async function openProtected(call, gate) {
  let result = await call(null);
  try {
    while (result?.locked) {
      const password = await gate.ask({ name: result.name, wrong: result.wrong });
      if (password == null) {
        const err = new Error(`${result.name} is protected with a password, and it stays closed until the password is given.`);
        err.locked = result.name;
        throw err;
      }
      result = await call(password);
    }
  } finally {
    gate.done();
  }
  return result;
}

/** The error screen's way back to the Password dialog. */
export function LockedAction() {
  return <Button primary icon="lock" label="Enter password…" className="pw-retry" onClick={() => window.location.reload()} />;
}

/* ── File → Info ────────────────────────────────────────────────────────── */

function formatSize(n) {
  if (!(n > 0)) return '—';
  if (n < 1024) return `${n} bytes`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(n < 10240 ? 1 : 0)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

/**
 * File → Info, as Office's Info page lays it out: the file's name and
 * where it lives, a Protect card whose menu holds Encrypt with Password
 * (and whatever else the app protects), and the file's properties beside
 * it. A protected file's card turns amber and says what protects it.
 *
 * `items`: [{ id, icon, label, detail, run }] — the Protect menu.
 * `notes`: sentences the card shows when something protects the file.
 */
export function InfoDialog({ app, doc, size = null, items = [], notes = [], properties = [], onClose }) {
  const noun = NOUN[app] || 'document';
  const [open, setOpen] = useState(false);
  const menuRef = useRef(null);
  const popRef = useRef(null);
  // The menu hangs from the tile but is drawn on the page itself, so the
  // dialog's scrolling body cannot clip it.
  const [at, setAt] = useState(null);
  useLayoutEffect(() => {
    if (!open) return;
    const r = menuRef.current?.getBoundingClientRect();
    if (r) setAt({ left: r.left, top: r.bottom + 6 });
  }, [open]);
  useEffect(() => {
    if (!open) return undefined;
    const off = (e) => {
      if (menuRef.current?.contains(e.target) || popRef.current?.contains(e.target)) return;
      setOpen(false);
    };
    const key = (e) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', off, true);
    window.addEventListener('keydown', key, true);
    return () => {
      document.removeEventListener('mousedown', off, true);
      window.removeEventListener('keydown', key, true);
    };
  }, [open]);
  const said = [...(doc?.encrypted ? [`A password is required to open this ${noun}.`] : []), ...notes];
  const label = PROTECT_LABEL[app] || 'Protect Document';
  const folder = doc?.path ? doc.path.replace(/[\\/][^\\/]*$/, '') : null;
  return (
    <Dialog title="Info" width={700} onClose={onClose} actions={<Button primary label="Close" className="pw-info-close" onClick={onClose} />}>
      <style>{PROTECT_CSS}</style>
      <div className="pw-info">
        <div className="pw-info-file">
          <Icon name={app === 'sheets' ? 'sheets' : app === 'slides' ? 'slides' : 'word'} size={22} />
          <div>
            <div className="pw-info-name">{doc?.name || '—'}</div>
            <div className="pw-info-path" data-tip={doc?.path || ''}>{folder || 'Not saved yet'}</div>
          </div>
        </div>
        <div className="pw-info-cols">
          <div className="pw-info-main">
            <div className={`pw-card${said.length ? ' on' : ''}`} data-card="protect">
              <div className="pw-card-tile-wrap" ref={menuRef}>
                <button
                  type="button"
                  className={`pw-card-tile${open ? ' open' : ''}`}
                  data-tip={`${label} — choose how this ${noun} is protected`}
                  aria-haspopup="menu"
                  aria-expanded={open}
                  onClick={() => setOpen((o) => !o)}
                >
                  <Icon name="lock" size={26} />
                  <span>{label}</span>
                  <Icon name="chevronDown" size={12} />
                </button>
                {open && at ? createPortal(
                  <div className="pw-menu" role="menu" ref={popRef} style={{ left: at.left, top: at.top }}>
                    <style>{PROTECT_CSS}</style>
                    {items.map((item) => (
                      <button
                        key={item.id}
                        type="button"
                        role="menuitem"
                        className="pw-menu-item"
                        data-item={item.id}
                        disabled={item.disabled}
                        onClick={() => {
                          setOpen(false);
                          item.run();
                        }}
                      >
                        <span className="pw-menu-icon"><Icon name={item.icon || 'lock'} size={18} /></span>
                        <span className="pw-menu-words">
                          <span className="pw-menu-label">{item.label}</span>
                          {item.detail ? <span className="pw-menu-detail">{item.detail}</span> : null}
                        </span>
                      </button>
                    ))}
                  </div>,
                  document.body
                ) : null}
              </div>
              <div className="pw-card-words">
                <h4>{label}</h4>
                {said.length ? (
                  <ul className="pw-card-notes">
                    {said.map((line) => (
                      <li key={line}>
                        <Icon name={line.startsWith('A password') ? 'lock' : 'shield'} size={13} />
                        <span>{line}</span>
                      </li>
                    ))}
                  </ul>
                ) : (
                  <p>Control what types of changes people can make to this {noun}.</p>
                )}
              </div>
            </div>
          </div>
          <aside className="pw-info-props">
            <h5>Properties</h5>
            <dl>
              <dt>Size</dt>
              <dd>{formatSize(size)}</dd>
              <dt>Type</dt>
              <dd>{(doc?.path || doc?.name || '').split('.').pop().toUpperCase() || '—'}</dd>
              {properties.map(([k, v]) => (
                <React.Fragment key={k}>
                  <dt>{k}</dt>
                  <dd>{v}</dd>
                </React.Fragment>
              ))}
              <dt>Encryption</dt>
              <dd className="pw-info-encryption" data-tip={doc?.encrypted ? 'Agile Encryption: AES-256 with SHA-512, as Office 2013 onwards writes it' : undefined}>{doc?.encrypted ? 'AES-256' : 'None'}</dd>
              <dt>Unsaved changes</dt>
              <dd>{doc?.dirty ? 'Yes' : 'No'}</dd>
            </dl>
          </aside>
        </div>
      </div>
    </Dialog>
  );
}

/* ── Encrypt with Password ──────────────────────────────────────────────── */

/**
 * Encrypt Document: the password, typed twice, and Office's caution that a
 * lost password cannot be recovered. On a document that has one already,
 * both boxes left empty take it off again.
 */
export function EncryptDialog({ app, encrypted = false, onClose, onSet }) {
  const noun = NOUN[app] || 'document';
  const [password, setPassword] = useState('');
  const [again, setAgain] = useState('');
  const [busy, setBusy] = useState(false);
  const clearing = encrypted && !password && !again;
  const mismatch = again !== '' && password !== again;
  const ok = !busy && (clearing || (password !== '' && password === again));
  const submit = async () => {
    if (!ok) return;
    setBusy(true);
    try {
      await onSet(clearing ? '' : password);
    } finally {
      setBusy(false);
    }
  };
  const enter = (e) => {
    if (e.key === 'Enter') submit();
  };
  return (
    <Dialog
      title="Encrypt Document"
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" className="pw-set-cancel" onClick={onClose} />
          <Button primary label={clearing ? 'Remove password' : 'OK'} className="pw-set-ok" disabled={!ok} onClick={submit} />
        </>
      }
    >
      <style>{PROTECT_CSS}</style>
      <div className="pw-set">
        <p className="pw-set-lead">Encrypt the contents of this file. It will open only with the password.</p>
        {encrypted ? (
          <p className="pw-set-current">
            <Icon name="lock" size={13} />
            <span>This {noun} already needs a password to open. Type a new one to change it, or leave both boxes empty to remove it.</span>
          </p>
        ) : null}
        <label className="pw-label">Password</label>
        <SecretInput className="pw-set-password" label="Password" value={password} onChange={setPassword} onKeyDown={enter} autoFocus />
        <label className="pw-label">Reenter password</label>
        <SecretInput className="pw-set-again" label="Reenter password" value={again} onChange={setAgain} onKeyDown={enter} invalid={mismatch} />
        {mismatch ? (
          <div className="pw-error pw-set-mismatch" role="alert">
            <Icon name="info" size={14} />
            <span>The two passwords are not the same.</span>
          </div>
        ) : null}
        <div className="pw-caution">
          <Icon name="shield" size={16} />
          <span>
            <b>Caution:</b> if you lose or forget the password, it cannot be recovered — nobody can open the {noun} without it. Keep a list of passwords and the files they open somewhere safe. Passwords are case-sensitive.
          </span>
        </div>
      </div>
    </Dialog>
  );
}

/**
 * File → Info and Encrypt with Password for one window. `items` are the
 * app's own entries in the Protect menu after Encrypt with Password; each
 * closes Info before it runs. Answers the app menu's Info entry, a way to
 * open Info, and the dialogs to render.
 */
export function useProtection({ app, shell, doc, setDoc, toast, items = [], notes = [], properties = [] }) {
  const [view, setView] = useState(null);
  const [size, setSize] = useState(null);
  const noun = NOUN[app] || 'document';
  const path = doc?.path || null;
  const openInfo = useCallback(() => {
    setView('info');
    setSize(null);
    if (path) shell.fs.stat({ path }).then((s) => setSize(s?.size ?? null)).catch(() => {});
  }, [shell, path]);
  const encrypt = {
    id: 'encrypt',
    icon: 'lock',
    label: 'Encrypt with Password',
    detail: `Require a password to open this ${noun}.`,
    run: () => setView('encrypt'),
  };
  const menu = [
    encrypt,
    ...items.map((item) => ({
      ...item,
      run: () => {
        setView(null);
        item.run();
      },
    })),
  ];
  const setPassword = async (password) => {
    if (!doc?.id) return;
    try {
      const meta = await shell.doc.setPassword({ id: doc.id, password });
      setDoc((d) => ({ ...d, ...meta }));
      toast(password ? `Encrypted with a password. Save to protect the file on disk.` : `Password removed. Save to write the ${noun} without it.`, { tone: 'good', ms: 5200 });
      setView('info');
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  };
  const node =
    view === 'info' ? (
      <InfoDialog app={app} doc={doc} size={size} items={menu} notes={notes} properties={properties} onClose={() => setView(null)} />
    ) : view === 'encrypt' ? (
      <EncryptDialog app={app} encrypted={Boolean(doc?.encrypted)} onClose={() => setView('info')} onSet={setPassword} />
    ) : null;
  return { openInfo, node, menuItem: { label: 'Info', icon: 'info', run: openInfo } };
}

export const PROTECT_CSS = `
.pw-prompt { display: flex; flex-direction: column; gap: 6px; padding: 2px 0 12px; }
.pw-prompt-head { display: flex; align-items: center; gap: 12px; margin-bottom: 10px; }
.pw-badge { flex: none; width: 38px; height: 38px; border-radius: 10px; display: grid; place-items: center; background: var(--accent-soft); color: var(--accent); }
.pw-prompt-words { min-width: 0; }
.pw-file { font-weight: 600; font-size: 13.5px; color: var(--ink); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 340px; }
.pw-sub { font-size: 12.5px; color: var(--ink-2); margin-top: 1px; }
.pw-label { font-size: 11.5px; color: var(--ink-2); font-weight: 600; margin-top: 4px; }
.pw-secret-wrap { display: flex; flex-direction: column; gap: 4px; }
.pw-secret { position: relative; display: flex; }
.pw-secret .pw-secret-input { flex: 1; padding-right: 36px; letter-spacing: 0.02em; }
.pw-secret.invalid .pw-secret-input { border-color: var(--bad); box-shadow: 0 0 0 3px color-mix(in srgb, var(--bad) 16%, transparent); }
.pw-reveal { position: absolute; right: 4px; top: 50%; transform: translateY(-50%); width: 28px; height: 26px; border: 0; border-radius: 6px; background: transparent; color: var(--ink-3); display: grid; place-items: center; cursor: pointer; }
.pw-reveal:hover { background: var(--hover); color: var(--ink); }
.pw-reveal[aria-pressed="true"] { color: var(--accent); }
.pw-caps { font-size: 11.5px; color: var(--warn); }
.pw-error { display: flex; align-items: flex-start; gap: 7px; margin-top: 6px; padding: 8px 10px; border-radius: 8px; font-size: 12.5px; line-height: 1.4; color: var(--bad); background: color-mix(in srgb, var(--bad) 9%, var(--surface)); border: 1px solid color-mix(in srgb, var(--bad) 28%, transparent); }
.pw-error svg { flex: none; margin-top: 1px; }
.rw-dialog-foot .pw-open-ok { display: inline-flex; align-items: center; gap: 7px; min-width: 76px; justify-content: center; }
.rw-dialog-foot .pw-open-ok .rw-spinner { border-color: color-mix(in srgb, var(--ink-on-accent) 35%, transparent); border-top-color: var(--ink-on-accent); }

.pw-set { display: flex; flex-direction: column; gap: 6px; padding: 0 0 12px; }
.pw-set-lead { margin: 0 0 4px; font-size: 12.5px; color: var(--ink-2); line-height: 1.5; }
.pw-set-current { display: flex; gap: 7px; align-items: flex-start; margin: 0 0 4px; padding: 8px 10px; border-radius: 8px; font-size: 12.5px; line-height: 1.45; color: var(--ink); background: color-mix(in srgb, var(--warn) 12%, var(--surface)); border: 1px solid color-mix(in srgb, var(--warn) 30%, transparent); }
.pw-set-current svg { flex: none; margin-top: 2px; color: var(--warn); }
.pw-caution { display: flex; gap: 9px; align-items: flex-start; margin-top: 10px; padding: 10px 12px; border-radius: 8px; font-size: 12px; line-height: 1.5; color: var(--ink-2); background: var(--surface-2); border: 1px solid var(--line-soft); }
.pw-caution svg { flex: none; color: var(--warn); margin-top: 1px; }
.pw-caution b { color: var(--ink); font-weight: 600; }

.pw-info { display: flex; flex-direction: column; gap: 16px; padding: 2px 0 14px; }
.pw-info-file { display: flex; align-items: center; gap: 12px; padding-bottom: 14px; border-bottom: 1px solid var(--line-soft); }
.pw-info-file > svg { flex: none; color: var(--accent); }
.pw-info-name { font-family: var(--font-display); font-size: 17px; font-weight: 600; color: var(--ink); letter-spacing: -0.01em; }
.pw-info-path { font-size: 12px; color: var(--ink-3); margin-top: 2px; word-break: break-all; }
.pw-info-cols { display: grid; grid-template-columns: minmax(0, 1fr) 210px; gap: 22px; align-items: start; }
.pw-card { display: flex; gap: 16px; align-items: flex-start; padding: 14px; border-radius: 12px; border: 1px solid var(--line-soft); background: var(--surface); }
.pw-card.on { background: color-mix(in srgb, var(--warn) 11%, var(--surface)); border-color: color-mix(in srgb, var(--warn) 34%, transparent); }
.pw-card-tile-wrap { position: relative; flex: none; }
.pw-card-tile { width: 96px; height: 92px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 5px; padding: 8px 6px; border-radius: 10px; border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; font-size: 11.5px; line-height: 1.25; text-align: center; cursor: pointer; transition: background var(--fast), border-color var(--fast); }
.pw-card-tile > svg:first-child { color: var(--accent); }
.pw-card-tile > svg:last-child { color: var(--ink-3); }
.pw-card-tile:hover, .pw-card-tile.open { background: var(--hover); border-color: var(--line-strong); }
.pw-card.on .pw-card-tile > svg:first-child { color: var(--warn); }
.pw-card-words { min-width: 0; padding-top: 2px; }
.pw-card-words h4 { margin: 0 0 6px; font-family: var(--font-display); font-size: 14.5px; font-weight: 600; color: var(--ink); }
.pw-card-words p { margin: 0; font-size: 12.5px; color: var(--ink-2); line-height: 1.5; }
.pw-card-notes { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 5px; }
.pw-card-notes li { display: flex; gap: 7px; align-items: flex-start; font-size: 12.5px; line-height: 1.45; color: var(--ink); }
.pw-card-notes svg { flex: none; margin-top: 2px; color: var(--warn); }
.pw-menu { position: fixed; z-index: 130; width: 330px; padding: 6px; border-radius: 12px; border: 1px solid var(--line); background: var(--surface); box-shadow: var(--shadow-3); animation: rw-pop 120ms var(--ease); }
.pw-menu-item { display: flex; gap: 12px; align-items: flex-start; width: 100%; padding: 9px 10px; border: 0; border-radius: 8px; background: transparent; color: var(--ink); font: inherit; text-align: left; cursor: pointer; }
.pw-menu-item:hover:not(:disabled) { background: var(--hover); }
.pw-menu-item:disabled { opacity: 0.5; cursor: default; }
.pw-menu-icon { flex: none; width: 32px; height: 32px; border-radius: 8px; display: grid; place-items: center; background: var(--accent-soft); color: var(--accent); }
.pw-menu-words { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
.pw-menu-label { font-weight: 600; font-size: 12.5px; }
.pw-menu-detail { font-size: 11.5px; color: var(--ink-2); line-height: 1.4; }
.pw-info-props h5 { margin: 0 0 8px; font-size: 11px; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; color: var(--ink-3); }
.pw-info-props dl { display: grid; grid-template-columns: auto 1fr; gap: 7px 12px; margin: 0; font-size: 12px; }
.pw-info-props dt { color: var(--ink-3); white-space: nowrap; }
.pw-info-props dd { margin: 0; color: var(--ink); min-width: 0; overflow-wrap: anywhere; }
`;
