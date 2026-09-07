// The shared shell: everything seven apps have in common, and nothing they
// don't. A window frame with its own caption, a ribbon, panels, dialogs, menus
// and toasts — all of it unaware of documents, mail or media.

import React, { createContext, useContext, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from './icons.js';

export { Icon };
export * from './commands.js';

/* ── theme ──────────────────────────────────────────────────────────────── */

const ThemeContext = createContext({ dark: false, setTheme: () => {} });
export const useTheme = () => useContext(ThemeContext);

/**
 * Theme follows the operating system until somebody chooses otherwise, and then
 * the choice sticks. Both facts are stored, because "system" is a setting too.
 */
export function ThemeProvider({ app, initial = 'system', onChange, children }) {
  const [mode, setMode] = useState(initial);
  const [systemDark, setSystemDark] = useState(
    () => typeof matchMedia === 'function' && matchMedia('(prefers-color-scheme: dark)').matches
  );

  useEffect(() => {
    if (typeof matchMedia !== 'function') return undefined;
    const mq = matchMedia('(prefers-color-scheme: dark)');
    const listen = (e) => setSystemDark(e.matches);
    mq.addEventListener('change', listen);
    return () => mq.removeEventListener('change', listen);
  }, []);

  const dark = mode === 'system' ? systemDark : mode === 'dark';

  useEffect(() => {
    const root = document.documentElement;
    root.dataset.theme = dark ? 'dark' : 'light';
    if (app) root.dataset.app = app;
  }, [dark, app]);

  const setTheme = useCallback(
    (next) => {
      setMode(next);
      onChange?.(next);
    },
    [onChange]
  );

  const value = useMemo(() => ({ dark, mode, setTheme }), [dark, mode, setTheme]);
  return <ThemeContext.Provider value={value}>{children}</ThemeContext.Provider>;
}

/* ── toasts ─────────────────────────────────────────────────────────────── */

const ToastContext = createContext(() => {});
export const useToast = () => useContext(ToastContext);

export function ToastProvider({ children }) {
  const [items, setItems] = useState([]);
  const seq = useRef(0);

  const push = useCallback((message, { tone = 'plain', ms = 3200, action } = {}) => {
    const id = ++seq.current;
    setItems((list) => {
      // The same thing going wrong fifteen times is one problem, not fifteen.
      // A repeat is counted on the notice already showing rather than stacked
      // behind it, which is what turns a failing shortcut into a wall of red.
      const at = list.findIndex((t) => t.message === message && t.tone === tone);
      if (at >= 0) {
        const existing = list[at];
        const next = list.slice();
        next[at] = { ...existing, repeats: (existing.repeats || 1) + 1 };
        return next;
      }
      return [...list, { id, message, tone, action, repeats: 1 }];
    });
    if (ms) setTimeout(() => setItems((list) => list.filter((t) => t.id !== id)), ms);
    return id;
  }, []);

  return (
    <ToastContext.Provider value={push}>
      {children}
      <div className="rw-toast-host">
        {items.map((t) => (
          <div key={t.id} className={`rw-toast ${t.tone === 'plain' ? '' : t.tone}`} role="status">
            {t.tone === 'bad' ? <Icon name="info" /> : t.tone === 'good' ? <Icon name="check" /> : null}
            <span>{t.message}</span>
            {t.repeats > 1 ? <span className="rw-toast-count">×{t.repeats}</span> : null}
            {t.action ? (
              <button
                type="button"
                className="rw-btn ghost"
                style={{ color: 'inherit', textDecoration: 'underline' }}
                onClick={() => {
                  t.action.run();
                  setItems((list) => list.filter((x) => x.id !== t.id));
                }}
              >
                {t.action.label}
              </button>
            ) : null}
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
}

/* ── window frame ───────────────────────────────────────────────────────── */

export function Window({ app, children }) {
  return (
    <div className="rw" data-app={app}>
      {children}
    </div>
  );
}

/**
 * The caption bar. Frameless on every platform; macOS keeps its own traffic
 * lights, so the controls are drawn only where the platform has none.
 */
export function TitleBar({ app, title, subtitle, dirty, platform, shell, right, onMenu }) {
  const [maximized, setMaximized] = useState(false);
  const isMac = platform === 'darwin';

  useEffect(() => {
    let alive = true;
    shell?.win.state().then((s) => alive && setMaximized(Boolean(s.maximized)));
    const off = shell?.on('win:state', (s) => setMaximized(Boolean(s.maximized)));
    return () => {
      alive = false;
      off?.();
    };
  }, [shell]);

  return (
    <div className={`rw-titlebar${isMac ? ' mac' : ''}`}>
      <button type="button" className="rw-appmark interactive rw-btn ghost" onClick={onMenu} title="Rutba Office">
        <span className="glyph">
          <Icon name={app?.icon || 'home'} size={12} />
        </span>
        <span>{app?.short || 'Rutba Office'}</span>
        <Icon name="chevronDown" size={12} style={{ opacity: 0.5 }} />
      </button>

      <div className="rw-doctitle" title={title}>
        {title}
        {dirty ? <span className="dirty">•</span> : null}
        {subtitle ? <span style={{ opacity: 0.6 }}> — {subtitle}</span> : null}
      </div>

      <div className="interactive" style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
        {right}
      </div>

      {isMac ? null : (
        <div className="rw-wincontrols">
          <button type="button" onClick={() => shell?.win.minimize()} title="Minimise" aria-label="Minimise">
            <Icon name="minimize" size={14} />
          </button>
          <button
            type="button"
            onClick={async () => setMaximized((await shell?.win.toggleMaximize())?.maximized ?? false)}
            title={maximized ? 'Restore' : 'Maximise'}
            aria-label={maximized ? 'Restore' : 'Maximise'}
          >
            <Icon name={maximized ? 'restore' : 'maximize'} size={13} />
          </button>
          <button type="button" className="close" onClick={() => shell?.win.close()} title="Close" aria-label="Close">
            <Icon name="close" size={14} />
          </button>
        </div>
      )}
    </div>
  );
}

/* ── ribbon ─────────────────────────────────────────────────────────────── */

export function Ribbon({ tabs, active, onTab, quick, children }) {
  return (
    <div className="rw-ribbon">
      <div className="rw-tabs" role="tablist">
        {tabs.map((t) => (
          <button
            key={t.id}
            type="button"
            role="tab"
            className="rw-tab"
            aria-selected={t.id === active}
            onClick={() => onTab(t.id)}
          >
            {t.label}
          </button>
        ))}
        <div className="rw-quick">{quick}</div>
      </div>
      <div className="rw-groups" role="tabpanel">
        {children}
      </div>
    </div>
  );
}

export function Group({ label, children }) {
  return (
    <div className="rw-group">
      <div className="rw-group-items">{children}</div>
      {label ? <div className="rw-group-label">{label}</div> : null}
    </div>
  );
}

export function Button({
  icon,
  label,
  tall,
  primary,
  ghost,
  pressed,
  disabled,
  title,
  onClick,
  children,
  className = '',
  ...rest
}) {
  const classes = ['rw-btn', tall ? 'tall' : '', primary ? 'primary' : '', ghost ? 'ghost' : '', !label && !children ? 'icon' : '', className]
    .filter(Boolean)
    .join(' ');
  return (
    <button
      type="button"
      className={classes}
      aria-pressed={pressed === undefined ? undefined : Boolean(pressed)}
      disabled={disabled}
      title={title || label}
      onClick={onClick}
      {...rest}
    >
      {icon ? <Icon name={icon} size={tall ? 20 : 16} /> : null}
      {label ? <span>{label}</span> : null}
      {children}
    </button>
  );
}

export const Separator = () => <div className="rw-sep" />;

/* ── layout ─────────────────────────────────────────────────────────────── */

export const Body = ({ children, style }) => (
  <div className="rw-body" style={style}>
    {children}
  </div>
);

export function Panel({ title, width, right, actions, children, style, resizable, onResize }) {
  const ref = useRef(null);
  const drag = useRef(null);

  const start = (e) => {
    drag.current = { x: e.clientX, w: ref.current?.offsetWidth || width || 240 };
    const move = (ev) => {
      const next = Math.max(160, Math.min(640, drag.current.w + (right ? -1 : 1) * (ev.clientX - drag.current.x)));
      if (ref.current) ref.current.style.width = `${next}px`;
      onResize?.(next);
    };
    const stop = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', stop);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', stop);
  };

  return (
    <div
      ref={ref}
      className={`rw-panel${right ? ' right' : ''}`}
      style={{ width: width ?? 240, flex: 'none', position: 'relative', ...style }}
    >
      {title ? (
        <div className="rw-panel-head">
          <span style={{ flex: 1 }}>{title}</span>
          {actions}
        </div>
      ) : null}
      <div className="rw-panel-body">{children}</div>
      {resizable ? (
        <div
          onMouseDown={start}
          style={{
            position: 'absolute',
            top: 0,
            bottom: 0,
            [right ? 'left' : 'right']: -3,
            width: 6,
            cursor: 'col-resize',
            zIndex: 5,
          }}
        />
      ) : null}
    </div>
  );
}

export const Content = ({ children, style }) => (
  <div style={{ flex: 1, minWidth: 0, display: 'flex', flexDirection: 'column', overflow: 'hidden', ...style }}>
    {children}
  </div>
);

export function StatusBar({ children }) {
  return <div className="rw-status">{children}</div>;
}

export const Spacer = () => <div className="spacer" />;

export const Chip = ({ children, title }) => (
  <span className="chip" title={title}>
    {children}
  </span>
);

/* ── lists ──────────────────────────────────────────────────────────────── */

export function List({ children }) {
  return <div className="rw-list">{children}</div>;
}

export function Item({ icon, label, count, current, selected, indent = 0, onClick, onContextMenu, children, title }) {
  return (
    <button
      type="button"
      className={`rw-item${selected ? ' selected' : ''}`}
      aria-current={current || undefined}
      onClick={onClick}
      onContextMenu={onContextMenu}
      title={title || label}
      style={indent ? { paddingLeft: 9 + indent * 14 } : undefined}
    >
      {icon ? <Icon name={icon} size={15} /> : null}
      {label ? <span className="label">{label}</span> : null}
      {children}
      {count != null ? <span className="count">{count}</span> : null}
    </button>
  );
}

/* ── inputs ─────────────────────────────────────────────────────────────── */

export function Search({ value, onChange, placeholder = 'Search', style, onKeyDown, autoFocus }) {
  return (
    <div className="rw-search" style={style}>
      <Icon name="find" />
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        onKeyDown={onKeyDown}
        autoFocus={autoFocus}
        spellCheck={false}
      />
      {value ? (
        <button type="button" className="rw-btn icon ghost" onClick={() => onChange('')} title="Clear">
          <Icon name="close" size={13} />
        </button>
      ) : null}
    </div>
  );
}

export function Field({ label, hint, children }) {
  return (
    <div className="rw-field">
      {label ? <label>{label}</label> : null}
      {children}
      {hint ? <div className="rw-hint">{hint}</div> : null}
    </div>
  );
}

export const Input = (props) => <input className="rw-input" spellCheck={false} {...props} />;
export const Select = ({ children, ...rest }) => (
  <select className="rw-select" {...rest}>
    {children}
  </select>
);

/* ── dialogs and menus ──────────────────────────────────────────────────── */

export function Dialog({ title, children, actions, onClose, width }) {
  useEffect(() => {
    const onKey = (e) => {
      if (e.key === 'Escape') onClose?.();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  return (
    <div className="rw-scrim" onMouseDown={(e) => e.target === e.currentTarget && onClose?.()}>
      <div className="rw-dialog" style={width ? { width } : undefined} role="dialog" aria-modal="true" aria-label={title}>
        {title ? <div className="rw-dialog-head">{title}</div> : null}
        <div className="rw-dialog-body">{children}</div>
        {actions ? <div className="rw-dialog-foot">{actions}</div> : null}
      </div>
    </div>
  );
}

/**
 * A context menu placed where it was asked for, then nudged back on screen —
 * a menu opened near the bottom edge should not run off it.
 */
export function Menu({ x, y, items, onClose }) {
  const ref = useRef(null);
  const [pos, setPos] = useState({ x, y });

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    setPos({
      x: Math.min(x, window.innerWidth - r.width - 8),
      y: Math.min(y, window.innerHeight - r.height - 8),
    });
  }, [x, y]);

  useEffect(() => {
    const off = () => onClose?.();
    const key = (e) => e.key === 'Escape' && onClose?.();
    window.addEventListener('mousedown', off);
    window.addEventListener('blur', off);
    window.addEventListener('keydown', key);
    return () => {
      window.removeEventListener('mousedown', off);
      window.removeEventListener('blur', off);
      window.removeEventListener('keydown', key);
    };
  }, [onClose]);

  return (
    <div ref={ref} className="rw-menu" style={{ left: pos.x, top: pos.y }} onMouseDown={(e) => e.stopPropagation()}>
      {items.map((item, i) =>
        item === '-' ? (
          <hr key={`sep${i}`} />
        ) : (
          <button
            key={item.label}
            type="button"
            disabled={item.disabled}
            onClick={() => {
              item.run?.();
              onClose?.();
            }}
          >
            {item.icon ? <Icon name={item.icon} /> : <span style={{ width: 15 }} />}
            <span>{item.label}</span>
            {item.key ? <span className="key">{item.key}</span> : null}
          </button>
        )
      )}
    </div>
  );
}

/** Wire a right-click to a menu, and keep the state for it. */
export function useMenu() {
  const [menu, setMenu] = useState(null);
  const open = useCallback((e, items) => {
    e.preventDefault();
    e.stopPropagation();
    setMenu({ x: e.clientX, y: e.clientY, items });
  }, []);
  const close = useCallback(() => setMenu(null), []);
  const node = menu ? <Menu {...menu} onClose={close} /> : null;
  return { open, close, node };
}

/* ── states ─────────────────────────────────────────────────────────────── */

export function Empty({ icon = 'file', title, children, action }) {
  return (
    <div className="rw-empty">
      <div className="art">
        <Icon name={icon} size={28} />
      </div>
      {title ? <h3>{title}</h3> : null}
      {children ? <p>{children}</p> : null}
      {action}
    </div>
  );
}

export const Spinner = ({ style }) => <div className="rw-spinner" style={style} />;

export function Progress({ value, max = 1 }) {
  const pct = Math.max(0, Math.min(100, (value / (max || 1)) * 100));
  return (
    <div className="rw-progress">
      <div style={{ width: `${pct}%` }} />
    </div>
  );
}

/* ── helpers ────────────────────────────────────────────────────────────── */

export function formatBytes(n) {
  if (!Number.isFinite(n) || n <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.min(units.length - 1, Math.floor(Math.log(n) / Math.log(1024)));
  const v = n / 1024 ** i;
  return `${v >= 100 || i === 0 ? Math.round(v) : v.toFixed(1)} ${units[i]}`;
}

/** Dates the way a mail list shows them: time today, weekday this week, else date. */
export function formatWhen(value, { long = false } = {}) {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return '';
  if (long) return d.toLocaleString(undefined, { dateStyle: 'full', timeStyle: 'short' });
  const now = new Date();
  const sameDay = d.toDateString() === now.toDateString();
  if (sameDay) return d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  const days = (now - d) / 86400000;
  if (days < 7 && days > 0) return d.toLocaleDateString(undefined, { weekday: 'short' });
  if (d.getFullYear() === now.getFullYear()) return d.toLocaleDateString(undefined, { day: 'numeric', month: 'short' });
  return d.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
}

export const basename = (p) => String(p || '').split(/[\\/]/).pop() || '';
export const dirname = (p) => String(p || '').split(/[\\/]/).slice(0, -1).join('/');
export const extname = (p) => {
  const name = basename(p);
  const at = name.lastIndexOf('.');
  return at > 0 ? name.slice(at).toLowerCase() : '';
};
