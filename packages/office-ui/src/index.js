// The shared shell: everything seven apps have in common, and nothing they
// don't. A window frame with its own caption, a ribbon, panels, dialogs, menus
// and toasts — all of it unaware of documents, mail or media.

import React, { createContext, useContext, useCallback, useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { Icon } from './icons.js';
import { useControlNames, splitTip, labelFirstControl, modality, focusDocument } from './a11y.js';
export { labelFirstControl, nameControls, splitTip, cycleRegions, focusDocument } from './a11y.js';
import { useRibbonKeys } from './keytips.js';
import { t, tFilled } from './messages.js';

export { Icon };
export * from './commands.js';
export * from './zoom.js';
export { t, tn, msg, tFilled, setLanguage, language, languages, setDigits, numberTag, registerCatalogue, untranslated } from './messages.js';

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

  const push = useCallback((said, { tone = 'plain', ms = 3200, action } = {}) => {
    const id = ++seq.current;
    // An engine's message shown as it was thrown, in the window's language as well.
    const message = typeof said === 'string' ? tFilled(said) : said;
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
      <div className="rw-toast-host" aria-live="polite" aria-relevant="additions text">
        {items.map((t) => (
          <div key={t.id} className={`rw-toast ${t.tone === 'plain' ? '' : t.tone}`} role={t.tone === 'bad' ? 'alert' : undefined}>
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
  const [fullscreen, setFullscreen] = useState(false);
  const isMac = platform === 'darwin';

  useEffect(() => {
    let alive = true;
    shell?.win.state().then((s) => {
      if (!alive) return;
      setMaximized(Boolean(s.maximized));
      setFullscreen(Boolean(s.fullscreen));
    });
    const off = shell?.on('win:state', (s) => {
      setMaximized(Boolean(s.maximized));
      setFullscreen(Boolean(s.fullscreen));
    });
    return () => {
      alive = false;
      off?.();
    };
  }, [shell]);

  // Full screen has a key everywhere and a button almost nowhere. It is one
  // press on every platform here, and F11 (or Escape, to come back) still
  // works — the two stay in step because both go through the same call.
  const toggleFullscreen = async () => setFullscreen(Boolean((await shell?.win.fullscreen({ on: !fullscreen }))?.fullscreen));

  useEffect(() => {
    const onKey = (e) => {
      if (e.key === 'F11' || (e.key === 'f' && (e.metaKey || e.ctrlKey) && e.shiftKey)) {
        e.preventDefault();
        toggleFullscreen();
      } else if (e.key === 'Escape' && fullscreen) {
        toggleFullscreen();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  useTips();
  useControlNames();

  return (
    <div className={`rw-titlebar${isMac ? ' mac' : ''}`} role="banner">
      {/* words-ok: the suite's own name */}
      <button type="button" className="rw-appmark interactive rw-btn ghost" onClick={onMenu} title="Rutba Office">
        <span className="glyph">
          <Icon name={app?.icon || 'home'} size={12} />
        </span>
        <span>{app?.short || 'Rutba Office'}</span>
        <Icon name="chevronDown" size={12} style={{ opacity: 0.5 }} />
      </button>

      <h1 className="rw-doctitle" title={title}>
        {title}
        {dirty ? <span className="dirty" aria-hidden="true">•</span> : null}
        {subtitle ? <span style={{ opacity: 0.6 }}> — {subtitle}</span> : null}
      </h1>

      <div className="interactive" style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
        {right}
        {/*
          On macOS the traffic lights are the system's, so the full-screen
          button lives here beside them rather than in a set of controls this
          application does not draw.
        */}
        {isMac ? (
          <button
            type="button"
            className="rw-btn ghost"
            onClick={toggleFullscreen}
            title={fullscreen ? t('Leave full screen (Escape)') : t('Full screen (⌃⌘F)')}
            aria-label={fullscreen ? t('Leave full screen') : t('Full screen')}
          >
            <Icon name={fullscreen ? 'restore' : 'maximize'} size={13} />
          </button>
        ) : null}
      </div>

      {isMac ? null : (
        <div className="rw-wincontrols">
          <button
            type="button"
            onClick={toggleFullscreen}
            title={fullscreen ? t('Leave full screen (F11)') : t('Full screen (F11)')}
            aria-label={fullscreen ? t('Leave full screen') : t('Full screen')}
            className={fullscreen ? 'on' : undefined}
          >
            <Icon name={fullscreen ? 'zoomOut' : 'zoomIn'} size={13} />
          </button>
          <button type="button" onClick={() => shell?.win.minimize()} title={t('Minimise')} aria-label={t('Minimise')}>
            <Icon name="minimize" size={14} />
          </button>
          <button
            type="button"
            onClick={async () => setMaximized((await shell?.win.toggleMaximize())?.maximized ?? false)}
            title={maximized ? t('Restore') : t('Maximise')}
            aria-label={maximized ? t('Restore') : t('Maximise')}
          >
            <Icon name={maximized ? 'restore' : 'maximize'} size={13} />
          </button>
          <button type="button" className="close" onClick={() => shell?.win.close()} title={t('Close')} aria-label={t('Close')}>
            <Icon name="close" size={14} />
          </button>
        </div>
      )}
    </div>
  );
}

/* ── ribbon ─────────────────────────────────────────────────────────────── */

/**
 * A toolbar press must not take focus from the editor.
 *
 * Every button here acts on a selection in the page below it — bold applies to
 * the selected words, a font size to the selected cells. A click on a button
 * normally moves focus to the button, and the moment that happens the editor's
 * selection is what the browser considers "not the active selection": the DOM
 * range can survive it, but the editor has lost focus, the caret is gone, and
 * whatever the app does next runs against a page nobody is in.
 *
 * Preventing the default on mousedown stops the focus move and nothing else:
 * the click still fires. Inputs and selects are left alone — a dropdown that
 * cannot take focus cannot open.
 */
const keepEditorFocus = (e) => {
  const target = e.target instanceof Element ? e.target.closest('button, select, input, textarea') : null;
  if (target && target.tagName === 'BUTTON') e.preventDefault();
};

/**
 * The ribbon: a tab strip, the active tab's groups, and the quick-access
 * buttons at the strip's end.
 *
 * It COLLAPSES, the way Office's does: the chevron at the strip's end (or
 * Ctrl+F1) folds it to the tabs alone, and while folded a tab click peeks
 * the groups OVER the content — they float, the content does not move —
 * until a command is pressed or the pointer goes down elsewhere. An app
 * that wants to decide passes `collapsed`/`onCollapse` (the picture viewer
 * folds it whenever a picture is open, so the picture gets the room);
 * otherwise the ribbon keeps its own state.
 */
export function Ribbon({ tabs, active, onTab, quick, children, collapsed: collapsedProp, onCollapse, collapsible = true }) {
  const [own, setOwn] = useState(false);
  const collapsed = collapsible && (collapsedProp ?? own);
  const [peek, setPeek] = useState(false);
  const ref = useRef(null);
  // Icons only: every button its picture, its name in its tip. A choice for
  // every window, kept where the window keeps its own small things.
  const [iconsOnly, setIconsOnlyState] = useState(() => {
    try { return localStorage.getItem('rutba.ribbon.icons') === 'only'; } catch { return false; }
  });
  const setIconsOnly = (next) => {
    setIconsOnlyState(next);
    try { localStorage.setItem('rutba.ribbon.icons', next ? 'only' : 'names'); } catch { /* kept for this window alone */ }
  };
  const [options, setOptions] = useState(null);
  const setCollapsed = (next) => {
    if (onCollapse) onCollapse(next);
    if (collapsedProp === undefined) setOwn(next);
    setPeek(false);
  };

  // A peek closes on a press anywhere outside the ribbon, or on a command.
  useEffect(() => {
    if (!peek) return undefined;
    const close = (e) => { if (ref.current && !ref.current.contains(e.target)) setPeek(false); };
    document.addEventListener('mousedown', close, true);
    return () => document.removeEventListener('mousedown', close, true);
  }, [peek]);
  useEffect(() => { if (!collapsed) setPeek(false); }, [collapsed]);
  useEffect(() => {
    if (!collapsible) return undefined;
    const onKey = (e) => { if (e.ctrlKey && e.key === 'F1') { e.preventDefault(); setCollapsed(!collapsed); } };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [collapsed, collapsible, collapsedProp, onCollapse]);

  const showGroups = !collapsed || peek;
  const ids = useId();
  const tabId = (id) => `${ids}tab-${id}`;
  // KeyTips and the arrow keys (keytips.js).
  const keys = useRibbonKeys({ ref, tabs, active, onTab, collapsed, peek: setPeek });

  // The groups are wider than most windows. There is no bar (a bar under a
  // toolbar reads as a fault); a chevron appears on whichever side has more,
  // the wheel scrolls sideways, and the edge shadows say the rest.
  const groupsRef = useRef(null);
  const [more, setMore] = useState({ left: false, right: false });
  useEffect(() => {
    const el = groupsRef.current;
    if (!el || !showGroups) return undefined;
    const measure = () => {
      // Right to left, the groups start at the right and scroll to negative offsets.
      const rtl = getComputedStyle(el).direction === 'rtl';
      const gone = Math.abs(el.scrollLeft);
      const rest = gone + el.clientWidth < el.scrollWidth - 2;
      const left = rtl ? rest : gone > 2;
      const right = rtl ? gone > 2 : rest;
      setMore((m) => (m.left === left && m.right === right ? m : { left, right }));
    };
    measure();
    el.addEventListener('scroll', measure, { passive: true });
    const ro = typeof ResizeObserver === 'function' ? new ResizeObserver(measure) : null;
    ro?.observe(el);
    return () => {
      el.removeEventListener('scroll', measure);
      ro?.disconnect();
    };
  }, [showGroups, active, children]);

  // The groups fitted to the window, again whenever it or what is in the
  // groups changes (a tab chosen, a contextual group come and gone).
  useLayoutEffect(() => {
    const el = groupsRef.current;
    if (!el || !showGroups) return undefined;
    let frame = 0;
    const fit = () => {
      frame = 0;
      fitGroups(el, iconsOnly);
      el.dispatchEvent(new Event('scroll'));
    };
    const later = () => { if (!frame) frame = requestAnimationFrame(fit); };
    fitGroups(el, iconsOnly);
    const ro = typeof ResizeObserver === 'function' ? new ResizeObserver(later) : null;
    ro?.observe(el);
    // Changes the fitting makes itself (a folded group's picture) are not a reason to fit again.
    const mo = typeof MutationObserver === 'function' ? new MutationObserver((list) => { if (list.some((m) => !(m.target instanceof Element && m.target.closest('.rw-fold')))) later(); }) : null;
    mo?.observe(el, { childList: true, subtree: true, characterData: true });
    document.fonts?.ready?.then(later).catch(() => {});
    return () => {
      ro?.disconnect();
      mo?.disconnect();
      if (frame) cancelAnimationFrame(frame);
    };
  }, [showGroups, active, iconsOnly]);
  const scrollGroups = (dir) => {
    const el = groupsRef.current;
    if (el) el.scrollBy({ left: dir * Math.max(240, el.clientWidth * 0.6), behavior: 'smooth' });
  };
  const wheelSideways = (e) => {
    const el = groupsRef.current;
    if (!el || el.scrollWidth <= el.clientWidth) return;
    if (Math.abs(e.deltaY) > Math.abs(e.deltaX) && !e.target.closest('select')) {
      // The wheel moves on through the groups, whichever way they run.
      el.scrollLeft += getComputedStyle(el).direction === 'rtl' ? -e.deltaY : e.deltaY;
      e.preventDefault();
    }
  };

  return (
    <div ref={ref} className={`rw-ribbon${collapsed ? ' collapsed' : ''}${collapsed && peek ? ' peek' : ''}${iconsOnly ? ' icons-only' : ''}${keys.keyTips ? ' keytips' : ''}`} role="region" aria-label={t('Ribbon')} onMouseDown={keepEditorFocus} onKeyDown={keys.onKeyDown}>
      <div className="rw-tabs">
        <div className="rw-tablist" role="tablist" aria-label={t('Ribbon tabs')}>
        {tabs.map((t, i) => (
          <React.Fragment key={t.id}>
            {/* The contextual tabs, set off from the others by a line, as a group of their own. */}
            {t.contextual && !tabs[i - 1]?.contextual ? <span className="rw-tab-sep" aria-hidden="true" /> : null}
            <button
              type="button"
              role="tab"
              id={tabId(t.id)}
              className={t.contextual ? 'rw-tab contextual' : 'rw-tab'}
              data-label={typeof t.label === 'string' ? t.label : undefined}
              aria-selected={t.id === active}
              aria-controls={showGroups && t.id === active ? `${ids}panel` : undefined}
              tabIndex={t.id === active ? 0 : -1}
              onClick={() => {
                onTab(t.id);
                if (collapsed) setPeek((p) => !(p && t.id === active));
              }}
            >
              {t.label}
            </button>
          </React.Fragment>
        ))}
        </div>
        {options ? (
          <Menu
            x={options.x}
            y={options.y}
            rtl={options.rtl}
            onClose={() => setOptions(null)}
            items={[
              { label: t('Show labels'), icon: iconsOnly ? undefined : 'check', run: () => setIconsOnly(false) },
              { label: t('Icons only'), icon: iconsOnly ? 'check' : undefined, title: t('Every button its picture alone; its name shows when the pointer rests on it'), run: () => setIconsOnly(true) },
              ...(collapsible ? ['-', { label: collapsed ? t('Expand the ribbon') : t('Collapse the ribbon'), key: 'Ctrl+F1', run: () => setCollapsed(!collapsed) }] : []),
            ]}
          />
        ) : null}
        <div className="rw-quick">
          {quick}
          <button
            type="button"
            className="rw-ribbon-options"
            title={t('Ribbon display options')}
            aria-label={t('Ribbon display options')}
            aria-haspopup="menu"
            onClick={(e) => {
              const r = e.currentTarget.getBoundingClientRect();
              const rtl = getComputedStyle(e.currentTarget).direction === 'rtl';
              setOptions(options ? null : { x: rtl ? r.left : r.right, y: r.bottom + 4, rtl: !rtl });
            }}
          >
            <Icon name="sliders" size={13} />
          </button>
          {collapsible ? (
            <button
              type="button"
              className="rw-collapse"
              title={collapsed ? t('Expand the ribbon (Ctrl+F1)') : t('Collapse the ribbon (Ctrl+F1)')}
              aria-label={collapsed ? t('Expand the ribbon') : t('Collapse the ribbon')}
              aria-expanded={!collapsed}
              onClick={() => setCollapsed(!collapsed)}
            >
              <Icon name={collapsed ? 'chevronDown' : 'chevronUp'} size={12} />
            </button>
          ) : null}
        </div>
      </div>
      {showGroups ? (
        <div className="rw-groups-wrap">
          {more.left ? (
            <button type="button" className="rw-more left" title={t('More, to the left')} aria-label={t('Scroll the ribbon left')} onClick={() => scrollGroups(-1)}>
              <Icon name="chevronLeft" size={14} />
            </button>
          ) : null}
          <div
            ref={groupsRef}
            className="rw-groups"
            role="tabpanel"
            id={`${ids}panel`}
            aria-labelledby={tabs.some((x) => x.id === active) ? tabId(active) : undefined}
            onWheel={wheelSideways}
            onClick={(e) => {
              // A command pressed from a peek closes the peek, as in Office; a
              // dropdown keeps it, because a dropdown that closes cannot open.
              if (collapsed && peek && e.target.closest('button') && !e.target.closest('select, .rw-fold')) setPeek(false);
            }}
          >
            {children}
          </div>
          {more.right ? (
            <button type="button" className="rw-more right" title={t('More, to the right')} aria-label={t('Scroll the ribbon right')} onClick={() => scrollGroups(1)}>
              <Icon name="chevronRight" size={14} />
            </button>
          ) : null}
        </div>
      ) : null}
      {keys.overlay}
    </div>
  );
}

/**
 * A group of the ribbon: its controls, and its name under them.
 *
 * When the window is too narrow for every group, the ribbon folds the groups
 * from its end inwards (see `fitGroups`), and a folded group is one button
 * the height of the ribbon, with the group's picture and name: pressed, it
 * drops the group's controls under it, as they are when there is room, and
 * a command pressed there, a press elsewhere or Escape puts them away.
 * `icon` names the folded button's picture; without one it takes the
 * picture of the group's first control.
 */
export function Group({ label, icon, children }) {
  const ref = useRef(null);
  const [open, setOpen] = useState(false);

  useLayoutEffect(() => {
    if (!open) return undefined;
    const group = ref.current;
    const fold = group?.querySelector(':scope > .rw-fold');
    const pop = group?.querySelector(':scope > .rw-group-body');
    if (!group || !fold || !pop) return undefined;
    // Under the folded button, from its start edge, kept inside the window.
    const r = fold.getBoundingClientRect();
    const w = pop.offsetWidth;
    const rtl = getComputedStyle(group).direction === 'rtl';
    const x = Math.max(6, Math.min(window.innerWidth - w - 6, rtl ? r.right - w : r.left));
    pop.style.left = `${Math.round(x)}px`;
    pop.style.top = `${Math.round(r.bottom + 4)}px`;
    // The keyboard goes with it: to the first control, and back on Escape.
    if (group.dataset.keyOpened) pop.querySelector(FOCUSABLE)?.focus();
    const away = (e) => { if (!group.contains(e.target) && !e.target.closest?.('.rw-menu')) setOpen(false); };
    const key = (e) => {
      if (e.key !== 'Escape') return;
      e.stopPropagation();
      setOpen(false);
      fold.focus();
    };
    // A group that unfolds (the window grew) or a ribbon that moves puts it away.
    const gone = () => setOpen(false);
    const watch = typeof MutationObserver === 'function' ? new MutationObserver(() => { if (!group.classList.contains('folded')) setOpen(false); }) : null;
    watch?.observe(group, { attributes: true, attributeFilter: ['class'] });
    document.addEventListener('mousedown', away, true);
    group.addEventListener('keydown', key);
    window.addEventListener('resize', gone);
    window.addEventListener('blur', gone);
    return () => {
      delete group.dataset.keyOpened;
      watch?.disconnect();
      document.removeEventListener('mousedown', away, true);
      group.removeEventListener('keydown', key);
      window.removeEventListener('resize', gone);
      window.removeEventListener('blur', gone);
    };
  }, [open]);

  return (
    <div ref={ref} className="rw-group" data-open={open ? 'true' : undefined}>
      {/* The folded group's button; drawn only when the group is folded. Its name is drawn from data-label. */}
      <button
        type="button"
        className="rw-fold"
        data-label={typeof label === 'string' ? label : undefined}
        aria-label={typeof label === 'string' ? label : undefined}
        aria-haspopup="true"
        aria-expanded={open}
        onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ' || e.key === 'ArrowDown') { ref.current.dataset.keyOpened = '1'; if (e.key === 'ArrowDown') { e.preventDefault(); setOpen(true); } } }}
        onClick={() => setOpen((o) => !o)}
      >
        <span className="rw-fold-icon">{icon ? <Icon name={icon} size={24} strokeWidth={1.5} /> : null}</span>
        <Icon name="chevronDown" size={9} strokeWidth={2.2} className="rw-caret" />
      </button>
      <div
        className="rw-group-body"
        onClick={(e) => {
          // A command pressed in a dropped group puts the group away, as in
          // Office; a field or a list stays, since it is not done with.
          if (open && e.target.closest('button') && !e.target.closest('select, input')) setTimeout(() => setOpen(false), 0);
        }}
      >
        <div className="rw-group-items">{children}</div>
        {label ? <div className="rw-group-label">{label}</div> : null}
      </div>
    </div>
  );
}

/**
 * Fit the groups to the ribbon's width, as Office does when its window
 * narrows: first a gallery shows fewer of its items (Word's Styles,
 * PowerPoint's Themes), one at a time down to its fewest, then the large
 * buttons become small ones stacked three high, then the small buttons lose
 * their names and keep their pictures, and last whole groups fold into a
 * button each; every step taken from the end of the ribbon inwards, and no
 * more of them than it takes. A group whose pictures repeat keeps its
 * names, since its buttons could not be told apart without them. Classes on
 * the groups' own nodes: React never sets them, so a render keeps them.
 */
const STAGES = ['compact', 'small', 'folded'];
function fitGroups(el, iconsOnly = false) {
  const groups = [...el.querySelectorAll(':scope > .rw-group')];
  for (const g of groups) g.classList.remove(...STAGES);
  const galleries = (g) => [...g.querySelectorAll('.rw-gallery')];
  for (const g of groups) for (const gal of galleries(g)) galleryShow(gal, Number(gal.dataset.max) || Infinity);
  const fits = () => el.scrollWidth <= el.clientWidth + 1;
  if (!groups.length || fits()) return;
  // Icons only, every group is small already; what does not fit still folds.
  for (const stage of iconsOnly ? ['gallery', 'folded'] : ['gallery', ...STAGES]) {
    for (let i = groups.length - 1; i >= 0; i--) {
      const g = groups[i];
      if (g.classList.contains('folded')) continue;
      if (stage === 'gallery') {
        for (const gal of galleries(g).reverse()) {
          const fewest = Math.max(1, Number(gal.dataset.min) || 1);
          while (Number(gal.dataset.shown) > fewest) {
            galleryShow(gal, Number(gal.dataset.shown) - 1);
            if (fits()) return giveBack(groups, fits);
          }
        }
        continue;
      }
      if (stage === 'folded') {
        g.classList.remove('compact', 'small');
        // Dropped from its button, a folded group's gallery has the room to show all it would.
        for (const gal of galleries(g)) galleryShow(gal, Number(gal.dataset.max) || Infinity);
        const slot = g.querySelector(':scope > .rw-fold > .rw-fold-icon');
        if (slot && !slot.firstChild) {
          const first = g.querySelector('.rw-group-items .rw-btn > svg:not(.rw-caret)');
          if (first) slot.appendChild(first.cloneNode(true));
        }
        g.classList.add('folded');
      } else {
        if (stage === 'small' && !namesCanGo(g)) continue;
        const before = g.getBoundingClientRect().width;
        g.classList.add(stage);
        if (g.getBoundingClientRect().width >= before - 1) {
          g.classList.remove(stage);
          continue;
        }
      }
      if (fits()) return giveBack(groups, fits);
    }
  }
}

/**
 * A group shrunk early may have room again once a later one has folded:
 * from the start of the ribbon, each gets back what it gave up, where the
 * ribbon still fits, its buttons' names first and then its gallery's items.
 */
function giveBack(groups, fits) {
  for (const g of groups) {
    for (const stage of ['small', 'compact']) {
      if (!g.classList.contains(stage)) continue;
      g.classList.remove(stage);
      if (!fits()) {
        g.classList.add(stage);
        break;
      }
    }
  }
  for (const g of groups) {
    if (g.classList.contains('folded')) continue;
    for (const gal of g.querySelectorAll('.rw-gallery')) {
      const most = Number(gal.dataset.max) || Infinity;
      while (Number(gal.dataset.shown) < Math.min(most, gal.querySelector(':scope > .rw-gallery-items')?.children.length || 0)) {
        const was = Number(gal.dataset.shown);
        galleryShow(gal, was + 1);
        if (!fits()) {
          galleryShow(gal, was);
          break;
        }
      }
    }
  }
}

/**
 * A gallery shows its first `count` items' worth of width (fewer when it
 * has fewer), the rest reached with its arrows. The items keep their own
 * widths, so the viewport is the span of the first `count` of them; the
 * step its arrows scroll by is the span from one page to the next.
 */
export function galleryShow(gal, count) {
  const items = gal?.querySelector(':scope > .rw-gallery-items');
  if (!items) return;
  const kids = [...items.children];
  if (!kids.length) {
    items.style.maxWidth = '';
    gal.dataset.shown = '0';
    return;
  }
  const n = Math.max(1, Math.min(kids.length, Number.isFinite(count) ? count : kids.length));
  const a = kids[0].getBoundingClientRect();
  const b = kids[n - 1].getBoundingClientRect();
  const span = Math.max(a.right, b.right) - Math.min(a.left, b.left);
  const cs = getComputedStyle(items);
  const pad = (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.paddingRight) || 0) + (cs.boxSizing === 'border-box' ? (parseFloat(cs.borderLeftWidth) || 0) + (parseFloat(cs.borderRightWidth) || 0) : 0);
  items.style.maxWidth = `${Math.ceil(span + (cs.boxSizing === 'border-box' ? pad : 0))}px`;
  const next = kids[n]?.getBoundingClientRect();
  gal.dataset.step = String(Math.round(next ? Math.abs(next.left - a.left) : span));
  gal.dataset.shown = String(n);
}

/**
 * A gallery in the ribbon, as Office draws one: a row of choices (Word's
 * Styles, PowerPoint's Themes and Variants) with three arrows at its end.
 * Up and down page back and on through the choices it has no room to show;
 * the third, under a bar, opens them all (`onMore`). `max` is how many it
 * shows when there is room, `min` the fewest it is shrunk to, a choice at
 * a time, before its group folds into a button. `label` names the choices
 * in the arrows' tips; `moreTip` and `moreClass` dress the third arrow.
 */
export function Gallery({ children, max = 6, min = 3, label = '', onMore, moreTip, moreClass = '', className = '', itemsClassName = '' }) {
  const ref = useRef(null);
  const [edge, setEdge] = useState({ start: true, end: true });
  useLayoutEffect(() => {
    const gal = ref.current;
    const items = gal?.querySelector(':scope > .rw-gallery-items');
    if (!items) return undefined;
    // Shown at its length until the ribbon fits it (fitGroups shows it again when the window changes).
    if (!gal.dataset.shown) galleryShow(gal, max);
    const measure = () => {
      const at = Math.abs(items.scrollLeft);
      const start = at < 1;
      const end = at + items.clientWidth >= items.scrollWidth - 1;
      setEdge((e) => (e.start === start && e.end === end ? e : { start, end }));
    };
    measure();
    items.addEventListener('scroll', measure, { passive: true });
    const ro = typeof ResizeObserver === 'function' ? new ResizeObserver(measure) : null;
    ro?.observe(items);
    return () => {
      items.removeEventListener('scroll', measure);
      ro?.disconnect();
    };
  }, [max, children]);
  const page = (dir) => {
    const gal = ref.current;
    const items = gal?.querySelector(':scope > .rw-gallery-items');
    if (!items) return;
    const rtl = getComputedStyle(items).direction === 'rtl';
    const step = Number(gal.dataset.step) || items.clientWidth;
    items.scrollBy({ left: dir * step * (rtl ? -1 : 1), behavior: 'smooth' });
  };
  const named = label || t('choices');
  return (
    <div ref={ref} className={`rw-gallery ${className}`.trim()} data-max={max} data-min={min}>
      <div className={`rw-gallery-items ${itemsClassName}`.trim()}>{children}</div>
      <div className="rw-gallery-arrows">
        <button type="button" className="rw-gallery-up" disabled={edge.start} data-tip={t('Previous {label}', { label: named })} aria-label={t('Previous {label}', { label: named })} onClick={() => page(-1)}>
          <Icon name="chevronUp" size={10} strokeWidth={2.2} />
        </button>
        <button type="button" className="rw-gallery-down" disabled={edge.end} data-tip={t('Next {label}', { label: named })} aria-label={t('Next {label}', { label: named })} onClick={() => page(1)}>
          <Icon name="chevronDown" size={10} strokeWidth={2.2} />
        </button>
        <button type="button" className={`rw-gallery-more ${moreClass}`.trim()} disabled={!onMore} data-tip={moreTip || t('More: every one of the {label}', { label: named })} aria-label={moreTip || t('More: every one of the {label}', { label: named })} onClick={(e) => onMore?.(e)}>
          <span className="rw-gallery-bar" aria-hidden="true" />
          <Icon name="chevronDown" size={10} strokeWidth={2.2} />
        </button>
      </div>
    </div>
  );
}

/** Whether a group's buttons can be told apart by their pictures alone. */
function namesCanGo(g) {
  const named = [...g.querySelectorAll('.rw-group-items .rw-btn')].filter((b) => b.querySelector(':scope > span') && b.querySelector(':scope > svg:not(.rw-caret)'));
  if (!named.length) return false;
  const all = [...g.querySelectorAll('.rw-group-items .rw-btn > svg:not(.rw-caret)')].map((s) => s.getAttribute('data-icon'));
  const mine = named.map((b) => b.querySelector(':scope > svg:not(.rw-caret)').getAttribute('data-icon'));
  return mine.every((icon) => all.filter((x) => x === icon).length === 1);
}

/**
 * A name in two lines, split at the space nearest its middle (Read / aloud,
 * in any language), each line kept whole; its text reads the same.
 */
function inTwo(label) {
  if (typeof label !== 'string') return label;
  const words = label.split(' ');
  if (words.length < 2) return label;
  let at = 1;
  for (let i = 2; i < words.length; i++) {
    const gap = (n) => Math.abs(words.slice(0, n).join(' ').length - words.slice(n).join(' ').length);
    if (gap(i) < gap(at)) at = i;
  }
  return <><span className="rw-line">{words.slice(0, at).join(' ')}</span> <span className="rw-line">{words.slice(at).join(' ')}</span></>;
}

/**
 * A large button's name with its menu's arrow: the last word and the arrow
 * kept on one line, so the arrow never drops to a line of its own under a
 * two-line name (Paragraph Spacing), where the name's two lines cut it off.
 */
function withCaret(label, caret) {
  if (typeof label !== 'string') return label;
  const at = label.lastIndexOf(' ');
  // One word: its arrow on the line under it, as Office draws Colours and Fonts.
  return at < 0
    ? <>{label}{caret}</>
    : <>{label.slice(0, at + 1)}<span className="rw-btn-last">{label.slice(at + 1)}{caret}</span></>;
}

/**
 * Small buttons stacked in rows beside the tall ones, the way Office lays a
 * group out: each child is one row. A Home ribbon of single-row groups was
 * 2,372 px wide; in rows it fits a normal window.
 */
export function Rows({ children }) {
  return (
    <div className="rw-rows">
      {React.Children.map(children, (row, i) => (row === null || row === undefined || row === false ? null : <div className="rw-row" key={i}>{row}</div>))}
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
  busy,
  dropdown,
  twoLine,
  title,
  onClick,
  children,
  className = '',
  ...rest
}) {
  // `busy`: its work is under way (a password being checked): it cannot be
  // pressed again, and a small spinner says so in place of its icon.
  // `dropdown`: it opens a menu, and a small chevron says so: after its name
  // on a large button, as Office draws it, and after its picture on a small one.
  // `twoLine`: a large button's name on two lines whatever the room, as
  // Office draws Read Aloud, so its group is no wider than its picture.
  const classes = ['rw-btn', tall ? 'tall' : '', tall && twoLine ? 'two-line' : '', primary ? 'primary' : '', ghost ? 'ghost' : '', !label && !children ? 'icon' : '', busy ? 'busy' : '', className]
    .filter(Boolean)
    .join(' ');
  const caret = dropdown ? <Icon name="chevronDown" size={9} strokeWidth={2.2} className="rw-caret" /> : null;
  return (
    <button
      type="button"
      className={classes}
      aria-pressed={pressed === undefined ? undefined : Boolean(pressed)}
      aria-haspopup={dropdown ? 'menu' : undefined}
      aria-expanded={dropdown ? false : undefined}
      disabled={disabled || busy}
      aria-busy={busy || undefined}
      data-tip={title || label}
      aria-label={!label && !children ? splitTip(title).name || undefined : undefined}
      onClick={onClick}
      {...rest}
    >
      {busy ? <Spinner /> : icon ? <Icon name={icon} size={tall ? 24 : 16} strokeWidth={tall ? 1.5 : 1.7} /> : null}
      {label ? <span>{tall && caret ? withCaret(label, caret) : tall && twoLine ? inTwo(label) : label}{tall && caret && typeof label !== 'string' ? caret : null}</span> : null}
      {children}
      {tall && label ? null : caret}
    </button>
  );
}

export const Separator = () => <div className="rw-sep" />;

/**
 * Tooltips of our own: one small dark chip under the pointer's element,
 * after a short rest, for anything with a `data-tip` (every ribbon button)
 * or a `title`. The browser's tooltip is a grey box that arrives late and
 * looks like nothing else in the window; this one is the suite's. An
 * element's `title` is held aside while the pointer is on it — the browser
 * draws its own tip from the attribute, and two tips are one too many — and
 * given back when the pointer leaves.
 */
function useTips() {
  useEffect(() => {
    // One per document, kept between mounts: a second frame in the same
    // window would otherwise leave an empty first one behind.
    let tip = document.querySelector('.rw-tip');
    tip?.setAttribute('aria-hidden', 'true');
    if (!tip) {
      tip = document.createElement('div');
      tip.className = 'rw-tip';
      tip.setAttribute('role', 'tooltip');
      tip.setAttribute('aria-hidden', 'true');
      document.body.appendChild(tip);
    }
    let timer = 0;
    let held = null;
    const hide = () => {
      clearTimeout(timer);
      timer = 0;
      tip.classList.remove('on');
      if (held) {
        held.el.setAttribute('title', held.title);
        held = null;
      }
    };
    const show = (el, text, lead = '') => {
      tip.textContent = '';
      if (lead) {
        const b = document.createElement('b');
        b.textContent = lead;
        tip.append(b, document.createElement('br'));
      }
      tip.append(text);
      tip.classList.add('on');
      const r = el.getBoundingClientRect();
      const w = tip.offsetWidth;
      const h = tip.offsetHeight;
      let left = Math.round(r.left + r.width / 2 - w / 2);
      left = Math.max(6, Math.min(window.innerWidth - w - 6, left));
      let top = Math.round(r.bottom + 8);
      if (top + h > window.innerHeight - 6) top = Math.round(r.top - h - 8);
      tip.style.left = `${left}px`;
      tip.style.top = `${top}px`;
    };
    const over = (e) => {
      const el = e.target.closest?.('[data-tip], [title]');
      if (!el || el === held?.el) return;
      hide();
      let text = el.getAttribute('data-tip') || el.getAttribute('title');
      // A button's name, when its tip does not begin with it, comes first: with
      // the ribbon showing icons only, the tip is where the name is read.
      const name = el.matches?.('.rw-btn') ? el.querySelector(':scope > span')?.textContent.trim() : '';
      const lead = name && text && !text.startsWith(name) ? name : '';
      if (!text || el.disabled === true && !el.getAttribute('data-tip')) return;
      if (el.hasAttribute('title')) {
        held = { el, title: el.getAttribute('title') };
        el.removeAttribute('title');
      }
      timer = setTimeout(() => show(el, text, lead), 420);
      const leave = () => {
        el.removeEventListener('mouseleave', leave);
        hide();
      };
      el.addEventListener('mouseleave', leave);
    };
    document.addEventListener('mouseover', over);
    document.addEventListener('mousedown', hide, true);
    document.addEventListener('keydown', hide, true);
    window.addEventListener('blur', hide);
    return () => {
      document.removeEventListener('mouseover', over);
      document.removeEventListener('mousedown', hide, true);
      document.removeEventListener('keydown', hide, true);
      window.removeEventListener('blur', hide);
      hide();
    };
  }, []);
}

/**
 * The zoom slider every document app keeps at the right of its status bar:
 * minus, a slider, plus, and the level as a button that puts it back to
 * 100% (or, for a deck, to fit). `value` is a factor, 1 being actual size.
 */
export function ZoomSlider({ value = 1, min = 0.5, max = 2, step = 0.05, onChange, onReset, resetLabel = null }) {
  const pct = Math.round((value || 1) * 100);
  const clamp = (v) => Math.max(min, Math.min(max, Math.round(v * 100) / 100));
  return (
    <div className="rw-zoom" role="group" aria-label={t('Zoom')}>
      <button type="button" className="rw-zoom-step" data-tip={t('Zoom out (Ctrl+-)')} onClick={() => onChange?.(clamp(value - 0.1))}>−</button>
      <input
        type="range"
        min={min}
        max={max}
        step={step}
        value={Math.max(min, Math.min(max, value || 1))}
        aria-label={t('Zoom level')}
        onChange={(e) => onChange?.(clamp(Number(e.target.value)))}
      />
      <button type="button" className="rw-zoom-step" data-tip={t('Zoom in (Ctrl++)')} onClick={() => onChange?.(clamp(value + 0.1))}>+</button>
      <button type="button" className="rw-zoom-pct" data-tip={resetLabel || t('Back to 100%')} onClick={() => onReset?.()}>{pct}%</button>
    </div>
  );
}

/* ── layout ─────────────────────────────────────────────────────────────── */

export const Body = ({ children, style }) => (
  <main className="rw-body" style={style}>
    {children}
  </main>
);

/** Where a panel keeps the width it was given: by the window's app and the panel's name. */
const panelKey = (name, right) => `rutba.panel.${(typeof document !== 'undefined' && document.querySelector('.rw[data-app]')?.dataset.app) || 'app'}.${name || (right ? 'right' : 'side')}`;
const keptWidth = (key) => { try { const v = Number(localStorage.getItem(key)); return Number.isFinite(v) && v > 0 ? v : null; } catch { return null; } };
const keepWidth = (key, w) => { try { if (w == null) localStorage.removeItem(key); else localStorage.setItem(key, String(Math.round(w))); } catch { /* kept for this window alone */ } };

/**
 * A panel beside the window's work: a list of folders, a properties pane.
 * `resizable`, it has a grip on its inner edge that shows when the pointer
 * or the keyboard is on it: dragged, or moved with the arrow keys, it gives
 * the panel more or less room, between `min` and `max` (and never more
 * than half the window); a double-click puts it back. The width is kept on
 * this computer for the window's app and the panel's `name` (its title when
 * it has no name), so it comes back the next time.
 */
export function Panel({ title, width, right, actions, children, style, resizable, onResize, name, min = 160, max = 640 }) {
  const ref = useRef(null);
  const base = width ?? 240;
  const key = resizable ? panelKey(name || (typeof title === 'string' ? title : ''), right) : null;
  const [own, setOwn] = useState(() => (key ? keptWidth(key) : null));
  // Another panel's name (a pane that changes what it shows) brings its own width.
  const lastKey = useRef(key);
  useEffect(() => {
    if (lastKey.current === key) return;
    lastKey.current = key;
    setOwn(key ? keptWidth(key) : null);
  }, [key]);
  // A width the window sets itself (Move Split's arrow keys) is the width from then on.
  const lastBase = useRef(base);
  useEffect(() => {
    if (lastBase.current === base) return;
    lastBase.current = base;
    if (key) { setOwn(null); keepWidth(key, null); }
  }, [base, key]);
  const most = () => Math.max(min, Math.min(max, Math.round((typeof window !== 'undefined' ? window.innerWidth : 1600) * 0.5)));
  const limit = (w) => Math.round(Math.max(min, Math.min(most(), w)));
  const shown = resizable && own ? limit(own) : base;
  const set = (w) => {
    const next = limit(w);
    setOwn(next);
    if (key) keepWidth(key, next);
    onResize?.(next);
  };
  const towards = () => (right ? -1 : 1) * (ref.current && getComputedStyle(ref.current).direction === 'rtl' ? -1 : 1);

  const start = (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    const x0 = e.clientX;
    const w0 = ref.current?.offsetWidth || shown;
    const sign = towards();
    let next = w0;
    ref.current?.classList.add('resizing');
    const move = (ev) => {
      next = limit(w0 + sign * (ev.clientX - x0));
      if (ref.current) ref.current.style.width = `${next}px`;
      onResize?.(next);
    };
    const stop = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', stop);
      ref.current?.classList.remove('resizing');
      set(next);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', stop);
  };
  const key2 = (e) => {
    const step = e.shiftKey ? 64 : 16;
    const by = e.key === 'ArrowRight' ? step : e.key === 'ArrowLeft' ? -step : 0;
    if (e.key === 'Home') { e.preventDefault(); setOwn(null); if (key) keepWidth(key, null); onResize?.(base); return; }
    if (!by) return;
    e.preventDefault();
    set((ref.current?.offsetWidth || shown) + by * towards());
  };
  const reset = () => { setOwn(null); if (key) keepWidth(key, null); onResize?.(base); };

  return (
    <div
      ref={ref}
      className={`rw-panel${right ? ' right' : ''}`}
      role={typeof title === 'string' && title ? 'region' : undefined}
      aria-label={typeof title === 'string' && title ? title : undefined}
      data-panel={key ? key.split('.').pop() || undefined : undefined}
      style={{ width: shown, flex: 'none', position: 'relative', ...style }}
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
          className="rw-panel-grip"
          role="separator"
          aria-orientation="vertical"
          aria-valuenow={Math.round(shown)}
          aria-valuemin={min}
          aria-valuemax={most()}
          aria-label={typeof title === 'string' && title ? t('Width of {name}', { name: title }) : t('Width of the panel')}
          data-tip={t('Drag to make the panel wider or narrower, or press the arrow keys; double-click puts it back')}
          tabIndex={0}
          onMouseDown={start}
          onDoubleClick={reset}
          onKeyDown={key2}
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
  return <div className="rw-status" role="region" aria-label={t('Status bar')}>{children}</div>;
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

export function Search({ value, onChange, placeholder, style, onKeyDown, autoFocus }) {
  return (
    <div className="rw-search" style={style}>
      <Icon name="find" />
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder ?? t('Search')}
        onKeyDown={onKeyDown}
        autoFocus={autoFocus}
        spellCheck={false}
      />
      {value ? (
        <button type="button" className="rw-btn icon ghost" onClick={() => onChange('')} title={t('Clear')}>
          <Icon name="close" size={13} />
        </button>
      ) : null}
    </div>
  );
}

export function Field({ label, hint, children }) {
  const ref = useRef(null);
  const id = useId();
  // The first control in the field that has no name of its own takes the
  // field's label (and its hint as its description), whatever drew it.
  useLayoutEffect(() => {
    if (label) labelFirstControl(ref.current, `${id}label`, hint ? `${id}hint` : null);
  });
  return (
    <div ref={ref} className="rw-field">
      {label ? <label id={`${id}label`}>{label}</label> : null}
      {children}
      {hint ? <div className="rw-hint" id={`${id}hint`}>{hint}</div> : null}
    </div>
  );
}

// A field or a list in a ribbon is named by its tip when nothing else names it.
export const Input = (props) => <input className="rw-input" spellCheck={false} aria-label={props['aria-label'] ?? (props['aria-labelledby'] ? undefined : props.title)} {...props} />;
export const Select = ({ children, ...rest }) => (
  <select className="rw-select" aria-label={rest['aria-label'] ?? (rest['aria-labelledby'] ? undefined : rest.title)} {...rest}>
    {children}
  </select>
);

/* ── dialogs and menus ──────────────────────────────────────────────────── */

/** The dialogs open now, the last on top: Escape closes the top one only. */
const openDialogs = [];

const FOCUSABLE = 'a[href], button:not([disabled]), input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"]), [contenteditable="true"]';

/** What Tab can reach inside `root`, in order, leaving out what is not drawn. */
export function focusablesIn(root) {
  return [...root.querySelectorAll(FOCUSABLE)].filter((el) => el.getClientRects().length > 0);
}

/**
 * A modal dialog. Focus goes into it when it opens — to what its author
 * marked `autoFocus`, or else the first thing in it that takes focus — and
 * Tab and Shift+Tab go round inside it rather than into the page behind;
 * when it closes, focus goes back to where it was. Escape closes the top
 * dialog only, so a dialog opened from another leaves the first open.
 */
export function Dialog({ title, children, actions, onClose, width, className = null }) {
  const ref = useRef(null);
  const close = useRef(onClose);
  close.current = onClose;
  // Where focus was, read as the dialog is first drawn: by the time its
  // effects run, a field marked autoFocus has already taken it.
  const [before] = useState(() => (typeof document === 'undefined' ? null : document.activeElement));
  useEffect(() => {
    const self = {};
    openDialogs.push(self);
    const box = ref.current;
    if (box && !box.contains(document.activeElement)) (focusablesIn(box)[0] || box).focus({ preventScroll: true });
    const onKey = (e) => {
      if (e.key === 'Escape' && openDialogs[openDialogs.length - 1] === self) close.current?.();
    };
    window.addEventListener('keydown', onKey);
    return () => {
      window.removeEventListener('keydown', onKey);
      openDialogs.splice(openDialogs.indexOf(self), 1);
      const focus = document.activeElement;
      if (!focus || focus === document.body || box?.contains(focus)) {
        if (before?.isConnected && before !== document.body) before.focus?.({ preventScroll: true });
        else setTimeout(() => { if (!document.activeElement || document.activeElement === document.body) focusDocument(); }, 0);
      }
    };
  }, []);
  const trap = (e) => {
    if (e.key !== 'Tab' || !ref.current) return;
    const all = focusablesIn(ref.current);
    if (!all.length) { e.preventDefault(); return; }
    const first = all[0];
    const last = all[all.length - 1];
    const at = document.activeElement;
    if (e.shiftKey && (at === first || !ref.current.contains(at) || at === ref.current)) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && (at === last || !ref.current.contains(at))) { e.preventDefault(); first.focus(); }
  };

  return (
    <div className="rw-scrim" onMouseDown={(e) => e.target === e.currentTarget && onClose?.()}>
      <div ref={ref} className={`rw-dialog${className ? ` ${className}` : ''}`} style={width ? { width } : undefined} role="dialog" aria-modal="true" aria-label={title} tabIndex={-1} onKeyDown={trap}>
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
export function Menu({ x, y, rtl = false, items, onClose, keyboard = false }) {
  const ref = useRef(null);
  const [pos, setPos] = useState({ x, y });
  const [from] = useState(() => (typeof document === 'undefined' ? null : document.activeElement));
  const entries = () => [...(ref.current?.querySelectorAll('button:not(:disabled)') || [])];
  // The keyboard goes back while the items are still on the page, so it is never on nothing for a moment
  // (a sheet takes the keyboard whenever it finds it there).
  useLayoutEffect(() => {
    const box = ref.current;
    return () => {
      if (box?.contains(document.activeElement) && from && from.isConnected && from !== document.body) from.focus({ preventScroll: true });
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (keyboard || modality.last === 'key') entries()[0]?.focus({ preventScroll: true });
    const box = ref.current;
    return () => {
      const now = document.activeElement;
      if (!now || now === document.body || box?.contains(now)) {
        if (from && from.isConnected && from !== document.body) from.focus({ preventScroll: true });
        else focusDocument();
      }
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const keys = (e) => {
    const all = entries();
    if (!all.length) return;
    const at = all.indexOf(document.activeElement);
    const to = e.key === 'ArrowDown' ? (at + 1) % all.length : e.key === 'ArrowUp' ? (at <= 0 ? all.length - 1 : at - 1) : e.key === 'Home' ? 0 : e.key === 'End' ? all.length - 1 : -1;
    if (to >= 0) { e.preventDefault(); e.stopPropagation(); all[to].focus(); return; }
    if (e.key === 'Tab') { e.preventDefault(); onClose?.(); }
  };

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    setPos({
      x: Math.max(8, Math.min(rtl ? x - r.width : x, window.innerWidth - r.width - 8)),
      y: Math.min(y, window.innerHeight - r.height - 8),
    });
  }, [x, y, rtl]);

  useEffect(() => {
    const off = () => onClose?.();
    const key = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); onClose?.(); return; }
      // The arrows take the keyboard into a menu opened by the pointer.
      if ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && !ref.current?.contains(document.activeElement)) {
        const all = entries();
        if (all.length) { e.preventDefault(); (e.key === 'ArrowDown' ? all[0] : all[all.length - 1]).focus(); }
      }
    };
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
    <div ref={ref} className="rw-menu" role="menu" style={{ left: pos.x, top: pos.y }} onMouseDown={(e) => e.stopPropagation()} onKeyDown={keys}>
      {items.map((item, i) =>
        item === '-' ? (
          <hr key={`sep${i}`} role="separator" />
        ) : item.heading ? (
          // A section's name over its items, the way Office heads Entrance, Emphasis and Exit.
          <div key={`head${i}`} className="rw-menu-head" role="presentation">{item.label}</div>
        ) : (
          <button
            key={`${i}:${item.label}`}
            type="button"
            role={item.icon === 'check' || item.checked !== undefined ? 'menuitemcheckbox' : 'menuitem'}
            aria-checked={item.icon === 'check' || item.checked !== undefined ? Boolean(item.checked ?? item.icon === 'check') : undefined}
            disabled={item.disabled}
            title={item.title}
            onClick={() => {
              item.run?.();
              onClose?.();
            }}
          >
            {/* A gallery's sample, where an item shows what it makes (a WordArt style), in place of an icon. */}
            {item.preview ? <span className="rw-menu-preview">{item.preview}</span> : item.icon ? <Icon name={item.icon} /> : <span style={{ width: 15 }} />}
            <span>{item.label}</span>
            {item.key ? <span className="key" aria-hidden="true">{item.key}</span> : null}
          </button>
        )
      )}
    </div>
  );
}

/** Wire a right-click to a menu, and keep the state for it. */
export function useMenu() {
  const [menu, setMenu] = useState(null);
  const opener = useRef(null);
  const open = useCallback((e, items) => {
    e.preventDefault();
    e.stopPropagation();
    const from = e.currentTarget instanceof Element ? e.currentTarget.closest('button, [role="button"]') : null;
    if (opener.current && opener.current !== from) opener.current.setAttribute('aria-expanded', 'false');
    opener.current = from;
    if (from?.getAttribute('aria-haspopup')) from.setAttribute('aria-expanded', 'true');
    // A ribbon button's menu drops from the button, as Office's do, rather
    // than from wherever the pointer was on it, and from its far edge right
    // to left. A right click anywhere else opens where it was made.
    const keyboard = (e.type === 'click' && e.detail === 0) || (e.type === 'contextmenu' && !e.clientX && !e.clientY);
    const button = e.currentTarget instanceof Element && e.type !== 'contextmenu' ? e.currentTarget.closest('.rw-ribbon button') : null;
    const r = button?.getBoundingClientRect();
    if (r && r.width) {
      const rtl = getComputedStyle(button).direction === 'rtl';
      setMenu({ x: rtl ? r.right : r.left, y: r.bottom + 2, rtl, items, keyboard });
    } else if (e.type === 'contextmenu' && !e.clientX && !e.clientY && e.target instanceof Element) {
      const at = (document.activeElement && document.activeElement !== document.body ? document.activeElement : e.target).getBoundingClientRect();
      setMenu({ x: at.left + Math.min(24, at.width / 2), y: at.top + Math.min(at.height, 24), items, keyboard });
    } else {
      setMenu({ x: e.clientX, y: e.clientY, items, keyboard });
    }
  }, []);
  const close = useCallback(() => {
    if (opener.current?.getAttribute('aria-haspopup')) opener.current.setAttribute('aria-expanded', 'false');
    setMenu(null);
  }, []);
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
    <div className="rw-progress" role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(pct)}>
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
