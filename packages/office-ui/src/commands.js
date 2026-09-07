// Commands and keyboard.
//
// Every action a person can take is a command with an id, and the ribbon, the
// menus and the keyboard all invoke the same one. That is what stops the three
// from drifting — a button that does something the shortcut does not is the
// classic office-suite bug, and it cannot happen if there is one implementation
// and three ways to reach it.

import { useCallback, useEffect, useRef, useState } from 'react';

const IS_MAC = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent || '');

/** `Mod+S` reads as ⌘S on a Mac and Ctrl+S everywhere else. */
export function prettyKey(combo) {
  if (!combo) return '';
  return combo
    .split('+')
    .map((part) => {
      const p = part.trim().toLowerCase();
      if (p === 'mod') return IS_MAC ? '⌘' : 'Ctrl';
      if (p === 'shift') return IS_MAC ? '⇧' : 'Shift';
      if (p === 'alt') return IS_MAC ? '⌥' : 'Alt';
      if (p === 'ctrl') return IS_MAC ? '⌃' : 'Ctrl';
      if (p === 'enter') return '↵';
      if (p === 'escape') return 'Esc';
      if (p === 'delete') return 'Del';
      if (p === 'arrowup') return '↑';
      if (p === 'arrowdown') return '↓';
      if (p === 'arrowleft') return '←';
      if (p === 'arrowright') return '→';
      return part.length === 1 ? part.toUpperCase() : part[0].toUpperCase() + part.slice(1);
    })
    .join(IS_MAC ? '' : '+');
}

function matches(combo, event) {
  const parts = combo.toLowerCase().split('+').map((s) => s.trim());
  const key = parts[parts.length - 1];
  const want = {
    mod: parts.includes('mod'),
    shift: parts.includes('shift'),
    alt: parts.includes('alt'),
    ctrl: parts.includes('ctrl'),
  };
  const mod = IS_MAC ? event.metaKey : event.ctrlKey;
  if (want.mod !== mod) return false;
  if (want.shift !== event.shiftKey) return false;
  if (want.alt !== event.altKey) return false;
  if (want.ctrl && !IS_MAC && !event.ctrlKey) return false;
  const pressed = (event.key || '').toLowerCase();
  return pressed === key || (key.length === 1 && pressed === key);
}

/**
 * Bind a command table to the window.
 * @param {Record<string, { run: Function, key?: string, when?: () => boolean, label?: string, icon?: string }>} commands
 */
export function useCommands(commands, deps = []) {
  const table = useRef(commands);
  table.current = commands;

  const run = useCallback((id, ...args) => {
    const cmd = table.current[id];
    if (!cmd || cmd.disabled) return false;
    if (cmd.when && !cmd.when()) return false;
    cmd.run(...args);
    return true;
  }, []);

  useEffect(() => {
    const onKey = (event) => {
      // Typing in a field is typing, not a shortcut — except for the few that
      // must always work, which say so.
      const el = event.target;
      const typing =
        el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable || el.tagName === 'SELECT');

      for (const [id, cmd] of Object.entries(table.current)) {
        if (!cmd?.key) continue;
        if (typing && !cmd.global) continue;
        if (!matches(cmd.key, event)) continue;
        if (cmd.when && !cmd.when()) continue;
        event.preventDefault();
        cmd.run();
        return;
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  return run;
}

/** Menu items for a set of command ids, ready for the context menu. */
export function menuItems(commands, ids) {
  return ids.map((id) =>
    id === '-'
      ? '-'
      : {
          label: commands[id]?.label || id,
          icon: commands[id]?.icon,
          key: prettyKey(commands[id]?.key),
          disabled: commands[id]?.disabled || (commands[id]?.when ? !commands[id].when() : false),
          run: () => commands[id]?.run(),
        }
  );
}

/** Listen for menu commands sent by the backend's application menu. */
export function useShellCommands(shell, handler, deps = []) {
  const ref = useRef(handler);
  ref.current = handler;
  useEffect(() => {
    if (!shell?.on) return undefined;
    return shell.on('app:command', (payload) => ref.current(payload.command, payload.args));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);
}

export const isMac = IS_MAC;

/**
 * Load something asynchronously and re-render when it lands.
 *
 * Late answers are dropped: if the folder changes while the previous folder's
 * messages are still in flight, the stale list must not win the race and
 * replace the list the person is now looking at.
 */
export function useAsync(load, deps = [], initial = null) {
  const [state, setState] = useState({ value: initial, loading: true, error: null });
  const run = useRef(0);

  const reload = useCallback(() => {
    const ticket = ++run.current;
    setState((s) => ({ ...s, loading: true, error: null }));
    Promise.resolve()
      .then(load)
      .then((value) => {
        if (ticket === run.current) setState({ value, loading: false, error: null });
      })
      .catch((error) => {
        if (ticket === run.current) setState({ value: initial, loading: false, error });
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  useEffect(() => {
    reload();
  }, [reload]);

  return { ...state, reload, set: (value) => setState((s) => ({ ...s, value })) };
}
