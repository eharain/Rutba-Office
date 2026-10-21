// Global shortcuts, keys that work wherever the person is on the computer
// (Recorder's stop and pause, Take Picture), through an adapter: the system
// one is Electron's globalShortcut, and a check run is given a fake one
// (machine-fakes.js) that a check presses itself. The system adapter is
// checked on its own, in the gate's machine lane.

import { globalShortcut } from 'electron';

/** Electron's own global shortcuts. */
export const systemShortcuts = {
  kind: 'system',
  register: (accelerator, fn) => globalShortcut.register(accelerator, fn),
  unregister: (accelerator) => globalShortcut.unregister(accelerator),
  isRegistered: (accelerator) => globalShortcut.isRegistered(accelerator),
};

let current = systemShortcuts;

/** Use `adapter` for every global shortcut from now on. */
export function useShortcuts(adapter) {
  current = adapter || systemShortcuts;
}

/** The adapter in use. */
export const shortcuts = () => current;
