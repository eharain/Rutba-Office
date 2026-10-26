// The profile password window's one bridge (profile-lock.js): the password
// typed, handed to the main process to try, and Cancel. Nothing else.

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('rutbaLock', {
  open: (password) => ipcRenderer.invoke('rutba-office/profile-lock:open', String(password ?? '')),
  cancel: () => ipcRenderer.invoke('rutba-office/profile-lock:cancel'),
});
