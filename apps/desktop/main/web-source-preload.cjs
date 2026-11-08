// A web page source's one bridge (web-source.js): the page may say what it is
// showing, and Office may pass it the show page's own messages. Nothing else.

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('rutbaShow', {
  state: (s) => ipcRenderer.send('rutba-show:state', {
    index: Number.isInteger(s?.index) ? s.index : 0,
    count: Number.isInteger(s?.count) ? s.count : 0,
    title: typeof s?.title === 'string' ? s.title.slice(0, 200) : '',
  }),
});

ipcRenderer.on('rutba-show:command', (_event, message) => {
  window.postMessage(message, '*');
});
