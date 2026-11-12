// The renderer entry point.
//
// Which app a window is showing comes from its URL, set when the window was
// created. There is no router because there is no navigation: a window is one
// app for its whole life, and opening a document of another kind opens another
// window. So each app is a bundle of its own, loaded when its window opens
// (owner, 2026-10-20, the build split by app): a window parses its own app
// and what it shares, and an app turned off in Choose apps is never loaded.

// First: the window's language, settled before any app's words are made.
import './language.js';
import React from 'react';
import { createRoot } from 'react-dom/client';
import shell, { boot, isInstalled, packWindows } from '@rutba/office-shell/client';
import { ThemeProvider, ToastProvider } from '@rutba/office-ui';
import { APPS } from '@rutba/office-formats/registry';
import '@rutba/office-ui/theme.css';
import { installLostMouseupGuard } from './lost-mouseup.js';

const APP_LOADERS = {
  home: () => import('./apps/home.js'),
  word: () => import('./apps/word.js'),
  sheets: () => import('./apps/sheets.js'),
  slides: () => import('./apps/slides.js'),
  pictures: () => import('./apps/pictures.js'),
  image: () => import('./apps/image.js'),
  video: () => import('./apps/video/index.js'),
  recorder: () => import('./apps/recorder.js'),
  broadcast: () => import('./apps/broadcast.js'),
  talk: () => import('./apps/talk.js'),
  base: () => import('./apps/base.js'),
  pdf: () => import('./apps/pdf/index.js'),
  mail: () => import('./apps/mail.js'),
  calendar: () => import('./apps/calendar.js'),
  contacts: () => import('./apps/contacts.js'),
  ...(process.env.RUTBA_LOCKER === '1' ? { locker: () => import('./apps/locker.js') } : {}),
};

// An app that is not in the registry, carried only by a build made with its switch.
const EXTRA_APPS = process.env.RUTBA_LOCKER === '1' ? { locker: { key: 'locker', name: 'Locker', short: 'Locker', icon: 'lock', accent: '#2f4a6b' } } : {};

// The window says it is ready, for whatever waits on it (the checks, the screen runs): the app's first render is in
// and two frames are drawn (a window drawing none, hidden, says so after half a second). The checks used to wait
// for a class no window has, .rw-app, and ran out at 15 s on every window of an app with no mark of its own:
// Calendar, Broadcast, Recorder and the Locker among them (12 Nov gate, 26 Calendar windows).
function Ready({ children }) {
  React.useEffect(() => {
    let said = false;
    const say = () => {
      if (said) return;
      said = true;
      document.documentElement.dataset.ready = '1';
      window.dispatchEvent(new Event('rutba-ready'));
    };
    requestAnimationFrame(() => requestAnimationFrame(say));
    const late = setTimeout(say, 500);
    return () => clearTimeout(late);
  }, []);
  return children;
}

function Root({ start, App, appKey: key }) {
  const app = key === 'home' ? { key: 'home', name: 'Rutba Office', short: 'Rutba Office', icon: 'home' } : (APPS[key] || EXTRA_APPS[key]);

  return (
    <ThemeProvider
      app={key}
      initial={window.__rutbaTheme || 'system'}
      onChange={(mode) => shell.store.set({ key: 'theme', value: mode })}
    >
      <ToastProvider>
        <Ready><App app={app} boot={start} shell={shell} /></Ready>
      </ToastProvider>
    </ThemeProvider>
  );
}

// A drag whose mouseup is lost — let go outside the window, or the window left mid-drag — still ends.
installLostMouseupGuard(window);

// The stored theme is read before the first paint so a dark-mode window never
// flashes white on the way up, and the window's app is loaded meanwhile. (The
// language is settled earlier: language.js.)
const start = boot();
// Only an installed app is loaded (docs/APP-PACKS.md): the main process makes
// no window for an absent one, and a window that asks for one all the same
// shows the launcher. A pack that carries its own window module is loaded
// from its archive; an app still built into the core, from its chunk here.
const own = (key) => packWindows(key);
const appKey = start.app !== 'home' && isInstalled(start.app) && (own(start.app) || APP_LOADERS[start.app]) ? start.app : 'home';
const loaded = own(appKey) ? import(new URL(own(appKey), document.baseURI).href) : APP_LOADERS[appKey]();
Promise.all([
  shell.store.get({ key: 'theme', fallback: 'system' }).then((mode) => { window.__rutbaTheme = mode || 'system'; }),
])
  .catch(() => {})
  .finally(async () => {
    const App = (await loaded).default;
    createRoot(document.getElementById('root')).render(<Root start={start} App={App} appKey={appKey} />);
  });
