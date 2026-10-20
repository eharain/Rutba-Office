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
import shell, { boot } from '@rutba/office-shell/client';
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
  talk: () => import('./apps/talk.js'),
  pdf: () => import('./apps/pdf/index.js'),
  mail: () => import('./apps/mail.js'),
  calendar: () => import('./apps/calendar.js'),
  contacts: () => import('./apps/contacts.js'),
};

function Root({ start, App, appKey: key }) {
  const app = key === 'home' ? { key: 'home', name: 'Rutba Office', short: 'Rutba Office', icon: 'home' } : APPS[key];

  return (
    <ThemeProvider
      app={key}
      initial={window.__rutbaTheme || 'system'}
      onChange={(mode) => shell.store.set({ key: 'theme', value: mode })}
    >
      <ToastProvider>
        <App app={app} boot={start} shell={shell} />
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
const appKey = APP_LOADERS[start.app] ? start.app : 'home';
const loaded = APP_LOADERS[appKey]();
Promise.all([
  shell.store.get({ key: 'theme', fallback: 'system' }).then((mode) => { window.__rutbaTheme = mode || 'system'; }),
])
  .catch(() => {})
  .finally(async () => {
    const App = (await loaded).default;
    createRoot(document.getElementById('root')).render(<Root start={start} App={App} appKey={appKey} />);
  });
