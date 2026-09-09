// The renderer entry point.
//
// One bundle serves all seven apps; which one a window is showing comes from
// its URL, set when the window was created. There is no router because there is
// no navigation: a window is one app for its whole life, and opening a document
// of another kind opens another window.

import React from 'react';
import { createRoot } from 'react-dom/client';
import shell, { boot } from '@rutba/office-shell/client';
import { ThemeProvider, ToastProvider } from '@rutba/office-ui';
import { APPS } from '@rutba/office-formats/registry';
import '@rutba/office-ui/theme.css';

import Home from './apps/home.js';
import Word from './apps/word.js';
import Sheets from './apps/sheets.js';
import Slides from './apps/slides.js';
import Pictures from './apps/pictures.js';
import ImageTool from './apps/image.js';
import VideoTool from './apps/video.js';
import Mail from './apps/mail.js';
import Calendar from './apps/calendar.js';
import Contacts from './apps/contacts.js';

const APP_COMPONENTS = {
  home: Home,
  word: Word,
  sheets: Sheets,
  slides: Slides,
  pictures: Pictures,
  image: ImageTool,
  video: VideoTool,
  mail: Mail,
  calendar: Calendar,
  contacts: Contacts,
};

function Root() {
  const start = boot();
  const key = APP_COMPONENTS[start.app] ? start.app : 'home';
  const App = APP_COMPONENTS[key];
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

// The stored theme is read before the first paint so a dark-mode window never
// flashes white on the way up.
shell.store
  .get({ key: 'theme', fallback: 'system' })
  .then((mode) => {
    window.__rutbaTheme = mode || 'system';
  })
  .catch(() => {})
  .finally(() => {
    createRoot(document.getElementById('root')).render(<Root />);
  });
