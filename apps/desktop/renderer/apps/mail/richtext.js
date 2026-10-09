// Rich text in Mail: the compose editor's and the signature editor's.
//
// Both are a contentEditable, so both need the same three things: plain text
// turned into the markup the editor expects, the editor's markup turned back
// into plain text line for line (the plain part of a message, and a switch
// to Plain text that keeps what was typed), and markup from outside — a
// signature pasted from a web page — cut down to what a signature can carry
// before it is kept. That last is done in a DOMParser document, which runs
// no script and fetches nothing, never in the window's own.

import React, { useState } from 'react';
import { Button, Input, t } from '@rutba/office-ui';

/** Plain text into the markup the editor expects, quoting preserved. */
export function textToHtml(text) {
  return String(text)
    .split('\n')
    .map((line) => {
      const safe = line.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
      return line.startsWith('>') ? `<blockquote>${safe.replace(/^&gt;\s?/, '')}</blockquote>` : `<div>${safe || '<br>'}</div>`;
    })
    .join('');
}

const BLOCK = new Set(['DIV', 'P', 'BLOCKQUOTE', 'LI', 'UL', 'OL', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'TABLE', 'TR', 'PRE']);

/**
 * The editor's markup as plain text, the inverse of `textToHtml`: each block
 * a line, a <br> a line break, a quote's lines marked "> ", a link's
 * address after its words when they are not the address already.
 */
export function htmlToText(html) {
  const doc = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html');
  const lines = [];
  let line = '';
  let depth = 0;
  const flush = () => {
    lines.push('> '.repeat(depth) + line);
    line = '';
  };
  // Whether what was walked ended its line — on a <br>, or with a block of
  // its own — so the block round it does not end the line a second time.
  const walk = (node) => {
    let ended = false;
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        line += n.nodeValue.replace(/ /g, ' ');
        ended = false;
      } else if (n.nodeType !== 1 || n.nodeName === 'SCRIPT' || n.nodeName === 'STYLE') continue;
      else if (n.nodeName === 'BR') {
        flush();
        ended = true;
      } else if (n.nodeName === 'IMG') {
        line += n.getAttribute('alt') || '';
        ended = false;
      } else if (BLOCK.has(n.nodeName)) {
        if (line) flush();
        const quote = n.nodeName === 'BLOCKQUOTE';
        if (quote) depth++;
        if (!walk(n) || line) flush();
        if (quote) depth--;
        ended = true;
      } else if (n.nodeName === 'A') {
        walk(n);
        const href = n.getAttribute('href') || '';
        if (/^https?:/i.test(href) && n.textContent.trim() !== href) line += ` <${href}>`;
        ended = false;
      } else ended = walk(n);
    }
    return ended;
  };
  walk(doc.body);
  if (line) flush();
  return lines.join('\n');
}

/** What a signature may carry: the tags, and each one's attributes. */
const ALLOWED = {
  B: [], STRONG: [], I: [], EM: [], U: [], S: [], STRIKE: [], BR: [], DIV: ['style'], P: ['style'], SPAN: ['style'],
  FONT: ['color', 'size', 'face'], A: ['href'], IMG: ['src', 'alt', 'width', 'height'], UL: [], OL: [], LI: [],
  TABLE: [], TBODY: [], TR: [], TD: ['style'], SMALL: [], BIG: [], SUB: [], SUP: [], BLOCKQUOTE: [],
};
/** The style properties kept: the look of words, never their position or a fetch. */
const STYLES = new Set(['color', 'background-color', 'font-weight', 'font-style', 'font-size', 'font-family', 'text-decoration', 'text-align']);

/**
 * Markup cut down to what a signature can carry: words and their look,
 * links to web and mail addresses, pictures held in the signature itself or
 * on the web. Everything else — scripts, event handlers, forms, styles that
 * place things or fetch them — goes; a tag not allowed leaves its words.
 */
export function cleanHtml(html) {
  const doc = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html');
  const clean = (node) => {
    for (const n of [...node.childNodes]) {
      if (n.nodeType === 8) {
        n.remove();
        continue;
      }
      if (n.nodeType !== 1) continue;
      const allowed = ALLOWED[n.nodeName];
      if (!allowed) {
        if (['SCRIPT', 'STYLE', 'IFRAME', 'OBJECT', 'EMBED', 'TEMPLATE', 'FORM', 'INPUT', 'BUTTON', 'TEXTAREA', 'SELECT', 'SVG', 'MATH', 'LINK', 'META'].includes(n.nodeName)) n.remove();
        else {
          clean(n);
          n.replaceWith(...n.childNodes);
        }
        continue;
      }
      for (const attr of [...n.attributes]) {
        const name = attr.name.toLowerCase();
        const value = attr.value.trim();
        let keep = allowed.includes(name);
        if (keep && name === 'href') keep = /^(https?:|mailto:)/i.test(value);
        if (keep && name === 'src') keep = /^(https:|data:image\/(png|jpe?g|gif|webp);base64,)/i.test(value);
        if (keep && name === 'style') {
          const kept = value.split(';').map((d) => d.trim()).filter((d) => {
            const [prop, ...rest] = d.split(':');
            return STYLES.has(prop.trim().toLowerCase()) && !/url\(|expression|javascript:/i.test(rest.join(':'));
          });
          if (kept.length) {
            n.setAttribute('style', kept.join('; '));
            continue;
          }
          keep = false;
        }
        if (!keep) n.removeAttribute(attr.name);
      }
      clean(n);
    }
  };
  clean(doc.body);
  return doc.body.innerHTML.trim();
}

/** The selection inside an editor, kept, or null when it is elsewhere. */
export function selectionIn(editor) {
  const sel = window.getSelection();
  return sel?.rangeCount && editor?.contains(sel.anchorNode) ? sel.getRangeAt(0).cloneRange() : null;
}

/**
 * The link row a rich editor opens under its toolbar: an address, and Add.
 * Electron has no window.prompt, so a link asked for that way was never
 * made. `range` is the selection taken when the row was asked for — the
 * address box takes the focus, and with it the selection — put back before
 * the link is made.
 */
export function LinkRow({ editor, range, onDone }) {
  const [url, setUrl] = useState('https://');
  const add = () => {
    const href = url.trim();
    if (!/^(https?:\/\/.+|mailto:.+)/i.test(href)) return;
    editor?.focus();
    const sel = window.getSelection();
    if (range) {
      sel.removeAllRanges();
      sel.addRange(range);
    }
    try {
      // Nothing selected: the address itself becomes the link's words.
      if (sel.isCollapsed) document.execCommand('insertHTML', false, `<a href="${href.replace(/"/g, '&quot;')}">${href.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]))}</a>`);
      else document.execCommand('createLink', false, href);
    } catch {
      /* an engine without the command leaves the words as they were */
    }
    onDone();
  };
  return (
    <div className="ml-linkrow">
      <Input value={url} autoFocus onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); add(); } if (e.key === 'Escape') onDone(); }} aria-label={t('Link address')} />
      <Button primary label={t('Add link')} disabled={!/^(https?:\/\/.+|mailto:.+)/i.test(url.trim())} onClick={add} />
      <Button label={t('Cancel')} onClick={onDone} />
    </div>
  );
}
