// The pieces Rutba Mail is assembled from.
//
// Kept apart from the app itself because they are the parts with rules rather
// than the parts with state: how a message body is made safe to draw, how a
// conversation is recognised, how a person becomes a coloured circle. Each is
// a pure function, which is also what makes them testable without a window.

/* ── the reading frame ───────────────────────────────────────────────────── */

export const BLOCKED_NOTE = 'blocked-remote';

/**
 * Build the document shown in the sandboxed frame.
 *
 * Everything remote is stripped unless the reader asked for it; `cid:` images
 * are rewritten to the URLs the backend is holding for this message, so inline
 * pictures work with no network at all.
 */
export function bodyDocument(message, { remote = false, dark = false, plain = false } = {}) {
  const escape = (s) => s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  let html = plain ? '' : message.html || '';

  if (!html) {
    const text = message.text || (plain ? stripTags(message.html || '') : '');
    html = `<pre class="plain">${escape(text)}</pre>`;
  }

  // Inline images first — these are ours and are safe.
  for (const [cid, url] of Object.entries(message.inlineUrls || {})) {
    html = html.split(`cid:${cid}`).join(url);
  }

  if (!remote) {
    html = html.replace(/(<img\b[^>]*?\bsrc=)(["'])(https?:[^"']*)\2/gi, `$1$2$2 data-${BLOCKED_NOTE}=$2$3$2`);
  }

  const csp = remote
    ? "default-src 'none'; img-src https: data: rutba: blob:; style-src 'unsafe-inline'; font-src data:"
    : "default-src 'none'; img-src data: rutba: blob:; style-src 'unsafe-inline'; font-src data:";

  return `<!doctype html><html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<style>
  :root { color-scheme: ${dark ? 'dark' : 'light'}; }
  body {
    margin: 0; padding: 18px 24px;
    font: 14px/1.6 "Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif;
    color: ${dark ? '#e6e9ee' : '#1a1c20'}; background: ${dark ? '#1e2127' : '#ffffff'};
    word-wrap: break-word; overflow-wrap: anywhere;
  }
  img { max-width: 100%; height: auto; }
  pre.plain { white-space: pre-wrap; font: inherit; margin: 0; }
  blockquote { margin: 0 0 0 12px; padding-left: 12px; border-left: 3px solid ${dark ? '#3b4048' : '#d5d9e0'}; color: ${dark ? '#a9b0bb' : '#5b626d'}; }
  a { color: ${dark ? '#7fb0ff' : '#1a56c4'}; }
  table { max-width: 100%; }
</style></head><body>${html}</body></html>`;
}

/** HTML to something readable, for the plain-text view and for quoting a reply. */
export function stripTags(html) {
  return String(html || '')
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|li|h[1-6])>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/* ── people ──────────────────────────────────────────────────────────────── */

// Enough hues to tell a dozen correspondents apart at a glance, all at the same
// lightness so no one of them shouts. The choice is deterministic, so the same
// person is the same colour in every folder, on every machine, forever.
const TONES = ['#5b6ee1', '#c2477d', '#0d8f6f', '#b4690e', '#7a4fd0', '#0f7ea8', '#c14343', '#4a7c1f', '#8a5a2b', '#2a6fb5'];

/** A stable colour and initial for an address. */
export function avatarFor(person) {
  const address = String(person?.address || person?.name || '?').toLowerCase();
  let hash = 0;
  for (let i = 0; i < address.length; i++) hash = (hash * 31 + address.charCodeAt(i)) >>> 0;
  const label = String(person?.name || person?.address || '?').trim();
  // The first letter of a name, or of the mailbox part of a bare address.
  const initial = (/^[a-z0-9]/i.test(label) ? label : label.replace(/^[^a-z0-9]+/i, '') || '?').charAt(0).toUpperCase();
  return { colour: TONES[hash % TONES.length], initial: initial || '?' };
}

export const displayName = (person) => person?.name || person?.address || '(unknown sender)';

/* ── conversations ───────────────────────────────────────────────────────── */

// Every reply prefix in wide use. Mail is international and a German client
// writes "AW:", a French one "RE :", a Dutch one "Antw:" — a thread that only
// recognises "Re:" splits in half the moment someone answers from Outlook in
// another language.
const REPLY_PREFIX = /^\s*(?:(?:re|aw|antw|antwort|sv|vs|ref|odp|res|r|fw|fwd|wg|tr|rv|enc|doorst|vb|回复|转发)\s*(?:\[\d+\])?\s*[:：]\s*)+/i;

/** The subject a whole conversation shares. */
export function threadSubject(subject) {
  return String(subject || '')
    .replace(REPLY_PREFIX, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Group a flat list into conversations.
 *
 * Message-Id chains are the correct way to do this and are used first: a reply
 * names what it answers, and that survives a changed subject. Subject is the
 * fallback, because a great deal of real mail — anything that went through a
 * list, anything forwarded by hand, everything in an imported archive whose
 * References header was dropped — has no chain to follow.
 *
 * @param {object[]} rows newest-first headers
 * @returns {object[]} one entry per conversation, newest first, each with
 *   `messages` in date order and the newest message's fields on the entry
 */
export function buildThreads(rows) {
  const byId = new Map();
  for (const row of rows) if (row.messageId) byId.set(row.messageId, row);

  // Union-find over the chain links, then subject as a second pass.
  const parent = new Map();
  const find = (k) => {
    let root = k;
    while (parent.get(root) !== root) root = parent.get(root);
    while (parent.get(k) !== root) {
      const next = parent.get(k);
      parent.set(k, root);
      k = next;
    }
    return root;
  };
  const union = (a, b) => {
    const ra = find(a);
    const rb = find(b);
    if (ra !== rb) parent.set(ra, rb);
  };

  const keyOf = (row) => `${row.accountId || ''}|${row.id}`;
  for (const row of rows) parent.set(keyOf(row), keyOf(row));

  for (const row of rows) {
    for (const ref of [row.inReplyTo, ...(row.references || [])]) {
      const other = ref && byId.get(ref);
      if (other && other !== row) union(keyOf(row), keyOf(other));
    }
  }

  // Subject fallback, scoped to the account so two people's unrelated "Invoice"
  // threads in a unified list do not become one conversation.
  const bySubject = new Map();
  for (const row of rows) {
    const subject = threadSubject(row.subject);
    if (!subject) continue;
    const key = `${row.accountId || ''}|${subject}`;
    if (bySubject.has(key)) union(keyOf(row), bySubject.get(key));
    else bySubject.set(key, keyOf(row));
  }

  const groups = new Map();
  for (const row of rows) {
    const root = find(keyOf(row));
    if (!groups.has(root)) groups.set(root, []);
    groups.get(root).push(row);
  }

  const threads = [];
  for (const messages of groups.values()) {
    messages.sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    const newest = messages[messages.length - 1];
    threads.push({
      ...newest,
      threadKey: `${newest.accountId || ''}|${newest.id}`,
      messages,
      count: messages.length,
      unread: messages.some((m) => m.unread),
      flagged: messages.some((m) => m.flagged),
      pinned: messages.some((m) => m.pinned),
      hasAttachments: messages.some((m) => m.hasAttachments),
      // Who is in this conversation, oldest first, without repeats: the line
      // Yahoo and Gmail both show, and the fastest way to recognise a thread.
      participants: [...new Map(messages.map((m) => [m.from?.address || displayName(m.from), m.from])).values()],
    });
  }

  threads.sort((a, b) => Number(b.pinned || false) - Number(a.pinned || false) || (b.date || '').localeCompare(a.date || ''));
  return threads;
}

/* ── views ───────────────────────────────────────────────────────────────── */

export const FILTERS = [
  { id: 'all', label: 'All', test: () => true },
  { id: 'unread', label: 'Unread', icon: 'mail', test: (r) => r.unread },
  { id: 'starred', label: 'Starred', icon: 'star', test: (r) => r.flagged },
  { id: 'attachments', label: 'Attachments', icon: 'attach', test: (r) => r.hasAttachments },
  { id: 'people', label: 'People', icon: 'reply', test: (r) => !r.bulk },
];

export const SORTS = [
  { id: 'date', label: 'Date', of: (r) => r.date || '' },
  { id: 'from', label: 'Sender', of: (r) => String(r.from?.name || r.from?.address || '').toLowerCase() },
  { id: 'subject', label: 'Subject', of: (r) => threadSubject(r.subject) },
  { id: 'size', label: 'Size', of: (r) => String(r.size || 0).padStart(12, '0') },
];

/** Apply the chosen filter, search and order to a page of headers. */
export function arrange(rows, { filter = 'all', sort = 'date', ascending = false } = {}) {
  const test = FILTERS.find((f) => f.id === filter)?.test ?? (() => true);
  const of = SORTS.find((s) => s.id === sort)?.of ?? SORTS[0].of;
  const out = rows.filter(test);
  out.sort((a, b) => {
    const cmp = String(of(a)).localeCompare(String(of(b)));
    return ascending ? cmp : -cmp;
  });
  return out;
}

/* ── the stylesheet ──────────────────────────────────────────────────────── */

export const CSS = `
/* rail ------------------------------------------------------------------- */
.ml-compose-cta { padding: 10px 10px 8px; }
.ml-compose-cta button {
  width: 100%; display: flex; align-items: center; justify-content: center; gap: 8px;
  padding: 9px 12px; border: 0; border-radius: var(--r-2); background: var(--accent);
  color: #fff; font: inherit; font-weight: 600; font-size: 13px; cursor: pointer;
  transition: filter var(--fast);
}
.ml-compose-cta button:hover { filter: brightness(1.08); }
.ml-compose-cta button:disabled { opacity: .5; cursor: default; }
.ml-accounts { padding: 2px 6px; border-bottom: 1px solid var(--line-soft); }

/* filter strip ----------------------------------------------------------- */
.ml-tools { padding: 8px 10px 6px; display: flex; flex-direction: column; gap: 7px; }
.ml-filters { display: flex; gap: 4px; align-items: center; }
.ml-filter {
  display: inline-flex; align-items: center; gap: 5px; padding: 3px 10px; border-radius: 20px;
  border: 1px solid transparent; background: transparent; color: var(--ink-2);
  font: inherit; font-size: 11.5px; cursor: pointer; white-space: nowrap;
}
.ml-filter:hover { background: var(--hover); }
.ml-filter.on { background: var(--selected); border-color: var(--accent-line); color: var(--accent-ink, var(--ink)); font-weight: 600; }
.ml-listbar {
  display: flex; align-items: center; gap: 6px; padding: 4px 10px; font-size: 11.5px;
  color: var(--ink-3); border-bottom: 1px solid var(--line-soft); min-height: 30px;
}
.ml-listbar .rw-btn { padding: 2px 6px; }
.ml-check { width: 14px; height: 14px; accent-color: var(--accent); cursor: pointer; margin: 0; }

/* the list --------------------------------------------------------------- */
.ml-list { display: flex; flex-direction: column; overflow: auto; flex: 1; min-height: 0; }
.ml-row {
  display: grid; grid-template-columns: auto auto 1fr auto; align-items: start; gap: 9px;
  padding: 9px 12px 9px 9px; text-align: left; border: 0; width: 100%;
  border-bottom: 1px solid var(--line-soft); background: transparent; color: var(--ink);
  font: inherit; position: relative; cursor: pointer; transition: background var(--fast);
}
.ml-row:hover { background: var(--hover); }
.ml-row.selected { background: var(--selected); box-shadow: inset 3px 0 0 var(--accent); }
.ml-row.checked { background: var(--selected); }
.ml-row.pinned { background: color-mix(in srgb, var(--accent) 6%, transparent); }
.ml-list.compact .ml-row { padding: 5px 12px 5px 9px; }
.ml-list.compact .ml-preview { display: none; }
.ml-list.compact .ml-avatar-sm { width: 20px; height: 20px; font-size: 10px; }

.ml-avatar-sm {
  width: 30px; height: 30px; border-radius: 50%; display: grid; place-items: center;
  color: #fff; font-weight: 600; font-size: 12.5px; flex: none; user-select: none;
}
.ml-row-main { min-width: 0; display: flex; flex-direction: column; gap: 1px; }
.ml-row-top { display: flex; align-items: baseline; gap: 8px; }
.ml-from { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12.5px; }
.ml-row.unread .ml-from { font-weight: 700; }
.ml-count {
  flex: none; font-size: 10.5px; color: var(--ink-3); border: 1px solid var(--line);
  border-radius: 20px; padding: 0 6px; font-variant-numeric: tabular-nums;
}
.ml-subject { font-size: 12.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; display: flex; align-items: center; gap: 5px; }
.ml-row.unread .ml-subject { font-weight: 600; }
.ml-preview { font-size: 11.5px; color: var(--ink-3); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-row-side { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; flex: none; }
.ml-when { font-size: 11px; color: var(--ink-3); white-space: nowrap; font-variant-numeric: tabular-nums; }
.ml-marks { display: flex; align-items: center; gap: 4px; color: var(--ink-3); }
.ml-star { border: 0; background: none; padding: 0; cursor: pointer; color: var(--ink-4, var(--ink-3)); display: grid; }
.ml-star.on { color: #e0a800; }
.ml-star:hover { color: #e0a800; }
.ml-folder-tag {
  font-size: 10px; color: var(--ink-3); background: var(--sunken); border-radius: 3px;
  padding: 0 5px; max-width: 110px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.ml-thread-children { background: var(--sunken); }
.ml-thread-children .ml-row { padding-left: 34px; border-bottom-color: transparent; }

/* selection bar ---------------------------------------------------------- */
.ml-bulk {
  display: flex; align-items: center; gap: 6px; padding: 5px 10px;
  background: var(--selected); border-bottom: 1px solid var(--accent-line); font-size: 12px;
}
.ml-bulk strong { font-variant-numeric: tabular-nums; }

/* reader ----------------------------------------------------------------- */
.ml-reader { flex: 1; display: flex; flex-direction: column; min-height: 0; background: var(--surface); }
.ml-head { padding: 16px 22px 10px; border-bottom: 1px solid var(--line); }
.ml-head h2 { margin: 0 0 12px; font-family: var(--font-display); font-size: 18px; font-weight: 600; display: flex; align-items: center; gap: 9px; }
.ml-meta { display: flex; align-items: center; gap: 11px; }
.ml-avatar { width: 36px; height: 36px; border-radius: 50%; color: #fff; display: grid; place-items: center; font-weight: 600; flex: none; }
.ml-meta-text { flex: 1; min-width: 0; font-size: 12.5px; }
.ml-addr { color: var(--ink-3); }
.ml-to { color: var(--ink-3); font-size: 11.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-date { color: var(--ink-3); font-size: 11.5px; flex: none; }
.ml-attachments { display: flex; flex-wrap: wrap; gap: 7px; margin-top: 12px; }
.ml-keep { margin-left: 8px; border: 0; background: none; padding: 0 4px; font: inherit; font-size: 11.5px; color: var(--accent); cursor: pointer; display: inline-flex; align-items: center; gap: 3px; vertical-align: middle; }
.ml-invite { display: flex; align-items: center; gap: 12px; margin: 12px 0 0; padding: 12px 14px; border-radius: var(--r-2); background: var(--sunken); border-left: 4px solid #1a9f7a; }
.ml-invite-text { flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 2px; }
.ml-invite-title { font-weight: 600; }
.ml-invite-when, .ml-invite-who { font-size: 12.5px; color: var(--ink-2); }
.ml-invite-actions { display: flex; gap: 6px; }
.ml-suggest { position: absolute; left: 0; right: 0; top: 100%; z-index: 20; margin-top: 2px; padding: 4px; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); box-shadow: 0 6px 20px rgba(0,0,0,.12); display: flex; flex-direction: column; }
.ml-suggest button { display: flex; align-items: baseline; gap: 8px; border: 0; background: none; padding: 5px 8px; border-radius: 4px; font: inherit; text-align: left; color: var(--ink); cursor: pointer; }
.ml-suggest button.on, .ml-suggest button:hover { background: var(--selected); }
.ml-suggest .n { font-weight: 500; }
.ml-suggest .e { color: var(--ink-3); font-size: 12px; }
.ml-suggest .s { margin-left: auto; color: var(--ink-3); font-size: 11px; }
.ml-address { position: relative; }
.ml-attachment {
  display: inline-flex; align-items: center; gap: 7px; padding: 5px 10px;
  border: 1px solid var(--line); border-radius: 20px; background: var(--chrome);
  color: var(--ink); font: inherit; font-size: 12px; cursor: pointer;
}
.ml-attachment:hover { border-color: var(--accent-line); background: var(--selected); }
.ml-attachment .size { color: var(--ink-3); font-size: 11px; }
.ml-body { flex: 1; min-height: 0; }
.ml-body iframe { width: 100%; border: 0; display: block; }
.ml-headers { border-top: 1px solid var(--line); padding: 8px 22px; font-size: 11.5px; color: var(--ink-3); }
.ml-headers pre { white-space: pre-wrap; max-height: 220px; overflow: auto; font-family: var(--mono); font-size: 11px; }

/* the privacy strip ------------------------------------------------------ */
.ml-strip { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-top: 11px; }
.ml-badge {
  display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px; border-radius: 20px;
  background: var(--sunken); font-size: 11.5px; color: var(--ink-2); border: 1px solid transparent;
  font: inherit; font-size: 11.5px; cursor: default;
}
.ml-badge.act { cursor: pointer; }
.ml-badge.act:hover { border-color: var(--accent-line); background: var(--selected); }
.ml-badge.warn { color: #a8620a; background: color-mix(in srgb, #e0a800 14%, transparent); }
.ml-badge.good { color: var(--good); background: color-mix(in srgb, var(--good) 12%, transparent); }
.ml-badge.bad { color: var(--bad); background: color-mix(in srgb, var(--bad) 12%, transparent); }
.ml-trackers { margin-top: 10px; border: 1px solid var(--line); border-radius: var(--r-2); overflow: hidden; }
.ml-tracker {
  display: flex; align-items: center; gap: 9px; padding: 6px 11px; font-size: 12px;
  border-bottom: 1px solid var(--line-soft);
}
.ml-tracker:last-child { border-bottom: 0; }
.ml-tracker .host { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-family: var(--mono); font-size: 11px; }
.ml-tracker .net { font-weight: 600; }
.ml-tracker .pix { color: var(--bad); font-size: 11px; }

/* dialogs and forms ------------------------------------------------------ */
.ml-form { display: flex; flex-direction: column; gap: 11px; padding: 4px 0 10px; }
.ml-servers { display: grid; grid-template-columns: 1fr 90px; gap: 10px; }
.ml-servers3 { display: grid; grid-template-columns: 1fr 80px 118px; gap: 10px; }
.ml-advanced { display: flex; flex-direction: column; gap: 11px; padding-top: 4px; border-top: 1px solid var(--line); }
.ml-search { display: flex; flex-direction: column; gap: 6px; padding: 8px 11px; border-radius: var(--r-2); background: var(--sunken); font-size: 12.5px; }
.ml-search[data-state='found'] { background: color-mix(in srgb, var(--good) 8%, var(--sunken)); }
.ml-search .ml-note { padding: 0; background: none; }
.ml-search-line { display: flex; align-items: center; gap: 8px; color: var(--ink-2); }
.ml-search-line.quiet { color: var(--ink-3); }
.ml-search-line b { color: var(--ink); font-weight: 600; }
.ml-steps-toggle { align-self: flex-start; border: 0; background: none; padding: 0; color: var(--accent); font: inherit; font-size: 12px; cursor: pointer; }
.ml-steps { list-style: none; margin: 2px 0 0; padding: 0; display: flex; flex-direction: column; gap: 3px; font-size: 12px; }
.ml-step { display: grid; grid-template-columns: 16px 150px 1fr; gap: 6px; align-items: baseline; color: var(--ink-2); }
.ml-step .name { color: var(--ink); }
.ml-step .detail { overflow-wrap: anywhere; }
.ml-step.ok svg { color: var(--good); }
.ml-step.failed svg { color: var(--bad); }
select.rw-input { appearance: auto; }
.ml-note, .ml-result {
  display: flex; align-items: center; gap: 8px; padding: 7px 11px; border-radius: var(--r-2);
  background: var(--sunken); font-size: 12px; color: var(--ink-2);
}
.ml-result.good { color: var(--good); }
.ml-result.bad { color: var(--bad); }
.ml-formats { margin: 10px 0; padding-left: 20px; font-size: 12.5px; line-height: 1.9; }
.ml-import-summary { display: flex; gap: 7px; margin: 4px 0 12px; }
.ml-import-summary .chip { background: var(--sunken); padding: 2px 9px; border-radius: 20px; font-size: 11.5px; }
.ml-import-folders { max-height: 300px; overflow: auto; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--chrome); }
.ml-import-folder { display: flex; align-items: center; gap: 8px; padding: 5px 10px; font-size: 12.5px; border-bottom: 1px solid var(--line-soft); }
.ml-import-folder:last-child { border-bottom: 0; }
.ml-import-folder .name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-import-folder .count { color: var(--ink-3); font-size: 11.5px; font-variant-numeric: tabular-nums; }

/* found on this computer ------------------------------------------------- */
.ml-found { display: flex; flex-direction: column; gap: 8px; margin: 4px 0 12px; }
.ml-found-item {
  display: flex; align-items: center; gap: 11px; padding: 10px 12px; text-align: left;
  border: 1px solid var(--line); border-radius: var(--r-2); background: var(--chrome);
  font: inherit; color: var(--ink); cursor: pointer; width: 100%;
}
.ml-found-item:hover { border-color: var(--accent-line); background: var(--selected); }
.ml-found-item .grow { flex: 1; min-width: 0; }
.ml-found-item .who { font-weight: 600; font-size: 12.5px; }
.ml-found-item .what { color: var(--ink-3); font-size: 11.5px; }
.ml-found-logo { width: 30px; height: 30px; border-radius: var(--r-1); display: grid; place-items: center; background: var(--sunken); flex: none; }

/* compose ---------------------------------------------------------------- */
.ml-compose-body { font-family: var(--font); resize: vertical; min-height: 200px; }
.ml-rich {
  min-height: 220px; max-height: 46vh; overflow: auto; padding: 10px 12px;
  border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface);
  font: 14px/1.6 var(--font); outline: none;
}
.ml-rich:focus { border-color: var(--accent-line); }
.ml-rich blockquote { margin: 0 0 0 10px; padding-left: 10px; border-left: 3px solid var(--line); color: var(--ink-3); }
.ml-toolbar { display: flex; align-items: center; gap: 2px; flex-wrap: wrap; padding: 4px 0; }
.ml-to-row { display: flex; align-items: center; gap: 8px; }
.ml-to-row .grow { flex: 1; }
.ml-chipbar { display: flex; gap: 5px; }
.ml-linkchip {
  border: 0; background: none; color: var(--ink-3); font: inherit; font-size: 11.5px;
  cursor: pointer; padding: 2px 4px; border-radius: var(--r-1);
}
.ml-linkchip:hover { color: var(--accent); background: var(--hover); }
.ml-linkchip.on { color: var(--accent); font-weight: 600; }
.ml-files { display: flex; flex-wrap: wrap; gap: 6px; }

.ml-rule-row { display: flex; align-items: center; gap: 7px; }
.ml-rule-row .rw-select, .ml-rule-row .rw-input { width: 100%; }

/* attachment and people views -------------------------------------------- */
.ml-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 10px; padding: 14px; overflow: auto; }
.ml-card {
  border: 1px solid var(--line); border-radius: var(--r-2); background: var(--chrome);
  padding: 11px 12px; text-align: left; font: inherit; color: var(--ink); cursor: pointer;
  display: flex; flex-direction: column; gap: 4px;
}
.ml-card:hover { border-color: var(--accent-line); background: var(--selected); }
.ml-card .name { font-weight: 600; font-size: 12.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-card .sub { color: var(--ink-3); font-size: 11.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-people { display: flex; flex-direction: column; overflow: auto; }
.ml-person { display: flex; align-items: center; gap: 11px; padding: 9px 16px; border-bottom: 1px solid var(--line-soft); font: inherit; color: var(--ink); background: none; border-width: 0 0 1px; text-align: left; cursor: pointer; }
.ml-person:hover { background: var(--hover); }
.ml-person .grow { flex: 1; min-width: 0; }
.ml-person .nm { font-size: 12.5px; font-weight: 600; }
.ml-person .ad { font-size: 11.5px; color: var(--ink-3); }
.ml-person .n { font-size: 11.5px; color: var(--ink-3); font-variant-numeric: tabular-nums; }

/* the outbox banner ------------------------------------------------------ */
.ml-outbox {
  position: fixed; left: 50%; transform: translateX(-50%); bottom: 34px; z-index: 60;
  display: flex; align-items: center; gap: 12px; padding: 9px 14px; border-radius: 22px;
  background: var(--ink); color: var(--surface); font-size: 12.5px;
  box-shadow: 0 6px 24px rgba(0,0,0,.28);
}
.ml-outbox button {
  border: 0; background: none; color: #7fb0ff; font: inherit; font-weight: 600;
  cursor: pointer; padding: 0;
}
`;

/** Injected once, rather than per render of a large tree. */
export function installStyles() {
  if (typeof document === 'undefined' || document.getElementById('rutba-mail-css')) return;
  const style = document.createElement('style');
  style.id = 'rutba-mail-css';
  style.textContent = CSS;
  document.head.appendChild(style);
}
