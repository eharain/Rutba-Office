// Group conversations in Talk, seen in its window (Office 2.27.0): a group
// made of one's contacts, written in by three people, renamed, someone a
// contact brought who is not one's own, someone taken out, and leaving.
//
// Talk is the stand-in introducer of verify-talk.js with the groups below, a
// stand-in for the introducer's own rules (workers, introducer/src/
// conversations.js): who may be put in a group, what each person is told and
// by whom, an address shown only to a person's contacts, and "arrived" said
// once, when everyone a message went to has it. The other people are plain
// connections from this process.

import { BrowserWindow } from 'electron';
import { standInIntroducer } from './verify-talk.js';

/** Groups as the introducer keeps them, for standInIntroducer's `more`. */
export function standInGroups() {
  const groups = new Map(); // id -> { id, name, starter, created, members: [ids] }
  const waiting = new Map(); // message id -> { from, group, to: Set of ids still to have it }
  let tools = null;
  let n = 0;
  const person = (id) => ({ id, name: Object.values(tools.people).find((p) => p.id === id)?.name || '' });
  const emailOf = (id) => Object.values(tools.people).find((p) => p.id === id)?.email || '';
  const view = (g, viewer) => ({
    id: g.id, name: g.name, starter: g.starter, created: g.created,
    members: g.members.map((id) => {
      const contact = id !== viewer && tools.isContact(viewer, id);
      return { ...person(id), ...(contact || id === viewer ? { email: emailOf(id) } : {}), contact };
    }),
  });
  const tellGroup = (g, make, except = null) => { for (const id of g.members) if (id !== except) tools.to(id, make(view(g, id), id)); };
  const settle = (id) => {
    const m = waiting.get(id);
    if (!m || m.to.size) return;
    waiting.delete(id);
    n += 1;
    tools.to(m.from, { type: 'notices', notices: [{ id: `note-${n}`, kind: 'delivered', group: m.group, ids: [id], all: true }] });
  };

  const more = (msg, me, t) => {
    tools = t;
    if (!String(msg.type || '').startsWith('group.')) return false;
    const answer = (m) => t.conn.send(JSON.stringify(m));
    const fail = (code, extra = {}) => { answer({ type: 'error', for: msg.type, code, ref: msg.ref, ...extra }); return true; };
    const distinct = (ids) => [...new Set([].concat(ids || []).map(String))].filter((id) => id && id !== me.id);
    if (msg.type === 'group.create') {
      const members = distinct(msg.members);
      if (members.some((id) => !t.isContact(me.id, id))) return fail('not-a-contact');
      if (members.length < 2) return fail('too-few');
      n += 1;
      const g = { id: `g-${n}`, name: String(msg.name || '').trim(), starter: me.id, created: Date.now(), members: [me.id, ...members] };
      groups.set(g.id, g);
      answer({ type: 'group.created', ref: msg.ref, group: view(g, me.id) });
      for (const id of members) t.to(id, { type: 'group.added', group: view(g, id), by: person(me.id) });
      return true;
    }
    if (msg.type === 'group.ack') {
      for (const id of [].concat(msg.ids || [])) { waiting.get(id)?.to.delete(me.id); settle(id); }
      return true;
    }
    const g = groups.get(String(msg.groupId || ''));
    if (!g || !g.members.includes(me.id)) return fail('no-group', { id: msg.id, groupId: msg.groupId });
    if (msg.type === 'group.send') {
      const message = { id: String(msg.id), group: g.id, from: me.id, text: String(msg.text || ''), at: Date.now() };
      const to = g.members.filter((id) => id !== me.id);
      waiting.set(message.id, { from: me.id, group: g.id, to: new Set(to) });
      for (const id of to) t.to(id, { type: 'group.messages', messages: [message] });
      answer({ type: 'group.sent', id: message.id, groupId: g.id, at: message.at, to });
      return true;
    }
    if (msg.type === 'group.add') {
      const asked = distinct(msg.members).filter((id) => !g.members.includes(id));
      if (!asked.length) return fail('nobody');
      if (asked.some((id) => !t.isContact(me.id, id))) return fail('not-a-contact');
      const before = [...g.members];
      g.members.push(...asked);
      const event = { kind: 'added', by: person(me.id), who: asked.map(person), at: Date.now() };
      answer({ type: 'group.changed', ref: msg.ref, group: view(g, me.id), event });
      for (const id of before) if (id !== me.id) t.to(id, { type: 'group.changed', group: view(g, id), event });
      for (const id of asked) t.to(id, { type: 'group.added', group: view(g, id), by: person(me.id) });
      return true;
    }
    if (msg.type === 'group.rename') {
      g.name = String(msg.name || '').replace(/\s+/g, ' ').trim();
      const event = { kind: 'renamed', by: person(me.id), name: g.name, at: Date.now() };
      answer({ type: 'group.changed', ref: msg.ref, group: view(g, me.id), event });
      tellGroup(g, (seen) => ({ type: 'group.changed', group: seen, event }), me.id);
      return true;
    }
    const gone = (who) => {
      g.members = g.members.filter((id) => id !== who);
      for (const [id, m] of waiting) if (m.group === g.id) { m.to.delete(who); settle(id); }
    };
    if (msg.type === 'group.leave') {
      gone(me.id);
      answer({ type: 'group.left', ref: msg.ref, groupId: g.id });
      const event = { kind: 'left', who: [person(me.id)], at: Date.now() };
      tellGroup(g, (seen) => ({ type: 'group.changed', group: seen, event }));
      return true;
    }
    if (msg.type === 'group.remove') {
      const whom = String(msg.member || '');
      // Only the person who started a group takes someone else out of it.
      if (g.starter !== me.id) return fail('not-yours', { groupId: g.id });
      if (whom === me.id || !g.members.includes(whom)) return fail('nobody', { groupId: g.id });
      gone(whom);
      t.to(whom, { type: 'group.left', groupId: g.id, by: person(me.id) });
      const event = { kind: 'removed', by: person(me.id), who: [person(whom)], at: Date.now() };
      answer({ type: 'group.changed', ref: msg.ref, group: view(g, me.id), event });
      tellGroup(g, (seen) => ({ type: 'group.changed', group: seen, event }), me.id);
      return true;
    }
    return false;
  };
  return { more, get: (id) => groups.get(id), all: () => [...groups.values()] };
}

const windowsOf = (app) => BrowserWindow.getAllWindows().filter((w) => {
  if (w.isDestroyed()) return false;
  // A window with no address yet (one Office is still opening, or a hidden one of its own) is no app's window.
  try { return new URL(w.webContents.getURL()).searchParams.get('app') === app; } catch { return false; }
});

export async function verifyTalkGroups(h) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const kept = standInGroups();
  const intro = standInIntroducer({
    more: kept.more,
    people: { chen: { id: 'u-chen', name: 'Chen Wei', email: 'chen@example.com' }, dana: { id: 'u-dana', name: 'Dana Reyes', email: 'dana@example.com' } },
    contacts: [['u-amira', 'u-ben'], ['u-amira', 'u-chen'], ['u-ben', 'u-dana']],
  });
  const others = [];
  let spare = null;
  try {
    const port = await intro.listen();
    const url = `ws://127.0.0.1:${port}/ws`;

    // Talk windows left by the checks before this one are done with; another window keeps Office open meanwhile.
    const left = windowsOf('talk');
    spare = await open('word');
    for (const w of left) w.destroy();

    // Amira signs in with a test server in Talk's window; Ben and Chen are her contacts, Dana is Ben's alone.
    const win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const seen = (sel, what, ms = 8000) => ok(until(() => js(`Boolean(${q(sel)})`), what, ms));
    const gone = (sel, what, ms = 8000) => ok(until(() => js(`!${q(sel)}`), what, ms));
    const type = (sel, value) => js(`(() => { const n = ${q(sel)}; Object.getOwnPropertyDescriptor(n instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await js(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
    await until(() => js(`Boolean(${q('.tk-signin')})`), 'the sign-in', 8000).catch(() => {});
    await js(`${q('.tk-test')}?.setAttribute('open', ''), 1`);
    await type('.tk-server', url);
    await type('.tk-token', 'amira');
    await click('.tk-connect');
    const signedIn = await ok(until(() => js(`Boolean(${q('.tk-contact[data-contact="u-ben"]')}) && Boolean(${q('.tk-contact[data-contact="u-chen"]')})`), 'signed in, Ben and Chen among the contacts', 10000));

    // The others, each on their own connection.
    const join = async (token) => {
      const ws = new WebSocket(url);
      const heard = [];
      await new Promise((resolve, reject) => { ws.onopen = resolve; ws.onerror = () => reject(new Error(`${token} could not connect`)); });
      ws.onmessage = (e) => { try { heard.push(JSON.parse(e.data)); } catch { /* not for this check */ } };
      const who = { ws, heard, say: (m) => ws.send(JSON.stringify(m)), told: (test, what, ms = 8000) => ok(until(() => heard.some(test) || null, what, ms)) };
      others.push(who);
      who.say({ type: 'hello', token, can: ['groups', 'notices'] });
      await who.told((m) => m.type === 'welcome', `${token} signed in`);
      return who;
    };
    const ben = await join('ben');
    const chen = await join('chen');
    const dana = await join('dana');
    // As their Office does: what arrives is acknowledged, so the introducer can say it arrived.
    const acks = (who) => { who.ws.addEventListener('message', (e) => { try { const m = JSON.parse(e.data); if (m.type === 'group.messages') who.say({ type: 'group.ack', ids: m.messages.map((x) => x.id) }); } catch { /* not for this check */ } }); };
    acks(ben);

    // A new group: Start is not offered until two are chosen; made, it is in the list and open, and each of the
    // two is told who put them there.
    await click('.tk-group-new');
    const offered = await seen('.tk-new-group', 'New group');
    const pick = (id) => click(`.tk-new-group .tk-pick[data-pick="${id}"] input`);
    const notYet = await js(`${q('.tk-group-make')}?.disabled === true`);
    await pick('u-ben');
    const stillNot = await js(`${q('.tk-group-make')}?.disabled === true`);
    await pick('u-chen');
    await type('.tk-new-group .tk-group-name', 'Site visit');
    await ok(until(() => js(`${q('.tk-group-make')}?.disabled === false`), 'Start ready', 5000));
    await click('.tk-group-make');
    const made = await seen('.tk-groups .tk-contact[data-group]', 'the group in the list', 10000);
    const gid = await js(`${q('.tk-groups .tk-contact[data-group]')}?.dataset.group || ''`);
    const CONV = `.tk-conversation[data-conversation="group:${gid}"]`;
    const opened = await seen(CONV, 'the group\'s conversation open');
    const row = await js(`${q(`.tk-groups .tk-contact[data-group="${gid}"] .tk-who`)}?.textContent || ''`);
    const added = (who) => who.told((m) => m.type === 'group.added' && m.group.id === gid && m.by.id === 'u-amira' && m.group.members.length === 3, 'told of the group');
    const benAdded = await added(ben);
    const chenAdded = await added(chen);
    check('talk groups: New asks for two contacts or more before Start is offered; the group is then in the list and open, and each person put in it is told who put them there',
      signedIn && offered && notYet && stillNot && made && Boolean(gid) && opened && /Site visit/.test(row) && benAdded && chenAdded && kept.get(gid)?.members.join() === 'u-amira,u-ben,u-chen',
      JSON.stringify({ signedIn, offered, notYet, stillNot, made, gid, opened, row, benAdded, chenAdded }));

    // A message of hers goes to both. It says sent while one of them has not had it, and arrived when both have.
    const stateOf = (id) => js(`${q(`${CONV} .tk-bubble[data-message="${id}"]`)}?.dataset.state || ''`);
    await type(`${CONV} .tk-input`, 'Meet at the gate at nine.');
    await click(`${CONV} .tk-send`);
    const got = (who) => who.told((m) => m.type === 'group.messages' && m.messages.some((x) => x.group === gid && x.from === 'u-amira' && x.text === 'Meet at the gate at nine.'), 'her message');
    const benGot = await got(ben);
    const chenGot = await got(chen);
    const mid = chen.heard.find((m) => m.type === 'group.messages')?.messages[0]?.id || '';
    const sent = await ok(until(() => stateOf(mid).then((s) => s === 'sent'), 'sent', 6000));
    await wait(600);
    const waitsForChen = await stateOf(mid);
    chen.say({ type: 'group.ack', ids: [mid] });
    const arrived = await ok(until(() => stateOf(mid).then((s) => s === 'arrived'), 'arrived', 6000));
    acks(chen);
    check('talk groups: a message she writes in the group goes to everyone in it, says sent while one of them has not had it yet, and arrived once all of them have',
      benGot && chenGot && Boolean(mid) && sent && waitsForChen === 'sent' && arrived, JSON.stringify({ benGot, chenGot, mid, sent, waitsForChen, arrived }));

    // What the others write shows in the group with who wrote it; written while she looks at another conversation,
    // it is counted on the group's row until she opens it.
    ben.say({ type: 'group.send', groupId: gid, id: 'g-ben-000001', text: 'I will bring the drawings.' });
    const shown = await ok(until(() => js(`[...document.querySelectorAll(${JSON.stringify(`${CONV} .tk-bubble`)})].some((b) => b.dataset.message === 'g-ben-000001')`), 'Ben\'s message in the group', 8000));
    const said = await js(`(() => { const b = ${q(`${CONV} .tk-bubble[data-message="g-ben-000001"]`)}; return b ? { mine: b.classList.contains('mine'), text: b.textContent, around: b.parentElement?.textContent || '' } : null; })()`);
    await click('.tk-contact[data-contact="u-ben"]');
    await seen('.tk-conversation[data-conversation="u-ben"]', 'Ben\'s own conversation');
    chen.say({ type: 'group.send', groupId: gid, id: 'g-chen-00001', text: 'Nine is fine.' });
    chen.say({ type: 'group.send', groupId: gid, id: 'g-chen-00002', text: 'I have the keys.' });
    const ROW = `.tk-groups .tk-contact[data-group="${gid}"]`;
    const counted = await ok(until(() => js(`${q(`${ROW} .tk-count`)}?.textContent.trim() === '2'`), 'two not read on the group', 8000));
    const notInBens = await js(`![...document.querySelectorAll('.tk-conversation[data-conversation="u-ben"] .tk-bubble')].some((b) => /^g-chen/.test(b.dataset.message || ''))`);
    await click(ROW);
    await seen(CONV, 'the group again');
    const read = await gone(`${ROW} .tk-count`, 'the count gone');
    const both = await js(`['g-chen-00001', 'g-chen-00002'].every((id) => Boolean(document.querySelector(${JSON.stringify(CONV)} + ' .tk-bubble[data-message="' + id + '"]')))`);
    check('talk groups: what the others write shows in the group with who wrote it and in no other conversation; written while she is elsewhere it is counted on the group\'s row, and opening the group clears the count',
      shown && said?.mine === false && /I will bring the drawings\./.test(said?.text || '') && /Ben/.test(said?.around || '') && counted && notInBens && read && both,
      JSON.stringify({ shown, said: said ? { mine: said.mine, text: said.text, names: /Ben/.test(said.around) } : null, counted, notInBens, read, both }));

    // People: the three of them. Ben brings Dana, his contact and not hers: she is in the list, said not to be a
    // contact, with no address shown, and a line in the conversation says who added her.
    await click('.tk-group-people-btn');
    const listed = await seen('.tk-group-people .tk-member-list', 'the people of the group');
    const members = () => js(`[...document.querySelectorAll('.tk-group-people .tk-member')].map((m) => ({ id: m.dataset.member, says: m.querySelector('.tk-who')?.textContent || '', out: Boolean(m.querySelector('.tk-member-remove')) }))`);
    const three = await members();
    ben.say({ type: 'group.add', groupId: gid, members: ['u-dana'], ref: 'b-add' });
    const four = await ok(until(() => members().then((l) => l.length === 4), 'Dana in the list', 8000));
    const withDana = await members();
    const danaRow = withDana.find((m) => m.id === 'u-dana');
    const danaTold = await dana.told((m) => m.type === 'group.added' && m.group.id === gid && m.by.id === 'u-ben', 'Dana told of the group');
    const group = (await js(`window.rutbaOffice.talk.state()`)).groups?.find((g) => g.id === gid);
    const danaAsKept = group?.members.find((m) => m.id === 'u-dana');
    const chenRow = withDana.find((m) => m.id === 'u-chen');
    const chenAsKept = group?.members.find((m) => m.id === 'u-chen');
    check('talk groups: People lists who is in the group; someone a contact brings who is not her own contact is listed as such and no address of theirs reaches this computer, while the address of a contact of hers does',
      listed && three.length === 3 && four && Boolean(danaRow) && /Dana Reyes/.test(danaRow.says) && !/dana@example\.com/.test(JSON.stringify(withDana)) && danaAsKept?.contact === false && !('email' in (danaAsKept || {})) && danaTold && /Chen Wei/.test(chenRow?.says || '') && chenRow.says === 'Chen Wei' && chenAsKept?.contact === true && chenAsKept.email === 'chen@example.com' && danaRow.says !== 'Dana Reyes',
      JSON.stringify({ listed, three: three.map((m) => m.id), four, danaRow, danaAsKept, danaTold, chenRow, chenAsKept }));

    // The name is changed and Dana taken out, by the one who started the group: everyone else is told each, the one
    // taken out is told she is out, and the conversation keeps a line for each change.
    await type('.tk-group-people .tk-group-name', 'Site visit, Tuesday');
    await ok(until(() => js(`${q('.tk-group-rename-save')}?.disabled === false`), 'Save ready', 5000));
    await click('.tk-group-rename-save');
    const renamedHere = await ok(until(() => js(`/Site visit, Tuesday/.test(${q(`${ROW} .tk-who`)}?.textContent || '')`), 'the new name on the row', 8000));
    const benRenamed = await ben.told((m) => m.type === 'group.changed' && m.event?.kind === 'renamed' && m.event.name === 'Site visit, Tuesday' && m.event.by.id === 'u-amira', 'Ben told of the name');
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer: { response: 0 } });
    await click('.tk-group-people .tk-member[data-member="u-dana"] .tk-member-remove');
    const threeAgain = await ok(until(() => members().then((l) => l.length === 3 && !l.some((m) => m.id === 'u-dana')), 'Dana out of the list', 8000));
    const danaOut = await dana.told((m) => m.type === 'group.left' && m.groupId === gid && m.by?.id === 'u-amira', 'Dana told she is out');
    const benSawOut = await ben.told((m) => m.type === 'group.changed' && m.event?.kind === 'removed' && m.event.who[0]?.id === 'u-dana', 'Ben told Dana is out');
    const canTakeOut = three.filter((m) => m.id !== 'u-amira').every((m) => m.out) && !three.find((m) => m.id === 'u-amira')?.out;
    const lines = await js(`[...document.querySelectorAll(${JSON.stringify(`${CONV} .tk-line`)})].map((l) => l.dataset.line)`);
    check('talk groups: the person who started a group renames it and takes someone out; the others are told each change, the one taken out is told so, and the conversation keeps a line for someone added, the new name and someone taken out',
      renamedHere && benRenamed && threeAgain && danaOut && benSawOut && canTakeOut && ['added', 'renamed', 'removed'].every((k) => lines.includes(k)) && kept.get(gid)?.members.join() === 'u-amira,u-ben,u-chen',
      JSON.stringify({ renamedHere, benRenamed, threeAgain, danaOut, benSawOut, canTakeOut, lines }));

    // Leaving: asked first, then the group is off her list and the others are told she left.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer: { response: 0 } });
    await click('.tk-group-people .tk-group-leave');
    const offList = await gone(ROW, 'the group off the list', 8000);
    const dialogShut = await gone('.tk-group-people', 'People closed');
    const convShut = await gone(CONV, 'its conversation closed');
    const benSawLeave = await ben.told((m) => m.type === 'group.changed' && m.event?.kind === 'left' && m.event.who[0]?.id === 'u-amira' && m.group.members.length === 2, 'Ben told she left');
    check('talk groups: Leave group asks first, then the group is off her list with its conversation closed, and the others are told she left',
      offList && dialogShut && convShut && benSawLeave && kept.get(gid)?.members.join() === 'u-ben,u-chen', JSON.stringify({ offList, dialogShut, convShut, benSawLeave }));

    await wait(300);
    const complaints = await errorsIn(win);
    check('talk groups: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    await js(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
  } catch (err) {
    check('talk groups: the check ran to the end', false, err?.message || String(err));
  } finally {
    for (const who of others) { try { who.ws.close(); } catch { /* gone */ } }
    try { spare?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
