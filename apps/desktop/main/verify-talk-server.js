// A meeting held on a Talk Server, as far as Talk's window shows it with no
// Talk Server there: Amira is Rutba Office, Ben is plain web code in a hidden
// window, the stand-in introducer knows meetings (verify-talk-meet.js) and
// answers what Talk's service asks of it about servers by http, the way the
// real one does (workers/introducer: GET /servers/mine, POST
// /servers/meetings). What is checked is what a person sees where a meeting
// is set up and on the way into one: no switch at all at an introducer that
// does not say it knows such a meeting, whatever servers the person has; no
// switch for someone with no server;
// a line that names the server and says why when it cannot hold a meeting
// now; the switch, what it changes, and how many can then be asked; what is
// asked of the introducer at Start, the line said while a server is found,
// and the sentence when none takes the meeting; a small meeting set up all
// the same by a Talk service that holds none on servers, and said; and,
// asked into someone else's, what is said before joining, the wait for the
// person who lets people in, the asking, and being told when not let in;
// and, in Talk's settings, an owner's server with who may hold meetings on
// it, the people chosen among the contacts, and a change not taken.
// Nothing here enters a server's room: that needs a Talk Server, and is not
// seen by this check. Run alone with RUTBA_VERIFY_ONLY=talk-server.

import { BrowserWindow } from 'electron';
import { standInIntroducer } from './verify-talk.js';
import { standInMeetings } from './verify-talk-meet.js';

const MOST = 25;
const SERVER = { id: 'srv-office-1', name: 'Office server' };

/**
 * What the introducer answers about Talk Servers, for the stand-in
 * introducer (its `rest`). `mine` is what it says of the person's own
 * servers (null: it has no servers at all, and answers as for any address it
 * does not have); `carry` what it says when a meeting is brought to a
 * server; `hold` keeps that answer back until `release`.
 */
export function standInServers() {
  const calls = []; // every question as it arrived: { method, path, bearer, body }
  const state = { mine: null, carry: { status: 404, body: { error: 'no-server' } } };
  let kept = null;
  let letGo = null;
  const rest = (req, res) => {
    const at = /^\/servers\/(.*)$/.exec(req.url || '');
    if (!at) return false;
    let text = '';
    req.on('data', (d) => { text += d; });
    req.on('end', async () => {
      let body = null;
      try { body = text ? JSON.parse(text) : null; } catch { body = null; }
      calls.push({ method: req.method, path: at[1], bearer: String(req.headers.authorization || '').replace(/^Bearer /, ''), body });
      const say = (status, json) => { res.writeHead(status, { 'content-type': 'application/json' }); res.end(JSON.stringify(json)); };
      if (req.method === 'GET' && at[1] === 'mine') { if (state.mine) say(200, state.mine); else say(404, { error: 'not-found' }); return; }
      if (req.method === 'POST' && at[1] === 'meetings') {
        if (kept) await kept;
        say(state.carry.status, state.carry.body);
        return;
      }
      const share = /^mine\/([\w.:-]+)$/.exec(at[1]);
      if (req.method === 'PATCH' && share) {
        const s = state.mine?.servers?.find((x) => x.id === share[1]);
        if (!s) { say(404, { error: 'no-server' }); return; }
        if (state.noShares) { say(409, { error: 'refused' }); return; }
        s.settings = { ...s.settings, ...(body?.settings || {}) };
        say(200, { server: s });
        return;
      }
      say(404, { error: 'no-meeting' });
    });
    return true;
  };
  return {
    rest,
    calls,
    /** A change of who may hold meetings on a server is turned away from now on, or taken again. */
    refuseShares: (on = true) => { state.noShares = on === true; },
    /** What is said of the person's own servers from now on. */
    mine: (answer) => { state.mine = answer; },
    /** The answer to a meeting brought to a server is kept back. */
    hold: () => { kept = new Promise((resolve) => { letGo = resolve; }); },
    /** That answer given, as `answer`: { status, body }. */
    release: (answer) => { state.carry = answer; const go = letGo; kept = null; letGo = null; go?.(); },
  };
}

/**
 * One server of the person's own as the introducer describes it, with
 * whether a meeting of theirs would find a server now (`carry`): on their
 * own when it can hold one, on none when it cannot.
 */
const own = (more) => ({
  plan: { id: 'plan-check', maxPeople: MOST, maxServers: 1, rutbaServers: false, recording: false },
  servers: [{ ...SERVER, state: 'active', answering: true, settings: { who: 'me', named: [] }, ...more }],
  carry: more.canCarry === true ? { can: true, on: 'own', maxPeople: MOST } : { can: false, on: null, maxPeople: null },
});

// Ben here only hears and says: nothing of a meeting on a server is connected between the computers themselves.
const BEN = (url) => `<!doctype html><meta charset="utf-8"><body><script>
const log = [];
const ws = new WebSocket(${JSON.stringify(url)});
window.ben = { log, send: (m) => ws.send(JSON.stringify(m)) };
ws.onopen = () => window.ben.send({ type: 'hello', token: 'ben' });
ws.onmessage = (e) => { log.push(JSON.parse(e.data)); };
</script></body>`;

export async function verifyTalkServer(h) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const meetings = standInMeetings({ carries: true });
  const servers = standInServers();
  // Whether the introducer says in its welcome that it knows a meeting held on a server: not at first.
  let knows = false;
  const intro = standInIntroducer({
    page: BEN,
    more: meetings.more,
    rest: servers.rest,
    can: () => ['groups', 'notices', 'meetings', ...(knows ? ['carried'] : [])],
    contacts: [['u-amira', 'u-ben'], ['u-amira', 'u-cara'], ['u-amira', 'u-dev'], ['u-amira', 'u-eli'], ['u-amira', 'u-fay']],
    people: {
      cara: { id: 'u-cara', name: 'Cara Diaz', email: 'cara@example.com' },
      dev: { id: 'u-dev', name: 'Dev Patel', email: 'dev@example.com' },
      eli: { id: 'u-eli', name: 'Eli Stone', email: 'eli@example.com' },
      fay: { id: 'u-fay', name: 'Fay Okoro', email: 'fay@example.com' },
    },
  });
  let ben = null;
  let win = null;
  try {
    const port = await intro.listen();
    const url = `ws://127.0.0.1:${port}/ws`;
    win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const seen = (sel, what, ms = 8000) => ok(until(() => js(`Boolean(${q(sel)})`), what, ms));
    const gone = (sel, what, ms = 8000) => ok(until(() => js(`!${q(sel)}`), what, ms));
    const happens = (test, what, ms = 8000) => ok(until(async () => test(), what, ms));
    const set = (cls, value) => js(`(() => { const n = document.querySelector('.${cls}'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const pick = (id) => click(`.tk-pick[data-pick="${id}"] input`);
    const picked = () => js(`document.querySelectorAll('.tk-pick input:checked').length`);
    const shut = (id) => js(`${q(`.tk-pick[data-pick="${id}"] input`)}?.disabled === true`);
    const cancel = async () => { await js(`(${q('.tk-meet-start')}?.previousElementSibling?.click(), 1)`); return gone('.tk-meet-start', 'the set-up closed'); };
    const asked = (method, path) => servers.calls.filter((c) => c.method === method && c.path === path).length;
    /** Meet now opened, with what Talk's service said of her servers arrived and drawn. */
    const setUp = async () => {
      const before = asked('GET', 'mine');
      await click('.tk-meet-now');
      const there = await seen('.tk-meet-start', 'the set-up');
      const answered = await happens(() => asked('GET', 'mine') > before, 'her servers asked for');
      await wait(400);
      return there && answered;
    };
    const badToasts = () => js(`(window.__badToasts || []).slice()`);
    const plainToasts = () => js(`[...document.querySelectorAll('.rw-toast')].filter((n) => !n.classList.contains('bad')).map((n) => n.textContent.trim())`);
    const canStart = () => ok(until(() => js(`${q('.tk-meet-start')}?.disabled === false`), 'Start can be pressed', 10000));
    // The sentences this check asks for itself, kept out of the complaints at the end.
    const expected = [];

    // Amira signs in with a test server; five people are her contacts already.
    const signIn = async () => {
      await js(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
      await until(() => js(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000).catch(() => {});
      await js(`document.querySelector('.tk-test')?.setAttribute('open', ''), 1`);
      await set('tk-server', url);
      await set('tk-token', 'amira');
      await click('.tk-connect');
      await until(() => js(`/Signed in as Amira Khan/.test(document.querySelector('.tk-statusbar')?.textContent || '')`), 'signed in', 8000).catch(() => {});
      return seen('.tk-contact[data-contact="u-fay"]', 'her contacts');
    };
    const signedIn = await signIn();

    // An introducer that does not say it knows a meeting held on a server, and a server of her own that it says can
    // hold one: no switch and nothing said of servers, and three people asked as before.
    servers.mine(own({ canCarry: true, why: null }));
    const unsaidSetUp = await setUp();
    const unsaid = await js(`({ switch: Boolean(${q('.tk-meet-server')}), why: Boolean(${q('.tk-meet-server-why')}), live: Boolean(${q('.tk-meet-live')}), plain: Boolean(${q('.tk-meet-server-plain')}), words: ${q('.rw-dialog')}?.textContent || '' })`);
    await pick('u-ben'); await pick('u-cara'); await pick('u-dev');
    const threeUnsaid = (await picked()) === 3 && await shut('u-eli');
    const closed0 = await cancel();
    check('talk server meetings: at an introducer that does not say it knows a meeting held on a server, someone whose own Talk Server could hold one is offered no switch and told nothing of servers, and can ask three people as before',
      signedIn && unsaidSetUp && !unsaid.switch && !unsaid.why && unsaid.live && !unsaid.plain && unsaid.words.length > 20 && !unsaid.words.includes(SERVER.name) && threeUnsaid && closed0,
      JSON.stringify({ signedIn, unsaidSetUp, unsaid: { ...unsaid, words: unsaid.words.length }, threeUnsaid, closed0 }));

    // Nor do Talk's settings say anything of her server there.
    const settingsOpen = async () => { const before = asked('GET', 'mine'); await click('.tk-settings-btn'); const there = await seen('.tk-settings', 'Talk\'s settings'); const answered = await happens(() => asked('GET', 'mine') > before, 'her servers asked for'); await wait(400); return there && answered; };
    const settingsShut = async () => { await click('.tk-settings .rw-dialog-foot button'); return gone('.tk-settings', 'the settings closed'); };
    const openedUnsaid = await settingsOpen();
    const partUnsaid = await js(`Boolean(${q('.tk-settings .tk-own-server')}) || Boolean(${q('.tk-settings .tk-server-who')})`);
    const shutUnsaid = await settingsShut();
    check('talk server meetings: at such an introducer Talk\'s settings show nothing of her Talk Server either', openedUnsaid && !partUnsaid && shutUnsaid, JSON.stringify({ openedUnsaid, partUnsaid, shutUnsaid }));

    // From here the introducer says it: she signs in again, and is told so in its welcome.
    knows = true;
    servers.mine(null);
    const signedInAgain = await signIn();
    ben = new BrowserWindow({ show: false, webPreferences: { contextIsolation: true, nodeIntegration: false, backgroundThrottling: false, partition: 'talk-check-ben-server' } });
    await ben.loadURL(`http://127.0.0.1:${port}/ben`);
    const bjs = (code) => ben.webContents.executeJavaScript(code);
    const benHeard = (test, what, ms = 8000) => ok(until(() => bjs(`window.ben.log.some((m) => ${test})`), what, ms));
    await until(() => bjs(`window.ben.log.some((m) => m.type === 'welcome')`), 'Ben signed in', 6000).catch(() => {});

    // No server of her own (this Talk service has none): the set-up is as it always was.
    const plainSetUp = await setUp();
    const none = await js(`({ switch: Boolean(${q('.tk-meet-server')}), why: Boolean(${q('.tk-meet-server-why')}), live: Boolean(${q('.tk-meet-live')}), plain: Boolean(${q('.tk-meet-server-plain')}) })`);
    await pick('u-ben'); await pick('u-cara'); await pick('u-dev');
    const three = (await picked()) === 3 && await shut('u-eli');
    const closed1 = await cancel();
    check('talk server meetings: someone with no Talk Server is offered no switch and told nothing of servers where a meeting is set up, and can ask three people as before',
      signedInAgain && plainSetUp && !none.switch && !none.why && none.live && !none.plain && three && closed1 && servers.calls.at(-1)?.bearer === 'amira', JSON.stringify({ signedInAgain, plainSetUp, none, three, closed1, bearer: servers.calls.at(-1)?.bearer }));

    // A server of her own that cannot hold a meeting now: no switch, and a line that names it.
    servers.mine(own({ canCarry: false, why: 'paused' }));
    const pausedSetUp = await setUp();
    const paused = await js(`({ switch: Boolean(${q('.tk-meet-server')}), why: ${q('.tk-meet-server-why[data-server="no"]')}?.textContent || '', live: Boolean(${q('.tk-meet-live')}) })`);
    const closed2 = await cancel();
    check('talk server meetings: a Talk Server of her own that cannot hold a meeting now gives no switch, and a line names the server and says why',
      pausedSetUp && !paused.switch && paused.why.includes(SERVER.name) && paused.why.length > SERVER.name.length + 20 && paused.live && closed2, JSON.stringify({ pausedSetUp, paused, closed2 }));

    // A server that can: the switch, off at first; on, the count is said, more than three can be asked and the
    // meeting is not one streamed live; off again, three stay asked.
    servers.mine(own({ canCarry: true, why: null }));

    // Talk's settings name her server and who may hold meetings on it: only her at first, with nothing to type. She
    // takes "the people I choose" and ticks Ben; each change goes to Talk's service, and one it does not take is
    // said and put back as the service holds it.
    const OWN = `.tk-settings .tk-own-server[data-server="${SERVER.id}"]`;
    const shares = () => servers.calls.filter((c) => c.method === 'PATCH' && c.path === `mine/${SERVER.id}`);
    const noteOf = () => js(`${q(`${OWN} .tk-server-note`)}?.textContent || ''`);
    const openedOwn = await settingsOpen();
    const part = await js(`(() => { const n = ${q(OWN)}; return n ? { who: n.dataset.who, ways: [...n.querySelectorAll('.tk-server-who option')].map((o) => o.value), styled: n.querySelector('.tk-server-who')?.classList.contains('rw-select') === true, says: n.querySelector('.tk-field span')?.textContent || '', note: n.querySelector('.tk-server-note')?.textContent || '', people: n.querySelectorAll('.tk-server-people .tk-pick').length, typed: n.querySelectorAll('input:not([type="checkbox"])').length } : null; })()`);
    await js(`(() => { const s = ${q(`${OWN} .tk-server-who`)}; if (!s) return 0; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'named'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    const namedShown = await seen(`${OWN}[data-who="named"] .tk-server-people .tk-pick[data-pick="u-fay"]`, 'her contacts to choose among');
    const sentWho = await happens(() => shares().length === 1, 'the way sent');
    const noteNone = await noteOf();
    await click(`${OWN} .tk-pick[data-pick="u-ben"] input`);
    const sentBen = await happens(() => shares().length === 2, 'Ben sent');
    const ticked = await ok(until(() => js(`${q(`${OWN} .tk-pick[data-pick="u-ben"] input`)}?.checked === true`), 'Ben ticked', 4000));
    const noteOne = await noteOf();
    await js(`(window.__badToasts = [], 1)`);
    servers.refuseShares(true);
    await click(`${OWN} .tk-pick[data-pick="u-cara"] input`);
    const saidNo = await ok(until(() => js(`(window.__badToasts || []).some((s) => s.includes(${JSON.stringify(SERVER.name)}))`), 'the refusal said', 6000));
    const putBack = await ok(until(() => js(`${q(`${OWN} .tk-pick[data-pick="u-cara"] input`)}?.checked === false && ${q(`${OWN} .tk-pick[data-pick="u-ben"] input`)}?.checked === true`), 'put back as the Talk service holds it', 6000));
    expected.push(...await badToasts());
    servers.refuseShares(false);
    const shutOwn = await settingsShut();
    const [first, second, third] = shares();
    check('talk server meetings: Talk\'s settings name her Talk Server and who may hold meetings on it, only her at first with nothing to type; the people she chooses are ticked among her contacts and each change goes to Talk\'s service; one it does not take is said and put back',
      openedOwn && part?.who === 'me' && part.ways.join() === 'me,contacts,named,org' && part.styled && part.says.includes(SERVER.name) && part.note.length > 10 && part.people === 0 && part.typed === 0
        && namedShown && sentWho && first?.bearer === 'amira' && first.body?.settings?.who === 'named' && noteNone.length > 10 && noteNone !== part.note
        && sentBen && second?.body?.settings?.named?.join() === 'u-ben' && ticked && noteOne.includes('1') && noteOne !== noteNone
        && saidNo && third?.body?.settings?.named?.join() === 'u-ben,u-cara' && putBack && shutOwn,
      JSON.stringify({ openedOwn, part, namedShown, sentWho, first, noteNone, sentBen, second, ticked, noteOne, saidNo, third, putBack, shutOwn }));

    const heldSetUp = await setUp();
    const off = await js(`({ there: Boolean(${q('.tk-meet-server')}), on: ${q('.tk-meet-server')}?.checked, says: ${q('.tk-meet-server-switch small')}?.dataset.says, words: ${q('.tk-meet-server-switch small')}?.textContent || '', lead: ${q('.rw-dialog .tk-lead')}?.textContent || '', live: Boolean(${q('.tk-meet-live')}), plain: Boolean(${q('.tk-meet-server-plain')}) })`);
    await pick('u-ben'); await pick('u-cara'); await pick('u-dev');
    const threeOff = (await picked()) === 3 && await shut('u-eli');
    await click('.tk-meet-server');
    const turnedOn = await seen('.tk-meet-server-switch small[data-says="on"]', 'the switch on');
    const on = await js(`({ words: ${q('.tk-meet-server-switch small')}?.textContent || '', lead: ${q('.rw-dialog .tk-lead')}?.textContent || '', live: Boolean(${q('.tk-meet-live')}), plain: ${q('.tk-meet-server-plain[data-carried="yes"]')}?.textContent || '' })`);
    const eliOpen = !(await shut('u-eli'));
    await pick('u-eli'); await pick('u-fay');
    const five = await picked();
    await click('.tk-meet-server');
    const turnedOff = await seen('.tk-meet-server-switch small[data-says="off"]', 'the switch off again');
    const back = { picked: await picked(), live: await js(`Boolean(${q('.tk-meet-live')})`), plain: await js(`Boolean(${q('.tk-meet-server-plain')})`) };
    check('talk server meetings: a Talk Server that can hold a meeting gives the switch, off at first with the set-up as before; on, it says how many people the meeting holds, more than three can be asked and the streamed live switch gives way to a line; off again, three stay asked',
      heldSetUp && off.there && off.on === false && off.says === 'off' && off.words.length > 10 && !off.lead.includes(String(MOST)) && off.live && !off.plain && threeOff
        && turnedOn && on.words.includes(String(MOST)) && on.words !== off.words && on.lead.includes(String(MOST)) && !on.live && on.plain.length > 10 && eliOpen && five === 5
        && turnedOff && back.picked === 3 && back.live && !back.plain,
      JSON.stringify({ heldSetUp, off, threeOff, turnedOn, on, eliOpen, five, turnedOff, back }));

    // Start, held on the server: the introducer is asked for such a meeting, then for a server; while it is found a
    // line says so; and when no server takes it the meeting is ended, with why in a sentence that names her server.
    await click('.tk-meet-server');
    await seen('.tk-meet-server-switch small[data-says="on"]', 'the switch on again');
    await pick('u-eli');
    const fourAsked = (await picked()) === 4;
    servers.hold();
    const carriesBefore = asked('POST', 'meetings');
    const ready = await canStart();
    await click('.tk-meet-start');
    const inIt = await seen('.tk-meeting[data-state="in"][data-meeting]', 'in the meeting', 12000);
    const id = await js(`${q('.tk-meeting')}?.dataset.meeting || ''`);
    const finding = await seen('.tk-meet-step[data-step="starting"]', 'a server being found');
    const findingWords = await js(`(${q('.tk-meet-step')}?.textContent || '').trim()`);
    const made = meetings.asked.at(-1) || {};
    const brought = await happens(() => asked('POST', 'meetings') === carriesBefore + 1, 'a server asked for');
    const bring = servers.calls.filter((c) => c.method === 'POST' && c.path === 'meetings').at(-1) || {};
    const benAsked = await benHeard(`m.type === 'meet.invited' && m.meeting.id === ${JSON.stringify(id)} && m.meeting.carried === 'server'`, 'Ben asked in');
    await js(`(window.__badToasts = [], 1)`);
    servers.release({ status: 409, body: { error: 'no-server', own: [{ ...SERVER, why: 'paused' }] } });
    const out = await gone('.tk-meeting', 'out of the meeting', 10000);
    const told = await ok(until(() => js(`(window.__badToasts || []).some((s) => s.includes(${JSON.stringify(SERVER.name)}))`), 'why, in a sentence', 6000));
    const refusal = (await badToasts()).find((s) => s.includes(SERVER.name)) || '';
    expected.push(...await badToasts());
    const ended = await happens(() => !meetings.get(id), 'the meeting ended at the introducer', 6000);
    check('talk server meetings: Start asks the introducer for a meeting held on a server for that many people and tells those asked, a line says a server is being found, and when no server takes it the meeting is ended and a sentence that names her server says why',
      fourAsked && ready && inIt && Boolean(id) && finding && findingWords.length > 10 && made.carried === 'server' && made.most === MOST && made.held === true && made.live === false && made.invite?.length === 4 && made.tell === true
        && brought && bring.bearer === 'amira' && String(bring.body?.meetingId || '').startsWith(`${id}-`) && bring.body?.live === false && benAsked && out && told && refusal.length > SERVER.name.length + 20 && ended,
      JSON.stringify({ fourAsked, ready, inIt, id, finding, findingWords, made, brought, bring, benAsked, out, told, refusal, ended }));

    // A Talk service that holds no meetings on servers sets a small one up all the same: said, and no server is asked for.
    meetings.carry(false);
    await setUp();
    await click('.tk-meet-server');
    await seen('.tk-meet-server-switch small[data-says="on"]', 'the switch on');
    await pick('u-ben');
    const carriesNow = asked('POST', 'meetings');
    await js(`(window.__badToasts = [], 1)`);
    const saidBefore = await plainToasts();
    await canStart();
    await click('.tk-meet-start');
    const inSmall = await seen('.tk-meeting[data-state="in"][data-meeting]', 'in a small meeting', 12000);
    const idSmall = await js(`${q('.tk-meeting')}?.dataset.meeting || ''`);
    const said = await happens(async () => (await plainToasts()).some((s) => s.length > 20 && !saidBefore.includes(s)), 'said to be a small meeting', 6000);
    await wait(600);
    const small = { step: await js(`Boolean(${q('.tk-meet-step')})`), wanted: meetings.asked.at(-1)?.carried, held: meetings.asked.at(-1)?.held, carries: asked('POST', 'meetings') - carriesNow, bad: (await badToasts()).length };
    await click('.tk-meet-leave');
    await seen('.tk-meet-end-all', 'the question');
    await click('.tk-meet-end-all');
    const leftSmall = await gone('.tk-meeting', 'out of the small meeting');
    check('talk server meetings: a Talk service that holds no meetings on servers sets up a small meeting all the same, says so, and no server is asked for',
      inSmall && Boolean(idSmall) && said && !small.step && small.wanted === 'server' && small.held === false && small.carries === 0 && small.bad === 0 && leftSmall, JSON.stringify({ inSmall, idSmall, said, small, leftSmall }));

    // Asked into Ben's meeting, held on a server: before joining it says how it is held and that Ben lets people in;
    // in it she waits for him and no server is asked for anything; once he is in, he is asked; not let in, she is
    // told and is out of the meeting, which goes on for him.
    meetings.carry(true);
    await bjs(`window.ben.send({ type: 'meet.create', ref: 'b-make-1', name: 'All hands', live: false, carried: 'server', most: ${MOST}, invite: ['u-amira'], tell: true }), 1`);
    await benHeard(`m.type === 'meet.created' && m.ref === 'b-make-1'`, 'Ben\'s meeting');
    const id2 = await bjs(`window.ben.log.find((m) => m.type === 'meet.created' && m.ref === 'b-make-1')?.meeting.id || ''`);
    const card = await seen(`.tk-meet-card[data-meeting="${id2}"]`, 'the card');
    await click('.tk-meet-card-join');
    const looked = await seen(`.tk-meet-seen[data-meeting="${id2}"][data-can="yes"]`, 'the meeting before joining');
    const heldWords = await js(`${q('.tk-meet-seen .tk-meet-held[data-carried="yes"]')}?.textContent || ''`);
    const questions = servers.calls.length;
    await click('.tk-meet-join-go');
    const inBens = await seen(`.tk-meeting[data-state="in"][data-meeting="${id2}"]`, 'in Ben\'s meeting', 12000);
    const waiting = await seen('.tk-meet-step[data-step="waiting"]', 'waiting for Ben');
    const waitingWords = await js(`(${q('.tk-meet-step')}?.textContent || '').trim()`);
    await wait(600);
    const nothingAsked = servers.calls.length === questions;
    await bjs(`window.ben.send({ type: 'meet.join', id: ${JSON.stringify(id2)}, ref: 'b-join-1' }), 1`);
    const asking = await seen('.tk-meet-step[data-step="asking"]', 'asking Ben');
    const askingWords = await js(`(${q('.tk-meet-step')}?.textContent || '').trim()`);
    const benWasAsked = await benHeard(`m.type === 'meet.signal' && m.from === 'u-amira' && m.data && m.data.t === 'srv-ask' && typeof m.data.pub === 'string' && m.data.pub.length > 20`, 'Ben asked to let her in');
    check('talk server meetings: asked into a meeting held on a server, she is told before joining how it is held and who lets people in; in it she waits for that person with no server asked for anything, and asks them once they are in',
      card && looked && heldWords.includes('Ben Ortiz') && inBens && waiting && waitingWords.includes('Ben Ortiz') && nothingAsked && asking && askingWords.includes('Ben Ortiz') && askingWords !== waitingWords && benWasAsked,
      JSON.stringify({ card, looked, heldWords, inBens, waiting, waitingWords, nothingAsked, asking, askingWords, benWasAsked }));

    await js(`(window.__badToasts = [], 1)`);
    await bjs(`window.ben.send({ type: 'meet.signal', id: ${JSON.stringify(id2)}, to: 'u-amira', data: { t: 'srv-no', why: 'full' } }), 1`);
    const notLetIn = await gone('.tk-meeting', 'out of Ben\'s meeting', 10000);
    const toldWhy = await ok(until(() => js(`(window.__badToasts || []).some((s) => s.trim().length > 20)`), 'why, in a sentence', 6000));
    expected.push(...await badToasts());
    const benToldLeft = await benHeard(`m.type === 'meet.person' && m.event === 'left' && m.person.id === 'u-amira'`, 'Ben told she left');
    const goesOn = meetings.inside(id2).join() === 'u-ben' && Boolean(meetings.get(id2));
    check('talk server meetings: not let in, she is told why in a sentence and is out of the meeting, which goes on for the person who set it up, and still no server was asked for anything',
      notLetIn && toldWhy && benToldLeft && goesOn && servers.calls.length === questions, JSON.stringify({ notLetIn, toldWhy, benToldLeft, goesOn, inside: meetings.inside(id2), questions: servers.calls.length - questions }));

    await wait(300);
    const complaints = (await errorsIn(win)).filter((s) => !expected.includes(s));
    check('talk server meetings: none of that reported an error but the three sentences asked for', complaints.length === 0 && expected.length === 3, [...complaints, `asked for: ${expected.length}`].join(' | '));
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`).catch(() => {});
  } catch (err) {
    check('talk server meetings: the check ran to the end', false, err?.message || String(err));
  } finally {
    try { ben?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
