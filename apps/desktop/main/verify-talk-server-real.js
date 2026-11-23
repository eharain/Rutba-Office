// A meeting held on a real Talk Server, in Office's own Talk window, between
// two copies of Office running on this computer at once: one starts the
// meeting, the other is asked in and joins. Nothing here is a stand-in. The
// introducer and the Talk Server are real programs somebody started before
// the run and named to it, and each copy signs in to that introducer the way
// the window checks sign in to a stand-in: a token and an address given
// outright (talk.js, testServers).
//
// What is seen is what the stand-in check (verify-talk-server.js) says it
// cannot see: the server's room entered from Talk's window, with the
// meeting's engine loaded by that window, the way through the server's door
// made by Talk's service, and the other person's picture, sound and written
// words arriving through the server and opening with the meeting's key.
//
// Never in a whole run and never in the gate: it needs a real introducer and
// a real Talk Server, which Office's folder does not hold. Run alone with
// RUTBA_VERIFY_ONLY=talk-server-real and RUTBA_TALK_REAL naming a file:
//
//   { "role": "starts" | "joins", "server": "ws://localhost:<port>/ws",
//     "token": "<sign-in token>", "dir": "<a folder both copies write notes in>",
//     "most": <how many people the plan holds>, "serverName": "<the server's name>" }
//
// The copy that starts must own a Talk Server at that introducer that can
// hold a meeting now. The two copies tell each other where they are with
// small files in `dir`: the invitation link, that the one who joins is
// ready, and that they have seen what they came to see. Without the file a
// run by that name fails at its first check and does nothing else.

import fs from 'node:fs';
import path from 'node:path';

const ENV = 'RUTBA_TALK_REAL';
/** How long one copy waits for the other to get somewhere. */
const OTHER = 240000;

function readRun() {
  try {
    const run = JSON.parse(fs.readFileSync(String(process.env[ENV] || ''), 'utf8'));
    if (!['starts', 'joins'].includes(run.role) || !/^wss?:\/\//.test(String(run.server || '')) || !run.token || !run.dir || !fs.statSync(run.dir).isDirectory()) return null;
    return run;
  } catch {
    return null;
  }
}

export async function verifyTalkServerReal(h) {
  const { open, check, until, wait, errorsIn } = h;
  const run = readRun();
  if (!run) {
    check(`talk server, real: the run was given a real introducer, a sign-in, its part and a folder for notes (${ENV})`, false, `${ENV} names no such file, or the file does not say them`);
    return;
  }
  const P = `talk server, real (${run.role})`;
  const ok = (p) => p.then(() => true).catch(() => false);
  /** A note for the other copy, there whole or not at all. */
  const leave = (name, what) => { const file = path.join(run.dir, name); fs.writeFileSync(`${file}.part`, JSON.stringify(what)); fs.renameSync(`${file}.part`, file); };
  /** The other copy's note, once it is there. */
  const note = async (name, what, ms = OTHER) => {
    const file = path.join(run.dir, name);
    try { await until(() => fs.existsSync(file), what, ms); return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return null; }
  };
  let win = null;
  try {
    win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const seen = (sel, what, ms = 10000) => ok(until(() => js(`Boolean(${q(sel)})`), what, ms));
    const gone = (sel, what, ms = 10000) => ok(until(() => js(`!${q(sel)}`), what, ms));
    const talk = (method, payload = {}) => js(`window.rutbaOffice.talk.${method}(${JSON.stringify(payload)})`);

    // Everything the window says in a toast, and every step of the way into the server's room, kept as it happens.
    await js(`(() => {
      window.__said = []; window.__steps = [];
      const look = () => {
        for (const n of document.querySelectorAll('.rw-toast')) { const s = n.textContent.trim(); if (s && !window.__said.includes(s)) window.__said.push(s); }
        const step = document.querySelector('.tk-meet-step')?.dataset.step;
        if (step && window.__steps.at(-1) !== step) window.__steps.push(step);
      };
      new MutationObserver(look).observe(document.body, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ['data-step'] });
      return 1;
    })()`);

    /** In the meeting, through every step, with the server's room entered. */
    const intoTheRoom = async (ms) => {
      const inMeeting = await seen('.tk-meeting[data-state="in"][data-meeting]', 'in the meeting', 20000);
      const entered = await ok(until(() => js(`Boolean(${q('.tk-meeting[data-state="in"]')}) && !${q('.tk-meet-step')} && window.__steps.length > 0`), 'the server\'s room entered', ms));
      return { inMeeting, entered, steps: await js(`window.__steps.slice()`) };
    };
    /** Said to be encrypted, once somebody else's sound and picture arrive too. */
    const saidEncrypted = () => seen('.tk-meet-held.in[data-carried="yes"]', 'said to be encrypted', 20000);
    const cameraOn = async () => { await click('.tk-device-camera'); return seen('.tk-devices[data-camera="on"]', 'the camera on', 15000); };
    /**
     * What arrives of one person in their tile: a picture that has a size,
     * moves and is not black, and sound that is not silence. Read off the
     * tile's own stream, as the window plays it.
     */
    const arrives = async (id) => {
      const tile = `.tk-tile[data-person="${id}"]`;
      const there = await seen(`${tile}[data-connected="yes"]`, 'their tile, connected', OTHER);
      await until(() => js(`(${q(`${tile} video`)}?.videoWidth || 0) > 0`), 'their picture', 60000).catch(() => {});
      const got = await js(`(async () => {
        const v = ${q(`${tile} video`)};
        const s = v?.srcObject;
        if (!v || !s) return null;
        const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
        const t0 = v.currentTime;
        let loud = 0;
        try {
          const ctx = new AudioContext();
          await ctx.resume();
          const an = ctx.createAnalyser();
          an.fftSize = 2048;
          ctx.createMediaStreamSource(new MediaStream(s.getAudioTracks())).connect(an);
          const buf = new Uint8Array(an.fftSize);
          const end = Date.now() + 5000;
          while (Date.now() < end && loud < 8) { an.getByteTimeDomainData(buf); for (const b of buf) loud = Math.max(loud, Math.abs(b - 128)); await sleep(40); }
          await ctx.close();
        } catch (err) { loud = -1; }
        await sleep(300);
        let lit = 0;
        try {
          const c = document.createElement('canvas');
          c.width = 64; c.height = 48;
          const g = c.getContext('2d');
          g.drawImage(v, 0, 0, 64, 48);
          for (const b of g.getImageData(0, 0, 64, 48).data) lit = Math.max(lit, b === 255 ? 0 : b);
        } catch (err) { lit = -1; }
        return { width: v.videoWidth, height: v.videoHeight, moves: v.currentTime > t0, shown: v.classList.contains('tk-tile-video'), sounds: s.getAudioTracks().length, pictures: s.getVideoTracks().length, loud, lit };
      })()`);
      return { there, ...(got || { none: true }), good: Boolean(there && got && got.width > 0 && got.moves && got.shown && got.pictures > 0 && got.sounds > 0 && got.loud >= 8 && got.lit > 0) };
    };
    const write = async (text) => {
      await js(`(() => { const n = ${q('.tk-meet-input')}; if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(n, ${JSON.stringify(text)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
      await until(() => js(`${q('.tk-meet-send')}?.disabled === false`), 'Send can be pressed', 5000).catch(() => {});
      await click('.tk-meet-send');
      return ok(until(() => js(`[...document.querySelectorAll('.tk-meet-says .tk-bubble.mine .tk-text')].some((n) => n.textContent === ${JSON.stringify(text)})`), 'her own words shown', 8000));
    };
    const reads = (text, ms = OTHER) => ok(until(() => js(`[...document.querySelectorAll('.tk-meet-says .tk-bubble:not(.mine) .tk-text')].some((n) => n.textContent === ${JSON.stringify(text)})`), 'the other\'s words', ms));
    const messagesOpen = async () => { if (!await js(`Boolean(${q('.tk-meet-input')})`)) await click('.tk-meet-messages'); return seen('.tk-meet-input', 'the messages'); };

    // Signed in to the real introducer, with a token and its address given outright.
    await js(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
    await until(() => js(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 10000).catch(() => {});
    const refused = await js(`window.rutbaOffice.talk.signIn(${JSON.stringify({ token: run.token, server: run.server })}).then(() => '', (err) => String(err?.message || err))`);
    const signedIn = await seen('.tk-add', 'signed in', 20000);
    const me = (await talk('state'))?.me || null;
    const FROM_STARTER = 'Can you read this, held on a Talk Server?';
    const FROM_JOINER = 'Read here, and this is the answer.';

    if (run.role === 'starts') {
      // Her invitation link makes the other person a contact, as two people become contacts anywhere.
      const made = await talk('inviteLink', { note: 'For a meeting held on a Talk Server' });
      const link = made?.invite?.url || made?.invite?.code || '';
      if (link) leave('invite.json', { link, me });
      const ready = await note('joins-ready.json', 'the other copy of Office to be ready');
      const other = ready?.me || null;
      const contact = other ? await seen(`.tk-contact[data-contact="${other.id}"]`, 'the other person among her contacts', 20000) : false;
      check(`${P}: she signs in to a real introducer, and the person who took her invitation link in another copy of Office is among her contacts`,
        !refused && signedIn && Boolean(me?.id) && Boolean(link) && Boolean(other?.id) && contact, JSON.stringify({ refused, signedIn, me, link: Boolean(link), other, contact }));

      // Talk's settings name the server she runs: only she may hold meetings on it.
      await click('.tk-settings-btn');
      const partShown = await seen('.tk-settings .tk-own-server[data-server]', 'her server in Talk\'s settings', 15000);
      const part = await js(`(() => { const n = ${q('.tk-settings .tk-own-server[data-server]')}; return n ? { servers: document.querySelectorAll('.tk-settings .tk-own-server').length, who: n.dataset.who, says: n.querySelector('.tk-field span')?.textContent || '' } : null; })()`);
      await click('.tk-settings .rw-dialog-foot button');
      const partShut = await gone('.tk-settings', 'the settings closed');
      check(`${P}: Talk's settings name the Talk Server she runs, as the real introducer describes it, with only her to hold meetings on it`,
        partShown && part?.servers === 1 && part.who === 'me' && (!run.serverName || part.says.includes(run.serverName)) && partShut, JSON.stringify({ partShown, part, partShut }));

      // Where a meeting is set up her server gives the switch, for as many people as her plan holds.
      await click('.tk-meet-now');
      const setUp = await seen('.tk-meet-start', 'the set-up');
      const offered = await seen('.tk-meet-server', 'the switch', 15000);
      await click('.tk-meet-server');
      const turnedOn = await seen('.tk-meet-server-switch small[data-says="on"]', 'the switch on', 15000);
      const says = await js(`${q('.tk-meet-server-switch small')}?.textContent || ''`);
      const camera = await cameraOn();
      await click(`.tk-pick[data-pick="${other?.id}"] input`);
      const asked = await js(`document.querySelectorAll('.tk-pick input:checked').length`);
      const canStart = await ok(until(() => js(`${q('.tk-meet-start')}?.disabled === false`), 'Start can be pressed', 15000));
      check(`${P}: where a meeting is set up her Talk Server gives the switch, and turned on it says how many people her plan holds`,
        setUp && offered && turnedOn && (!run.most || says.includes(String(run.most))) && camera && asked === 1 && canStart, JSON.stringify({ setUp, offered, turnedOn, says, most: run.most || null, camera, asked, canStart }));

      // Start: a server is asked for, its room is entered from this window, and she is in it with her own picture.
      await click('.tk-meet-start');
      const room = await intoTheRoom(90000);
      const id = await js(`${q('.tk-meeting')}?.dataset.meeting || ''`);
      const mine = await seen(`.tk-tile[data-person="${me?.id}"]`, 'her own tile');
      const noFault = !await js(`Boolean(${q('.tk-device-fault')})`);
      check(`${P}: Start asks for a server and Talk's window enters its room: she is in the meeting, with her microphone and camera and nothing said to be wrong`,
        room.inMeeting && room.entered && room.steps[0] === 'starting' && Boolean(id) && mine && noFault, JSON.stringify({ ...room, id, mine, noFault, faults: await js(`[...document.querySelectorAll('.tk-device-fault')].map((n) => n.dataset.fault + ': ' + n.textContent)`) }));

      // The other person comes in, let in by this copy with nobody pressing anything: picture and sound through the server.
      const theirs = await arrives(other?.id);
      const encrypted = await saidEncrypted();
      check(`${P}: the other person comes into the meeting from their copy of Office, let in by hers with nothing pressed; their picture and their sound arrive through the server and open with the meeting's key, and the meeting is said to be encrypted`,
        theirs.good && encrypted, JSON.stringify({ theirs, encrypted }));

      // What is written goes through the server too, sealed with the meeting's key.
      const side = await messagesOpen();
      const wrote = await write(FROM_STARTER);
      const answered = await reads(FROM_JOINER);
      check(`${P}: what she writes in the meeting is read in the other copy, and the answer written there is read in hers`, side && wrote && answered, JSON.stringify({ side, wrote, answered }));

      // Once the other copy has seen all it came to see, she ends the meeting for everyone.
      const sawIt = await note('joins-seen.json', 'the other copy to have seen it all');
      await click('.tk-meet-leave');
      const question = await seen('.tk-meet-end-all', 'leave, or end for everyone');
      await click('.tk-meet-end-all');
      const out = await gone('.tk-meeting', 'out of the meeting', 20000);
      check(`${P}: her red button asks, and End for everyone ends the meeting held on the server`, Boolean(sawIt) && question && out, JSON.stringify({ sawIt, question, out }));
    } else {
      // The other person's link is taken here: they are contacts.
      const invite = await note('invite.json', 'the invitation link from the other copy of Office');
      const starter = invite?.me || null;
      const taken = invite?.link ? await talk('acceptInvite', { link: invite.link }) : null;
      const contact = starter ? await seen(`.tk-contact[data-contact="${starter.id}"]`, 'the other person among his contacts', 20000) : false;
      check(`${P}: he signs in to the same real introducer from another copy of Office, takes her invitation link, and she is among his contacts`,
        !refused && signedIn && Boolean(me?.id) && Boolean(starter?.id) && taken?.contact?.id === starter?.id && contact, JSON.stringify({ refused, signedIn, me, starter, taken, contact }));

      // He runs no server and nobody chose him for theirs: nothing is offered and nothing is said of servers.
      await click('.tk-meet-now');
      const setUp = await seen('.tk-meet-start', 'the set-up');
      await wait(3000);
      const nothing = await js(`({ switch: Boolean(${q('.tk-meet-server')}), why: Boolean(${q('.tk-meet-server-why')}), live: Boolean(${q('.tk-meet-live')}) })`);
      await js(`(${q('.tk-meet-start')}?.previousElementSibling?.click(), 1)`);
      const shut = await gone('.tk-meet-start', 'the set-up closed');
      await click('.tk-settings-btn');
      const settings = await seen('.tk-settings', 'Talk\'s settings');
      await wait(3000);
      const partShown = await js(`Boolean(${q('.tk-settings .tk-own-server')}) || Boolean(${q('.tk-settings .tk-server-who')})`);
      await click('.tk-settings .rw-dialog-foot button');
      const settingsShut = await gone('.tk-settings', 'the settings closed');
      check(`${P}: a contact of someone who runs a Talk Server, not chosen for it, is offered no switch where a meeting is set up and sees nothing of servers in Talk's settings`,
        setUp && !nothing.switch && !nothing.why && nothing.live && shut && settings && !partShown && settingsShut, JSON.stringify({ setUp, nothing, shut, settings, partShown, settingsShut }));
      leave('joins-ready.json', { me });

      // Asked in: before joining he reads how the meeting is held and who lets people in.
      const card = await seen('.tk-meet-card[data-meeting]', 'asked into the meeting', OTHER);
      const id = await js(`${q('.tk-meet-card[data-meeting]')}?.dataset.meeting || ''`);
      await click('.tk-meet-card-join');
      const looked = await seen(`.tk-meet-seen[data-meeting="${id}"][data-can="yes"]`, 'the meeting before joining', 15000);
      const heldWords = await js(`${q('.tk-meet-seen .tk-meet-held[data-carried="yes"]')}?.textContent || ''`);
      const camera = await cameraOn();
      check(`${P}: asked into her meeting, he reads before joining that it is held on a server and that she lets people in`,
        card && Boolean(id) && looked && Boolean(starter?.name) && heldWords.includes(starter.name) && camera, JSON.stringify({ card, id, looked, heldWords, camera }));

      // Join: he asks her copy, is let in by it, and enters the server's room from this window.
      await click('.tk-meet-join-go');
      const room = await intoTheRoom(120000);
      const noFault = !await js(`Boolean(${q('.tk-device-fault')})`);
      check(`${P}: Join asks the person who set the meeting up, whose copy lets him in, and Talk's window enters the server's room: he is in the meeting with nothing said to be wrong`,
        room.inMeeting && room.entered && room.steps.includes('entering') && noFault, JSON.stringify({ ...room, noFault, faults: await js(`[...document.querySelectorAll('.tk-device-fault')].map((n) => n.dataset.fault + ': ' + n.textContent)`) }));

      const theirs = await arrives(starter?.id);
      const encrypted = await saidEncrypted();
      check(`${P}: her picture and her sound arrive through the server and open with the meeting's key, and the meeting is said to be encrypted`, theirs.good && encrypted, JSON.stringify({ theirs, encrypted }));

      const side = await messagesOpen();
      const read = await reads(FROM_STARTER);
      const wrote = await write(FROM_JOINER);
      check(`${P}: what she wrote in the meeting is read here, and he writes his answer`, side && read && wrote, JSON.stringify({ side, read, wrote }));
      leave('joins-seen.json', { theirs: theirs.good, read, wrote });

      // She ends it for everyone: he is out, and told who ended it.
      const out = await gone('.tk-meeting', 'out of the meeting once she ended it', OTHER);
      const told = await ok(until(() => js(`window.__said.some((s) => s.includes(${JSON.stringify(starter?.name || '\u0000')}))`), 'told who ended it', 8000));
      check(`${P}: when she ends the meeting for everyone he is out of it and is told who ended it`, out && told, JSON.stringify({ out, told, said: await js(`window.__said.slice(-4)`) }));
    }

    await wait(500);
    const complaints = await errorsIn(win);
    check(`${P}: none of that reported an error`, complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`).catch(() => {});
  } catch (err) {
    check(`${P}: the check ran to the end`, false, err?.message || String(err));
  }
}
