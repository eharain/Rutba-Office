// Choose apps (owner, 2026-10-20: "since the number of applications have
// grown we should let users to choose what they wanted to install and what
// not"), checked in the launcher: on first start the eight Office apps and a
// card offering more (D-2026-10-20-19: the further apps and AI Tools are
// added by choice), the Choose apps window in its groups, an app turned off
// leaving the tiles and the templates and one added joining them, AI Tools
// added, the choice kept for the next window, and every app on again.

export async function verifyAppsChoice(h) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  let js = null;
  try {
    let home = await open('home');
    js = (code) => home.webContents.executeJavaScript(code);
    await js(`window.rutbaOffice.store.delete({ key: 'apps.choice' }).then(() => 1)`);
    await home.webContents.reload();
    await until(() => js(`document.querySelectorAll('.home-card').length > 0`), 'the launcher', 10000);
    const card = await ok(until(() => js(`Boolean(document.querySelector('.home-choose'))`), 'the first-start card', 5000));
    const tiles = () => js(`[...document.querySelectorAll('.home-card')].map((c) => c.dataset.app)`);
    const all = await tiles();
    check('apps choice: on first start the eight Office apps are on, Image, Video and Recorder are not, and a card offers to choose',
      card && all.length === 8 && all.includes('sheets') && all.includes('pictures') && !all.includes('image') && !all.includes('video') && !all.includes('recorder') && !all.includes('talk'), JSON.stringify({ card, all }));

    await js(`document.querySelector('.home-choose-btn')?.click(), 1`);
    const shown = await ok(until(() => js(`document.querySelectorAll('.apps-chooser .apps-tile').length === 13`), 'Choose apps', 5000));
    const groups = await js(`[...document.querySelectorAll('.apps-chooser .apps-group')].map((g) => g.dataset.group + ':' + g.querySelectorAll('.apps-tile').length).join(' ')`);
    const onAtFirst = await js(`document.querySelectorAll('.apps-chooser .apps-tile.on').length`);
    await js(`(() => { for (const k of ['sheets', 'recorder', 'ai']) document.querySelector('.apps-chooser .apps-tile[data-app="' + k + '"]')?.click(); return 1; })()`);
    const count = await js(`document.querySelector('.apps-count')?.textContent || ''`);
    await js(`document.querySelector('.apps-done')?.click(), 1`);
    const gone = await ok(until(() => js(`!document.querySelector('.apps-chooser') && !document.querySelector('.home-card[data-app="sheets"]')`), 'the tiles without Worksheets', 4000));
    const after = await tiles();
    const templates = await js(`[...document.querySelectorAll('.home-template .tpl-glyph')].map((g) => g.dataset.app).filter(Boolean)`);
    const recordShown = await js(`Boolean(document.querySelector('.home-record'))`);
    const cardGone = await js(`!document.querySelector('.home-choose')`);
    const stored = await js(`window.rutbaOffice.store.get({ key: 'apps.choice', fallback: null })`);
    check('apps choice: Choose apps shows Office apps, More apps and Add-ons with the Office apps on; Worksheets off and Recorder and AI Tools added change the tiles and templates, and the card goes',
      shown && groups === 'office:8 more:4 addons:1' && onAtFirst === 8 && /8 of 12/.test(count) && gone && after.length === 8 && !after.includes('sheets') && after.includes('recorder') && !after.includes('video')
        && !templates.includes('sheets') && templates.includes('word') && !recordShown && cardGone
        && JSON.stringify(stored) === JSON.stringify({ off: ['sheets'], on: ['recorder'], addons: ['ai'] }),
      JSON.stringify({ shown, groups, onAtFirst, count, after, templates, recordShown, cardGone, stored }));

    // The template gallery from the launcher shows only the apps chosen.
    await js(`document.querySelector('.home-more')?.click(), 1`);
    const gallery = await ok(until(() => js(`document.querySelectorAll('.tg-thumb').length > 0`), 'the template gallery', 6000));
    const kinds = await js(`[...new Set([...document.querySelectorAll('.tg-thumb')].map((n) => [...n.classList].find((c) => ['word', 'sheets', 'slides'].includes(c))).filter(Boolean))]`);
    check('apps choice: the template gallery from the launcher leaves out the templates of an app turned off',
      gallery && kinds.includes('word') && kinds.includes('slides') && !kinds.includes('sheets'), JSON.stringify({ gallery, kinds }));

    // Another launcher keeps the choice; All apps brings everything back.
    const first = home;
    home = await open('home');
    first.destroy();
    js = (code) => home.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.home-card').length > 0`), 'the launcher again', 10000);
    await wait(300);
    const kept = await tiles();
    const noCard = await js(`!document.querySelector('.home-choose')`);
    await js(`document.querySelector('.home-apps-btn')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.apps-chooser'))`), 'Choose apps again', 4000).catch(() => {});
    await js(`document.querySelector('.apps-all')?.click(), 1`);
    await js(`document.querySelector('.apps-done')?.click(), 1`);
    const back = await ok(until(() => js(`document.querySelectorAll('.home-card').length === 12`), 'every app back', 4000));
    const addonsKept = await js(`window.rutbaOffice.store.get({ key: 'apps.choice', fallback: null }).then((c) => c && c.addons)`);
    check('apps choice: a new launcher window keeps the choice without asking again, and All apps turns every app on, keeping the add-ons',
      kept.length === 8 && !kept.includes('sheets') && kept.includes('recorder') && noCard && back && JSON.stringify(addonsKept) === '["ai"]', JSON.stringify({ kept, noCard, back, addonsKept }));

    const complaints = await errorsIn(home);
    check('apps choice: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('apps choice: the check ran to the end', false, err?.message || String(err));
  } finally {
    // Every app and add-on on for the checks that follow, as a check run starts (main.js).
    await js?.(`window.rutbaOffice.store.set({ key: 'apps.choice', value: { off: [], on: ['image', 'video', 'recorder', 'talk'], addons: ['ai'] } }).then(() => 1)`).catch(() => {});
  }
}
