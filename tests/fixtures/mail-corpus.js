// A large mailbox, made up the same way every time.
//
// For the archive's performance budget: tens of thousands of messages over
// years and folders, with bodies of ordinary length drawn from a vocabulary
// of a few thousand words (common ones often, rare ones rarely, as in real
// mail), a few hundred correspondents, and attachments carrying most of the
// bytes, as they do in a real .pst. Nothing random from run to run: a seeded
// generator, so a search's answer is the same each time.

const SYLLABLES = 'ka lo mi ne ra su te vo ba di fe gu ha ji ko lu ma no pi qu re si to vu wa xe yo za an el in or us ar en il on ur'.split(' ');
const COMMON = (
  'the of and to in is for on that with as this by be are from at it was or an have not will can all your our we they has more ' +
  'meeting report contract invoice budget project review draft schedule quarter figures payment order delivery client supplier ' +
  'office team plan agenda notes update please thanks regards attached following week month year today tomorrow monday friday'
).split(' ');

/** A small, fast, seeded generator (mulberry32). */
export function rng(seed) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/**
 * Messages, one at a time.
 * @param {{ count: number, seed?: number, from?: number, years?: number, vocabulary?: number,
 *   attachEvery?: number, attachBytes?: number, bodyWords?: number }} opts
 * @yields {{ folder: string, message: object }}
 */
export function* corpus({ count, seed = 7, from = Date.UTC(2014, 0, 1), years = 10, vocabulary = 6000, attachEvery = 4, attachBytes = 60_000, bodyWords = 160 }) {
  const random = rng(seed);
  const pick = (list) => list[Math.floor(random() * list.length)];
  const rare = [];
  for (let i = 0; i < vocabulary; i++) {
    const n = 2 + Math.floor(random() * 3);
    let w = '';
    for (let k = 0; k < n; k++) w += pick(SYLLABLES);
    rare.push(w);
  }
  // Zipf-like: word i is drawn about 1/(i+1) as often as the first.
  const word = () => {
    if (random() < 0.45) return pick(COMMON);
    const r = random();
    return rare[Math.min(rare.length - 1, Math.floor(Math.exp(r * Math.log(rare.length + 1)) - 1))];
  };
  const people = Array.from({ length: 400 }, (_, i) => {
    const first = rare[(i * 7) % rare.length];
    const last = rare[(i * 13 + 5) % rare.length];
    const name = `${first[0].toUpperCase()}${first.slice(1)} ${last[0].toUpperCase()}${last.slice(1)}`;
    return { name, address: `${first}.${last}@${['northwind', 'contoso', 'fabrikam', 'tailspin', 'adatum'][i % 5]}.example` };
  });
  const me = { name: 'Archive Owner', address: 'owner@example.com' };
  const folders = ['Inbox', 'Inbox/Clients', 'Inbox/Suppliers', 'Inbox/Projects', 'Inbox/Projects/Harbour', 'Sent Items', 'Archive', 'Archive/2014-2017', 'Finance', 'Legal'];
  const span = years * 365.25 * 86_400_000;
  // One pool of bytes that attachments are cut from, so they differ without
  // the cost of generating each one.
  const pool = Buffer.alloc(attachBytes * 4 + 65_536);
  for (let i = 0; i < pool.length; i += 4) pool.writeUInt32LE((random() * 4294967296) >>> 0, i);
  const kinds = [['pdf', 'application/pdf'], ['xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'], ['docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'], ['jpg', 'image/jpeg']];

  for (let i = 0; i < count; i++) {
    const folder = folders[i % 17 < 9 ? 0 : (i % folders.length)];
    const sent = folder === 'Sent Items';
    const other = pick(people);
    const words = Math.max(20, Math.round(bodyWords * (0.4 + random() * 1.2)));
    const body = [];
    for (let w = 0; w < words; w++) body.push(word());
    const subjectWords = [word(), word(), word(), word()].join(' ');
    const attachments = [];
    if (attachEvery && i % attachEvery === 0) {
      const size = Math.round(attachBytes * (0.5 + random()));
      const at = Math.floor(random() * (pool.length - size));
      const [ext, type] = kinds[i % kinds.length];
      attachments.push({ filename: `${word()}-${i}.${ext}`, type, bytes: pool.subarray(at, at + size) });
    }
    yield {
      folder,
      message: {
        subject: `${subjectWords[0].toUpperCase()}${subjectWords.slice(1)} ${i % 50 === 0 ? 'invoice' : ''}`.trim(),
        from: sent ? me : other,
        to: sent ? [other] : [me],
        cc: i % 5 === 0 ? [pick(people)] : [],
        date: new Date(from + (span * i) / count).toISOString(),
        body: `${body.join(' ')}.\n\nRef ${i}`,
        messageId: `m${i}.${seed}@corpus.example`,
        unread: i % 9 === 0,
        attachments,
      },
    };
  }
}

/** The folder paths `corpus` files messages under. */
export const CORPUS_FOLDERS = ['Inbox', 'Inbox/Clients', 'Inbox/Suppliers', 'Inbox/Projects', 'Inbox/Projects/Harbour', 'Sent Items', 'Archive', 'Archive/2014-2017', 'Finance', 'Legal'];
