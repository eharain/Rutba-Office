// What a folder's `.rutbalocker` files say: gitignore's wording, the other way
// round. Lines say what to lock; `!` makes exceptions; a line in a subfolder's
// file wins over its parent's, because it is read later (the last match
// decides, as in git). One thing is kinder than git: an exception inside a
// folder that another line locks does work, because every rule is checked
// against the file and each of its parent folders, rather than git's "you
// cannot re-include a file under an excluded folder".
//
// With no rules file anywhere above a file, every file is locked. These are
// never locked, whatever a file says: the temporary and system files Office
// and Windows make, and the Locker's own.

const EXCLUDED = [
  { re: /^~\$/, why: 'Office\'s own temporary file' },
  { re: /\.tmp$/i, why: 'A temporary file' },
  { re: /^desktop\.ini$/i, why: 'A Windows system file' },
  { re: /^thumbs\.db$/i, why: 'A Windows system file' },
  { re: /^ehthumbs(_vista)?\.db$/i, why: 'A Windows system file' },
  { re: /^\.DS_Store$/, why: 'A macOS system file' },
  { re: /^\.dropbox(\.attr)?$/i, why: 'A sync drive\'s own file' },
];

/** The hidden folder beside a file that holds its plain copy while a program has it open. */
export const OPEN_DIR = '.rutba-open';

/** Names that are the Locker's own and are never looked at as a person's file. */
export const isLockerOwn = (name) => name === OPEN_DIR || name === '.rutbalocker' || name === '.rutbalocker-keys' || name === '.rlocker-old' || name.endsWith('.rlock') || name.endsWith('.rlock.d') || /^\.rl-[0-9a-f]{24}\.part$/.test(name);

function toRegex(glob) {
  let re = '';
  for (let i = 0; i < glob.length; i += 1) {
    const c = glob[i];
    if (c === '\\' && i + 1 < glob.length) {
      i += 1;
      re += esc(glob[i]);
    } else if (c === '*') {
      if (glob[i + 1] === '*') {
        // `**/` any folders, including none; `/**` everything below; `**` alone anything.
        i += 1;
        if (glob[i + 1] === '/') {
          i += 1;
          re += '(?:[^/]+/)*';
        } else {
          re += '.*';
        }
      } else {
        re += '[^/]*';
      }
    } else if (c === '?') {
      re += '[^/]';
    } else if (c === '[') {
      let j = i + 1;
      if (glob[j] === '!' || glob[j] === '^') j += 1;
      if (glob[j] === ']') j += 1;
      while (j < glob.length && glob[j] !== ']') j += 1;
      if (j >= glob.length) {
        re += '\\[';
      } else {
        let body = glob.slice(i + 1, j);
        if (body[0] === '!') body = `^${body.slice(1)}`;
        re += `[${body.replace(/\\/g, '\\\\')}]`;
        i = j;
      }
    } else {
      re += esc(c);
    }
  }
  return re;
}

const esc = (c) => c.replace(/[.+*?^${}()|[\]\\]/g, '\\$&');

/** One file's text as rules: lines kept in order, each with what a person wrote. */
export function parseRules(text) {
  const rules = [];
  for (const raw of String(text).split(/\r?\n/)) {
    let line = raw.replace(/(?<!\\)\s+$/, '').replace(/\\ $/, ' ');
    if (!line || line[0] === '#') continue;
    let negate = false;
    if (line[0] === '!') {
      negate = true;
      line = line.slice(1);
    } else if (line.startsWith('\\!') || line.startsWith('\\#')) {
      line = line.slice(1);
    }
    if (!line) continue;
    const dirOnly = line.endsWith('/');
    if (dirOnly) line = line.slice(0, -1);
    const anchored = line.includes('/');
    const body = line.startsWith('/') ? line.slice(1) : line;
    const re = new RegExp(`^${toRegex(body)}$`, 'i');
    rules.push({ raw: raw.trim(), negate, dirOnly, anchored, re });
  }
  return rules;
}

/** Whether a rule says something about this path or one of the folders above it. */
function ruleMatches(rule, rel, isDir) {
  const segs = rel.split('/');
  const matchAt = (path, atDir) => {
    if (rule.dirOnly && !atDir) return false;
    return rule.re.test(rule.anchored ? path : path.slice(path.lastIndexOf('/') + 1));
  };
  // The file itself, then each folder above it, innermost first.
  if (matchAt(rel, isDir)) return true;
  for (let i = segs.length - 1; i > 0; i -= 1) {
    if (matchAt(segs.slice(0, i).join('/'), true)) return true;
  }
  return false;
}

/**
 * The decision for one path (forward slashes, from the folder's top): whether
 * to lock it, and why, in the words of the rule that decided.
 *
 * @param {string} rel
 * @param {Array<{ dir: string, rules: object[] }>} sets  rules files that apply, top folder first, deepest last
 * @returns {{ lock: boolean, why: string, kind: 'rule'|'default'|'excluded' }}
 */
export function decide(rel, sets, { isDir = false } = {}) {
  const name = rel.slice(rel.lastIndexOf('/') + 1);
  for (const e of EXCLUDED) if (e.re.test(name)) return { lock: false, why: e.why, kind: 'excluded' };
  if (!sets.length) return { lock: true, why: 'Every file', kind: 'default' };
  let verdict = { lock: false, why: 'Not named by the rules', kind: 'default' };
  for (const { dir, rules } of sets) {
    const here = dir ? (rel.startsWith(`${dir}/`) ? rel.slice(dir.length + 1) : null) : rel;
    if (here === null) continue;
    for (const rule of rules) {
      if (ruleMatches(rule, here, isDir)) verdict = { lock: !rule.negate, why: rule.raw, kind: 'rule' };
    }
  }
  return verdict;
}
