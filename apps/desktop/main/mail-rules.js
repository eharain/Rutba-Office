// Rules.
//
// The oldest feature in mail and still the one people ask for first: when a
// message arrives, put it where it belongs without me. Outlook has it,
// Thunderbird has it, and a client without it makes every arriving message a
// small decision.
//
// A rule is conditions and actions. It runs when mail arrives, and can be run
// over a folder on demand — because the reason anybody writes a rule is the
// four thousand messages already sitting there.
//
// Two decisions worth stating:
//
//   Rules run against the stored header, not the body. The index has the
//   sender, the subject, the preview and the flags, and running a rule over
//   50,000 messages must not mean opening 50,000 files. A rule that needs the
//   whole body is a rule that takes a minute, and this is not that.
//
//   A rule never deletes outright unless it says so in those words. "Delete"
//   moves to Trash, which is what every mail client means by it and what
//   anybody writing a rule at speed assumes.

/** What a condition can look at. Each reads from the index row. */
const FIELDS = {
  from: (row) => `${row.from?.name || ''} ${row.from?.address || ''}`,
  fromAddress: (row) => row.from?.address || '',
  to: (row) => (row.to || []).map((t) => `${t.name || ''} ${t.address || ''}`).join(' '),
  subject: (row) => row.subject || '',
  body: (row) => row.preview || '',
  folder: (row) => row.folder || '',
  list: (row) => row.listId || '',
};

/** How a condition compares. */
const OPERATORS = {
  contains: (value, want) => value.toLowerCase().includes(want.toLowerCase()),
  notContains: (value, want) => !value.toLowerCase().includes(want.toLowerCase()),
  is: (value, want) => value.trim().toLowerCase() === want.trim().toLowerCase(),
  isNot: (value, want) => value.trim().toLowerCase() !== want.trim().toLowerCase(),
  startsWith: (value, want) => value.toLowerCase().startsWith(want.toLowerCase()),
  endsWith: (value, want) => value.toLowerCase().endsWith(want.toLowerCase()),
  matches: (value, want) => {
    // A rule with a bad expression must not stop every other rule from running.
    try {
      return new RegExp(want, 'i').test(value);
    } catch {
      return false;
    }
  },
};

/** Conditions that are about the message rather than about its text. */
const FLAGS = {
  hasAttachment: (row) => Boolean(row.hasAttachments),
  noAttachment: (row) => !row.hasAttachments,
  isUnread: (row) => Boolean(row.unread),
  isBulk: (row) => Boolean(row.listId),
  isNotBulk: (row) => !row.listId,
};

export const RULE_FIELDS = Object.keys(FIELDS);
export const RULE_OPERATORS = Object.keys(OPERATORS);
export const RULE_FLAGS = Object.keys(FLAGS);
export const RULE_ACTIONS = ['move', 'copy', 'star', 'pin', 'markRead', 'markUnread', 'junk', 'trash', 'delete', 'stop'];

/** Does one condition hold? */
export function conditionHolds(row, condition) {
  if (condition.flag) return FLAGS[condition.flag]?.(row) ?? false;
  const read = FIELDS[condition.field];
  const compare = OPERATORS[condition.op];
  if (!read || !compare) return false;
  return compare(String(read(row) ?? ''), String(condition.value ?? ''));
}

/** Does a whole rule match? `all` means every condition; otherwise any. */
export function ruleMatches(row, rule) {
  const conditions = rule.conditions || [];
  if (!conditions.length) return false;
  return rule.all === false
    ? conditions.some((c) => conditionHolds(row, c))
    : conditions.every((c) => conditionHolds(row, c));
}

/**
 * Run the rules over a set of rows.
 *
 * Returns what should happen rather than doing it, so the same function decides
 * both what a rule will do on arrival and what it *would* do — which is what
 * makes "test this rule" possible without moving anybody's mail.
 *
 * @returns {{ row, rule, actions }[]} one entry per message that matched
 */
export function planRules(rows, rules) {
  const plan = [];
  const active = (rules || []).filter((r) => r.enabled !== false);

  for (const row of rows) {
    const actions = [];
    let rule = null;
    for (const candidate of active) {
      if (!ruleMatches(row, candidate)) continue;
      rule = rule || candidate;
      for (const action of candidate.actions || []) {
        if (action.type === 'stop') {
          if (actions.length) plan.push({ row, rule, actions });
          actions.length = 0;
          rule = null;
          break;
        }
        actions.push({ ...action, rule: candidate.name || candidate.id });
      }
      // `stopOnMatch` is the common case written as a checkbox rather than as
      // an action nobody remembers to add.
      if (candidate.stopOnMatch) break;
    }
    if (actions.length) plan.push({ row, rule, actions });
  }

  return plan;
}

/**
 * Carry out a plan against the store.
 *
 * @param {object} deps `store`, and `folderFor(accountId, role)` so "trash"
 *   resolves to whatever this account actually calls its Trash.
 * @returns {{ moved, starred, read, deleted, matched }}
 */
export function applyPlan(plan, { store, accountId, folderFor }) {
  const tally = { matched: plan.length, moved: 0, starred: 0, read: 0, deleted: 0, pinned: 0 };

  for (const { row, actions } of plan) {
    // A message can only be moved once, and moving it makes every later action
    // address a folder it is no longer in. So the flags go on first and the
    // move happens last.
    let destination = null;
    let remove = false;

    for (const action of actions) {
      switch (action.type) {
        case 'star':
          store.setFlags(accountId, row.folder, row.id, { flagged: true });
          tally.starred++;
          break;
        case 'pin':
          store.setFlags(accountId, row.folder, row.id, { pinned: true });
          tally.pinned++;
          break;
        case 'markRead':
          store.setFlags(accountId, row.folder, row.id, { unread: false });
          tally.read++;
          break;
        case 'markUnread':
          store.setFlags(accountId, row.folder, row.id, { unread: true });
          break;
        case 'move':
          destination = action.value;
          break;
        case 'copy': {
          const message = store.get(accountId, row.folder, row.id);
          if (message && action.value) {
            store.put(accountId, action.value, message, { force: true });
            store.upsertFolder(accountId, { path: action.value, name: String(action.value).split('/').pop() });
            tally.moved++;
          }
          break;
        }
        case 'junk':
          destination = folderFor(accountId, 'junk') || 'Junk';
          break;
        case 'trash':
          destination = folderFor(accountId, 'trash') || 'Trash';
          break;
        case 'delete':
          // The one action that means it. Everything else that sounds like
          // deleting is a move to Trash.
          remove = true;
          break;
        default:
          break;
      }
    }

    if (remove) {
      store.remove(accountId, row.folder, row.id);
      tally.deleted++;
      continue;
    }
    if (destination && destination !== row.folder) {
      const message = store.get(accountId, row.folder, row.id);
      if (message) {
        store.put(accountId, destination, message, { force: true });
        store.remove(accountId, row.folder, row.id);
        store.upsertFolder(accountId, { path: destination, name: String(destination).split('/').pop() });
        tally.moved++;
      }
    }
  }

  return tally;
}
