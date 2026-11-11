// A node --test reporter for the gate: one JSON line a test, with its file,
// line and depth, beside the TAP the log keeps. TAP names the file of a test
// that failed, never of one that passed; the failure list needs both, so a
// pass later closes the entry its failure opened (tools/gate-hosts.js).
//
//   node --test --test-reporter=tap --test-reporter-destination=stdout \
//     --test-reporter=./tools/gate-reporter.js --test-reporter-destination=<file>

export default async function* gateReporter(source) {
  for await (const event of source) {
    if (event.type !== 'test:start' && event.type !== 'test:pass' && event.type !== 'test:fail') continue;
    const d = event.data || {};
    const error = d.details?.error;
    yield `${JSON.stringify({
      event: event.type.slice(5),
      name: d.name,
      nesting: d.nesting,
      file: d.file || null,
      line: d.line ?? null,
      skip: Boolean(d.skip),
      todo: Boolean(d.todo),
      failureType: error?.failureType || null,
      error: error ? String(error.cause?.message || error.message || '').split('\n')[0].slice(0, 300) : null,
    })}\n`;
  }
}
