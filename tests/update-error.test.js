// A failed update check is told as a sentence by its kind, not as the
// updater's raw HTTP exchange and stack (the About box, 2026-10-12).
import test from 'node:test';
import assert from 'node:assert/strict';
import { updateErrorKind } from '../apps/desktop/main/update-error.js';

const published404 = 'Cannot find latest.yml in the latest release artifacts (https://github.com/eharain/Rutba-Office/releases/download/v1.44.0/latest.yml): HttpError: 404\n"method: GET url: https://github.com/eharain/Rutba-Office/releases/download/v1.44.0/latest.yml"\n Headers: { "cache-control": "no-cache" }\n    at createHttpError (httpExecutor.js:53:12)';

test('a release whose files are not up yet reads as still being published, its detail one short line', () => {
  const r = updateErrorKind(published404);
  assert.equal(r.kind, 'publishing');
  assert.ok(!r.detail.includes('\n') && r.detail.length <= 160 && !/Headers|at create/.test(r.detail), r.detail);
});

test('no connection, a refusal and anything else each have their kind', () => {
  assert.equal(updateErrorKind('getaddrinfo ENOTFOUND github.com').kind, 'offline');
  assert.equal(updateErrorKind('net::ERR_INTERNET_DISCONNECTED').kind, 'offline');
  assert.equal(updateErrorKind('HttpError: 403 rate limit exceeded').kind, 'refused');
  assert.equal(updateErrorKind('sha512 checksum mismatch').kind, 'other');
  assert.equal(updateErrorKind(undefined).kind, 'other');
});
