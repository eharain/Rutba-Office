// A meeting held on a Talk Server: the worker that encrypts this person's
// sound and picture before they leave the window and decrypts the others' as
// they arrive, with the meeting's key (renderer/meeting/server-call.js). It is the worker of
// the engine such a meeting runs on (LiveKit's client), built as a bundle of
// its own (talk-worker.js) and started only when such a meeting is entered.
// It reaches nothing: frames come in from the window and go back to it.

import 'livekit-client/e2ee-worker';
