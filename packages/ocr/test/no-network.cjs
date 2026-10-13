// Loaded with --require before anything else, in the process and in every
// thread it starts: each way out to the network counts an attempt and
// fails. The recognition test reads the count afterwards.

const attempts = [];
globalThis.__networkAttempts = attempts;
const refuse = (what) => function refused() {
  attempts.push(what);
  throw new Error(`network refused: ${what}`);
};

const net = require('node:net');
net.Socket.prototype.connect = refuse('net.connect');
net.connect = refuse('net.connect');
net.createConnection = refuse('net.createConnection');
for (const name of ['http', 'https']) {
  const mod = require(`node:${name}`);
  mod.request = refuse(`${name}.request`);
  mod.get = refuse(`${name}.get`);
}
require('node:tls').connect = refuse('tls.connect');
const dns = require('node:dns');
dns.lookup = refuse('dns.lookup');
dns.resolve = refuse('dns.resolve');
dns.promises.lookup = refuse('dns.promises.lookup');
require('node:dgram').createSocket = refuse('dgram.createSocket');
globalThis.fetch = refuse('fetch');
globalThis.WebSocket = refuse('WebSocket');
globalThis.XMLHttpRequest = refuse('XMLHttpRequest');
globalThis.__networkGuard = true;
