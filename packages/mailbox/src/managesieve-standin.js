// A stand-in ManageSieve server for the checks: plain TCP on this computer,
// the commands Office uses, scripts kept in memory, one active at a time.

import net from 'node:net';

export function standInSieve({ user = 'amira@checks.example', pass = 'check', vacation = true, scripts: given = {} } = {}) {
  const scripts = new Map(Object.entries(given));
  let active = null;
  const log = [];
  const caps = ['"IMPLEMENTATION" "Stand-in"', `"SIEVE" "fileinto ${vacation ? 'vacation ' : ''}date relational include"`, '"SASL" "PLAIN"', '"VERSION" "1.0"', 'OK "ready"'];
  const q = (s) => `"${String(s).replace(/[\\"]/g, (c) => `\\${c}`)}"`;
  const server = net.createServer((sock) => {
    let buf = Buffer.alloc(0);
    let authed = false;
    let pending = null; // { name, need } for a PUTSCRIPT literal
    sock.write(`${caps.join('\r\n')}\r\n`);
    sock.on('data', (chunk) => {
      buf = Buffer.concat([buf, chunk]);
      for (;;) {
        if (pending) {
          if (buf.length < pending.need + 2) return;
          const body = buf.subarray(0, pending.need).toString('utf8');
          buf = buf.subarray(pending.need + 2);
          scripts.set(pending.name, body);
          log.push(['PUTSCRIPT', pending.name]);
          pending = null;
          sock.write('OK\r\n');
          continue;
        }
        const i = buf.indexOf('\r\n');
        if (i < 0) return;
        const l = buf.subarray(0, i).toString('utf8');
        buf = buf.subarray(i + 2);
        const args = [...l.matchAll(/"((?:[^"\\]|\\.)*)"/g)].map((m) => m[1].replace(/\\(.)/g, '$1'));
        const cmd = l.split(' ')[0].toUpperCase();
        log.push([cmd, ...args]);
        if (cmd === 'AUTHENTICATE') {
          const [, u, p] = Buffer.from(args[1] || '', 'base64').toString('utf8').split('\0');
          authed = u === user && p === pass;
          sock.write(authed ? 'OK\r\n' : 'NO "Authentication failed."\r\n');
        } else if (!authed && cmd !== 'LOGOUT') sock.write('NO "Sign in first."\r\n');
        else if (cmd === 'LISTSCRIPTS') sock.write(`${[...scripts.keys()].map((n) => `${q(n)}${n === active ? ' ACTIVE' : ''}\r\n`).join('')}OK\r\n`);
        else if (cmd === 'GETSCRIPT') {
          const body = scripts.get(args[0]);
          if (body == null) sock.write('NO (NONEXISTENT) "No such script."\r\n');
          else { const b = Buffer.from(body, 'utf8'); sock.write(`{${b.length}}\r\n`); sock.write(b); sock.write('\r\nOK\r\n'); }
        } else if (cmd === 'PUTSCRIPT') {
          const m = /\{(\d+)\+?\}$/.exec(l);
          pending = { name: args[0], need: Number(m[1]) };
        } else if (cmd === 'SETACTIVE') {
          if (args[0] && !scripts.has(args[0])) sock.write('NO (NONEXISTENT) "No such script."\r\n');
          else { active = args[0] || null; sock.write('OK\r\n'); }
        } else if (cmd === 'DELETESCRIPT') { scripts.delete(args[0]); sock.write('OK\r\n'); }
        else if (cmd === 'LOGOUT') { sock.write('OK\r\n'); sock.end(); }
        else sock.write('NO "Unknown command."\r\n');
      }
    });
    sock.on('error', () => {});
  });
  return {
    listen: () => new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(server.address().port))),
    close: () => new Promise((resolve) => server.close(() => resolve())),
    scripts,
    active: () => active,
    setActive: (n) => { active = n; },
    log,
  };
}
