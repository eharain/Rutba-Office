// The Model Context Protocol over standard input and output, as an AI app on
// this computer speaks it when it starts `Rutba Office.exe --mcp` (owner,
// 2026-10-16, D-2026-10-16-06: "it should be off unless some one opens it for
// single session or want to keep it always open").
//
// One JSON-RPC 2.0 message a line, each way. The server answers initialize,
// ping, tools/list and tools/call, and ignores the notifications it is sent;
// nothing else is offered, so nothing else is asked of it. Written here rather
// than taken from a library: the protocol's part this needs is a few dozen
// lines, and every line of it can be read.

export const PROTOCOL_VERSIONS = ['2025-06-18', '2025-03-26', '2024-11-05'];

/**
 * The server's answers, apart from the streams, so the tests can talk to it:
 * `handle(message)` answers one parsed message with the reply to send, or
 * null for a notification. `tools` is { list(), call(name, args) } (tools.js).
 */
export function createServer({ name = 'rutba-office', version = '0.0.0', tools, instructions = '', onInitialize = null, onCall = null }) {
  const ok = (id, result) => ({ jsonrpc: '2.0', id, result });
  const fail = (id, code, message) => ({ jsonrpc: '2.0', id, error: { code, message } });

  async function handle(msg) {
    if (!msg || typeof msg !== 'object' || msg.jsonrpc !== '2.0') return fail(msg?.id ?? null, -32600, 'Not a JSON-RPC 2.0 message.');
    const { id, method, params } = msg;
    const notification = id === undefined || id === null;
    if (notification) return null;
    try {
      switch (method) {
        case 'initialize': {
          const asked = params?.protocolVersion;
          onInitialize?.({ client: params?.clientInfo || null, protocolVersion: asked || null });
          return ok(id, {
            protocolVersion: PROTOCOL_VERSIONS.includes(asked) ? asked : PROTOCOL_VERSIONS[0],
            capabilities: { tools: { listChanged: false } },
            serverInfo: { name, version },
            ...(instructions ? { instructions } : {}),
          });
        }
        case 'ping':
          return ok(id, {});
        case 'tools/list':
          return ok(id, { tools: await tools.list() });
        case 'tools/call': {
          onCall?.(params?.name);
          const result = await tools.call(params?.name, params?.arguments || {});
          return ok(id, result);
        }
        default:
          return fail(id, -32601, `Rutba Office does not answer ${method}.`);
      }
    } catch (err) {
      return fail(id, -32603, String(err?.message || err));
    }
  }

  return { handle };
}

/**
 * Serve `server` on two streams: a line in, a line out, in the order asked.
 * Answers when `input` ends. Nothing but protocol goes to `output`; whatever
 * else the process prints belongs on standard error (run.js sees to that).
 */
export function serve(server, input, output) {
  return new Promise((resolve) => {
    let buffered = '';
    let chain = Promise.resolve();
    const reply = (obj) => { if (obj) output.write(`${JSON.stringify(obj)}\n`); };
    input.setEncoding?.('utf8');
    input.on('data', (chunk) => {
      buffered += chunk;
      let at = buffered.indexOf('\n');
      while (at >= 0) {
        const line = buffered.slice(0, at).trim();
        buffered = buffered.slice(at + 1);
        if (line) {
          let msg = null;
          try {
            msg = JSON.parse(line);
          } catch {
            chain = chain.then(() => reply({ jsonrpc: '2.0', id: null, error: { code: -32700, message: 'That line is not JSON.' } }));
          }
          if (msg) chain = chain.then(() => server.handle(msg)).then(reply);
        }
        at = buffered.indexOf('\n');
      }
    });
    input.on('end', () => chain.then(() => resolve()));
  });
}
