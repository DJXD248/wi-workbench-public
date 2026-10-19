import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const handlerPath = process.argv[2];
if (!handlerPath) throw Error('Usage: node server.mjs <path/to/handler.mjs> [port]');
const port = Number(process.argv[3] ?? 8000);
if (!Number.isInteger(port) || port < 0 || port > 65535) throw Error('Invalid port');
const { handleDashboard } = await import(pathToFileURL(resolve(handlerPath)).href);

const server = createServer(async (incoming, outgoing) => {
  try {
    const url = new URL(incoming.url, 'http://127.0.0.1');
    if (url.pathname !== '/functions/v1/app') { outgoing.writeHead(404); outgoing.end(); return; }
    const response = await handleDashboard({
      request: new Request(url, { method: incoming.method }),
      supabase: null,
    });
    outgoing.writeHead(response.status, Object.fromEntries(response.headers));
    outgoing.end(Buffer.from(await response.arrayBuffer()));
  } catch {
    outgoing.writeHead(500, { 'content-type': 'application/json' });
    outgoing.end('{"error":"local_preview_failed"}');
  }
});
server.listen(port, '127.0.0.1', () => console.log(`Function preview: http://127.0.0.1:${server.address().port}/`));
