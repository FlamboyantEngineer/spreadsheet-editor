import { createServer } from 'node:http';
import { createReadStream } from 'node:fs';
import { realpath, stat } from 'node:fs/promises';
import { resolve, extname, sep } from 'node:path';
import { pathToFileURL } from 'node:url';
import { Readable } from 'node:stream';
import { buildLabs, output } from './build-labs.mjs';
import { reportLabsError } from './lib/labs-setup.mjs';

function option(name, fallback) {
  const index = process.argv.indexOf(`--${name}`);
  return index < 0 ? fallback : process.argv[index + 1];
}
async function main() {
  const port = Number(option('port', process.env.LABS_PORT ?? '4173'));
  const host = option('host', process.env.LABS_HOST ?? '127.0.0.1');
  if (!Number.isInteger(port) || port < 0 || port > 65535)
    throw new RangeError('Invalid labs port.');
  await buildLabs();
  const { handleCalculationRequest } = await import(
    pathToFileURL(resolve(output, 'calculation-server.mjs')).href
  );
  const types = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.map': 'application/json',
    '.svg': 'image/svg+xml',
  };
  const server = createServer(async (request, response) => {
    try {
      const url = new URL(request.url ?? '/', 'http://localhost');
      if (url.pathname === '/api/calculate') {
        const webRequest = new Request(url, {
          method: request.method,
          headers: request.headers,
          ...(request.method === 'GET' || request.method === 'HEAD'
            ? {}
            : { body: Readable.toWeb(request), duplex: 'half' }),
        });
        const result = await handleCalculationRequest(webRequest);
        response.writeHead(result.status, Object.fromEntries(result.headers));
        response.end(await result.text());
        return;
      }
      if (request.method !== 'GET' && request.method !== 'HEAD') {
        response.writeHead(405).end();
        return;
      }
      const pathname = decodeURIComponent(url.pathname);
      // Serve browser assets only; the server evaluator bundle is private to this process.
      const file = await realpath(
        resolve(output, `.${pathname === '/' ? '/index.html' : pathname}`),
      );
      if (!file.startsWith(output + sep) || !types[extname(file)]) {
        response.writeHead(404).end('Not found');
        return;
      }
      const details = await stat(file);
      if (!details.isFile()) {
        response.writeHead(404).end('Not found');
        return;
      }
      response.writeHead(200, {
        'content-type': types[extname(file)],
        'content-length': details.size,
        'cache-control': 'no-store',
      });
      if (request.method === 'HEAD') response.end();
      else createReadStream(file).pipe(response);
    } catch (error) {
      if (!response.headersSent)
        response.writeHead(error instanceof URIError ? 400 : 404).end('Not found');
      else response.end();
    }
  });
  server.on('error', (error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
  server.listen(port, host, () => {
    console.log(`Spreadsheet labs: http://${host}:${server.address().port}/`);
  });
  process.on('SIGTERM', () => server.close());
  process.on('SIGINT', () => server.close());
}

await main().catch(reportLabsError);
