import { createReadStream } from 'node:fs';
import { realpath, stat } from 'node:fs/promises';
import { createServer } from 'node:http';
import { extname, resolve, sep } from 'node:path';
import { Readable } from 'node:stream';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { buildLabs, output } from './build-labs.mjs';
import { listenLabsServer } from './lib/labs-server.mjs';
import { reportLabsError } from './lib/labs-setup.mjs';

async function main() {
  const { values } = parseArgs({
    options: {
      port: { type: 'string', default: process.env.LABS_PORT ?? '4173' },
      host: { type: 'string', default: process.env.LABS_HOST ?? '127.0.0.1' },
      'strict-port': { type: 'boolean', default: false },
    },
  });
  const port = Number(values.port);
  const host = values.host;
  if (!Number.isInteger(port) || port < 0 || port > 65535)
    throw new RangeError('Invalid labs port.');
  await buildLabs();
  const { handleCalculationRequest } = await import(
    pathToFileURL(resolve(output, 'calculation-server.mjs')).href
  );
  const { importXlsx, exportXlsx } = await import(
    pathToFileURL(resolve(output, 'excel-server.mjs')).href
  );
  const { calculateWorkbook } = await import(
    pathToFileURL(resolve(output, 'calculation-server.mjs')).href
  );
  async function readBody(request, maxBytes) {
    const chunks = [];
    let size = 0;
    for await (const chunk of request) {
      size += chunk.length;
      if (size > maxBytes) throw new RangeError('Request is too large.');
      chunks.push(chunk);
    }
    return Buffer.concat(chunks);
  }
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
      if (url.pathname === '/api/xlsx/import' || url.pathname === '/api/xlsx/export') {
        if (request.method !== 'POST') {
          response.writeHead(405).end();
          return;
        }
        try {
          const body = await readBody(request, 20 * 1024 * 1024);
          if (url.pathname.endsWith('/import')) {
            const imported = await importXlsx(body);
            const result = calculateWorkbook(imported.workbook);
            response
              .writeHead(200, { 'content-type': 'application/json' })
              .end(JSON.stringify({ ...imported, result }));
          } else {
            const { workbook, result } = JSON.parse(body.toString());
            const bytes = await exportXlsx(workbook, result);
            response
              .writeHead(200, {
                'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'content-disposition': 'attachment; filename="workbook.xlsx"',
              })
              .end(bytes);
          }
        } catch (error) {
          response
            .writeHead(error instanceof RangeError ? 413 : 400, {
              'content-type': 'application/json',
            })
            .end(JSON.stringify({ error: error.message }));
        }
        return;
      }
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
      if (
        !file.startsWith(output + sep) ||
        !types[extname(file)] ||
        file.startsWith(resolve(output, 'worker.'))
      ) {
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
  const listening = await listenLabsServer(server, {
    port,
    host,
    strictPort: values['strict-port'],
  });
  server.on('error', (error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
  process.on('SIGTERM', () => server.close());
  process.on('SIGINT', () => server.close());
  const moved = port !== 0 && listening.port !== port;
  console.log(`Spreadsheet labs: ${listening.url}${moved ? ` (port ${port} was in use)` : ''}`);
}

await main().catch(reportLabsError);
