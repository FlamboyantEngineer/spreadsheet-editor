import { EventEmitter } from 'node:events';
import { createServer, type Server } from 'node:http';
import { afterEach, describe, expect, it, vi } from 'vitest';

interface ListenerOptions {
  host: string;
  port: number;
  strictPort?: boolean;
}
const { listenLabsServer } = (await import(
  new URL('../../scripts/lib/labs-server.mjs', import.meta.url).href
)) as {
  listenLabsServer(
    server: Server,
    options: ListenerOptions,
  ): Promise<{ port: number; url: string }>;
};
const host = '127.0.0.1';
const servers: Server[] = [];
function server() {
  const instance = createServer((request, response) => {
    response.writeHead(200, { 'content-type': 'application/json' });
    response.end(JSON.stringify({ route: request.url }));
  });
  servers.push(instance);
  return instance;
}
afterEach(async () => {
  await Promise.all(
    servers.splice(0).map(
      (instance) =>
        new Promise<void>((resolve, reject) => {
          if (!instance.listening) return resolve();
          instance.closeAllConnections();
          instance.close((error) => (error ? reject(error) : resolve()));
        }),
    ),
  );
});

// Deterministic failures exercise exhaustion and OS errors without reserving fixed ports.
class RejectedServer extends EventEmitter {
  attempts: number[] = [];
  constructor(private code: string) {
    super();
  }
  listen(port: number) {
    this.attempts.push(port);
    queueMicrotask(() =>
      this.emit('error', Object.assign(new Error(this.code), { code: this.code })),
    );
    return this;
  }
  asServer() {
    return this as unknown as Server;
  }
}

describe('SE-LABS-1: shared lab port policy with custom API handlers', () => {
  it('binds an assigned port, reports it and retains request and existing event handlers', async () => {
    const instance = server();
    const listening = vi.fn();
    instance.on('listening', listening);
    const listeners = instance.listeners('listening');
    const bound = await listenLabsServer(instance, { host, port: 0 });
    expect(bound.port).toBeGreaterThan(0);
    expect(bound.url).toBe(`http://${host}:${bound.port}/`);
    expect(await (await fetch(`${bound.url}api/calculate`)).json()).toEqual({
      route: '/api/calculate',
    });
    expect(listening).toHaveBeenCalledTimes(1);
    expect(instance.listeners('listening')).toEqual(listeners);
    expect(instance.listenerCount('error')).toBe(0);
  });

  it('skips occupied ports while leaving their owners running', async () => {
    const first = server();
    const start = await listenLabsServer(first, { host, port: 0 });
    const second = server();
    const occupied = await listenLabsServer(second, { host, port: start.port });
    const instance = server();
    const listeners = instance.listeners('listening');
    const selected = await listenLabsServer(instance, { host, port: start.port });
    expect(occupied.port).toBeGreaterThan(start.port);
    expect(selected.port).toBeGreaterThan(occupied.port);
    expect(selected.port).toBeLessThan(start.port + 20);
    expect(first.listening).toBe(true);
    expect(second.listening).toBe(true);
    expect(await (await fetch(`${selected.url}api/xlsx/export`)).json()).toEqual({
      route: '/api/xlsx/export',
    });
    expect(instance.listeners('listening')).toEqual(listeners);
    expect(instance.listenerCount('error')).toBe(0);
  });

  it('fails on an occupied strict port and leaves the existing listener available', async () => {
    const blocker = server();
    const occupied = await listenLabsServer(blocker, { host, port: 0 });
    const instance = server();
    const listeners = instance.listeners('listening');
    await expect(
      listenLabsServer(instance, { host, port: occupied.port, strictPort: true }),
    ).rejects.toMatchObject({ code: 'EADDRINUSE' });
    expect(instance.listening).toBe(false);
    expect(instance.listeners('listening')).toEqual(listeners);
    expect(instance.listenerCount('error')).toBe(0);
    expect((await fetch(occupied.url)).status).toBe(200);
  });

  it.each([
    { port: 4173, strictPort: false, code: 'EADDRINUSE', attempts: 20 },
    { port: 4173, strictPort: true, code: 'EADDRINUSE', attempts: 1 },
    { port: 0, strictPort: false, code: 'EADDRINUSE', attempts: 1 },
    { port: 65534, strictPort: false, code: 'EADDRINUSE', attempts: 2 },
    { port: 4173, strictPort: false, code: 'EACCES', attempts: 1 },
    { port: 4173, strictPort: false, code: 'ENOTFOUND', attempts: 1 },
  ])('stops with the original error for %j and releases temporary listeners', async (options) => {
    const instance = new RejectedServer(options.code);
    await expect(listenLabsServer(instance.asServer(), { host, ...options })).rejects.toMatchObject(
      {
        code: options.code,
      },
    );
    expect(instance.attempts).toEqual(
      Array.from({ length: options.attempts }, (_, index) => options.port + index),
    );
    expect(instance.listenerCount('listening')).toBe(0);
    expect(instance.listenerCount('error')).toBe(0);
  });

  it.each([-1, 65536, 1.5, NaN])('rejects invalid port %s before listening', async (port) => {
    const instance = server();
    await expect(listenLabsServer(instance, { host, port })).rejects.toThrow('Invalid labs port.');
    expect(instance.listening).toBe(false);
  });

  it('cleans up when listen throws synchronously', async () => {
    const instance = server();
    const listeners = instance.listeners('listening');
    await listenLabsServer(instance, { host, port: 0 });
    await expect(listenLabsServer(instance, { host, port: 0 })).rejects.toMatchObject({
      code: 'ERR_SERVER_ALREADY_LISTEN',
    });
    expect(instance.listeners('listening')).toEqual(listeners);
    expect(instance.listenerCount('error')).toBe(0);
  });
});
