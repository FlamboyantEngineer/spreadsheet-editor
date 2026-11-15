// Follow ui-components' UIC-LABS-4.3 port policy while retaining this lab's API handlers
// and browser-only asset boundary. Its static server does not accept custom handlers.
const PORT_ATTEMPTS = 20;

/**
 * Bind the actual server, avoiding a separate availability probe and its race window.
 * @param {import('node:http').Server} server
 * @param {{ host: string, port: number, strictPort?: boolean }} options
 * @returns {Promise<{ port: number, url: string }>}
 */
export async function listenLabsServer(server, { host, port, strictPort = false }) {
  if (!Number.isInteger(port) || port < 0 || port > 65535)
    throw new RangeError('Invalid labs port.');
  const last = strictPort || port === 0 ? port : Math.min(65535, port + PORT_ATTEMPTS - 1);
  for (let candidate = port; ; candidate += 1) {
    try {
      await new Promise((resolve, reject) => {
        const cleanup = () => {
          server.off('error', failed);
          server.off('listening', listening);
        };
        const failed = (error) => {
          cleanup();
          reject(error);
        };
        const listening = () => {
          cleanup();
          resolve();
        };
        server.once('error', failed);
        server.once('listening', listening);
        try {
          server.listen(candidate, host);
        } catch (error) {
          failed(error);
        }
      });
      const address = server.address();
      const bound = address.port;
      const urlHost = host.includes(':') ? `[${host}]` : host;
      return { port: bound, url: `http://${urlHost}:${bound}/` };
    } catch (error) {
      if (error.code !== 'EADDRINUSE' || candidate >= last) throw error;
    }
  }
}
