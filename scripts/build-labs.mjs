import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadEsbuild, reportLabsError } from './lib/labs-setup.mjs';

export const root = fileURLToPath(new URL('../', import.meta.url));
export const output = resolve(root, 'labs/build');
export async function buildLabs() {
  const { build } = await loadEsbuild();
  await mkdir(output, { recursive: true });
  await build({
    absWorkingDir: root,
    entryPoints: { labs: 'labs/src/main.tsx' },
    outdir: output,
    bundle: true,
    splitting: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2022',
    sourcemap: true,
    define: { 'process.env.NODE_ENV': '"development"' },
    loader: { '.png': 'dataurl', '.gif': 'dataurl' },
  });
  await build({
    absWorkingDir: root,
    entryPoints: ['src/calculation/http.ts'],
    outfile: resolve(output, 'calculation-server.mjs'),
    bundle: true,
    format: 'esm',
    platform: 'node',
    target: 'node22',
  });
  await writeFile(
    resolve(output, 'index.html'),
    '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Spreadsheet editor labs</title><link rel="stylesheet" href="/labs.css"></head><body><div id="root"></div><script type="module" src="/labs.js"></script></body></html>',
  );
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url))
  await buildLabs().catch(reportLabsError);
