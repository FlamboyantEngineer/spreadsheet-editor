import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { ensureLabsDependencies } from './lib/labs-dependencies.mjs';
import { loadEsbuild, reportLabsError } from './lib/labs-setup.mjs';

export const root = fileURLToPath(new URL('../', import.meta.url));
export const output = resolve(root, 'labs/build');
export async function buildLabs() {
  await ensureLabsDependencies(root);
  const { build } = await loadEsbuild();
  await mkdir(output, { recursive: true });
  await build({
    absWorkingDir: root,
    entryPoints: { labs: 'labs/src/main.tsx', reference: 'labs/src/reference.css' },
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
    entryPoints: ['src/calculation/index.ts'],
    outfile: resolve(output, 'calculation-server.mjs'),
    bundle: true,
    format: 'esm',
    platform: 'node',
    target: 'node22',
  });
  await build({
    absWorkingDir: root,
    entryPoints: ['src/excel/index.ts'],
    outfile: resolve(output, 'excel-server.mjs'),
    bundle: true,
    format: 'esm',
    platform: 'node',
    target: 'node22',
    banner: {
      js: "import { createRequire as _createRequire } from 'node:module'; const require=_createRequire(import.meta.url);",
    },
  });
  await build({
    absWorkingDir: root,
    entryPoints: ['src/excel/worker.ts'],
    outfile: resolve(output, 'worker.js'),
    bundle: true,
    format: 'esm',
    platform: 'node',
    target: 'node22',
    banner: {
      js: "import { createRequire as _createRequire } from 'node:module'; const require=_createRequire(import.meta.url);",
    },
  });
  for (const [page, title] of Object.entries({
    index: 'Spreadsheet editor',
    minimal: 'Minimal',
    embedded: 'Embedded in an app',
    full: 'Full editor',
  })) {
    await writeFile(
      resolve(output, `${page}.html`),
      `<!doctype html><html lang="en" data-reference-document><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${title} · Labs</title><link rel="icon" href="data:,"><link id="lab-reference" rel="stylesheet" href="/reference.css"><link rel="stylesheet" href="/labs.css"></head><body><div id="root"></div><script type="module" src="/labs.js"></script></body></html>`,
    );
  }
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url))
  await buildLabs().catch(reportLabsError);
