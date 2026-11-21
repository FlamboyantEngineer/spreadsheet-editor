import { execFileSync } from 'node:child_process';
import { copyFile, rm } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
await rm(new URL('../dist/', import.meta.url), { recursive: true, force: true });
execFileSync(
  process.execPath,
  [
    fileURLToPath(new URL('../node_modules/typescript/bin/tsc', import.meta.url)),
    '-p',
    'tsconfig.build.json',
  ],
  { cwd: root, stdio: 'inherit' },
);

// The two opt-in stylesheets and the parts they import, side by side as in labs/src.
for (const sheet of ['reference.css', 'shadcn.css', 'editor.css', 'shadcn-retouch.css']) {
  await copyFile(
    new URL(`../labs/src/${sheet}`, import.meta.url),
    new URL(`../dist/${sheet}`, import.meta.url),
  );
}
