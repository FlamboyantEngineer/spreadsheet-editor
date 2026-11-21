import { execFile, spawn } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { promisify } from 'node:util';
import { assertLabsNodeVersion, LabsSetupError } from './labs-setup.mjs';

const execute = promisify(execFile);
const requiredEntries = [
  'react/jsx-runtime',
  'react-dom/client',
  '@jspreadsheet-ce/react',
  'jspreadsheet-ce/dist/jspreadsheet.css',
  'jsuites/dist/jsuites.css',
  '@flamboyantengineer/ui-components/reference.css',
  '@flamboyantengineer/ui-components/shadcn.css',
  '@flamboyantengineer/ui-components/labs',
  'decimal.js',
  'exceljs',
  'fflate',
];
const packageName = (specifier) =>
  specifier
    .split('/')
    .slice(0, specifier.startsWith('@') ? 2 : 1)
    .join('/');
async function json(file) {
  return JSON.parse(await readFile(file, 'utf8'));
}

async function installationProblem(root, manifest, lock) {
  const direct = { ...manifest.dependencies, ...manifest.devDependencies };
  for (const section of ['dependencies', 'devDependencies']) {
    const entries = (value) =>
      JSON.stringify(Object.entries(value ?? {}).sort(([a], [b]) => a.localeCompare(b)));
    if (entries(manifest[section]) !== entries(lock.packages?.['']?.[section]))
      return 'package.json and package-lock.json differ';
  }
  let installed;
  try {
    installed = await json(join(root, 'node_modules/.package-lock.json'));
  } catch {
    return 'dependencies are missing or incomplete';
  }
  const required = Object.entries(lock.packages ?? {}).filter(
    ([location, expected]) =>
      location.startsWith('node_modules/') &&
      !expected.link &&
      (!expected.optional || Object.hasOwn(direct, location.slice(13))),
  );
  const problems = await Promise.all(
    required.map(async ([location, expected]) => {
      const recorded = installed.packages?.[location];
      if (
        !recorded ||
        ['version', 'resolved', 'integrity'].some((key) => expected[key] !== recorded[key])
      )
        return `${location.slice(13)} is missing or out of date`;
      try {
        const actual = await json(join(root, location, 'package.json'));
        return actual.version === expected.version
          ? null
          : `${location.slice(13)} has the wrong version`;
      } catch {
        return `${location.slice(13)} is incomplete`;
      }
    }),
  );
  const problem = problems.find(Boolean);
  if (problem) return problem;
  // A separate process avoids retaining a broken/wrong-platform bundler in Node's import cache
  // before npm repairs it. Resolve prepared assets without importing browser-only modules.
  const entries = requiredEntries.filter((specifier) =>
    Object.hasOwn(direct, packageName(specifier)),
  );
  const probe = `
    import { access } from 'node:fs/promises';
    import { fileURLToPath } from 'node:url';
    for (const entry of ${JSON.stringify(entries)}) await access(fileURLToPath(import.meta.resolve(entry)));
    const esbuild = await import('esbuild');
    try { await esbuild.transform('const ready = 1;', { loader: 'js' }); }
    finally { esbuild.stop?.(); }
  `;
  try {
    await execute(process.execPath, ['--input-type=module', '--eval', probe], {
      cwd: root,
      timeout: 15000,
    });
    return null;
  } catch {
    return 'the bundler or prepared UI assets need repair';
  }
}

function install(root) {
  const npm = process.env.npm_execpath;
  if (!npm)
    throw new LabsSetupError(
      'Run npm run labs from the repository root so the launcher can install its dependencies.',
    );
  return new Promise((resolve, reject) => {
    // npm_execpath is npm's JavaScript CLI, including on Windows. No shell/npm.cmd quoting.
    const child = spawn(
      process.execPath,
      [
        npm,
        'ci',
        '--include=dev',
        '--include=optional',
        '--ignore-scripts=false',
        '--no-audit',
        '--no-fund',
      ],
      {
        cwd: root,
        stdio: 'inherit',
        env: {
          ...process.env,
          NODE_ENV: 'development',
          npm_config_omit: '',
          npm_config_ignore_scripts: 'false',
        },
      },
    );
    child.once('error', reject);
    child.once('exit', (code, signal) => {
      if (code === 0) resolve();
      else
        reject(
          Object.assign(
            new LabsSetupError(
              `Labs dependency installation ${signal ? `stopped (${signal})` : `failed (exit ${code})`}. Fix the npm error above, then rerun npm run labs.\n` +
                'The ui-components dependency is public: https://github.com/FlamboyantEngineer/ui-components. No private-library token is required.',
            ),
            { exitCode: code || 1 },
          ),
        );
    });
  });
}

/** Install the locked toolchain only when absent, stale, incomplete or unusable on this platform. */
export async function ensureLabsDependencies(root) {
  assertLabsNodeVersion();
  const manifest = await json(join(root, 'package.json'));
  const lock = await json(join(root, 'package-lock.json'));
  const problem = await installationProblem(root, manifest, lock);
  if (!problem) return;
  console.log(`Preparing labs: ${problem}. Installing locked dependencies…`);
  await install(root);
  const remaining = await installationProblem(root, manifest, lock);
  if (remaining)
    throw new LabsSetupError(
      `Labs dependencies are still unavailable after installation: ${remaining}. See the npm output above.`,
    );
}
