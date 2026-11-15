import { spawnSync } from 'node:child_process';
import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';

const scripts = fileURLToPath(new URL('../../scripts/', import.meta.url));
const fixtures: string[] = [];
afterEach(() => {
  for (const directory of fixtures.splice(0)) rmSync(directory, { recursive: true, force: true });
});
const healthyEsbuild =
  'export async function transform() { return { code: "const ready = 1;" }; } export function build() { throw new Error("Fixture build failure"); }';
function writeJson(path: string, value: unknown) {
  writeFileSync(path, JSON.stringify(value));
}
function fakePackage(root: string, location: string, version: string, source = '') {
  const directory = join(root, location);
  mkdirSync(directory, { recursive: true });
  writeJson(join(directory, 'package.json'), {
    name: location.slice(13),
    version,
    type: 'module',
    exports: './index.js',
  });
  writeFileSync(join(directory, 'index.js'), source);
}
function lockFor(directory: string) {
  return JSON.parse(readFileSync(join(directory, 'package-lock.json'), 'utf8'));
}
function healthyTree(directory: string) {
  const root = join(directory, 'healthy');
  const lock = lockFor(directory);
  for (const [location, expected] of Object.entries(lock.packages) as [
    string,
    { version: string },
  ][]) {
    if (!location) continue;
    fakePackage(
      root,
      location,
      expected.version,
      location === 'node_modules/esbuild' ? healthyEsbuild : '',
    );
  }
  const packages = { ...lock.packages };
  delete packages[''];
  writeJson(join(root, 'node_modules/.package-lock.json'), { lockfileVersion: 3, packages });
}
function checkout(installed = false) {
  const directory = mkdtempSync(join(tmpdir(), 'spreadsheet labs '));
  fixtures.push(directory);
  cpSync(scripts, join(directory, 'scripts'), { recursive: true });
  const manifest = {
    type: 'module',
    dependencies: { 'fixture-runtime': '1.0.0' },
    devDependencies: { esbuild: '^1.0.0' },
  };
  writeJson(join(directory, 'package.json'), manifest);
  writeJson(join(directory, 'package-lock.json'), {
    lockfileVersion: 3,
    packages: {
      '': manifest,
      'node_modules/esbuild': { version: '1.0.0', resolved: 'fixture:esbuild', dev: true },
      'node_modules/fixture-runtime': { version: '1.0.0' },
      'node_modules/fixture-transitive': { version: '2.0.0' },
    },
  });
  healthyTree(directory);
  if (installed)
    cpSync(join(directory, 'healthy/node_modules'), join(directory, 'node_modules'), {
      recursive: true,
    });
  mkdirSync(join(directory, 'fixture tools'));
  writeFileSync(
    join(directory, 'fixture tools/npm cli.js'),
    `
    import { existsSync, readFileSync, writeFileSync, cpSync, rmSync } from 'node:fs';
    const file = 'install-calls.json';
    const calls = existsSync(file) ? JSON.parse(readFileSync(file, 'utf8')) : [];
    calls.push({ args: process.argv.slice(2), cwd: process.cwd(), nodeEnv: process.env.NODE_ENV, ignoreScripts: process.env.npm_config_ignore_scripts });
    writeFileSync(file, JSON.stringify(calls));
    if (process.env.FIXTURE_INSTALL_FAIL) { console.error('Fixture npm failure'); process.exit(7); }
    if (!process.env.FIXTURE_EMPTY_INSTALL) {
      rmSync('node_modules', { recursive: true, force: true });
      cpSync('healthy/node_modules', 'node_modules', { recursive: true });
    }
  `,
  );
  return directory;
}
function run(directory: string, script: string, env: NodeJS.ProcessEnv = {}) {
  const overrides = { npm_execpath: join(directory, 'fixture tools/npm cli.js'), ...env };
  // Windows environment names are case-insensitive. Vitest workers can inherit uppercase
  // npm variables, which must not take precedence over this fixture's lowercase overrides.
  const inherited = Object.fromEntries(
    Object.entries(process.env).filter(
      ([key]) => !Object.keys(overrides).some((name) => name.toLowerCase() === key.toLowerCase()),
    ),
  );
  return spawnSync(process.execPath, [join(directory, 'scripts', script)], {
    cwd: directory,
    env: { ...inherited, ...overrides },
    encoding: 'utf8',
    timeout: 15_000,
  });
}
function ensure(directory: string, env: NodeJS.ProcessEnv = {}) {
  writeFileSync(
    join(directory, 'scripts/ensure-fixture.mjs'),
    'import { ensureLabsDependencies } from "./lib/labs-dependencies.mjs"; import { reportLabsError } from "./lib/labs-setup.mjs"; await ensureLabsDependencies(process.cwd()).catch(reportLabsError);',
  );
  return run(directory, 'ensure-fixture.mjs', env);
}
function calls(directory: string) {
  return JSON.parse(readFileSync(join(directory, 'install-calls.json'), 'utf8'));
}

describe('SE-LABS-2.1: self-bootstrapping portable startup', () => {
  it.each(['labs.mjs', 'build-labs.mjs'])(
    'installs before building %s from a fresh checkout, including under production/omit settings',
    (script) => {
      const directory = checkout();
      const result = run(directory, script, {
        NODE_ENV: 'production',
        npm_config_omit: 'dev optional',
        npm_config_ignore_scripts: 'true',
        PATH: '',
      });
      expect(result.error).toBeUndefined();
      expect(result.stderr).toContain('Fixture build failure'); // Build reached only after successful setup.
      expect(result.stderr).not.toContain('ERR_MODULE_NOT_FOUND');
      expect(calls(directory)).toEqual([
        {
          args: [
            'ci',
            '--include=dev',
            '--include=optional',
            '--ignore-scripts=false',
            '--no-audit',
            '--no-fund',
          ],
          cwd: directory,
          nodeEnv: 'development',
          ignoreScripts: 'false',
        },
      ]);
    },
  );
  it('reuses an already complete locked installation on subsequent starts', () => {
    const directory = checkout(true);
    expect(ensure(directory).status).toBe(0);
    expect(ensure(directory).status).toBe(0);
    expect(existsSync(join(directory, 'install-calls.json'))).toBe(false);
  });
  it.each(['esbuild', 'fixture-runtime', 'fixture-transitive'])(
    'repairs missing %s even when the install record remains',
    (name) => {
      const directory = checkout(true);
      rmSync(join(directory, 'node_modules', name), { recursive: true });
      expect(ensure(directory).status).toBe(0);
      expect(calls(directory)).toHaveLength(1);
      expect(existsSync(join(directory, 'node_modules', name, 'package.json'))).toBe(true);
      expect(ensure(directory).status).toBe(0);
      expect(calls(directory)).toHaveLength(1);
    },
  );
  it('reinstalls when the lockfile changes, including a Git revision with the same package version', () => {
    const directory = checkout(true),
      lock = lockFor(directory);
    lock.packages['node_modules/esbuild'].resolved = 'fixture:another-commit';
    writeJson(join(directory, 'package-lock.json'), lock);
    healthyTree(directory);
    expect(ensure(directory).status).toBe(0);
    expect(calls(directory)).toHaveLength(1);
  });
  it('repairs an unusable native bundler before loading it in the launcher process', () => {
    const directory = checkout(true);
    writeFileSync(
      join(directory, 'node_modules/esbuild/index.js'),
      'export function transform() { throw new Error("Wrong-platform binary"); }',
    );
    const result = run(directory, 'labs.mjs');
    expect(calls(directory)).toHaveLength(1);
    expect(result.stderr).toContain('Fixture build failure');
    expect(result.stderr).not.toContain('Wrong-platform binary');
  });
  it.each(['reference.css', 'labs.js'])(
    'repairs missing prepared UI asset %s even if package versions match',
    (asset) => {
      const directory = checkout(true),
        lock = lockFor(directory);
      const name = '@flamboyantengineer/ui-components';
      lock.packages[''].dependencies[name] = 'fixture:ui';
      lock.packages[`node_modules/${name}`] = { version: '0.1.0', resolved: 'fixture:ui' };
      writeJson(join(directory, 'package.json'), lock.packages['']);
      writeJson(join(directory, 'package-lock.json'), lock);
      healthyTree(directory);
      const ui = join(directory, 'healthy/node_modules', name);
      writeJson(join(ui, 'package.json'), {
        name,
        version: '0.1.0',
        exports: { './reference.css': './reference.css', './labs': './labs.js' },
      });
      writeFileSync(join(ui, 'reference.css'), '/* prepared */');
      writeFileSync(join(ui, 'labs.js'), 'export {};');
      cpSync(join(directory, 'healthy/node_modules'), join(directory, 'node_modules'), {
        recursive: true,
      });
      rmSync(join(directory, 'node_modules', name, asset));
      expect(ensure(directory).status).toBe(0);
      expect(calls(directory)).toHaveLength(1);
    },
  );
  it('keeps actual build failures and does not reinstall for a source error', () => {
    const directory = checkout(true),
      result = run(directory, 'labs.mjs');
    expect(result.status).toBe(1);
    expect(result.stderr).toContain('Fixture build failure');
    expect(existsSync(join(directory, 'install-calls.json'))).toBe(false);
  });
  it('preserves npm failure output and exit code without retrying or starting the server', () => {
    const directory = checkout(),
      result = run(directory, 'labs.mjs', { FIXTURE_INSTALL_FAIL: '1' });
    expect(result.status).toBe(7);
    expect(result.stderr).toContain('Fixture npm failure');
    expect(result.stderr).toContain('rerun npm run labs');
    expect(result.stderr).toContain('No private-library token is required');
    expect(result.stdout).not.toContain('Spreadsheet labs: http');
    expect(calls(directory)).toHaveLength(1);
  });
  it('does not continue when npm succeeds without producing a usable install', () => {
    const directory = checkout(),
      result = ensure(directory, { FIXTURE_EMPTY_INSTALL: '1' });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain('still unavailable after installation');
    expect(calls(directory)).toHaveLength(1);
  });
  it('keeps the standalone loader diagnostic for a broken transitive module', () => {
    const directory = checkout(true);
    writeFileSync(
      join(directory, 'node_modules/esbuild/index.js'),
      'import "missing-esbuild-internal-fixture";',
    );
    writeFileSync(
      join(directory, 'scripts/loader-fixture.mjs'),
      'import { loadEsbuild, reportLabsError } from "./lib/labs-setup.mjs"; await loadEsbuild().catch(reportLabsError);',
    );
    const result = run(directory, 'loader-fixture.mjs');
    expect(result.status).toBe(1);
    expect(result.stderr).toContain('missing-esbuild-internal-fixture');
    expect(result.stderr).not.toContain('esbuild is missing');
  });
  it.each(['20.20.2', '22.12.0', '23.11.0'])(
    'reports the Node requirement for %s before attempting to install anything',
    (version) => {
      const directory = checkout();
      const script = `const m = await import(process.argv[1]); try { m.assertLabsNodeVersion(${JSON.stringify(version)}); } catch (error) { m.reportLabsError(error); }`;
      const result = spawnSync(
        process.execPath,
        [
          '--input-type=module',
          '-e',
          script,
          pathToFileURL(join(directory, 'scripts/lib/labs-setup.mjs')).href,
        ],
        { encoding: 'utf8' },
      );
      expect(result.status).toBe(1);
      expect(result.stderr).toContain('Node.js 22.13+ or 24+');
      expect(result.stderr).toContain(version);
      expect(existsSync(join(directory, 'install-calls.json'))).toBe(false);
    },
  );
  it.each(['22.13.0', '22.23.4', '24.0.0', '24.22.2'])('accepts supported Node %s', (version) => {
    const directory = checkout();
    const result = spawnSync(
      process.execPath,
      [
        '--input-type=module',
        '-e',
        'const m = await import(process.argv[1]); m.assertLabsNodeVersion(process.argv[2]);',
        pathToFileURL(join(directory, 'scripts/lib/labs-setup.mjs')).href,
        version,
      ],
      { encoding: 'utf8' },
    );
    expect(result.status).toBe(0);
    expect(result.stderr).toBe('');
  });
  it('prepares a Git dependency with spaces in its path without npm.cmd or a shell', () => {
    const directory = checkout();
    writeFileSync(
      join(directory, 'scripts/build.mjs'),
      'import { writeFileSync } from "node:fs"; writeFileSync(new URL("../prepared.txt", import.meta.url), "prepared");',
    );
    const result = run(directory, 'prepare.mjs', {
      PATH: '',
      npm_package_resolved: 'git+https://example.invalid/editor.git',
      npm_config_ignore_scripts: 'false',
    });
    expect(result.status).toBe(0);
    expect(result.error).toBeUndefined();
    expect(readFileSync(join(directory, 'prepared.txt'), 'utf8')).toBe('prepared');
  });
});
