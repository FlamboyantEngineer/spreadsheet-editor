import { spawnSync } from 'node:child_process';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';

const scripts = fileURLToPath(new URL('../../scripts/', import.meta.url));
const fixtures: string[] = [];
afterEach(() => {
  for (const directory of fixtures.splice(0)) rmSync(directory, { recursive: true, force: true });
});

function checkout() {
  const directory = mkdtempSync(join(tmpdir(), 'spreadsheet labs '));
  fixtures.push(directory);
  cpSync(scripts, join(directory, 'scripts'), { recursive: true });
  return directory;
}
function run(directory: string, script: string, env: NodeJS.ProcessEnv = process.env) {
  return spawnSync(process.execPath, [join(directory, 'scripts', script)], {
    cwd: directory,
    env,
    encoding: 'utf8',
    timeout: 10_000,
  });
}
function fakeEsbuild(directory: string, source: string) {
  const module = join(directory, 'node_modules', 'esbuild');
  mkdirSync(module, { recursive: true });
  writeFileSync(
    join(module, 'package.json'),
    JSON.stringify({ name: 'esbuild', type: 'module', exports: './index.js' }),
  );
  writeFileSync(join(module, 'index.js'), source);
}

describe('SE-LABS-2: startup diagnostics and portable preparation', () => {
  it.each(['labs.mjs', 'build-labs.mjs'])(
    'explains missing development dependencies for %s in a fresh checkout',
    (script) => {
      const directory = checkout(),
        result = run(directory, script);
      expect(result.status).toBe(1);
      expect(result.error).toBeUndefined();
      expect(result.stderr).toContain('esbuild is missing');
      expect(result.stderr).toContain('npm ci --include=dev');
      expect(result.stderr).toContain('https://github.com/FlamboyantEngineer/ui-components');
      expect(result.stderr).toContain('no private-library token is required');
      expect(result.stderr).not.toContain('ERR_MODULE_NOT_FOUND');
    },
  );
  it('preserves actual build failures', () => {
    const directory = checkout();
    fakeEsbuild(directory, 'export function build() { throw new Error("Fixture build failure"); }');
    const result = run(directory, 'labs.mjs');
    expect(result.status).toBe(1);
    expect(result.stderr).toContain('Fixture build failure');
    expect(result.stderr).not.toContain('esbuild is missing');
  });
  it('does not misreport a broken transitive module as a missing esbuild install', () => {
    const directory = checkout();
    fakeEsbuild(directory, 'import "missing-esbuild-internal-fixture";');
    const result = run(directory, 'build-labs.mjs');
    expect(result.status).toBe(1);
    expect(result.stderr).toContain('missing-esbuild-internal-fixture');
    expect(result.stderr).not.toContain('esbuild is missing');
  });
  it('reports the Node requirement without loading the bundler', () => {
    const directory = checkout();
    const script =
      'const m = await import(process.argv[1]); try { m.assertLabsNodeVersion("20.19.0"); } catch (error) { m.reportLabsError(error); }';
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
    expect(result.stderr).toContain('Node.js 22 or newer');
    expect(result.stderr).toContain('20.19.0');
  });
  it('prepares a Git dependency with spaces in its path without npm.cmd or a shell', () => {
    const directory = checkout();
    writeFileSync(
      join(directory, 'scripts/build.mjs'),
      'import { writeFileSync } from "node:fs"; writeFileSync(new URL("../prepared.txt", import.meta.url), "prepared");',
    );
    const result = run(directory, 'prepare.mjs', {
      ...process.env,
      PATH: '',
      npm_package_resolved: 'git+https://example.invalid/editor.git',
      npm_config_ignore_scripts: 'false',
    });
    expect(result.status).toBe(0);
    expect(result.error).toBeUndefined();
    expect(readFileSync(join(directory, 'prepared.txt'), 'utf8')).toBe('prepared');
  });
});
