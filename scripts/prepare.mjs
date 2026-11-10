import { execFileSync } from 'node:child_process';
if (process.env.npm_package_resolved && process.env.npm_config_ignore_scripts !== 'true') {
  execFileSync(process.platform === 'win32' ? 'npm.cmd' : 'npm', ['run', 'build'], {
    stdio: 'inherit',
  });
}
