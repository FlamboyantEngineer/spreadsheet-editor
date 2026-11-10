import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
if (process.env.npm_package_resolved && process.env.npm_config_ignore_scripts !== 'true') {
  execFileSync(process.execPath, [fileURLToPath(new URL('./build.mjs', import.meta.url))], {
    stdio: 'inherit',
  });
}
