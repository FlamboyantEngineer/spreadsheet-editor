export class LabsSetupError extends Error {}

export function assertLabsNodeVersion(version = process.versions.node) {
  if (Number(version.split('.')[0]) < 22) {
    throw new LabsSetupError(
      `Spreadsheet labs require Node.js 22 or newer; found ${version}. Switch Node versions, then run npm ci --include=dev.`,
    );
  }
}

export async function loadEsbuild() {
  assertLabsNodeVersion();
  try {
    return await import('esbuild');
  } catch (error) {
    if (error.code === 'ERR_MODULE_NOT_FOUND' && error.message.includes("package 'esbuild'")) {
      throw new LabsSetupError(
        'Spreadsheet labs need development dependencies (esbuild is missing).\n' +
          'From the repository root, run:\n\n' +
          '  npm ci --include=dev\n' +
          '  npm run labs\n\n' +
          'If installation fails while fetching https://github.com/FlamboyantEngineer/ui-components, resolve the installation error and rerun npm ci --include=dev. This dependency is public; no private-library token is required.',
      );
    }
    throw error;
  }
}

export function reportLabsError(error) {
  console.error(error instanceof LabsSetupError ? error.message : error);
  process.exitCode = 1;
}
