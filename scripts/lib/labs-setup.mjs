export class LabsSetupError extends Error {}

export function assertLabsNodeVersion(version = process.versions.node) {
  if (Number(version.split('.')[0]) < 22) {
    throw new LabsSetupError(
      `Spreadsheet labs require Node.js 22 or newer; found ${version}. Switch Node versions, then rerun npm run labs.`,
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
        'The labs bundler became unavailable. Rerun npm run labs to install or repair dependencies automatically.\n' +
          'The ui-components dependency is public: https://github.com/FlamboyantEngineer/ui-components. No private-library token is required.',
      );
    }
    throw error;
  }
}

export function reportLabsError(error) {
  console.error(error instanceof LabsSetupError ? error.message : error);
  process.exitCode = error instanceof LabsSetupError ? (error.exitCode ?? 1) : 1;
}
