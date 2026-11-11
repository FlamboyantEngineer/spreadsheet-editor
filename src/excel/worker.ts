import { exportXlsx, exportXlsxWithReport } from './write.js';
import { validateFormulaCapabilities } from './capabilities.js';
import { decodeImportRequest } from './protocol.js';
import { convertXlsx } from './convert.js';

// The importer owns one process. Neither heap exhaustion nor malformed XML can take down its host.
const chunks: Buffer[] = [];
let size = 0;
const memory = setInterval(() => {
  if (process.memoryUsage().rss > 768 * 1024 * 1024) process.exit(1);
}, 100);
process.stdin.on('data', (chunk: Buffer) => {
  size += chunk.length;
  if (size > (process.argv[2] === 'import' ? 20 * 1024 * 1024 + 65_540 : 100 * 1024 * 1024))
    process.exit(1);
  chunks.push(chunk);
});
process.stdin.on('end', () => {
  const bytes = Buffer.concat(chunks);
  const task = async () => {
    if (process.argv[2] === 'export' || process.argv[2] === 'export-report') {
      const { workbook, result, formulaCapabilities: supplied } = JSON.parse(bytes.toString());
      const formulaCapabilities = validateFormulaCapabilities(supplied);
      if (process.argv[2] === 'export-report') {
        const output = await exportXlsxWithReport(workbook, result, { formulaCapabilities });
        process.stdout.write(
          JSON.stringify({
            bytes: Buffer.from(output.bytes).toString('base64'),
            report: output.report,
          }),
        );
      } else process.stdout.write(await exportXlsx(workbook, result, { formulaCapabilities }));
    } else if (process.argv[2] === 'import') {
      const request = decodeImportRequest(bytes);
      process.stdout.write(
        JSON.stringify(await convertXlsx(request.bytes, {}, request.formulaCapabilities)),
      );
    } else throw new TypeError('Unknown Excel worker operation.');
  };
  task()
    .catch((error: unknown) => {
      process.stderr.write(error instanceof Error ? error.message : 'Excel processing failed.');
      process.exitCode = 1;
    })
    .finally(() => clearInterval(memory));
});
