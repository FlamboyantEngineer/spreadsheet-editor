import { validateFormulaCapabilities } from './capabilities.js';
import { convertXlsx } from './convert.js';
import { decodeImportRequest, encodeExcelHeader, excelInputLimit } from './protocol.js';
import { exportXlsxWithReport } from './write.js';

// The importer owns one process. Neither heap exhaustion nor malformed XML can take down its host.
const chunks: Buffer[] = [];
const operation = process.argv[2];
if (operation !== 'import' && operation !== 'export')
  throw new TypeError('Unknown Excel worker operation.');
const inputLimit = excelInputLimit(operation);
let size = 0;
const memory = setInterval(() => {
  if (process.memoryUsage().rss > 768 * 1024 * 1024) process.exit(1);
}, 100);
process.stdin.on('data', (chunk: Buffer) => {
  size += chunk.length;
  if (size > inputLimit) process.exit(1);
  chunks.push(chunk);
});
process.stdin.on('end', () => {
  const requestBytes = Buffer.concat(chunks);
  const task = async () => {
    if (operation === 'export') {
      const {
        workbook,
        result,
        formulaCapabilities: supplied,
      } = JSON.parse(requestBytes.toString());
      const formulaCapabilities = validateFormulaCapabilities(supplied);
      const { report, bytes } = await exportXlsxWithReport(workbook, result, {
        formulaCapabilities,
      });
      process.stdout.write(encodeExcelHeader({ report }));
      process.stdout.write(bytes);
    } else {
      const request = decodeImportRequest(requestBytes);
      process.stdout.write(
        JSON.stringify(await convertXlsx(request.bytes, {}, request.formulaCapabilities)),
      );
    }
  };
  task()
    .catch((error: unknown) => {
      process.stderr.write(error instanceof Error ? error.message : 'Excel processing failed.');
      process.exitCode = 1;
    })
    .finally(() => clearInterval(memory));
});
