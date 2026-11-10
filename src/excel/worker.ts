import { exportXlsx } from './write.js';
import { convertXlsx } from './convert.js';

// The importer owns one process. Neither heap exhaustion nor malformed XML can take down its host.
const chunks: Buffer[] = [];
let size = 0;
const memory = setInterval(() => {
  if (process.memoryUsage().rss > 768 * 1024 * 1024) process.exit(1);
}, 100);
process.stdin.on('data', (chunk: Buffer) => {
  size += chunk.length;
  if (size > (process.argv[2] === 'export' ? 100 : 20) * 1024 * 1024) process.exit(1);
  chunks.push(chunk);
});
process.stdin.on('end', () => {
  const bytes = Buffer.concat(chunks);
  const task = async () => {
    if (process.argv[2] === 'export') {
      const { workbook, result } = JSON.parse(bytes.toString());
      process.stdout.write(await exportXlsx(workbook, result));
    } else process.stdout.write(JSON.stringify(await convertXlsx(bytes)));
  };
  task()
    .catch((error: unknown) => {
      process.stderr.write(error instanceof Error ? error.message : 'Excel processing failed.');
      process.exitCode = 1;
    })
    .finally(() => clearInterval(memory));
});
