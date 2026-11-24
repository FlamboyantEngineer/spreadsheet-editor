import type { SpreadsheetCopy } from '../model/clipboard.js';
import type { CellInput } from '../model/workbook.js';

/** Plain text plus the optional opaque identity of an editor copy. */
export interface SpreadsheetClipboardData {
  text: string;
  origin?: string;
}

export interface ClipboardCopy {
  copy: SpreadsheetCopy;
  origin: string;
}

const originPattern = /\bdata-spreadsheet-editor-origin="(\d{1,10}(?:-\d{1,10}){3})"/;
const htmlCharacters: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;' };

/** HTML is shared by both browser clipboard entry points, unlike custom clipboard types. */
export function clipboardHtml(origin: string, rows: CellInput[][]) {
  const body = rows
    .map(
      (row) =>
        `<tr>${row.map((value) => `<td>${String(value ?? '').replace(/[&<>]/g, (character) => htmlCharacters[character])}</td>`).join('')}</tr>`,
    )
    .join('');
  return `<table data-spreadsheet-editor-origin="${origin}"><tbody>${body}</tbody></table>`;
}

function browserData(text: string, html: string): SpreadsheetClipboardData {
  const origin = originPattern.exec(html.slice(0, 4096))?.[1];
  return origin ? { text, origin } : { text };
}

export function readClipboardEvent(transfer: DataTransfer): SpreadsheetClipboardData | null {
  if (!transfer.types.includes('text/plain')) return null;
  return browserData(transfer.getData('text/plain'), transfer.getData('text/html'));
}

export function writeClipboardEvent(transfer: DataTransfer, data: { text: string; html: string }) {
  transfer.setData('text/plain', data.text);
  transfer.setData('text/html', data.html);
}

function browserClipboard() {
  const clipboard = globalThis.navigator?.clipboard;
  if (!clipboard) throw new Error('Clipboard access is unavailable. Use the grid shortcut.');
  return clipboard;
}

export async function readBrowserClipboard(): Promise<SpreadsheetClipboardData | null> {
  const clipboard = browserClipboard();
  if (typeof clipboard.read === 'function') {
    try {
      const item = (await clipboard.read()).find((item) => item.types.includes('text/plain'));
      if (!item) return null;
      const text = await (await item.getType('text/plain')).text();
      const html = item.types.includes('text/html')
        ? await (await item.getType('text/html')).slice(0, 4096).text()
        : '';
      return browserData(text, html);
    } catch {
      // Browsers may grant access to text only.
    }
  }
  return { text: await clipboard.readText() };
}

/** Returns whether the browser retained the origin alongside the text. */
export async function writeBrowserClipboard(data: {
  text: string;
  html: string;
}): Promise<boolean> {
  const clipboard = browserClipboard();
  if (typeof clipboard.write === 'function' && typeof ClipboardItem === 'function') {
    try {
      await clipboard.write([
        new ClipboardItem({
          'text/plain': new Blob([data.text], { type: 'text/plain' }),
          'text/html': new Blob([data.html], { type: 'text/html' }),
        }),
      ]);
      return true;
    } catch {
      // Browsers that cannot write custom formats can still copy plain text.
    }
  }
  await clipboard.writeText(data.text);
  return false;
}
