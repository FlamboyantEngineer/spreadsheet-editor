import { forgetCell } from './cell-paint.js';

/** Track cells whose contents the engine changed between presentation passes. */
export function observeGridCellWrites(table: HTMLTableElement) {
  const Observer = table.ownerDocument.defaultView?.MutationObserver;
  if (!Observer) return undefined;
  const rows = new Set<number>();
  function capture(records: MutationRecord[]) {
    for (const record of records) {
      const node = record.type === 'characterData' ? record.target.parentElement : record.target;
      if (!(node instanceof HTMLElement) || node.tagName !== 'TD' || node.dataset.x === undefined)
        continue;
      const y = Number(node.dataset.y);
      if (Number.isSafeInteger(y) && y >= 0) {
        rows.add(y);
        forgetCell(node);
      }
    }
  }
  const observer = new Observer(capture);
  observer.observe(table, { subtree: true, childList: true, characterData: true });
  return {
    table,
    changedRows() {
      capture(observer.takeRecords());
      const dirty = new Set(rows);
      rows.clear();
      return dirty;
    },
    ignorePaint() {
      observer.takeRecords();
    },
    dispose() {
      observer.disconnect();
    },
  };
}
