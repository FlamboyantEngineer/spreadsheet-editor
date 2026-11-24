import { type CellPosition, keyFor } from './cell-evaluation.js';

interface ActiveCell {
  cell: CellPosition;
  key: string;
  cyclic: boolean;
}

/** The suspended formula path and the members of any cycle discovered on it. */
export class FormulaStack {
  private readonly frames: ActiveCell[] = [];
  private readonly indexes = new Map<string, number>();

  get current() {
    return this.frames.at(-1);
  }

  push(cell: CellPosition): void {
    const key = keyFor(cell);
    this.indexes.set(key, this.frames.length);
    this.frames.push({ cell, key, cyclic: false });
  }

  pop(): void {
    const frame = this.frames.pop();
    if (frame) this.indexes.delete(frame.key);
  }

  /** Abort the active read so an error-handling function inside the cycle cannot catch it. */
  rejectCycle(target: string): void {
    const first = this.indexes.get(target);
    if (first === undefined) return;
    for (let index = first; index < this.frames.length; index++) this.frames[index].cyclic = true;
    throw new Error('Circular formula dependency.');
  }
}
