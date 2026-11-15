import type { ReactNode } from 'react';
import { SUPPORTED_FUNCTIONS } from '../calculation/evaluator.js';
import type { CalculationCellState } from '../calculation/publication.js';
import type { CellValue } from '../calculation/values.js';
import type { CellInput, Workbook } from '../model/workbook.js';

export interface FormulaFunctionDescriptor {
  name: string;
  signature: string;
  description: string;
  aliases?: readonly string[];
  insertText?: string;
}
export const BUILTIN_FORMULA_FUNCTIONS: readonly FormulaFunctionDescriptor[] =
  SUPPORTED_FUNCTIONS.map((name) => ({
    name,
    signature: `${name}(…)`,
    description: `${name} function`,
  }));

/** Registers the upper-case name and aliases of a descriptor, refusing invalid or seen names. */
function descriptorNames(descriptor: FormulaFunctionDescriptor, seen: Set<string>) {
  if ((descriptor.aliases?.length ?? 0) > 32) throw new RangeError('Too many formula aliases.');
  const names = [descriptor.name, ...(descriptor.aliases ?? [])].map((name) => name.toUpperCase());
  for (const name of names) {
    if (!/^[A-Z_][A-Z0-9_.]*$/.test(name) || name.length > 128 || seen.has(name))
      throw new TypeError(`Invalid or duplicate formula function: ${name}`);
    seen.add(name);
  }
  return names;
}

function assertDescriptorText({ signature, description, insertText }: FormulaFunctionDescriptor) {
  const validInsert =
    insertText === undefined || (typeof insertText === 'string' && insertText.length <= 1024);
  if (
    typeof signature !== 'string' ||
    typeof description !== 'string' ||
    signature.length > 1024 ||
    description.length > 8192 ||
    !validInsert
  )
    throw new RangeError('Formula descriptor exceeds the text limit.');
}

export function formulaFunctionCatalog(
  host: readonly FormulaFunctionDescriptor[] = [],
  includeBuiltins = true,
) {
  if (host.length > 1024) throw new RangeError('Too many host formula descriptors.');
  const seen = new Set<string>();
  return [...(includeBuiltins ? BUILTIN_FORMULA_FUNCTIONS : []), ...host].map((descriptor) => {
    const names = descriptorNames(descriptor, seen);
    assertDescriptorText(descriptor);
    return { ...descriptor, name: names[0], aliases: names.slice(1) };
  });
}

export interface CellPresentation {
  text?: string;
  tooltip?: string;
  ariaDescription?: string;
  defaultUnit?: string;
  status?: 'resolved' | 'stale' | 'missing' | 'error';
  decoration?: string;
}
export interface CellPresentationSnapshot {
  inputKey: string;
  revision: number;
  cells: Readonly<Record<string, CellPresentation>>;
}
export interface CellRenderContext {
  cellKey: string;
  sheetId: string;
  x: number;
  y: number;
  inputKey: string;
  revision: number;
  rawInput: CellInput;
  value: CellValue;
  presentation?: CellPresentation;
  calculatedText: string;
}
export type RenderCellContent = (context: CellRenderContext) => ReactNode;
export interface SpreadsheetEditorState {
  dirty: boolean;
  draftDirty: boolean;
  saving: boolean;
  flushing: boolean;
  error: string | null;
  calculation: { inputKey: string; revision: number; status: CalculationCellState } | null;
}
export interface SpreadsheetEditorHandle {
  setPanel(id: string | null): void;
  restoreDraft(draft: { sheetId: string; x: number; y: number; text: string }): void;
  commitDraft(): Workbook;
  cancelDraft(): void;
  getSnapshot(): Workbook;
  select(target: { sheetId: string; range: string }): void;
  getState(): SpreadsheetEditorState;
  subscribe(listener: () => void): () => void;
  save(): Promise<void>;
  flush(): Promise<void>;
}
