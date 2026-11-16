import {
  BUILTIN_FORMULA_FUNCTIONS,
  type CellPresentation,
  type CellPresentationSnapshot,
  type CellRenderContext,
  type ExternalCalculationSnapshot,
  formulaFunctionCatalog,
  type FormulaFunctionDescriptor,
  type NumericValue,
  SpreadsheetEditor,
  type SpreadsheetEditorHandle,
  type SpreadsheetEditorProps,
  type SpreadsheetEditorState,
  type WorkbookChangeContext,
  type WorkbookIdFactory,
} from '@flamboyantengineer/spreadsheet-editor';
import {
  assertExternalCalculationSnapshot,
  calculateWorkbook,
  calculationCellKey,
  type CalculationDependencies,
  type CalculationOptions,
  type CalculationProvider,
  type CalculationResult,
  createClientCalculator,
  type FormulaCellContext,
  type FormulaHost,
  type JsonValue,
  workbookCalculationKey,
} from '@flamboyantengineer/spreadsheet-editor/calculation';
import {
  createDocumentInsertion,
  type DocumentInsertion,
} from '@flamboyantengineer/spreadsheet-editor/document';
import type { ExcelFormulaCapabilities } from '@flamboyantengineer/spreadsheet-editor/excel';
import {
  assertWorkbook,
  changeWorkbookStructureWithOperations,
  createWorkbook,
  invertWorkbookOperations,
  type Workbook,
  type WorkbookOperation,
  worksheetExtent,
} from '@flamboyantengineer/spreadsheet-editor/model';
import { createRef } from 'react';

// This file deliberately imports built public entrypoints, never src/ or internal types.
export const canonical: NumericValue = {
  kind: 'number',
  numerator: '30315',
  denominator: '100',
  approximate: false,
  dimension: [0, 0, 0, 0, 1, 0, 0],
  temperature: 'absolute',
  provenance: { constants: 'immutable-revision' },
};
export const workbook: Workbook = createWorkbook([['30 °C', '=A1']]);
workbook.worksheets[0].inputValues = { A1: canonical };
workbook.worksheets[0].rowIds = ['row-1'];
workbook.worksheets[0].columnIds = ['column-1', 'column-2'];
workbook.worksheets[0].minDimensions = [2, 1];
workbook.worksheets[0].formats = {
  A1: { kind: 'unit', unit: '°C', inputUnit: '°C', locale: 'de-DE', currency: 'EUR' },
};
assertWorkbook(workbook);

export const calculationOptions: CalculationOptions = { precision: 80, wordSizeBits: '64' };
export const dependencies: CalculationDependencies = {
  cells: { [calculationCellKey('sheet-1', 1, 0)]: [calculationCellKey('sheet-1', 0, 0)] },
  extents: { [calculationCellKey('sheet-1', 1, 0)]: ['sheet-1'] },
};
export const opaqueContext: JsonValue = {
  revision: 'immutable',
  records: [true, null, '9007199254740993'],
};
export const formulaContext: FormulaCellContext = {
  sheetId: 'sheet-1',
  x: 0,
  y: 0,
  formula: '=HOSTVALUE(1)',
};
export const result: CalculationResult = calculateWorkbook(workbook, calculationOptions);
export const external: ExternalCalculationSnapshot = {
  inputKey: workbookCalculationKey(workbook),
  revision: 1,
  status: 'ready',
  result,
  cellStates: { [calculationCellKey('sheet-1', 0, 0)]: 'ready' },
  hostContext: { capturedRevision: 'immutable-revision', operands: [null, true, '1E-25'] },
};
assertExternalCalculationSnapshot(external, workbook);

export const hostFunctions: readonly FormulaFunctionDescriptor[] = [
  {
    name: 'HOSTVALUE',
    aliases: ['HOST_VALUE'],
    signature: 'HOSTVALUE(name)',
    description: 'Read a host-authorized value.',
    insertText: 'HOSTVALUE(',
  },
];
export const catalog = formulaFunctionCatalog(hostFunctions);
export const builtins: readonly FormulaFunctionDescriptor[] = BUILTIN_FORMULA_FUNCTIONS;
export const presentation: CellPresentation = {
  text: '30 °C',
  tooltip: 'Captured value',
  ariaDescription: 'Resolved temperature',
  defaultUnit: '°C',
  status: 'resolved',
  decoration: 'external-value',
};
export const cellPresentation: CellPresentationSnapshot = {
  inputKey: external.inputKey,
  revision: external.revision,
  cells: { [calculationCellKey('sheet-1', 0, 0)]: presentation },
};
export const editorRef = createRef<SpreadsheetEditorHandle>();
export const idFactory: WorkbookIdFactory = (kind) => `${kind}-${crypto.randomUUID()}`;

export function inspectTransaction(next: Workbook, context: WorkbookChangeContext): void {
  const operations: readonly WorkbookOperation[] = context.operations;
  const baseline: Workbook = context.baseline;
  for (const operation of operations) {
    if (operation.kind === 'cells')
      for (const cell of operation.changes) {
        const before: NumericValue | undefined = cell.before.inputValue;
        const after: NumericValue | undefined = cell.after.inputValue;
        void before;
        void after;
      }
    if (operation.kind === 'lines') {
      const identities: readonly string[] = operation.ids;
      void identities;
    }
  }
  void next;
  void baseline;
}
export async function exerciseHandle(
  handle: SpreadsheetEditorHandle,
): Promise<SpreadsheetEditorState> {
  const unsubscribe = handle.subscribe(() => {
    void handle.getState();
  });
  const snapshot: Workbook = handle.commitDraft();
  handle.select({ sheetId: snapshot.worksheets[0].id, range: 'A1:B1' });
  handle.cancelDraft();
  void handle.getSnapshot();
  await handle.save();
  await handle.flush();
  unsubscribe();
  return handle.getState();
}
export const modernProps: SpreadsheetEditorProps = {
  value: workbook,
  editorRef,
  idFactory,
  onChange: inspectTransaction,
  calculation: { mode: 'external', snapshot: external },
  formulaFunctions: hostFunctions,
  cellPresentation,
  renderCellContent(context: CellRenderContext) {
    const current:
      NumericValue | string | boolean | null | { readonly kind: 'error'; readonly code: string } =
      context.value;
    void current;
    return context.presentation?.status === 'resolved' ? (
      <span>{context.calculatedText}</span>
    ) : undefined;
  },
  undo: { canUndo: true, canRedo: false, undo: async () => {}, redo: () => {} },
};
// Existing single-argument handlers and providers remain assignable.
export const legacyProvider: CalculationProvider = {
  calculate: async (snapshot) => calculateWorkbook(snapshot),
};
export const legacyProps: SpreadsheetEditorProps = {
  defaultValue: createWorkbook([['1E-25']]),
  onChange: (snapshot) => {
    void snapshot.worksheets;
  },
  calculation: { mode: 'server', provider: legacyProvider },
};
export const legacyHost: FormulaHost = {
  readName: (name) => {
    void name;
    return canonical;
  },
};
export const modernHost: FormulaHost = {
  readName: (_name, context) => {
    void context?.formula;
    return canonical;
  },
};
export const provider = createClientCalculator(modernHost);
export const table: DocumentInsertion = createDocumentInsertion(workbook, result, {
  mode: 'table',
});
export const extent = worksheetExtent(workbook.worksheets[0], workbook);
export const structure = changeWorkbookStructureWithOperations(
  workbook,
  { sheetId: 'sheet-1', idFactory },
  { axis: 'row', kind: 'insert', index: 1, count: 1 },
);
export const inverse: WorkbookOperation[] = invertWorkbookOperations(structure.operations);
export const capabilities: ExcelFormulaCapabilities = {
  hostFunctions: [{ name: 'HOSTVALUE', capturedText: 'html-to-text' }],
};
export const modernElement = <SpreadsheetEditor {...modernProps} />;
export const legacyElement = <SpreadsheetEditor {...legacyProps} />;

// Version-one granular application, causal receipt and linked-capture contracts.
export {
  createLinkedSpreadsheetInsertion,
  type LinkedSpreadsheetInsertion,
  updateLinkedSpreadsheetInsertion,
} from '@flamboyantengineer/spreadsheet-editor/document';
export { createTransactionReceiptTracker } from '@flamboyantengineer/spreadsheet-editor/host';
export {
  createWorkbookOperationBinding,
  initializeWorkbookIdentities,
} from '@flamboyantengineer/spreadsheet-editor/model';
