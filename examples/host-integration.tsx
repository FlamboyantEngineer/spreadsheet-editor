import type { Ref } from 'react';
import {
  SpreadsheetEditor,
  createWorkbookOperationBinding,
  type CellPresentationSnapshot,
  type ExternalCalculationSnapshot,
  type FormulaFunctionDescriptor,
  type RenderCellContent,
  type SpreadsheetEditorHandle,
  type SpreadsheetEditorProps,
  type Workbook,
  type WorkbookChangeContext,
} from '@flamboyantengineer/spreadsheet-editor';

/** Project the host's current accepted state and commit synchronously in its serialized lane. */
export function createHostWorkbookBinding(host: {
  getSnapshot(): Workbook;
  /** Authorize and commit atomically; throw before mutation on rejection. */
  commit(next: Workbook, transaction: WorkbookChangeContext): void;
}): HostWorkbookBinding {
  return createWorkbookOperationBinding(host);
}

/** Host persistence and transport remain independent of operation application. */
export interface HostWorkbookBinding {
  /** Preflight the complete transaction before mutation; throw only before applying it. */
  applyTransaction(next: Workbook, transaction: WorkbookChangeContext): void;
}

/** A projection of an existing host provider; this example never creates or destroys it. */
export interface HostDurabilityProvider {
  whenStored(): Promise<boolean>;
}

export interface HostIntegrationProps {
  sessionKey: string;
  workbook: Workbook;
  publication: ExternalCalculationSnapshot | null;
  binding: HostWorkbookBinding;
  provider: HostDurabilityProvider;
  canEdit: boolean;
  editorRef?: Ref<SpreadsheetEditorHandle>;
  undo?: SpreadsheetEditorProps['undo'];
  formulaFunctions?: readonly FormulaFunctionDescriptor[];
  cellPresentation?: CellPresentationSnapshot;
  renderCellContent?: RenderCellContent;
  onSave?: SpreadsheetEditorProps['onSave'];
  onRefresh?: SpreadsheetEditorProps['onRefresh'];
}

/**
 * The host projects accepted native snapshots and authoritative publications into props.
 * Opening or receiving a publication does not run the local calculator or write a snapshot.
 * Keep a stable binding/provider per session and supply trusted host edit capabilities.
 */
export function HostIntegratedSpreadsheet(props: HostIntegrationProps) {
  return (
    <SpreadsheetEditor
      sessionKey={props.sessionKey}
      value={props.workbook}
      editorRef={props.editorRef}
      calculation={{ mode: 'external', snapshot: props.publication }}
      readOnly={!props.canEdit}
      permissions={{ edit: props.canEdit, comment: false, suggest: false }}
      features={{ comments: false, suggestions: false, modeControl: false }}
      undo={props.undo}
      formulaFunctions={props.formulaFunctions}
      cellPresentation={props.cellPresentation}
      renderCellContent={props.renderCellContent}
      onChange={(next, transaction) => props.binding.applyTransaction(next, transaction)}
      onSave={props.onSave}
      onRefresh={props.onRefresh}
      onFlush={async (_snapshot, { signal }) => {
        // commitDraft already invoked the synchronous binding before this durability barrier.
        signal.throwIfAborted();
        const stored = await props.provider.whenStored();
        signal.throwIfAborted();
        if (!stored) throw new Error('The workbook has not reached durable storage.');
      }}
    />
  );
}

/** A failed draft commit or receipt keeps the current route and editor mounted. */
export async function flushBeforeNavigation(
  editor: SpreadsheetEditorHandle,
  navigate: () => void | Promise<void>,
): Promise<void> {
  await editor.flush();
  await navigate();
}
