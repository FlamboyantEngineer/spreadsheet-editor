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
import {
  READ_ONLY_CAPABILITIES,
  type WorkbookSession,
  type WorkbookStoreAdapter,
} from '@flamboyantengineer/spreadsheet-editor/host';

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

/**
 * A workbook store for a host that persists each accepted transaction through its own binding: a
 * save is a durability barrier for the transactions it covers, not a second write. Open a session
 * over it with `useWorkbookSession({ workbooks, workbookId })`.
 */
export function createDurabilityStore(
  current: () => Workbook,
  provider: HostDurabilityProvider,
): WorkbookStoreAdapter {
  let sequence = 0;
  return {
    async load(id) {
      return {
        ok: true,
        value: { id, content: current(), revision: `stored-${sequence}`, sequence },
      };
    },
    async save() {
      if (!(await provider.whenStored()))
        return {
          ok: false,
          error: { kind: 'unavailable', message: 'The workbook has not reached durable storage.' },
        };
      sequence++;
      return { ok: true, value: { revision: `stored-${sequence}`, sequence } };
    },
  };
}

export interface HostIntegrationProps {
  sessionKey: string;
  workbook: Workbook;
  publication: ExternalCalculationSnapshot | null;
  binding: HostWorkbookBinding;
  /** The session over the host's store, such as `createDurabilityStore`'s. */
  session: WorkbookSession | null;
  canEdit: boolean;
  editorRef?: Ref<SpreadsheetEditorHandle>;
  undo?: SpreadsheetEditorProps['undo'];
  formulaFunctions?: readonly FormulaFunctionDescriptor[];
  cellPresentation?: CellPresentationSnapshot;
  renderCellContent?: RenderCellContent;
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
      capabilities={{ ...READ_ONLY_CAPABILITIES, edit: props.canEdit }}
      features={{ comments: false, suggestions: false, modeControl: false }}
      undo={props.undo}
      formulaFunctions={props.formulaFunctions}
      cellPresentation={props.cellPresentation}
      renderCellContent={props.renderCellContent}
      onChange={(next, transaction) => props.binding.applyTransaction(next, transaction)}
      session={props.session}
      onRefresh={props.onRefresh}
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
