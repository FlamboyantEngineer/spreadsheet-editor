import { selectionAddress, writeCells } from '../model/editing.js';
import { createWorkbookEntityId } from '../model/extent.js';
import type {
  WorkbookChangeContext,
  WorkbookCommandResult,
  WorkbookOperation,
} from '../model/operations.js';
import { lockedChange } from '../model/presence.js';
import { attributeCellChanges } from '../model/provenance.js';
import { proposeWorkbookChange, type ReviewStamp } from '../model/review.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { StructureChange } from '../model/structure.js';
import { assertWorkbook, copyWorkbook, type Workbook } from '../model/workbook.js';
import { commandOperations } from './command-operations.js';
import { type CellTargets, immutableBaseline, mergeTargets } from './edit-targets.js';
import type { DraftText, EditorCore } from './editor-core.js';
import { historyEntry } from './history.js';

/** A structure command's line change, which a review proposal records. */
type StructureOperation = { sheetId: string; change: StructureChange };

/** Operations a command already knows, and the cells it wrote, so it needs no workbook diff. */
export interface CommitOptions {
  operations?: WorkbookOperation[] | undefined;
  targets?: CellTargets | undefined;
}
export interface MutateOptions extends CommitOptions {
  operation?: StructureOperation | undefined;
}
export interface DraftMutateOptions {
  operation?: StructureOperation | undefined;
  written?: CellTargets | undefined;
}

export function stampOf({ actor }: EditorCore): ReviewStamp {
  return { id: globalThis.crypto.randomUUID(), time: new Date().toISOString(), actor };
}

function readOnlyMessage({ viewedVersion, props }: EditorCore) {
  if (viewedVersion) return 'A saved version is read-only.';
  if (props.session === undefined) return null;
  if (props.value && typeof props.acceptChange !== 'function')
    return 'A controlled workbook with a session requires acceptChange.';
  return props.session?.workbook() ? null : 'The workbook is still opening.';
}

/** Hands `next` to the host, then to the session or the editor's own state. */
export function publishChange(
  core: EditorCore,
  next: Workbook,
  {
    operations = commandOperations(core.workbook, next),
    source = 'user',
  }: { operations?: WorkbookOperation[]; source?: WorkbookChangeContext['source'] } = {},
) {
  const { props, workbook: previous } = core;
  assertWorkbook(next);
  let baseline: Workbook | undefined;
  const context: WorkbookChangeContext = {
    transactionId: (props.idFactory ?? createWorkbookEntityId)('transaction'),
    source,
    // The detached, frozen copy is made only for a host that reads it (SE-PERF-2.1).
    get baseline() {
      return (baseline ??= immutableBaseline(previous));
    },
    operations: structuredClone(operations),
  };
  const refused = readOnlyMessage(core);
  if (refused) {
    core.setNotice(refused);
    throw new Error(refused);
  }
  let accepted = next;
  try {
    if (props.acceptChange) {
      accepted = props.acceptChange(next, context);
      if (accepted !== next) assertWorkbook(accepted);
    } else {
      props.onChange?.(next, context);
    }
    if (props.session) props.session.change(accepted, context);
    else if (!props.value) core.setLocal(accepted);
  } catch (error) {
    core.setNotice(error instanceof Error ? error.message : 'The workbook change was rejected.');
    throw error;
  }
  core.latestWorkbook.current = accepted;
  return accepted;
}

/** Publishes `next` as one undo step. */
export function commitChange(
  core: EditorCore,
  next: Workbook,
  { label, operations, targets }: CommitOptions & { label: string },
) {
  const { workbook, currentHistory } = core;
  if (next === workbook) return;
  const acceptedOperations = operations ?? commandOperations(workbook, next, targets);
  const accepted = publishChange(core, next, { operations: acceptedOperations });
  // An inverse based on the candidate cannot safely replace fields changed by host rebasing.
  const past =
    accepted === next
      ? [
          ...currentHistory.past.slice(-49),
          historyEntry({ previous: workbook, next }, label, structuredClone(acceptedOperations)),
        ]
      : [];
  core.setHistory({
    head: accepted,
    past,
    future: [],
  });
  core.setEditingCell(false);
}

/** Why a change cannot be made now, already shown to the reader, or null when it can. */
function refusedChange(core: EditorCore, next: Workbook) {
  const { workbook } = core;
  const { canWrite, canEdit, peers } = core.live.current;
  if (!core.canWrite || !canWrite || core.canEdit !== canEdit) {
    core.warn('This workbook is read-only.');
    return true;
  }
  if (core.latestWorkbook.current !== workbook) {
    core.setNotice('The workbook changed before this action completed. Try again.');
    return true;
  }
  const locked = lockedChange(workbook, next, peers);
  if (locked) {
    core.warn(`${locked.name} is editing a cell in this selection. Try again after they finish.`);
    return true;
  }
  return false;
}

/** Commits a change, or proposes it in Reviewing mode, attributing the cells it writes. */
export function mutateChange(
  core: EditorCore,
  next: Workbook,
  { label, operation, operations, targets }: MutateOptions & { label: string },
) {
  if (!core.alive.current || refusedChange(core, next)) return false;
  const { workbook, canSuggest, actor } = core;
  const accepted = canSuggest
    ? proposeWorkbookChange(workbook, { proposed: next, summary: label, operation }, stampOf(core))
    : operation
      ? next
      : attributeCellChanges(workbook, next, {
          actor,
          source: 'human',
          at: new Date().toISOString(),
          targets,
        });
  commitChange(core, accepted, {
    label: canSuggest ? `Suggest: ${label}` : label,
    operations: canSuggest ? undefined : operations,
    targets,
  });
  if (canSuggest) {
    core.setPanel('suggestions');
    core.setPanelOpen(true);
  }
  return true;
}

/** Whether `pending` is a dirty formula-bar draft of the selected cell. */
const isDirtyDraft = (core: EditorCore, pending: DraftText): pending is NonNullable<DraftText> =>
  pending?.key === core.display.draftKey && pending.value !== String(core.display.authored);

/** The workbook with the selected cell's pending draft written, or null when it is locked. */
function withPendingDraft(
  core: EditorCore,
  pending: DraftText,
  operation: StructureOperation | undefined,
) {
  const { workbook, sheet, selection, props, peers, actor } = core;
  const draftTargets: CellTargets = new Map();
  if (!isDirtyDraft(core, pending)) return { base: workbook, draftTargets };
  draftTargets.set(sheet.id, [formatSpreadsheetAddress(selection.x1, selection.y1)]);
  let base = writeCells(workbook, [[pending.value]], {
    sheetId: sheet.id,
    x: selection.x1,
    y: selection.y1,
    idFactory: props.idFactory,
  });
  const locked = lockedChange(workbook, base, peers);
  if (locked) {
    core.warn(`${locked.name} is editing a cell in this selection. Try again after they finish.`);
    return null;
  }
  // Structural transforms move this attribution together with the authored cell; other
  // commands are attributed once, in `mutateChange`.
  if (core.canEdit && operation)
    base = attributeCellChanges(workbook, base, {
      actor,
      source: 'human',
      at: new Date().toISOString(),
      targets: draftTargets,
    });
  return { base, draftTargets };
}

/** Compose a protected formula draft and a chrome command into one history entry. */
export function mutateWithDraftChange(
  core: EditorCore,
  transform: (base: Workbook) => Workbook | WorkbookCommandResult,
  { label, operation, written }: DraftMutateOptions & { label: string },
): boolean {
  const pending = core.draftRef.current;
  const prepared = withPendingDraft(core, pending, operation);
  if (!prepared) return false;
  const { base, draftTargets } = prepared;
  const transformed = transform(base);
  const result =
    'workbook' in transformed ? transformed : { workbook: transformed, operations: undefined };
  const changed = mutateChange(core, result.workbook, {
    label,
    operation,
    operations: result.operations
      ? [...commandOperations(core.workbook, base, draftTargets), ...result.operations]
      : undefined,
    targets: written && mergeTargets(written, draftTargets),
  });
  if (!changed) return false;
  if (core.draftRef.current === pending && pending?.key === core.display.draftKey)
    core.setDraft(null);
  return true;
}

export type ReplaceReason = 'import' | 'new-workbook' | 'reset';

export function replaceWorkbookChange(
  core: EditorCore,
  next: Workbook,
  { label, reason }: { label: string; reason: ReplaceReason },
): boolean {
  const features = core.props.features;
  if (
    (reason === 'import' && features?.import === false) ||
    (reason === 'new-workbook' && features?.newWorkbook === false)
  )
    return false;
  if (
    !core.canEdit ||
    !mutateChange(core, next, {
      label,
      operations: [
        { kind: 'replace', reason, before: copyWorkbook(core.workbook), after: copyWorkbook(next) },
      ],
    })
  )
    return false;
  core.setDraft(null);
  return true;
}

/** Runs a chrome command, showing a thrown or rejected failure as a notice. */
export function attemptAction(core: EditorCore, action: () => unknown) {
  try {
    void Promise.resolve(action()).catch((error) =>
      core.setNotice(error instanceof Error ? error.message : 'The action failed.'),
    );
  } catch (error) {
    core.setNotice(error instanceof Error ? error.message : 'The action failed.');
  }
}

/** Commits the selected cell's dirty formula-bar draft. */
export function applyFormula(core: EditorCore): boolean {
  if (!isDirtyDraft(core, core.draftRef.current)) return true;
  return mutateWithDraftChange(core, (base) => base, {
    label: `Edit ${selectionAddress(core.selection)}`,
    written: new Map(),
  });
}
