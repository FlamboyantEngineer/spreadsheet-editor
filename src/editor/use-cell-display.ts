'use client';

import { useMemo } from 'react';
import { createClientCalculator } from '../calculation/provider.js';
import { calculationCellKey, pendingCalculationCells } from '../calculation/publication.js';
import { type CellSelection, parseCellStyle, selectionAddress } from '../model/editing.js';
import { worksheetCalculationExtent } from '../model/extent.js';
import { requireDefined } from '../model/require-defined.js';
import { previewSuggestions } from '../model/review.js';
import type { Workbook, Worksheet } from '../model/workbook.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { NO_CELLS, readyValues, withoutPending, withoutPendingValues } from './pending-values.js';
import {
  acceptedPresentation,
  usePublication,
  useRecalculatingDisplay,
} from './use-publication.js';

/** The provider and options every publication of this editor calculates with. */
function useCalculationSetup(props: SpreadsheetEditorProps, refreshVersion: number) {
  const clientHost = props.calculation?.mode === 'client' ? props.calculation.host : undefined;
  const client = useMemo(() => createClientCalculator(clientHost), [clientHost]);
  const baseProvider = props.calculation?.mode === 'server' ? props.calculation.provider : client;
  const provider = useMemo(
    () => ({ calculate: baseProvider.calculate.bind(baseProvider) }),
    // A new provider object on Refresh starts a new calculation of the same workbook.
    // eslint-disable-next-line react-hooks/exhaustive-deps -- refreshVersion is the trigger
    [baseProvider, refreshVersion],
  );
  const hostOptions =
    props.calculation?.mode === 'external' ? undefined : props.calculation?.options;
  const options = useMemo(
    () => ({ timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone, ...hostOptions }),
    [hostOptions],
  );
  return { provider, options };
}

interface ShownCalculationInputs {
  props: SpreadsheetEditorProps;
  workbook: Workbook;
  canSuggest: boolean;
  actorId: string;
  refreshVersion: number;
}

/** The accepted calculation, the reviewer's suggestion preview, and what the grid shows of them. */
function useShownCalculation({
  props,
  workbook,
  canSuggest,
  actorId,
  refreshVersion,
}: ShownCalculationInputs) {
  const { provider, options } = useCalculationSetup(props, refreshVersion);
  const calculation = usePublication(workbook, { mode: props.calculation, provider, options });
  const previewWorkbook = useMemo(
    () => (canSuggest ? previewSuggestions(workbook, actorId) : workbook),
    [workbook, actorId, canSuggest],
  );
  const previewCalculation = usePublication(previewWorkbook, {
    mode:
      props.calculation?.mode === 'external'
        ? { mode: 'external', snapshot: props.reviewCalculation ?? null }
        : props.calculation,
    provider,
    options,
    enabled: previewWorkbook !== workbook,
  });
  const visibleCalculation = previewWorkbook !== workbook ? previewCalculation : calculation;
  const shown = useRecalculatingDisplay(
    previewWorkbook,
    visibleCalculation,
    props.calculation?.mode !== 'external',
  );
  const publication = useMemo(
    () => ({ inputKey: visibleCalculation.inputKey, revision: visibleCalculation.revision }),
    [visibleCalculation.inputKey, visibleCalculation.revision],
  );
  return { calculation, previewWorkbook, visibleCalculation, shown, publication };
}

type DraftText = { key: string; value: string } | null;
type DraftCell = { x: number; y: number; text: string } | null;

/** The selected cell's authored input and its formula-bar draft, if that draft is dirty. */
function selectedDraft(
  { sheet, previewSheet }: { sheet: Worksheet; previewSheet: Worksheet },
  selection: CellSelection,
  { draft, nativeDraft }: { draft: DraftText; nativeDraft: DraftCell },
) {
  const authored = previewSheet.data[selection.y1]?.[selection.x1] ?? '';
  const draftKey = `${sheet.id}:${selection.x1}:${selection.y1}`;
  const draftValue = draft?.key === draftKey ? draft.value : String(authored);
  const draftDirty = draft?.key === draftKey && draft.value !== String(authored);
  const draftCell = draftDirty
    ? { x: selection.x1, y: selection.y1, text: requireDefined(draft).value }
    : nativeDraft;
  return { authored, draftKey, draftValue, draftDirty, draftCell };
}

/** The ready cells and values of the active worksheet, with a draft's pending cells blanked. */
function usePendingDisplay(
  shown: ReturnType<typeof useRecalculatingDisplay>,
  sheet: Worksheet,
  draftCell: DraftCell,
) {
  const calculationExtent = useMemo(() => worksheetCalculationExtent(sheet), [sheet]);
  const draftX = draftCell?.x,
    draftY = draftCell?.y;
  const pendingCells = useMemo(
    () =>
      draftX !== undefined && draftY !== undefined
        ? pendingCalculationCells(
            shown.result,
            [calculationCellKey(sheet.id, draftX, draftY)],
            draftX >= calculationExtent.columns || draftY >= calculationExtent.rows
              ? [sheet.id]
              : [],
          )
        : NO_CELLS,
    [shown.result, sheet.id, calculationExtent.columns, calculationExtent.rows, draftX, draftY],
  );
  const readyCells = useMemo(
    () => withoutPending(shown.readyCells, pendingCells),
    [shown.readyCells, pendingCells],
  );
  const acceptedOutput = shown.result?.worksheets.find((result) => result.id === sheet.id)?.values;
  // Ready values change with the result; a draft only blanks the rows its pending cells are in.
  const readyOutput = useMemo(
    () => readyValues(acceptedOutput, shown.readyCells, sheet.id),
    [acceptedOutput, shown.readyCells, sheet.id],
  );
  const output = useMemo(
    () => withoutPendingValues(readyOutput, pendingCells, sheet.id),
    [readyOutput, pendingCells, sheet.id],
  );
  return { readyCells, output };
}

/** The style and number format of the selection's first cell. */
function selectedFormatting(previewSheet: Worksheet, selection: CellSelection) {
  const address = selectionAddress({ ...selection, x2: selection.x1, y2: selection.y1 });
  return {
    style: parseCellStyle(previewSheet.style?.[address]),
    cellFormat: previewSheet.formats?.[address] ?? { kind: 'general' as const },
  };
}

/**
 * What the grid and chrome show for the active worksheet: the calculation and its suggestion
 * preview, the selected cell's draft, and the ready values with a draft's pending cells blanked.
 */
export function useCellDisplay({
  props,
  workbook,
  sheet,
  selection,
  draft,
  nativeDraft,
  canSuggest,
  actorId,
  refreshVersion,
}: ShownCalculationInputs & {
  sheet: Worksheet;
  selection: CellSelection;
  draft: DraftText;
  nativeDraft: DraftCell;
}) {
  const { calculation, previewWorkbook, visibleCalculation, shown, publication } =
    useShownCalculation({ props, workbook, canSuggest, actorId, refreshVersion });
  const previewSheet = requireDefined(
    previewWorkbook.worksheets.find((value) => value.id === sheet.id),
  );
  const { authored, draftKey, draftValue, draftDirty, draftCell } = selectedDraft(
    { sheet, previewSheet },
    selection,
    { draft, nativeDraft },
  );
  const { readyCells, output } = usePendingDisplay(shown, sheet, draftCell);
  const presentation = acceptedPresentation(props.cellPresentation, visibleCalculation);
  const { style, cellFormat } = selectedFormatting(previewSheet, selection);
  return {
    authored,
    calculation,
    cellFormat,
    draftCell,
    draftDirty,
    draftKey,
    draftValue,
    output,
    presentation,
    previewSheet,
    previewWorkbook,
    publication,
    readyCells,
    shown,
    style,
    visibleCalculation,
  };
}
