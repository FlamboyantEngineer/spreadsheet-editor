import { useMemo } from 'react';
import { calculationCellKey, pendingCalculationCells } from '../calculation/publication.js';
import { createClientCalculator } from '../calculation/provider.js';
import { parseCellStyle, selectionAddress, type CellSelection } from '../model/editing.js';
import { worksheetCalculationExtent } from '../model/extent.js';
import { previewSuggestions } from '../model/review.js';
import type { Workbook, Worksheet } from '../model/workbook.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { NO_CELLS, readyValues, withoutPending, withoutPendingValues } from './pending-values.js';
import {
  acceptedPresentation,
  usePublication,
  useRecalculatingDisplay,
} from './use-publication.js';

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
}: {
  props: SpreadsheetEditorProps;
  workbook: Workbook;
  sheet: Worksheet;
  selection: CellSelection;
  draft: { key: string; value: string } | null;
  nativeDraft: { x: number; y: number; text: string } | null;
  canSuggest: boolean;
  actorId: string;
  refreshVersion: number;
}) {
  const client = useMemo(
    () =>
      createClientCalculator(
        props.calculation?.mode === 'client' ? props.calculation.host : undefined,
      ),
    [props.calculation?.mode === 'client' ? props.calculation.host : undefined],
  );
  const baseProvider = props.calculation?.mode === 'server' ? props.calculation.provider : client;
  const provider = useMemo(
    () => ({ calculate: baseProvider.calculate.bind(baseProvider) }),
    [baseProvider, refreshVersion],
  );
  const calculationOptions = useMemo(
    () => ({
      timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      ...(props.calculation?.mode === 'external' ? undefined : props.calculation?.options),
    }),
    [props.calculation?.mode === 'external' ? undefined : props.calculation?.options],
  );
  const calculation = usePublication(workbook, props.calculation, provider, calculationOptions);
  const previewWorkbook = useMemo(
    () => (canSuggest ? previewSuggestions(workbook, actorId) : workbook),
    [workbook, actorId, canSuggest],
  );
  const previewCalculation = usePublication(
    previewWorkbook,
    props.calculation?.mode === 'external'
      ? { mode: 'external', snapshot: props.reviewCalculation ?? null }
      : props.calculation,
    provider,
    calculationOptions,
    previewWorkbook !== workbook,
  );
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
  const previewSheet = previewWorkbook.worksheets.find((value) => value.id === sheet.id)!;
  const authored = previewSheet.data[selection.y1]?.[selection.x1] ?? '';
  const draftKey = `${sheet.id}:${selection.x1}:${selection.y1}`;
  const draftValue = draft?.key === draftKey ? draft.value : String(authored);
  const draftDirty = draft?.key === draftKey && draft.value !== String(authored);
  const draftCell = draftDirty
    ? { x: selection.x1, y: selection.y1, text: draft!.value }
    : nativeDraft;
  const calculationExtent = useMemo(() => worksheetCalculationExtent(sheet), [sheet]);
  const pendingCells = useMemo(
    () =>
      draftCell
        ? pendingCalculationCells(
            shown.result,
            [calculationCellKey(sheet.id, draftCell.x, draftCell.y)],
            draftCell.x >= calculationExtent.columns || draftCell.y >= calculationExtent.rows
              ? [sheet.id]
              : [],
          )
        : NO_CELLS,
    [
      shown.result,
      sheet.id,
      calculationExtent.columns,
      calculationExtent.rows,
      draftCell?.x,
      draftCell?.y,
      draftCell?.text,
    ],
  );
  const readyCells = useMemo(
    () => withoutPending(shown.readyCells, pendingCells),
    [shown.readyCells, pendingCells],
  );
  const presentation = acceptedPresentation(
    props.cellPresentation,
    visibleCalculation.inputKey,
    visibleCalculation.revision,
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
  const style = parseCellStyle(
    previewSheet.style?.[selectionAddress({ ...selection, x2: selection.x1, y2: selection.y1 })],
  );
  const cellFormat = previewSheet.formats?.[
    selectionAddress({ ...selection, x2: selection.x1, y2: selection.y1 })
  ] ?? { kind: 'general' as const };
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
