'use client';

import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CellValue } from '../calculation/values.js';
import type { DisplayOptions } from '../calculation/format.js';
import type { CellInput, Worksheet } from '../model/workbook.js';
import { createWorkbook } from '../model/workbook.js';
import type { SpreadsheetComment, SpreadsheetSuggestion } from '../model/review.js';
import type { GridController } from '../editor/contracts.js';
import { writeCells, parseDelimited, type CellSelection } from '../model/editing.js';
import {
  JSpreadsheetHost,
  releaseJSpreadsheetKeyboardOwner,
  type JSpreadsheetHostHandle,
  type JSpreadsheetSelectionRange,
  type JSpreadsheetWorksheetInstance,
  type JSpreadsheetAnchorRect,
  type JSpreadsheetCellCoords,
} from './jspreadsheet-host.js';
import { renderCalculatedCells, renderReviewMarkers } from './cell-renderer.js';

export interface WorkbookGridProps {
  worksheet: Worksheet;
  values?: CellValue[][];
  display?: DisplayOptions;
  readOnly?: boolean;
  unavailableText?: string;
  selection: CellSelection;
  reviewMode?: boolean;
  showFormulas?: boolean;
  comments?: SpreadsheetComment[];
  suggestions?: SpreadsheetSuggestion[];
  onController(controller: GridController | null): void;
  onContextMenu(coords: JSpreadsheetCellCoords, rect: JSpreadsheetAnchorRect): void;
  onCommentRequest(): void;
  onChange(worksheet: Worksheet): void;
  onSelectionChange(range: JSpreadsheetSelectionRange): void;
}

export default function WorkbookGrid(props: WorkbookGridProps) {
  const handle = useRef<JSpreadsheetHostHandle | null>(null);
  const latest = useRef(props),
    suppress = useRef(false),
    selecting = useRef(false);
  const [revision, setRevision] = useState(0);
  useLayoutEffect(() => {
    latest.current = props;
  });
  function worksheet(): JSpreadsheetWorksheetInstance | null {
    const instance = handle.current?.getInstance();
    if (!instance) return null;
    return Array.isArray(instance)
      ? (instance[0] ?? null)
      : 'worksheets' in instance
        ? (instance.worksheets?.[0] ?? null)
        : (instance as JSpreadsheetWorksheetInstance);
  }
  const controller = useMemo<GridController>(
    () => ({
      select(range) {
        if (selecting.current) return;
        selecting.current = true;
        try {
          const sheet = worksheet();
          sheet?.updateSelectionFromCoords?.(range.x1, range.y1, range.x2, range.y2);
          sheet
            ?.getCellFromCoords?.(range.x1, range.y1)
            ?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
        } finally {
          selecting.current = false;
        }
      },
      getAnchor(range) {
        return (
          worksheet()?.getCellFromCoords?.(range.x1, range.y1)?.getBoundingClientRect() ?? null
        );
      },
      releaseKeyboard() {
        releaseJSpreadsheetKeyboardOwner();
      },
    }),
    [],
  );
  useEffect(() => {
    props.onController(controller);
    return () => props.onController(null);
  }, [controller]);
  const config = useMemo(() => {
    function capture(value: unknown, allowReview = false) {
      if (
        suppress.current ||
        latest.current.readOnly ||
        (latest.current.reviewMode && !allowReview)
      )
        return;
      const worksheet = value as JSpreadsheetWorksheetInstance,
        data = worksheet?.getData?.();
      if (!data) return;
      const current = latest.current.worksheet;
      const normalized = data.map((row) =>
        row.map((cell) => (cell === undefined ? null : (cell as CellInput))),
      );
      const merge = worksheet.getMerge?.();
      const next: Worksheet = {
        ...current,
        data: normalized.map((row) => [...row]),
        ...(merge && !Array.isArray(merge) ? { mergeCells: structuredClone(merge) } : {}),
      };
      const width = Math.max(current.columns?.length ?? 0, ...normalized.map((row) => row.length));
      next.columns = Array.from({ length: width }, (_, x) => {
        const measured = Number(worksheet.getWidth?.(x));
        return {
          ...current.columns?.[x],
          ...(Number.isFinite(measured) && measured > 0 ? { width: measured } : {}),
        };
      });
      next.rows = {};
      normalized.forEach((_row, y) => {
        const height = Number(worksheet.getHeight?.(y));
        if (Number.isFinite(height) && height > 0) next.rows![String(y)] = { height };
      });
      if (JSON.stringify(next) !== JSON.stringify(current)) latest.current.onChange(next);
    }
    function propose(values: CellInput[][], x: number, y: number) {
      const current = latest.current.worksheet,
        workbook = createWorkbook();
      workbook.worksheets = [structuredClone(current)];
      latest.current.onChange(
        writeCells(workbook, current.id, Number(x), Number(y), values).worksheets[0],
      );
    }
    return {
      about: false,
      toolbar: false,
      tabs: false,
      allowExport: false,
      parseFormulas: false,
      contextMenu: () => false,
      onbeforechange: (
        _worksheet: unknown,
        _cell: unknown,
        x: number,
        y: number,
        value: CellInput,
      ) => {
        if (latest.current.readOnly)
          return latest.current.worksheet.data[Number(y)]?.[Number(x)] ?? '';
        if (!latest.current.reviewMode || suppress.current) return undefined;
        propose([[value]], x, y);
        return latest.current.worksheet.data[Number(y)]?.[Number(x)] ?? '';
      },
      onbeforepaste: (_worksheet: unknown, data: unknown, x: number, y: number) => {
        if (latest.current.readOnly) return false;
        if (!latest.current.reviewMode) return undefined;
        const values =
          typeof data === 'string'
            ? parseDelimited(data)
            : Array.isArray(data)
              ? data.map((row: unknown) =>
                  Array.isArray(row)
                    ? row.map((cell: unknown) =>
                        cell && typeof cell === 'object' && 'value' in cell
                          ? (cell.value as CellInput)
                          : (cell as CellInput),
                      )
                    : [],
                )
              : [];
        propose(values, x, y);
        return false;
      },
      worksheets: [
        {
          worksheetName: props.worksheet.name,
          data: props.worksheet.data.map((row) => [...row]),
          columns: props.worksheet.columns?.map((column) => ({ ...column, type: 'text' })),
          mergeCells: structuredClone(props.worksheet.mergeCells ?? {}),
          style: { ...props.worksheet.style },
          minDimensions: [8, 12],
          rows: structuredClone(props.worksheet.rows ?? {}),
          freezeColumns: props.worksheet.freezeColumns ?? 0,
          editable: !props.readOnly,
          // React commands move every affected formula and review anchor together.
          allowInsertRow: false,
          allowDeleteRow: false,
          allowInsertColumn: false,
          allowDeleteColumn: false,
          tableOverflow: true,
          tableWidth: '100%',
          tableHeight: '400px',
        },
      ],
      onafterchanges: (worksheet: unknown) => capture(worksheet),
      onresizecolumn: (worksheet: unknown) => capture(worksheet, true),
      onresizerow: (worksheet: unknown) => capture(worksheet, true),
      // Canonical React history includes cell data, formatting and review state.
      onundo: () => false,
      onredo: () => false,
    };
  }, []);
  useLayoutEffect(() => {
    if (!handle.current) return;
    suppress.current = true;
    try {
      handle.current.setData(props.worksheet.data.map((row) => [...row]));
    } finally {
      suppress.current = false;
    }
  }, [props.worksheet.data]);
  useLayoutEffect(() => {
    if (handle.current) controller.select(props.selection);
  }, [props.selection, controller]);
  function paint(api: JSpreadsheetHostHandle, current: WorkbookGridProps) {
    renderCalculatedCells(
      api,
      current.values,
      current.display,
      current.worksheet.data,
      current.unavailableText,
      current.worksheet.formats,
      current.showFormulas,
    );
    renderReviewMarkers(api, current.worksheet.id, current.comments, current.suggestions);
  }
  useEffect(() => {
    if (handle.current) paint(handle.current, props);
  }, [
    props.values,
    props.display,
    props.worksheet,
    props.unavailableText,
    props.comments,
    props.suggestions,
    props.showFormulas,
    revision,
  ]);
  return (
    <JSpreadsheetHost
      ref={handle}
      config={config}
      className="se-editor__engine"
      onMounted={(_instance, api) => {
        handle.current = api;
        paint(api, latest.current);
        controller.select(latest.current.selection);
      }}
      onSelectionChange={(range) => {
        if (!selecting.current) props.onSelectionChange(range);
      }}
      onContextMenuRequest={props.onContextMenu}
      onCellCommentRequest={props.onCommentRequest}
      onCellEditEnd={() => setRevision((value) => value + 1)}
    />
  );
}
