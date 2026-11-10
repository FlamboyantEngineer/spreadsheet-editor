'use client';

import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CellValue } from '../calculation/values.js';
import type { DisplayOptions } from '../calculation/format.js';
import type { CellInput, Worksheet } from '../model/workbook.js';
import {
  JSpreadsheetHost,
  type JSpreadsheetHostHandle,
  type JSpreadsheetSelectionRange,
  type JSpreadsheetWorksheetInstance,
} from './jspreadsheet-host.js';
import { renderCalculatedCells } from './cell-renderer.js';

export interface WorkbookGridProps {
  worksheet: Worksheet;
  values?: CellValue[][];
  display?: DisplayOptions;
  readOnly?: boolean;
  unavailableText?: string;
  onChange(worksheet: Worksheet): void;
  onSelectionChange(range: JSpreadsheetSelectionRange): void;
}

export default function WorkbookGrid(props: WorkbookGridProps) {
  const handle = useRef<JSpreadsheetHostHandle | null>(null);
  const latest = useRef(props),
    suppress = useRef(false);
  const [revision, setRevision] = useState(0);
  useLayoutEffect(() => {
    latest.current = props;
  });
  const config = useMemo(() => {
    function capture(value: unknown) {
      if (suppress.current || latest.current.readOnly) return;
      const worksheet = value as JSpreadsheetWorksheetInstance;
      const data = worksheet?.getData?.();
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
      const rows: Record<string, { height: number }> = {};
      normalized.forEach((_row, y) => {
        const height = Number(worksheet.getHeight?.(y));
        if (Number.isFinite(height) && height > 0) rows[String(y)] = { height };
      });
      next.rows = rows;
      if (JSON.stringify(next) !== JSON.stringify(current)) latest.current.onChange(next);
    }
    return {
      about: false,
      toolbar: false,
      tabs: false,
      allowExport: false,
      parseFormulas: false,
      worksheets: [
        {
          worksheetName: props.worksheet.name,
          data: props.worksheet.data.map((row) => [...row]),
          columns: props.worksheet.columns?.map((column) => ({ ...column, type: 'text' })),
          mergeCells: structuredClone(props.worksheet.mergeCells ?? {}),
          style: { ...props.worksheet.style },
          minDimensions: [8, 12],
          rows: structuredClone(props.worksheet.rows ?? {}),
          editable: !props.readOnly,
          // Structure changes require reference rewriting; expose them after that contract is migrated.
          allowInsertRow: false,
          allowDeleteRow: false,
          allowInsertColumn: false,
          allowDeleteColumn: false,
          tableOverflow: true,
          tableWidth: '100%',
          tableHeight: '360px',
        },
      ],
      onafterchanges: capture,
      onmerge: capture,
      onunmerge: capture,
      onresizecolumn: capture,
      onresizerow: capture,
      onundo: capture,
      onredo: capture,
    };
    // A keyed surface owns one engine configuration for its lifetime.
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
  useEffect(() => {
    if (handle.current)
      renderCalculatedCells(
        handle.current,
        props.values,
        props.display,
        props.worksheet.data,
        props.unavailableText,
      );
  }, [props.values, props.display, props.worksheet.data, props.unavailableText, revision]);
  return (
    <JSpreadsheetHost
      ref={handle}
      config={config}
      className="se-editor__engine"
      onMounted={(_instance, api) => {
        handle.current = api;
        renderCalculatedCells(
          api,
          latest.current.values,
          latest.current.display,
          latest.current.worksheet.data,
          latest.current.unavailableText,
        );
      }}
      onSelectionChange={props.onSelectionChange}
      onCellEditEnd={() => setRevision((value) => value + 1)}
    />
  );
}
