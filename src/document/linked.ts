import type { Workbook } from '../model/workbook.js';
import { copyWorkbook } from '../model/workbook.js';
import {
  assertExternalCalculationSnapshot,
  orderedJson,
  type ExternalCalculationSnapshot,
} from '../calculation/publication.js';
import {
  createDocumentInsertion,
  type DocumentTableInsertion,
  type InsertionOptions,
} from './insertion.js';

/** Opaque source identity and host-authoritative revision, independent of storage/transport. */
export interface SpreadsheetSourceReference {
  id: string;
  revision: string;
}
export interface LinkedSpreadsheetInsertion {
  kind: 'linked-spreadsheet';
  schemaVersion: 1;
  source: SpreadsheetSourceReference;
  sheetId: string;
  range: string;
  workbook: Workbook;
  publication: ExternalCalculationSnapshot;
  /** Deterministic read-only/print rendering captured with this exact source revision. */
  table: DocumentTableInsertion;
}

/** Never invokes a calculator. Missing/stale/non-ready authoritative values reject capture. */
export function createLinkedSpreadsheetInsertion(
  workbook: Workbook,
  publication: ExternalCalculationSnapshot,
  options: Omit<InsertionOptions, 'mode'> & { source: SpreadsheetSourceReference },
): LinkedSpreadsheetInsertion {
  if (
    typeof options.source.id !== 'string' ||
    !options.source.id ||
    typeof options.source.revision !== 'string' ||
    !options.source.revision
  )
    throw new TypeError('A linked spreadsheet requires a source identity and revision.');
  assertExternalCalculationSnapshot(publication, workbook);
  if (
    publication.status !== 'ready' ||
    !publication.result ||
    Object.values(publication.cellStates ?? {}).some((state) => state !== 'ready')
  )
    throw new Error('A linked spreadsheet requires ready authoritative results.');
  const table = createDocumentInsertion(workbook, publication.result, {
    ...options,
    mode: 'table',
  }) as DocumentTableInsertion;
  return {
    kind: 'linked-spreadsheet',
    schemaVersion: 1,
    source: structuredClone(options.source),
    sheetId: table.sheetId,
    range: table.range,
    workbook: copyWorkbook(workbook),
    publication: structuredClone(publication),
    table,
  };
}

/** The host resolves the source and explicitly authorizes refresh of a current document. */
export function updateLinkedSpreadsheetInsertion(
  current: LinkedSpreadsheetInsertion,
  workbook: Workbook,
  publication: ExternalCalculationSnapshot,
  options: {
    source: SpreadsheetSourceReference;
    expectedRevision: string;
    display?: InsertionOptions['display'];
  },
): LinkedSpreadsheetInsertion {
  if (
    options.source.id !== current.source.id ||
    options.expectedRevision !== current.source.revision
  )
    throw new Error('The linked spreadsheet source or captured revision changed.');
  if (
    options.source.revision === current.source.revision &&
    (orderedJson(workbook) !== orderedJson(current.workbook) ||
      orderedJson(publication) !== orderedJson(current.publication))
  )
    throw new Error('A source revision cannot identify different captured data.');
  return createLinkedSpreadsheetInsertion(workbook, publication, {
    ...options,
    sheetId: current.sheetId,
    range: current.range,
  });
}
