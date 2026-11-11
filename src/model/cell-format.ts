import { validateNumberFormat } from '../calculation/cell-format.js';
import type { CellFormat } from './review.js';

/** Shared by authored cells and serialized suggestion baselines. */
export function assertCellFormat(value: unknown): asserts value is CellFormat {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new TypeError('Invalid cell format.');
  const format = value as CellFormat;
  if (
    ![
      'general',
      'number',
      'percent',
      'currency',
      'scientific',
      'fraction',
      'date',
      'time',
      'datetime',
      'unit',
    ].includes(format.kind) ||
    (format.kind === 'unit' && !format.unit) ||
    [format.unit, format.inputUnit, format.expectedDimension, format.quantityKind].some(
      (entry) => entry !== undefined && (typeof entry !== 'string' || entry.length > 128),
    ) ||
    (format.decimals !== undefined &&
      (!Number.isInteger(format.decimals) || format.decimals < 0 || format.decimals > 100))
  )
    throw new TypeError('Invalid cell format.');
  validateNumberFormat(format.locale, format.currency);
}
