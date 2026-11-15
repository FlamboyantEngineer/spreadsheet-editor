import { describe, expect, it } from 'vitest';
import {
  FULL_CAPABILITIES,
  READ_ONLY_CAPABILITIES,
  resolveSpreadsheetMode,
} from './capabilities.js';

describe('SE-ADAPTERS-2: capabilities', () => {
  it('SE-ADAPTERS-2.1: maps capabilities to the most the reader may do', () => {
    expect(resolveSpreadsheetMode(FULL_CAPABILITIES)).toBe('author');
    expect(resolveSpreadsheetMode(FULL_CAPABILITIES, { historical: true })).toBe('read');
    expect(resolveSpreadsheetMode({ ...READ_ONLY_CAPABILITIES, suggest: true })).toBe('review');
    expect(resolveSpreadsheetMode({ ...READ_ONLY_CAPABILITIES, comment: true })).toBe('review');
    expect(resolveSpreadsheetMode({ ...READ_ONLY_CAPABILITIES, decide: true })).toBe('read');
    expect(resolveSpreadsheetMode(READ_ONLY_CAPABILITIES)).toBe('read');
  });

  it('SE-ADAPTERS-2.1: the shared capability sets cannot be changed', () => {
    expect(Object.isFrozen(FULL_CAPABILITIES)).toBe(true);
    expect(Object.isFrozen(READ_ONLY_CAPABILITIES)).toBe(true);
  });
});
