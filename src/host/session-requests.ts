import type { HostError, HostResult } from './adapters.js';

// Timing and failures shared by a workbook session's loads and saves.
export const DEFAULT_AUTOSAVE_DELAY = 1000;
export const DEFAULT_DRAFT_DELAY = 300;
/** Autosaves retry an unavailable store, with backoff, for this long. */
export const RETRY_WINDOW = 30_000;

export type Timer = ReturnType<typeof setTimeout>;

function failure(error: unknown): HostError {
  return { kind: 'unavailable', message: error instanceof Error ? error.message : undefined };
}

export class HostRequestError extends Error {
  constructor(readonly error: HostError) {
    super(error.message ?? `The workbook could not be saved (${error.kind}).`);
  }
}

export const closedError = () => new Error('The workbook session has closed.');
export const wait = (milliseconds: number) =>
  new Promise((resolve) => setTimeout(resolve, milliseconds));

/** A host request's answer; a request that throws answers that the store is unavailable. */
export async function answer<T>(request: () => Promise<HostResult<T>>): Promise<HostResult<T>> {
  try {
    return await request();
  } catch (error) {
    return { ok: false, error: failure(error) };
  }
}
