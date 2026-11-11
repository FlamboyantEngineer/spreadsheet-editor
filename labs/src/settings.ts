import { createSafeStorage, type UiStorage } from '@flamboyantengineer/ui-components/core';

/** Lab preferences are optional and separate from editor/document persistence. */
export function labStorage(): UiStorage | null {
  try {
    const storage = window.localStorage;
    const prefix = 'spreadsheet-editor-labs:';
    return createSafeStorage({
      getItem: (key) => storage.getItem(prefix + key),
      setItem: (key, value) => storage.setItem(prefix + key, value),
      removeItem: (key) => storage.removeItem(prefix + key),
    });
  } catch {
    return null;
  }
}
