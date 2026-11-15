// @vitest-environment jsdom
import { jspreadsheet } from '@jspreadsheet-ce/react';
import { afterEach, expect, it } from 'vitest';
import {
  type JSpreadsheetWorksheetInstance,
  releaseJSpreadsheetKeyboardOwner,
} from './jspreadsheet-host.js';

const vendor = jspreadsheet as unknown as { current: JSpreadsheetWorksheetInstance | null };
afterEach(() => {
  vendor.current = null;
  document.body.replaceChildren();
});

it('SE-HOST-1: late grid teardown cannot release a newer keyboard owner', () => {
  const oldElement = document.createElement('div'),
    nextElement = document.createElement('div');
  const old = { element: oldElement } as unknown as JSpreadsheetWorksheetInstance;
  const next = { element: nextElement } as unknown as JSpreadsheetWorksheetInstance;
  document.body.append(oldElement, nextElement);
  vendor.current = next;
  expect(releaseJSpreadsheetKeyboardOwner(old)).toBeNull();
  expect(vendor.current).toBe(next);
  const restore = releaseJSpreadsheetKeyboardOwner(next);
  expect(vendor.current).toBeNull();
  restore?.();
  expect(vendor.current).toBe(next);
});
