// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { FULL_CAPABILITIES } from '../host/capabilities.js';
import { createMemoryHost } from '../host/memory-host.js';
import { createWorkbookSession } from '../host/workbook-session.js';
import { writeCells } from '../model/editing.js';
import { createWorkbook } from '../model/workbook.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { useSpreadsheetEditor } from './use-editor.js';

async function openVersions(capabilities = FULL_CAPABILITIES) {
  const host = createMemoryHost({
    reader: { id: 'alex', name: 'Alex' },
    workbooks: { book: createWorkbook([['first']]) },
  });
  const session = createWorkbookSession({
    workbooks: host.workbooks,
    workbookId: 'book',
    autosaveDelay: Infinity,
  });
  await session.ready;
  const props: SpreadsheetEditorProps = {
    session,
    versions: host.versions,
    workbookId: 'book',
    identity: host.identity,
    capabilities,
  };
  const rendered = renderHook((current: SpreadsheetEditorProps) => useSpreadsheetEditor(current), {
    initialProps: props,
  });
  const edit = (value: string) =>
    act(() => {
      rendered.result.current.mutate(
        writeCells(rendered.result.current.workbook, [[value]], { sheetId: 'sheet-1', x: 0, y: 0 }),
        'Edit',
      );
    });
  return { host, session, props, edit, ...rendered };
}

describe('SE-ADAPTERS-5.1: saved versions through the version store', () => {
  it('lists versions while the history pane is open and follows the store', async () => {
    const { result, edit } = await openVersions();
    expect(result.current.versionList.versions).toEqual([]);
    act(() => {
      result.current.openPanel('history');
    });
    await waitFor(() => expect(result.current.versionList.versions).toHaveLength(1));
    expect(result.current.versionList.versions[0]).toMatchObject({ revision: 'r1', label: null });
    edit('second');
    await act(() => result.current.save());
    await waitFor(() =>
      expect(result.current.versionList.versions.map((version) => version.revision)).toEqual([
        'r2',
        'r1',
      ]),
    );
  });

  it('opens a version read-only in place of the workbook and returns to the current one', async () => {
    const { result, edit, session } = await openVersions();
    edit('second');
    act(() => {
      result.current.openPanel('history');
    });
    await waitFor(() => expect(result.current.versionList.versions).toHaveLength(1));
    const first = result.current.versionList.versions[0];
    await act(() => result.current.openVersion(first));
    // Opening a version saves the change made before it.
    expect(session.status()).toMatchObject({ dirty: false, revision: 'r2' });
    expect(result.current.viewedVersion).toEqual(first);
    expect(result.current.sheet.data[0][0]).toBe('first');
    expect(result.current.mode).toBe('read');
    expect(result.current.canEdit).toBe(false);
    expect(result.current.canSaveVersion).toBe(false);
    edit('refused');
    expect(result.current.sheet.data[0][0]).toBe('first');
    expect(session.status().dirty).toBe(false);
    act(() => {
      result.current.returnToCurrent();
    });
    expect(result.current.viewedVersion).toBeNull();
    expect(result.current.sheet.data[0][0]).toBe('second');
    expect(result.current.mode).toBe('author');
  });

  it('saves a named version of the saved workbook, renames it and restores another', async () => {
    const { result, edit, session, host } = await openVersions();
    act(() => {
      result.current.openPanel('history');
    });
    edit('second');
    await act(() => result.current.saveVersion('Review copy'));
    await waitFor(() =>
      expect(result.current.versionList.versions.map((version) => version.label)).toEqual([
        'Review copy',
        null,
        null,
      ]),
    );
    const [named, , seed] = result.current.versionList.versions;
    expect(named.author).toBe('alex');
    await act(() => result.current.renameVersion(named, 'Final'));
    await waitFor(() => expect(result.current.versionList.versions[0].label).toBe('Final'));
    await act(() => result.current.openVersion(seed));
    await act(() => result.current.restoreVersion(seed));
    expect(result.current.viewedVersion).toBeNull();
    expect(result.current.sheet.data[0][0]).toBe('first');
    expect(session.status()).toMatchObject({ dirty: false });
    const stored = await host.workbooks.load('book');
    expect(stored.ok && stored.value.content.worksheets[0].data[0][0]).toBe('first');
  });

  it('offers only what the capabilities and the store allow, and shows a failed load', async () => {
    const { result, rerender, props, host } = await openVersions({
      ...FULL_CAPABILITIES,
      saveVersion: false,
      edit: false,
    });
    expect(result.current.canSaveVersion).toBe(false);
    expect(result.current.canRenameVersion).toBe(false);
    expect(result.current.canRestoreVersion).toBe(false);
    rerender({ ...props, versions: { list: host.versions.list, load: host.versions.load } });
    expect(result.current.canRenameVersion).toBe(false);
    rerender({ ...props, capabilities: { ...FULL_CAPABILITIES, history: false } });
    act(() => {
      result.current.openPanel('history');
    });
    await Promise.resolve();
    expect(result.current.versionList.versions).toEqual([]);
    rerender(props);
    await waitFor(() => expect(result.current.versionList.versions).toHaveLength(1));
    await act(() =>
      result.current.openVersion({ ...result.current.versionList.versions[0], revision: 'gone' }),
    );
    expect(result.current.viewedVersion).toBeNull();
    expect(result.current.notice).toMatch(/gone/);
  });
});
