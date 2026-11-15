'use client';

import { type RefObject, useRef } from 'react';
import type { HostResult, WorkbookVersion } from '../host/adapters.js';
import type { Workbook } from '../model/workbook.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import type { WorkbookVersionList } from './use-workbook-versions.js';

type ViewedVersion = { version: WorkbookVersion; content: Workbook } | null;

/** Opening, saving, renaming and restoring saved versions (SE-ADAPTERS-5). */
export function useVersionActions({
  viewedVersion,
  setViewedVersion,
  versionList,
  commitDraft,
  flush,
  props,
  alive,
  setNotice,
}: {
  viewedVersion: ViewedVersion;
  setViewedVersion: (next: ViewedVersion | ((viewed: ViewedVersion) => ViewedVersion)) => void;
  versionList: WorkbookVersionList;
  commitDraft: () => Workbook;
  flush: () => Promise<void>;
  props: () => SpreadsheetEditorProps;
  alive: RefObject<boolean>;
  setNotice: (message: string | null) => void;
}) {
  /**
   * Commits the active cell edit and saves every change made so far, so a version action never
   * leaves an edit behind. Throws when an edit stays unsaved.
   */
  async function settleChanges() {
    commitDraft();
    const current = props().session;
    if (current?.status().dirty || current?.status().phase === 'saving') {
      await flush();
      if (current.status().dirty)
        throw new Error('Changes are not saved. Save or discard them first.');
    }
  }
  /** Runs a version action, showing its failure as a notice. Resolves whether it succeeded. */
  async function versionAction(action: () => Promise<void>, fallback: string) {
    try {
      await action();
      return true;
    } catch (error) {
      if (alive.current) setNotice(error instanceof Error ? error.message : fallback);
      return false;
    }
  }
  function versionStore() {
    const { versions, workbookId } = props();
    if (!versions || workbookId === undefined) throw new Error('This workbook has no versions.');
    return { versions, workbookId };
  }
  function unwrap<T>(result: HostResult<T>, fallback: string): T {
    if (!result.ok) throw new Error(result.error.message ?? fallback);
    return result.value;
  }
  const versionRequest = useRef(0);
  /** Saves pending changes, then shows `version` read-only in place of the workbook. */
  function openVersion(version: WorkbookVersion) {
    const request = ++versionRequest.current;
    return versionAction(async () => {
      if (!viewedVersion) await settleChanges();
      const { versions, workbookId } = versionStore();
      const content = unwrap(
        await versions.load(workbookId, version.revision),
        'Could not open the version.',
      );
      if (alive.current && request === versionRequest.current) {
        setNotice(null);
        setViewedVersion({ version, content });
      }
    }, 'Could not open the version.');
  }
  function returnToCurrent() {
    versionRequest.current++;
    setViewedVersion(null);
  }
  /** Saves pending changes, then saves the workbook as a new version named `label`. */
  function saveVersion(label: string | null) {
    return versionAction(async () => {
      await settleChanges();
      const { versions, workbookId } = versionStore();
      if (!versions.create) throw new Error('This workbook cannot save versions.');
      unwrap(
        await versions.create(workbookId, { content: commitDraft(), label }),
        'Could not save the version.',
      );
      versionList.refresh();
    }, 'Could not save the version.');
  }
  function renameVersion(version: WorkbookVersion, label: string | null) {
    return versionAction(async () => {
      const { versions, workbookId } = versionStore();
      if (!versions.rename) throw new Error('This workbook cannot rename versions.');
      const renamed = unwrap(
        await versions.rename(workbookId, version.revision, label),
        'Could not rename the version.',
      );
      if (alive.current)
        setViewedVersion((viewed) =>
          viewed?.version.revision === renamed.revision ? { ...viewed, version: renamed } : viewed,
        );
      versionList.refresh();
    }, 'Could not rename the version.');
  }
  /** Makes `version` the workbook's current content and returns to it. */
  function restoreVersion(version: WorkbookVersion) {
    return versionAction(async () => {
      const { versions, workbookId } = versionStore();
      if (!versions.restore) throw new Error('This workbook cannot restore versions.');
      if (!viewedVersion) await settleChanges();
      unwrap(
        await versions.restore(workbookId, version.revision),
        'Could not restore the version.',
      );
      await props().session?.takeStored();
      if (alive.current) returnToCurrent();
      versionList.refresh();
    }, 'Could not restore the version.');
  }
  return { openVersion, returnToCurrent, saveVersion, renameVersion, restoreVersion };
}
