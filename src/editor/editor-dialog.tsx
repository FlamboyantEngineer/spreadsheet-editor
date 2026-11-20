'use client';

import { Dialog } from '@flamboyantengineer/ui-components/dialog';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import type { EditorDialog } from './editor-core.js';
import type { SpreadsheetEditorController } from './use-editor.js';

const DIALOG_TITLES: Record<Exclude<EditorDialog, null>, string> = {
  'confirm-structure': 'Change merged cells',
  'row-height': 'Row height',
  'column-width': 'Column width',
  import: 'Import workbook',
  export: 'Export workbook',
};

const primaryLabel = (dialog: EditorDialog) =>
  dialog === 'confirm-structure'
    ? 'Continue'
    : dialog === 'export'
      ? 'Done'
      : dialog === 'import'
        ? 'Import'
        : 'Apply';

const fieldLabel = (dialog: EditorDialog) =>
  dialog === 'import'
    ? 'CSV or workbook JSON'
    : dialog === 'export'
      ? 'Workbook JSON'
      : 'Size in pixels';

/** The editor's dialog: line sizes, import and export, and merged-cell confirmation. */
export function EditorDialogModal({ editor: e }: { editor: SpreadsheetEditorController }) {
  const { dialog } = e;
  return (
    <Dialog
      isOpen={dialog !== null}
      onOpenChange={(open) => {
        if (!open) e.setDialog(null);
      }}
      title={dialog === null ? 'Export workbook' : DIALOG_TITLES[dialog]}
      primaryAction={{
        label: primaryLabel(dialog),
        onPress: () => e.applyDialog(),
        isDisabled: !!e.dialogError,
      }}
      cancelAction={dialog === 'export' ? null : {}}
      description={
        dialog === 'import'
          ? 'Paste CSV data or a serialized workbook. All digits in CSV numeric fields are preserved.'
          : undefined
      }
    >
      <div onFocusCapture={() => e.grid.current?.releaseKeyboard()}>
        {dialog === 'confirm-structure' ? (
          <p>
            This change intersects merged cells. Continue with the row or column change and adjusted
            merge spans?
          </p>
        ) : (
          <TextField
            autoFocus
            label={fieldLabel(dialog)}
            tracksChanges={false}
            value={e.dialogValue}
            onChange={e.setDialogValue}
            multiline={dialog === 'import' || dialog === 'export'}
            isReadOnly={dialog === 'export'}
            minRows={8}
            maxRows={18}
          />
        )}
        {e.dialogError && <p role="alert">{e.dialogError}</p>}
      </div>
    </Dialog>
  );
}
