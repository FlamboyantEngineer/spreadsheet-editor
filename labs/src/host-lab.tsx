import { Button } from '@flamboyantengineer/ui-components/button';
import { useEffect, useMemo, useRef, useState } from 'react';
import { calculateWorkbook } from '../../src/calculation/evaluator.js';
import {
  calculationCellKey,
  type ExternalCalculationSnapshot,
  workbookCalculationKey,
} from '../../src/calculation/publication.js';
import type {
  SpreadsheetEditorHandle,
  SpreadsheetEditorState,
} from '../../src/editor/contracts.js';
import type {
  CellPresentationSnapshot,
  FormulaFunctionDescriptor,
} from '../../src/editor/public-host.js';
import {
  createMemoryHost,
  READ_ONLY_CAPABILITIES,
  type WorkbookStoreAdapter,
} from '../../src/host/index.js';
import { createWorkbook, SpreadsheetEditor, useWorkbookSession } from '../../src/index.js';
import type { Workbook } from '../../src/model/workbook.js';

const descriptors: readonly FormulaFunctionDescriptor[] = [
  {
    name: 'CAPTURED',
    signature: 'CAPTURED()',
    description: 'A captured value supplied by this example host.',
    aliases: ['SNAPSHOTVALUE'],
  },
];
function exampleWorkbook(): Workbook {
  const workbook = createWorkbook([
    ['2', '=A1*3', '=SUM(20,22)', '=CAPTURED()'],
    ['=CAPTURED()', '=A1+1', '=CONCAT("Hello ","world")', '=1 m'],
  ]);
  workbook.worksheets[0].id = 'host-main';
  workbook.worksheets[0].name = 'Inputs';
  workbook.worksheets[0].minDimensions = [26, 100];
  workbook.worksheets.push({
    id: 'host-second',
    name: 'Another sheet',
    data: [['7', '=A1*2']],
    minDimensions: [26, 100],
  });
  workbook.worksheets.push({
    id: 'host-edge',
    name: 'Anchored extent',
    data: [[]],
    minDimensions: [2, 2],
  });
  workbook.comments = [
    {
      id: 'edge-thread',
      sheetId: 'host-edge',
      range: 'D5',
      state: 'open',
      entries: [
        {
          id: 'edge-post',
          author: 'host',
          text: 'An annotation extends this otherwise blank worksheet.',
          createdAt: '2026-11-11T12:00:00Z',
        },
      ],
    },
  ];
  return workbook;
}
const host = (text: string) => ({
  callFunction(name: string) {
    return name === 'CAPTURED' || name === 'SNAPSHOTVALUE' ? text : undefined;
  },
});

/** This is an application-owned example: calculation publication and durability
 * sit outside the editor, and rich content uses React's ordinary text escaping.
 */
export function HostLab() {
  const [workbook, setWorkbook] = useState(exampleWorkbook);
  const current = useRef(workbook);
  const [handle, setHandle] = useState<SpreadsheetEditorHandle | null>(null);
  const [state, setState] = useState<SpreadsheetEditorState | null>(null);
  const [external, setExternal] = useState(false);
  const [snapshot, setSnapshot] = useState<ExternalCalculationSnapshot | null>(null);
  const [readOnly, setReadOnly] = useState(false);
  const [rejectDurability, setRejectDurability] = useState(false);
  const [rejectChanges, setRejectChanges] = useState(false);
  const [transactions, setTransactions] = useState(0);
  const [saves, setSaves] = useState(0);
  const [saved, setSaved] = useState<Workbook | null>(null);
  const [error, setError] = useState<string | null>(null);
  const revision = useRef(0);
  // The store reads the switch when it saves, outside a render.
  const rejecting = useRef(rejectDurability);
  // The application's own store: it counts saves and can refuse them, as an unreachable server would.
  const store = useMemo<WorkbookStoreAdapter>(() => {
    const memory = createMemoryHost({ workbooks: { host: exampleWorkbook() } }).workbooks;
    return {
      load: (id, signal) => memory.load(id, signal),
      async save(id, save) {
        if (rejecting.current)
          return {
            ok: false,
            error: { kind: 'unavailable', message: 'Host durability receipt was rejected.' },
          };
        const result = await memory.save(id, save);
        if (result.ok) {
          setSaved(structuredClone(save.content));
          setSaves((count) => count + 1);
        }
        return result;
      },
    };
  }, []);
  const session = useWorkbookSession({
    workbooks: store,
    workbookId: 'host',
    autosaveDelay: Infinity,
  });
  const [captured, setCaptured] = useState('Captured source');
  const client = useMemo(() => ({ mode: 'client' as const, host: host('Captured source') }), []);
  const calculation = useMemo(
    () => (external ? { mode: 'external' as const, snapshot } : client),
    [external, snapshot, client],
  );
  useEffect(() => {
    if (!handle) return;
    const read = () => setState(handle.getState());
    read();
    return handle.subscribe(read);
  }, [handle]);
  const presentation = useMemo<CellPresentationSnapshot | undefined>(
    () =>
      snapshot?.status === 'ready'
        ? {
            inputKey: snapshot.inputKey,
            revision: snapshot.revision,
            cells: {
              [calculationCellKey('host-main', 3, 0)]: {
                text: captured,
                tooltip: `Captured source revision ${snapshot.revision}`,
                ariaDescription: 'Host supplied captured value',
                status: 'resolved',
                decoration: 'external',
              },
              [calculationCellKey('host-main', 3, 1)]: { defaultUnit: 'cm', status: 'resolved' },
            },
          }
        : undefined,
    [snapshot, captured],
  );
  function publish(text = captured) {
    const accepted = handle?.getSnapshot() ?? current.current;
    setCaptured(text);
    setSnapshot({
      inputKey: workbookCalculationKey(accepted),
      revision: ++revision.current,
      status: 'ready',
      result: calculateWorkbook(accepted, {}, host(text)),
      hostContext: { example: 'immutable captured source', revision: String(revision.current) },
    });
  }
  async function attempt(action: () => void | Promise<void>) {
    setError(null);
    try {
      await action();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }
  return (
    <>
      <section className="lab-demo" aria-label="Public host integration example">
        <SpreadsheetEditor
          aria-label="Host spreadsheet"
          editorRef={setHandle}
          value={workbook}
          acceptChange={(next) => {
            if (rejectChanges) throw new Error('Host workbook change was rejected.');
            current.current = next;
            setWorkbook(next);
            setTransactions((count) => count + 1);
            if (external) setSnapshot(null);
            return next;
          }}
          calculation={calculation}
          formulaFunctions={descriptors}
          cellPresentation={external ? presentation : undefined}
          renderCellContent={(context) =>
            external && context.cellKey === calculationCellKey('host-main', 3, 0) ? (
              <strong data-testid="captured-rich-content">
                {context.presentation?.text ?? context.calculatedText}
              </strong>
            ) : undefined
          }
          capabilities={readOnly ? READ_ONLY_CAPABILITIES : undefined}
          features={{ comments: false, suggestions: false, panels: false, modeControl: false }}
          session={session}
        />
      </section>
      <section className="lab-controls" aria-label="Example host controls">
        <h2>Application-owned publication and durability</h2>
        <p>
          A1 feeds B1 and B2; C1 is independent. CAPTURED is advertised by the host on both editing
          surfaces. Authoritative mode changes only when this example host publishes a result.
        </p>
        <Button
          onPress={() =>
            void attempt(async () => {
              await handle?.flush();
              if (external) {
                setExternal(false);
                setSnapshot(null);
              } else {
                publish();
                setExternal(true);
              }
            })
          }
        >
          {external ? 'Use client calculations' : 'Use authoritative results'}
        </Button>
        <Button isDisabled={!external} onPress={() => publish('Refreshed source')}>
          Publish updated result
        </Button>
        <Button
          isDisabled={!external}
          onPress={() => {
            setSnapshot({
              inputKey: workbookCalculationKey(current.current),
              revision: ++revision.current,
              status: 'pending',
            });
          }}
        >
          Publish pending result
        </Button>
        <Button
          isDisabled={!external}
          onPress={() => publish('<img src=x onerror=alert(1)> escaped source')}
        >
          Publish escaped text
        </Button>
        <Button onPress={() => setReadOnly((value) => !value)}>
          {readOnly ? 'Enable editing' : 'Enable read-only'}
        </Button>
        <Button onPress={() => void attempt(() => handle?.save())}>Save host editor</Button>
        <Button onPress={() => void attempt(() => handle?.flush())}>Flush host editor</Button>
        <Button
          onPress={() => {
            rejecting.current = !rejectDurability;
            setRejectDurability(!rejectDurability);
          }}
        >
          {rejectDurability ? 'Allow saves' : 'Reject saves'}
        </Button>
        <Button onPress={() => setRejectChanges((value) => !value)}>
          {rejectChanges ? 'Allow workbook changes' : 'Reject workbook changes'}
        </Button>
        <dl>
          <dt>Accepted A1</dt>
          <dd>
            <output aria-label="Authored A1">{String(workbook.worksheets[0].data[0][0])}</output>
          </dd>
          <dt>Saved A1</dt>
          <dd>
            <output aria-label="Saved A1">
              {saved ? String(saved.worksheets[0].data[0][0]) : 'No save'}
            </output>
          </dd>
          <dt>Draft</dt>
          <dd>
            <output aria-label="Host draft state">{state?.draftDirty ? 'Editing' : 'Clean'}</output>
          </dd>
          <dt>Unsaved changes</dt>
          <dd>
            <output aria-label="Host dirty state">{state?.dirty ? 'Unsaved' : 'Saved'}</output>
          </dd>
          <dt>Transactions</dt>
          <dd>
            <output aria-label="Host transaction count">{transactions}</output>
          </dd>
          <dt>Saves</dt>
          <dd>
            <output aria-label="Host save count">{saves}</output>
          </dd>
        </dl>
        {error && <p role="alert">{error}</p>}
      </section>
    </>
  );
}
