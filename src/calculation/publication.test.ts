import { describe, expect, it } from 'vitest';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook, type Workbook } from '../model/workbook.js';
import { calculateWorkbook, type CalculationResult } from './evaluator.js';
import {
  assertExternalCalculationSnapshot,
  calculationCellKey,
  type ExternalCalculationSnapshot,
  orderedJson,
  pendingCalculationCells,
  workbookCalculationKey,
} from './publication.js';
import type { NumericValue } from './values.js';

const exact = (numerator = '1'): NumericValue => ({
  kind: 'number',
  numerator,
  denominator: '1',
  approximate: false,
});
const workbookFixture = (): Workbook => ({
  schemaVersion: 1,
  name: 'Document title',
  worksheets: [
    {
      id: 'inputs',
      name: 'Inputs',
      data: [
        ['1', '=A1+1'],
        ['other', ''],
      ],
      minDimensions: [2, 2],
    },
    { id: 'other', name: 'Other', data: [['2']] },
  ],
});
const publication = (workbook: Workbook): ExternalCalculationSnapshot => ({
  inputKey: workbookCalculationKey(workbook),
  revision: 0,
  status: 'ready',
  result: calculateWorkbook(workbook),
});
const mutate = (workbook: Workbook, change: (next: Workbook) => void) => {
  const next = structuredClone(workbook);
  change(next);
  return next;
};

describe('SE-HOST-CALC-1: deterministic authored calculation identity', () => {
  it('uses ordered JSON independent of object insertion order and omits undefined fields', () => {
    expect(orderedJson({ z: [1, { b: 2, a: null }], a: 'x', missing: undefined })).toBe(
      orderedJson({ a: 'x', z: [1, { a: null, b: 2 }] }),
    );
    const a = workbookFixture(),
      b = workbookFixture();
    a.worksheets[0].inputValues = { A1: { ...exact(), provenance: { z: '2', a: '1' } } };
    b.worksheets[0].inputValues = { A1: { ...exact(), provenance: { a: '1', z: '2' } } };
    expect(workbookCalculationKey(a)).toBe(workbookCalculationKey(b));
  });

  it('SE-3: is a short digest, whatever the size of the workbook', () => {
    const large = createWorkbook(
      Array.from({ length: 5_000 }, (_, y) => Array.from({ length: 10 }, (_, x) => `${x + y}`)),
    );
    const key = workbookCalculationKey(large);
    expect(key).toMatch(/^wck1-[0-9a-f]{28}$/);
    expect(workbookCalculationKey(structuredClone(large))).toBe(key);
    const changed = structuredClone(large);
    changed.worksheets[0].data[4_999][9] = '0';
    expect(workbookCalculationKey(changed)).not.toBe(key);
  });

  it('excludes title, contributor provenance and annotation contents from freshness', () => {
    const workbook = workbookFixture(),
      key = workbookCalculationKey(workbook);
    const next = mutate(workbook, (value) => {
      value.name = 'Renamed document';
      value.worksheets[0].cellProvenance = {
        A1: { actor: { id: 'peer', name: 'Peer' }, source: 'human' },
      };
      value.worksheets[0].style = { B2: 'font-weight:bold;background-color:#ffffff' };
      value.comments = [
        {
          id: 'thread',
          sheetId: 'inputs',
          range: 'A1',
          state: 'open',
          entries: [
            {
              id: 'post',
              author: 'peer',
              text: 'Comment',
              createdAt: '2026-11-11T12:00:00Z',
            },
          ],
        },
      ];
      value.suggestions = [
        {
          id: 'suggestion',
          author: { id: 'peer', name: 'Peer' },
          createdAt: '2026-11-11T12:00:00Z',
          status: 'pending',
          summary: 'Proposal',
          changes: [{ sheetId: 'inputs', address: 'A1', before: '1', after: '3' }],
        },
      ];
    });
    expect(workbookCalculationKey(next)).toBe(key);
    expect(
      workbookCalculationKey(
        mutate(next, (value) => {
          requireDefined(value.comments)[0].entries[0].text = 'Edited comment';
          requireDefined(value.suggestions)[0].changes[0].after = '4';
        }),
      ),
    ).toBe(key);
  });

  it('excludes purely display-format changes while retaining calculation-relevant unit fields', () => {
    const workbook = workbookFixture();
    workbook.worksheets[0].formats = {
      A1: { kind: 'unit', unit: 'm', inputUnit: 'mm', expectedDimension: 'Length' },
    };
    const changedDisplay = mutate(workbook, (value) => {
      requireDefined(value.worksheets[0].formats).A1 = {
        ...requireDefined(value.worksheets[0].formats).A1,
        decimals: 6,
        locale: 'de-DE',
        currency: 'EUR',
      };
    });
    expect(workbookCalculationKey(changedDisplay)).toBe(workbookCalculationKey(workbook));
    for (const [field, input] of [
      ['unit', 's'],
      ['inputUnit', 'cm'],
      ['expectedDimension', 'Time'],
      ['quantityKind', 'torque'],
    ] as const) {
      const changed = mutate(workbook, (value) => {
        requireDefined(value.worksheets[0].formats).A1 = {
          ...requireDefined(value.worksheets[0].formats).A1,
          [field]: input,
        };
      });
      expect(workbookCalculationKey(changed)).not.toBe(workbookCalculationKey(workbook));
    }
  });

  it('ignores distant visual/annotation anchors without changing calculation extent', () => {
    const workbook = workbookFixture(),
      key = workbookCalculationKey(workbook);
    const next = mutate(workbook, (value) => {
      value.worksheets[0].style = { Z100: 'font-weight:bold' };
      value.worksheets[0].cellProvenance = {
        AA101: { actor: { id: 'peer', name: 'Peer' }, source: 'human' },
      };
      value.worksheets[0].formats = {
        AB102: { kind: 'currency', locale: 'de-DE', currency: 'EUR' },
      };
      value.comments = [
        {
          id: 'thread',
          sheetId: 'inputs',
          range: 'AC103',
          state: 'open',
          entries: [
            {
              id: 'post',
              author: 'peer',
              text: 'Far comment',
              createdAt: '2026-11-11T12:00:00Z',
            },
          ],
        },
      ];
      value.suggestions = [
        {
          id: 'suggestion',
          author: { id: 'peer', name: 'Peer' },
          createdAt: '2026-11-11T12:00:00Z',
          status: 'pending',
          summary: 'Far proposal',
          changes: [{ sheetId: 'inputs', address: 'AD104', before: '', after: '3' }],
        },
      ];
    });
    expect(workbookCalculationKey(next)).toBe(key);
    const geometry = mutate(workbook, (value) => {
      value.worksheets[0].columns = Array.from({ length: 3 }, () => ({}));
    });
    expect(workbookCalculationKey(geometry)).not.toBe(key);
    const merged = mutate(workbook, (value) => {
      value.worksheets[0].mergeCells = { C3: [2, 2] };
    });
    expect(workbookCalculationKey(merged)).not.toBe(key);
  });

  it('includes worksheet identity, order, name, authored data, extent and stable line identities', () => {
    const workbook = workbookFixture(),
      key = workbookCalculationKey(workbook);
    const changes: ((value: Workbook) => void)[] = [
      (value) => {
        value.worksheets[0].id = 'new-id';
      },
      (value) => {
        value.worksheets.reverse();
      },
      (value) => {
        value.worksheets[0].name = 'Renamed sheet';
      },
      (value) => {
        value.worksheets[0].data[0][0] = '1E-25';
      },
      (value) => {
        value.worksheets[0].minDimensions = [3, 2];
      },
      (value) => {
        value.worksheets[0].rowIds = ['r0', 'r1'];
      },
      (value) => {
        value.worksheets[0].columnIds = ['c0', 'c1'];
      },
    ];
    changes.forEach((change) =>
      expect(workbookCalculationKey(mutate(workbook, change))).not.toBe(key),
    );
  });

  it('includes canonical value, temperature, approximation and numeric provenance', () => {
    const workbook = workbookFixture();
    workbook.worksheets[0].inputValues = {
      A1: {
        ...exact('5'),
        dimension: [0, 0, 0, 0, 1, 0, 0],
        temperature: 'difference',
        provenance: { constants: 'rev1' },
      },
    };
    const key = workbookCalculationKey(workbook);
    for (const metadata of [
      { numerator: '6' },
      { temperature: 'absolute' as const },
      { approximate: true, precision: 80 },
      { provenance: { constants: 'rev2' } },
    ]) {
      const next = mutate(workbook, (value) => {
        requireDefined(value.worksheets[0].inputValues).A1 = {
          ...requireDefined(value.worksheets[0].inputValues).A1,
          ...metadata,
        };
      });
      expect(workbookCalculationKey(next)).not.toBe(key);
    }
    const invalid = mutate(workbook, (value) => {
      requireDefined(value.worksheets[0].inputValues).B1 = exact();
    });
    expect(() => workbookCalculationKey(invalid)).toThrow('nonformula');
  });
});

describe('SE-HOST-CALC-2: external snapshot ingress', () => {
  it('accepts a detached ready publication with finite opaque context and valid quantity metadata', () => {
    const workbook = workbookFixture(),
      snapshot = publication(workbook);
    snapshot.hostContext = {
      generation: 1,
      historical: true,
      records: [null, 'captured', { revision: 'a' }],
    };
    requireDefined(snapshot.result).worksheets[0].values[0][0] = {
      ...exact('5'),
      dimension: [0, 0, 0, 0, 1, 0, 0],
      temperature: 'difference',
      provenance: { catalog: 'revision' },
    };
    expect(() => assertExternalCalculationSnapshot(snapshot, workbook)).not.toThrow();
    expect(() =>
      assertExternalCalculationSnapshot(JSON.parse(JSON.stringify(snapshot)), workbook),
    ).not.toThrow();
    expect(requireDefined(snapshot.result).worksheets[0].values[0][0]).toMatchObject({
      temperature: 'difference',
      provenance: { catalog: 'revision' },
    });
  });

  it('requires a matching authored key, nonnegative safe revision and known global state', () => {
    const workbook = workbookFixture();
    for (const revision of [
      -1,
      0.5,
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.MAX_SAFE_INTEGER + 1,
    ])
      expect(() =>
        assertExternalCalculationSnapshot({ ...publication(workbook), revision }, workbook),
      ).toThrow();
    for (const status of ['', 'resolved', null])
      expect(() =>
        assertExternalCalculationSnapshot({ ...publication(workbook), status }, workbook),
      ).toThrow();
    expect(() =>
      assertExternalCalculationSnapshot({ ...publication(workbook), inputKey: 'other' }, workbook),
    ).toThrow('input key');
    expect(() =>
      assertExternalCalculationSnapshot(
        { ...publication(workbook), revision: Number.MAX_SAFE_INTEGER },
        workbook,
      ),
    ).not.toThrow();
    for (const value of [null, undefined, false, 'ready', []])
      expect(() => assertExternalCalculationSnapshot(value, workbook)).toThrow();
  });

  it('rejects prototype-dependent snapshots and state maps', () => {
    const workbook = workbookFixture(),
      snapshot = publication(workbook);
    for (const value of [
      Object.assign([], snapshot),
      Object.assign(new Map(), snapshot),
      Object.assign(new Date(), snapshot),
    ])
      expect(() => assertExternalCalculationSnapshot(value, workbook)).toThrow();
    const cellStates = Object.create({ [calculationCellKey('inputs', 0, 0)]: 'pending' });
    expect(() =>
      assertExternalCalculationSnapshot({ ...snapshot, cellStates }, workbook),
    ).toThrow();
    expect(() =>
      assertExternalCalculationSnapshot({ ...snapshot, cellStates: Object.create(null) }, workbook),
    ).not.toThrow();
  });

  it('rejects ready without results, missing authored values and malformed worksheet/result shapes', () => {
    const workbook = workbookFixture();
    expect(() =>
      assertExternalCalculationSnapshot(
        { inputKey: workbookCalculationKey(workbook), revision: 0, status: 'ready' },
        workbook,
      ),
    ).toThrow('require a result');
    for (const change of [
      (result: CalculationResult) => {
        result.worksheets[0].values[0][0] = null;
      },
      (result: CalculationResult) => {
        result.worksheets.pop();
      },
      (result: CalculationResult) => {
        result.worksheets[0].id = 'other';
      },
      (result: CalculationResult) => {
        result.worksheets[0].values.pop();
      },
      (result: CalculationResult) => {
        result.worksheets[0].values[0].pop();
      },
    ]) {
      const snapshot = publication(workbook);
      change(requireDefined(snapshot.result));
      expect(() => assertExternalCalculationSnapshot(snapshot, workbook)).toThrow();
    }
    const blank = createWorkbook([['', null]]);
    expect(() => assertExternalCalculationSnapshot(publication(blank), blank)).not.toThrow();
  });

  it('honors per-cell pending/unavailable states and forbids presenting old cached values', () => {
    const workbook = workbookFixture(),
      snapshot = publication(workbook);
    const key = calculationCellKey('inputs', 1, 0);
    snapshot.cellStates = { [key]: 'pending' };
    expect(() => assertExternalCalculationSnapshot(snapshot, workbook)).toThrow('cached values');
    requireDefined(snapshot.result).worksheets[0].values[0][1] = null;
    expect(() => assertExternalCalculationSnapshot(snapshot, workbook)).not.toThrow();
    snapshot.cellStates[key] = 'unavailable';
    expect(() => assertExternalCalculationSnapshot(snapshot, workbook)).not.toThrow();
    snapshot.cellStates[key] = 'ready';
    expect(() => assertExternalCalculationSnapshot(snapshot, workbook)).toThrow(
      'require calculated values',
    );
  });

  it('allows pending/unavailable publications with no result and mixed explicitly ready cells', () => {
    const workbook = workbookFixture();
    for (const status of ['pending', 'unavailable'] as const) {
      const empty: ExternalCalculationSnapshot = {
        inputKey: workbookCalculationKey(workbook),
        revision: 1,
        status,
      };
      expect(() => assertExternalCalculationSnapshot(empty, workbook)).not.toThrow();
      const mixed: ExternalCalculationSnapshot = { ...empty, result: calculateWorkbook(workbook) };
      requireDefined(mixed.result).worksheets.forEach((sheet) =>
        sheet.values.forEach((row) => row.fill(null)),
      );
      requireDefined(mixed.result).worksheets[0].values[0][0] = exact();
      mixed.cellStates = { [calculationCellKey('inputs', 0, 0)]: 'ready' };
      expect(() => assertExternalCalculationSnapshot(mixed, workbook)).not.toThrow();
      delete mixed.cellStates;
      expect(() => assertExternalCalculationSnapshot(mixed, workbook)).toThrow('cached values');
    }
  });

  it('validates tuple shape, known sheet, bounded coordinates and state-map size', () => {
    const workbook = workbookFixture();
    for (const key of [
      'not-json',
      '{}',
      '["inputs",0]',
      '["missing",0,0]',
      '["inputs",-1,0]',
      '["inputs",0.5,0]',
      '["inputs",0,"0"]',
      '["inputs",16384,0]',
      '["inputs",0,1048576]',
    ])
      expect(() =>
        assertExternalCalculationSnapshot(
          { ...publication(workbook), cellStates: { [key]: 'pending' } },
          workbook,
        ),
      ).toThrow();
    for (const cellStates of [null, [], { [calculationCellKey('inputs', 0, 0)]: 'resolved' }])
      expect(() =>
        assertExternalCalculationSnapshot({ ...publication(workbook), cellStates }, workbook),
      ).toThrow();
    const cellStates = Object.fromEntries(
      Array.from({ length: 250001 }, (_, x) => [calculationCellKey('inputs', x, 0), 'pending']),
    );
    expect(() =>
      assertExternalCalculationSnapshot({ ...publication(workbook), cellStates }, workbook),
    ).toThrow('Too many');
  });

  it('rejects non-JSON, over-depth and over-byte host context plus invalid error metadata', () => {
    const workbook = workbookFixture(),
      snapshot = publication(workbook);
    const cycle: Record<string, unknown> = {};
    cycle.self = cycle;
    let deep: unknown = null;
    for (let i = 0; i < 33; i++) deep = { nested: deep };
    for (const hostContext of [
      Number.NaN,
      Number.POSITIVE_INFINITY,
      () => 1,
      { value: undefined },
      new Date(),
      new Map(),
      cycle,
      deep,
      'x'.repeat(16 * 1024 * 1024),
    ])
      expect(() =>
        assertExternalCalculationSnapshot({ ...snapshot, hostContext }, workbook),
      ).toThrow();
    let atLimit: unknown = null;
    for (let i = 0; i < 32; i++) atLimit = { nested: atLimit };
    expect(() =>
      assertExternalCalculationSnapshot({ ...snapshot, hostContext: atLimit }, workbook),
    ).not.toThrow();
    for (const error of [false, null, 'x'.repeat(8193)])
      expect(() => assertExternalCalculationSnapshot({ ...snapshot, error }, workbook)).toThrow(
        'error',
      );
    expect(() =>
      assertExternalCalculationSnapshot({ ...snapshot, error: 'x'.repeat(8192) }, workbook),
    ).not.toThrow();
  });

  it('validates all numeric metadata before authoritative values become current', () => {
    const workbook = workbookFixture();
    for (const metadata of [
      { denominator: '0' },
      { numerator: '1'.repeat(100002) },
      { approximate: true },
      { precision: 15 },
      { precision: 1001 },
      { dimension: [1, 2] },
      { dimension: [33, 0, 0, 0, 0, 0, 0] },
      { temperature: 'difference' },
      { quantityKind: 'x'.repeat(129) },
      { provenance: { constants: 1 } },
      {
        provenance: Object.fromEntries(
          Array.from({ length: 129 }, (_, i) => [`p${i}`, 'revision']),
        ),
      },
    ]) {
      const snapshot = publication(workbook);
      requireDefined(snapshot.result).worksheets[0].values[0][0] = {
        ...exact(),
        ...metadata,
      } as NumericValue;
      expect(() => assertExternalCalculationSnapshot(snapshot, workbook)).toThrow();
    }
  });
});

describe('SE-HOST-CALC-3: bounded draft invalidation', () => {
  it('invalidates transitive readers and keeps unrelated results outside the pending set', () => {
    const workbook = createWorkbook([['1', '=A1+1', '=B1+1', '=42']]);
    const result = calculateWorkbook(workbook),
      key = (x: number) => calculationCellKey('sheet-1', x, 0);
    expect(pendingCalculationCells(result, [key(0)])).toEqual(new Set([key(0), key(1), key(2)]));
    expect(pendingCalculationCells(result, [key(3)])).toEqual(new Set([key(3)]));
    expect(pendingCalculationCells(null, [key(0)])).toEqual(new Set([key(0)]));
  });
  it('SE-4: records formula cells only, and a value cell reads nothing', () => {
    const workbook = createWorkbook([['1', '2', '=A1+B1', '=C1*2']]),
      result = calculateWorkbook(workbook),
      key = (x: number) => calculationCellKey('sheet-1', x, 0);
    expect(result.dependencies).toEqual({
      cells: { [key(2)]: [key(0), key(1)], [key(3)]: [key(2)] },
      scope: 'formulas',
    });
    expect(pendingCalculationCells(result, [key(1)])).toEqual(new Set([key(1), key(2), key(3)]));
    expect(pendingCalculationCells(result, [key(3)])).toEqual(new Set([key(3)]));
  });
  it('invalidates conservatively without dependency metadata and terminates through cycles', () => {
    const result = calculateWorkbook(createWorkbook([['1', '=A1+1', '=42']]));
    delete result.dependencies;
    expect(pendingCalculationCells(result, [calculationCellKey('sheet-1', 0, 0)]).size).toBe(3);
    const [a, b, c] = [0, 1, 2].map((x) => calculationCellKey('sheet-1', x, 0));
    result.dependencies = { cells: { [a]: [b], [b]: [a], [c]: [b] } };
    expect(pendingCalculationCells(result, [a])).toEqual(new Set([a, b, c]));
  });
  it('invalidates missing dependency entries and extent readers when a draft extends a sheet', () => {
    const workbook = createWorkbook([['1', '=ROWS(A:A)', '=42']]),
      result = calculateWorkbook(workbook);
    const key = (x: number, y = 0) => calculationCellKey('sheet-1', x, y);
    // A result without formula scope (for example from an older host) may omit entries.
    delete requireDefined(result.dependencies).cells[key(2)];
    delete requireDefined(result.dependencies).scope;
    requireDefined(result.dependencies).cells[key(0)] = [];
    expect(pendingCalculationCells(result, [key(0)])).toEqual(new Set([key(0), key(1), key(2)]));
    expect(pendingCalculationCells(result, [key(0, 10)], ['sheet-1'])).toEqual(
      new Set([key(0, 10), key(1), key(2)]),
    );
  });
  it('rejects executable JSON hooks before serialization can invoke them', () => {
    const workbook = workbookFixture(),
      snapshot = publication(workbook);
    let calls = 0;
    const hostContext = {
      toJSON() {
        calls++;
        return {};
      },
    };
    expect(() => assertExternalCalculationSnapshot({ ...snapshot, hostContext }, workbook)).toThrow(
      'JSON',
    );
    expect(calls).toBe(0);
  });
});
