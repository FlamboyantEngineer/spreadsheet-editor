import type { Workbook } from '../../src/model/index.js';

export const sample: Workbook = {
  schemaVersion: 1,
  name: 'Precision workbook',
  worksheets: [
    {
      id: 'calculations',
      name: 'Calculations',
      data: [
        ['Example', 'Value', 'Check'],
        ['Small number', '1E-25', '=B2=0'],
        ['Near one', '=1-B2', '=B3=1'],
        ['Recover difference', '=1-B3', '=B4=B2'],
        ['Exact third', '=1/3', '=B5*3'],
        ['Approximate root', '=SQRT(2)', '=B6*B6'],
        ['Cross-sheet', '=Inputs!A1*B2', '=SUM(B2:B4)'],
      ],
      columns: [{ width: 200 }, { width: 300 }, { width: 260 }],
    },
    {
      id: 'inputs',
      name: 'Inputs',
      data: [['3'], ['=Calculations!B3']],
      columns: [{ width: 300 }],
    },
  ],
  comments: [
    {
      id: 'sample-comment',
      sheetId: 'calculations',
      range: 'B3',
      state: 'open',
      entries: [
        {
          id: 'sample-post',
          author: 'morgan',
          text: 'Keep this value distinct from 1 when changing the display precision.',
          createdAt: '2026-11-10T12:00:00.000Z',
        },
      ],
    },
  ],
  suggestions: [
    {
      id: 'sample-suggestion',
      author: { id: 'morgan', name: 'Morgan' },
      createdAt: '2026-11-10T12:00:00.000Z',
      status: 'pending',
      summary: 'Try another input value',
      changes: [{ sheetId: 'inputs', address: 'A1', before: '3', after: '4' }],
    },
  ],
};

export const engineeringSample: Workbook = {
  schemaVersion: 1,
  name: 'Payload mass budget',
  worksheets: [
    {
      id: 'mass',
      name: 'Mass budget',
      data: [
        ['Subsystem', 'Quantity', 'Unit mass (kg)', 'Total (kg)'],
        ['Avionics', '2', '1.275', '=B2*C2'],
        ['Thermal protection', '4', '0.325', '=B3*C3'],
        ['Payload structure', '1', '8.45', '=B4*C4'],
        ['Harness', '3', '0.185', '=B5*C5'],
        ['Total', '', '', '=SUM(D2:D5)'],
        ['Allocation', '', '', '15'],
        ['Margin', '', '', '=D7-D6'],
      ],
      columns: [{ width: 200 }, { width: 100 }, { width: 160 }, { width: 160 }],
      style: {
        A1: 'font-weight: bold',
        B1: 'font-weight: bold',
        C1: 'font-weight: bold',
        D1: 'font-weight: bold',
        A6: 'font-weight: bold',
        D6: 'font-weight: bold',
      },
    },
  ],
};
