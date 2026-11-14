import { expect, test, type Page } from '@playwright/test';

async function open(page: Page, theme = 'light', lab = 'full') {
  await page.goto(`/${lab}.html?theme=${theme}`);
  const editor = page.locator('.se-editor').first();
  await expect(editor.getByRole('status', { name: 'Calculation status' })).toHaveText(
    'Calculated on client',
  );
  return editor;
}

for (const theme of ['light', 'dark']) {
  test(`SE-CHROME-5: fixed chrome, scaled cells and four selection edges in ${theme}`, async ({
    page,
  }) => {
    const editor = await open(page, theme);
    const cell = editor.locator('td[data-x="1"][data-y="1"]').first();
    await cell.click();
    const measure = () =>
      editor.evaluate((root) => {
        const box = (selector: string) => {
          const rect = root.querySelector(selector)!.getBoundingClientRect();
          return { width: rect.width, height: rect.height };
        };
        const cell = root.querySelector('td[data-x="1"][data-y="1"]')!;
        const style = getComputedStyle(cell);
        const outlines = Array.from(root.querySelectorAll('.se-editor__selection-outline')).map(
          (node) => {
            const s = getComputedStyle(node);
            return {
              widths: [
                s.borderTopWidth,
                s.borderRightWidth,
                s.borderBottomWidth,
                s.borderLeftWidth,
              ],
              color: s.borderTopColor,
            };
          },
        );
        return {
          editor: box('.se-editor__workspace'),
          grid: box('.se-editor__grid'),
          toolbar: box('.se-editor__toolbar'),
          formula: box('.se-editor__formula-bar'),
          tabs: box('.se-editor__sheets'),
          cell: box('td[data-x="1"][data-y="1"]'),
          font: parseFloat(style.fontSize),
          strokes: [
            style.borderTopWidth,
            style.borderRightWidth,
            style.borderBottomWidth,
            style.borderLeftWidth,
          ],
          outlines,
        };
      });
    const baseline = await measure();
    expect(baseline.outlines).toHaveLength(1);
    expect(baseline.outlines[0].widths).toEqual(['1px', '1px', '1px', '1px']);
    expect(baseline.outlines[0].color).not.toBe('rgba(0, 0, 0, 0)');
    const zoom = editor.getByRole('group', { name: 'Spreadsheet zoom', exact: true });
    await expect(zoom.getByRole('button', { name: /reset/i })).toHaveCount(0);
    await zoom.getByRole('button', { name: 'Zoom in', exact: true }).click();
    await expect(zoom).toContainText('125%');
    await expect.poll(async () => (await measure()).font).toBeCloseTo(baseline.font * 1.25, 1);
    const scaled = await measure();
    for (const key of ['editor', 'grid', 'toolbar', 'formula', 'tabs'] as const)
      expect(scaled[key]).toEqual(baseline[key]);
    expect(scaled.cell.width).toBeCloseTo(baseline.cell.width * 1.25, 0);
    expect(scaled.cell.height).toBeCloseTo(baseline.cell.height * 1.25, 0);
    expect(scaled.strokes).toEqual(baseline.strokes);
    expect(scaled.outlines[0].widths).toEqual(['1px', '1px', '1px', '1px']);
    await expect(editor.getByRole('button', { name: 'All worksheets', exact: true })).toHaveCount(
      0,
    );
  });
}

test('SE-CHROME-5: embedded chrome uses accessible icon actions and public formula composition', async ({
  page,
}) => {
  const editor = await open(page, 'light', 'embedded');
  const toolbar = editor.getByRole('toolbar', { name: 'Spreadsheet formatting', exact: true });
  expect(await toolbar.locator('button').count()).toBeGreaterThan(12);
  for (const button of await toolbar.locator('button').all()) {
    await expect(button).toHaveAccessibleName(/\S/);
    await expect(button.locator('svg')).toHaveCount(1);
    expect((await button.innerText()).trim()).toBe('');
  }
  await expect(toolbar).toHaveClass(/ui-toolbar-row/);
  await expect(editor.locator('.se-editor__formula-widget')).toHaveCount(1);
  await expect(editor.locator('.se-editor__formula-input textarea')).toHaveCount(1);
  await expect(editor.locator('.se-editor__formula-overlay')).toHaveCount(0);
  for (const name of ['Apply', 'Cancel cell edit', 'Restore', 'All worksheets'])
    await expect(editor.getByRole('button', { name, exact: true })).toHaveCount(0);
});

test('SE-CHROME-5: pending formula survives formatting and inactive inline worksheet rename', async ({
  page,
}) => {
  const editor = await open(page);
  const formula = editor.getByRole('textbox', { name: 'Cell value or formula', exact: true });
  await editor.locator('td[data-x="1"][data-y="1"]').first().click();
  await formula.fill('2E-25');
  await editor.getByRole('button', { name: 'Bold selection', exact: true }).click();
  await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveText('2e-25');
  await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveCSS(
    'font-weight',
    '700',
  );
  await formula.fill('3E-25');
  await editor.getByRole('tab', { name: 'Inputs', exact: true }).click({ button: 'right' });
  await expect(editor.getByRole('tab', { name: 'Calculations', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  const rename = page.getByRole('menuitem', { name: 'Rename', exact: true });
  await expect(rename.locator('svg')).toHaveCount(1);
  await expect(rename).toContainText('F2');
  await rename.click();
  const input = editor.getByRole('textbox', { name: 'Rename Inputs', exact: true });
  await input.fill('Parameters');
  await input.press('Enter');
  await expect(editor.getByRole('tab', { name: 'Parameters', exact: true })).toBeVisible();
  await expect(editor.getByRole('dialog', { name: 'Rename worksheet', exact: true })).toHaveCount(
    0,
  );
  await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveText('3e-25');
  await expect(editor.locator('td[data-x="1"][data-y="6"]').first()).toHaveText('9e-25');
});

test('SE-CHROME-5: cancelling an existing formula restores ordinary grid selection', async ({
  page,
}) => {
  const editor = await open(page);
  const formula = editor.getByRole('textbox', { name: 'Cell value or formula', exact: true });
  await editor.locator('td[data-x="1"][data-y="2"]').first().click();
  await formula.fill('=B2+7');
  await formula.press('Escape');
  await expect(formula).toHaveValue('=1-B2');
  await editor.locator('td[data-x="0"][data-y="4"]').first().click();
  await expect(
    editor.getByRole('textbox', { name: 'Selected cell or range', exact: true }),
  ).toHaveValue('A5');
  await expect(editor.locator('td[data-x="1"][data-y="2"]').first()).toHaveText(
    '0.9999999999999999999999999',
  );
});
