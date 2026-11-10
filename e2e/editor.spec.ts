import { expect, test } from '@playwright/test';

test('SE-UI-1: exact examples render in the real grid without browser errors', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await expect(editor.getByRole('status').last()).toHaveText('Calculated on client');
  await expect(editor.locator('td[data-x="1"][data-y="2"]').first()).toHaveText(
    '0.9999999999999999999999999',
  );
  await expect(editor.locator('td[data-x="1"][data-y="3"]').first()).toHaveText('1e-25');
  await expect(editor.locator('td[data-x="2"][data-y="2"]').first()).toHaveText('FALSE');
  expect(errors).toEqual([]);
});

test('SE-UI-2: server mode posts edits and preserves exact values', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Calculate on server', exact: true }).click();
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await expect(editor.getByRole('status').last()).toHaveText('Calculated on server');
  await editor
    .getByRole('textbox', { name: 'Cell value or formula', exact: true })
    .fill('=1-(1-1E-25)');
  const posted = page.waitForResponse(
    (response) =>
      response.url().endsWith('/api/calculate') && response.request().method() === 'POST',
  );
  await editor.getByRole('button', { name: 'Apply', exact: true }).click();
  const response = await posted;
  expect(response.ok()).toBe(true);
  const body = await response.json();
  expect(body.result.worksheets[0].values[0][0]).toMatchObject({
    numerator: '1',
    denominator: '10000000000000000000000000',
    approximate: false,
  });
  await expect(editor.getByLabel('Calculated cell value')).toHaveText('1e-25');
  await expect(editor.locator('td[data-x="0"][data-y="0"]').first()).toHaveText('1e-25');
});

test('SE-UI-3: direct grid edits preserve literals and update dependents', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  const cell = editor.locator('td[data-x="1"][data-y="1"]').first();
  await expect(cell).toHaveText('1e-25');
  await cell.dblclick();
  const input = cell.getByRole('textbox');
  await input.fill('2E-25');
  await input.press('Enter');
  await expect(cell).toHaveText('2e-25');
  await expect(editor.locator('td[data-x="1"][data-y="3"]').first()).toHaveText('2e-25');
  await cell.dblclick();
  await cell.getByRole('textbox').fill('0.9999999999999999999999999');
  await cell.getByRole('textbox').press('Enter');
  await expect(cell).toHaveText('0.9999999999999999999999999');
});

test('SE-UI-7: toolbar edits support undo, redo, and new worksheets', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await editor.getByRole('textbox', { name: 'Cell value or formula', exact: true }).fill('=1/3*3');
  await editor.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(editor.getByLabel('Calculated cell value')).toHaveText('1');
  await editor.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.getByLabel('Calculated cell value')).toHaveText('Example');
  await editor.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor.getByLabel('Calculated cell value')).toHaveText('1');
  await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
  await expect(editor.getByRole('button', { name: 'Sheet3', exact: true })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await editor
    .getByRole('textbox', { name: 'Cell value or formula', exact: true })
    .fill('=Calculations!B2*3');
  await editor.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(editor.getByLabel('Calculated cell value')).toHaveText('3e-25');
});

test('SE-DOC-2: inserts ordinary tables and isolated live spreadsheets', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Preview full table', exact: true }).click();
  const table = page.getByRole('table', { name: 'Inserted document table' });
  await expect(table).toContainText('0.9999999999999999999999999');
  await expect(table).toContainText('≈1.4142135623730950488016887242097');
  await page.getByRole('button', { name: 'Preview live spreadsheet', exact: true }).click();
  const embedded = page.getByRole('region', { name: 'Inserted live spreadsheet', exact: true });
  await expect(embedded.getByRole('status').last()).toHaveText('Calculated on client');
  await embedded
    .getByRole('textbox', { name: 'Cell value or formula', exact: true })
    .fill('=1/3*3');
  await embedded.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(embedded.getByLabel('Calculated cell value')).toHaveText('1');
  const source = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await expect(source.locator('td[data-x="0"][data-y="0"]').first()).toHaveText('Example');
});

test('SE-UI-4: display precision never changes exact comparison results', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: 'Display significant digits', exact: true }).fill('3');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await expect(editor.locator('td[data-x="1"][data-y="2"]').first()).toHaveText('1');
  await expect(editor.locator('td[data-x="2"][data-y="2"]').first()).toHaveText('FALSE');
});

test('SE-UI-5: a failed server request is visible and never falls back locally', async ({
  page,
}) => {
  await page.goto('/');
  await page.route('**/api/calculate', (route) =>
    route.fulfill({ status: 503, body: 'Unavailable' }),
  );
  await page.getByRole('button', { name: 'Calculate on server', exact: true }).click();
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await expect(editor.getByRole('alert')).toContainText('503');
  await expect(editor.locator('td[data-x="1"][data-y="2"]').first()).toHaveText('Unavailable');
  await expect(
    editor.getByRole('button', { name: 'Insert as document table', exact: true }),
  ).toBeDisabled();
});
