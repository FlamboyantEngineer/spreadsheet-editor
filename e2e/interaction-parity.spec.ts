import { expect, test, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
const source = (page: Page) =>
  page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
async function open(page: Page) {
  await page.goto('/full.html');
  await expect(source(page).locator('td[data-x="1"][data-y="1"]').first()).toHaveText('1e-25');
}

test('SE-CR-2.2/2.18: resize previews, cancel, zoom and undo preserve the engine root', async ({
  page,
}) => {
  await open(page);
  const editor = source(page);
  const root = await editor.locator('[data-jspreadsheet-host]').elementHandle();
  const cell = editor.locator('td[data-x="0"][data-y="0"]').first();
  await cell.click();
  await editor.getByRole('button', { name: 'Bold selection', exact: true }).click();
  expect(await root!.evaluate((element) => element.isConnected)).toBe(true);
  const header = editor.locator('thead td[data-x="0"]').first();
  const before = await header.boundingBox();
  await page.mouse.move(before!.x + before!.width - 1, before!.y + 10);
  await page.mouse.down();
  await page.mouse.move(before!.x + before!.width + 39, before!.y + 10);
  await expect
    .poll(async () => (await header.boundingBox())!.width)
    .toBeGreaterThan(before!.width + 30);
  await page.keyboard.press('Escape');
  await page.mouse.up();
  await expect.poll(async () => (await header.boundingBox())!.width).toBeCloseTo(before!.width, 0);
  await editor
    .getByRole('group', { name: 'Spreadsheet zoom', exact: true })
    .getByRole('button', { name: 'Zoom in', exact: true })
    .click();
  await header.scrollIntoViewIfNeeded();
  const zoomed = (await header.boundingBox())!;
  await page.mouse.move(zoomed.x + zoomed.width - 1, zoomed.y + 10);
  await page.mouse.down();
  await page.mouse.move(zoomed.x + zoomed.width + 49, zoomed.y + 10);
  await page.mouse.up();
  await expect
    .poll(async () => (await header.boundingBox())!.width)
    .toBeCloseTo(zoomed.width + 50, 0);
  expect(await root!.evaluate((element) => element.isConnected)).toBe(true);
  await editor.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(async () => (await header.boundingBox())!.width).toBeCloseTo(zoomed.width, 0);
});

test('SE-CR-2.7: native formula completion and reference picking', async ({ page }) => {
  await open(page);
  const editor = source(page),
    cell = editor.locator('td[data-x="1"][data-y="1"]').first();
  await cell.dblclick();
  const input = cell.getByRole('textbox');
  await input.fill('=SUMI');
  await expect(page.getByRole('listbox', { name: 'Formula functions' })).toBeVisible();
  await input.press('Enter');
  await expect(input).toHaveValue('=SUMIF(');
  await input.fill('=');
  await editor.locator('td[data-x="0"][data-y="1"]').first().click();
  await expect(input).toHaveValue('=A2');
  await input.press('Escape');
  await expect(cell).toHaveText('1e-25');
});

test('SE-CR-3.6 and collaboration: reviewer controls and advisory peer locks', async ({ page }) => {
  await open(page);
  const editor = source(page);
  await page.getByRole('button', { name: 'Show collaborator lock', exact: true }).click();
  await editor.locator('td[data-x="1"][data-y="1"]').first().dblclick();
  await expect(editor.locator('.se-editor__grid-surface input')).toHaveCount(0);
  await expect(editor.getByRole('status').filter({ hasText: 'Morgan' })).toBeVisible();
  await expect(editor.getByRole('group', { name: 'Active collaborators' })).toBeVisible();
  await page.getByRole('button', { name: 'Reviewer permissions', exact: true }).click();
  await expect(editor).toHaveAttribute('data-mode', 'review');
  await expect(editor.getByRole('button', { name: 'Save', exact: true })).toBeDisabled();
  await expect(editor.getByRole('button', { name: 'Add worksheet', exact: true })).toBeDisabled();
  await editor.locator('td[data-x="0"][data-y="1"]').first().click({ button: 'right' });
  await expect(
    page.getByRole('menuitem', { name: 'Insert rows above', exact: true }),
  ).toHaveAttribute('aria-disabled', 'true');
});

test('SE-XL-1/2: import preview ignores caches, cancellation preserves the workbook, export uses captured results', async ({
  page,
}) => {
  const document = new ExcelJS.Workbook(),
    sheet = document.addWorksheet('Imported');
  sheet.getCell('A1').value = '0.9999999999999999999999999';
  sheet.getCell('B1').value = { formula: '1-A1', result: 999 };
  sheet.getCell('A2').value = { richText: [{ text: 'Example' }] };
  const bytes = Buffer.from(await document.xlsx.writeBuffer());
  await open(page);
  const editor = source(page);
  async function prepare() {
    await editor.getByRole('menuitem', { name: 'File', exact: true }).click();
    await page.getByRole('menuitem', { name: 'Import Excel…', exact: true }).click();
    const dialog = page.getByRole('dialog', { name: 'Import Excel workbook', exact: true });
    await dialog.locator('input[type=file]').setInputFiles({
      name: 'sample.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: bytes,
    });
    const preview = dialog.getByRole('region', { name: 'Excel import preview', exact: true });
    await expect(preview.locator('td[data-x="1"][data-y="0"]').first()).toHaveText('1e-25');
    await expect(dialog).toContainText('Rich text');
    return dialog;
  }
  const first = await prepare();
  await first.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveText('1e-25');
  const second = await prepare();
  await second.getByRole('button', { name: 'Replace workbook', exact: true }).click();
  await expect(editor.getByRole('tab', { name: 'Imported', exact: true })).toBeVisible();
  await expect(editor.locator('td[data-x="1"][data-y="0"]').first()).toHaveText('1e-25');
  await editor.getByRole('menuitem', { name: 'File', exact: true }).click();
  const exported = page.waitForResponse((response) => response.url().endsWith('/api/xlsx/export'));
  await page.getByRole('menuitem', { name: 'Export to Excel', exact: true }).click();
  const response = await exported;
  expect(response.ok()).toBe(true);
  const downloading = page.waitForEvent('download');
  await page.getByRole('link', { name: 'Download workbook.xlsx', exact: true }).click();
  const download = await downloading;
  const stream = await download.createReadStream();
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.from(chunk));
  const result = new ExcelJS.Workbook();
  await result.xlsx.load(Buffer.concat(chunks) as never);
  expect(result.worksheets[0].getCell('B1').formula).toBe('1-A1');
  expect(result.worksheets[0].getCell('B1').result).toBe(1e-25);
  await expect(
    page.getByRole('link', { name: 'Download workbook.xlsx', exact: true }),
  ).toBeVisible();
});

test('SE-CR-2.11/2.13/2.19: border grid and header dialogs retain selection and keyboard ownership', async ({
  page,
}) => {
  await open(page);
  const editor = source(page),
    cell = editor.locator('td[data-x="1"][data-y="1"]').first();
  await cell.click();
  await editor.getByRole('button', { name: 'Borders', exact: true }).click();
  const borders = page.getByRole('menu', { name: 'Borders', exact: true });
  await expect(borders.getByRole('menuitem')).toHaveCount(7);
  await borders.getByRole('menuitem', { name: 'Top borders', exact: true }).click();
  await expect(cell).toHaveCSS('border-top-style', 'solid');
  const row = editor.locator('tbody td.jss_row[data-y="1"]').first();
  await row.click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Row height…', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Row height', exact: true });
  const size = dialog.getByRole('textbox', { name: 'Size in pixels', exact: true });
  await size.fill('0');
  await expect(dialog.getByRole('button', { name: 'Apply', exact: true })).toBeDisabled();
  await expect(dialog.getByRole('alert')).toContainText('positive');
  await size.fill('34.5');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect.poll(async () => (await row.boundingBox())!.height).toBeCloseTo(34.5, 0);
  await expect(cell).toHaveText('1e-25');
  await row.click();
  await page.keyboard.press('Delete');
  await expect(editor.locator('td[data-x="0"][data-y="1"]').first()).toHaveText('Near one');
});

test('SE-VR-1: saved history opens read-only through context action and returns to current', async ({
  page,
}) => {
  await open(page);
  const editor = source(page);
  await editor.getByRole('button', { name: 'Save', exact: true }).click();
  await editor.getByRole('button', { name: 'History', exact: true }).click();
  const snapshot = editor.getByRole('button').filter({ hasText: 'State: Draft' });
  await expect(snapshot).toHaveCount(1);
  await snapshot.click();
  await expect(editor.locator('.se-editor__historical')).toHaveCount(0);
  await snapshot.click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'View historical version', exact: true }).click();
  await expect(editor).toHaveAttribute('data-mode', 'read');
  await expect(editor.locator('.se-editor__historical')).toContainText('historical version');
  await expect(editor.getByRole('button', { name: 'Save', exact: true })).toBeDisabled();
  await editor.getByRole('button', { name: 'View Current Version', exact: true }).click();
  await expect(editor).toHaveAttribute('data-mode', 'author');
  await expect(editor.locator('.se-editor__historical')).toHaveCount(0);
});

test('SE-CR-3.2: comment single-click opens a bubble while double-click keeps native cell editing', async ({
  page,
}) => {
  await open(page);
  const cell = source(page).locator('td[data-x="1"][data-y="2"]').first();
  await cell.click();
  const thread = page.getByRole('group', { name: 'Comment thread Calculations!B3', exact: true });
  await expect(thread).toBeVisible();
  await expect(thread).toContainText('Keep this value distinct from 1');
  await page.keyboard.press('Escape');
  await expect(thread).not.toBeVisible();
  await cell.dblclick();
  await expect(cell.getByRole('textbox')).toHaveValue('=1-B2');
  await cell.getByRole('textbox').press('Escape');
  await expect(cell).toHaveText('0.9999999999999999999999999');
});
