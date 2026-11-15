import { expect, test } from '@playwright/test';
import { requireDefined } from '../src/model/require-defined.js';

test.use({ viewport: { width: 1440, height: 1000 } });

test('SE-UI-3: finishing unchanged native entries preserves logical types and canonical inputs', async ({
  page,
}) => {
  await page.goto('/full.html');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await editor.getByRole('menuitem', { name: 'File', exact: true }).click();
  await page.getByRole('menuitem', { name: /^Import CSV or workbook JSON/ }).click();
  await page.getByRole('textbox', { name: 'CSV or workbook JSON', exact: true }).fill(
    JSON.stringify({
      schemaVersion: 1,
      worksheets: [
        {
          id: 'typed',
          name: 'Typed',
          data: [[true, '=ISLOGICAL(A1)', '1', '=C1*2']],
          inputValues: {
            C1: { kind: 'number', numerator: '7', denominator: '3', approximate: false },
          },
        },
      ],
    }),
  );
  await page.getByRole('button', { name: 'Import', exact: true }).click();
  const cell = (x: number) => editor.locator(`td[data-x="${x}"][data-y="0"]`).first();
  await expect(cell(1)).toHaveText('TRUE');
  const canonicalResult = await cell(3).textContent();
  for (const x of [0, 2]) {
    await cell(x).dblclick();
    await cell(x).getByRole('textbox').press('Enter');
  }
  await expect(cell(0)).toHaveText('TRUE');
  await expect(cell(1)).toHaveText('TRUE');
  await expect(cell(3)).toHaveText(requireDefined(canonicalResult));
});

test('SE-HOST-1: native completion includes host function documentation', async ({ page }) => {
  await page.goto('/host.html');
  const editor = page.getByRole('region', { name: 'Host spreadsheet', exact: true });
  const cell = editor.locator('td[data-x="0"][data-y="0"]').first();
  await cell.dblclick();
  await cell.getByRole('textbox').fill('=CAP');
  const option = page.getByRole('option', { name: /CAPTURED/ }).first();
  await expect(option).toContainText('CAPTURED()');
  await expect(option).toContainText('A captured value supplied by this example host.');
});

for (const theme of ['light', 'dark'])
  for (const merged of [false, true])
    test(`SE-CHROME-4: ${merged ? 'merged' : 'ordinary'} native editing preserves geometry and visible text in ${theme}`, async ({
      page,
    }) => {
      await page.goto(`/full.html?theme=${theme}`);
      const editor = page.locator('.se-editor').first();
      await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
      const cell = editor.locator('td[data-x="0"][data-y="0"]').first();
      await cell.click();
      if (merged) {
        await page.keyboard.down('Shift');
        await editor.locator('td[data-x="1"][data-y="1"]').first().click();
        await page.keyboard.up('Shift');
        await editor.getByRole('button', { name: 'Merge cells', exact: true }).click();
        await editor
          .getByRole('group', { name: 'Spreadsheet zoom', exact: true })
          .getByRole('button', { name: 'Zoom in', exact: true })
          .click();
        await editor.getByRole('menuitem', { name: 'View', exact: true }).click();
        await page
          .getByRole('menuitem', { name: 'Freeze through selected column', exact: true })
          .click();
      }
      const geometry = async () =>
        editor.evaluate((root) =>
          ['td[data-x="0"][data-y="0"]', 'td[data-x="2"][data-y="2"]'].map((selector) => {
            const element = root.querySelector(selector);
            if (!element) throw new Error(`No ${selector}`);
            const { x, y, width, height } = element.getBoundingClientRect();
            return { x, y, width, height };
          }),
        );
      const before = await geometry();
      const color = await cell.evaluate((element) => getComputedStyle(element).color);
      await cell.click();
      await page.keyboard.type('Visible text');
      const input = cell.getByRole('textbox');
      await expect(input).toHaveValue('Visible text');
      await expect(input).toBeVisible();
      expect(await geometry()).toEqual(before);
      await expect(input).toHaveCSS('color', color);
      await expect(input).toHaveCSS('text-align', 'left');
      const inputBox = requireDefined(await input.boundingBox());
      expect(inputBox.width).toBeLessThanOrEqual(before[0].width);
      expect(inputBox.height).toBeLessThanOrEqual(before[0].height);
      await input.press('Enter');
      await expect(cell).toHaveText('Visible text');
      expect(await geometry()).toEqual(before);
      await cell.dblclick();
      const formula = cell.getByRole('textbox');
      await formula.fill('=B4+1E-25');
      await expect(editor.locator('.se-editor__native-formula')).toContainText('=B4+1E-25');
      await expect(editor.locator('.se-editor__native-formula')).toHaveCSS('color', color);
      expect(await geometry()).toEqual(before);
      await formula.press('Escape');
      await expect(cell).toHaveText('Visible text');
      expect(await geometry()).toEqual(before);
    });
