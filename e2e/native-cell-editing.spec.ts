import { expect, test } from '@playwright/test';

test.use({ viewport: { width: 1440, height: 1000 } });

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
            const { x, y, width, height } = root.querySelector(selector)!.getBoundingClientRect();
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
      const inputBox = (await input.boundingBox())!;
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
