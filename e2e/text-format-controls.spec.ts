import { expect, test, type Page } from '@playwright/test';

async function open(page: Page, chrome: 'full' | 'embedded') {
  await page.goto(`/${chrome}.html`);
  const editor = page.getByRole('region', {
    name: chrome === 'embedded' ? 'Embedded spreadsheet' : 'Spreadsheet editor',
    exact: true,
  });
  await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
  await editor.locator('td[data-x="0"][data-y="0"]').first().click();
  await page.keyboard.down('Shift');
  await editor.locator('td[data-x="1"][data-y="1"]').first().click();
  await page.keyboard.up('Shift');
  return editor;
}

for (const chrome of ['full', 'embedded'] as const) {
  test(`SE-CHROME-7: custom font sizes and preset stepping format the range in ${chrome}`, async ({
    page,
  }) => {
    const editor = await open(page, chrome);
    if (chrome === 'embedded')
      await editor.getByRole('button', { name: 'Font size', exact: true }).click();
    const input = page.getByRole('textbox', { name: 'Font size', exact: true });
    await input.fill('13.5');
    await input.press('Enter');
    const selected = editor.locator('td[data-x="1"][data-y="1"]').first();
    await expect(selected).toHaveCSS('font-size', '13.5px');
    await page.getByRole('button', { name: 'Increase Font size', exact: true }).click();
    await expect(selected).toHaveCSS('font-size', '14px');
    await input.press('ArrowDown');
    await expect(selected).toHaveCSS('font-size', '12px');
    await input.fill('17');
    await input.press('Tab');
    await expect(selected).toHaveCSS('font-size', '17px');
    if (chrome === 'embedded') await page.keyboard.press('Escape');
    for (const x of [0, 1])
      for (const y of [0, 1])
        await expect(editor.locator(`td[data-x="${x}"][data-y="${y}"]`).first()).toHaveCSS(
          'font-size',
          '17px',
        );
  });

  test(`SE-CHROME-7: glyph color palettes, custom entry and clear format cells in ${chrome}`, async ({
    page,
  }) => {
    const editor = await open(page, chrome);
    const cell = editor.locator('td[data-x="1"][data-y="1"]').first();
    const textColor = editor.getByRole('button', { name: 'Text color', exact: true });
    const fillColor = editor.getByRole('button', { name: 'Cell fill', exact: true });
    await textColor.click();
    await page.getByRole('option', { name: '#b42318', exact: true }).click();
    await expect(cell).toHaveCSS('color', 'rgb(180, 35, 24)');
    await expect(textColor.locator('.ui-color-swatch-button__swatch')).toHaveCSS(
      'background-color',
      'rgb(180, 35, 24)',
    );
    await expect(textColor).toBeFocused();
    await fillColor.click();
    await page.getByRole('option', { name: '#fef0c7', exact: true }).click();
    await expect(cell).toHaveCSS('background-color', 'rgb(254, 240, 199)');
    await fillColor.click();
    const custom = page.getByRole('textbox', { name: 'Custom cell fill', exact: true });
    await custom.fill('#336699');
    await custom.press('Enter');
    await expect(cell).toHaveCSS('background-color', 'rgb(51, 102, 153)');
    await page.getByRole('button', { name: 'Clear cell fill', exact: true }).click();
    expect(await cell.evaluate((element) => element.style.backgroundColor)).toBe('');
    await editor.locator('td[data-x="0"][data-y="2"]').first().click();
    await expect(cell).toHaveCSS('background-color', 'rgb(255, 255, 255)');
    await expect(fillColor.locator('.ui-color-swatch-button__swatch')).toHaveAttribute(
      'data-empty',
      'true',
    );
    // Clearing the fill leaves the separately authored text color intact.
    await expect(cell).toHaveCSS('color', 'rgb(180, 35, 24)');
  });
}

test('SE-CHROME-7: palette formatting commits a pending formula with one undoable change', async ({
  page,
}) => {
  const editor = await open(page, 'embedded');
  const formula = editor.locator('.se-editor__formula-input textarea');
  await formula.fill('=1E-25');
  await editor.getByRole('button', { name: 'Text color', exact: true }).click();
  await page.getByRole('option', { name: '#175cd3', exact: true }).click();
  const cell = editor.locator('td[data-x="0"][data-y="0"]').first();
  await expect(formula).toHaveValue('=1E-25');
  await expect(cell).toHaveCSS('color', 'rgb(23, 92, 211)');
  await editor.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(formula).toHaveValue('');
  await expect(cell).not.toHaveCSS('color', 'rgb(23, 92, 211)');
  await editor.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(formula).toHaveValue('=1E-25');
  await expect(cell).toHaveCSS('color', 'rgb(23, 92, 211)');
});
