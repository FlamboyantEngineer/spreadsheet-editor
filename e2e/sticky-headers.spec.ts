import { expect, type Locator, type Page, test } from '@playwright/test';
import { requireDefined } from '../src/model/require-defined.js';

test.use({ viewport: { width: 1440, height: 1000 } });

async function headerPixels(page: Page, header: Locator) {
  const box = requireDefined(await header.boundingBox());
  // Align inside the header: fractional table origins otherwise include pixels
  // outside its top edge in a locator screenshot when it first becomes sticky.
  return page.screenshot({
    clip: {
      x: Math.ceil(box.x) + 1,
      y: Math.ceil(box.y) + 1,
      width: Math.floor(box.width) - 2,
      height: Math.floor(box.height) - 2,
    },
  });
}

for (const theme of ['light', 'dark'])
  for (const zoom of [100, 125])
    test(`SE-CHROME-6.1: sticky headers stay opaque and retain their fixed edges in ${theme} at ${zoom}%`, async ({
      page,
    }) => {
      await page.goto(`/full.html?theme=${theme}`);
      const editor = page.locator('.se-editor').first();
      await expect(editor.getByRole('status', { name: 'Calculation status' })).toHaveText(
        'Calculated on client',
      );
      if (zoom === 125)
        await editor
          .getByRole('group', { name: 'Spreadsheet zoom', exact: true })
          .getByRole('button', { name: 'Zoom in', exact: true })
          .click();
      await editor.locator('td[data-x="0"][data-y="2"]').first().click();
      await page.keyboard.down('Shift');
      await editor.locator('td[data-x="1"][data-y="2"]').first().click();
      await page.keyboard.up('Shift');
      const column = editor.locator('thead td[data-x="0"]').first();
      // A selected header must keep an opaque base beneath the selection tint.
      const color = await column.evaluate((element) => getComputedStyle(element).backgroundColor);
      expect(color).not.toMatch(/rgba\([^)]*,\s*0?\.\d+\)/);
      await expect(column).toHaveCSS('border-bottom-width', '1px');
      await expect(column).not.toHaveCSS('border-bottom-color', 'rgba(0, 0, 0, 0)');
      const corner = editor.locator('thead td:first-child').first();
      await expect(corner).not.toHaveCSS('border-right-color', 'rgba(0, 0, 0, 0)');
      const row = editor.locator('tbody td.jss_row[data-y="2"]').first();
      await expect(row).not.toHaveCSS('border-right-color', 'rgba(0, 0, 0, 0)');

      // These are actual rendered pixels, not just overlay bounds: different
      // selected values, cell text and gridlines scroll underneath the same header.
      const baseline = await headerPixels(page, column);
      for (const top of [52, 76, 103]) {
        await editor.locator('.jss_content').evaluate((viewport, top) => {
          viewport.scrollTop = top;
        }, top);
        await expect
          .poll(() => editor.locator('.jss_content').evaluate((v) => v.scrollTop))
          .toBe(top);
        expect((await headerPixels(page, column)).equals(baseline)).toBe(true);
      }
      await editor.locator('.jss_content').evaluate((viewport) => {
        viewport.scrollTop = 0;
        viewport.scrollLeft = 80;
      });
      await expect(row).toHaveCSS('border-right-width', '1px');
      const first = await headerPixels(page, row);
      await editor.locator('.jss_content').evaluate((viewport) => {
        viewport.scrollLeft = 140;
      });
      expect((await headerPixels(page, row)).equals(first)).toBe(true);
    });
