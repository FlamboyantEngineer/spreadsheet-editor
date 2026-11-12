import { expect, test } from '@playwright/test';

test('SE-CHROME-5: drag highlight and border reach the same range in one paint frame', async ({
  page,
}) => {
  await page.goto('/full.html');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
  const first = editor.locator('td[data-x="0"][data-y="0"]').first();
  const last = editor.locator('td[data-x="4"][data-y="7"]').first();
  await first.click();
  await page.evaluate(
    () =>
      new Promise<void>((done) => requestAnimationFrame(() => requestAnimationFrame(() => done()))),
  );
  const start = (await first.boundingBox())!,
    end = (await last.boundingBox())!;
  await page.evaluate(() => {
    const queue = new Map<number, FrameRequestCallback>();
    let id = 0;
    window.requestAnimationFrame = (callback) => (queue.set(++id, callback), id);
    window.cancelAnimationFrame = (handle) => {
      queue.delete(handle);
    };
    (window as unknown as { paintSelectionFrame(): void }).paintSelectionFrame = () => {
      const callbacks = [...queue.values()];
      queue.clear();
      for (const callback of callbacks) callback(performance.now());
    };
  });
  await page.mouse.move(start.x + 20, start.y + 10);
  await page.mouse.down();
  for (const point of [
    { x: end.x + 20, y: end.y + 10 },
    { x: start.x + 320, y: start.y + 82 },
  ]) {
    await page.mouse.move(point.x, point.y, { steps: 4 });
    const result = await editor.evaluate((root) => {
      (window as unknown as { paintSelectionFrame(): void }).paintSelectionFrame();
      const cells = [...root.querySelectorAll<HTMLElement>('tbody td.highlight:not(.jss_row)')];
      const rectangles = cells.map((cell) => cell.getBoundingClientRect());
      const border = root.querySelector('.se-editor__selection-outline')!.getBoundingClientRect();
      return {
        border: { left: border.left, top: border.top, right: border.right, bottom: border.bottom },
        highlight: {
          left: Math.min(...rectangles.map((r) => r.left)),
          top: Math.min(...rectangles.map((r) => r.top)),
          right: Math.max(...rectangles.map((r) => r.right)),
          bottom: Math.max(...rectangles.map((r) => r.bottom)),
        },
      };
    });
    for (const side of ['left', 'top', 'right', 'bottom'] as const)
      expect(result.border[side]).toBeCloseTo(result.highlight[side], 1);
  }
  await page.mouse.up();
});

for (const theme of ['light', 'dark'])
  test(`SE-CHROME-5: scrolled selection overlays exclude both header bands in ${theme}`, async ({
    page,
  }) => {
    await page.goto(`/full.html?theme=${theme}`);
    const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
    await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
    const first = editor.locator('td[data-x="0"][data-y="0"]').first();
    await first.click();
    await page.keyboard.down('Shift');
    await editor.locator('td[data-x="5"][data-y="9"]').first().click();
    await page.keyboard.up('Shift');
    await editor.locator('.jss_content').evaluate((viewport) => {
      viewport.scrollTop = 100;
      viewport.scrollLeft = 200;
    });
    await expect
      .poll(() =>
        editor.evaluate((root) => {
          const frame = root.querySelector('.se-editor__reference-frame')?.getBoundingClientRect();
          if (!frame) return false;
          const corner = root.querySelector('thead td:first-child')!.getBoundingClientRect();
          const column = root.querySelector('thead td[data-x="0"]')!.getBoundingClientRect();
          const row = root.querySelector('tbody td.jss_row')!.getBoundingClientRect();
          return (
            frame.left >= Math.max(corner.right, row.right) - 0.1 &&
            frame.top >= Math.max(corner.bottom, column.bottom) - 0.1
          );
        }),
      )
      .toBe(true);
    await expect(editor.locator('.se-editor__selection-outline')).toHaveCount(1);
  });

test('SE-CHROME-5: the native fill handle cannot cover headers when selection scrolls out of view', async ({
  page,
}) => {
  await page.goto('/full.html');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
  await editor.locator('td[data-x="0"][data-y="0"]').first().click();
  await page.keyboard.down('Shift');
  await editor.locator('td[data-x="5"][data-y="9"]').first().click();
  await page.keyboard.up('Shift');
  await editor.locator('.jss_content').evaluate((viewport) => {
    viewport.scrollTop = 248;
    viewport.scrollLeft = 620;
  });
  await expect
    .poll(() =>
      editor.locator('.jss_corner').evaluate((handle) => {
        const rect = handle.getBoundingClientRect();
        return (
          document
            .elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2)
            ?.closest('.jss_corner') !== null
        );
      }),
    )
    .toBe(false);
  await editor.locator('.jss_content').evaluate((viewport) => {
    viewport.scrollTop = 0;
    viewport.scrollLeft = 0;
  });
  await expect
    .poll(() =>
      editor.locator('.jss_corner').evaluate((handle) => {
        const rect = handle.getBoundingClientRect();
        return (
          document.elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2) ===
          handle
        );
      }),
    )
    .toBe(true);
});

test('SE-CHROME-5: frozen selection fragments remain inside the cell body at 150% zoom', async ({
  page,
}) => {
  await page.goto('/full.html?theme=dark');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
  const zoom = editor.getByRole('group', { name: 'Spreadsheet zoom', exact: true });
  await zoom.getByRole('button', { name: 'Zoom in', exact: true }).click();
  await zoom.getByRole('button', { name: 'Zoom in', exact: true }).click();
  await editor.locator('td[data-x="0"][data-y="0"]').first().click();
  await editor.getByRole('menuitem', { name: 'View', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Freeze through selected column', exact: true }).click();
  await page.keyboard.down('Shift');
  await editor.locator('td[data-x="5"][data-y="7"]').first().click();
  await page.keyboard.up('Shift');
  await editor.locator('.jss_content').evaluate((viewport) => {
    viewport.scrollLeft = 200;
    viewport.scrollTop = 100;
  });
  await expect(editor.locator('.se-editor__selection-outline')).toHaveCount(2);
  await expect
    .poll(() =>
      editor.evaluate((root) => {
        const frozen = root.querySelector('thead td[data-x="0"]')!.getBoundingClientRect();
        const pieces = [...root.querySelectorAll('.se-editor__selection-outline')].map((node) =>
          node.getBoundingClientRect(),
        );
        return (
          pieces.every((piece) => piece.top >= frozen.bottom - 0.1) &&
          pieces[0].right <= frozen.right + 0.1 &&
          pieces[1].left >= frozen.right - 0.1
        );
      }),
    )
    .toBe(true);
});

test('SE-CHROME-5: ending a drag preserves grid scroll instead of revealing its offscreen anchor', async ({
  page,
}) => {
  await page.goto('/full.html');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
  const first = (await editor.locator('td[data-x="0"][data-y="0"]').first().boundingBox())!;
  const last = (await editor.locator('td[data-x="4"][data-y="7"]').first().boundingBox())!;
  await page.mouse.move(first.x + 20, first.y + 10);
  await page.mouse.down();
  await page.mouse.move(last.x + 20, last.y + 10);
  await editor.locator('.jss_content').evaluate((viewport) => {
    viewport.scrollLeft = 200;
    viewport.scrollTop = 100;
  });
  await page.mouse.up();
  await page.evaluate(
    () =>
      new Promise<void>((done) => requestAnimationFrame(() => requestAnimationFrame(() => done()))),
  );
  expect(
    await editor
      .locator('.jss_content')
      .evaluate((viewport) => ({ left: viewport.scrollLeft, top: viewport.scrollTop })),
  ).toEqual({ left: 200, top: 100 });
});
