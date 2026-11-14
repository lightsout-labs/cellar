import { expect, test } from '@playwright/test';
import { openApp } from './grid';

// The longest a frame may take to draw. A 60 frames per second frame lasts
// 16.7 ms, and drawing must leave the browser time for the rest of it; CI
// machines are slower than the laptop of R45, which the human checks.
const FRAME_BUDGET_MS = 16;
const ROWS = 100_000;
const COLUMNS = 100;

test('a sheet of 100,000 by 100 populated cells scrolls within the frame budget', async ({
  page,
  browserName,
}) => {
  test.skip(browserName !== 'chromium', 'Measured on Chromium only.');
  test.setTimeout(120_000);
  await openApp(page);

  // Stored as the app stores a workbook, so that it opens it on reload.
  await page.evaluate(
    ({ rows, columns }) =>
      new Promise<void>((resolve, reject) => {
        // The test plays the platform here: no file format can be opened yet.
        // eslint-disable-next-line no-restricted-globals
        const request = indexedDB.open('cellar');
        request.onerror = () => reject(request.error!);
        request.onsuccess = () => {
          const transaction = request.result.transaction(
            ['workbooks', 'sheets'],
            'readwrite',
          );
          transaction.objectStore('workbooks').put({
            id: 'default',
            name: 'Big',
            modified: Date.now(),
            version: 1,
            sheetIds: ['sheet1'],
          });
          transaction.objectStore('sheets').put({
            workbook: 'default',
            id: 'sheet1',
            name: 'Sheet1',
            columns: Array.from({ length: columns }, (_, col) =>
              Array.from({ length: rows }, (_, row) => row * columns + col),
            ),
            formulas: [],
          });
          transaction.oncomplete = () => resolve();
          transaction.onerror = () => reject(transaction.error!);
        };
      }),
    { rows: ROWS, columns: COLUMNS },
  );
  await page.reload();
  await expect(page.getByRole('status')).toHaveText('A1 0', {
    timeout: 30_000,
  });

  // Each frame scrolls by a few rows and columns, as a wheel or trackpad
  // does, and the time to draw it is measured from the scroll event.
  const times = await page.locator('canvas').evaluate(async (canvas) => {
    // The element that scrolls, around the sheet-sized element.
    const grid = canvas.parentElement!.parentElement!;
    const result: number[] = [];
    for (let step = 1; step <= 300; step++) {
      await new Promise(requestAnimationFrame);
      grid.scrollTop = step * 97;
      grid.scrollLeft = (step % 50) * 31;
      const start = performance.now();
      grid.dispatchEvent(new Event('scroll'));
      result.push(performance.now() - start);
    }
    return result;
  });
  expect(Math.max(...times)).toBeLessThan(FRAME_BUDGET_MS);
});
