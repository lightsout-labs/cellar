import { expect, type Page } from '@playwright/test';
import { cellRect } from '../../src/app/grid/geometry';

/** The sheet scrolled to the top at 100%, where these helpers work. */
const TOP = { scrollX: 0, scrollY: 0, zoom: 1 };

/**
 * Opens the app and waits for the grid. With CELLAR_COMMIT set, first waits
 * until the address serves the build of that commit, not a cached older one.
 */
export async function openApp(page: Page): Promise<void> {
  const commit = process.env.CELLAR_COMMIT;
  if (commit) {
    await expect(async () => {
      await page.goto('./');
      await expect(page.locator('html')).toHaveAttribute(
        'data-commit',
        commit,
        { timeout: 1000 },
      );
    }).toPass({ timeout: 120_000 });
  }
  await page.goto('./');
  await expect(page.locator('canvas')).toBeVisible();
}

/** Clicks the middle of a cell of the canvas grid. */
export async function clickCell(
  page: Page,
  row: number,
  col: number,
): Promise<void> {
  const rect = cellRect({ row, col }, TOP);
  await page.locator('canvas').click({
    position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
  });
}

/** Double-clicks the middle of a cell of the canvas grid. */
export async function doubleClickCell(
  page: Page,
  row: number,
  col: number,
): Promise<void> {
  const rect = cellRect({ row, col }, TOP);
  await page.locator('canvas').dblclick({
    position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
  });
}

/** Types an entry into the active cell and commits it with Enter. */
export async function enter(page: Page, text: string): Promise<void> {
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
}

/** Selects a cell and expects the grid to announce its name and contents. */
export async function expectCell(
  page: Page,
  row: number,
  col: number,
  announced: string,
): Promise<void> {
  await clickCell(page, row, col);
  await expect(page.getByRole('status')).toHaveText(announced);
}

/**
 * Whether the canvas has drawn anything inside a cell, or in its left or
 * right half, away from its grid lines: whether any pixel there differs
 * from the top-left one.
 */
export async function isCellPainted(
  page: Page,
  row: number,
  col: number,
  part: 'whole' | 'left' | 'right' = 'whole',
): Promise<boolean> {
  const cell = cellRect({ row, col }, TOP);
  const half = cell.width / 2;
  const rect = {
    ...cell,
    x: part === 'right' ? cell.x + half : cell.x,
    width: part === 'whole' ? cell.width : half,
  };
  return page.locator('canvas').evaluate((canvas: HTMLCanvasElement, rect) => {
    const scale = canvas.width / canvas.clientWidth;
    const inset = 4;
    const { data } = canvas
      .getContext('2d')!
      .getImageData(
        (rect.x + inset) * scale,
        (rect.y + inset) * scale,
        (rect.width - 2 * inset) * scale,
        (rect.height - 2 * inset) * scale,
      );
    for (let i = 4; i < data.length; i += 4) {
      if (
        data[i] !== data[0] ||
        data[i + 1] !== data[1] ||
        data[i + 2] !== data[2]
      ) {
        return true;
      }
    }
    return false;
  }, rect);
}

/** The red, green and blue of the background the canvas drew in a cell. */
export async function cellBackground(
  page: Page,
  row: number,
  col: number,
): Promise<number[]> {
  const rect = cellRect({ row, col }, TOP);
  return page.locator('canvas').evaluate((canvas: HTMLCanvasElement, rect) => {
    const scale = canvas.width / canvas.clientWidth;
    const { data } = canvas
      .getContext('2d')!
      .getImageData((rect.x + 4) * scale, (rect.y + 4) * scale, 1, 1);
    return [data[0]!, data[1]!, data[2]!];
  }, rect);
}

/**
 * Stores a workbook of one sheet as the app stores it, and reloads the page
 * so that the app opens it. The test plays the platform here.
 */
export async function openStoredSheet(
  page: Page,
  sheet: { columns: unknown[][]; formulas: [number, number, string][] },
): Promise<void> {
  await page.evaluate(
    (sheet) =>
      new Promise<void>((resolve, reject) => {
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
            ...sheet,
          });
          transaction.oncomplete = () => resolve();
          transaction.onerror = () => reject(transaction.error!);
        };
      }),
    sheet,
  );
  await page.reload();
}
