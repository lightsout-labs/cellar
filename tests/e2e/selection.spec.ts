import { expect, test, type Page } from '@playwright/test';
import {
  cellRect,
  COLUMN_HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
} from '../../src/app/grid/geometry';
import { clickCell, enter, openApp } from './grid';

const nameBox = (page: Page) => page.getByLabel('Name box');
const status = (page: Page) => page.getByRole('contentinfo');
const canvas = (page: Page) => page.locator('canvas');
const top = { scrollX: 0, scrollY: 0, zoom: 1 };

/** The middle of a cell, at the top of the sheet and at 100%. */
function middle(row: number, col: number) {
  const rect = cellRect({ row, col }, top);
  return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
}

/** Fills A1 = 1, A2 = 2, B1 = 3, B2 = 4 and C5 = 10. */
async function fill(page: Page): Promise<void> {
  await clickCell(page, 0, 0);
  await enter(page, '1');
  await enter(page, '2');
  await clickCell(page, 0, 1);
  await enter(page, '3');
  await enter(page, '4');
  await clickCell(page, 4, 2);
  await enter(page, '10');
}

/**
 * Clicks a point of the canvas with the mouse itself: a locator click
 * would first scroll the grid to bring the canvas into view.
 */
async function clickAt(page: Page, point: { x: number; y: number }) {
  const box = (await canvas(page).boundingBox())!;
  await page.mouse.click(box.x + point.x, box.y + point.y);
}

async function expectSummary(page: Page, summary: string): Promise<void> {
  await expect(page.getByLabel('Summary of the selection')).toHaveText(summary);
}

test('the mouse selects cells, ranges, rows, columns, the sheet and further ranges', async ({
  page,
}) => {
  await openApp(page);
  await fill(page);
  const box = (await canvas(page).boundingBox())!;
  const at = (point: { x: number; y: number }) => ({
    x: box.x + point.x,
    y: box.y + point.y,
  });

  // A range by dragging.
  const a1 = at(middle(0, 0));
  const b2 = at(middle(1, 1));
  await page.mouse.move(a1.x, a1.y);
  await page.mouse.down();
  await page.mouse.move(b2.x, b2.y, { steps: 4 });
  await page.mouse.up();
  await expect(nameBox(page)).toHaveText('A1');
  await expectSummary(page, 'Average: 2.5Count: 4Sum: 10');

  // A range by Shift+click.
  await clickCell(page, 1, 1);
  await canvas(page).click({ position: middle(0, 0), modifiers: ['Shift'] });
  await expect(nameBox(page)).toHaveText('B2');
  await expectSummary(page, 'Average: 2.5Count: 4Sum: 10');

  // A column and a row by their headers.
  await canvas(page).click({ position: { x: middle(0, 0).x, y: 4 } });
  await expectSummary(page, 'Average: 1.5Count: 2Sum: 3');
  await canvas(page).click({ position: { x: 4, y: middle(0, 0).y } });
  await expectSummary(page, 'Average: 2Count: 2Sum: 4');
  await expect(nameBox(page)).toHaveText('A1');

  // The whole sheet by the corner.
  await canvas(page).click({ position: { x: 4, y: 4 } });
  await expectSummary(page, 'Average: 4Count: 5Sum: 20');

  // Further ranges with Ctrl or Cmd.
  await clickCell(page, 0, 0);
  await expectSummary(page, '');
  await canvas(page).click({
    position: middle(4, 2),
    modifiers: ['ControlOrMeta'],
  });
  await expect(nameBox(page)).toHaveText('C5');
  await expectSummary(page, 'Average: 5.5Count: 2Sum: 11');
});

test('the keyboard moves and extends the selection as in Excel', async ({
  page,
}) => {
  await openApp(page);
  await fill(page);
  await clickCell(page, 0, 0);

  await page.keyboard.press('ArrowRight');
  await expect(nameBox(page)).toHaveText('B1');
  await page.keyboard.press('Shift+ArrowLeft');
  await page.keyboard.press('Shift+ArrowDown');
  await expect(nameBox(page)).toHaveText('B1');
  await expectSummary(page, 'Average: 2.5Count: 4Sum: 10');
  await page.keyboard.press('ArrowLeft');
  await expect(nameBox(page)).toHaveText('A1');
  await expectSummary(page, '');

  // To the edges of the data, and of the sheet.
  await page.keyboard.press('ControlOrMeta+ArrowDown');
  await expect(nameBox(page)).toHaveText('A2');
  await page.keyboard.press('ControlOrMeta+Shift+ArrowRight');
  await expectSummary(page, 'Average: 3Count: 2Sum: 6');
  await page.keyboard.press('ControlOrMeta+ArrowDown');
  await expect(nameBox(page)).toHaveText('A100000');
  await page.keyboard.press('ControlOrMeta+ArrowRight');
  await expect(nameBox(page)).toHaveText('CV100000');
  await page.keyboard.press('ControlOrMeta+Home');
  await expect(nameBox(page)).toHaveText('A1');
  await page.keyboard.press('ControlOrMeta+End');
  await expect(nameBox(page)).toHaveText('C5');

  // Home, and End mode, shown in the status bar until the next key.
  await page.keyboard.press('Home');
  await expect(nameBox(page)).toHaveText('A5');
  await page.keyboard.press('End');
  await expect(status(page)).toContainText('End Mode');
  await page.keyboard.press('ArrowUp');
  await expect(nameBox(page)).toHaveText('A2');
  await expect(status(page)).not.toContainText('End Mode');

  // The data around the active cell, then the whole sheet.
  await page.keyboard.press('ControlOrMeta+a');
  await expectSummary(page, 'Average: 2.5Count: 4Sum: 10');
  await page.keyboard.press('ControlOrMeta+a');
  await expectSummary(page, 'Average: 4Count: 5Sum: 20');

  // A page at a time.
  await page.keyboard.press('ArrowDown');
  await expect(nameBox(page)).toHaveText('A3');
  await page.keyboard.press('PageDown');
  await expect(nameBox(page)).not.toHaveText('A3');
  await page.keyboard.press('PageUp');
  await expect(nameBox(page)).toHaveText('A3');
});

test('the grid scrolls to keep the active cell visible, and reaches the last cell', async ({
  page,
}) => {
  await openApp(page);
  // The element that scrolls, around the canvas and the sheet-sized element.
  const grid = canvas(page).locator('xpath=../..');
  const scroll = () =>
    grid.evaluate((node) => ({
      top: node.scrollTop,
      left: node.scrollLeft,
      bottom: node.scrollHeight - node.scrollTop - node.clientHeight,
      right: node.scrollWidth - node.scrollLeft - node.clientWidth,
    }));
  await clickCell(page, 0, 0);

  await page.keyboard.press('ControlOrMeta+ArrowDown');
  await page.keyboard.press('ControlOrMeta+ArrowRight');
  await expect.poll(scroll).toMatchObject({ bottom: 0, right: 0 });
  await page.keyboard.press('ControlOrMeta+Home');
  await expect.poll(scroll).toMatchObject({ top: 0, left: 0 });

  // Every cell is reachable with the scroll bars or the wheel.
  await page.mouse.move(400, 400);
  await page.mouse.wheel(0, 1e9);
  await expect.poll(async () => (await scroll()).bottom).toBe(0);
  await grid.evaluate((node) => node.scrollTo(1e9, 1e9));
  await expect.poll(scroll).toMatchObject({ bottom: 0, right: 0 });
  const box = (await canvas(page).boundingBox())!;
  await clickAt(page, { x: box.width - 4, y: box.height - 4 });
  await expect(nameBox(page)).toHaveText('CV100000');
});

test('the zoom control scales the grid from 50% to 200%', async ({ page }) => {
  await openApp(page);
  const zoomIn = page.getByRole('button', { name: 'Zoom in' });
  const zoomOut = page.getByRole('button', { name: 'Zoom out' });
  const slider = page.getByRole('slider', { name: 'Zoom' });
  await expect(status(page)).toContainText('100%');

  await zoomIn.click();
  await expect(status(page)).toContainText('110%');
  await slider.fill('200');
  await expect(status(page)).toContainText('200%');
  await expect(zoomIn).toBeDisabled();
  // At twice the size, the middle of B2 at 100% is inside A1.
  await clickAt(page, middle(1, 1));
  await expect(nameBox(page)).toHaveText('A1');

  await slider.fill('50');
  await expect(status(page)).toContainText('50%');
  await expect(zoomOut).toBeDisabled();
  await expect(zoomIn).toBeEnabled();
  // At half the size, the middle of B2 at 100% is inside D5.
  await clickAt(page, middle(1, 1));
  await expect(nameBox(page)).toHaveText('D5');
  // The headers are scaled too.
  await canvas(page).click({
    position: { x: ROW_HEADER_WIDTH / 2 + 4, y: COLUMN_HEADER_HEIGHT / 2 + 4 },
  });
  await expect(nameBox(page)).toHaveText('A1');
});
