import { expect, test, type Page } from '@playwright/test';
import { cellRect } from '../../src/app/grid/geometry';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US' });

const formulaBar = (page: Page) =>
  page.getByRole('textbox', { name: 'Formula bar' });
const suggestions = (page: Page) =>
  page.getByRole('listbox', { name: 'Functions' });
const hint = (page: Page) => page.getByRole('note', { name: 'Arguments' });

/** Where the middle of a cell is on the page, at the top of the sheet. */
async function cellPoint(page: Page, row: number, col: number) {
  const canvas = (await page.locator('canvas').boundingBox())!;
  const rect = cellRect({ row, col }, { scrollX: 0, scrollY: 0, zoom: 1 });
  return {
    x: canvas.x + rect.x + rect.width / 2,
    y: canvas.y + rect.y + rect.height / 2,
  };
}

/** The color the canvas drew just inside the left edge of a cell. */
async function leftEdgeColor(page: Page, row: number, col: number) {
  const rect = cellRect({ row, col }, { scrollX: 0, scrollY: 0, zoom: 1 });
  return page.locator('canvas').evaluate((canvas: HTMLCanvasElement, rect) => {
    const scale = canvas.width / canvas.clientWidth;
    const { data } = canvas
      .getContext('2d')!
      .getImageData(
        Math.floor((rect.x + 1) * scale),
        Math.floor((rect.y + rect.height / 2) * scale),
        1,
        1,
      );
    return `rgb(${data[0]}, ${data[1]}, ${data[2]})`;
  }, rect);
}

test('function names are completed and their arguments shown while typing', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '4');
  await enter(page, '6');
  await clickCell(page, 3, 0);
  await page.keyboard.type('=sumi');
  await expect(suggestions(page).getByRole('option')).toHaveText([
    'SUMIF',
    'SUMIFS',
  ]);
  await expect(
    suggestions(page).getByRole('option', { name: 'SUMIF', exact: true }),
  ).toHaveAttribute('aria-selected', 'true');
  // The arrows move through the list rather than leaving the cell.
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Tab');
  await expect(formulaBar(page)).toHaveValue('=SUMIF(');
  await expect(suggestions(page)).toBeHidden();

  await expect(hint(page)).toHaveText('SUMIF(range, criteria, [sum_range])');
  await expect(hint(page).locator('strong')).toHaveText('range');
  await page.keyboard.type('A1:A2,">5",');
  await expect(hint(page).locator('strong')).toHaveText('[sum_range]');
  await page.keyboard.type('A1:A2)');
  await expect(hint(page)).toBeHidden();
  await page.keyboard.press('Enter');
  await expectCell(page, 3, 0, 'A4 6');

  // A click chooses too; a name no function starts with offers nothing,
  // and Escape closes the list but keeps the entry.
  await clickCell(page, 4, 0);
  await page.keyboard.type('=ro');
  await suggestions(page).getByRole('option', { name: 'ROUNDUP' }).click();
  await expect(formulaBar(page)).toHaveValue('=ROUNDUP(');
  await page.keyboard.type('2.1,0)+xyz');
  await expect(suggestions(page)).toBeHidden();
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.type('ab');
  await expect(suggestions(page).getByRole('option')).toHaveText(['ABS']);
  await page.keyboard.press('Escape');
  await expect(suggestions(page)).toBeHidden();
  await expect(formulaBar(page)).toHaveValue('=ROUNDUP(2.1,0)+ab');
});

test('clicking and dragging on the grid inserts references, on another sheet too', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '1');
  await enter(page, '2');
  await enter(page, '3');
  await clickCell(page, 0, 1);
  await enter(page, '10');

  await clickCell(page, 4, 0);
  await page.keyboard.type('=SUM(');
  const from = await cellPoint(page, 0, 0);
  const to = await cellPoint(page, 2, 1);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await expect(formulaBar(page)).toHaveValue('=SUM(A1');
  await page.mouse.move(to.x, to.y, { steps: 4 });
  await page.mouse.up();
  await expect(formulaBar(page)).toHaveValue('=SUM(A1:B3');
  // The entry goes on, in the cell it started in.
  await expect(page.getByRole('textbox', { name: 'A5' })).toBeFocused();
  await page.keyboard.type(')');
  await page.keyboard.press('Enter');
  await expectCell(page, 4, 0, 'A5 16');

  // Where no reference can go, a click commits the entry and selects.
  await clickCell(page, 5, 0);
  await page.keyboard.type('=5');
  await clickCell(page, 0, 3);
  await expect(page.getByRole('status')).toHaveText('D1');
  await expectCell(page, 5, 0, 'A6 5');

  // A reference to another sheet.
  await page.getByRole('button', { name: 'New sheet' }).click();
  await clickCell(page, 2, 2);
  await enter(page, '7');
  await page.getByRole('tab', { name: 'Sheet1', exact: true }).click();
  await clickCell(page, 6, 0);
  await page.keyboard.type('=A1+');
  await page.getByRole('tab', { name: 'Sheet2', exact: true }).click();
  await clickCell(page, 2, 2);
  await expect(formulaBar(page)).toHaveValue('=A1+Sheet2!C3');
  await expect(formulaBar(page)).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(
    page.getByRole('tab', { name: 'Sheet1', exact: true }),
  ).toHaveAttribute('aria-selected', 'true');
  await expectCell(page, 6, 0, 'A7 8');
});

test('each reference has its own color, and its range is outlined in it', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 3);
  await page.keyboard.type('=A1+B2:C3*a1');
  const editor = page.getByRole('textbox', { name: 'D1' });
  const references = page.locator('[data-place="cell"] [data-reference]');
  await expect(references).toHaveText(['A1', 'B2:C3', 'a1']);
  const colors = await references.evaluateAll((spans) =>
    spans.map((span) => getComputedStyle(span).color),
  );
  expect(colors[0]).toBe(colors[2]);
  expect(colors[0]).not.toBe(colors[1]);
  await expect(editor).toBeFocused();

  expect(await leftEdgeColor(page, 0, 0)).toBe(colors[0]);
  expect(await leftEdgeColor(page, 1, 1)).toBe(colors[1]);
  expect(await leftEdgeColor(page, 2, 1)).toBe(colors[1]);
  const plain = await leftEdgeColor(page, 5, 5);
  expect(await leftEdgeColor(page, 1, 2)).toBe(plain); // Inside the range.

  await page.keyboard.press('Escape');
  await expect(references).toHaveCount(0);
  expect(await leftEdgeColor(page, 0, 0)).toBe(plain);
});

test('F4 cycles the reference at the caret through absolute and relative', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await page.keyboard.type('=1+B2');
  for (const text of ['=1+$B$2', '=1+B$2', '=1+$B2', '=1+B2']) {
    await page.keyboard.press('F4');
    await expect(formulaBar(page)).toHaveValue(text);
  }
  // In Edit mode the arrows move the caret onto another reference, and
  // the caret stays after the reference that F4 changes.
  await page.keyboard.type('+C3');
  await page.keyboard.press('F2');
  for (let left = 0; left < 3; left++) await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('F4');
  await page.keyboard.type('*2');
  await expect(formulaBar(page)).toHaveValue('=1+$B$2*2+C3');
});

test('Insert function on the Formulas tab lists every function for a formula', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '-3');
  await page.getByRole('tab', { name: 'Formulas' }).click();
  await clickCell(page, 1, 0);
  const insert = page.getByRole('button', { name: 'Insert function' });
  await insert.hover();
  await expect(
    page.getByRole('tooltip', { name: 'Choose a function to insert' }),
  ).toBeVisible();
  await insert.click();
  await expect(formulaBar(page)).toHaveValue('=');
  await expect(formulaBar(page)).toBeFocused();
  await expect(suggestions(page).getByRole('option').first()).toHaveText('ABS');
  await expect(suggestions(page).getByRole('option')).not.toHaveCount(1);
  await suggestions(page).getByRole('option', { name: 'ABS' }).click();
  await expect(formulaBar(page)).toHaveValue('=ABS(');
  await expect(hint(page)).toHaveText('ABS(number)');

  // In a formula, the function goes at the caret.
  await page.keyboard.type('A1)+');
  await insert.click();
  await expect(formulaBar(page)).toBeFocused();
  await page.keyboard.type('SQ');
  await expect(suggestions(page).getByRole('option')).toHaveText([
    'SQRT',
    'SQRTPI',
  ]);
  await page.keyboard.press('Enter');
  await expect(formulaBar(page)).toHaveValue('=ABS(A1)+SQRT(');
  await page.keyboard.type('16)');
  await page.keyboard.press('Enter');
  await expectCell(page, 1, 0, 'A2 7');
});

test.describe('in Chinese and the dark theme', () => {
  test.use({ locale: 'zh-CN', colorScheme: 'dark' });

  test('the list and the hint are translated and dark, with hover states', async ({
    page,
  }) => {
    await openApp(page);
    await clickCell(page, 0, 0);
    await page.keyboard.type('=ro');
    const list = page.getByRole('listbox', { name: '函数' });
    await expect(list).toBeVisible();
    const background = (selector: string) =>
      page
        .locator(selector)
        .first()
        .evaluate((node) => getComputedStyle(node).backgroundColor);
    const dark = (color: string) => {
      const [red, green, blue] = color.match(/\d+/g)!.map(Number);
      return (red! + green! + blue!) / 3 < 128;
    };
    expect(dark(await background('[role="listbox"]'))).toBe(true);
    const option = list.getByRole('option').nth(1);
    const resting = await option.evaluate(
      (node) => getComputedStyle(node).backgroundColor,
    );
    await option.hover();
    await expect
      .poll(() =>
        option.evaluate((node) => getComputedStyle(node).backgroundColor),
      )
      .not.toBe(resting);
    await page.keyboard.press('Tab');
    await expect(page.getByRole('note', { name: '参数' })).toBeVisible();
    expect(dark(await background('[role="note"]'))).toBe(true);
    await page.getByRole('tab', { name: '公式' }).click();
    await expect(page.getByRole('button', { name: '插入函数' })).toBeVisible();
  });
});
