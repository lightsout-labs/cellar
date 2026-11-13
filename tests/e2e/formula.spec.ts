import { expect, test } from '@playwright/test';
import { clickCell, enter, expectCell, isCellPainted, openApp } from './grid';

// Also the smoke test that the deployment runs against the live address.
test('a formula typed in the grid shows its result @smoke', async ({
  page,
}) => {
  test.setTimeout(150_000); // A new deployment can take a while to be served.
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '2'); // Enter moves down to A2.
  await enter(page, '=A1+1');

  await expectCell(page, 1, 0, 'A2 3');
  await expectCell(page, 0, 0, 'A1 2');

  // The result is drawn on the canvas, in A2 and not beside it. The
  // selection is moved away first so that its border is not what is seen.
  await clickCell(page, 5, 3);
  expect(await isCellPainted(page, 1, 0)).toBe(true);
  expect(await isCellPainted(page, 1, 1)).toBe(false);
});

test('Escape abandons an entry', async ({ page }) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await page.keyboard.type('7');
  await page.keyboard.press('Escape');
  await expectCell(page, 0, 0, 'A1');
});
