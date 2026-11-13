import { test } from '@playwright/test';
import { clickCell, enter, expectCell, openApp } from './grid';

test('cells are restored from browser storage after a reload', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '2');
  await enter(page, '=A1+1');
  await expectCell(page, 1, 0, 'A2 3');

  // Autosave runs half a second after the last change.
  await page.waitForTimeout(1000);
  await page.reload();
  await expectCell(page, 1, 0, 'A2 3');
  await expectCell(page, 0, 0, 'A1 2');

  // The formula itself was stored, not only its result.
  await clickCell(page, 0, 0);
  await enter(page, '10');
  await expectCell(page, 1, 0, 'A2 11');
});
