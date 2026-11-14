import { expect, test, type Locator, type Page } from '@playwright/test';
import { cellBackground, clickCell, enter, openApp } from './grid';

const nameBox = (page: Page) => page.getByLabel('Name box');
const formulaBar = (page: Page) => page.getByLabel('Formula bar');

/** Whether a color such as `rgb(27, 26, 25)` is nearer black than white. */
async function isDark(element: Locator, property: string): Promise<boolean> {
  const color = await element.evaluate(
    (node, property) => getComputedStyle(node).getPropertyValue(property),
    property,
  );
  const [red, green, blue] = color.match(/\d+/g)!.map(Number);
  return (red! + green! + blue!) / 3 < 128;
}

test('the shell is laid out as in Excel for the web at 1024 px', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1024, height: 768 });
  await openApp(page);

  // From top to bottom, each part below the one before and inside the window.
  const parts = [
    page.getByRole('tablist', { name: 'Ribbon' }),
    page.getByRole('tabpanel', { name: 'View' }),
    formulaBar(page),
    page.locator('canvas'),
    page.getByRole('tablist', { name: 'Sheets' }),
    page.getByRole('contentinfo'),
  ];
  let bottom = 0;
  for (const part of parts) {
    await expect(part).toBeVisible();
    const box = (await part.boundingBox())!;
    expect(box.y).toBeGreaterThanOrEqual(bottom);
    expect(box.x).toBeGreaterThanOrEqual(0);
    expect(box.x + box.width).toBeLessThanOrEqual(1024);
    bottom = box.y + box.height;
  }
  expect(bottom).toBeLessThanOrEqual(768);

  // The name box is at the left of the formula bar, on its line.
  const name = (await nameBox(page).boundingBox())!;
  const formula = (await formulaBar(page).boundingBox())!;
  expect(name.x + name.width).toBeLessThanOrEqual(formula.x);
  expect(name.y).toBeLessThan(formula.y + formula.height);
  expect(formula.y).toBeLessThan(name.y + name.height);

  await expect(page.getByRole('tab', { name: 'Sheet1' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(page.getByRole('contentinfo')).toHaveText('Ready');
});

test('the name box and the formula bar follow the active cell', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '2');
  await enter(page, '=A1+1');

  await clickCell(page, 1, 0);
  await expect(nameBox(page)).toHaveText('A2');
  await expect(formulaBar(page)).toHaveText('=A1+1');
  await clickCell(page, 0, 0);
  await expect(nameBox(page)).toHaveText('A1');
  await expect(formulaBar(page)).toHaveText('2');

  // An entry shows in the formula bar while it is typed.
  await clickCell(page, 2, 1);
  await expect(nameBox(page)).toHaveText('B3');
  await expect(formulaBar(page)).toHaveText('');
  await page.keyboard.type('45');
  await expect(formulaBar(page)).toHaveText('45');
  await expect(page.getByRole('contentinfo')).toHaveText('Enter');
  await page.keyboard.press('Escape');
  await expect(formulaBar(page)).toHaveText('');
});

test.describe('in a browser set to English', () => {
  test.use({ locale: 'en-US' });

  test('the interface starts in English', async ({ page }) => {
    await openApp(page);
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
    await expect(page.getByRole('tab', { name: 'View' })).toBeVisible();
  });
});

test.describe('in a browser set to Chinese', () => {
  test.use({ locale: 'zh-CN' });

  test('the interface starts in Chinese, switches, and remembers', async ({
    page,
  }) => {
    await openApp(page);
    await expect(page.locator('html')).toHaveAttribute('lang', 'zh-CN');
    const chinese = [
      page.getByRole('tablist', { name: '功能区' }),
      page.getByRole('tab', { name: '视图' }),
      page.getByRole('group', { name: '主题' }),
      page.getByRole('button', { name: '浅色' }),
      page.getByRole('button', { name: '深色' }),
      page.getByRole('group', { name: '语言' }),
      page.getByLabel('名称框'),
      page.getByLabel('编辑栏'),
      page.getByRole('tablist', { name: '工作表' }),
      page.getByRole('contentinfo').getByText('就绪'),
    ];
    for (const each of chinese) await expect(each).toBeVisible();
    await page.getByRole('button', { name: '深色' }).hover();
    await expect(page.getByRole('tooltip')).toHaveText('使用深色主题');

    await page.getByRole('button', { name: 'English' }).click();
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
    const english = [
      page.getByRole('tablist', { name: 'Ribbon' }),
      page.getByRole('tab', { name: 'View' }),
      page.getByRole('group', { name: 'Theme' }),
      page.getByRole('button', { name: 'Light' }),
      page.getByRole('button', { name: 'Dark' }),
      page.getByRole('group', { name: 'Language' }),
      nameBox(page),
      formulaBar(page),
      page.getByRole('tablist', { name: 'Sheets' }),
      page.getByRole('contentinfo').getByText('Ready'),
    ];
    for (const each of english) await expect(each).toBeVisible();
    for (const each of chinese) await expect(each).toHaveCount(0);

    await page.reload();
    await expect(page.getByRole('tab', { name: 'View' })).toBeVisible();
    await page.getByRole('button', { name: '中文' }).click();
    await expect(page.getByRole('tab', { name: '视图' })).toBeVisible();
  });
});

for (const system of ['light', 'dark'] as const) {
  const other = system === 'dark' ? 'Light' : 'Dark';

  test.describe(`on a system set to the ${system} theme`, () => {
    test.use({ colorScheme: system });

    test('the interface starts in that theme, switches, and remembers', async ({
      page,
    }) => {
      // Polled, because the colors of the controls change in a transition.
      const expectTheme = async (dark: boolean) => {
        const ribbon = page.getByRole('tabpanel', { name: 'View' });
        const status = page.getByRole('contentinfo');
        const button = page.getByRole('button', { name: 'English' });
        await expect
          .poll(async () => {
            const [red, green, blue] = await cellBackground(page, 3, 3);
            return [
              (red! + green! + blue!) / 3 < 128,
              await isDark(ribbon, 'background-color'),
              await isDark(status, 'background-color'),
              await isDark(formulaBar(page), 'background-color'),
              !(await isDark(ribbon, 'color')),
              !(await isDark(button, 'color')),
            ];
          })
          .toEqual(Array(6).fill(dark));
      };

      await openApp(page);
      await expectTheme(system === 'dark');
      await expect(page.getByRole('button', { name: other })).toHaveAttribute(
        'aria-pressed',
        'false',
      );

      await page.getByRole('button', { name: other }).click();
      await expect(page.getByRole('button', { name: other })).toHaveAttribute(
        'aria-pressed',
        'true',
      );
      await expectTheme(system !== 'dark');

      await page.reload();
      await expect(page.locator('canvas')).toBeVisible();
      await expectTheme(system !== 'dark');
    });
  });
}

test.describe('the controls', () => {
  test.use({ locale: 'en-US', colorScheme: 'light' });

  test('each have a tooltip', async ({ page }) => {
    await openApp(page);
    const tooltips = [
      ['tab', 'View', 'Appearance and language'],
      ['button', 'Light', 'Use the light theme'],
      ['button', 'Dark', 'Use the dark theme'],
      ['button', 'English', 'Show the interface in English'],
      ['button', '中文', 'Show the interface in Simplified Chinese'],
    ] as const;
    for (const [role, name, tooltip] of tooltips) {
      await page.getByRole(role, { name }).hover();
      await expect(page.getByRole('tooltip')).toHaveText(tooltip);
    }
    // No control is left out of the list above.
    await expect(page.getByRole('button')).toHaveCount(4);
    await expect(
      page.getByRole('tablist', { name: 'Ribbon' }).getByRole('tab'),
    ).toHaveCount(1);
  });

  test('are rounded and show hover and keyboard focus with a short transition', async ({
    page,
    browserName,
  }) => {
    await openApp(page);
    const button = page.getByRole('button', { name: 'Dark' });
    const style = (property: string) =>
      button.evaluate(
        (node, property) => getComputedStyle(node).getPropertyValue(property),
        property,
      );

    expect(parseFloat(await style('border-radius'))).toBeGreaterThan(0);
    const duration = parseFloat(await style('transition-duration'));
    expect(duration).toBeGreaterThan(0);
    expect(duration).toBeLessThanOrEqual(0.2);

    const resting = await style('background-color');
    await button.hover();
    await expect.poll(() => style('background-color')).not.toBe(resting);

    // Focus is outlined in the accent color when it comes from the keyboard.
    expect(await style('outline-style')).toBe('none');
    await page.getByRole('tab', { name: 'View' }).focus();
    // Safari reaches buttons with Option+Tab; Tab alone skips them.
    const tab = browserName === 'webkit' ? 'Alt+Tab' : 'Tab';
    for (let presses = 0; presses < 3; presses++) {
      await page.keyboard.press(tab); // The panel, Light, then Dark.
    }
    await expect(button).toBeFocused();
    await expect.poll(() => style('outline-style')).toBe('solid');
    const accent = await page
      .getByRole('button', { name: 'Light' })
      .evaluate((node) => getComputedStyle(node).color);
    expect(await style('outline-color')).toBe(accent);
  });
});
