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
    page.getByRole('list', { name: 'Sheets' }),
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

  await expect(
    page.getByRole('list', { name: 'Sheets' }).getByRole('listitem'),
  ).toHaveText(['Sheet1']);
  await expect(page.getByRole('contentinfo')).toContainText('Ready');
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
  await expect(page.getByRole('contentinfo')).toContainText('Enter');
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
      page.getByRole('list', { name: '工作表' }),
      page.getByRole('contentinfo').getByText('就绪'),
      page.getByRole('button', { name: '放大' }),
      page.getByRole('button', { name: '缩小' }),
      page.getByRole('slider', { name: '缩放' }),
    ];
    for (const each of chinese) await expect(each).toBeVisible();
    await page.getByRole('button', { name: '深色' }).hover();
    await expect(page.getByRole('tooltip')).toHaveText('使用深色主题');
    // The summary of a selection is in Chinese too.
    await page.locator('canvas').click({ position: { x: 4, y: 4 } });
    await page.locator('canvas').click({ position: { x: 100, y: 36 } });
    await page.keyboard.type('5');
    await page.keyboard.press('Enter');
    await page.keyboard.press('Shift+ArrowUp');
    await expect(page.getByLabel('选定区域的汇总')).toHaveText(
      '平均值: 5计数: 1求和: 5',
    );

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
      page.getByRole('list', { name: 'Sheets' }),
      page.getByRole('contentinfo').getByText('Ready'),
      page.getByRole('button', { name: 'Zoom in' }),
      page.getByRole('button', { name: 'Zoom out' }),
      page.getByRole('slider', { name: 'Zoom' }),
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
      ['button', 'Zoom out', 'Zoom out by 10%'],
      ['button', 'Zoom in', 'Zoom in by 10%'],
      ['slider', 'Zoom', 'Size of the grid'],
    ] as const;
    for (const [role, name, tooltip] of tooltips) {
      await page.getByRole(role, { name }).hover();
      await expect(page.getByRole('tooltip')).toHaveText(tooltip);
    }
    // No control is left out of the list above.
    await expect(page.getByRole('button')).toHaveCount(6);
    await expect(page.getByRole('slider')).toHaveCount(1);
    await expect(
      page.getByRole('tablist', { name: 'Ribbon' }).getByRole('tab'),
    ).toHaveCount(1);
  });

  test('are rounded and show hover, keyboard focus and disabled states with a short transition', async ({
    page,
    browserName,
  }) => {
    await openApp(page);
    const style = (control: Locator, property: string) =>
      control.evaluate(
        (node, property) => getComputedStyle(node).getPropertyValue(property),
        property,
      );
    const tab = page.getByRole('tab', { name: 'View' });
    const button = page.getByRole('button', { name: 'Dark' });
    const accent = await style(
      page.getByRole('button', { name: 'Light' }),
      'color',
    );

    await page.mouse.move(500, 400); // Over the grid, away from the controls.
    for (const control of [tab, button]) {
      expect(parseFloat(await style(control, 'border-radius'))).toBeGreaterThan(
        0,
      );
      const duration = parseFloat(await style(control, 'transition-duration'));
      expect(duration).toBeGreaterThan(0);
      expect(duration).toBeLessThanOrEqual(0.2);

      const resting = await style(control, 'background-color');
      await control.hover();
      await expect
        .poll(() => style(control, 'background-color'))
        .not.toBe(resting);
      expect(await style(control, 'outline-style')).toBe('none');
    }

    // Focus is outlined in the accent color when it comes from the keyboard.
    // Safari reaches buttons with Option+Tab; Tab alone skips them.
    const key = browserName === 'webkit' ? 'Alt+Tab' : 'Tab';
    await tab.focus();
    for (let presses = 0; presses < 3; presses++) {
      await page.keyboard.press(key); // The panel, Light, then Dark.
    }
    await expect(button).toBeFocused();
    for (let presses = 0; presses < 3; presses++) {
      expect(await style(page.locator(':focus'), 'outline-style')).toBe(
        'solid',
      );
      expect(await style(page.locator(':focus'), 'outline-color')).toBe(accent);
      await page.keyboard.press(key.replace('Tab', 'Shift+Tab'));
    }
    await expect(tab).toBeFocused();
    expect(await style(tab, 'outline-style')).toBe('solid');

    // No control is disabled yet, so one is disabled here to see the state:
    // faded, with the plain cursor, and unchanged under the pointer.
    await page.mouse.move(500, 400);
    await button.evaluate((node: HTMLButtonElement) => (node.disabled = true));
    await expect.poll(() => style(button, 'opacity')).toBe('0.5');
    expect(await style(button, 'cursor')).toBe('default');
    await button.hover({ force: true });
    await page.waitForTimeout(300); // Longer than the transition.
    expect(await style(button, 'background-color')).toBe('rgba(0, 0, 0, 0)');
  });
});
