import { ESLint } from 'eslint';
import { describe, expect, it } from 'vitest';

const eslint = new ESLint();

async function problems(code: string, filePath: string): Promise<string[]> {
  const [result] = await eslint.lintText(code, { filePath });
  return result!.messages.map((message) => message.ruleId ?? message.message);
}

const platformApis = [
  ['browser storage', 'indexedDB.open("cellar");'],
  ['browser storage', 'localStorage.setItem("theme", "dark");'],
  ['browser storage', 'window.localStorage.getItem("theme");'],
  ['browser storage', 'void navigator.storage.persist();'],
  ['the clipboard', 'void navigator.clipboard.writeText("a");'],
  [
    'the clipboard',
    'export const read = (e: ClipboardEvent) => e.clipboardData;',
  ],
  ['a file', 'export const input = <input type="file" />;'],
  [
    'a file',
    'const input = document.createElement("input"); input.type = "file";',
  ],
  ['a file', 'void showOpenFilePicker();'],
  ['a file', 'export const reader = new FileReader();'],
  ['a file', 'export const link = <a href="blob:x" download="a.xlsx" />;'],
  [
    'a file',
    'const link = document.createElement("a"); link.download = "a.csv";',
  ],
  ['printing', 'window.print();'],
  ['printing', 'print();'],
] as const;

describe('module boundaries enforced by lint', () => {
  it.each(platformApis)(
    'fails on direct use of %s outside the web platform: %s',
    async (_api, code) => {
      for (const file of [
        'src/app/ui/Example.tsx',
        'src/engine/io/example.tsx',
      ]) {
        expect(await problems(code, file)).toEqual([
          expect.stringMatching(/^no-restricted-/),
        ]);
      }
    },
  );

  it.each(platformApis)(
    'allows %s inside the web platform: %s',
    async (_api, code) => {
      expect(await problems(code, 'src/app/platform/web/example.tsx')).toEqual(
        [],
      );
    },
  );

  it('fails when the engine imports the app or the UI library', async () => {
    const file = 'src/engine/workbook/example.ts';
    for (const code of ['import "../../app/platform";', 'import "react";']) {
      expect(await problems(code, file)).toEqual(['no-restricted-imports']);
    }
    expect(await problems('import "../formula";', file)).toEqual([]);
  });
});
