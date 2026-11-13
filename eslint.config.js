import js from '@eslint/js';
import tseslint from 'typescript-eslint';

const viaPlatform =
  'Reach files, storage, the clipboard and printing through the Platform interface (src/app/platform).';

// Globals that only src/app/platform/web may use (R51).
const platformGlobals = [
  'indexedDB',
  'localStorage',
  'sessionStorage',
  'showOpenFilePicker',
  'showSaveFilePicker',
  'print',
  'caches',
  'FileReader',
  'ClipboardItem',
];

const platformBoundary = {
  'no-restricted-globals': [
    'error',
    ...platformGlobals.map((name) => ({ name, message: viaPlatform })),
  ],
  'no-restricted-properties': [
    'error',
    ...['window', 'globalThis', 'self'].flatMap((object) =>
      platformGlobals.map((property) => ({
        object,
        property,
        message: viaPlatform,
      })),
    ),
    { property: 'clipboard', message: viaPlatform },
    { object: 'navigator', property: 'storage', message: viaPlatform },
    { object: 'document', property: 'cookie', message: viaPlatform },
    { object: 'document', property: 'execCommand', message: viaPlatform },
    { property: 'clipboardData', message: viaPlatform },
  ],
  'no-restricted-syntax': [
    'error',
    {
      // <input type="file">
      selector:
        "JSXOpeningElement[name.name='input'] > JSXAttribute[name.name='type'][value.value='file']",
      message: viaPlatform,
    },
    {
      // input.type = 'file'
      selector:
        "AssignmentExpression[left.property.name='type'][right.value='file']",
      message: viaPlatform,
    },
    {
      // <a download>
      selector: "JSXAttribute[name.name='download']",
      message: viaPlatform,
    },
    {
      // link.download = name
      selector: "AssignmentExpression[left.property.name='download']",
      message: viaPlatform,
    },
  ],
};

export default tseslint.config(
  { ignores: ['dist', 'coverage', 'test-results', 'playwright-report'] },
  js.configs.recommended,
  tseslint.configs.recommended,
  { rules: platformBoundary },
  {
    // Only the entry point chooses the implementation of Platform.
    files: ['src/app/**', 'src/engine/**'],
    ignores: ['src/app/platform/web/**'],
    rules: {
      'no-restricted-imports': [
        'error',
        { patterns: [{ regex: 'platform/web(/|$)', message: viaPlatform }] },
      ],
    },
  },
  {
    files: ['src/app/platform/web/**'],
    rules: {
      'no-restricted-globals': 'off',
      'no-restricted-properties': 'off',
      'no-restricted-syntax': 'off',
    },
  },
  {
    // The engine has no DOM and does not know the app (R50).
    files: ['src/engine/**'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          patterns: [
            { regex: 'platform/web(/|$)', message: viaPlatform },
            {
              regex: '(^|/)app(/|$)',
              message: 'The engine must not import from src/app.',
            },
            {
              group: ['react', 'react-dom', 'react-dom/*'],
              message: 'The engine must not depend on the UI.',
            },
          ],
        },
      ],
    },
  },
);
