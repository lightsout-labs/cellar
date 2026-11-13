import react from '@vitejs/plugin-react';
import { defineConfig } from 'vitest/config';

export default defineConfig({
  base: '/cellar/',
  plugins: [react()],
  preview: { port: 4173, strictPort: true },
  test: {
    environment: 'node',
    include: ['src/**/*.test.ts', 'tests/*.test.ts'],
    coverage: {
      provider: 'v8',
      // The engine, and the parts of the app that have no DOM: DOM code
      // lives in `.tsx` files and in the web platform, and is covered by
      // the end-to-end tests.
      include: ['src/engine/**/*.ts', 'src/app/**/*.ts'],
      exclude: ['src/app/platform/web/**', '**/*.test.ts', '**/*.d.ts'],
      thresholds: {
        'src/engine/**': { lines: 90 },
        'src/app/**': { lines: 90 },
      },
    },
  },
});
