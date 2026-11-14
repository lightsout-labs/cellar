import { defineConfig, devices } from '@playwright/test';

// GRIDWORKS_URL points the tests at a deployed app instead of a local build.
const deployed = process.env.GRIDWORKS_URL;

export default defineConfig({
  testDir: 'tests/e2e',
  fullyParallel: true,
  forbidOnly: Boolean(process.env.CI),
  retries: process.env.CI ? 1 : 0,
  reporter: process.env.CI ? 'github' : 'list',
  use: {
    baseURL: deployed ?? 'http://localhost:4173/gridworks/',
    trace: 'retain-on-failure',
  },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    { name: 'firefox', use: { ...devices['Desktop Firefox'] } },
    { name: 'webkit', use: { ...devices['Desktop Safari'] } },
  ],
  webServer: deployed
    ? undefined
    : {
        command: 'npm run build && npm run preview',
        url: 'http://localhost:4173/gridworks/',
        reuseExistingServer: false,
      },
});
