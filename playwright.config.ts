import { defineConfig } from '@playwright/test';
export default defineConfig({
  testDir: './test/browser',
  workers: 1,
  timeout: 45000,
  use: { baseURL: 'http://127.0.0.1:4187', headless: true, acceptDownloads: true,
    launchOptions: process.env.PLAYWRIGHT_EXECUTABLE_PATH ? { executablePath: process.env.PLAYWRIGHT_EXECUTABLE_PATH } : {} },
  webServer: { command: 'node dist/src/server.js', url: 'http://127.0.0.1:4187', env: { LEDGER_CHECK_PORT: '4187' }, reuseExistingServer: false },
  reporter: 'list',
});
