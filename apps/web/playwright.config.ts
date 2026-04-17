import { defineConfig, devices } from "@playwright/test";

/**
 * Playwright config for the @officeai/web smoke suite (W1, batch P1.1).
 *
 * The specs exercise the in-browser editor against the bundled sample
 * DOCX (apps/web/app/lib/sample-docx.ts) so they don't depend on the
 * real-world fixture corpus. They are not part of `make verify` — they
 * run as a dedicated `make e2e-web` target and a dedicated CI job
 * (`web-e2e`) so a developer machine without Playwright browsers
 * installed can still ship.
 *
 * Run locally:
 *   pnpm --filter @officeai/web e2e:install    # one-off browser install
 *   pnpm --filter @officeai/web build          # compile workspace deps
 *   pnpm --filter @officeai/web e2e
 */
const PORT = Number(process.env.E2E_PORT ?? 3100);
const baseURL = process.env.E2E_BASE_URL ?? `http://localhost:${PORT}`;
const isCI = !!process.env.CI;

export default defineConfig({
  testDir: "./e2e",
  testMatch: ["**/*.spec.ts"],
  fullyParallel: false,
  forbidOnly: isCI,
  retries: isCI ? 1 : 0,
  workers: 1,
  reporter: isCI ? [["list"], ["html", { open: "never" }]] : "list",
  timeout: 30_000,
  expect: { timeout: 7_500 },
  use: {
    baseURL,
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
    video: "off",
    viewport: { width: 1280, height: 800 },
  },
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },
  ],
  webServer: process.env.E2E_BASE_URL
    ? undefined
    : {
        // `next start` requires `next build` to have run; the Makefile
        // target does that before invoking Playwright. We bind to a
        // dedicated port so a dev-server on :3000 doesn't conflict.
        command: `pnpm exec next start --port ${PORT}`,
        port: PORT,
        reuseExistingServer: !isCI,
        timeout: 120_000,
        stdout: "pipe",
        stderr: "pipe",
      },
});
