import { defineConfig } from "vitest/config";

/**
 * Unit tests for pure helpers that ship with the Next.js app.
 *
 * E2E lives under `e2e/` and runs with Playwright. Vitest is scoped to
 * `app/**` so it never accidentally picks up Playwright specs.
 */
export default defineConfig({
  test: {
    environment: "node",
    include: ["app/**/*.test.ts"],
    passWithNoTests: true,
  },
});
