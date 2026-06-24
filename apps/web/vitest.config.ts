import path from "node:path";
import { defineConfig } from "vitest/config";

/**
 * Unit tests for pure helpers that ship with the Next.js app.
 *
 * E2E lives under `e2e/` and runs with Playwright. Vitest is scoped to
 * `app/**` so it never accidentally picks up Playwright specs.
 *
 * The `@/*` alias mirrors the one in `tsconfig.json` so non-trivial
 * helpers (e.g. the comments provider, which imports
 * `@/lib/format-helpers`) are loadable from the test runner.
 */
export default defineConfig({
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "app"),
      "@officeai/agent/projections": path.resolve(__dirname, "../../packages/agent/src/projections.ts"),
      "@officeai/agent/session-store": path.resolve(__dirname, "../../packages/agent/src/session-store.ts"),
    },
  },
  test: {
    environment: "node",
    include: ["app/**/*.test.ts"],
    passWithNoTests: true,
  },
});
