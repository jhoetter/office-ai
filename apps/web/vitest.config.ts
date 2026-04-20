import { defineConfig } from "vitest/config";
import { fileURLToPath } from "node:url";

/**
 * Unit tests for pure helpers that ship with the Next.js app.
 *
 * E2E lives under `e2e/` and runs with Playwright. Vitest is scoped to
 * `app/**` so it never accidentally picks up Playwright specs.
 */
export default defineConfig({
  resolve: {
    // Mirror `tsconfig.json#compilerOptions.paths` for `@/*` so
    // component-rendering tests (`*.test.tsx`) can import shared
    // libs the same way runtime code does.
    alias: {
      "@": fileURLToPath(new URL("./app", import.meta.url)),
    },
  },
  // The app's tsconfig uses `jsx: "preserve"` so Next can transform
  // it. Vitest needs the automatic runtime so JSX in tests doesn't
  // require an explicit `import * as React from "react"`.
  esbuild: {
    jsx: "automatic",
  },
  test: {
    environment: "node",
    // `.test.tsx` files render React components and opt in to
    // jsdom via the `// @vitest-environment jsdom` docblock.
    include: ["app/**/*.test.{ts,tsx}"],
    passWithNoTests: true,
  },
});
