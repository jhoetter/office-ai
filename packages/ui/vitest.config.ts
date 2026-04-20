import { defineConfig } from "vitest/config";

/**
 * Most UI primitives don't ship behaviour worth unit-testing in
 * isolation (a button is a button); the consumers' own tests
 * already cover wiring. The exception is anything that owns
 * input handling — `resize-handles.tsx` reports a synthesised
 * mousedown payload that ALL editors depend on, so we pin it
 * here. We opt that single suite into jsdom via the
 * `// @vitest-environment jsdom` docblock to keep the default
 * environment cheap.
 */
export default defineConfig({
  test: {
    environment: "node",
    include: ["src/**/*.test.{ts,tsx}"],
    passWithNoTests: true,
  },
});
