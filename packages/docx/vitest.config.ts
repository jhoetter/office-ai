import { defineConfig } from "vitest/config";

/**
 * Most of the docx tests are pure model / parser / serializer
 * checks that run happily in plain Node. The renderer mount tests
 * (`renderer/mount.test.ts`) need a DOM because ProseMirror's
 * `EditorView` calls into `document` / `window`. We opt those
 * specific files in via the `// @vitest-environment jsdom`
 * docblock at the top of the file rather than paying the jsdom
 * startup cost for every other suite.
 */
export default defineConfig({
  test: {
    environment: "node",
    environmentMatchGlobs: [["src/renderer/mount.test.ts", "jsdom"]],
    include: ["src/**/*.test.ts"],
  },
});
