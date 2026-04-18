// Root ESLint flat config for the entire monorepo.
//
// Goals (in priority order):
//   1. Catch obvious correctness bugs (unused vars, no-undef, etc).
//   2. Enforce architecture boundaries via import restrictions, in addition
//      to the dep-graph check in scripts/check-architecture.mjs.
//   3. Stay fast — no type-aware rules in the default lane so `make lint`
//      finishes in seconds even as the project grows.
//
// apps/web ships its own Next.js-flavoured config that extends this one.

import js from "@eslint/js";
import tsParser from "@typescript-eslint/parser";
import tsPlugin from "@typescript-eslint/eslint-plugin";
import importPlugin from "eslint-plugin-import";

const SHARED_IGNORES = [
  "**/node_modules/**",
  "**/dist/**",
  "**/.next/**",
  "**/.turbo/**",
  "**/coverage/**",
  "**/*.d.ts",
  "fixtures/**",
  "scripts/**",
];

/**
 * Architecture boundaries enforced at the import level.
 * The dep-graph check (scripts/check-architecture.mjs) backs this up at
 * package.json level — both layers must agree.
 */
const ARCH_RESTRICTED_IMPORTS = {
  "packages/core/**": [
    { group: ["@officeai/docx", "@officeai/docx/*"], message: "core must not depend on docx (it's a leaf)." },
    { group: ["@officeai/agent", "@officeai/agent/*"], message: "core must not depend on agent." },
    { group: ["@officeai/ui", "@officeai/ui/*"], message: "core must not depend on UI." },
    { group: ["@officeai/web", "@officeai/web/*"], message: "core must not depend on the web app." },
    { group: ["next", "next/*", "react", "react-dom"], message: "core is headless — no React/Next imports." },
  ],
  "packages/docx/**": [
    { group: ["@officeai/agent", "@officeai/agent/*"], message: "docx must not depend on agent." },
    { group: ["@officeai/ui", "@officeai/ui/*"], message: "docx must not depend on UI." },
    { group: ["@officeai/web", "@officeai/web/*"], message: "docx must not depend on the web app." },
    { group: ["next", "next/*"], message: "docx must not depend on Next.js." },
  ],
  "packages/agent/**": [
    { group: ["@officeai/ui", "@officeai/ui/*"], message: "agent must not depend on UI." },
    { group: ["@officeai/web", "@officeai/web/*"], message: "agent must not depend on the web app." },
    { group: ["next", "next/*"], message: "agent is a CLI / library — no Next.js." },
    { group: ["react", "react-dom"], message: "agent is headless — no React." },
  ],
  "packages/ui/**": [
    { group: ["@officeai/core", "@officeai/core/*"], message: "ui is presentation-only — no model deps." },
    { group: ["@officeai/docx", "@officeai/docx/*"], message: "ui is presentation-only — no docx deps." },
    { group: ["@officeai/agent", "@officeai/agent/*"], message: "ui is presentation-only — no agent deps." },
    { group: ["@officeai/web", "@officeai/web/*"], message: "ui must not depend on the host app." },
  ],
  "packages/design-tokens/**": [
    { group: ["@officeai/*"], message: "design-tokens is a leaf — no internal deps." },
    { group: ["react", "react-dom", "next", "next/*"], message: "design-tokens has no runtime." },
  ],
  "packages/text-formatting/**": [
    { group: ["@officeai/*"], message: "text-formatting is a leaf — no internal deps." },
    { group: ["react", "react-dom", "next", "next/*"], message: "text-formatting is headless — no React/Next imports." },
  ],
  "packages/comments/**": [
    { group: ["@officeai/*"], message: "comments is a leaf — no internal deps." },
    { group: ["react", "react-dom", "next", "next/*"], message: "comments is headless — no React/Next imports." },
  ],
};

/**
 * Helper: turn the arch map into one ESLint config block per file pattern.
 */
const archConfigs = Object.entries(ARCH_RESTRICTED_IMPORTS).map(([files, patterns]) => ({
  files: [`${files}/*.{ts,tsx,mjs,cjs,js}`, `${files}/**/*.{ts,tsx,mjs,cjs,js}`],
  rules: {
    "no-restricted-imports": ["error", { patterns }],
  },
}));

export default [
  { ignores: SHARED_IGNORES },

  js.configs.recommended,

  // Default TS rules for every TS file in packages + tests.
  {
    files: ["**/*.{ts,tsx,mts,cts}"],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: "latest",
        sourceType: "module",
        ecmaFeatures: { jsx: true },
      },
      globals: {
        // Node + browser globals available everywhere; specific files can
        // tighten if needed.
        console: "readonly",
        process: "readonly",
        Buffer: "readonly",
        setTimeout: "readonly",
        clearTimeout: "readonly",
        setInterval: "readonly",
        clearInterval: "readonly",
        URL: "readonly",
        URLSearchParams: "readonly",
        TextEncoder: "readonly",
        TextDecoder: "readonly",
        fetch: "readonly",
        crypto: "readonly",
        window: "readonly",
        document: "readonly",
        HTMLElement: "readonly",
        Element: "readonly",
        Event: "readonly",
        File: "readonly",
        Blob: "readonly",
        FormData: "readonly",
        navigator: "readonly",
        localStorage: "readonly",
        sessionStorage: "readonly",
        Node: "readonly",
        // DOM types that ProseMirror touches:
        MutationObserver: "readonly",
        IntersectionObserver: "readonly",
        ResizeObserver: "readonly",
      },
    },
    plugins: {
      "@typescript-eslint": tsPlugin,
      import: importPlugin,
    },
    rules: {
      // Correctness
      "no-unused-vars": "off", // delegated to TS rule below
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_", caughtErrorsIgnorePattern: "^_" },
      ],
      "no-undef": "off", // TS handles this; ESLint's no-undef is noisy with TS-only globals
      "no-empty": ["error", { allowEmptyCatch: true }],
      "no-constant-condition": ["error", { checkLoops: false }],
      "prefer-const": "error",

      // Architecture & hygiene
      "no-restricted-syntax": [
        "error",
        {
          selector: "ImportDeclaration[source.value=/^@officeai\\/[^/]+\\/src\\//]",
          message:
            "Import from a package's public entry (e.g. @officeai/docx) — never reach into its `src/` (that breaks the public API contract).",
        },
        {
          selector: "TSEnumDeclaration",
          message: "Use string-literal unions instead of TypeScript enums (better tree-shaking, no runtime emit).",
        },
      ],

      // Workspace rule "no inline imports": keep imports at top of file.
      "import/first": "error",
    },
  },

  // Per-package architecture restrictions (layered on top of the default).
  ...archConfigs,

  // Test files: relax a few rules.
  {
    files: ["**/*.test.{ts,tsx}", "**/test-utils/**", "tests/**"],
    rules: {
      "@typescript-eslint/no-unused-vars": "warn",
    },
  },
];
