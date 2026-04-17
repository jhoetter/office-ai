# Quality gates

Date: 2026-04-17

## Why now

The repo grew an editor surface, a parser/serializer pair, an agent
CLI, and a web host without any CI. Easy to slip in cross-package
imports that erode the architecture (e.g. `core` reaching into `docx`)
or land typo-level bugs in PRs. This change adds a single canonical
quality gate (`make verify`) that runs everything an early CI should
care about, plus a GitHub Actions workflow that runs the same pipeline.

## What runs

In order, fail-fast:

1. **`format-check`** — Prettier with the root `.prettierrc.json`. The
   complement of `make format`. Cheapest gate; catches accidental
   noise so reviewers don't have to.
2. **`lint`** — ESLint flat config at the repo root
   (`eslint.config.mjs`). Default rules: `@typescript-eslint`
   recommended subset, `prefer-const`, `no-unused-vars`. Architectural
   rules:
   - `no-restricted-syntax`: bans deep imports
     (`@officeai/<pkg>/src/...`) — must use the public entry — and
     bans TS enums (use string-literal unions).
   - `import/first`: enforces top-of-file imports (workspace rule
     "no inline imports").
   - Per-package `no-restricted-imports` patterns enforce the dep
     graph at import-statement level (e.g. nothing in `packages/core/`
     may import from `@officeai/docx` or React).
     `apps/web` retains its Next.js ESLint config; everything else uses
     the root config.
3. **`architecture`** — `scripts/check-architecture.mjs` validates the
   `package.json` dependency graph against a declarative allow-list:
   ```
   core             → ∅
   design-tokens    → ∅
   ui               → design-tokens
   docx             → core
   agent            → core, docx
   integration-tests → core, docx, agent
   web              → core, docx, agent, ui, design-tokens
   ```
   It also bans `react` / `react-dom` / `next` from headless packages.
   Both layers (this manifest check + the ESLint `no-restricted-imports`
   in step 2) must agree — defense in depth.
4. **`typecheck`** — `turbo typecheck` (per-package `tsc --noEmit`).
5. **`test`** — `turbo test` (Vitest, all packages).
6. **`build`** — `turbo build` to catch build-time integration issues.

## Files added / changed

| File                             | Purpose                                                            |
| -------------------------------- | ------------------------------------------------------------------ |
| `.prettierrc.json`               | Single source of truth for formatting                              |
| `.prettierignore`                | Excludes `dist`, fixtures, lock files                              |
| `eslint.config.mjs`              | Root ESLint flat config + per-package architecture rules           |
| `scripts/check-architecture.mjs` | Manifest-level dep-graph linter                                    |
| `Makefile`                       | Adds `verify`, `ci`, `format-check`, `architecture`, `lint`        |
| `package.json`                   | New scripts: `format:check`, `architecture`, `lint:root`, `verify` |
| `packages/*/package.json`        | Stub `lint` scripts replaced with real `eslint ...` calls          |
| `.github/workflows/ci.yml`       | Runs `make verify` on push to main + every PR                      |

## What surfaced

Running ESLint for the first time caught real issues already in `main`:

- `packages/docx/src/parser/parse.ts` had an unused `CommentReference` type import.
- `packages/docx/src/renderer/mount.ts` had a `let` that was never reassigned (`prefer-const`).

Both fixed in the same commit. Prettier found 62 unformatted files;
`make format` cleaned them in one pass.

## Local + CI parity

Both `make verify` (local) and `.github/workflows/ci.yml` (CI) run the
same target. If the gate is green locally, CI is green — no surprise
deltas.

## Deliberately deferred

- **Coverage thresholds.** `@vitest/coverage-v8` is easy to add but
  enforcing per-package thresholds early creates more noise than
  signal. A `make coverage` target can be added when there's a real
  reason to look at uncovered lines.
- **Dead-code detection (`knip` / `ts-prune`).** Useful but
  opinionated; would generate a stream of false positives against the
  current code while patterns are still being established.
- **Mutation testing.** Out of scope for early-stage.
- **Branch protection on `main`.** A GitHub UI setting; not in code.
  Once the workflow has run once, switch on "Require status checks to
  pass before merging → CI / verify".
