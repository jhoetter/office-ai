# CI / quality pipeline — one-pager

Operational reference for the CI/CD pipeline: what runs where, what's
fast vs. heavy, what's advisory vs. blocking, and the local commands
that mirror each gate.

The narrative of _why_ each gate exists lives in
[`build-log/quality-gates.md`](build-log/quality-gates.md). This page
is the cheat-sheet you reach for when something fails or you want to
run a specific check.

## Two-tier model

The pipeline splits into two tiers so a slow heavy gate never blocks
the inner-loop quality signal.

| Tier             | What it covers                                             | Local entry point            | CI job(s)                                                                             |
| ---------------- | ---------------------------------------------------------- | ---------------------------- | ------------------------------------------------------------------------------------- |
| **Quality gate** | format / lint / arch / typecheck / test / build            | `make verify` (or `make ci`) | `verify`                                                                              |
| **Heavy gates**  | LibreOffice roundtrip, perf budgets, OOXML XSD, Playwright | `make heavy`, `make e2e-web` | `*-libreoffice-roundtrip`, `*-perf`, `*-schema-validation`, `web-e2e`, `license-scan` |

`make heavy` runs every per-format heavy gate locally
(roundtrip + schema + perf for docx/xlsx/pptx). It's the local
equivalent of "what every CI heavy job runs", in one command.

## What `make verify` runs (in order, fail-fast)

1. `format-check` — Prettier (cheapest; catches whitespace noise).
2. `lint-root` — ESLint over `packages/**` + `tests/**`.
3. `lint-web` — Next.js ESLint over `apps/web`. **Promoted to its own
   step** so a single unescaped JSX entity fails in seconds rather than
   cascading into the slow build phase. (`next build` itself has
   `eslint.ignoreDuringBuilds: true` for the same reason.)
4. `architecture` — package dep-graph guard.
5. `typecheck` — `tsc --noEmit` across the workspace.
6. `test` — every `*.test.ts(x)` via Vitest.
7. `build` — turbo build of every package + the Next.js web host.

If any step fails, the gate stops there. Total wall time: ~30–60 s on
a warm cache.

## What the heavy CI jobs run

12 jobs total. Format parity: every product (DOCX, XLSX, PPTX) has the
same three heavy gates wired up.

| Job                          | Make target                 | Format | Notes                                                                                                                       |
| ---------------------------- | --------------------------- | ------ | --------------------------------------------------------------------------------------------------------------------------- |
| `verify`                     | `make verify`               | -      | Quality gate above.                                                                                                         |
| `docx-libreoffice-roundtrip` | `make roundtrip-libre-docx` | docx   | Roundtrips every fixture through `soffice`, asserts no repair text.                                                         |
| `xlsx-libreoffice-roundtrip` | `make roundtrip-libre-xlsx` | xlsx   | Same, with `libreoffice-calc`.                                                                                              |
| `pptx-libreoffice-roundtrip` | `make roundtrip-libre-pptx` | pptx   | Same, with `libreoffice-impress`. Skips fixtures whose input PDF is dirty before our roundtrip (e.g. corrupt embedded PNG). |
| `web-e2e`                    | `make e2e-web`              | -      | Playwright against `next start`.                                                                                            |
| `docx-perf`                  | `make perf-docx`            | docx   | Synthetic 100-page DOCX; parse + 1k commands + serialize budgets.                                                           |
| `xlsx-perf`                  | `make perf-xlsx`            | xlsx   | Synthetic 5-sheet workbook (~60k cells); same phases.                                                                       |
| `pptx-perf`                  | `make perf-pptx`            | pptx   | Synthetic 100-slide deck; same phases.                                                                                      |
| `docx-schema-validation`     | `make schema-validate-docx` | docx   | Advisory. Validates input + agent re-emit against `wml.xsd`.                                                                |
| `xlsx-schema-validation`     | `make schema-validate-xlsx` | xlsx   | Advisory. Validates against `sml.xsd`.                                                                                      |
| `pptx-schema-validation`     | `make schema-validate-pptx` | pptx   | Advisory. Validates against `pml.xsd`.                                                                                      |
| `license-scan`               | `make licenses`             | -      | Hard-fails on AGPL/SSPL/BUSL etc.; warns on LGPL/GPL-with-exception.                                                        |

### Why the schema-validation jobs are advisory

ECMA-376 Transitional flags pre-existing extension attributes that the
upstream emitters (`docx`, `xlsx` SheetJS, `pptxgenjs`) write
verbatim, plus genuine serializer quirks we want to fix in follow-up
PRs (e.g. PPTX `buSzPct` raw-percent values). Keeping the jobs
advisory means a fresh PR isn't blocked on a known issue, but the
artifact still surfaces every violation as a precise hit list.

## Skip semantics

Every heavy script is designed to run unattended on dev machines that
may not have `soffice` / `xmllint` installed:

- **Missing binary** → exit 0 with a `⚠ ...` warning. Local devs see
  it; CI fails the install step earlier so it never silently skips
  there.
- **Missing fixtures** → exit 0 with the same shape of warning.
- **Missing XSDs** → exit 0; `make xsd-fetch` (or the cached CI step)
  populates them.
- **Dirty input fixture** (LibreOffice roundtrip only) → marked as
  `skip` in the table, not a failure. The gate is for regressions
  _we_ introduce, not for pre-existing PNG corruption in real-world
  fixtures.

## Local commands cheat-sheet

| Goal                                            | Command                      |
| ----------------------------------------------- | ---------------------------- |
| Mirror what CI's `verify` job runs              | `make verify`                |
| Lighter pre-push gate (skip arch/build/format)  | `make precommit`             |
| Run every heavy gate (3 formats × 3 gate types) | `make heavy`                 |
| Run a single heavy gate for one product         | `make perf-xlsx` (etc.)      |
| Just run one product's tests                    | `make test-docx` (etc.)      |
| Run Playwright suite                            | `make e2e-web`               |
| Print a holistic snapshot of the repo           | `make metrics`               |
| Auto-format every file (Prettier `--write`)     | `make format`                |
| Re-fetch the OOXML XSD bundle                   | `make xsd-fetch`             |
| Regenerate fixtures                             | `make fixtures` (or per fmt) |
| List every target with a one-line description   | `make help`                  |

## CI-only operational notes

- **Concurrency.** `concurrency.cancel-in-progress: true` per ref;
  superseded commits are cancelled mid-run.
- **Caching.** `actions/setup-node@v4` with `cache: pnpm` for every
  job. The OOXML XSD bundle is additionally cached under
  `vendor/ooxml-xsd` keyed on the SHA of `scripts/fetch-ooxml-xsd.mjs`
  — a script change invalidates the cache automatically.
- **Build artifact.** The Playwright HTML report is uploaded on every
  `web-e2e` run (always, including on success).
- **System packages installed in CI.**
  - `libreoffice-{core,writer,calc,impress}` for the three roundtrip
    jobs (split per-job so the matrix keeps install time honest).
  - `libxml2-utils` for `xmllint` in the three schema-validation jobs.

## Adding a new gate

1. Write the script under `scripts/<gate>-<format>.mjs`. Make it skip
   gracefully when its system dependency is missing.
2. Add a `make <gate>-<format>` target in the **Heavy / opt-in
   checks** block of [`Makefile`](../Makefile). Add it to the
   `<gate>-all` aggregator and `make heavy`.
3. Add a CI job in [`.github/workflows/ci.yml`](../.github/workflows/ci.yml)
   following the per-format pattern (checkout → setup pnpm → install
   → install system deps → build → run). Mark `continue-on-error: true`
   if the gate is advisory until follow-up work lands.
4. Add a row to the table above and (if the gate is non-trivial) a
   note in [`build-log/quality-gates.md`](build-log/quality-gates.md).
