# ============================================
# officeAI — convenience targets
#
# `make verify` is the canonical quality gate: identical to what CI
# runs. Pass it locally before pushing.
# ============================================

.PHONY: help install dev build lint lint-root lint-web format format-check architecture \
        typecheck test test-docx test-xlsx test-pptx test-core test-web verify ci precommit \
        clean cli fixtures fixtures-real fixtures-xlsx fixtures-pptx fixtures-pptx-real \
        roundtrip-libre roundtrip-libre-docx roundtrip-libre-xlsx roundtrip-libre-pptx \
        roundtrip-libre-all \
        e2e-web perf perf-docx perf-xlsx perf-pptx perf-all \
        licenses metrics heavy \
        xsd-fetch schema-validate schema-validate-docx schema-validate-xlsx \
        schema-validate-pptx schema-validate-all

help:
	@echo "officeAI — available targets (see docs/ci-pipeline.md for the full map)"
	@echo ""
	@echo "  install        Install all dependencies"
	@echo "  dev            Start the Next.js editor host"
	@echo "  build          Build all packages"
	@echo ""
	@echo "  Quality gates  (== what CI runs in the 'verify' job):"
	@echo "  format         Auto-format all files (prettier --write)"
	@echo "  format-check   Check formatting (CI-safe, no writes)"
	@echo "  lint           Lint root + apps/web (== lint-root + lint-web)"
	@echo "  architecture   Validate package dep graph (separation of concerns)"
	@echo "  typecheck      Typecheck all packages"
	@echo "  test           Run every test in the workspace (turbo)"
	@echo "  test-docx      Run only @officeai/docx tests"
	@echo "  test-xlsx      Run only @officeai/xlsx tests"
	@echo "  test-pptx      Run only @officeai/pptx tests"
	@echo "  test-core      Run only @officeai/core tests"
	@echo "  test-web       Run only @officeai/web tests (vitest)"
	@echo "  verify         Run the full quality gate (== what CI runs)"
	@echo "  ci             Alias for verify"
	@echo "  precommit      Lighter gate for fast local pre-push (format + lint + typecheck + test)"
	@echo ""
	@echo "  Misc:"
	@echo "  clean          Remove build artifacts and dependencies"
	@echo "  cli            Build and link the office-agent CLI"
	@echo "  fixtures       Regenerate synthetic DOCX + XLSX + PPTX fixtures"
	@echo "  fixtures-real  Regenerate real-shape DOCX fixtures (Word-grade)"
	@echo "  fixtures-xlsx  Regenerate synthetic XLSX fixtures only"
	@echo "  fixtures-pptx  Regenerate synthetic PPTX fixtures only"
	@echo "  fixtures-pptx-real  Regenerate third-party-emitter PPTX fixtures"
	@echo "  metrics        Print a holistic snapshot (LOC, packages, tests, fixtures, CI jobs)"
	@echo ""
	@echo "  Heavy / opt-in checks (NOT part of make verify; CI runs them in dedicated jobs):"
	@echo "  heavy                    Run every heavy gate (roundtrip + schema + perf for all 3 formats)"
	@echo "  roundtrip-libre          Alias for roundtrip-libre-docx (back-compat)"
	@echo "  roundtrip-libre-docx     Headless LibreOffice roundtrip on DOCX fixtures"
	@echo "  roundtrip-libre-xlsx     Headless LibreOffice roundtrip on XLSX fixtures"
	@echo "  roundtrip-libre-pptx     Headless LibreOffice roundtrip on PPTX fixtures"
	@echo "  roundtrip-libre-all      Run the LibreOffice roundtrip across all three formats"
	@echo "  e2e-web                  Playwright smoke tests against the web app"
	@echo "  perf                     Run every perf budget (docx + xlsx + pptx)"
	@echo "  perf-docx                DOCX perf budgets (parse / 1k commands / serialize)"
	@echo "  perf-xlsx                XLSX perf budgets (parse / 1k commands / serialize)"
	@echo "  perf-pptx                PPTX perf budgets (parse / 1k commands / serialize)"
	@echo "  perf-all                 Alias for perf"
	@echo "  licenses                 SPDX license-graph scan against the resolved deps"
	@echo "  xsd-fetch                Download the ECMA-376 OOXML XSDs into vendor/ooxml-xsd/"
	@echo "  schema-validate          Alias for schema-validate-docx (back-compat)"
	@echo "  schema-validate-docx     Validate DOCX fixtures (input + agent re-emit) against wml.xsd"
	@echo "  schema-validate-xlsx     Validate XLSX fixtures (input + agent re-emit) against sml.xsd"
	@echo "  schema-validate-pptx     Validate PPTX fixtures (input + agent re-emit) against pml.xsd"
	@echo "  schema-validate-all      Run XSD validation across all three formats"

install:
	pnpm install

dev:
	pnpm --filter @officeai/web dev

build:
	pnpm build

lint: lint-root lint-web

lint-root:
	pnpm lint:root

# `lint-web` runs Next.js's ESLint config against apps/web. Promoted to a
# dedicated, fail-fast step in `verify` so a single unescaped JSX entity
# (or any other apps/web lint regression) surfaces in ~5s instead of
# cascading into every CI job that touches `pnpm build`. See
# docs/ci-pipeline.md for the rationale.
lint-web:
	pnpm --filter @officeai/web lint

format:
	pnpm format

format-check:
	pnpm format:check

architecture:
	pnpm architecture

typecheck:
	pnpm typecheck

test:
	pnpm test

# ── Per-package test entrypoints ─────────────────────────────────────
# Useful while iterating on a single product (or for a CI matrix split).
# Each delegates to turbo so dependent packages still build first.
test-docx:
	pnpm --filter @officeai/docx test

test-xlsx:
	pnpm --filter @officeai/xlsx test

test-pptx:
	pnpm --filter @officeai/pptx test

test-core:
	pnpm --filter @officeai/core test

test-web:
	pnpm --filter @officeai/web test

# ── The quality gate ──────────────────────────────────────────────────
# Order is intentional and fail-fast:
#   1. format-check  — cheapest, catches accidental noise
#   2. lint-root     — packages/** + tests/** lint (correctness + arch boundaries)
#   3. lint-web      — apps/web lint (Next.js config). Split out so an apps/web
#                       regression fails in seconds, before the slower steps.
#   4. architecture  — catches package-level dep-graph violations
#   5. typecheck     — catches type-system violations
#   6. test          — catches behavioural regressions
#   7. build         — catches build-time/integration issues
verify: format-check lint-root lint-web architecture typecheck test build
	@echo ""
	@echo "✅ verify: all quality gates passed."

ci: verify

# `precommit` is a shorter, fast-feedback gate intended for a local
# pre-push hook. It deliberately drops `format-check` (devs typically
# auto-format on save), `architecture` (slow, rarely violated mid-edit),
# and `build` (turbo will catch it on the verify pass). Use `make verify`
# before opening a PR; CI runs `verify` regardless.
precommit: lint-root lint-web typecheck test
	@echo ""
	@echo "✅ precommit: lightweight gate passed. Run \`make verify\` before opening a PR."

clean:
	pnpm clean || true

cli:
	pnpm --filter @officeai/agent build
	@echo "Run via: pnpm --filter @officeai/agent exec office-agent --help"

fixtures:
	pnpm fixtures:docx
	pnpm fixtures:xlsx
	pnpm fixtures:pptx

fixtures-real:
	node scripts/generate-real-fixtures.mjs

fixtures-xlsx:
	pnpm fixtures:xlsx

fixtures-pptx:
	pnpm fixtures:pptx

fixtures-pptx-real:
	pnpm fixtures:pptx-real

# ── Heavy / opt-in checks ────────────────────────────────────────────
# These are intentionally NOT wired into `make verify` because they
# require system-level dependencies (LibreOffice, xmllint, Playwright
# browsers) that not every dev machine has installed. CI runs each one
# in a dedicated job (see .github/workflows/ci.yml). They all skip
# gracefully when their system dependency is absent, so wrappers can
# call them blindly.
#
# Naming convention: <gate>-<format> for per-format runners,
# <gate>-all for the union.
#
# `make heavy` is the catch-all that runs every heavy gate. Useful
# before a release or a major refactor when you want to catch
# everything CI runs in heavy jobs in a single command.
heavy: roundtrip-libre-all schema-validate-all perf-all
	@echo ""
	@echo "✅ heavy: all heavy gates passed (roundtrip + schema + perf for docx/xlsx/pptx)."

# ── LibreOffice roundtrip (per-format) ───────────────────────────────
# Each runner shells out to `soffice --headless --convert-to pdf` on
# every fixture, both for the input and for the agent re-emit, and
# fails on any "repair" / "error" / libreoffice non-zero exit. Inputs
# whose PDF conversion fails before our roundtrip even runs (e.g.
# corrupt embedded PNG in the original fixture) are reported as a
# skip, not a failure — the gate is for *regressions* we introduce.
roundtrip-libre: roundtrip-libre-docx

roundtrip-libre-docx:
	node scripts/run-libreoffice-roundtrip.mjs --format docx

roundtrip-libre-xlsx:
	node scripts/run-libreoffice-roundtrip.mjs --format xlsx

roundtrip-libre-pptx:
	node scripts/run-libreoffice-roundtrip.mjs --format pptx

roundtrip-libre-all: roundtrip-libre-docx roundtrip-libre-xlsx roundtrip-libre-pptx
	@echo ""
	@echo "✅ roundtrip-libre-all: docx + xlsx + pptx fixtures roundtrip clean."

# `e2e-web` builds the workspace first so apps/web's Next.js compile and
# the @officeai/docx dist outputs are ready, then runs Playwright against
# `next start`. Run `pnpm --filter @officeai/web e2e:install` once to
# fetch the browser binaries.
e2e-web:
	pnpm build
	pnpm --filter @officeai/web e2e

# ── Perf budgets (per-format) ────────────────────────────────────────
# Each script builds a representative synthetic doc/sheet/deck, then
# measures: parse + 1k commands + serialize against budgets pinned
# inline in the script. Designed for Apple Silicon dev machines and
# the GitHub-Actions ubuntu-latest runner; budgets are conservative.
# Each script exits non-zero when any phase blows the budget so CI
# fails honestly (no "performance creep").
perf: perf-all

perf-all: perf-docx perf-xlsx perf-pptx
	@echo ""
	@echo "✅ perf-all: all per-format perf budgets met."

perf-docx:
	node scripts/perf-docx.mjs

perf-xlsx:
	node scripts/perf-xlsx.mjs

perf-pptx:
	node scripts/perf-pptx.mjs

# `licenses` walks the resolved pnpm dependency graph and hard-fails on any
# AGPL / GPL-only / SSPL / BUSL entry. Permissive SPDX entries are summarized;
# LGPL / GPL-with-exception is surfaced as a warning. Reads each
# package.json's `license` field — no network calls.
licenses:
	node scripts/license-scan.mjs

# `metrics` prints a holistic snapshot of the repo: source LOC by package,
# fixture counts, test counts, CI job count, etc. Pure reporting — never
# fails the build. Used to keep the README / build log honest about scale.
metrics:
	node scripts/quality-metrics.mjs

# `xsd-fetch` downloads the ECMA-376 5th-edition Transitional OOXML XSDs into
# `vendor/ooxml-xsd/`. The script is idempotent (skips when wml.xsd is already
# present) and SHA-256-pins the source archive. CI caches the directory across
# runs keyed on the script hash. See `scripts/fetch-ooxml-xsd.mjs` for the
# pinned URL + SHA + license note.
xsd-fetch:
	node scripts/fetch-ooxml-xsd.mjs

# ── OOXML XSD validation (per-format) ────────────────────────────────
# Each runner: opens every fixture, enumerates its XML parts, runs the
# same fixture through the format-specific agent (DocxAgent /
# XlsxAgent / PptxAgent), and validates BOTH the input AND the
# re-emit against the ECMA-376 Transitional XSDs via `xmllint`.
# DOCX additionally applies a trivial `insert-text` edit before
# exporting so the diff path is exercised end-to-end (preserves the
# historical CI signal). XLSX/PPTX simply round-trip via
# `fromBuffer → exportFile()`.
#
# All three are advisory in CI today: ECMA-376 Transitional flags
# pre-existing extension attributes (e.g. mc:Ignorable, w15:*,
# DML pattern quirks like buSzPct) that the upstream generators emit
# verbatim. The gate stays advisory until those serializer fixes land;
# the artifact still surfaces every violation for follow-up.
schema-validate: schema-validate-docx

schema-validate-docx:
	node scripts/validate-ooxml-schemas.mjs --format docx

schema-validate-xlsx:
	node scripts/validate-ooxml-schemas.mjs --format xlsx

schema-validate-pptx:
	node scripts/validate-ooxml-schemas.mjs --format pptx

schema-validate-all: schema-validate-docx schema-validate-xlsx schema-validate-pptx
	@echo ""
	@echo "✅ schema-validate-all: docx + xlsx + pptx XSD validation complete."
