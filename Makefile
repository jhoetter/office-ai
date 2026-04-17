# ============================================
# officeAI — convenience targets
#
# `make verify` is the canonical quality gate: identical to what CI
# runs. Pass it locally before pushing.
# ============================================

.PHONY: help install dev build lint lint-root format format-check architecture \
        typecheck test verify ci clean cli fixtures fixtures-real \
        roundtrip-libre e2e-web perf-docx licenses xsd-fetch schema-validate

help:
	@echo "officeAI — available targets:"
	@echo ""
	@echo "  install        Install all dependencies"
	@echo "  dev            Start the Next.js editor host"
	@echo "  build          Build all packages"
	@echo ""
	@echo "  Quality gates:"
	@echo "  format         Auto-format all files (prettier --write)"
	@echo "  format-check   Check formatting (CI-safe, no writes)"
	@echo "  lint           Lint all packages (ESLint + per-package rules)"
	@echo "  architecture   Validate package dep graph (separation of concerns)"
	@echo "  typecheck      Typecheck all packages"
	@echo "  test           Run all tests"
	@echo "  verify         Run the full quality gate (== what CI runs)"
	@echo "  ci             Alias for verify"
	@echo ""
	@echo "  Misc:"
	@echo "  clean          Remove build artifacts and dependencies"
	@echo "  cli            Build and link the office-agent CLI"
	@echo "  fixtures       Regenerate synthetic DOCX fixtures"
	@echo "  fixtures-real  Regenerate real-shape DOCX fixtures (Word-grade)"
	@echo ""
	@echo "  Heavy / opt-in checks (NOT part of make verify):"
	@echo "  roundtrip-libre  Headless LibreOffice roundtrip on real-world fixtures"
	@echo "  e2e-web          Playwright smoke tests against the web app"
	@echo "  perf-docx        DOCX perf budgets (parse / 1k commands / serialize)"
	@echo "  licenses         SPDX license-graph scan against the resolved deps"
	@echo "  xsd-fetch        Download the ECMA-376 OOXML XSDs into vendor/ooxml-xsd/"
	@echo "  schema-validate  Validate every fixture (input + agent re-emit) against the OOXML XSDs"

install:
	pnpm install

dev:
	pnpm --filter @officeai/web dev

build:
	pnpm build

lint:
	pnpm lint:root
	pnpm --filter @officeai/web lint

lint-root:
	pnpm lint:root

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

# ── The quality gate ──────────────────────────────────────────────────
# Order is intentional and fail-fast:
#   1. format-check  — cheapest, catches accidental noise
#   2. lint          — catches correctness + import-boundary violations
#   3. architecture  — catches package-level dep-graph violations
#   4. typecheck     — catches type-system violations
#   5. test          — catches behavioural regressions
#   6. build         — catches build-time/integration issues
verify: format-check lint architecture typecheck test build
	@echo ""
	@echo "✅ verify: all quality gates passed."

ci: verify

clean:
	pnpm clean || true

cli:
	pnpm --filter @officeai/agent build
	@echo "Run via: pnpm --filter @officeai/agent exec office-agent --help"

fixtures:
	pnpm fixtures:docx

fixtures-real:
	node scripts/generate-real-fixtures.mjs

# ── Heavy / opt-in checks ────────────────────────────────────────────
# These are intentionally NOT wired into `make verify` because they
# require system-level dependencies (LibreOffice, Playwright browsers)
# that not every dev machine has installed. CI runs them in dedicated
# jobs (.github/workflows/ci.yml: docx-libreoffice-roundtrip, web-e2e).
#
# `roundtrip-libre` skips gracefully if `soffice` is not on PATH, so it
# is safe to invoke from any wrapper script without first probing for
# the binary.
roundtrip-libre:
	node scripts/run-libreoffice-roundtrip.mjs

# `e2e-web` builds the workspace first so apps/web's Next.js compile and
# the @officeai/docx dist outputs are ready, then runs Playwright against
# `next start`. Run `pnpm --filter @officeai/web e2e:install` once to
# fetch the browser binaries.
e2e-web:
	pnpm build
	pnpm --filter @officeai/web e2e

# `perf-docx` builds and serializes a synthetic 100-page DOCX, asserting
# the parse / dispatch / serialize budgets documented inline in the script.
# Skipped from `make verify` because perf is best run on the actual dev
# machine (Apple Silicon target); CI runs it in a dedicated job.
perf-docx:
	node scripts/perf-docx.mjs

# `licenses` walks the resolved pnpm dependency graph and hard-fails on any
# AGPL / GPL-only / SSPL / BUSL entry. Permissive SPDX entries are summarized;
# LGPL / GPL-with-exception is surfaced as a warning. Reads each
# package.json's `license` field — no network calls.
licenses:
	node scripts/license-scan.mjs

# `xsd-fetch` downloads the ECMA-376 5th-edition Transitional OOXML XSDs into
# `vendor/ooxml-xsd/`. The script is idempotent (skips when wml.xsd is already
# present) and SHA-256-pins the source archive. CI caches the directory across
# runs keyed on the script hash. See `scripts/fetch-ooxml-xsd.mjs` for the
# pinned URL + SHA + license note.
xsd-fetch:
	node scripts/fetch-ooxml-xsd.mjs

# `schema-validate` opens every fixture in `fixtures/docx/real-world/`,
# enumerates its XML parts, runs the same fixture through the agent
# (`DocxAgent.fromBuffer → trivial edit → exportFile()`), and validates BOTH
# sides against the ECMA-376 Transitional XSDs via `xmllint --schema`.
# Skipped from `make verify` (heavy / system-dep), same rationale as
# `roundtrip-libre`. Skips gracefully if `xmllint` or the XSDs are missing.
schema-validate:
	node scripts/validate-ooxml-schemas.mjs
