# ============================================
# officeAI — convenience targets
#
# `make verify` is the canonical quality gate: identical to what CI
# runs. Pass it locally before pushing.
# ============================================

.PHONY: help install dev build lint lint-root format format-check architecture \
        typecheck test verify ci clean cli fixtures

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
	@echo "  fixtures       Regenerate DOCX fixtures"

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
