# ============================================
# officeAI — convenience targets
# ============================================

.PHONY: help install dev dev-web build lint typecheck test format clean cli

help:
	@echo "officeAI — available targets:"
	@echo ""
	@echo "  install      Install all dependencies"
	@echo "  dev          Start the Next.js editor host"
	@echo "  build        Build all packages"
	@echo "  lint         Lint TS"
	@echo "  typecheck    Typecheck all packages"
	@echo "  test         Run all tests"
	@echo "  format       Auto-format"
	@echo "  clean        Remove build artifacts and dependencies"
	@echo "  cli          Build and link the office-agent CLI"

install:
	pnpm install

dev:
	pnpm --filter @officeai/web dev

build:
	pnpm build

lint:
	pnpm lint

typecheck:
	pnpm typecheck

test:
	pnpm test

format:
	pnpm format

clean:
	pnpm clean || true

cli:
	pnpm --filter @officeai/agent build
	@echo "Run via: pnpm --filter @officeai/agent exec office-agent --help"
