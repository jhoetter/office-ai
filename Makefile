# ============================================
# officeAI — convenience targets
# ============================================

.PHONY: help install install-js install-py dev dev-web dev-backend \
        build lint test format clean

help:
	@echo "officeAI — available targets:"
	@echo ""
	@echo "  install       Install all JS + Python dependencies"
	@echo "  dev           Start backend + frontend dev servers"
	@echo "  dev-web       Start only the Next.js frontend"
	@echo "  dev-backend   Start only the FastAPI backend"
	@echo "  build         Production build (all packages)"
	@echo "  lint          Lint TS + Python"
	@echo "  test          Run all tests"
	@echo "  format        Auto-format TS + Python"
	@echo "  clean         Remove build artifacts and dependencies"

install: install-js install-py

install-js:
	pnpm install

install-py:
	cd backend && uv sync --extra dev

dev:
	@echo "Starting backend (:8000) and frontend (:3000)…"
	@trap 'kill 0' INT; \
	  $(MAKE) dev-backend & \
	  $(MAKE) dev-web & \
	  wait

dev-web:
	pnpm --filter @officeai/web dev

dev-backend:
	cd backend && uv run uvicorn app.main:app --reload --host 0.0.0.0 --port 8000

build:
	pnpm build

lint:
	pnpm lint
	cd backend && uv run ruff check .

test:
	pnpm test
	cd backend && uv run pytest

format:
	pnpm format
	cd backend && uv run ruff format .

clean:
	pnpm clean || true
	rm -rf backend/.venv backend/.pytest_cache backend/.ruff_cache backend/.mypy_cache
	find . -type d -name "__pycache__" -prune -exec rm -rf {} +
