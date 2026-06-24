# ============================================
# office-ai — convenience targets
#
# `make verify` is the canonical quality gate: identical to what CI
# runs. Pass it locally before pushing.
# ============================================

.PHONY: help install dev dev-forwarded dev-forwarded-fugu dev-realtime kill-ports build lint lint-root lint-web format format-check architecture actions fixtures-check \
        typecheck test test-docx test-xlsx test-pptx test-core test-web verify ci precommit \
        clean cli fixtures fixtures-real fixtures-xlsx fixtures-pptx fixtures-pptx-real \
        roundtrip-libre roundtrip-libre-docx roundtrip-libre-xlsx roundtrip-libre-pptx \
        roundtrip-libre-all audit-roundtrip audit-roundtrip-pdf fixtures-pdf \
        e2e-web perf perf-docx perf-xlsx perf-pptx perf-all \
        licenses metrics heavy \
        xsd-fetch schema-validate schema-validate-docx schema-validate-xlsx \
        schema-validate-pptx schema-validate-all

# ----------------------------------------------------------------------
# Dev ports (Next.js editor host + realtime websocket)
#
# Default 3100 keeps office-ai out of the common localhost:3000 slot. Override
# with `PORT=3000 make dev` for standalone experiments that want the historical
# Next.js port. The realtime websocket server listens on 1234 by default.
#
# Forwarded ports follow the Sonaloop dev convention:
#   local:     3100 / 1234
#   forwarded: 23003 / 21234
#   fugu:      63003 / 61234
# ----------------------------------------------------------------------
PORT              ?= 3100
FORWARDED_PORT    ?= 23003
FUGU_PORT         ?= 63003
RT_PORT           ?= 1234
RT_HOST           ?= 127.0.0.1
FORWARDED_RT_PORT ?= 21234
FUGU_RT_PORT      ?= 61234
REALTIME_URL      ?= ws://localhost:$(RT_PORT)

# ----------------------------------------------------------------------
# Realtime-server reuse
#
# When another local stack is already using office-ai collaboration it may have
# started a compatible realtime server on :$(RT_PORT). If we then `kill-ports`
# and `pnpm --filter @officeai/realtime-server dev` blindly we:
#
#   1. tear down the healthy server that stack is depending on, and
#   2. race its respawn loop for the bind, surfacing a scary
#      `apps/realtime-server dev: Failed` line for what is actually a
#      benign duplicate-spawn collision.
#
# So `kill-ports` and `dev` both probe `/health` first (IPv4 *and* IPv6
# — the server binds `[::1]` and not every macOS dev box maps
# `localhost` → `::1`). If a healthy server already answers we leave it
# alone and let `next dev` share it. Standalone behaviour is unchanged
# (probe misses → kill + spawn as before).
#
# Override with `OAI_RT_REUSE=0 make dev` to force the historical
# kill+respawn behaviour (useful when iterating on realtime-server itself
# and you want a clean restart on every `make dev`).
# ----------------------------------------------------------------------
OAI_RT_REUSE ?= 1

help:
	@echo "office-ai — available targets (see docs/ci-pipeline.md for the full map)"
	@echo ""
	@echo "  install        Install all dependencies"
	@echo "  dev            Start the Next.js editor host (port \$$PORT, default 3100)"
	@echo "  dev-forwarded  Start web/realtime on tunnel-friendly ports $(FORWARDED_PORT)/$(FORWARDED_RT_PORT)"
	@echo "  dev-forwarded-fugu"
	@echo "                 Start web/realtime on Fugu ports $(FUGU_PORT)/$(FUGU_RT_PORT)"
	@echo "                   Reuses an existing healthy realtime server on \$$RT_PORT;"
	@echo "                   set OAI_RT_REUSE=0 to force a clean restart of the realtime server."
	@echo "  kill-ports     Free \$$PORT (default 3100) and \$$RT_PORT (default 1234, skipped when healthy) — auto-runs as a prereq of \`dev\`"
	@echo "  build          Build all packages"
	@echo ""
	@echo "  Quality gates  (== what CI runs in the 'verify' job):"
	@echo "  format         Auto-format all files (prettier --write)"
	@echo "  format-check   Check formatting (CI-safe, no writes)"
	@echo "  lint           Lint root + apps/web (== lint-root + lint-web)"
	@echo "  architecture   Validate package dep graph (separation of concerns)"
	@echo "  actions        Validate CLI/palette/UI action parity (every bus handler is catalogued)"
	@echo "  fixtures-check Validate fixtures/MATRIX.json and fixture file coverage"
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

# `make dev` must build the workspace packages first, otherwise Next.js
# loads STALE `dist/` artifacts for @officeai/{core,docx,xlsx,pptx} —
# `next.config.ts` only transpiles `@officeai/ui` + `@officeai/design-tokens`,
# the rest are consumed from their compiled `dist/`. A stale dist makes the
# editor look loaded (filename / chrome render from initial state) but the
# agent fails silently and the body is stuck on EmptyState. Turbo skips
# packages that are already up-to-date so the cost on a warm checkout is
# ~1s; on a fresh checkout it's a one-time ~30s build.
#
# Iterating on a package's source after `make dev` is running? Re-run
# `pnpm build` (or `pnpm --filter @officeai/<pkg> build`) in another shell
# — Next.js HMR will pick up the new dist on next request.
#
# PORT is forwarded into Next.js dev (it honours $PORT natively). Run with
# `PORT=3000 make dev` if you want the historical localhost:3000 in standalone
# office-ai workflows.
# Frees the ports we own before re-launching `make dev`. Mirrors
# mail-ai/collaboration-ai's pattern: kill the PIDs holding the port,
# AND `pkill` the supervisor processes in case Next/turbo respawn a
# child between `lsof` and the next bind. Polls until the ports are
# really free (max ~3s) so re-running `make dev` from any state Just
# Works.
#
# A naive `lsof | kill` loses to next dev / turbo because:
#   - `next dev` spawns a Turbopack worker that holds the port; killing
#     the parent leaves the child orphaned and still listening.
#   - Between kill and the next bind there's a tiny window in which the
#     supervisor can respawn.
#
# Realtime coexistence:
#   To avoid two local stacks fighting over the bind, both `kill-ports` and
#   `dev` first probe `http://localhost:$(RT_PORT)/health`. If a healthy server
#   already answers, we leave it alone and `next dev` shares it; if not we kill
#   the port and start our own. Set `OAI_RT_REUSE=0` to force the historical
#   kill+respawn behaviour. See the `OAI_RT_REUSE` block at the top of this file
#   for the full rationale.
kill-ports:
	@WS_TAG="$(CURDIR)"; \
	PORTS="$(PORT)"; \
	if [ "$(OAI_RT_REUSE)" = "1" ] \
	   && ( curl -sf -m 1 http://$(RT_HOST):$(RT_PORT)/health >/dev/null 2>&1 \
	     || curl -sf -m 1 http://localhost:$(RT_PORT)/health  >/dev/null 2>&1 \
	     || curl -sf -m 1 http://[::1]:$(RT_PORT)/health      >/dev/null 2>&1 ); then \
	  echo "kill-ports: realtime server on :$(RT_PORT) is healthy — reusing (set OAI_RT_REUSE=0 to force restart)."; \
	else \
	  PORTS="$$PORTS $(RT_PORT)"; \
	fi; \
	for _ in 1 2 3 4 5 6; do \
	  for p in $$PORTS; do \
	    pids=$$(lsof -ti :$$p 2>/dev/null); \
	    [ -n "$$pids" ] && kill -9 $$pids 2>/dev/null || true; \
	  done; \
	  pkill -9 -f "[n]ext-server.*$$WS_TAG"  2>/dev/null || true; \
	  pkill -9 -f "[n]ext dev.*$$WS_TAG"     2>/dev/null || true; \
	  pkill -9 -f "[t]sx.*$$WS_TAG"          2>/dev/null || true; \
	  pkill -9 -f "[t]urbo run dev"          2>/dev/null || true; \
	  busy=""; \
	  for p in $$PORTS; do \
	    lsof -ti :$$p >/dev/null 2>&1 && busy="$$busy $$p"; \
	  done; \
	  [ -z "$$busy" ] && exit 0; \
	  sleep 0.5; \
	done; \
	echo "kill-ports: still in use after retries:$$busy" >&2; \
	exit 1

dev: kill-ports
	pnpm build
	@if [ "$(OAI_RT_REUSE)" = "1" ] \
	   && ( curl -sf -m 1 http://$(RT_HOST):$(RT_PORT)/health >/dev/null 2>&1 \
	     || curl -sf -m 1 http://localhost:$(RT_PORT)/health  >/dev/null 2>&1 \
	     || curl -sf -m 1 http://[::1]:$(RT_PORT)/health      >/dev/null 2>&1 ); then \
	  echo ""; \
	  echo "→ next dev      http://localhost:$(PORT)"; \
	  echo "→ realtime ws   $(REALTIME_URL)   (reusing existing healthy server)"; \
	  echo ""; \
	  PORT=$(PORT) NEXT_PUBLIC_OAI_REALTIME_URL=$(REALTIME_URL) pnpm --filter @officeai/web dev; \
	else \
	  echo ""; \
	  echo "→ next dev      http://localhost:$(PORT)"; \
	  echo "→ realtime ws   $(REALTIME_URL)   (health: http://$(RT_HOST):$(RT_PORT)/health)"; \
	  echo ""; \
	  PORT=$(PORT) OAI_RT_PORT=$(RT_PORT) OAI_RT_HOST=$(RT_HOST) NEXT_PUBLIC_OAI_REALTIME_URL=$(REALTIME_URL) pnpm --parallel --filter @officeai/web --filter @officeai/realtime-server dev; \
	fi

# Forwarded dev profile for viewing this machine through an SSH tunnel:
#   ssh -L $(FORWARDED_PORT):127.0.0.1:$(FORWARDED_PORT) \
#       -L $(FORWARDED_RT_PORT):127.0.0.1:$(FORWARDED_RT_PORT) <host>
dev-forwarded:
	$(MAKE) dev PORT=$(FORWARDED_PORT) RT_PORT=$(FORWARDED_RT_PORT) REALTIME_URL=ws://localhost:$(FORWARDED_RT_PORT)

# Same, but on the Fugu (non-EU) dev host's port range so it can be tunnelled
# alongside the EU host without local port clashes (FUGU = FORWARDED + 40000).
dev-forwarded-fugu:
	$(MAKE) dev PORT=$(FUGU_PORT) RT_PORT=$(FUGU_RT_PORT) REALTIME_URL=ws://localhost:$(FUGU_RT_PORT)

dev-realtime: kill-ports
	OAI_RT_PORT=$(RT_PORT) OAI_RT_HOST=$(RT_HOST) pnpm --filter @officeai/realtime-server dev

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

# `actions` checks that every command-bus handler has a catalogue entry.
# Cheap (<1s) source-only scan; runs BEFORE typecheck so missing
# catalogue coverage surfaces in seconds, not after a slow build. See
# scripts/check-action-parity.mjs and packages/{format}/src/actions/.
actions:
	pnpm actions

fixtures-check:
	pnpm fixtures:check

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
#   4. actions       — catches CLI/palette parity drift (every bus handler is catalogued)
#   5. fixtures-check — catches missing/unindexed real document fixtures
#   6. architecture   — catches package-level dep-graph violations
#   7. typecheck      — catches type-system violations
#   8. test           — catches behavioural regressions
#   9. build          — catches build-time/integration issues
verify: format-check lint-root lint-web actions fixtures-check architecture typecheck test build
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

# ── Attribute-fidelity round-trip audit ──────────────────────────────
# Parses → exports → re-parses every fixture and counts the curated
# set of formatting attributes (font, size, alignment, list, page,
# charts, …) on both sides. Fails-soft (exit 0) but writes a JSON
# summary to docs/build-log/roundtrip-audit-night.json that ops can
# diff between runs to catch regressions. Cheap (~ 1s for the 30
# bundled fixtures) so it's safe to call from any developer machine
# without LibreOffice installed.
.PHONY: audit-roundtrip audit-roundtrip-pdf fixtures-pdf
audit-roundtrip:
	pnpm --filter @officeai/docx --filter @officeai/xlsx --filter @officeai/pptx --filter @officeai/pdf build
	node scripts/audit-roundtrip.mjs

# PDF-only variant. Useful while iterating on the PDF parser /
# serializer without paying the docx + xlsx + pptx audit cost. The
# script keeps the existing per-format JSON entries intact when run
# with --product, so this can be chained with the others.
audit-roundtrip-pdf:
	pnpm --filter @officeai/pdf build
	node scripts/audit-roundtrip.mjs --product pdf

# Regenerate the synthetic PDF fixture corpus (12 files). Idempotent;
# only rewrites a fixture when its bytes change.
fixtures-pdf:
	node fixtures/pdf/build-fixtures.mjs

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
