# officeAI

A minimal product scaffold sharing the same design language and tech stack
patterns as `hof-os`. Built to be a clean starting point for a new product
in our Python-leaning stack.

## Stack

- **Frontend**: Next.js 15, React 19, Tailwind CSS 4, TypeScript
- **Backend**: FastAPI, Pydantic v2, async Python 3.12+
- **Design system**: Shared `@officeai/design-tokens` + `@officeai/ui` packages
  (Notion-like aesthetic, light/dark mode via `next-themes`)
- **Monorepo**: Turborepo + pnpm workspaces
- **Python tooling**: uv (package manager), Ruff (linter/formatter)

## Prerequisites

| Tool    | Version | Install                                            |
| ------- | ------- | -------------------------------------------------- |
| Node.js | 20+     | [nodejs.org](https://nodejs.org)                   |
| pnpm    | 9+      | `npm install -g pnpm`                              |
| Python  | 3.12+   | [python.org](https://python.org)                   |
| uv      | latest  | `curl -LsSf https://astral.sh/uv/install.sh \| sh` |

## Quick start

```bash
make install       # install JS + Python deps
cp .env.example .env
make dev           # backend on :8000, frontend on :3000
```

- Frontend: http://localhost:3000
- Backend API: http://localhost:8000
- API Docs: http://localhost:8000/docs

The "hello world" is a minimal in-memory text editor at `/editor`. It
demonstrates the shared design tokens, UI primitives, theme toggle, and the
full frontend↔backend round-trip.

## Layout

```
office-ai/
├── apps/web/                    # Next.js 15 App Router (UI)
├── packages/ui/                 # Shared React component library
├── packages/design-tokens/      # Brand colors, typography, spacing as code
├── backend/                     # FastAPI backend (domain-driven, async)
│   └── app/domains/             # documents, ...
└── Makefile                     # convenience commands
```

## Commands

| Command            | Description                                |
| ------------------ | ------------------------------------------ |
| `make install`     | Install all JS + Python dependencies       |
| `make dev`         | Start backend + frontend together          |
| `make dev-web`     | Start only the Next.js frontend            |
| `make dev-backend` | Start only the FastAPI backend             |
| `make build`       | Production build (all packages)            |
| `make lint`        | Lint everything (TS + Python)              |
| `make test`        | Run backend + frontend tests               |
| `make format`      | Auto-format all code                       |
| `make clean`       | Remove build artifacts and dependencies    |
