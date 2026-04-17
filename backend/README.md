# officeAI backend

FastAPI service powering the officeAI scaffold.

## Run locally

```bash
uv sync --extra dev
uv run uvicorn app.main:app --reload --port 8000
```

- API docs: http://localhost:8000/docs
- Health: http://localhost:8000/health

## Layout

```
backend/
├── app/
│   ├── main.py           # FastAPI app factory + router wiring
│   ├── config.py         # Pydantic settings (env-driven)
│   ├── core/             # cross-cutting: exceptions, logging
│   └── domains/          # domain-driven feature modules
│       └── documents/    # router · service · schemas
└── tests/
```

The single `documents` domain is the hello-world surface — an in-memory
CRUD that backs the editor in `apps/web`. Replace the in-memory store with
SQLAlchemy + a real DB when the product direction is clearer; the interface
is already async so the upgrade is non-breaking.
