import logging
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.config import settings
from app.core.exceptions import NotFoundError, OfficeAIError, ValidationError
from app.core.logging import setup_logging

logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI) -> AsyncIterator[None]:
    level = "DEBUG" if settings.debug else settings.log_level
    setup_logging(level=level)
    logger.info("officeAI backend starting — level=%s debug=%s", level, settings.debug)
    yield
    logger.info("officeAI backend shutting down")


def create_app() -> FastAPI:
    """Application factory."""
    app = FastAPI(
        title=settings.app_name,
        description="officeAI — minimal product scaffold",
        version="0.1.0",
        docs_url="/docs",
        redoc_url="/redoc",
        openapi_url="/openapi.json",
        lifespan=lifespan,
    )

    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.backend_cors_origins,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @app.exception_handler(NotFoundError)
    async def not_found_handler(request: Request, exc: NotFoundError) -> JSONResponse:
        return JSONResponse(status_code=404, content={"detail": exc.detail})

    @app.exception_handler(ValidationError)
    async def validation_handler(request: Request, exc: ValidationError) -> JSONResponse:
        return JSONResponse(status_code=400, content={"detail": exc.detail})

    @app.exception_handler(OfficeAIError)
    async def officeai_error_handler(request: Request, exc: OfficeAIError) -> JSONResponse:
        return JSONResponse(status_code=500, content={"detail": exc.detail})

    @app.get("/health", tags=["system"])
    async def health_check() -> dict[str, str]:
        return {"status": "ok", "service": "officeai-backend"}

    from app.domains.documents.router import router as documents_router

    app.include_router(
        documents_router,
        prefix=f"{settings.api_v1_prefix}/documents",
        tags=["documents"],
    )

    return app


app = create_app()
