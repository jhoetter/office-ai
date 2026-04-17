"""In-memory document store.

Intentionally trivial — swap for SQLAlchemy + a real DB once the product
direction is clearer. The interface is async so the upgrade is non-breaking.
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import uuid4

from app.core.exceptions import NotFoundError
from app.domains.documents.schemas import Document, DocumentCreate, DocumentUpdate

_store: dict[str, Document] = {}


async def list_documents() -> list[Document]:
    return sorted(_store.values(), key=lambda d: d.updated_at, reverse=True)


async def get_document(doc_id: str) -> Document:
    doc = _store.get(doc_id)
    if doc is None:
        raise NotFoundError(f"Document {doc_id} not found")
    return doc


async def create_document(payload: DocumentCreate) -> Document:
    doc = Document(
        id=str(uuid4()),
        title=payload.title,
        content=payload.content,
        updated_at=datetime.now(UTC),
    )
    _store[doc.id] = doc
    return doc


async def update_document(doc_id: str, payload: DocumentUpdate) -> Document:
    if doc_id not in _store:
        raise NotFoundError(f"Document {doc_id} not found")
    doc = Document(
        id=doc_id,
        title=payload.title,
        content=payload.content,
        updated_at=datetime.now(UTC),
    )
    _store[doc_id] = doc
    return doc


async def reset() -> None:
    """Test helper — wipe the in-memory store."""
    _store.clear()
