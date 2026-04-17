from fastapi import APIRouter

from app.domains.documents import service
from app.domains.documents.schemas import Document, DocumentCreate, DocumentUpdate

router = APIRouter()


@router.get("", response_model=list[Document])
async def list_documents() -> list[Document]:
    return await service.list_documents()


@router.post("", response_model=Document, status_code=201)
async def create_document(payload: DocumentCreate) -> Document:
    return await service.create_document(payload)


@router.get("/{doc_id}", response_model=Document)
async def get_document(doc_id: str) -> Document:
    return await service.get_document(doc_id)


@router.put("/{doc_id}", response_model=Document)
async def update_document(doc_id: str, payload: DocumentUpdate) -> Document:
    return await service.update_document(doc_id, payload)
