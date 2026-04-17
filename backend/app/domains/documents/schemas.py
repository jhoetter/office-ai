from datetime import datetime

from pydantic import BaseModel, Field


class DocumentCreate(BaseModel):
    title: str = Field(min_length=1, max_length=200)
    content: str = ""


class DocumentUpdate(BaseModel):
    title: str = Field(min_length=1, max_length=200)
    content: str = ""


class Document(BaseModel):
    id: str
    title: str
    content: str
    updated_at: datetime
