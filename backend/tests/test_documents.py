import pytest
from httpx import ASGITransport, AsyncClient

from app.domains.documents import service
from app.main import app


@pytest.fixture(autouse=True)
async def _reset_store() -> None:
    await service.reset()


async def test_health() -> None:
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        res = await client.get("/health")
    assert res.status_code == 200
    assert res.json() == {"status": "ok", "service": "officeai-backend"}


async def test_create_and_list_document() -> None:
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        created = await client.post(
            "/api/v1/documents",
            json={"title": "Hello", "content": "World"},
        )
        assert created.status_code == 201
        doc = created.json()
        assert doc["title"] == "Hello"
        assert doc["content"] == "World"

        listed = await client.get("/api/v1/documents")
    assert listed.status_code == 200
    assert len(listed.json()) == 1
