"""Domain exceptions used across the API."""


class OfficeAIError(Exception):
    """Base class for application-level errors."""

    detail: str = "Internal server error"

    def __init__(self, detail: str | None = None) -> None:
        if detail:
            self.detail = detail
        super().__init__(self.detail)


class NotFoundError(OfficeAIError):
    detail = "Resource not found"


class ValidationError(OfficeAIError):
    detail = "Invalid request"
