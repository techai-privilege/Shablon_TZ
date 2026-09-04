"""Small network-safety helpers used by all remote data sources."""

from __future__ import annotations

from typing import Any, Protocol

MAX_PAGE_BYTES = 8 * 1024 * 1024
MAX_IMAGE_BYTES = 15 * 1024 * 1024
MAX_PDF_BYTES = 30 * 1024 * 1024


class GetSession(Protocol):
    """Small structural interface shared by requests and deterministic test clients."""

    def get(self, url: str, **kwargs: Any) -> Any: ...


def limited_response_content(response, maximum: int, description: str) -> bytes:
    """Return response bytes after rejecting unexpectedly large payloads."""

    headers = getattr(response, "headers", {}) or {}
    declared = headers.get("Content-Length", "")
    try:
        declared_size = int(declared)
    except (TypeError, ValueError):
        declared_size = 0
    if declared_size > maximum:
        raise ValueError(f"{description} слишком большой ({declared_size} байт).")
    content = response.content
    if len(content) > maximum:
        raise ValueError(f"{description} слишком большой ({len(content)} байт).")
    return content


def checked_image_content(response, description: str) -> tuple[bytes, str]:
    """Validate basic image response metadata and size."""

    headers = getattr(response, "headers", {}) or {}
    content_type = headers.get("Content-Type", "").split(";", 1)[0].strip()
    if not content_type.casefold().startswith("image/"):
        raise ValueError(f"{description} вернул вместо изображения неподдерживаемый файл.")
    content = limited_response_content(response, MAX_IMAGE_BYTES, f"Изображение {description}")
    if not content:
        raise ValueError(f"{description} вернул пустое изображение.")
    return content, content_type
