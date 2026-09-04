"""Safe filesystem helpers shared by all generated documents."""

from __future__ import annotations

import os
from contextlib import contextmanager
from pathlib import Path
from tempfile import mkstemp
from typing import Iterator


@contextmanager
def atomic_output_path(destination: str | Path) -> Iterator[Path]:
    """Yield a sibling temporary path and atomically replace the destination.

    Keeping the temporary file in the destination directory avoids cross-device
    rename failures (for example, when the user saves to an external drive).
    The old document remains intact if generation fails.
    """

    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = mkstemp(
        dir=destination.parent,
        prefix=f".{destination.name}.",
        suffix=".tmp",
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        yield temporary
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def atomic_write_bytes(destination: str | Path, content: bytes) -> Path:
    """Write bytes without exposing a partial document at the final path."""

    destination = Path(destination)
    with atomic_output_path(destination) as temporary:
        with temporary.open("wb") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
    return destination
