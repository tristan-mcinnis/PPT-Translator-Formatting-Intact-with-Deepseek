"""Utility helpers for CLI and filesystem handling."""
from __future__ import annotations

from pathlib import Path
from typing import Iterator


def clean_path(path: str) -> str:
    """Normalise shell provided paths, removing quotes and escaped spaces."""
    normalised = path.strip("'\"")
    normalised = normalised.replace("\\ ", " ")
    normalised = normalised.replace("\\'", "'")
    return normalised


def iter_presentation_files(target: Path) -> Iterator[Path]:
    """Yield PowerPoint files contained in *target*.

    If *target* is a single file it will be yielded when it has the expected
    suffix. When *target* is a directory the function walks the tree and yields
    any ``.ppt`` or ``.pptx`` files that are discovered.
    """
    suffixes = {".ppt", ".pptx"}
    if target.is_file():
        if target.suffix.lower() in suffixes:
            yield target
        return

    if not target.exists():
        return

    for path in target.rglob("*"):
        if path.is_file() and path.suffix.lower() in suffixes:
            yield path
