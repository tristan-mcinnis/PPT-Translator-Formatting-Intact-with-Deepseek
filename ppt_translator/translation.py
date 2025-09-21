"""Translation service orchestrating providers, caching and chunking."""
from __future__ import annotations

import re
import threading
from typing import Dict, List

from .providers.base import TranslationProvider

_SENTENCE_SPLIT_PATTERN = re.compile(r"(?<=[.!?。！？])\s+")


class TranslationService:
    """Translate text using a configured provider with caching support."""

    def __init__(self, provider: TranslationProvider, *, max_chunk_size: int = 1000) -> None:
        self.provider = provider
        self.max_chunk_size = max_chunk_size
        self._cache: Dict[str, str] = {}
        self._lock = threading.Lock()

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        """Translate ``text`` and cache repeated requests."""
        if not text or text.isspace():
            return text

        with self._lock:
            if text in self._cache:
                return self._cache[text]

        chunks = self.chunk_text(text, self.max_chunk_size)
        translated_chunks: List[str] = []
        for chunk in chunks:
            stripped = chunk.strip()
            if not stripped:
                translated_chunks.append(chunk)
                continue
            translated = self.provider.translate(chunk, source_lang, target_lang)
            translated_chunks.append(translated.strip())

        combined = " ".join(part for part in translated_chunks if part)
        if not combined:
            combined = text

        with self._lock:
            self._cache[text] = combined
        return combined

    @staticmethod
    def chunk_text(text: str, max_chunk_size: int = 1000) -> List[str]:
        """Split long text into smaller chunks preserving sentence boundaries."""
        if len(text) <= max_chunk_size:
            return [text]

        sentences = [segment.strip() for segment in _SENTENCE_SPLIT_PATTERN.split(text) if segment.strip()]
        if not sentences:
            sentences = [text]

        chunks: List[str] = []
        current: List[str] = []
        current_len = 0

        for sentence in sentences:
            sentence_len = len(sentence)
            if current and current_len + sentence_len + 1 > max_chunk_size:
                chunks.append(" ".join(current))
                current = []
                current_len = 0
            if sentence_len > max_chunk_size:
                if current:
                    chunks.append(" ".join(current))
                    current = []
                    current_len = 0
                chunks.extend(
                    [sentence[i : i + max_chunk_size] for i in range(0, sentence_len, max_chunk_size)]
                )
                continue
            current.append(sentence)
            current_len += sentence_len + 1

        if current:
            chunks.append(" ".join(current))

        if not chunks:
            return [text]
        return chunks

    def clear_cache(self) -> None:
        """Drop cached translations."""
        with self._lock:
            self._cache.clear()

    def cache_size(self) -> int:
        """Return the number of cached entries."""
        with self._lock:
            return len(self._cache)
