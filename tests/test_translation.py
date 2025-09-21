from __future__ import annotations

from ppt_translator.translation import TranslationService
from ppt_translator.providers.base import TranslationProvider


class DummyProvider(TranslationProvider):
    def __init__(self) -> None:
        super().__init__(model="dummy")
        self.calls: list[str] = []

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        self.calls.append(text)
        return f"{text}->{target_lang}"


def test_chunk_text_respects_maximum_size():
    provider = DummyProvider()
    service = TranslationService(provider, max_chunk_size=20)
    text = "Sentence one. Sentence two. Sentence three."
    chunks = service.chunk_text(text, max_chunk_size=20)
    assert len(chunks) >= 2
    assert all(len(chunk) <= 20 for chunk in chunks)


def test_translate_uses_cache():
    provider = DummyProvider()
    service = TranslationService(provider, max_chunk_size=100)
    result_a = service.translate("Hello world", "en", "fr")
    result_b = service.translate("Hello world", "en", "fr")
    assert result_a == result_b
    assert provider.calls.count("Hello world") == 1
    assert service.cache_size() == 1


def test_chunk_text_handles_very_long_sentence():
    provider = DummyProvider()
    service = TranslationService(provider, max_chunk_size=50)
    text = "a" * 130
    chunks = service.chunk_text(text, max_chunk_size=50)
    assert all(len(chunk) <= 50 for chunk in chunks)
    assert sum(len(chunk) for chunk in chunks) >= len(text)


def test_clear_cache_resets_entries():
    provider = DummyProvider()
    service = TranslationService(provider, max_chunk_size=100)
    service.translate("Cache me", "en", "de")
    assert service.cache_size() == 1
    service.clear_cache()
    assert service.cache_size() == 0
