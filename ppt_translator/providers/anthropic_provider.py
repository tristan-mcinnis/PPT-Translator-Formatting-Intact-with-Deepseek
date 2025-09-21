"""Anthropic provider implementation."""
from __future__ import annotations

import os

from anthropic import Anthropic

from .base import ProviderConfigurationError, TranslationProvider


class AnthropicProvider(TranslationProvider):
    """Translate content using Anthropic's Messages API."""

    api_key_env = "ANTHROPIC_API_KEY"

    def __init__(self, model: str, *, api_key: str | None = None, max_tokens: int = 4096, temperature: float = 0.3) -> None:
        super().__init__(model, temperature=temperature)
        resolved_key = api_key or os.getenv(self.api_key_env)
        if not resolved_key:
            raise ProviderConfigurationError(
                "Missing API key for provider 'Anthropic'. "
                f"Set the {self.api_key_env} environment variable."
            )
        self.client = Anthropic(api_key=resolved_key)
        self.max_tokens = max_tokens

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        system_prompt = (
            "You are a translation assistant. Translate the user provided text "
            f"from {source_lang} to {target_lang} while preserving tone and formatting."
        )
        response = self.client.messages.create(
            model=self.model,
            max_tokens=self.max_tokens,
            temperature=self.temperature,
            system=system_prompt,
            messages=[{"role": "user", "content": text}],
        )
        return " ".join(part.text.strip() for part in response.content if getattr(part, "text", "")).strip()
