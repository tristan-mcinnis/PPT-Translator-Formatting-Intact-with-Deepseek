"""Base classes for translation providers."""
from __future__ import annotations

import os
from abc import ABC, abstractmethod
from typing import Dict, List

from openai import OpenAI


class ProviderConfigurationError(RuntimeError):
    """Raised when a provider cannot be configured properly."""


class TranslationProvider(ABC):
    """Abstract provider responsible for translating text."""

    def __init__(self, model: str, temperature: float = 0.3) -> None:
        self.model = model
        self.temperature = temperature

    @abstractmethod
    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        """Translate ``text`` from ``source_lang`` to ``target_lang``."""


class OpenAICompatibleProvider(TranslationProvider):
    """Provider implementation for OpenAI compatible chat completion APIs."""

    api_key_env: str = "OPENAI_API_KEY"
    default_base_url: str | None = None

    def __init__(
        self,
        model: str,
        *,
        api_key: str | None = None,
        base_url: str | None = None,
        temperature: float = 0.3,
        organization: str | None = None,
    ) -> None:
        super().__init__(model, temperature=temperature)
        resolved_key = api_key or os.getenv(self.api_key_env)
        if not resolved_key:
            raise ProviderConfigurationError(
                f"Missing API key for provider '{self.__class__.__name__}'. "
                f"Set the {self.api_key_env} environment variable."
            )
        self.client = OpenAI(api_key=resolved_key, base_url=base_url or self.default_base_url, organization=organization)

    def build_messages(self, text: str, source_lang: str, target_lang: str) -> List[Dict[str, str]]:
        """Construct chat messages sent to the model."""
        system_prompt = (
            "You are a translation assistant. Translate the user provided text "
            f"from {source_lang} to {target_lang} while preserving tone and formatting."
        )
        return [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text},
        ]

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        response = self.client.chat.completions.create(
            model=self.model,
            messages=self.build_messages(text, source_lang, target_lang),
            temperature=self.temperature,
            stream=False,
        )
        return response.choices[0].message.content.strip()
