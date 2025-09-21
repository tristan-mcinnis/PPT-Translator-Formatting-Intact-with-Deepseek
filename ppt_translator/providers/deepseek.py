"""DeepSeek provider implementation."""
from __future__ import annotations

import os

from .base import OpenAICompatibleProvider


class DeepSeekProvider(OpenAICompatibleProvider):
    """Translate through DeepSeek's OpenAI compatible API."""

    api_key_env = "DEEPSEEK_API_KEY"

    def __init__(self, model: str, **kwargs) -> None:
        base_url = kwargs.pop("base_url", None) or os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com")
        super().__init__(model=model, base_url=base_url, **kwargs)
