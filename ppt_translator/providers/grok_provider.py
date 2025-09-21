"""Grok provider implementation."""
from __future__ import annotations

import os

from .base import OpenAICompatibleProvider


class GrokProvider(OpenAICompatibleProvider):
    """Translate using Grok's OpenAI compatible endpoint."""

    api_key_env = "GROK_API_KEY"

    def __init__(self, model: str, **kwargs) -> None:
        base_url = kwargs.pop("base_url", None) or os.getenv("GROK_API_BASE", "https://api.grok.com/v1")
        super().__init__(model=model, base_url=base_url, **kwargs)
