"""OpenAI provider implementation."""
from __future__ import annotations

from .base import OpenAICompatibleProvider


class OpenAIProvider(OpenAICompatibleProvider):
    """Translate using OpenAI's chat completion API."""

    api_key_env = "OPENAI_API_KEY"
    default_base_url = None
