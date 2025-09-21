"""Provider factory for translation services."""
from __future__ import annotations

from typing import Dict, Type

from .anthropic_provider import AnthropicProvider
from .base import ProviderConfigurationError, TranslationProvider
from .deepseek import DeepSeekProvider
from .grok_provider import GrokProvider
from .openai_provider import OpenAIProvider

PROVIDER_REGISTRY: Dict[str, Type[TranslationProvider]] = {
    "deepseek": DeepSeekProvider,
    "openai": OpenAIProvider,
    "anthropic": AnthropicProvider,
    "grok": GrokProvider,
}

PROVIDER_DEFAULTS: Dict[str, Dict[str, str]] = {
    "deepseek": {"model": "deepseek-chat"},
    "openai": {"model": "gpt-5"},
    "anthropic": {"model": "claude-3.7-sonnet"},
    "grok": {"model": "grok-beta"},
}


def list_providers() -> list[str]:
    """Return the available provider identifiers."""
    return sorted(PROVIDER_REGISTRY.keys())


def create_provider(provider_name: str, *, model: str | None = None, **kwargs) -> TranslationProvider:
    """Instantiate a provider by name."""
    name = provider_name.lower()
    if name not in PROVIDER_REGISTRY:
        raise ValueError(f"Unsupported provider '{provider_name}'. Available: {', '.join(list_providers())}")
    provider_cls = PROVIDER_REGISTRY[name]
    default_options = PROVIDER_DEFAULTS.get(name, {}).copy()
    if model:
        default_options["model"] = model
    if "model" not in default_options:
        raise ValueError(f"No model specified for provider '{provider_name}'.")
    default_options.update(kwargs)
    return provider_cls(**default_options)


__all__ = [
    "AnthropicProvider",
    "DeepSeekProvider",
    "GrokProvider",
    "OpenAIProvider",
    "ProviderConfigurationError",
    "TranslationProvider",
    "create_provider",
    "list_providers",
]
