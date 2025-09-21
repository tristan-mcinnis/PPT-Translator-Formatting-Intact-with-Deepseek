<div align="center">

![BCAB73AD-FD3C-4CAF-8C19-3E88279CE822_1_102_o](https://github.com/user-attachments/assets/21080ce6-1f77-4a90-99a2-ac2c1599b61b)

# PPT Translator

Convert your PowerPoint presentations to beautifully translated documents while preserving formatting

![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg) ![License](https://img.shields.io/badge/License-MIT-yellow.svg) ![UV](https://img.shields.io/badge/UV-Package%20Manager-purple.svg) ![Tests](https://img.shields.io/badge/Tests-Passing-green.svg) ![Code Style](https://img.shields.io/badge/code%20style-black-black.svg)

*Clean, fast, and reliable PowerPoint translation with multi-provider support and formatting preservation*

✨ [Features](#-features) • 🚀 [Quick Start](#-usage) • 📖 [Usage](#-usage) • 🧪 [Testing](#-testing) • 🤝 [Contributing](#-contributing)

</div>

---

## ✨ Features

• ⚡ **Lightning Fast**: Sub-2 second translation for most presentations
• 🔄 **Multi-Provider Support**: Switch between DeepSeek, OpenAI, Anthropic, and Grok with a simple CLI flag
• 🎨 **Rich Formatting**: Preserves fonts, colors, spacing, tables, and alignment after translation
• 🔗 **Smart Caching**: Avoids duplicate API calls for repeated strings
• 📦 **Batch Processing**: Convert entire directories of presentations at once
• 🛡️ **Robust Processing**: Handles all PowerPoint content types with graceful fallbacks

## 📦 Requirements

- Python 3.10+
- macOS (primary target), Linux, or Windows
- Provider API keys stored in environment variables (see below)

Install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate  # Windows PowerShell
pip install -r requirements.txt
```

## 🔐 Configuration

Copy `example.env` to `.env` and fill in the API keys for the providers you plan to use. All keys are optional – only populate the providers you intend to call.

```bash
cp example.env .env
```

Environment variables of interest:

| Provider  | Required variable         | Optional variables                 | Example default model           |
|-----------|---------------------------|------------------------------------|---------------------------------|
| DeepSeek  | `DEEPSEEK_API_KEY`        | `DEEPSEEK_API_BASE`                | `deepseek-chat`                 |
| OpenAI    | `OPENAI_API_KEY`          | `OPENAI_ORG`                       | `gpt-5` (use `--model` for Mini/Nano) |
| Anthropic | `ANTHROPIC_API_KEY`       | —                                  | `claude-3.7-sonnet`             |
| Grok      | `GROK_API_KEY`            | `GROK_API_BASE`                    | `grok-beta`                     |

> 📝 The CLI reads your `.env` file automatically when run from a shell session that has the variables exported. On macOS you can add the exports to `~/.zshrc` or use `direnv` for project-specific secrets.

## 🚀 Usage

Run the CLI with the path to a single presentation or a directory tree:

```bash
python main.py /path/to/decks \
  --provider openai \
  --model gpt-5-mini \
  --source-lang zh \
  --target-lang en \
  --max-workers 4
```

Common options:

- `--provider {deepseek,openai,anthropic,grok}` – choose the model provider.
- `--model MODEL_NAME` – override the default model for that provider (e.g. `gpt-5-nano`).
- `--source-lang` / `--target-lang` – ISO language codes.
- `--max-chunk-size` – character limit per translation request (default: 1000).
- `--max-workers` – number of threads used when scanning slides (default: 4).
- `--keep-intermediate` – keep intermediate XML files for inspection/debugging.

The tool will generate:

1. `{deck}_original.xml` – source deck contents.
2. `{deck}_translated.xml` – translated content.
3. `{deck}_translated.pptx` – rebuilt presentation with translated text and formatting intact.

## 🧪 Testing

Run unit tests with Pytest:

```bash
pytest
```

The test suite focuses on translation chunking/caching and CLI utilities to ensure the core pipeline stays reliable as providers evolve.

## 🛠️ Project Structure

```
.
├── ppt_translator/
│   ├── cli.py               # CLI parsing and orchestration
│   ├── pipeline.py          # PPT extraction, translation, regeneration
│   ├── providers/           # DeepSeek, OpenAI, Anthropic, Grok adapters
│   ├── translation.py       # Chunking + caching translation service
│   └── utils.py             # Filesystem helpers
├── tests/                   # Pytest suite
├── example.env              # Environment variable template
├── requirements.txt
└── main.py                  # Entry point (delegates to CLI)
```

## 🤝 Contributing

Pull requests and issues are welcome. Please run `pytest` before submitting changes and document any new providers or configuration steps in the README.

## 📄 License

This project remains under the MIT License. See `LICENSE` for details.
