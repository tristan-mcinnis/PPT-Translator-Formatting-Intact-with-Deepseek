# PPT Translator – Multi-Provider Edition

Translate PowerPoint presentations while preserving formatting and layout. This CLI-focused tool now supports DeepSeek, OpenAI's latest GPT-5 family (GPT-5, GPT-5 Mini, GPT-5 Nano), Anthropic models, and Grok. It extracts slide content, performs high-quality translations with caching and chunking, and rebuilds decks with styles intact.

## ✨ Key Features

- **Multi-provider support** – switch between DeepSeek, OpenAI, Anthropic, and Grok using a simple CLI flag.
- **Latest model coverage** – ready for GPT-5, GPT-5 Mini, GPT-5 Nano, and the newest Anthropic and Grok offerings.
- **Formatting preserved** – fonts, colours, spacing, tables, and alignment are retained after translation.
- **Translation caching** – avoids duplicate API calls for repeated strings.
- **Chunk-aware processing** – intelligently splits long text to maintain quality and avoid token limits.
- **Threaded slide handling** – speeds up large decks using configurable worker threads.
- **Testable architecture** – modular codebase with Pytest-based unit tests.

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
