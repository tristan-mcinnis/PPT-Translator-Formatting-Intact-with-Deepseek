# PPT Translator (Formatting Intact) with Deepseek 🎯

A powerful PowerPoint translation tool that preserves all formatting while translating content using the Deepseek API. This tool maintains fonts, colors, layouts, and other styling elements while providing accurate translations between languages.

## ✨ Features

- Preserves all PowerPoint formatting during translation
- Supports tables, text boxes, and other PowerPoint elements
- Maintains font styles, sizes, colors, and alignments
- Intelligent text chunking for better translation quality
- Caches translations to avoid duplicate API calls
- Multi-threaded processing for faster execution
- Creates intermediate backups during translation
- Supports custom source and target languages

## 🚀 Installation

1. Clone this repository:
```bash
git clone https://github.com/tristan-mcinnis/PPT-Translator-Formatting-Intact-with-Deepseek.git
cd PPT-Translator-Formatting-Intact-with-Deepseek
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

4. Create a `.env` file in the project root and add your Deepseek API key:
```
DEEPSEEK_API_KEY=your_deepseek_api_key_here
```

## 💻 Usage

1. Run the script:
```bash
python main.py
```

2. Follow the prompts:
   - Enter the path to your PowerPoint file
   - Specify source language code (default: 'zh' for Chinese)
   - Specify target language code (default: 'en' for English)

3. The script will:
   - Generate an XML representation of your PowerPoint
   - Translate the content while preserving formatting
   - Create a new PowerPoint file with the translated content
   - Save the output file as `{original_filename}_translated.pptx`

## 📝 Example

```bash
Please enter the path to your PowerPoint file: /path/to/your/presentation.pptx
Enter source language code (default 'zh' for Chinese): zh
Enter target language code (default 'en' for English): en
```

## ⚙️ Supported Languages

The tool supports all languages available through the Deepseek API. Common language codes include:
- 'zh': Chinese
- 'en': English
- 'es': Spanish
- 'fr': French
- 'de': German
- 'ja': Japanese
- 'ko': Korean

## 🔍 Notes

- The tool automatically adjusts font sizes for translated text to maintain layout integrity
- English translations use Arial font by default for better compatibility
- Font sizes are reduced by 20% for English text to accommodate typically longer translations
- Intermediate XML files are automatically cleaned up after successful processing

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to check [issues page](https://github.com/tristan-mcinnis/PPT-Translator-Formatting-Intact-with-Deepseek/issues).

## ⭐️ Show your support

Give a ⭐️ if this project helped you!

