<div align='right'>
English | [简体中文](README.zh-CN.md)
</div>

# Engineering Word Toolkit v0.1.0

A user-friendly toolkit for document style management, template governance, and document standardization. Designed for non-technical users to quickly organize and standardize their documents.

## Why this project?

Many people spend excessive time manually fixing document formatting.

Common problems include:
- Inconsistent document styles
- Repeated formatting work
- Difficulty maintaining document templates
- Lack of document standardization

Engineering Word Toolkit helps users standardize, organize, and maintain documents with minimal effort.

## Who is it for?
- Engineers
- Teachers
- Students
- Researchers
- Office workers
- Administrative staff

No programming knowledge required.

## Features
- Export styles from existing Word documents
- Import styles into new documents
- Template management
- Batch style processing
- Visual GUI interface

## Use Cases
- Engineering documentation
- Academic papers
- Reports
- Teaching materials
- Corporate documents
- Administrative documents

## Screenshots

### Style Clean
![Style Clean](docs/screenshots/clean-style.png)

### Export Template
![Export Template](docs/screenshots/export-template.png)

### Import Style
![Import Style](docs/screenshots/import-style.png)

### Template Library
![Template Library](docs/screenshots/template-library.png)

## Installation
```bash
pip install -r requirements.txt
```

## Usage
```bash
python main.py
```

## Project Structure
```
├── src/ewt/                # Core code
│   ├── config.py           # Constants
│   ├── core/               # Core engine
│   │   ├── style_engine.py # Style cleaning & analysis
│   │   ├── numbering.py    # Numbering / multi-level lists
│   │   └── templates.py    # Template export/import/library
│   ├── ui/                 # GUI
│   │   ├── app.py          # Main window
│   │   ├── dialogs.py      # Dialogs
│   │   └── widgets.py      # Reusable components
│   └── utils/              # Utilities
│       ├── helpers.py      # XML helpers
│       └── word_io.py      # Word file I/O
├── tests/                  # Tests
├── main.py                 # Entry point
├── pyproject.toml
└── requirements.txt
```

## Dependencies
- Python 3.10+
- python-docx, lxml — OOXML processing
- ttkbootstrap, tkinterdnd2 — GUI
- pywin32 — .doc to .docx conversion (requires Microsoft Word)

## License
MIT License. See [LICENSE](LICENSE).
