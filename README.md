[简体中文](README.zh-CN.md) | English

# Engineering Word Toolkit v0.2.0

A user-friendly toolkit for document style management, template governance, and document standardization. Designed for non-technical users to quickly organize and standardize their documents.

Repository: [ysbushe/word-style-manager](https://github.com/ysbushe/word-style-manager)

[![Tests](https://github.com/ysbushe/word-style-manager/actions/workflows/tests.yml/badge.svg)](https://github.com/ysbushe/word-style-manager/actions/workflows/tests.yml)

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
- Clean unused styles, bullets, numbering, and multilevel lists
- Export all or selected styles from Word documents
- Import all or selected styles from the template library or external files
- Auto-include style dependencies and synchronize multilevel numbering
- Property preview and fixed sample-layout preview
- Standalone template editor with version management
- Built-in and custom multilevel numbering presets
- Folder batch processing, task schemes, and repeatable workflows
- Import preflight and HTML reports
- Configurable template library and output directories
- In-app GitHub link, manual update checks, and daily automatic checks
- Portable builds can download a GitHub Release, preserve local settings/templates, replace files, and restart automatically

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

## Automatic Updates

- Open the repository with the **GitHub** button in the application header.
- Use **Check for updates** to query the latest GitHub Release.
- Daily automatic checks are enabled by default and can be disabled in the status bar.
- Portable builds can download and install a release ZIP while preserving `cleaner_config.json` and `templates`.
- Source mode opens the Release page instead of replacing source files.
- Release ZIP assets should include `绿色版` or `portable` in their filename.

## Project Structure
```
├── src/ewt/                # Core code
│   ├── config.py           # Constants
│   ├── core/               # Core engine
│   │   ├── style_engine.py # Style cleaning & analysis
│   │   ├── numbering.py    # Numbering / multi-level lists
│   │   ├── templates.py    # Template export/import/library
│   │   ├── editor.py       # Template style and numbering editor
│   │   └── reports.py      # HTML reports
│   ├── ui/                 # GUI
│   │   ├── app.py          # Main window
│   │   ├── dialogs.py      # Dialogs
│   │   ├── template_editor.py # Template editor
│   │   └── widgets.py      # Preview and reusable components
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
