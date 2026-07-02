[简体中文](README.zh-CN.md) | [English](README.md)

# Word Style Manager v0.4.0

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
- People who are not confident with Word/WPS formatting and want to reduce manual layout work
- Engineers
- Teachers
- Students
- Researchers
- Office workers
- Administrative staff

No programming knowledge required.

## What's New in v0.4.0

- New **Format Conversion** workspace for batch `.doc` to `.docx` and `.xls` to `.xlsx` conversion
- Choose Microsoft Office or WPS as the conversion engine, with Office used by default
- Detect unavailable engines before conversion and guide users to an available option
- Add legacy files by file picker, folder scan, or drag and drop
- Automatically rename outputs on name conflicts and keep source files unchanged
- Continue after per-file failures and show a clear success/failure summary

## Features
- Clean unused styles, bullets, numbering, and multilevel lists
- Safe and deep cleaning modes, with deeper built-in style removal protected by output validation
- Source documents are never overwritten; cleaned copies are checked for package, XML, style, and numbering integrity
- Hide unused built-in styles while allowing Word/WPS to reveal them when used
- Extract all or selected styles into one compact `.dotx` template
- Auto-include dependencies required by selected exported styles
- Property and sample-layout previews on the cleaning and template-library pages
- Chinese-localized template editor with Word/WPS-style units, live preview, undo, and direct style saving
- Batch-convert legacy `.doc` and `.xls` files to `.docx` and `.xlsx`
- Choose Microsoft Office or WPS for conversion, with smart availability checks
- Drag files or folders into the conversion page, with recursive folder scanning
- Folder batch processing and repeatable workflows
- Configurable template library and output directories
- In-app GitHub link, user guide, and manual update checks
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
Inspect style usage, choose safe or deep cleaning, and create a validated copy.

![Style Clean](docs/screenshots/clean-style.png)

### Export Template
Select styles from the current document and save a compact `.dotx` template.

![Export Template](docs/screenshots/export-template.png)

### Template Editor
Edit typography, spacing, indentation, and color with a live layout preview.

![Template Editor](docs/screenshots/template-editor.png)

### Template Library
Review, edit, and delete saved templates with complete names and timestamps.

![Template Library](docs/screenshots/template-library.png)

### Format Conversion
Convert legacy `.doc` and `.xls` files to modern `.docx` and `.xlsx` files.

Older binary formats such as `.doc` and `.xls` are less consistent for automation tools,
scripts, and AI agents to read directly. They often require extra conversion steps, and
content or formatting extraction can be incomplete or unstable. The newer `.docx` and
`.xlsx` formats are based on the open OOXML standard and are structured package files,
which makes them easier for related tools to read more reliably and completely. This
feature helps users convert existing legacy files in batches before using them in agent
or automation workflows.

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
- The application does not check automatically; GitHub is contacted only after
  the user clicks **Check for updates**.
- Portable builds can download and install a release ZIP while preserving local settings and user templates.
- Existing application files are backed up before replacement. Incomplete downloads,
  damaged ZIP files, unsafe paths, and replacement failures stop the update or trigger rollback.
- Settings remain in the application directory. Templates default to
  `Documents\模板库`, so application replacement does not affect them.
- If an older version used the application-side `templates` folder, its `.dotx`
  files are copied to the new default library without overwriting same-named files.
- Source mode opens the Release page instead of replacing source files.
- Release ZIP assets should include `绿色版` or `portable` in their filename.

## Project Structure
```
├── src/ewt/                # Core code
│   ├── config.py           # Constants
│   ├── core/               # Core engine
│   │   ├── style_engine.py # Style cleaning & analysis
│   │   ├── numbering.py    # Numbering / multi-level lists
│   │   ├── templates.py    # Template export and library management
│   │   ├── editor.py       # Template style editing
│   │   ├── converter.py    # Legacy Office format conversion
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
- pywin32 — .doc/.xls conversion through Microsoft Office or WPS

## User Guide

The packaged Chinese user guide is available at [使用说明.md](使用说明.md).

## License
This project is licensed under the MIT License. See [LICENSE](LICENSE).
Third-party dependency notices are listed in
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
