# Word 样式管理器 / Word Style Manager

## 简介 / Overview

这是一个面向非专业用户的工具，降低学习门槛。无须 Word 排版经验、无需编程知识，只需几步操作即可统一文档样式、整理文档。适用于工程文档、论文、报告、教学材料等多种场景。

This is a tool designed for non-technical users, lowering the learning barrier. No Word formatting experience or programming knowledge required. With just a few clicks, you can standardize document styles and organize your files. Applicable to engineering documents, academic papers, reports, teaching materials, and more.

## 核心功能 / Key Features

| 功能 Feature | 说明 Description |
|------|------|
| **一键导入/导出样式模板** | 从文档提取样式为 .dotx/.docx 模板，批量导入到目标文档 |
| **批量处理文档** | 支持多文件拖拽、文件夹批量清理 |
| **可视化反馈** | 处理结果弹窗 + HTML 报告，成功/失败一目了然 |
| **跨领域适用** | 施工方案、QC 成果、技术交底、论文、报告均可 |

## 快速开始 / Quick Start

1. 选中文档 / Select your document(s)
2. 选择模板 / Choose a template
3. 点击"执行清理"或"开始导入" / Click 'Execute' or 'Start Import'
4. 完成 / Done

## 截图 / Screenshots

### 样式清理 / Style Clean
![样式清理](docs/screenshots/clean-style.png)

### 提取模板 / Export Template
![提取模板](docs/screenshots/export-template.png)

### 导入样式 / Import Style
![导入样式](docs/screenshots/import-style.png)

### 模板库 / Template Library
![模板库](docs/screenshots/template-library.png)

## 安装 / Installation

```bash
pip install -r requirements.txt
```

## 使用 / Usage

```bash
python main.py
```

## 项目目标 / Project Goals

- 降低文档样式管理的学习门槛 / Lower learning barrier for document style management
- 提高文档整理和标准化效率 / Improve document organization and standardization efficiency
- 支持多类型文档和多领域应用 / Support multiple document types and cross-domain usage

## 项目结构 / Project Structure

```
├── src/ewt/                # 核心代码
│   ├── config.py           # 配置常量
│   ├── core/               # 核心引擎
│   │   ├── style_engine.py # 样式清理与分析
│   │   ├── numbering.py    # 编号/多级列表
│   │   └── templates.py    # 模板导出/导入/库管理
│   ├── ui/                 # GUI
│   │   ├── app.py          # 主窗口
│   │   ├── dialogs.py      # 弹窗
│   │   └── widgets.py      # 复用组件
│   └── utils/              # 工具
│       ├── helpers.py      # XML 辅助
│       └── word_io.py      # Word 读写
├── tests/                  # 测试
├── main.py                 # 入口
├── pyproject.toml
└── requirements.txt
```

## 依赖 / Dependencies

- Python 3.10+
- python-docx、lxml — OOXML 处理
- ttkbootstrap、tkinterdnd2 — GUI
- pywin32 — .doc → .docx 转换（需 Microsoft Word）

## 许可 / License

MIT License. 详见 [LICENSE](LICENSE).
