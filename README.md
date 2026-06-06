# Word 样式管理器

清理、提取、模板化、批量导入 Word 文档样式与多级编号。

## 功能

- **样式清理** — 自动清理 Word 文档中未使用的样式和编号/多级列表
- **提取模板** — 从文档提取样式为 .dotx / .docx 模板，可勾选指定样式
- **导入样式** — 批量将样式模板导入多个目标文档，支持冲突处理
- **模板库** — 本地模板库管理，模板编辑、编号方案预设

## 运行环境

- Windows 10/11
- 需要安装 [Microsoft Word](https://www.microsoft.com/microsoft-365/word)（用于 .doc 格式转换）

## 使用方式

### 绿色版（打包好的 .exe）

从 [Releases](../../releases) 下载最新版 `Word 样式管理器.zip`，解压后双击 `Word 样式管理器.exe`。

### 从源码运行

```bash
pip install -r requirements.txt
python main.py
```

## 依赖

- Python 3.10+
- python-docx
- lxml
- ttkbootstrap
- tkinterdnd2
- pywin32（.doc 转换需要）

## 打包

```bash
pip install pyinstaller
pyinstaller 极简格式清理引擎.spec
```

## 许可

MIT License
