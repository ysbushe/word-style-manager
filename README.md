# Word 样式管理器 (Engineering Word Toolkit)

面向工程文档的 Word 自动化处理工具。解决施工方案、QC 成果、技术交底、论文等工程文档中的样式混乱、编号不一致、交叉引用断裂等结构性问题。

## 工程文档痛点

在工程文档实践中，以下问题反复出现且高度耗时：

- **样式不统一** — 多作者协作时每人使用不同样式，最终合并后格式混乱
- **编号断裂** — 多级列表（如"第1章 / 1.1 / 1.1.1"）在反复修改后编号层级错乱
- **未使用样式堆积** — 从模板复制粘贴导致文档累积数十甚至上百个无用样式
- **模板难以复用** — 好的样式配置无法快速应用到其他文档

这些问题的共同特征：手工逐一修改效率极低、容易遗漏、出错率高。本工具针对这些场景提供自动化处理能力。

## 功能

| 模块 | 说明 |
|------|------|
| **样式清理** | 扫描并清理文档中未使用的样式、编号/多级列表定义。支持预览后执行 |
| **提取模板** | 从文档提取样式生成 .dotx 模板或 .docx 样式文档。可勾选指定样式，自动补选依赖 |
| **导入样式** | 批量将样式模板导入目标文档。冲突策略：覆盖/跳过/重命名。同步编号/多级列表 |
| **模板库** | 本地模板库管理，模板索引、快速应用 |

## 截图

### 样式清理
![样式清理](docs/screenshots/clean-style.png)

### 提取模板
![提取模板](docs/screenshots/export-template.png)

### 导入样式
![导入样式](docs/screenshots/import-style.png)

### 模板库
![模板库](docs/screenshots/template-library.png)

## 安装

```bash
pip install -r requirements.txt
```

## 使用

```bash
python main.py
```

## 项目结构

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

## 依赖

- Python 3.10+
- python-docx、lxml — OOXML 处理
- ttkbootstrap、tkinterdnd2 — GUI
- pywin32 — .doc → .docx 转换（需 Microsoft Word）

## 许可

MIT License. 详见 [LICENSE](LICENSE).
