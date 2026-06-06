<div align='right'>
[English](README.md) | 简体中文
</div>

# 工程文档样式管理工具 / Engineering Word Toolkit v0.1.0

面向非专业用户的文档样式管理工具，降低学习门槛，无需编程经验或 Word 高级排版知识。

## 为什么开发这个项目？

很多人在日常工作、学习和科研中，经常遇到：
- 文档格式不统一
- 样式混乱
- 模板难以维护
- 重复调整格式耗费大量时间

本工具旨在帮助用户轻松标准化、整理和维护文档。

## 适用人群
- 工程人员
- 教师
- 学生
- 科研人员
- 行政办公人员
- 企业文档维护人员

无需编程经验。

## 核心功能
- 导出现有 Word 文档样式
- 导入样式到新文档
- 模板管理
- 批量样式处理
- 可视化 GUI 界面

## 适用场景
- 工程文档
- 论文
- 报告
- 教学材料
- 企业文档
- 行政公文

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
