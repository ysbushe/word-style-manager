"""Chinese display labels for OOXML style and unit values."""

from __future__ import annotations

import re


BUILTIN_STYLE_NAMES = {
    "Normal": "正文",
    "Default Paragraph Font": "默认段落字体",
    "DefaultParagraphFont": "默认段落字体",
    "Body Text": "正文文本",
    "BodyText": "正文文本",
    "Title": "标题",
    "Subtitle": "副标题",
    "Caption": "题注",
    "Quote": "引用",
    "Intense Quote": "强调引用",
    "IntenseQuote": "强调引用",
    "List Paragraph": "列表段落",
    "ListParagraph": "列表段落",
    "No Spacing": "无间隔",
    "NoSpacing": "无间隔",
    "Header": "页眉",
    "Footer": "页脚",
    "Footnote Text": "脚注文本",
    "Endnote Text": "尾注文本",
    "Hyperlink": "超链接",
    "Emphasis": "强调",
    "Strong": "强烈强调",
    "Intense Emphasis": "明显强调",
    "IntenseEmphasis": "明显强调",
    "Book Title": "书名",
    "BookTitle": "书名",
    "Macro Text": "宏文本",
    "MacroText": "宏文本",
    "Plain Text": "纯文本",
    "PlainText": "纯文本",
    "HTML Preformatted": "HTML 预设格式",
    "HTMLPreformatted": "HTML 预设格式",
    "Bibliography": "书目",
    "Table Normal": "普通表格",
    "TableNormal": "普通表格",
    "No List": "无列表",
    "NoList": "无列表",
}
CHINESE_LEVELS = "一二三四五六七八九"


def chinese_style_name(name="", style_id=""):
    """Return a Chinese UI label without changing the underlying style ID."""
    raw_name = str(name or "").strip()
    raw_id = str(style_id or "").strip()
    for value in (raw_name, raw_id):
        heading = re.fullmatch(r"Heading\s*([1-9])", value, re.I)
        if heading:
            return f"{CHINESE_LEVELS[int(heading.group(1)) - 1]}级标题"
        toc = re.fullmatch(r"TOC\s*([1-9])", value, re.I)
        if toc:
            return f"目录 {toc.group(1)}"
        list_bullet = re.fullmatch(r"List\s*Bullet\s*([1-9]?)", value, re.I)
        if list_bullet:
            level = list_bullet.group(1)
            return f"项目符号列表{f' {level}' if level else ''}"
        list_number = re.fullmatch(r"List\s*Number\s*([1-9]?)", value, re.I)
        if list_number:
            level = list_number.group(1)
            return f"编号列表{f' {level}' if level else ''}"
        table = re.fullmatch(r"Table\s*Grid", value, re.I)
        if table:
            return "网格型表格"
        if value in BUILTIN_STYLE_NAMES:
            return BUILTIN_STYLE_NAMES[value]
    return raw_name or raw_id


def style_choice_label(item):
    name = chinese_style_name(item.get("name"), item.get("style_id"))
    style_id = item.get("style_id", "")
    return f"{name}（内部样式：{style_id}）" if style_id else name


def make_choice_maps(items):
    label_to_id = {}
    id_to_label = {}
    for item in items:
        style_id = item.get("style_id", "")
        if not style_id:
            continue
        label = style_choice_label(item)
        label_to_id[label] = style_id
        id_to_label[style_id] = label
    return label_to_id, id_to_label
