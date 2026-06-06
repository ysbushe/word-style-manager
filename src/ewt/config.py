"""
工程文档自动化工具 (Engineering Word Toolkit) — 配置常量
"""

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W_NS = NS["w"]

STYLE_REF_TAGS = {
    f"{{{W_NS}}}pStyle",
    f"{{{W_NS}}}rStyle",
    f"{{{W_NS}}}tblStyle",
}
STYLE_DEP_TAGS = {
    f"{{{W_NS}}}basedOn",
    f"{{{W_NS}}}next",
    f"{{{W_NS}}}link",
}
NUM_REF_TAG = f"{{{W_NS}}}numId"
ILVL_TAG = f"{{{W_NS}}}ilvl"

WORD_TEMPLATE_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
WORD_DOC_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"

NUMBER_FORMATS = {
    "decimal": "1, 2, 3",
    "decimalZero": "01, 02, 03",
    "chineseCounting": "一, 二, 三",
    "chineseCountingThousand": "一, 二, 三",
    "upperLetter": "A, B, C",
    "lowerLetter": "a, b, c",
    "upperRoman": "I, II, III",
    "lowerRoman": "i, ii, iii",
}

DEFAULT_NUMBERING_PRESETS = [
    {
        "name": "1 / 1.1 / 1.1.1",
        "levels": [
            {"format": "%1", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3.%4", "numFmt": "decimal", "suffix": "space"},
        ],
    },
    {
        "name": "第1章 / 1.1 / 1.1.1",
        "levels": [
            {"format": "第%1章", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3.%4", "numFmt": "decimal", "suffix": "space"},
        ],
    },
    {
        "name": "第一章 / 1.1 / 1.1.1",
        "levels": [
            {"format": "第%1章", "numFmt": "chineseCounting", "suffix": "space"},
            {"format": "%1.%2", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3.%4", "numFmt": "decimal", "suffix": "space"},
        ],
    },
    {
        "name": "一、/ （一）/ 1. / （1）",
        "levels": [
            {"format": "%1、", "numFmt": "chineseCounting", "suffix": "space"},
            {"format": "（%2）", "numFmt": "chineseCounting", "suffix": "space"},
            {"format": "%3.", "numFmt": "decimal", "suffix": "space"},
            {"format": "（%4）", "numFmt": "decimal", "suffix": "space"},
        ],
    },
    {
        "name": "A / A.1 / A.1.1",
        "levels": [
            {"format": "%1", "numFmt": "upperLetter", "suffix": "space"},
            {"format": "%1.%2", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3.%4", "numFmt": "decimal", "suffix": "space"},
        ],
    },
    {
        "name": "Article 1 / 1.1 / 1.1.1",
        "levels": [
            {"format": "Article %1", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3", "numFmt": "decimal", "suffix": "space"},
            {"format": "%1.%2.%3.%4", "numFmt": "decimal", "suffix": "space"},
        ],
    },
]

APP_NAME = "Word 样式管理器"
CONFIG_FILE = "cleaner_config.json"
TEMPLATE_DIR = "templates"
