"""XML / OOXML 底层辅助函数"""

from src.ewt.config import NS, W_NS, STYLE_REF_TAGS, STYLE_DEP_TAGS


def _w_val(element):
    """读取 w:val 属性值"""
    return element.get(f"{{{W_NS}}}val")


def _style_id(style_element):
    return style_element.get(f"{{{W_NS}}}styleId")


def _style_name(style_element):
    node = style_element.find("w:name", NS)
    return _w_val(node) if node is not None else _style_id(style_element)


def _is_builtin_style(style_element):
    custom = style_element.get(f"{{{W_NS}}}customStyle", "").lower()
    if custom in {"1", "true", "on"}:
        return False
    return style_element.find("w:customStyle", NS) is None


def _style_type(style_element):
    return style_element.get(f"{{{W_NS}}}type", "")


def _find_val(parent, path):
    node = parent.find(path, NS) if parent is not None else None
    return _w_val(node) if node is not None else ""


def _ensure_child(parent, tag):
    node = parent.find(f"w:{tag}", NS)
    if node is None:
        from lxml import etree  # noqa
        node = etree.SubElement(parent, f"{{{W_NS}}}{tag}")
    return node


def _ensure_path(parent, tags):
    node = parent
    for tag in tags:
        node = _ensure_child(node, tag)
    return node


def _style_dict(styles_root):
    """构建 styleId → style_element 字典"""
    if styles_root is None:
        return {}
    return {_style_id(node): node for node in styles_root.findall("w:style", NS) if _style_id(node)}


def _iter_word_xml_names(zf):
    """遍历 docx 包内所有 word/ 下的 xml 文件"""
    for name in zf.namelist():
        lower = name.lower()
        if lower.startswith("word/") and lower.endswith(".xml"):
            if lower not in {"word/styles.xml", "word/numbering.xml"}:
                yield name
