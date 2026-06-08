"""样式清理与分析引擎"""

import os
import zipfile

from lxml import etree

from src.ewt.config import NS, W_NS, STYLE_REF_TAGS, STYLE_DEP_TAGS
from src.ewt.utils.helpers import _w_val, _style_id, _style_name, _is_builtin_style, _style_type, _style_dict, _iter_word_xml_names
from src.ewt.utils.word_io import prepare_document, _read_xml, _xml_bytes, _copy_docx_with_replacements
from src.ewt.core.numbering import (
    _used_numbering_ids,
    _clean_unused_numbering_root,
    _deduplicate_numbering_root,
    _style_numbering,
    _numbering_inventory,
)


def _used_style_ids(docx_path):
    """扫描文档正文中实际引用的样式ID"""
    used = set()
    with zipfile.ZipFile(docx_path, "r") as zf:
        for name in _iter_word_xml_names(zf):
            root = _read_xml(zf, name)
            if root is None:
                continue
            for node in root.iter():
                if node.tag in STYLE_REF_TAGS:
                    val = _w_val(node)
                    if val:
                        used.add(val)
    return used


def _keep_style_ids(styles_root, used_ids):
    """根据使用ID + 依赖关系，扩展需要保留的样式集合"""
    styles = _style_dict(styles_root)
    keep = set(used_ids)
    changed = True
    while changed:
        changed = False
        for sid in list(keep):
            style = styles.get(sid)
            if style is None:
                continue
            for dep in style.iter():
                if dep.tag in STYLE_DEP_TAGS:
                    dep_id = _w_val(dep)
                    if dep_id and dep_id not in keep:
                        keep.add(dep_id)
                        changed = True
    for core in ("Normal", "DefaultParagraphFont", "TableNormal", "NoList"):
        if core in styles:
            keep.add(core)
    return keep


def _clean_unused_styles_root(styles_root, used_ids):
    """从 styles.xml 中移除未使用样式"""
    if styles_root is None:
        return 0
    keep = _keep_style_ids(styles_root, used_ids)
    removed = 0
    for style in list(styles_root.findall("w:style", NS)):
        sid = _style_id(style)
        if sid and sid not in keep:
            styles_root.remove(style)
            removed += 1
    return removed


def _repair_broken_style_refs(styles_root, numbering_root=None):
    """Remove style dependencies and numbering links that point to missing definitions."""
    if styles_root is None:
        return 0
    valid_styles = set(_style_dict(styles_root))
    valid_nums = set()
    if numbering_root is not None:
        valid_nums = {
            node.get(f"{{{W_NS}}}numId")
            for node in numbering_root.findall("w:num", NS)
            if node.get(f"{{{W_NS}}}numId")
        }
    repaired = 0
    for style in styles_root.findall("w:style", NS):
        for tag in ("basedOn", "next", "link"):
            node = style.find(f"w:{tag}", NS)
            if node is not None and _w_val(node) not in valid_styles:
                style.remove(node)
                repaired += 1
        num_pr = style.find("w:pPr/w:numPr", NS)
        if num_pr is not None:
            num_id = num_pr.find("w:numId", NS)
            if num_id is not None and valid_nums and _w_val(num_id) not in valid_nums:
                num_pr.getparent().remove(num_pr)
                repaired += 1
    return repaired


def _selected_style_ids(styles_root, selected_ids=None, include_dependencies=True):
    """获取要保留的样式ID集合（支持勾选 + 自动补依赖）"""
    styles = _style_dict(styles_root)
    if not selected_ids:
        return set(styles)
    selected = {sid for sid in selected_ids if sid in styles}
    if include_dependencies:
        selected = _keep_style_ids(styles_root, selected)
    for core in ("Normal", "DefaultParagraphFont", "TableNormal", "NoList"):
        if core in styles:
            selected.add(core)
    return selected


def _filter_styles_root(styles_root, selected_ids=None, include_dependencies=True):
    """按选定ID过滤样式，返回 (移除数量, 自动补选列表)"""
    if styles_root is None or not selected_ids:
        return 0, []
    keep = _selected_style_ids(styles_root, selected_ids, include_dependencies)
    removed = []
    for style in list(styles_root.findall("w:style", NS)):
        sid = _style_id(style)
        if sid and sid not in keep:
            removed.append(sid)
            styles_root.remove(style)
    auto_added = sorted(keep - set(selected_ids))
    return len(removed), auto_added


def _style_properties(style_element):
    """提取样式的字体、段落、编号等属性"""
    rpr = style_element.find("w:rPr", NS)
    ppr = style_element.find("w:pPr", NS)
    font_node = rpr.find("w:rFonts", NS) if rpr is not None else None
    color_node = rpr.find("w:color", NS) if rpr is not None else None
    size_node = rpr.find("w:sz", NS) if rpr is not None else None
    spacing_node = ppr.find("w:spacing", NS) if ppr is not None else None
    ind_node = ppr.find("w:ind", NS) if ppr is not None else None
    return {
        "font": (font_node.get(f"{{{W_NS}}}ascii") or font_node.get(f"{{{W_NS}}}eastAsia") or "") if font_node is not None else "",
        "size": str(int(_w_val(size_node)) / 2).rstrip("0").rstrip(".") if size_node is not None and _w_val(size_node) and _w_val(size_node).isdigit() else "",
        "color": _w_val(color_node) if color_node is not None else "",
        "bold": rpr.find("w:b", NS) is not None if rpr is not None else False,
        "italic": rpr.find("w:i", NS) is not None if rpr is not None else False,
        "before": spacing_node.get(f"{{{W_NS}}}before", "") if spacing_node is not None else "",
        "after": spacing_node.get(f"{{{W_NS}}}after", "") if spacing_node is not None else "",
        "line": spacing_node.get(f"{{{W_NS}}}line", "") if spacing_node is not None else "",
        "left": ind_node.get(f"{{{W_NS}}}left", "") if ind_node is not None else "",
        "hanging": ind_node.get(f"{{{W_NS}}}hanging", "") if ind_node is not None else "",
        "based_on": _w_val(style_element.find("w:basedOn", NS)) if style_element.find("w:basedOn", NS) is not None else "",
        "next": _w_val(style_element.find("w:next", NS)) if style_element.find("w:next", NS) is not None else "",
        "link": _w_val(style_element.find("w:link", NS)) if style_element.find("w:link", NS) is not None else "",
        **_style_numbering(style_element),
    }


def _style_usage_counts(docx_path):
    """统计每个样式在正文中的使用次数"""
    counts = {}
    with zipfile.ZipFile(docx_path, "r") as zf:
        styles_root = _read_xml(zf, "word/styles.xml")
        for sid, style in _style_dict(styles_root).items():
            counts[sid] = {
                "style_id": sid,
                "name": _style_name(style),
                "type": _style_type(style),
                "builtin": _is_builtin_style(style),
                "count": 0,
            }

        for name in _iter_word_xml_names(zf):
            root = _read_xml(zf, name)
            if root is None:
                continue
            for node in root.iter():
                if node.tag in STYLE_REF_TAGS:
                    sid = _w_val(node)
                    if sid:
                        counts.setdefault(
                            sid,
                            {"style_id": sid, "name": sid, "type": "", "builtin": False, "count": 0},
                        )
                        counts[sid]["count"] += 1
    return counts


def inspect_document(file_path):
    """返回完整的样式属性清单 + 编号清单，供预览使用"""
    docx_path, is_temp = prepare_document(file_path)
    try:
        counts = _style_usage_counts(docx_path)
        with zipfile.ZipFile(docx_path, "r") as zf:
            styles_root = _read_xml(zf, "word/styles.xml")
            numbering_root = _read_xml(zf, "word/numbering.xml")
        styles = []
        for sid, style in _style_dict(styles_root).items():
            props = _style_properties(style)
            item = {
                "style_id": sid,
                "name": _style_name(style),
                "type": _style_type(style),
                "builtin": _is_builtin_style(style),
                "count": counts.get(sid, {}).get("count", 0),
            }
            item.update(props)
            styles.append(item)
        styles.sort(key=lambda item: (-item["count"], item["name"]))
        return {"success": True, "styles": styles, "numbering": _numbering_inventory(numbering_root)}
    except Exception as e:
        return {"success": False, "error": str(e), "styles": [], "numbering": []}
    finally:
        if is_temp:
            try:
                os.remove(docx_path)
            except Exception:
                pass


def analyze_document(file_path):
    """分析 .doc/.docx/.dotx 中的样式使用情况"""
    docx_path, is_temp = prepare_document(file_path)
    try:
        counts = _style_usage_counts(docx_path)
        builtin, custom = [], []
        for item in counts.values():
            row = (item["name"], item["count"], item["type"], item["style_id"])
            if item["builtin"]:
                builtin.append(row)
            else:
                custom.append(row)
        builtin.sort(key=lambda x: (-x[1], x[0]))
        custom.sort(key=lambda x: (-x[1], x[0]))
        return True, builtin, custom
    except Exception as e:
        return False, str(e), []
    finally:
        if is_temp:
            try:
                os.remove(docx_path)
            except Exception:
                pass


def process_styles(input_path, output_path, styles_to_remove=None, styles_to_hide=None):
    """清理未使用样式并生成新的 .docx"""
    try:
        docx_path, is_temp = prepare_document(input_path)
        used = _used_style_ids(docx_path)
        with zipfile.ZipFile(docx_path, "r") as zf:
            styles_root = _read_xml(zf, "word/styles.xml")
            numbering_root = _read_xml(zf, "word/numbering.xml")
        removed_count = _clean_unused_styles_root(styles_root, used)
        used_nums = _used_numbering_ids(docx_path, styles_root)
        removed_nums, removed_abs = _clean_unused_numbering_root(numbering_root, used_nums)
        deduplicated = _deduplicate_numbering_root(numbering_root)
        repaired = _repair_broken_style_refs(styles_root, numbering_root)
        replacements = {"word/styles.xml": _xml_bytes(styles_root)}
        if numbering_root is not None:
            replacements["word/numbering.xml"] = _xml_bytes(numbering_root)
        _copy_docx_with_replacements(docx_path, output_path, replacements)
        return True, removed_count + repaired, removed_nums + removed_abs + deduplicated
    except Exception as e:
        return False, str(e), 0
    finally:
        try:
            if "is_temp" in locals() and is_temp:
                os.remove(docx_path)
        except Exception:
            pass


def auto_clean_docx(input_path, output_path):
    """一键清理入口"""
    return process_styles(input_path, output_path)
