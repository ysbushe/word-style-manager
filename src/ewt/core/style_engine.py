"""样式清理与分析引擎"""

import os
import zipfile
from pathlib import Path
from uuid import uuid4

from lxml import etree

from src.ewt.config import NS, W_NS, STYLE_REF_TAGS, STYLE_DEP_TAGS
from src.ewt.utils.helpers import _w_val, _style_id, _style_name, _is_builtin_style, _style_type, _style_dict, _iter_word_xml_names
from src.ewt.utils.word_io import prepare_document, _read_xml, _xml_bytes, _copy_docx_with_replacements
from src.ewt.utils.display import chinese_style_name
from src.ewt.core.numbering import (
    _clean_unused_numbering_root,
    _deduplicate_numbering_root,
    _style_numbering,
    _numbering_inventory,
)

CORE_STYLE_IDS = {"Normal", "DefaultParagraphFont", "TableNormal", "NoList"}
NUMBERING_STYLE_TAGS = {
    f"{{{W_NS}}}pStyle",
    f"{{{W_NS}}}styleLink",
    f"{{{W_NS}}}numStyleLink",
}
DEFAULT_STYLE_TAGS = {
    f"{{{W_NS}}}defaultTableStyle",
}


def _style_refs_in_root(root, tags):
    if root is None:
        return set()
    return {
        value
        for node in root.iter()
        if node.tag in tags
        for value in [_w_val(node)]
        if value
    }


def _content_style_ids_from_archive(zf):
    """收集正文、页眉页脚、批注、脚注和设置中的样式引用。"""
    used = set()
    for name in _iter_word_xml_names(zf):
        root = _read_xml(zf, name)
        used.update(_style_refs_in_root(root, STYLE_REF_TAGS | DEFAULT_STYLE_TAGS))
    settings_root = _read_xml(zf, "word/settings.xml")
    used.update(_style_refs_in_root(settings_root, DEFAULT_STYLE_TAGS))
    return used


def _referenced_style_ids_from_archive(zf):
    """收集正文、编号和设置中所有显式样式引用。"""
    used = _content_style_ids_from_archive(zf)
    used.update(
        _style_refs_in_root(
            _read_xml(zf, "word/numbering.xml"),
            NUMBERING_STYLE_TAGS,
        )
    )
    return used


def _content_numbering_ids_from_archive(zf):
    used = set()
    for name in _iter_word_xml_names(zf):
        used.update(
            _style_refs_in_root(
                _read_xml(zf, name),
                {f"{{{W_NS}}}numId"},
            )
        )
    return used - {"0"}


def _reachable_style_and_numbering_ids(
    styles_root,
    numbering_root,
    content_style_ids,
    content_numbering_ids,
    dependency_tags,
):
    """从正文引用出发，计算真实可达的样式与编号集合。"""
    style_ids = set(content_style_ids)
    num_ids = set(content_numbering_ids) - {"0"}
    num_to_abs = {}
    abstract_by_id = {}
    num_by_id = {}
    if numbering_root is not None:
        for num in numbering_root.findall("w:num", NS):
            num_id = num.get(f"{{{W_NS}}}numId")
            ref = num.find("w:abstractNumId", NS)
            if num_id:
                num_by_id[num_id] = num
                if ref is not None and _w_val(ref):
                    num_to_abs[num_id] = _w_val(ref)
        abstract_by_id = {
            node.get(f"{{{W_NS}}}abstractNumId"): node
            for node in numbering_root.findall("w:abstractNum", NS)
            if node.get(f"{{{W_NS}}}abstractNumId")
        }

    while True:
        previous = (set(style_ids), set(num_ids))
        style_ids = _keep_style_ids(styles_root, style_ids, dependency_tags)
        styles = _style_dict(styles_root)
        for sid in style_ids:
            style = styles.get(sid)
            if style is not None:
                num_ids.update(
                    _style_refs_in_root(style, {f"{{{W_NS}}}numId"}) - {"0"}
                )
        for num_id in list(num_ids):
            num = num_by_id.get(num_id)
            abstract = abstract_by_id.get(num_to_abs.get(num_id))
            style_ids.update(_style_refs_in_root(num, NUMBERING_STYLE_TAGS))
            style_ids.update(_style_refs_in_root(abstract, NUMBERING_STYLE_TAGS))
        if previous == (style_ids, num_ids):
            return style_ids, num_ids


def _used_style_ids(docx_path):
    """扫描整个文档包中实际引用的样式 ID。"""
    with zipfile.ZipFile(docx_path, "r") as zf:
        return _referenced_style_ids_from_archive(zf)


def _keep_style_ids(styles_root, used_ids, dependency_tags=None):
    """根据使用ID + 依赖关系，扩展需要保留的样式集合"""
    styles = _style_dict(styles_root)
    keep = set(used_ids)
    dependency_tags = dependency_tags or STYLE_DEP_TAGS
    changed = True
    while changed:
        changed = False
        for sid in list(keep):
            style = styles.get(sid)
            if style is None:
                continue
            for dep in style.iter():
                if dep.tag in dependency_tags:
                    dep_id = _w_val(dep)
                    if dep_id and dep_id not in keep:
                        keep.add(dep_id)
                        changed = True
    for core in CORE_STYLE_IDS:
        if core in styles:
            keep.add(core)
    return keep


def _clean_unused_styles_root(
    styles_root,
    used_ids,
    preserve_builtin=False,
    dependency_tags=None,
):
    """从 styles.xml 中移除未使用样式"""
    if styles_root is None:
        return 0
    keep = _keep_style_ids(styles_root, used_ids, dependency_tags)
    removed = 0
    for style in list(styles_root.findall("w:style", NS)):
        sid = _style_id(style)
        if sid and sid not in keep:
            if preserve_builtin and _is_builtin_style(style):
                continue
            styles_root.remove(style)
            removed += 1
    return removed


def _prune_optional_style_links(styles_root, required_ids):
    """深度模式中移除不会影响现有正文排版的 next/link 关系。"""
    if styles_root is None:
        return 0
    changed = 0
    for style in styles_root.findall("w:style", NS):
        for tag in ("next", "link"):
            node = style.find(f"w:{tag}", NS)
            if node is not None and _w_val(node) not in required_ids:
                style.remove(node)
                changed += 1
    return changed


def _sync_secondary_styles_root(styles_root, effects_root):
    """让 stylesWithEffects.xml 与主样式表保持相同的样式集合。"""
    if effects_root is None:
        return 0
    valid_ids = set(_style_dict(styles_root))
    removed = 0
    for style in list(effects_root.findall("w:style", NS)):
        sid = _style_id(style)
        if sid and sid not in valid_ids:
            effects_root.remove(style)
            removed += 1
    return removed


def _hide_unused_builtin_styles_root(styles_root, used_ids):
    """隐藏未使用系统样式，并允许样式再次被使用时自动显示。"""
    if styles_root is None:
        return 0
    keep = _keep_style_ids(styles_root, used_ids)
    hidden = 0
    for style in styles_root.findall("w:style", NS):
        sid = _style_id(style)
        if not sid or sid in keep or not _is_builtin_style(style):
            continue
        changed = False
        if style.find("w:semiHidden", NS) is None:
            etree.SubElement(style, f"{{{W_NS}}}semiHidden")
            changed = True
        if style.find("w:unhideWhenUsed", NS) is None:
            etree.SubElement(style, f"{{{W_NS}}}unhideWhenUsed")
            changed = True
        qformat = style.find("w:qFormat", NS)
        if qformat is not None:
            style.remove(qformat)
            changed = True
        if changed:
            hidden += 1
    return hidden


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
            if num_id is not None and _w_val(num_id) not in valid_nums:
                num_pr.getparent().remove(num_pr)
                repaired += 1
    return repaired


def _style_validation_state(docx_path):
    """返回复检需要的样式、编号和悬空引用状态。"""
    with zipfile.ZipFile(docx_path, "r") as zf:
        bad_member = zf.testzip()
        if bad_member:
            raise RuntimeError(f"输出文件压缩包损坏：{bad_member}")
        styles_root = _read_xml(zf, "word/styles.xml")
        if styles_root is None:
            raise RuntimeError("输出文件缺少 word/styles.xml")
        effects_root = _read_xml(zf, "word/stylesWithEffects.xml")
        numbering_root = _read_xml(zf, "word/numbering.xml")
        style_ids = set(_style_dict(styles_root))
        effects_ids = set(_style_dict(effects_root))
        style_refs = _referenced_style_ids_from_archive(zf)
        num_ids = {
            node.get(f"{{{W_NS}}}numId")
            for node in numbering_root.findall("w:num", NS)
            if node.get(f"{{{W_NS}}}numId")
        } if numbering_root is not None else set()
        num_refs = set()
        for name in _iter_word_xml_names(zf):
            num_refs.update(
                _style_refs_in_root(_read_xml(zf, name), {f"{{{W_NS}}}numId"})
            )
        if styles_root is not None:
            num_refs.update(_style_refs_in_root(styles_root, {f"{{{W_NS}}}numId"}))
    return {
        "style_ids": style_ids,
        "style_refs": style_refs,
        "effects_extra": effects_ids - style_ids,
        "num_ids": num_ids,
        "num_refs": num_refs,
        "dangling_styles": style_refs - style_ids,
        "dangling_nums": num_refs - num_ids,
    }


def _validate_clean_output(output_path, source_state):
    """确认清理没有新增悬空引用，且双样式表保持一致。"""
    output_state = _style_validation_state(output_path)
    new_style_breaks = output_state["dangling_styles"] - source_state["dangling_styles"]
    new_num_breaks = output_state["dangling_nums"] - source_state["dangling_nums"]
    if output_state["effects_extra"]:
        raise RuntimeError(
            "辅助样式表仍包含已删除样式："
            + ", ".join(sorted(output_state["effects_extra"])[:8])
        )
    if new_style_breaks:
        raise RuntimeError(
            "清理产生了新的悬空样式引用："
            + ", ".join(sorted(new_style_breaks)[:8])
        )
    if new_num_breaks:
        raise RuntimeError(
            "清理产生了新的悬空编号引用："
            + ", ".join(sorted(new_num_breaks)[:8])
        )
    return output_state


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
        "line_rule": spacing_node.get(f"{{{W_NS}}}lineRule", "") if spacing_node is not None else "",
        "left": ind_node.get(f"{{{W_NS}}}left", "") if ind_node is not None else "",
        "hanging": ind_node.get(f"{{{W_NS}}}hanging", "") if ind_node is not None else "",
        "based_on": _w_val(style_element.find("w:basedOn", NS)) if style_element.find("w:basedOn", NS) is not None else "",
        "next": _w_val(style_element.find("w:next", NS)) if style_element.find("w:next", NS) is not None else "",
        "link": _w_val(style_element.find("w:link", NS)) if style_element.find("w:link", NS) is not None else "",
        **_style_numbering(style_element),
    }


def _style_usage_counts_from_archive(zf, styles_root):
    """在已打开的文档包中统计样式使用次数。"""
    counts = {}
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


def _style_usage_counts(docx_path):
    """统计每个样式在正文中的使用次数"""
    with zipfile.ZipFile(docx_path, "r") as zf:
        return _style_usage_counts_from_archive(zf, _read_xml(zf, "word/styles.xml"))


def inspect_document(file_path):
    """返回完整的样式属性清单 + 编号清单，供预览使用"""
    docx_path, is_temp = prepare_document(file_path)
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            styles_root = _read_xml(zf, "word/styles.xml")
            numbering_root = _read_xml(zf, "word/numbering.xml")
            counts = _style_usage_counts_from_archive(zf, styles_root)
            referenced_ids = _referenced_style_ids_from_archive(zf)
        directly_used = {
            sid
            for sid, usage in counts.items()
            if usage.get("count", 0) > 0
        } | referenced_ids
        required_ids = _keep_style_ids(styles_root, directly_used)
        styles = []
        for sid, style in _style_dict(styles_root).items():
            props = _style_properties(style)
            count = counts.get(sid, {}).get("count", 0)
            builtin = _is_builtin_style(style)
            dependency_required = count == 0 and sid in required_ids
            item = {
                "style_id": sid,
                "name": _style_name(style),
                "display_name": chinese_style_name(_style_name(style), sid),
                "type": _style_type(style),
                "builtin": builtin,
                "hidden": style.find("w:semiHidden", NS) is not None,
                "count": count,
                "dependency_required": dependency_required,
                "cleanable": count == 0 and not builtin and not dependency_required,
                "hideable": count == 0 and builtin and not dependency_required,
            }
            item.update(props)
            styles.append(item)
        type_order = {"paragraph": 0, "character": 1, "table": 2, "numbering": 3}
        styles.sort(
            key=lambda item: (
                0 if item.get("count", 0) == 0 else 1,
                type_order.get(item.get("type"), 9),
                item.get("name", "").casefold(),
            )
        )
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


def process_styles(input_path, output_path, styles_to_remove=None, styles_to_hide=None, options=None):
    """在不覆盖源文件的前提下清理样式、列表和多级列表。"""
    options = {
        "mode": "safe",
        "unused_styles": True,
        "unused_numbering": True,
        "unused_multilevel": True,
        "deduplicate_numbering": True,
        "repair_links": True,
        "hide_unused_builtin": True,
        **(options or {}),
    }
    source = Path(input_path).resolve()
    output = Path(output_path).resolve()
    if source == output:
        return {
            "success": False,
            "error": "为保护源文档，清理结果不能覆盖源文件。",
            "output": "",
        }
    try:
        docx_path, is_temp = prepare_document(input_path)
        source_state = _style_validation_state(docx_path)
        with zipfile.ZipFile(docx_path, "r") as zf:
            styles_root = _read_xml(zf, "word/styles.xml")
            effects_root = _read_xml(zf, "word/stylesWithEffects.xml")
            numbering_root = _read_xml(zf, "word/numbering.xml")
            content_used = _content_style_ids_from_archive(zf)
            content_num_ids = _content_numbering_ids_from_archive(zf)
        if styles_root is None:
            raise RuntimeError("源文档缺少样式表，无法安全清理。")

        deep_mode = options.get("mode") == "deep"
        dependency_tags = (
            {f"{{{W_NS}}}basedOn"}
            if deep_mode
            else STYLE_DEP_TAGS
        )
        used, used_nums = _reachable_style_and_numbering_ids(
            styles_root,
            numbering_root,
            content_used,
            content_num_ids,
            dependency_tags,
        )
        pruned_links = 0
        if deep_mode:
            required_for_layout = _keep_style_ids(
                styles_root,
                used,
                {f"{{{W_NS}}}basedOn"},
            )
            pruned_links = _prune_optional_style_links(
                styles_root,
                required_for_layout,
            )

        preserve_builtin = not deep_mode and options["hide_unused_builtin"]
        hidden_count = (
            _hide_unused_builtin_styles_root(styles_root, used)
            if preserve_builtin
            else 0
        )
        removed_count = (
            _clean_unused_styles_root(
                styles_root,
                used,
                preserve_builtin=preserve_builtin,
                dependency_tags=dependency_tags,
            )
            if options["unused_styles"]
            else 0
        )
        removed_nums, removed_abs = _clean_unused_numbering_root(
            numbering_root,
            used_nums,
            remove_single_level=options["unused_numbering"],
            remove_multilevel=options["unused_multilevel"],
        )
        deduplicated = _deduplicate_numbering_root(numbering_root) if options["deduplicate_numbering"] else 0
        repaired = _repair_broken_style_refs(styles_root, numbering_root) if options["repair_links"] else 0

        if options["unused_styles"]:
            used, used_nums = _reachable_style_and_numbering_ids(
                styles_root,
                numbering_root,
                content_used,
                content_num_ids,
                dependency_tags,
            )
            removed_count += _clean_unused_styles_root(
                styles_root,
                used,
                preserve_builtin=preserve_builtin,
                dependency_tags=dependency_tags,
            )

        if effects_root is not None:
            if deep_mode:
                _prune_optional_style_links(effects_root, set(_style_dict(styles_root)))
            elif preserve_builtin:
                _hide_unused_builtin_styles_root(effects_root, used)
        effects_removed = _sync_secondary_styles_root(styles_root, effects_root)
        if effects_root is not None and options["repair_links"]:
            repaired += _repair_broken_style_refs(effects_root, numbering_root)
        replacements = {"word/styles.xml": _xml_bytes(styles_root)}
        if effects_root is not None:
            replacements["word/stylesWithEffects.xml"] = _xml_bytes(effects_root)
        if numbering_root is not None:
            replacements["word/numbering.xml"] = _xml_bytes(numbering_root)
        candidate = output.with_name(f".{output.name}.{uuid4().hex}.verify")
        _copy_docx_with_replacements(docx_path, candidate, replacements)
        output_state = _validate_clean_output(candidate, source_state)
        output.parent.mkdir(parents=True, exist_ok=True)
        os.replace(candidate, output)
        return {
            "success": True,
            "output": str(output),
            "mode": "deep" if deep_mode else "safe",
            "removed_styles": removed_count,
            "hidden_builtin": hidden_count,
            "pruned_links": pruned_links,
            "repaired_links": repaired,
            "effects_removed": effects_removed,
            "removed_numbering": removed_nums + removed_abs,
            "deduplicated_numbering": deduplicated,
            "remaining_styles": len(output_state["style_ids"]),
            "verified": True,
        }
    except Exception as e:
        try:
            if "candidate" in locals() and candidate.exists():
                candidate.unlink()
        except OSError:
            pass
        return {
            "success": False,
            "error": str(e),
            "output": "",
        }
    finally:
        try:
            if "is_temp" in locals() and is_temp:
                os.remove(docx_path)
        except Exception:
            pass


def auto_clean_docx(input_path, output_path, options=None):
    """一键清理入口"""
    return process_styles(input_path, output_path, options=options)
