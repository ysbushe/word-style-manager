"""编号 / 多级列表处理引擎"""

from copy import deepcopy
from lxml import etree

from src.ewt.config import NS, W_NS, NUM_REF_TAG
from src.ewt.utils.helpers import _w_val, _find_val, _style_dict
from src.ewt.utils.word_io import _read_xml


def _numbering_maps(numbering_root):
    """构建 numId→abstractNumId 映射及使用中的 abstractNum 集合"""
    num_to_abs = {}
    if numbering_root is None:
        return num_to_abs, set()
    for num in numbering_root.findall("w:num", NS):
        num_id = num.get(f"{{{W_NS}}}numId")
        abs_node = num.find("w:abstractNumId", NS)
        if num_id and abs_node is not None:
            num_to_abs[num_id] = _w_val(abs_node)
    return num_to_abs, set(num_to_abs.values())


def _used_numbering_ids(docx_path, styles_root=None):
    """扫描正文 + 样式中的 numId 引用"""
    used_nums = set()
    import zipfile
    from src.ewt.utils.helpers import _iter_word_xml_names

    with zipfile.ZipFile(docx_path, "r") as zf:
        for name in _iter_word_xml_names(zf):
            root = _read_xml(zf, name)
            if root is None:
                continue
            for node in root.iter():
                if node.tag == NUM_REF_TAG:
                    val = _w_val(node)
                    if val:
                        used_nums.add(val)
    if styles_root is not None:
        for node in styles_root.findall(".//w:numId", NS):
            val = _w_val(node)
            if val:
                used_nums.add(val)
    return used_nums


def _clean_unused_numbering_root(
    numbering_root,
    used_nums,
    remove_single_level=True,
    remove_multilevel=True,
):
    """清理未使用编号，并允许分别控制单级编号和多级列表。"""
    if numbering_root is None:
        return 0, 0
    num_to_abs, _ = _numbering_maps(numbering_root)
    abstract_by_id = {
        node.get(f"{{{W_NS}}}abstractNumId"): node
        for node in numbering_root.findall("w:abstractNum", NS)
    }

    def should_remove(abstract_id):
        abstract = abstract_by_id.get(abstract_id)
        level_count = len(abstract.findall("w:lvl", NS)) if abstract is not None else 1
        return remove_multilevel if level_count > 1 else remove_single_level

    removed_nums = 0
    removed_abs = 0
    for num in list(numbering_root.findall("w:num", NS)):
        num_id = num.get(f"{{{W_NS}}}numId")
        abstract_id = num_to_abs.get(num_id)
        if num_id and num_id not in used_nums and should_remove(abstract_id):
            numbering_root.remove(num)
            removed_nums += 1

    remaining_abs = {
        _w_val(ref)
        for ref in numbering_root.findall("w:num/w:abstractNumId", NS)
        if _w_val(ref)
    }
    for abstract in list(numbering_root.findall("w:abstractNum", NS)):
        abs_id = abstract.get(f"{{{W_NS}}}abstractNumId")
        if abs_id and abs_id not in remaining_abs and should_remove(abs_id):
            numbering_root.remove(abstract)
            removed_abs += 1
    return removed_nums, removed_abs


def _deduplicate_numbering_root(numbering_root):
    """Merge duplicate abstract numbering definitions and remap num references."""
    if numbering_root is None:
        return 0
    seen = {}
    remap = {}
    removed = 0
    for abstract in list(numbering_root.findall("w:abstractNum", NS)):
        abs_id = abstract.get(f"{{{W_NS}}}abstractNumId", "")
        clone = deepcopy(abstract)
        clone.attrib.pop(f"{{{W_NS}}}abstractNumId", None)
        signature = etree.tostring(clone, encoding="UTF-8")
        if signature in seen:
            remap[abs_id] = seen[signature]
            numbering_root.remove(abstract)
            removed += 1
        else:
            seen[signature] = abs_id
    for num in numbering_root.findall("w:num", NS):
        ref = num.find("w:abstractNumId", NS)
        if ref is not None:
            old = _w_val(ref)
            if old in remap:
                ref.set(f"{{{W_NS}}}val", remap[old])
    return removed


def _style_numbering(style_element):
    """提取样式绑定的编号信息"""
    num = style_element.find(".//w:numId", NS)
    ilvl = style_element.find(".//w:ilvl", NS)
    return {
        "num_id": _w_val(num) if num is not None else "",
        "level": _w_val(ilvl) if ilvl is not None else "",
    }


def _numbering_inventory(numbering_root):
    """提取完整编号/多级列表清单"""
    if numbering_root is None:
        return []
    items = []
    num_to_abs, _ = _numbering_maps(numbering_root)
    for abstract in numbering_root.findall("w:abstractNum", NS):
        abs_id = abstract.get(f"{{{W_NS}}}abstractNumId", "")
        levels = []
        for lvl in abstract.findall("w:lvl", NS):
            ilvl = lvl.get(f"{{{W_NS}}}ilvl", "")
            levels.append(
                {
                    "level": ilvl,
                    "format": _find_val(lvl, "w:lvlText"),
                    "numFmt": _find_val(lvl, "w:numFmt"),
                    "suffix": _find_val(lvl, "w:suff"),
                    "linked_style": _find_val(lvl, "w:pStyle"),
                    "restart": _find_val(lvl, "w:lvlRestart"),
                    "start": _find_val(lvl, "w:start") or "1",
                }
            )
        nums = [num_id for num_id, aid in num_to_abs.items() if aid == abs_id]
        items.append({"abstract_id": abs_id, "num_ids": nums, "levels": levels})
    return items


def _merge_numbering(source_root, target_root, source_styles_root):
    """合并源编号定义到目标，返回 (merged_root, 新增数量)"""
    if source_root is None:
        return target_root, 0
    if target_root is None:
        return deepcopy(source_root), len(source_root.findall("w:num", NS))

    def max_id(root, tag, attr):
        values = []
        for node in root.findall(f"w:{tag}", NS):
            val = node.get(f"{{{W_NS}}}{attr}")
            if val and val.isdigit():
                values.append(int(val))
        return max(values) if values else 0

    abstract_map = {}
    num_map = {}
    next_abstract = max_id(target_root, "abstractNum", "abstractNumId") + 1
    next_num = max_id(target_root, "num", "numId") + 1

    for abstract in source_root.findall("w:abstractNum", NS):
        old = abstract.get(f"{{{W_NS}}}abstractNumId")
        new = str(next_abstract)
        next_abstract += 1
        abstract_map[old] = new
        copied = deepcopy(abstract)
        copied.set(f"{{{W_NS}}}abstractNumId", new)
        target_root.append(copied)

    for num in source_root.findall("w:num", NS):
        old = num.get(f"{{{W_NS}}}numId")
        new = str(next_num)
        next_num += 1
        num_map[old] = new
        copied = deepcopy(num)
        copied.set(f"{{{W_NS}}}numId", new)
        abs_node = copied.find("w:abstractNumId", NS)
        if abs_node is not None:
            abs_node.set(f"{{{W_NS}}}val", abstract_map.get(_w_val(abs_node), _w_val(abs_node)))
        target_root.append(copied)

    for num_id in source_styles_root.findall(".//w:numId", NS):
        val = _w_val(num_id)
        if val in num_map:
            num_id.set(f"{{{W_NS}}}val", num_map[val])

    return target_root, len(num_map)
