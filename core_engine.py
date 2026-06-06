from __future__ import annotations

import json
import os
import tempfile
import zipfile
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from docx import Document
from lxml import etree

try:
    import win32com.client
    _HAS_WIN32COM = True
except ImportError:
    _HAS_WIN32COM = False


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


def convert_doc_to_docx(doc_path):
    """Convert .doc to a temporary .docx file. Requires Microsoft Word."""
    if not _HAS_WIN32COM:
        raise RuntimeError("处理 .doc 文件需要安装 pywin32 库。")
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"找不到文件: {doc_path}")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(doc_path))
        name_no_ext = os.path.splitext(os.path.basename(doc_path))[0]
        docx_path = os.path.join(tempfile.gettempdir(), f"{name_no_ext}_temp_{os.getpid()}.docx")
        doc.SaveAs(docx_path, FileFormat=16)
        return docx_path
    except Exception as e:
        raise RuntimeError(f"Word 转换失败: {e}")
    finally:
        try:
            if doc:
                doc.Close(False)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass


def prepare_document(file_path):
    """Return a docx path and whether it is temporary."""
    if file_path.lower().endswith((".docx", ".dotx")):
        return file_path, False
    if file_path.lower().endswith(".doc"):
        return convert_doc_to_docx(file_path), True
    raise ValueError(f"不支持的文件格式，仅支持 .doc、.docx 和 .dotx: {file_path}")


def _read_xml(zf, name):
    try:
        return etree.fromstring(zf.read(name))
    except KeyError:
        return None


def _xml_bytes(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def _w_val(element):
    return element.get(f"{{{W_NS}}}val")


def _style_id(style_element):
    return style_element.get(f"{{{W_NS}}}styleId")


def _style_name(style_element):
    node = style_element.find("w:name", NS)
    return _w_val(node) if node is not None else _style_id(style_element)


def _is_builtin_style(style_element):
    return style_element.find("w:customStyle", NS) is None


def _style_type(style_element):
    return style_element.get(f"{{{W_NS}}}type", "")


def _find_val(parent, path):
    node = parent.find(path, NS) if parent is not None else None
    return _w_val(node) if node is not None else ""


def _ensure_child(parent, tag):
    node = parent.find(f"w:{tag}", NS)
    if node is None:
        node = etree.SubElement(parent, f"{{{W_NS}}}{tag}")
    return node


def _ensure_path(parent, tags):
    node = parent
    for tag in tags:
        node = _ensure_child(node, tag)
    return node


def _style_dict(styles_root):
    if styles_root is None:
        return {}
    return {_style_id(node): node for node in styles_root.findall("w:style", NS) if _style_id(node)}


def _iter_word_xml_names(zf):
    for name in zf.namelist():
        lower = name.lower()
        if lower.startswith("word/") and lower.endswith(".xml"):
            if lower not in {"word/styles.xml", "word/numbering.xml"}:
                yield name


def _used_style_ids(docx_path):
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


def _numbering_maps(numbering_root):
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
    used_nums = set()
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


def _clean_unused_numbering_root(numbering_root, used_nums):
    if numbering_root is None:
        return 0, 0
    num_to_abs, _ = _numbering_maps(numbering_root)
    used_abs = {num_to_abs[num_id] for num_id in used_nums if num_id in num_to_abs}
    removed_nums = 0
    removed_abs = 0
    for num in list(numbering_root.findall("w:num", NS)):
        num_id = num.get(f"{{{W_NS}}}numId")
        if num_id and num_id not in used_nums:
            numbering_root.remove(num)
            removed_nums += 1
    for abstract in list(numbering_root.findall("w:abstractNum", NS)):
        abs_id = abstract.get(f"{{{W_NS}}}abstractNumId")
        if abs_id and abs_id not in used_abs:
            numbering_root.remove(abstract)
            removed_abs += 1
    return removed_nums, removed_abs


def _style_numbering(style_element):
    num = style_element.find(".//w:numId", NS)
    ilvl = style_element.find(".//w:ilvl", NS)
    return {
        "num_id": _w_val(num) if num is not None else "",
        "level": _w_val(ilvl) if ilvl is not None else "",
    }


def _style_properties(style_element):
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
        "based_on": _find_val(style_element, "w:basedOn"),
        "next": _find_val(style_element, "w:next"),
        "link": _find_val(style_element, "w:link"),
        **_style_numbering(style_element),
    }


def _numbering_inventory(numbering_root):
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


def _style_usage_counts(docx_path):
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


def _keep_style_ids(styles_root, used_ids):
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


def _selected_style_ids(styles_root, selected_ids=None, include_dependencies=True):
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


def inspect_document(file_path):
    """Return rich style and numbering inventory for previews."""
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


def _copy_docx_with_replacements(input_path, output_path, replacements, as_template=False):
    output_path = str(output_path)
    with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        existing_names = set(zin.namelist())
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in replacements:
                data = replacements[item.filename]
            elif as_template and item.filename == "[Content_Types].xml":
                root = etree.fromstring(data)
                for override in root.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override"):
                    if override.get("PartName") == "/word/document.xml":
                        override.set("ContentType", WORD_TEMPLATE_TYPE)
                data = _xml_bytes(root)
            zout.writestr(item, data)
        for name, data in replacements.items():
            if name not in existing_names:
                zout.writestr(name, data)


def _empty_document_xml(source_root):
    body = source_root.find("w:body", NS)
    sect_pr = body.find("w:sectPr", NS) if body is not None else None
    doc = etree.Element(f"{{{W_NS}}}document", nsmap=source_root.nsmap)
    new_body = etree.SubElement(doc, f"{{{W_NS}}}body")
    if sect_pr is not None:
        new_body.append(deepcopy(sect_pr))
    return doc


def analyze_document(file_path):
    """Analyze styles for .doc/.docx/.dotx."""
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
    """Clean unused styles and save a new .docx file."""
    try:
        docx_path, is_temp = prepare_document(input_path)
        used = _used_style_ids(docx_path)
        with zipfile.ZipFile(docx_path, "r") as zf:
            styles_root = _read_xml(zf, "word/styles.xml")
            numbering_root = _read_xml(zf, "word/numbering.xml")
        removed_count = _clean_unused_styles_root(styles_root, used)
        used_nums = _used_numbering_ids(docx_path, styles_root)
        removed_nums, removed_abs = _clean_unused_numbering_root(numbering_root, used_nums)
        replacements = {"word/styles.xml": _xml_bytes(styles_root)}
        if numbering_root is not None:
            replacements["word/numbering.xml"] = _xml_bytes(numbering_root)
        _copy_docx_with_replacements(docx_path, output_path, replacements)
        return True, removed_count, removed_nums + removed_abs
    except Exception as e:
        return False, str(e), 0
    finally:
        try:
            if "is_temp" in locals() and is_temp:
                os.remove(docx_path)
        except Exception:
            pass


def auto_clean_docx(input_path, output_path):
    return process_styles(input_path, output_path)


def _next_output_path(path, suffix, ext=".docx"):
    base = Path(path)
    candidate = base.with_name(f"{base.stem}{suffix}{ext}")
    index = 2
    while candidate.exists():
        candidate = base.with_name(f"{base.stem}{suffix}_{index}{ext}")
        index += 1
    return str(candidate)


def make_clean_style_carrier(source_path, output_path, as_template=False, selected_styles=None, include_dependencies=True):
    """Create an empty style carrier document/template from source."""
    docx_path, is_temp = prepare_document(source_path)
    try:
        used = _used_style_ids(docx_path)
        with zipfile.ZipFile(docx_path, "r") as zf:
            styles_root = _read_xml(zf, "word/styles.xml")
            document_root = _read_xml(zf, "word/document.xml")
        removed = _clean_unused_styles_root(styles_root, used)
        filtered, auto_added = _filter_styles_root(styles_root, selected_styles, include_dependencies)
        replacements = {
            "word/styles.xml": _xml_bytes(styles_root),
            "word/document.xml": _xml_bytes(_empty_document_xml(document_root)),
        }
        _copy_docx_with_replacements(docx_path, output_path, replacements, as_template=as_template)
        return {"success": True, "output": output_path, "removed_unused": removed, "filtered": filtered, "auto_added": auto_added}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        if is_temp:
            try:
                os.remove(docx_path)
            except Exception:
                pass


def export_style_template(source_path, output_dir, template_name=None, formats=None, library_dir=None, selected_styles=None, include_dependencies=True):
    """Export source styles as .dotx/.docx and optional library entry."""
    formats = formats or ["dotx", "docx", "library"]
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    source = Path(source_path)
    safe_name = template_name or f"{source.stem}_样式模板"
    outputs = {}
    removed_unused = 0

    if "dotx" in formats:
        dotx_path = output_dir / f"{safe_name}.dotx"
        result = make_clean_style_carrier(source_path, str(dotx_path), as_template=True, selected_styles=selected_styles, include_dependencies=include_dependencies)
        if not result["success"]:
            return result
        outputs["dotx"] = str(dotx_path)
        removed_unused = result.get("removed_unused", 0)

    if "docx" in formats:
        docx_path = output_dir / f"{safe_name}.docx"
        result = make_clean_style_carrier(source_path, str(docx_path), as_template=False, selected_styles=selected_styles, include_dependencies=include_dependencies)
        if not result["success"]:
            return result
        outputs["docx"] = str(docx_path)
        removed_unused = result.get("removed_unused", removed_unused)

    if "library" in formats:
        lib_dir = Path(library_dir or output_dir / "templates")
        lib_dir.mkdir(parents=True, exist_ok=True)
        lib_path = lib_dir / f"{safe_name}.dotx"
        result = make_clean_style_carrier(source_path, str(lib_path), as_template=True, selected_styles=selected_styles, include_dependencies=include_dependencies)
        if not result["success"]:
            return result
        outputs["library"] = str(lib_path)
        _update_template_index(lib_dir, safe_name, source_path, lib_path)

    return {"success": True, "outputs": outputs, "removed_unused": removed_unused}


def _update_template_index(library_dir, name, source_path, template_path):
    index_path = Path(library_dir) / "template_index.json"
    try:
        items = json.loads(index_path.read_text(encoding="utf-8")) if index_path.exists() else []
    except Exception:
        items = []
    items = [item for item in items if item.get("name") != name]
    items.append(
        {
            "name": name,
            "source": str(source_path),
            "path": str(template_path),
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
    )
    index_path.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")


def list_template_library(library_dir):
    index_path = Path(library_dir) / "template_index.json"
    if not index_path.exists():
        return []
    try:
        return json.loads(index_path.read_text(encoding="utf-8"))
    except Exception:
        return []


def _merge_numbering(source_root, target_root, source_styles_root):
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


def _import_styles(source_styles_root, target_styles_root, conflict="overwrite"):
    if target_styles_root is None:
        return deepcopy(source_styles_root), {"added": 0, "overwritten": 0, "skipped": 0, "renamed": 0}

    stats = {"added": 0, "overwritten": 0, "skipped": 0, "renamed": 0}
    existing = _style_dict(target_styles_root)

    for style in source_styles_root.findall("w:style", NS):
        incoming = deepcopy(style)
        sid = _style_id(incoming)
        if not sid:
            continue
        if sid in existing:
            if conflict == "skip":
                stats["skipped"] += 1
                continue
            if conflict == "rename":
                base = sid
                index = 2
                while sid in existing:
                    sid = f"{base}_Imported{index}"
                    index += 1
                incoming.set(f"{{{W_NS}}}styleId", sid)
                name_node = incoming.find("w:name", NS)
                if name_node is not None:
                    name_node.set(f"{{{W_NS}}}val", f"{_w_val(name_node)} 导入")
                target_styles_root.append(incoming)
                existing[sid] = incoming
                stats["renamed"] += 1
                continue
            old = existing[sid]
            old.getparent().remove(old)
            target_styles_root.append(incoming)
            existing[sid] = incoming
            stats["overwritten"] += 1
        else:
            target_styles_root.append(incoming)
            existing[sid] = incoming
            stats["added"] += 1
    return target_styles_root, stats


def _html_escape(value):
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def generate_html_report(report_path, title, rows, summary=None):
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    summary = summary or {}
    body_rows = "\n".join(
        "<tr>" + "".join(f"<td>{_html_escape(cell)}</td>" for cell in row) + "</tr>"
        for row in rows
    )
    summary_html = "".join(f"<div><strong>{_html_escape(k)}：</strong>{_html_escape(v)}</div>" for k, v in summary.items())
    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{_html_escape(title)}</title>
<style>
body{{font-family:'Microsoft YaHei UI',Arial,sans-serif;background:#f6f8fb;color:#1f2937;margin:32px}}
.card{{background:white;border:1px solid #e5e7eb;border-radius:10px;padding:22px;box-shadow:0 8px 22px rgba(15,23,42,.06)}}
h1{{font-size:22px;margin:0 0 16px}}
table{{border-collapse:collapse;width:100%;margin-top:18px;background:white}}
th,td{{border-bottom:1px solid #e5e7eb;text-align:left;padding:10px;font-size:13px;vertical-align:top}}
th{{background:#eef6ff;color:#0f4c81}}
.muted{{color:#64748b;font-size:12px;margin-top:12px}}
</style>
</head>
<body><div class="card"><h1>{_html_escape(title)}</h1>{summary_html}
<table><tbody>{body_rows}</tbody></table>
<div class="muted">生成时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div></div></body></html>"""
    report_path.write_text(html, encoding="utf-8")
    return str(report_path)


def import_styles_to_document(
    source_path,
    target_path,
    output_path=None,
    conflict="overwrite",
    clean_unused=True,
    selected_styles=None,
    include_dependencies=True,
    report_path=None,
):
    """Import all styles from source/template to target, preserving numbering."""
    src_path, src_temp = prepare_document(source_path)
    tgt_path, tgt_temp = prepare_document(target_path)
    try:
        output_path = output_path or _next_output_path(target_path, "_样式导入版")
        used_target = _used_style_ids(tgt_path)

        with zipfile.ZipFile(src_path, "r") as src_zip, zipfile.ZipFile(tgt_path, "r") as tgt_zip:
            source_styles = _read_xml(src_zip, "word/styles.xml")
            source_numbering = _read_xml(src_zip, "word/numbering.xml")
            target_styles = _read_xml(tgt_zip, "word/styles.xml")
            target_numbering = _read_xml(tgt_zip, "word/numbering.xml")
            source_theme = _read_xml(src_zip, "word/theme/theme1.xml")

        source_used = _used_style_ids(src_path)
        if clean_unused and source_used:
            _clean_unused_styles_root(source_styles, source_used)
        filtered, auto_added = _filter_styles_root(source_styles, selected_styles, include_dependencies)
        if clean_unused:
            _clean_unused_styles_root(target_styles, used_target)

        target_numbering, numbering_count = _merge_numbering(source_numbering, target_numbering, source_styles)
        target_styles, stats = _import_styles(source_styles, target_styles, conflict=conflict)

        replacements = {"word/styles.xml": _xml_bytes(target_styles)}
        if target_numbering is not None:
            replacements["word/numbering.xml"] = _xml_bytes(target_numbering)
        if source_theme is not None:
            replacements["word/theme/theme1.xml"] = _xml_bytes(source_theme)

        _copy_docx_with_replacements(tgt_path, output_path, replacements)
        report = ""
        if report_path:
            report = generate_html_report(
                report_path,
                "Word 样式导入报告",
                [
                    ("来源", source_path),
                    ("目标", target_path),
                    ("输出", output_path),
                    ("新增样式", stats.get("added", 0)),
                    ("覆盖样式", stats.get("overwritten", 0)),
                    ("跳过样式", stats.get("skipped", 0)),
                    ("重命名样式", stats.get("renamed", 0)),
                    ("同步编号", numbering_count),
                    ("自动补选依赖", ", ".join(auto_added) if auto_added else "无"),
                    ("过滤未选样式", filtered),
                ],
            )
        stats.update({"success": True, "output": output_path, "numbering": numbering_count, "report": report, "auto_added": auto_added, "filtered": filtered})
        return stats
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        for path, is_temp in ((src_path, src_temp), (tgt_path, tgt_temp)):
            if is_temp:
                try:
                    os.remove(path)
                except Exception:
                    pass


def batch_import_styles(source_path, target_paths, output_dir=None, conflict="overwrite", clean_unused=True, selected_styles=None, include_dependencies=True, report_dir=None):
    results = []
    output_dir = Path(output_dir) if output_dir else None
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)
    for target in target_paths:
        out = None
        if output_dir:
            target_path = Path(target)
            out = str(output_dir / f"{target_path.stem}_样式导入版.docx")
        report_path = None
        if report_dir:
            Path(report_dir).mkdir(parents=True, exist_ok=True)
            report_path = str(Path(report_dir) / f"{Path(target).stem}_样式导入报告.html")
        results.append(import_styles_to_document(source_path, target, out, conflict, clean_unused, selected_styles, include_dependencies, report_path))
    return results
