"""模板导出、导入与模板库管理"""

import json
import os
import zipfile
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from lxml import etree

from src.ewt.config import NS, W_NS
from src.ewt.utils.helpers import _w_val, _style_id, _style_name, _style_dict
from src.ewt.utils.word_io import prepare_document, _read_xml, _xml_bytes, _copy_docx_with_replacements, _empty_document_xml
from src.ewt.core.style_engine import _used_style_ids, _clean_unused_styles_root, _filter_styles_root, _clean_unused_styles_root
from src.ewt.core.numbering import _merge_numbering


def make_clean_style_carrier(source_path, output_path, as_template=False, selected_styles=None, include_dependencies=True):
    """创建空正文样式载体文档/模板"""
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
    """导出样式模板为 .dotx/.docx 并可加入模板库"""
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
    """更新模板库 JSON 索引"""
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
    """列出模板库中所有模板"""
    index_path = Path(library_dir) / "template_index.json"
    if not index_path.exists():
        return []
    try:
        return json.loads(index_path.read_text(encoding="utf-8"))
    except Exception:
        return []


def _import_styles(source_styles_root, target_styles_root, conflict="overwrite"):
    """将源样式合并到目标样式树，返回 (merged_root, stats)"""
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
            # overwrite
            old = existing[sid]
            target_styles_root.replace(old, incoming)
            existing[sid] = incoming
            stats["overwritten"] += 1
        else:
            target_styles_root.append(incoming)
            existing[sid] = incoming
            stats["added"] += 1

    return target_styles_root, stats


def import_styles_to_document(
    source_path, target_path, output_path=None, conflict="overwrite", clean_unused=True,
    selected_styles=None, include_dependencies=True, report_path=None,
):
    """将源模板样式导入目标文档"""
    output_path = output_path or str(Path(target_path).with_name(f"{Path(target_path).stem}_样式导入版.docx"))
    source_docx, src_temp = prepare_document(source_path)
    target_docx, tgt_temp = prepare_document(target_path)

    try:
        with zipfile.ZipFile(source_docx, "r") as zf:
            source_styles_root = _read_xml(zf, "word/styles.xml")
            source_numbering_root = _read_xml(zf, "word/numbering.xml")

        with zipfile.ZipFile(target_docx, "r") as zf:
            target_styles_root = _read_xml(zf, "word/styles.xml")
            target_numbering_root = _read_xml(zf, "word/numbering.xml")

        # Filter selected styles if specified
        if selected_styles:
            from src.ewt.core.style_engine import _selected_style_ids, _filter_styles_root
            keep = _selected_style_ids(source_styles_root, selected_styles, include_dependencies)
            _filter_styles_root(source_styles_root, selected_styles, include_dependencies)

        merged_styles, stats = _import_styles(source_styles_root, target_styles_root, conflict)
        merged_numbering, numbering_count = _merge_numbering(
            source_numbering_root, target_numbering_root, source_styles_root
        )
        stats["numbering"] = numbering_count

        if clean_unused:
            from src.ewt.utils.helpers import _iter_word_xml_names
            used = set()
            with zipfile.ZipFile(target_docx, "r") as zf:
                for name in _iter_word_xml_names(zf):
                    root = _read_xml(zf, name)
                    if root is None:
                        continue
                    for node in root.iter():
                        tag = node.tag
                        if tag in {
                            f"{{{W_NS}}}pStyle",
                            f"{{{W_NS}}}rStyle",
                            f"{{{W_NS}}}tblStyle",
                        }:
                            val = _w_val(node)
                            if val:
                                used.add(val)
            _clean_unused_styles_root(merged_styles, used)

        replacements = {"word/styles.xml": _xml_bytes(merged_styles)}
        if merged_numbering is not None:
            replacements["word/numbering.xml"] = _xml_bytes(merged_numbering)
        _copy_docx_with_replacements(target_docx, output_path, replacements)

        return {"success": True, "output": output_path, "added": stats["added"], "overwritten": stats["overwritten"],
                "skipped": stats["skipped"], "renamed": stats["renamed"], "numbering": stats["numbering"]}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        for path, is_temp in [(source_docx, src_temp), (target_docx, tgt_temp)]:
            if is_temp:
                try:
                    os.remove(path)
                except Exception:
                    pass


def batch_import_styles(
    source_path, target_paths, output_dir=None, conflict="overwrite", clean_unused=True,
    selected_styles=None, include_dependencies=True, report_dir=None,
):
    """批量将模板样式导入多个目标文档"""
    results = []
    output_dir = Path(output_dir) if output_dir else None
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)
    for target in target_paths:
        out = None
        if output_dir:
            target_path_obj = Path(target)
            out = str(output_dir / f"{target_path_obj.stem}_样式导入版.docx")
        report_path = None
        if report_dir:
            Path(report_dir).mkdir(parents=True, exist_ok=True)
            report_path = str(Path(report_dir) / f"{Path(target).stem}_样式导入报告.html")
        results.append(import_styles_to_document(
            source_path, target, out, conflict, clean_unused,
            selected_styles, include_dependencies, report_path,
        ))
    return results
