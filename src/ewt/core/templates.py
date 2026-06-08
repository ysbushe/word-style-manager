"""Template export, import, library management, and batch task helpers."""

from __future__ import annotations

import json
import os
import zipfile
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from src.ewt.config import NS, W_NS
from src.ewt.core.numbering import _clean_unused_numbering_root, _merge_numbering, _used_numbering_ids
from src.ewt.core.reports import generate_html_report
from src.ewt.core.style_engine import (
    _clean_unused_styles_root,
    _filter_styles_root,
    _keep_style_ids,
    _selected_style_ids,
    _used_style_ids,
)
from src.ewt.utils.helpers import _style_dict, _style_id, _w_val
from src.ewt.utils.word_io import (
    _copy_docx_with_replacements,
    _empty_document_xml,
    _read_xml,
    _xml_bytes,
    prepare_document,
)


def _unique_path(path):
    path = Path(path)
    if not path.exists():
        return path
    index = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def _update_template_index(library_dir, name, source_path, template_path):
    library_dir = Path(library_dir)
    library_dir.mkdir(parents=True, exist_ok=True)
    index_path = library_dir / "template_index.json"
    try:
        items = json.loads(index_path.read_text(encoding="utf-8")) if index_path.exists() else []
    except Exception:
        items = []
    record = {
        "name": name,
        "source": str(source_path),
        "path": str(template_path),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "version_key": Path(template_path).stem,
    }
    items = [item for item in items if item.get("path") != str(template_path)]
    items.append(record)
    index_path.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    return record


def list_template_library(library_dir):
    library_dir = Path(library_dir)
    index_path = library_dir / "template_index.json"
    items = []
    if index_path.exists():
        try:
            items = json.loads(index_path.read_text(encoding="utf-8"))
        except Exception:
            items = []
    known = {str(Path(item.get("path", ""))) for item in items}
    for file in library_dir.glob("*.dotx"):
        if str(file) not in known:
            items.append(
                {
                    "name": file.stem,
                    "source": "",
                    "path": str(file),
                    "created_at": datetime.fromtimestamp(file.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    "version_key": file.stem,
                }
            )
    items.sort(key=lambda item: item.get("created_at", ""), reverse=True)
    return items


def make_clean_style_carrier(source_path, output_path, as_template=False, selected_styles=None, include_dependencies=True):
    docx_path, is_temp = prepare_document(source_path)
    try:
        used = _used_style_ids(docx_path)
        with zipfile.ZipFile(docx_path, "r") as zf:
            styles_root = _read_xml(zf, "word/styles.xml")
            numbering_root = _read_xml(zf, "word/numbering.xml")
            document_root = _read_xml(zf, "word/document.xml")
        selected_for_clean = set(selected_styles or [])
        removed_styles = _clean_unused_styles_root(styles_root, used | selected_for_clean)
        filtered, auto_added = _filter_styles_root(styles_root, selected_styles, include_dependencies)
        used_nums = _used_numbering_ids(docx_path, styles_root)
        removed_nums, removed_abs = _clean_unused_numbering_root(numbering_root, used_nums)
        replacements = {
            "word/styles.xml": _xml_bytes(styles_root),
            "word/document.xml": _xml_bytes(_empty_document_xml(document_root)),
        }
        if numbering_root is not None:
            replacements["word/numbering.xml"] = _xml_bytes(numbering_root)
        _copy_docx_with_replacements(docx_path, output_path, replacements, as_template=as_template)
        return {
            "success": True,
            "output": output_path,
            "removed_unused": removed_styles,
            "removed_numbering": removed_nums + removed_abs,
            "filtered": filtered,
            "auto_added": auto_added,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        if is_temp:
            try:
                os.remove(docx_path)
            except Exception:
                pass


def export_style_template(
    source_path,
    output_dir,
    template_name=None,
    formats=None,
    library_dir=None,
    selected_styles=None,
    include_dependencies=True,
):
    formats = formats or ["dotx", "docx", "library"]
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    source = Path(source_path)
    safe_name = template_name or f"{source.stem}_样式模板"
    outputs = {}
    removed_unused = 0
    removed_numbering = 0
    filtered = 0
    auto_added = []

    def run_one(path, as_template):
        return make_clean_style_carrier(
            source_path,
            str(path),
            as_template=as_template,
            selected_styles=selected_styles,
            include_dependencies=include_dependencies,
        )

    if "dotx" in formats:
        result = run_one(_unique_path(output_dir / f"{safe_name}.dotx"), True)
        if not result["success"]:
            return result
        outputs["dotx"] = result["output"]
        removed_unused = result.get("removed_unused", removed_unused)
        removed_numbering = result.get("removed_numbering", removed_numbering)
        filtered = result.get("filtered", filtered)
        auto_added = result.get("auto_added", auto_added)

    if "docx" in formats:
        result = run_one(_unique_path(output_dir / f"{safe_name}.docx"), False)
        if not result["success"]:
            return result
        outputs["docx"] = result["output"]
        removed_unused = result.get("removed_unused", removed_unused)
        removed_numbering = result.get("removed_numbering", removed_numbering)
        filtered = result.get("filtered", filtered)
        auto_added = result.get("auto_added", auto_added)

    if "library" in formats:
        lib_dir = Path(library_dir or output_dir / "templates")
        lib_dir.mkdir(parents=True, exist_ok=True)
        result = run_one(_unique_path(lib_dir / f"{safe_name}.dotx"), True)
        if not result["success"]:
            return result
        outputs["library"] = result["output"]
        _update_template_index(lib_dir, safe_name, source_path, result["output"])
        removed_unused = result.get("removed_unused", removed_unused)
        removed_numbering = result.get("removed_numbering", removed_numbering)
        filtered = result.get("filtered", filtered)
        auto_added = result.get("auto_added", auto_added)

    return {
        "success": True,
        "outputs": outputs,
        "removed_unused": removed_unused,
        "removed_numbering": removed_numbering,
        "filtered": filtered,
        "auto_added": auto_added,
    }


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
            target_styles_root.replace(existing[sid], incoming)
            existing[sid] = incoming
            stats["overwritten"] += 1
        else:
            target_styles_root.append(incoming)
            existing[sid] = incoming
            stats["added"] += 1
    return target_styles_root, stats


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
    output_path = output_path or str(_unique_path(Path(target_path).with_name(f"{Path(target_path).stem}_样式导入版.docx")))
    source_docx, src_temp = prepare_document(source_path)
    target_docx, tgt_temp = prepare_document(target_path)
    try:
        with zipfile.ZipFile(source_docx, "r") as zf:
            source_styles_root = _read_xml(zf, "word/styles.xml")
            source_numbering_root = _read_xml(zf, "word/numbering.xml")
            source_theme = _read_xml(zf, "word/theme/theme1.xml")
        with zipfile.ZipFile(target_docx, "r") as zf:
            target_styles_root = _read_xml(zf, "word/styles.xml")
            target_numbering_root = _read_xml(zf, "word/numbering.xml")

        filtered, auto_added = _filter_styles_root(source_styles_root, selected_styles, include_dependencies)
        if selected_styles and source_numbering_root is not None:
            source_nums = {
                _w_val(node)
                for node in source_styles_root.findall(".//w:numId", NS)
                if _w_val(node)
            }
            _clean_unused_numbering_root(source_numbering_root, source_nums)
        if clean_unused:
            target_used = _used_style_ids(target_docx)
            _clean_unused_styles_root(target_styles_root, target_used)
            used_nums = _used_numbering_ids(target_docx, target_styles_root)
            removed_nums, removed_abs = _clean_unused_numbering_root(target_numbering_root, used_nums)
        else:
            removed_nums = removed_abs = 0

        merged_numbering, numbering_count = _merge_numbering(source_numbering_root, target_numbering_root, source_styles_root)
        merged_styles, stats = _import_styles(source_styles_root, target_styles_root, conflict)

        replacements = {"word/styles.xml": _xml_bytes(merged_styles)}
        if merged_numbering is not None:
            replacements["word/numbering.xml"] = _xml_bytes(merged_numbering)
        if source_theme is not None:
            replacements["word/theme/theme1.xml"] = _xml_bytes(source_theme)
        _copy_docx_with_replacements(target_docx, output_path, replacements)

        report = ""
        if report_path:
            report = generate_html_report(
                report_path,
                "Word 样式导入报告",
                [
                    ("样式来源", source_path),
                    ("目标文件", target_path),
                    ("输出文件", output_path),
                    ("同名策略", conflict),
                    ("自动补选依赖", ", ".join(auto_added) if auto_added else "无"),
                    ("过滤未选样式", filtered),
                    ("清理未用编号", removed_nums + removed_abs),
                ],
                {
                    "新增": stats["added"],
                    "覆盖": stats["overwritten"],
                    "跳过": stats["skipped"],
                    "重命名": stats["renamed"],
                    "同步编号": numbering_count,
                },
            )
        return {
            "success": True,
            "output": output_path,
            "report": report,
            "added": stats["added"],
            "overwritten": stats["overwritten"],
            "skipped": stats["skipped"],
            "renamed": stats["renamed"],
            "numbering": numbering_count,
            "filtered": filtered,
            "auto_added": auto_added,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        for path, is_temp in ((source_docx, src_temp), (target_docx, tgt_temp)):
            if is_temp:
                try:
                    os.remove(path)
                except Exception:
                    pass


def batch_import_styles(
    source_path,
    target_paths,
    output_dir=None,
    conflict="overwrite",
    clean_unused=True,
    selected_styles=None,
    include_dependencies=True,
    report_dir=None,
):
    results = []
    output_dir = Path(output_dir) if output_dir else None
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)
    for target in target_paths:
        out = None
        if output_dir:
            out = str(_unique_path(output_dir / f"{Path(target).stem}_样式导入版.docx"))
        report_path = None
        if report_dir:
            report_path = Path(report_dir) / f"{Path(target).stem}_样式导入报告.html"
        results.append(
            import_styles_to_document(
                source_path,
                target,
                out,
                conflict,
                clean_unused,
                selected_styles,
                include_dependencies,
                str(report_path) if report_path else None,
            )
        )
    return results


def save_task_scheme(library_dir, name, data):
    path = Path(library_dir) / "task_schemes.json"
    try:
        items = json.loads(path.read_text(encoding="utf-8")) if path.exists() else []
    except Exception:
        items = []
    data = dict(data)
    data["name"] = name
    data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    items = [item for item in items if item.get("name") != name]
    items.append(data)
    path.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(path)


def list_task_schemes(library_dir):
    path = Path(library_dir) / "task_schemes.json"
    if not path.exists():
        return []
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []
