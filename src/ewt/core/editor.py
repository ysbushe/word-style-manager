"""Template style and multilevel-numbering editor."""

from __future__ import annotations

import json
import os
import re
import zipfile
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from lxml import etree

from src.ewt.config import DEFAULT_NUMBERING_PRESETS, NS, W_NS
from src.ewt.core.style_engine import inspect_document
from src.ewt.utils.helpers import _style_dict, _style_id
from src.ewt.utils.safe_io import atomic_write_json, preserve_corrupt_file
from src.ewt.utils.word_io import _copy_docx_with_replacements, _read_xml, _xml_bytes


def _w(tag):
    return f"{{{W_NS}}}{tag}"


def _child(parent, tag):
    node = parent.find(f"w:{tag}", NS)
    if node is None:
        node = etree.SubElement(parent, _w(tag))
    return node


def _path(parent, *tags):
    node = parent
    for tag in tags:
        node = _child(node, tag)
    return node


def _set_val(node, value):
    if value is None or value == "":
        return
    node.set(_w("val"), str(value))


def _set_toggle(parent, tag, enabled):
    old = parent.find(f"w:{tag}", NS)
    if enabled:
        if old is None:
            etree.SubElement(parent, _w(tag))
    elif old is not None:
        parent.remove(old)


def _apply_style_update(style, values):
    if values.get("name"):
        _set_val(_child(style, "name"), values["name"])
    rpr = _child(style, "rPr")
    ppr = _child(style, "pPr")

    font = values.get("font", "").strip()
    if font:
        fonts = _child(rpr, "rFonts")
        for attr in ("ascii", "hAnsi", "eastAsia"):
            fonts.set(_w(attr), font)
    size = values.get("size", "").strip()
    if size:
        half_points = str(round(float(size) * 2))
        _set_val(_child(rpr, "sz"), half_points)
        _set_val(_child(rpr, "szCs"), half_points)
    color = values.get("color", "").strip().lstrip("#")
    if color:
        _set_val(_child(rpr, "color"), color.upper())
    _set_toggle(rpr, "b", bool(values.get("bold")))
    _set_toggle(rpr, "i", bool(values.get("italic")))

    spacing = _child(ppr, "spacing")
    for key in ("before", "after", "line"):
        value = values.get(key, "")
        if value != "":
            spacing.set(_w(key), str(value))
    line_rule = values.get("line_rule", "").strip()
    if line_rule:
        spacing.set(_w("lineRule"), line_rule)
    ind = _child(ppr, "ind")
    for key in ("left", "hanging"):
        value = values.get(key, "")
        if value != "":
            ind.set(_w(key), str(value))
    for key, tag in (("based_on", "basedOn"), ("next", "next"), ("link", "link")):
        value = values.get(key, "").strip()
        if value:
            _set_val(_child(style, tag), value)


def _max_numeric_id(root, tag, attr):
    values = []
    if root is None:
        return 0
    for node in root.findall(f"w:{tag}", NS):
        value = node.get(_w(attr), "")
        if value.isdigit():
            values.append(int(value))
    return max(values, default=0)


def _ensure_numbering(numbering_root):
    if numbering_root is None:
        numbering_root = etree.Element(_w("numbering"), nsmap={"w": W_NS})
    abstract = numbering_root.find("w:abstractNum", NS)
    if abstract is None:
        abstract = etree.SubElement(numbering_root, _w("abstractNum"))
        abstract.set(_w("abstractNumId"), str(_max_numeric_id(numbering_root, "abstractNum", "abstractNumId") + 1))
        _set_val(etree.SubElement(abstract, _w("multiLevelType")), "multilevel")
    abs_id = abstract.get(_w("abstractNumId"))
    num = None
    for candidate in numbering_root.findall("w:num", NS):
        ref = candidate.find("w:abstractNumId", NS)
        if ref is not None and ref.get(_w("val")) == abs_id:
            num = candidate
            break
    if num is None:
        num = etree.SubElement(numbering_root, _w("num"))
        num.set(_w("numId"), str(_max_numeric_id(numbering_root, "num", "numId") + 1))
        _set_val(etree.SubElement(num, _w("abstractNumId")), abs_id)
    return numbering_root, abstract, num


def _apply_numbering(numbering_root, levels, auto_link=True, auto_restart=True):
    numbering_root, abstract, num = _ensure_numbering(numbering_root)
    for index in range(9):
        data = levels[index] if index < len(levels) else {}
        lvl = None
        for candidate in abstract.findall("w:lvl", NS):
            if candidate.get(_w("ilvl")) == str(index):
                lvl = candidate
                break
        if lvl is None:
            lvl = etree.SubElement(abstract, _w("lvl"))
            lvl.set(_w("ilvl"), str(index))
        _set_val(_child(lvl, "start"), data.get("start", "1"))
        _set_val(_child(lvl, "numFmt"), data.get("numFmt", "decimal"))
        default_format = ".".join(f"%{i}" for i in range(1, index + 2))
        _set_val(_child(lvl, "lvlText"), data.get("format", default_format))
        _set_val(_child(lvl, "suff"), data.get("suffix", "space"))
        linked = data.get("linked_style", "").strip()
        if auto_link and not linked:
            linked = f"Heading{index + 1}"
        if linked:
            _set_val(_child(lvl, "pStyle"), linked)
        if auto_restart and index > 0:
            _set_val(_child(lvl, "lvlRestart"), str(index))
        elif data.get("restart"):
            _set_val(_child(lvl, "lvlRestart"), data["restart"])
        ppr = _child(lvl, "pPr")
        ind = _child(ppr, "ind")
        ind.set(_w("left"), str(data.get("left", (index + 1) * 720)))
        ind.set(_w("hanging"), str(data.get("hanging", 360)))
    return numbering_root, num.get(_w("numId"))


def versioned_path(template_path):
    source = Path(template_path)
    stamp = datetime.now().strftime("v%Y%m%d_%H%M%S")
    return source.with_name(f"{source.stem}_{stamp}{source.suffix}")


def save_template_edits(
    template_path,
    style_updates=None,
    numbering_levels=None,
    auto_link=True,
    auto_restart=True,
    overwrite=False,
):
    source = Path(template_path)
    output = source if overwrite else versioned_path(source)
    with zipfile.ZipFile(source, "r") as zf:
        styles_root = _read_xml(zf, "word/styles.xml")
        numbering_root = _read_xml(zf, "word/numbering.xml")
    styles = _style_dict(styles_root)
    for style_id, values in (style_updates or {}).items():
        style = styles.get(style_id)
        if style is not None:
            _apply_style_update(style, values)
    num_id = ""
    if numbering_levels:
        numbering_root, num_id = _apply_numbering(numbering_root, numbering_levels, auto_link, auto_restart)
    replacements = {"word/styles.xml": _xml_bytes(styles_root)}
    if numbering_root is not None:
        replacements["word/numbering.xml"] = _xml_bytes(numbering_root)
    if overwrite:
        _copy_docx_with_replacements(source, source, replacements, as_template=source.suffix.lower() == ".dotx")
    else:
        _copy_docx_with_replacements(source, output, replacements, as_template=source.suffix.lower() == ".dotx")
    return {"success": True, "output": str(output), "num_id": num_id}


def load_numbering_presets(library_dir):
    path = Path(library_dir) / "numbering_presets.json"
    custom = []
    if path.exists():
        try:
            custom = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            preserve_corrupt_file(path)
            custom = []
    names = set()
    merged = []
    for item in list(DEFAULT_NUMBERING_PRESETS) + custom:
        if item.get("name") not in names:
            merged.append(item)
            names.add(item.get("name"))
    return merged


def save_custom_preset(library_dir, preset):
    path = Path(library_dir) / "numbering_presets.json"
    items = []
    if path.exists():
        try:
            items = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            preserve_corrupt_file(path)
            items = []
    items = [item for item in items if item.get("name") != preset.get("name")]
    items.append(preset)
    path.parent.mkdir(parents=True, exist_ok=True)
    atomic_write_json(path, items)
    return str(path)


def clean_old_template_versions(library_dir, base_name, keep=3, dry_run=False):
    library = Path(library_dir)
    pattern = re.compile(rf"^{re.escape(base_name)}_v\d{{8}}_\d{{6}}\.(dotx|docx)$", re.I)
    matches = sorted(
        [path for path in library.iterdir() if path.is_file() and pattern.match(path.name)],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    selected = matches[max(keep, 1):]
    if dry_run:
        return [str(path) for path in selected]
    removed = []
    for path in selected:
        path.unlink()
        removed.append(str(path))
    return removed


def template_editor_data(template_path):
    return inspect_document(template_path)
