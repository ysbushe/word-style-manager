"""Word 文件读写与格式转换"""

import os
import posixpath
import tempfile
import zipfile
import logging
from copy import deepcopy
from pathlib import Path
from uuid import uuid4

from lxml import etree

from src.ewt.config import NS, W_NS, WORD_TEMPLATE_TYPE

try:
    import pythoncom
    import win32com.client
    _HAS_WIN32COM = True
except ImportError:
    pythoncom = None
    _HAS_WIN32COM = False


OFFICE_PROGIDS = (
    ("WPS Writer", "KWPS.Application"),
    ("Microsoft Word", "Word.Application"),
)
logger = logging.getLogger(__name__)


def _office_product_name(application, fallback):
    """Identify the real office suite because WPS may expose itself as Word."""
    install_path = str(getattr(application, "Path", "") or "")
    build = str(getattr(application, "Build", "") or "")
    lowered = install_path.lower()
    if "kingsoft" in lowered or "wps office" in lowered or "\\wps" in lowered:
        return "WPS Writer"
    if "microsoft office" in lowered or "\\office" in lowered and "winword" in lowered:
        return "Microsoft Word"
    return fallback + (f" {build}" if build else "")


def create_office_application():
    """Start an isolated WPS Writer or Microsoft Word automation instance."""
    if not _HAS_WIN32COM:
        raise RuntimeError("处理 .doc 文件需要 pywin32，并需要安装 WPS Writer 或 Microsoft Word。")

    errors = []
    for label, progid in OFFICE_PROGIDS:
        try:
            application = win32com.client.DispatchEx(progid)
            return application, _office_product_name(application, label)
        except Exception as exc:
            errors.append(f"{label}: {exc}")
    raise RuntimeError(
        "未检测到可用于转换 .doc 的 WPS Writer 或 Microsoft Word。"
        "请确认至少安装其中一个，并允许其自动化组件运行。\n"
        + "\n".join(errors)
    )


def convert_doc_to_docx(doc_path):
    """将 .doc 转换为临时 .docx，自动使用 WPS Writer 或 Microsoft Word。"""
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"找不到文件: {doc_path}")
    if not _HAS_WIN32COM:
        raise RuntimeError("处理 .doc 文件需要 pywin32，并需要安装 WPS Writer 或 Microsoft Word。")

    pythoncom.CoInitialize()
    errors = []
    try:
        for label, progid in OFFICE_PROGIDS:
            word = None
            doc = None
            try:
                word = win32com.client.DispatchEx(progid)
                product = _office_product_name(word, label)
                word.Visible = False
                word.DisplayAlerts = False
                doc = word.Documents.Open(os.path.abspath(doc_path))
                name_no_ext = os.path.splitext(os.path.basename(doc_path))[0]
                docx_path = os.path.join(
                    tempfile.gettempdir(),
                    f"{name_no_ext}_temp_{os.getpid()}_{uuid4().hex}.docx",
                )
                doc.SaveAs(docx_path, FileFormat=16)
                return docx_path
            except Exception as exc:
                errors.append(f"{label}: {exc}")
            finally:
                try:
                    if doc:
                        doc.Close(False)
                except Exception as exc:
                    logger.warning("关闭 Office 文档失败：%s", exc)
                try:
                    if word:
                        word.Quit()
                except Exception as exc:
                    logger.warning("退出 Office 自动化进程失败：%s", exc)
        raise RuntimeError(
            "WPS Writer 和 Microsoft Word 都无法完成转换。"
            "请关闭正在占用该文件的 Office 窗口，并检查文件权限。\n"
            + "\n".join(errors)
        )
    finally:
        pythoncom.CoUninitialize()


def prepare_document(file_path):
    """返回 docx 路径及是否为临时文件"""
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


def _copy_docx_with_replacements(input_path, output_path, replacements, as_template=False):
    """复制 docx 并替换指定 XML 部件"""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_name(f".{output.name}.{uuid4().hex}.tmp")
    try:
        with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
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
        os.replace(temp, output)
    except Exception:
        try:
            temp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def _copy_lightweight_template(input_path, output_path, replacements):
    """Create a compact style template without source document content assets."""
    keep_exact = {
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml",
        "word/_rels/document.xml.rels",
        "word/styles.xml",
        "word/stylesWithEffects.xml",
        "word/numbering.xml",
        "word/settings.xml",
        "word/webSettings.xml",
        "word/fontTable.xml",
    }
    keep_prefixes = ("docProps/", "word/theme/")
    with zipfile.ZipFile(input_path, "r") as zin:
        source_names = set(zin.namelist())
        if "word/_rels/numbering.xml.rels" in source_names:
            keep_exact.add("word/_rels/numbering.xml.rels")
            numbering_rels = etree.fromstring(zin.read("word/_rels/numbering.xml.rels"))
            rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            for rel in numbering_rels.findall(f"{{{rel_ns}}}Relationship"):
                if rel.get("TargetMode") != "External":
                    keep_exact.add(posixpath.normpath(posixpath.join("word", rel.get("Target", ""))))
        kept_names = {
            name
            for name in source_names
            if name in keep_exact or name.startswith(keep_prefixes)
        }
        kept_names.update(replacements)
        payloads = {}
        for name in kept_names:
            if name in replacements:
                payloads[name] = replacements[name]
            elif name in source_names:
                payloads[name] = zin.read(name)

    content_types = payloads.get("[Content_Types].xml")
    if content_types:
        root = etree.fromstring(content_types)
        content_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        for override in list(root.findall(f"{{{content_ns}}}Override")):
            part_name = override.get("PartName", "").lstrip("/")
            if part_name not in kept_names:
                root.remove(override)
            elif part_name == "word/document.xml":
                override.set("ContentType", WORD_TEMPLATE_TYPE)
        payloads["[Content_Types].xml"] = _xml_bytes(root)

    relationships = payloads.get("word/_rels/document.xml.rels")
    if relationships:
        root = etree.fromstring(relationships)
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        for rel in list(root.findall(f"{{{rel_ns}}}Relationship")):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            normalized = posixpath.normpath(posixpath.join("word", target))
            if normalized not in kept_names:
                root.remove(rel)
        payloads["word/_rels/document.xml.rels"] = _xml_bytes(root)

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_name(f".{output.name}.{uuid4().hex}.tmp")
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in sorted(payloads):
                zout.writestr(name, payloads[name])
        os.replace(temp, output)
    except Exception:
        try:
            temp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def _empty_document_xml(source_root):
    """基于源文档生成空正文段落文档"""
    body = source_root.find("w:body", NS)
    sect_pr = body.find("w:sectPr", NS) if body is not None else None
    doc = etree.Element(f"{{{W_NS}}}document", nsmap=source_root.nsmap)
    new_body = etree.SubElement(doc, f"{{{W_NS}}}body")
    if sect_pr is not None:
        new_body.append(deepcopy(sect_pr))
    return doc


def _next_output_path(path, suffix, ext=".docx"):
    """避免覆盖已有文件，自动追加编号"""
    base = Path(path)
    candidate = base.with_name(f"{base.stem}{suffix}{ext}")
    index = 2
    while candidate.exists():
        candidate = base.with_name(f"{base.stem}{suffix}_{index}{ext}")
        index += 1
    return str(candidate)
