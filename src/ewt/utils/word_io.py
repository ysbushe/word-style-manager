"""Word 文件读写与格式转换"""

import os
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

from src.ewt.config import NS, W_NS, WORD_TEMPLATE_TYPE

try:
    import win32com.client
    _HAS_WIN32COM = True
except ImportError:
    _HAS_WIN32COM = False


def convert_doc_to_docx(doc_path):
    """将 .doc 转换为临时 .docx，需本机安装 Microsoft Word"""
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
