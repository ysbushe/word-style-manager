"""Feature tests for style selection, reports, task schemes, and template editing."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.ewt.core.editor import save_template_edits
from src.ewt.core.numbering import _deduplicate_numbering_root
from src.ewt.core.updater import _select_zip_asset, is_newer_version
from src.ewt.config import W_NS
from src.ewt.core.style_engine import inspect_document
from src.ewt.core.templates import (
    export_style_template,
    import_styles_to_document,
    list_task_schemes,
    save_task_scheme,
)


class FeatureTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / "source.docx"
        self.target = self.root / "target.docx"
        doc = Document()
        used = doc.styles.add_style("工程标题", WD_STYLE_TYPE.PARAGRAPH)
        used.font.name = "Microsoft YaHei"
        doc.styles.add_style("显式保留但未使用", WD_STYLE_TYPE.PARAGRAPH)
        doc.add_paragraph("标题", style="工程标题")
        doc.save(self.source)
        target = Document()
        target.add_paragraph("正文")
        target.save(self.target)

    def tearDown(self):
        self.temp.cleanup()

    def test_selected_unused_style_can_be_exported(self):
        info = inspect_document(str(self.source))
        selected = next(item["style_id"] for item in info["styles"] if item["name"] == "显式保留但未使用")
        result = export_style_template(
            str(self.source),
            str(self.root),
            "选择模板",
            ["dotx"],
            str(self.root / "library"),
            [selected],
            True,
        )
        self.assertTrue(result["success"])
        exported = inspect_document(result["outputs"]["dotx"])
        self.assertIn("显式保留但未使用", {item["name"] for item in exported["styles"]})

    def test_import_creates_html_report(self):
        info = inspect_document(str(self.source))
        selected = next(item["style_id"] for item in info["styles"] if item["name"] == "工程标题")
        exported = export_style_template(
            str(self.source), str(self.root), "报告模板", ["dotx"], None, [selected], True
        )
        report = self.root / "report.html"
        result = import_styles_to_document(
            exported["outputs"]["dotx"],
            str(self.target),
            str(self.root / "output.docx"),
            selected_styles=[selected],
            report_path=str(report),
        )
        self.assertTrue(result["success"])
        self.assertTrue(report.exists())
        self.assertIn("Word 样式导入报告", report.read_text(encoding="utf-8"))

    def test_template_edit_creates_version(self):
        info = inspect_document(str(self.source))
        selected = next(item["style_id"] for item in info["styles"] if item["name"] == "工程标题")
        exported = export_style_template(str(self.source), str(self.root), "编辑模板", ["dotx"])
        result = save_template_edits(
            exported["outputs"]["dotx"],
            {selected: {"name": "工程标题新版", "font": "SimSun", "size": "16", "bold": True}},
            [{"format": "第%1章", "numFmt": "decimal", "suffix": "space"}],
            True,
            True,
            False,
        )
        self.assertTrue(Path(result["output"]).exists())
        self.assertNotEqual(Path(result["output"]), Path(exported["outputs"]["dotx"]))

    def test_task_scheme_roundtrip(self):
        save_task_scheme(self.root, "每日任务", {"source": "a.dotx", "targets": ["b.docx"]})
        schemes = list_task_schemes(self.root)
        self.assertEqual(schemes[0]["name"], "每日任务")
        self.assertEqual(schemes[0]["source"], "a.dotx")

    def test_duplicate_abstract_numbering_is_merged(self):
        root = etree.Element(f"{{{W_NS}}}numbering", nsmap={"w": W_NS})
        for abs_id in ("1", "2"):
            abstract = etree.SubElement(root, f"{{{W_NS}}}abstractNum")
            abstract.set(f"{{{W_NS}}}abstractNumId", abs_id)
            level = etree.SubElement(abstract, f"{{{W_NS}}}lvl")
            level.set(f"{{{W_NS}}}ilvl", "0")
            fmt = etree.SubElement(level, f"{{{W_NS}}}numFmt")
            fmt.set(f"{{{W_NS}}}val", "decimal")
        num = etree.SubElement(root, f"{{{W_NS}}}num")
        num.set(f"{{{W_NS}}}numId", "9")
        ref = etree.SubElement(num, f"{{{W_NS}}}abstractNumId")
        ref.set(f"{{{W_NS}}}val", "2")
        removed = _deduplicate_numbering_root(root)
        self.assertEqual(removed, 1)
        self.assertEqual(ref.get(f"{{{W_NS}}}val"), "1")

    def test_update_version_comparison(self):
        self.assertTrue(is_newer_version("v0.2.1", "0.2.0"))
        self.assertFalse(is_newer_version("v0.2.0", "0.2.0"))
        self.assertFalse(is_newer_version("v0.1.9", "0.2.0"))

    def test_portable_release_asset_is_preferred(self):
        assets = [
            {"name": "source.zip", "browser_download_url": "source"},
            {"name": "Word样式管理器_v0.3.0_绿色版.zip", "browser_download_url": "portable"},
        ]
        self.assertEqual(_select_zip_asset(assets)["browser_download_url"], "portable")


if __name__ == "__main__":
    unittest.main()
