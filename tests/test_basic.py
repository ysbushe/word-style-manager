"""基础测试 — 验证模块导入与常量正确性"""

import sys
import os
import tomllib
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import unittest


class TestConfig(unittest.TestCase):
    def test_constants_defined(self):
        from src.ewt.config import NS, W_NS, APP_NAME, DEFAULT_NUMBERING_PRESETS, USER_GUIDE_FILE
        self.assertIsNotNone(NS)
        self.assertIsNotNone(W_NS)
        self.assertEqual(APP_NAME, "Word 样式管理器")
        self.assertEqual(USER_GUIDE_FILE, "使用说明.md")
        self.assertTrue(len(DEFAULT_NUMBERING_PRESETS) >= 6)

    def test_numbering_presets_structure(self):
        from src.ewt.config import DEFAULT_NUMBERING_PRESETS
        for preset in DEFAULT_NUMBERING_PRESETS:
            self.assertIn("name", preset)
            self.assertIn("levels", preset)
            self.assertIsInstance(preset["levels"], list)
            for lvl in preset["levels"]:
                self.assertIn("format", lvl)
                self.assertIn("numFmt", lvl)

    def test_release_version_is_consistent(self):
        from src.ewt.config import APP_VERSION

        root = Path(__file__).resolve().parents[1]
        with (root / "pyproject.toml").open("rb") as handle:
            project_version = tomllib.load(handle)["project"]["version"]
        self.assertEqual(APP_VERSION, "0.4.0")
        self.assertEqual(project_version, APP_VERSION)


class TestImports(unittest.TestCase):
    def test_can_import_core_modules(self):
        from src.ewt.core import style_engine
        from src.ewt.core import numbering
        from src.ewt.core import templates
        self.assertTrue(hasattr(style_engine, "analyze_document"))
        self.assertTrue(hasattr(numbering, "_numbering_maps"))
        self.assertTrue(hasattr(templates, "export_style_template"))

    def test_can_import_utils(self):
        from src.ewt.utils import helpers
        from src.ewt.utils import word_io
        self.assertTrue(hasattr(helpers, "_w_val"))
        self.assertTrue(hasattr(word_io, "prepare_document"))
        self.assertTrue(hasattr(word_io, "create_office_application"))

    def test_office_product_detection_recognizes_wps_path(self):
        from src.ewt.utils.word_io import _office_product_name

        application = type(
            "Application",
            (),
            {"Path": r"C:\Users\Test\AppData\Local\Kingsoft\WPS Office\office6", "Build": "12.1"},
        )()
        self.assertEqual(_office_product_name(application, "Microsoft Word"), "WPS Writer")


if __name__ == "__main__":
    unittest.main()
