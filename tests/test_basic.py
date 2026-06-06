"""基础测试 — 验证模块导入与常量正确性"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import unittest


class TestConfig(unittest.TestCase):
    def test_constants_defined(self):
        from src.ewt.config import NS, W_NS, APP_NAME, DEFAULT_NUMBERING_PRESETS
        self.assertIsNotNone(NS)
        self.assertIsNotNone(W_NS)
        self.assertEqual(APP_NAME, "Word 样式管理器")
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


if __name__ == "__main__":
    unittest.main()
