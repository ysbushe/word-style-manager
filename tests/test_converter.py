"""Tests for the UI-independent conversion core."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.ewt.core import converter


class ConverterCoreTest(unittest.TestCase):
    def test_application_identity_detects_wps(self):
        application = type(
            "Application",
            (),
            {
                "Name": "WPS Writer",
                "Version": "12.1",
                "Path": r"C:\Users\Test\AppData\Local\Kingsoft\WPS Office\office6",
            },
        )()

        identity = converter._application_identity(application)

        self.assertEqual(identity["product"], "wps")

    def test_application_identity_detects_microsoft_office(self):
        application = type(
            "Application",
            (),
            {
                "Name": "Microsoft Excel",
                "Version": "16.0",
                "Path": r"C:\Program Files\Microsoft Office\root\Office16",
            },
        )()

        identity = converter._application_identity(application)

        self.assertEqual(identity["product"], "office")

    def test_unique_output_path_uses_conversion_suffix_on_conflict(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "报告.doc"
            source.write_text("legacy", encoding="utf-8")
            (root / "报告.docx").write_text("existing", encoding="utf-8")
            (root / "报告_转换版.docx").write_text("existing", encoding="utf-8")

            output = converter._unique_output_path(source, ".docx")

            self.assertEqual(output, root / "报告_转换版_2.docx")

    def test_convert_files_returns_structured_failure_for_unsupported_file(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "notes.txt"
            source.write_text("not supported", encoding="utf-8")

            result = converter.convert_files([source])

            self.assertFalse(result["success"])
            self.assertEqual(result["succeeded"], 0)
            self.assertEqual(result["failed_count"], 1)
            self.assertIn("暂只支持", result["failed"][0]["error"])


if __name__ == "__main__":
    unittest.main()
