import tempfile
import unittest
from pathlib import Path

from office_tool.config import OfficeToolConfig
from office_tool.services import audit_many, collect_document_inputs, format_many, summarize_results


class ServiceTests(unittest.TestCase):
    def test_collect_document_inputs_reads_supported_files_from_directory(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_services_") as tmp:
            root = Path(tmp)
            (root / "a.txt").write_text("标题\n正文", encoding="utf-8")
            (root / "b.md").write_text("# 标题", encoding="utf-8")
            (root / "skip.xlsx").write_text("", encoding="utf-8")

            paths = collect_document_inputs([root])

        self.assertEqual([path.name for path in paths], ["a.txt", "b.md"])

    def test_batch_audit_and_format_write_reports_and_outputs(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_batch_") as tmp:
            root = Path(tmp)
            source = root / "sample.txt"
            source.write_text("示例标题\n主送机关：\n这是正文。", encoding="utf-8")
            report_dir = root / "reports"
            output_dir = root / "out"

            audits = audit_many([source], OfficeToolConfig(), report_dir=report_dir, markdown=True)
            formats = format_many([source], output_dir, OfficeToolConfig(), report_dir=report_dir, markdown=True)

            self.assertEqual(summarize_results(formats), "完成 1 个任务：成功 1，失败 0。")
            self.assertTrue(audits[0].ok)
            self.assertTrue((report_dir / "sample_audit.json").exists())
            self.assertTrue((report_dir / "sample_audit.md").exists())
            self.assertTrue(formats[0].ok)
            self.assertTrue((output_dir / "sample_formatted.docx").exists())


if __name__ == "__main__":
    unittest.main()
