import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from office_tool.config import OfficeToolConfig
from office_tool.services import audit_many, collect_document_inputs, format_many, summarize_results


class ServiceTests(unittest.TestCase):
    def test_content_generation_runs_only_on_export_service(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_generation_service_") as tmp:
            root = Path(tmp)
            source = root / "plain.txt"
            output = root / "generated.docx"
            source.write_text("关于测试工作的通知\n有关单位：\n请认真落实。", encoding="utf-8")
            config = OfficeToolConfig()
            config.audit.profile = "red_head"
            config.generation.add_red_head = True
            config.generation.red_head_title = "某某集团有限公司文件"
            config.generation.document_number = "某发〔2026〕68号"

            audit_result = audit_many([source], config)[0]
            format_result = format_many([source], output, config)[0]

            self.assertNotIn("generated_content", audit_result.report.stats)
            self.assertEqual(format_result.report.stats["generated_content"], ["red_head"])
            self.assertTrue(output.exists())

    def test_collect_document_inputs_reads_supported_files_from_directory(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_services_") as tmp:
            root = Path(tmp)
            (root / "a.txt").write_text("标题\n正文", encoding="utf-8")
            (root / "b.md").write_text("# 标题", encoding="utf-8")
            (root / "c.doc").write_bytes(b"legacy-word-placeholder")
            (root / "skip.xlsx").write_text("", encoding="utf-8")

            paths = collect_document_inputs([root])

        self.assertEqual([path.name for path in paths], ["a.txt", "b.md", "c.doc"])

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
            self.assertTrue((report_dir / "sample_proofreading.json").exists())
            self.assertTrue((report_dir / "sample_proofreading.md").exists())
            self.assertTrue(formats[0].ok)
            self.assertTrue((output_dir / "sample_formatted.docx").exists())

    def test_ai_failure_does_not_block_formatted_output(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_ai_fallback_") as tmp:
            root = Path(tmp)
            source = root / "sample.txt"
            output = root / "sample.docx"
            source.write_text("关于测试工作的通知\n这是正文。", encoding="utf-8")
            config = OfficeToolConfig()
            config.ai_review.enabled = True
            config.ai_review.base_url = "http://ai.invalid/v1"

            with patch(
                "office_tool.services.DeepSeekTextReviewer.review_into_report",
                side_effect=TimeoutError("连接超时"),
            ):
                result = format_many([source], output, config)[0]

            self.assertTrue(result.ok)
            self.assertTrue(output.exists())
            self.assertIn("ai_proofreading_failed", {finding.code for finding in result.report.findings})


if __name__ == "__main__":
    unittest.main()
