from pathlib import Path
import sys
import tempfile
import unittest

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from office_tool.audit import OfficialDocumentAuditor
from office_tool.config import OfficeToolConfig


class AuditTests(unittest.TestCase):
    def test_detects_red_head_structure(self):
        doc = Document()
        doc.add_paragraph("某某市人民政府文件")
        doc.add_paragraph("某政发〔2026〕1号")
        doc.add_paragraph("签发人：张三")
        doc.add_paragraph("关于推进办公助手建设的通知")
        doc.add_paragraph("各区人民政府：")
        doc.add_paragraph("现将有关事项通知如下。")
        doc.add_paragraph("某某市人民政府")
        doc.add_paragraph("2026年6月13日")

        report = OfficialDocumentAuditor().audit_document(doc)

        self.assertTrue(report.is_red_head)
        self.assertEqual(report.profile, "red_head")
        self.assertEqual(report.count("error"), 0)
        roles = {element.name: element.text for element in report.elements}
        self.assertEqual(roles["red_head"], "某某市人民政府文件")
        self.assertEqual(roles["document_number"], "某政发〔2026〕1号")
        self.assertEqual(roles["title"], "关于推进办公助手建设的通知")

    def test_missing_title_is_error(self):
        doc = Document()
        doc.add_paragraph("某政发〔2026〕1号")

        report = OfficialDocumentAuditor().audit_document(doc)

        self.assertEqual(report.count("error"), 1)
        self.assertEqual(report.findings[0].code, "missing_title")

    def test_missing_red_head_document_number_warns(self):
        doc = Document()
        doc.add_paragraph("某某市人民政府文件")
        doc.add_paragraph("关于推进办公助手建设的通知")
        doc.add_paragraph("各区人民政府：")
        doc.add_paragraph("正文。")
        doc.add_paragraph("2026年6月13日")

        report = OfficialDocumentAuditor().audit_document(doc)

        self.assertTrue(report.is_red_head)
        self.assertTrue(any(f.code == "missing_document_number" for f in report.findings))

    def test_page_layout_finding_is_fixable(self):
        doc = Document()
        doc.add_paragraph("关于推进办公助手建设的通知")
        doc.add_paragraph("正文。")
        config = OfficeToolConfig()
        report = OfficialDocumentAuditor(config).audit_document(doc)

        layout_findings = [f for f in report.findings if f.code.startswith("layout_")]
        self.assertTrue(layout_findings)
        self.assertTrue(all(f.can_fix for f in layout_findings))


if __name__ == "__main__":
    unittest.main()
