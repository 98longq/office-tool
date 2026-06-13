from pathlib import Path
import sys
import tempfile
import unittest

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from office_tool.formatter import OfficialDocumentFormatter
from office_tool.io import load_document


class FormatterTests(unittest.TestCase):
    def test_format_red_head_keeps_title_separate(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_test_") as tmp:
            source = Path(tmp) / "red.docx"
            output = Path(tmp) / "red_out.docx"
            doc = Document()
            doc.add_paragraph("某某市人民政府文件")
            doc.add_paragraph("某政发〔2026〕1号")
            doc.add_paragraph("关于推进办公助手建设的通知")
            doc.add_paragraph("各区人民政府：")
            doc.add_paragraph("一、总体要求")
            doc.add_paragraph("正文内容。")
            doc.add_paragraph("某某市人民政府")
            doc.add_paragraph("2026年6月13日")
            doc.save(source)

            report = OfficialDocumentFormatter().format_path(source, output)
            formatted = Document(output)

            self.assertTrue(report.is_red_head)
            self.assertEqual(formatted.paragraphs[0].runs[0].font.color.rgb, RGBColor(0xFF, 0, 0))
            self.assertEqual(formatted.paragraphs[2].alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertEqual(formatted.paragraphs[2].runs[0].font.size.pt, 22)
            self.assertAlmostEqual(formatted.sections[0].top_margin.cm, 3.7, places=2)

    def test_text_input_loads_as_document(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_txt_") as tmp:
            source = Path(tmp) / "sample.txt"
            source.write_text("关于测试的通知\n各部门：\n正文。", encoding="utf-8")

            doc, kind = load_document(source)

            self.assertEqual(kind, "txt")
            self.assertEqual([p.text for p in doc.paragraphs[:3]], ["关于测试的通知", "各部门：", "正文。"])


if __name__ == "__main__":
    unittest.main()
