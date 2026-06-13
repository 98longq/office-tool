from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from office_tool.ai.deepseek import DeepSeekTextReviewer, extract_json_text
from office_tool.config import AIReviewOptions


class AIReviewTests(unittest.TestCase):
    def test_extracts_json_from_fenced_response(self):
        content = """```json
[
  {"category": "措辞", "severity": "warning", "message": "表述偏口语", "quote": "马上办", "suggestion": "改为及时办理"}
]
```"""
        text = extract_json_text(content)
        findings = DeepSeekTextReviewer(AIReviewOptions()).parse_findings(text)

        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].severity, "warning")
        self.assertEqual(findings[0].category, "措辞")

    def test_non_json_response_becomes_info_finding(self):
        findings = DeepSeekTextReviewer(AIReviewOptions()).parse_findings("请检查标题。")

        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].severity, "info")
        self.assertIn("非 JSON", findings[0].message)


if __name__ == "__main__":
    unittest.main()
