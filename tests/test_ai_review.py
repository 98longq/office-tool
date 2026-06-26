from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from office_tool.ai.deepseek import DeepSeekTextReviewer, extract_json_text, stream_text_to_chat_completion, strip_reasoning_text
from office_tool.config import AIReviewOptions


class AIReviewTests(unittest.TestCase):
    def test_extracts_json_from_fenced_response(self):
        content = """```json
[
  {"category": "wording", "severity": "warning", "message": "Too casual", "quote": "ASAP", "suggestion": "Use a formal phrase."}
]
```"""
        text = extract_json_text(content)
        findings = DeepSeekTextReviewer(AIReviewOptions()).parse_findings(text)

        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].severity, "warning")
        self.assertEqual(findings[0].category, "wording")

    def test_non_json_response_becomes_info_finding(self):
        findings = DeepSeekTextReviewer(AIReviewOptions()).parse_findings("Please check the title.")

        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].severity, "info")
        self.assertIn("标准 JSON", findings[0].message)

    def test_streaming_chunks_are_combined(self):
        raw = "\n".join(
            [
                '{"choices":[{"delta":{"content":"[{\\"category\\":\\"risk\\","}}]}',
                '{"choices":[{"delta":{"content":"\\"severity\\":\\"info\\",\\"message\\":\\"ok\\"}]"}}]}',
            ]
        )

        response = stream_text_to_chat_completion(raw)
        content = response["choices"][0]["message"]["content"]

        self.assertIn('"category":"risk"', content)
        findings = DeepSeekTextReviewer(AIReviewOptions()).parse_findings(content)
        self.assertEqual(findings[0].category, "risk")

    def test_full_chat_completions_url_is_not_appended_again(self):
        reviewer = DeepSeekTextReviewer(AIReviewOptions(base_url="http://ai.local/deepseek/v1/chat/completions"))

        self.assertEqual(reviewer._resolve_url(), "http://ai.local/deepseek/v1/chat/completions")

    def test_authorization_can_be_raw_token(self):
        reviewer = DeepSeekTextReviewer(AIReviewOptions(api_key="abc", auth_prefix=""))

        self.assertEqual(reviewer._headers()["Authorization"], "abc")

    def test_strip_reasoning_text(self):
        content = "<think>private reasoning</think>\n[]"

        self.assertEqual(strip_reasoning_text(content), "[]")


if __name__ == "__main__":
    unittest.main()
