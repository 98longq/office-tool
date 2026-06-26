import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from office_tool.config import AIReviewOptions
from office_tool.gui import OfficeToolGUI, _attached_popup_geometry, _default_gui_config, _profile_label


class GuiLogicTests(unittest.TestCase):
    def test_virtual_direct_text_path_is_treated_as_single_document(self):
        self.assertTrue(OfficeToolGUI._is_single_document_source([Path("直接输入文稿.txt")]))

    def test_default_gui_config_uses_standard_profile_and_preserves_ai(self):
        ai = AIReviewOptions(enabled=True, base_url="http://ai.local", model="local-model")

        config = _default_gui_config(ai)

        self.assertEqual(config.audit.profile, "standard")
        self.assertTrue(config.ai_review.enabled)
        self.assertEqual(config.ai_review.model, "local-model")
        self.assertIsNot(config.ai_review, ai)
        self.assertEqual(config.generation.document_number, "某发〔2026〕1号")

    def test_profile_label_matches_all_document_schemes(self):
        self.assertEqual(_profile_label("standard"), "普通公文")
        self.assertEqual(_profile_label("red_head"), "红头文件")
        self.assertEqual(_profile_label("letter_head"), "红头文件（函）")
        self.assertEqual(_profile_label("meeting_minutes"), "红头文件（会议纪要）")

    def test_config_popup_is_attached_below_anchor(self):
        geometry = _attached_popup_geometry(
            anchor_x=900,
            anchor_y=120,
            anchor_width=160,
            anchor_height=36,
            requested_height=240,
            screen_x=0,
            screen_y=0,
            screen_width=1920,
            screen_height=1080,
        )

        self.assertEqual(geometry, (160, 240, 900, 158))

    def test_config_popup_clamps_long_content_to_screen(self):
        geometry = _attached_popup_geometry(
            anchor_x=10,
            anchor_y=700,
            anchor_width=140,
            anchor_height=36,
            requested_height=300,
            screen_x=0,
            screen_y=0,
            screen_width=1366,
            screen_height=768,
        )

        self.assertEqual(geometry, (140, 300, 10, 398))


if __name__ == "__main__":
    unittest.main()
