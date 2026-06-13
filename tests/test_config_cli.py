from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from office_tool.cli import main
from office_tool.config import OfficeToolConfig


class ConfigCliTests(unittest.TestCase):
    def test_nested_config_override(self):
        config = OfficeToolConfig()
        config.set_path("page.margin_top_cm", 3.9)
        config.set_path("styles.body.font", "仿宋")

        self.assertEqual(config.page.margin_top_cm, 3.9)
        self.assertEqual(config.styles["body"].font, "仿宋")

    def test_init_config_command(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_config_") as tmp:
            output = Path(tmp) / "config.json"
            code = main(["init-config", "-o", str(output)])

            self.assertEqual(code, 0)
            self.assertTrue(output.exists())
            self.assertIn("styles", output.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
