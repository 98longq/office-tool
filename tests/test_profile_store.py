import tempfile
import unittest
from pathlib import Path

from office_tool.config import OfficeToolConfig
from office_tool.profile_store import ConfigProfileStore


class ConfigProfileStoreTests(unittest.TestCase):
    def test_profile_round_trip_rename_delete_and_secret_removal(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_profiles_") as tmp:
            store = ConfigProfileStore(Path(tmp))
            config = OfficeToolConfig()
            config.audit.profile = "red_head"
            config.generation.add_red_head = True
            config.generation.red_head_title = "某某集团有限公司文件"
            config.ai_review.api_key = "temporary-secret"

            path = store.save("常用红头", config)
            loaded = store.load_all()["常用红头"]

            self.assertTrue(path.exists())
            self.assertEqual(loaded.audit.profile, "red_head")
            self.assertTrue(loaded.generation.add_red_head)
            self.assertEqual(loaded.generation.red_head_title, "某某集团有限公司文件")
            self.assertEqual(loaded.ai_review.api_key, "")

            store.rename("常用红头", "集团红头")
            self.assertEqual(list(store.load_all()), ["集团红头"])
            self.assertTrue(store.delete("集团红头"))
            self.assertEqual(store.load_all(), {})

    def test_duplicate_name_overwrites_same_profile(self):
        with tempfile.TemporaryDirectory(prefix="office_tool_profiles_") as tmp:
            store = ConfigProfileStore(Path(tmp))
            first = OfficeToolConfig()
            second = OfficeToolConfig()
            second.audit.profile = "meeting_minutes"

            first_path = store.save("同名配置", first)
            second_path = store.save("同名配置", second)

            self.assertEqual(first_path, second_path)
            self.assertEqual(store.load_all()["同名配置"].audit.profile, "meeting_minutes")


if __name__ == "__main__":
    unittest.main()
