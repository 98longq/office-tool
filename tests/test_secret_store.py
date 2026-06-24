from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from office_tool.secret_store import protect_secret, unprotect_secret


@unittest.skipUnless(sys.platform == "win32", "DPAPI is only available on Windows")
class SecretStoreTests(unittest.TestCase):
    def test_dpapi_round_trip_does_not_store_plaintext(self):
        protected = protect_secret("temporary-test-key")

        self.assertTrue(protected)
        self.assertNotIn("temporary-test-key", protected)
        self.assertEqual(unprotect_secret(protected), "temporary-test-key")


if __name__ == "__main__":
    unittest.main()
