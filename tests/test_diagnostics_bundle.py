#!/usr/bin/env python3

import json
import os
import tempfile
import unittest
import zipfile

from MSStoreHelper import StoreAPI


class DiagnosticsBundleTests(unittest.TestCase):
    def test_redact_diagnostic_text_removes_paths_and_secrets(self):
        temp_path = os.path.join(tempfile.gettempdir(), "msstorehelper-secret.txt")
        text = f"{temp_path}\napi_key=abc123\nauthorization: BearerToken"

        redacted = StoreAPI.redact_diagnostic_text(text)

        self.assertNotIn(temp_path, redacted)
        self.assertIn("%", redacted)
        self.assertIn("api_key=[REDACTED]", redacted)
        self.assertIn("authorization: [REDACTED]", redacted)

    def test_write_diagnostics_bundle_contains_redacted_support_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            bundle_path = os.path.join(temp_dir, "diagnostics.zip")
            local_path = os.path.join(temp_dir, "App_1.0.0.0_x64__test.msixbundle")
            log_text = f"[12:00:00] INFO Command: Add-AppxPackage {local_path}\npassword=secret"

            StoreAPI.write_diagnostics_bundle(
                bundle_path,
                "9.9.9",
                "x64",
                False,
                temp_dir,
                [{"Key": "rg-adguard", "Available": False, "Detail": "HTTP 403"}],
                [{"FileName": os.path.basename(local_path), "LocalPath": local_path, "Sha256": "abc"}],
                log_text,
            )

            with zipfile.ZipFile(bundle_path) as archive:
                names = set(archive.namelist())
                self.assertIn("diagnostics.json", names)
                self.assertIn("source-health.json", names)
                self.assertIn("queue.json", names)
                self.assertIn("app-log.txt", names)
                self.assertIn("powershell-transcript.txt", names)
                diagnostics = json.loads(archive.read("diagnostics.json"))
                queue = json.loads(archive.read("queue.json"))
                app_log = archive.read("app-log.txt").decode("utf-8")
                transcript = archive.read("powershell-transcript.txt").decode("utf-8")

            self.assertEqual(diagnostics["AppVersion"], "9.9.9")
            self.assertEqual(diagnostics["QueueCount"], 1)
            self.assertNotIn(temp_dir, json.dumps(queue))
            self.assertNotIn("password=secret", app_log)
            self.assertIn("password=[REDACTED]", app_log)
            self.assertIn("Add-AppxPackage", transcript)


if __name__ == "__main__":
    unittest.main()
