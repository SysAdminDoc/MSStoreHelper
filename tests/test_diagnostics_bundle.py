#!/usr/bin/env python3

import json
import os
import tempfile
import unittest
import zipfile

from MSStoreHelper import StoreAPI
from diagnostic_bundle import (
    DiagnosticRedactionError,
    diagnostic_preview_text,
    write_prepared_bundle,
)


class DiagnosticsBundleTests(unittest.TestCase):
    def test_redact_diagnostic_text_removes_paths_and_secrets(self):
        temp_path = os.path.join(tempfile.gettempdir(), "msstorehelper-secret.txt")
        text = f"{temp_path}\napi_key=abc123\nauthorization: BearerToken"

        redacted = StoreAPI.redact_diagnostic_text(text)

        self.assertNotIn(temp_path, redacted)
        self.assertIn("%", redacted)
        self.assertIn("api_key=[REDACTED]", redacted)
        self.assertIn("authorization: [REDACTED]", redacted)

    def test_text_redactor_handles_headers_quotes_urls_and_paths(self):
        text = "\n".join([
            "Authorization: Bearer abc.def.ghi extra data",
            'password = "spaced secret value"',
            "--token 'quoted token value'",
            (
                "Source: https://user:pass@example.test/file"
                "?safe=1&token=url-secret&X-Amz-Signature=signature"
            ),
            (
                r'Command: "C:\Users\Alice\Downloads\tool.exe" '
                r'--api-key "command secret"'
            ),
            r"Artifact: c:\USERS\ALICE\Temp\package.msix",
            r"Share: \\server\private\diagnostics.txt",
        ])

        redacted = StoreAPI.redact_diagnostic_text(text)

        for secret in (
            "abc.def.ghi",
            "spaced secret value",
            "quoted token value",
            "user:pass",
            "url-secret",
            "signature",
            "command secret",
            r"C:\Users\Alice",
            r"c:\USERS\ALICE",
            r"\\server\private",
        ):
            self.assertNotIn(secret, redacted)
        self.assertIn("safe=1", redacted)
        self.assertIn("Authorization: [REDACTED]", redacted)
        self.assertIn("password = [REDACTED]", redacted)
        self.assertIn("%LOCAL_PATH%", redacted)

    def test_structured_redaction_omits_secret_keys_and_safely_maps_paths(self):
        payload = {
            "Authorization": "Bearer should-disappear",
            "AccessToken": "also-disappear",
            "Nested": {
                "Password": "quoted secret",
                "SourceUrl": (
                    "https://user:pass@example.test/api"
                    "?safe=yes&sig=remove-me"
                ),
                "LocalPath": r"C:\Sensitive Folder\package.msix",
                "Executable": r"D:\Tools\private.exe",
                "Command": (
                    r'Add-AppxPackage "C:\Private\package.msix" '
                    '--token "command-token"'
                ),
                "StoreQuery": {
                    "Ring": "WIS",
                    "Language": "de-DE",
                    "token": "nested-query-secret",
                },
            },
        }

        redacted = StoreAPI.redact_diagnostic_structure(payload)
        serialized = json.dumps(redacted)

        self.assertNotIn("Authorization", redacted)
        self.assertNotIn("AccessToken", redacted)
        self.assertNotIn("Password", redacted["Nested"])
        self.assertEqual(
            redacted["Nested"]["SourceUrl"],
            "https://example.test/api?safe=yes",
        )
        self.assertIn("LocalPath", redacted["Nested"])
        self.assertIn("%LOCAL_PATH%", redacted["Nested"]["LocalPath"])
        self.assertNotIn("nested-query-secret", serialized)
        self.assertNotIn("command-token", serialized)
        self.assertNotIn(r"C:\Private", serialized)
        self.assertEqual(
            redacted["Nested"]["StoreQuery"]["Language"],
            "de-DE",
        )

    def test_unsupported_structured_value_fails_closed(self):
        with self.assertRaisesRegex(
            DiagnosticRedactionError,
            "Unsupported",
        ):
            StoreAPI.redact_diagnostic_structure({
                "Payload": b"secret bytes",
            })

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
                [{
                    "FileName": os.path.basename(local_path),
                    "LocalPath": local_path,
                    "Sha256": "abc",
                    "StoreQuery": {"Ring": "WIS", "Language": "pt-BR", "Market": "BR"},
                }],
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
            self.assertEqual(queue[0]["StoreQuery"]["Language"], "pt-BR")
            self.assertNotIn(temp_dir, json.dumps(queue))
            self.assertNotIn("password=secret", app_log)
            self.assertIn("password=[REDACTED]", app_log)
            self.assertIn("Add-AppxPackage", transcript)

    def test_prepared_preview_and_zip_are_byte_for_byte_identical(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            bundle_path = os.path.join(temp_dir, "reviewed.zip")
            secret_path = os.path.join(
                temp_dir,
                "secret package.msix",
            )
            entries = StoreAPI.prepare_diagnostics_bundle(
                "9.9.9",
                "x64",
                False,
                secret_path,
                [{
                    "Key": "source",
                    "Detail": (
                        "Authorization: Bearer source-token"
                    ),
                    "SourceUrl": (
                        "https://example.test/check"
                        "?safe=1&token=query-token"
                    ),
                }],
                [{
                    "FileName": "package.msix",
                    "LocalPath": secret_path,
                    "LastError": (
                        'password = "queue secret with spaces"'
                    ),
                }],
                (
                    f"Command: Add-AppxPackage {secret_path}\n"
                    "Authorization: Bearer log-token"
                ),
            )
            preview = diagnostic_preview_text(entries)

            write_prepared_bundle(bundle_path, entries)
            with zipfile.ZipFile(bundle_path) as archive:
                archived = {
                    name: archive.read(name)
                    for name in archive.namelist()
                }

        self.assertEqual(archived, entries)
        for name in entries:
            self.assertIn(f"===== {name} =====", preview)
        for secret in (
            secret_path,
            "source-token",
            "query-token",
            "queue secret with spaces",
            "log-token",
        ):
            self.assertNotIn(secret, preview)
        self.assertIn("safe=1", preview)
        self.assertIn("%TEMP%", preview)


if __name__ == "__main__":
    unittest.main()
