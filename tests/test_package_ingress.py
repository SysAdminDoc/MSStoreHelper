#!/usr/bin/env python3

import json
import os
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from MSStoreHelper import (
    MSStoreHelperApp,
    StoreAPI,
    _cli_download_selected,
)
from package_ingress import (
    PackageIngressError,
    ensure_path_within_root,
    package_path,
    validate_existing_package_path,
    validate_package_filename,
    validate_package_record,
    validate_package_url,
    validate_response_redirects,
)
from test_trust_utils import mark_package_trusted


class RedirectResponse:
    def __init__(self, url, *, location=None, history=None):
        self.url = url
        self.history = history or []
        self.headers = {}
        if location:
            self.headers["location"] = location
        self.status_code = 200

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return False

    def raise_for_status(self):
        return None

    def iter_content(self, chunk_size=8192):
        yield b"package"


class QueueHarness:
    def __init__(self):
        self.download_queue = []
        self.messages = []

    def _log(self, level, message):
        self.messages.append((level, message))

    def _save_download_state(self):
        raise AssertionError("invalid package must not mutate persisted state")


class PackageIngressTests(unittest.TestCase):
    def test_filename_validator_accepts_windows_safe_package_names(self):
        filenames = [
            "Contoso.App_1.0.0.0_x64__test.msix",
            "Contoso.O'Hare_1.0.0.0_neutral__test.appx",
            "Contoso.$(literal)_1.0.0.0_x64__test.msixbundle",
            "München.App_1.0.0.0_x64__test.appxbundle",
            "Contoso.Encrypted_1.0.0.0_x64__test.emsix",
        ]

        self.assertEqual(
            [validate_package_filename(filename) for filename in filenames],
            filenames,
        )

    def test_filename_validator_rejects_traversal_devices_ads_and_bad_extensions(self):
        invalid = [
            r"..\..\Users\Public\payload.msix",
            "nested/payload.msix",
            r"C:\Temp\payload.msix",
            r"\\server\share\payload.msix",
            "CON.msix",
            "CONOUT$.msix",
            "LPT9.appx",
            "COM¹.msix",
            "payload.msix:stream",
            "payload.msix.exe",
            "payload.msix.",
            " payload.msix",
            "payload.msix ",
            "payload\n.msix",
            ".msix",
            "payload",
        ]

        for filename in invalid:
            with self.subTest(filename=repr(filename)):
                with self.assertRaises(PackageIngressError):
                    validate_package_filename(filename)

    def test_url_validator_rejects_unsafe_and_ambiguous_urls(self):
        invalid = [
            "file:///C:/Temp/payload.msix",
            "data:application/octet-stream,payload",
            "ftp://example.test/payload.msix",
            "https://user:secret@example.test/payload.msix",
            "https://example.test/payload.msix#fragment",
            "https://example.test:invalid/payload.msix",
            "https://[invalid/payload.msix",
            "https://example.test/path%0Apayload.msix",
            "https://example.test/path%5Cpayload.msix",
            "https://example.test/bad%escape.msix",
            "https://example.test/path payload.msix",
        ]

        for url in invalid:
            with self.subTest(url=url):
                with self.assertRaises(PackageIngressError):
                    validate_package_url(url)

        self.assertEqual(
            validate_package_url("https://cdn.test/payload.msix?token=opaque"),
            "https://cdn.test/payload.msix?token=opaque",
        )
        self.assertEqual(
            validate_package_url("http://cdn.test/payload.msix"),
            "http://cdn.test/payload.msix",
        )

    def test_redirect_validator_rejects_https_downgrade(self):
        first = RedirectResponse(
            "https://cdn.test/payload.msix",
            location="http://cdn.test/payload.msix",
        )
        final = RedirectResponse(
            "http://cdn.test/payload.msix",
            history=[first],
        )

        with self.assertRaisesRegex(PackageIngressError, "downgraded"):
            validate_response_redirects(
                "https://cdn.test/payload.msix",
                final,
            )

    def test_confined_paths_reject_escape_and_symlink_targets(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = os.path.join(temp_dir, "root")
            outside = os.path.join(temp_dir, "outside")
            os.makedirs(root)
            os.makedirs(outside)
            outside_package = os.path.join(
                outside,
                "Contoso.App_1.0.0.0_x64__test.msix",
            )
            with open(outside_package, "wb") as handle:
                handle.write(b"package")

            with self.assertRaises(PackageIngressError):
                ensure_path_within_root(root, outside_package)
            with self.assertRaises(PackageIngressError):
                package_path(root, r"..\outside\payload.msix")

            link = os.path.join(
                root,
                "Contoso.Link_1.0.0.0_x64__test.msix",
            )
            try:
                os.symlink(outside_package, link)
            except OSError:
                return
            with self.assertRaises(PackageIngressError):
                validate_existing_package_path(
                    link,
                    root=root,
                    require_file=True,
                )

    def test_package_record_requires_matching_safe_filename(self):
        with self.assertRaisesRegex(PackageIngressError, "safe filename"):
            validate_package_record({
                "FileName": "Contoso.App_1.0.0.0_x64__test.msix",
                "SafeFileName": "Other.App_1.0.0.0_x64__test.msix",
            })

    def test_proxy_parser_drops_unsafe_rows_before_the_gui_queue(self):
        html = """
        <table class="tftable">
          <tr><td><a href="https://cdn.test/safe.msix">Contoso.Safe_1.0.0.0_x64__test.msix</a></td></tr>
          <tr><td><a href="https://cdn.test/escape.msix">..\\..\\Users\\Public\\escape.msix</a></td></tr>
          <tr><td><a href="file:///C:/Temp/payload.msix">Contoso.BadUrl_1.0.0.0_x64__test.msix</a></td></tr>
        </table>
        """
        response = SimpleNamespace(text=html)

        with patch(
            "MSStoreHelper.request_with_retries",
            return_value=(response, []),
        ):
            diagnostic = StoreAPI.get_packages_with_diagnostics("9TEST")

        self.assertEqual(
            [package["FileName"] for package in diagnostic["Packages"]],
            ["Contoso.Safe_1.0.0.0_x64__test.msix"],
        )
        self.assertEqual(len(diagnostic["Errors"]), 2)
        self.assertTrue(all("Rejected unsafe" in error for error in diagnostic["Errors"]))

    def test_gui_queue_rejects_persisted_or_injected_unsafe_metadata(self):
        harness = QueueHarness()
        count = MSStoreHelperApp._queue_unique_packages(
            harness,
            [{
                "FileName": r"..\..\Users\Public\payload.msix",
                "Url": "https://cdn.test/payload.msix",
            }],
            "x64",
        )

        self.assertEqual(count, 0)
        self.assertEqual(harness.download_queue, [])
        self.assertIn("Rejected unsafe package metadata", harness.messages[0][1])

    def test_cli_download_rejects_escape_before_network_or_write(self):
        with tempfile.TemporaryDirectory() as output_path:
            with patch("MSStoreHelper.requests.get") as get_mock:
                downloaded, records = _cli_download_selected(
                    [{
                        "FileName": r"..\..\Users\Public\payload.msix",
                        "Url": "https://cdn.test/payload.msix",
                    }],
                    output_path,
                    SimpleNamespace(write=lambda _text: None),
                )

        self.assertEqual(downloaded, [])
        self.assertEqual(records[0]["Status"], "failed")
        get_mock.assert_not_called()

    def test_download_rejects_out_of_root_path_and_unsafe_redirect(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = os.path.join(temp_dir, "root")
            outside = os.path.join(
                temp_dir,
                "Contoso.App_1.0.0.0_x64__test.msix",
            )
            os.makedirs(root)

            with patch("MSStoreHelper.requests.get") as get_mock:
                ok, message = StoreAPI.download_file(
                    "https://cdn.test/payload.msix",
                    outside,
                    destination_root=root,
                )
            self.assertFalse(ok)
            self.assertIn("escapes", message)
            get_mock.assert_not_called()

            target = os.path.join(
                root,
                "Contoso.App_1.0.0.0_x64__test.msix",
            )
            response = RedirectResponse("http://cdn.test/payload.msix")
            with patch("MSStoreHelper.requests.get", return_value=response):
                ok, message = StoreAPI.download_file(
                    "https://cdn.test/payload.msix",
                    target,
                    destination_root=root,
                )
            self.assertFalse(ok)
            self.assertIn("downgraded", message)
            self.assertFalse(os.path.exists(target))

    def test_cache_and_exports_reject_adversarial_package_names(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            downloads = os.path.join(temp_dir, "downloads")
            os.makedirs(downloads)
            source = os.path.join(
                downloads,
                "Contoso.App_1.0.0.0_x64__test.msix",
            )
            with open(source, "wb") as handle:
                handle.write(b"package")
            package = {
                "FileName": r"..\..\Users\Public\payload.msix",
                "LocalPath": source,
            }

            ok, _message = StoreAPI.cache_downloaded_artifact(
                package,
                os.path.join(temp_dir, "cache"),
            )
            self.assertFalse(ok)
            with self.assertRaises(PackageIngressError):
                StoreAPI.write_appinstaller_export(
                    [package],
                    downloads,
                    os.path.join(temp_dir, "Queue.appinstaller"),
                )
            with self.assertRaises(PackageIngressError):
                StoreAPI.prepare_intune_package_source(
                    [package],
                    os.path.join(temp_dir, "staging"),
                    downloads,
                )

    def test_corrupt_state_and_cache_paths_do_not_reenter_workflows(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "download-state.json")
            with open(state_path, "w", encoding="utf-8") as handle:
                json.dump({
                    "Version": 1,
                    "OutputPath": temp_dir,
                    "Queue": [
                        {
                            "FileName": r"..\..\Users\Public\payload.msix",
                            "Url": "https://cdn.test/payload.msix",
                        },
                        {
                            "FileName": "Contoso.Bad_1.0.0.0_x64__test.msix",
                            "Url": "file:///C:/Temp/payload.msix",
                        },
                    ],
                }, handle)

            self.assertEqual(StoreAPI.load_download_state(state_path)["Queue"], [])

            cache = os.path.join(temp_dir, "cache")
            outside = os.path.join(
                temp_dir,
                "Contoso.Outside_1.0.0.0_x64__test.msix",
            )
            os.makedirs(cache)
            with open(outside, "wb") as handle:
                handle.write(b"package")
            metadata = StoreAPI.artifact_metadata(
                {"FileName": os.path.basename(outside)},
                outside,
            )
            manifest = {
                "Version": 2,
                "Artifacts": {metadata["FileName"]: metadata},
                "History": {"contoso.outside": [metadata]},
            }
            StoreAPI.save_cache_manifest(cache, manifest)

            self.assertEqual(StoreAPI.cache_history_entries([cache]), [])

    def test_powershell_receives_package_path_only_through_environment(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            filename = "Contoso.$(literal)_1.0.0.0_x64__test.msix"
            package_path_value = os.path.join(temp_dir, filename)
            with open(package_path_value, "wb") as handle:
                handle.write(b"package")
            package = {"FileName": filename}
            mark_package_trusted(package, package_path_value)
            result = SimpleNamespace(returncode=0, stdout="", stderr="")

            with patch("MSStoreHelper.subprocess.run", return_value=result) as run_mock:
                ok, message = StoreAPI.install_package(
                    package_path_value,
                    package,
                )

            self.assertTrue(ok, message)
            command = run_mock.call_args.args[0][-1]
            environment = run_mock.call_args.kwargs["env"]
            self.assertNotIn(package_path_value, command)
            self.assertNotIn("$(literal)", command)
            self.assertEqual(
                environment["MSSTOREHELPER_PACKAGE_PATH"],
                os.path.realpath(package_path_value),
            )

    def test_rollback_rejects_command_like_identity_before_powershell(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            package_path_value = os.path.join(
                temp_dir,
                "Contoso.App_1.0.0.0_x64__test.msix",
            )
            with open(package_path_value, "wb") as handle:
                handle.write(b"package")

            with patch("MSStoreHelper.subprocess.run") as run_mock:
                ok, message = StoreAPI.rollback_package(
                    "Contoso.App'; Start-Process calc",
                    package_path_value,
                )

        self.assertFalse(ok)
        self.assertEqual(message, "Rollback package identity is invalid")
        run_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main()
