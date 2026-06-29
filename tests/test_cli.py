#!/usr/bin/env python3

import io
import json
import os
import tempfile
import unittest
from unittest.mock import patch

from MSStoreHelper import StoreAPI, main, run_cli


class CliWorkflowTests(unittest.TestCase):
    def test_resolve_cli_app_maps_package_identity_to_catalog_product(self):
        def fail_search(*_args, **_kwargs):
            raise AssertionError("catalog identity should not require Store search")

        app, error = StoreAPI.resolve_cli_app("Microsoft.WindowsTerminal", searcher=fail_search)

        self.assertIsNone(error)
        self.assertEqual(app["Name"], "Windows Terminal")
        self.assertEqual(app["ProductId"], "9N0DX20HK701")
        self.assertEqual(app["ResolvedFrom"], "package-identity")

    def test_run_cli_search_emits_json_without_gui(self):
        stdout = io.StringIO()
        stderr = io.StringIO()
        diagnostic = {
            "Source": "fixture",
            "Errors": [],
            "Results": [
                {
                    "Name": "Windows Terminal",
                    "ProductId": "9N0DX20HK701",
                    "Publisher": "Microsoft Corporation",
                }
            ],
        }

        with patch("MSStoreHelper.MSStoreHelperApp") as app_mock:
            with patch.object(StoreAPI, "search_store_with_diagnostics", return_value=diagnostic):
                exit_code = run_cli(["--search", "terminal", "--json"], stdout, stderr)

        self.assertEqual(exit_code, 0)
        app_mock.assert_not_called()
        payload = json.loads(stdout.getvalue())
        self.assertEqual(payload["Results"][0]["ProductId"], "9N0DX20HK701")
        self.assertEqual(stderr.getvalue(), "")

    def test_run_cli_download_uses_selected_package_without_gui(self):
        package = {
            "FileName": "Microsoft.WindowsTerminal_1.0.0.0_x64__8wekyb3d8bbwe.msixbundle",
            "Url": "https://example.test/terminal.msixbundle",
            "Architecture": "x64",
            "FileType": "MSIXBUNDLE",
        }
        stdout = io.StringIO()
        stderr = io.StringIO()

        with tempfile.TemporaryDirectory() as temp_dir:
            def fake_download(_url, filepath, progress_callback=None, package=None):
                with open(filepath, "wb") as handle:
                    handle.write(b"package")
                return True, "Success"

            with patch("MSStoreHelper.MSStoreHelperApp") as app_mock:
                with patch.object(StoreAPI, "get_packages_with_diagnostics", return_value={"Packages": [package], "Errors": [], "Query": {"ProductId": "9N0DX20HK701"}}):
                    with patch.object(StoreAPI, "smart_select", return_value=[package.copy()]):
                        with patch.object(StoreAPI, "order_packages_for_install", side_effect=lambda packages, _arch: packages):
                            with patch.object(StoreAPI, "download_file", side_effect=fake_download):
                                exit_code = run_cli(
                                    ["--download", "Microsoft.WindowsTerminal", "--output", temp_dir, "--json"],
                                    stdout,
                                    stderr,
                                )

        self.assertEqual(exit_code, 0)
        app_mock.assert_not_called()
        payload = json.loads(stdout.getvalue())
        self.assertEqual(payload["Action"], "download")
        self.assertEqual(payload["Packages"][0]["Status"], "downloaded")
        self.assertEqual(stderr.getvalue(), "")

    def test_run_cli_install_skips_current_package_without_installing(self):
        package = {
            "FileName": "Microsoft.WindowsTerminal_1.0.0.0_x64__8wekyb3d8bbwe.msixbundle",
            "Url": "https://example.test/terminal.msixbundle",
            "Architecture": "x64",
            "FileType": "MSIXBUNDLE",
        }
        stdout = io.StringIO()
        stderr = io.StringIO()

        with tempfile.TemporaryDirectory() as temp_dir:
            def fake_download(_url, filepath, progress_callback=None, package=None):
                with open(filepath, "wb") as handle:
                    handle.write(b"package")
                return True, "Success"

            with patch("MSStoreHelper.IS_ADMIN", True):
                with patch.object(StoreAPI, "get_packages_with_diagnostics", return_value={"Packages": [package], "Errors": [], "Query": {"ProductId": "9N0DX20HK701"}}):
                    with patch.object(StoreAPI, "smart_select", return_value=[package.copy()]):
                        with patch.object(StoreAPI, "order_packages_for_install", side_effect=lambda packages, _arch: packages):
                            with patch.object(StoreAPI, "download_file", side_effect=fake_download):
                                with patch.object(StoreAPI, "should_skip_installed_package", return_value=(True, "1.0.0.0", "Microsoft.WindowsTerminal")):
                                    with patch.object(StoreAPI, "verify_package_signature") as signature_mock:
                                        with patch.object(StoreAPI, "install_package") as install_mock:
                                            exit_code = run_cli(
                                                ["--install", "Microsoft.WindowsTerminal", "--output", temp_dir, "--json"],
                                                stdout,
                                                stderr,
                                            )

        self.assertEqual(exit_code, 0)
        signature_mock.assert_not_called()
        install_mock.assert_not_called()
        payload = json.loads(stdout.getvalue())
        self.assertEqual(payload["Skipped"], 1)
        self.assertEqual(payload["Failed"], 0)
        self.assertEqual(payload["Packages"][-1]["Status"], "skipped")

    def test_main_with_args_runs_cli_without_creating_window(self):
        with patch("MSStoreHelper.run_cli", return_value=0) as cli_mock:
            with patch("MSStoreHelper.MSStoreHelperApp") as app_mock:
                exit_code = main(["--search", "terminal"])

        self.assertEqual(exit_code, 0)
        cli_mock.assert_called_once_with(["--search", "terminal"])
        app_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main()
