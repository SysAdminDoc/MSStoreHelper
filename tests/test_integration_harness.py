#!/usr/bin/env python3

import json
import os
import tempfile
import unittest
from unittest.mock import patch

from MSStoreHelper import StoreAPI


class FakeResponse:
    def __init__(self, status_code=200, text="", json_data=None):
        self.status_code = status_code
        self.text = text
        self._json_data = json_data or {}
        self.headers = {}
        self.url = "https://example.test"

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def json(self):
        return self._json_data


class FakeDownloadResponse:
    headers = {"content-length": "11"}

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return False

    def raise_for_status(self):
        return None

    def iter_content(self, chunk_size=8192):
        yield b"partial"
        raise RuntimeError("network dropped")


class FakeRunResult:
    def __init__(self, returncode=0, stdout="", stderr=""):
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr


class MockedIntegrationHarnessTests(unittest.TestCase):
    def test_storeedgefd_search_fixture_without_network(self):
        payload = {
            "Data": [
                {
                    "PackageIdentifier": "9N0DX20HK701",
                    "PackageName": "Windows Terminal",
                    "Publisher": "Microsoft Corporation",
                }
            ]
        }
        with patch("MSStoreHelper.requests.post", return_value=FakeResponse(json_data=payload)) as post_mock:
            diagnostic = StoreAPI.search_store_with_diagnostics("terminal")

        self.assertEqual(diagnostic["Errors"], [])
        self.assertEqual(diagnostic["Results"][0]["ProductId"], "9N0DX20HK701")
        self.assertEqual(post_mock.call_count, 1)

    def test_rgadguard_package_fixture_without_network(self):
        html = """
        <table class="tftable">
          <tr><td><a href="https://cdn.test/runtime.appx">Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.appx</a></td></tr>
          <tr><td><a href="https://cdn.test/app.msixbundle">Contoso.App_1.0.0.0_x64__abc.msixbundle</a></td></tr>
        </table>
        """
        with patch("MSStoreHelper.requests.post", return_value=FakeResponse(text=html)):
            diagnostic = StoreAPI.get_packages_with_diagnostics("9TEST")

        self.assertEqual(len(diagnostic["Packages"]), 2)
        self.assertEqual(diagnostic["Packages"][0]["Architecture"], "x64")
        self.assertTrue(any(package["IsBundle"] for package in diagnostic["Packages"]))

    def test_failed_download_keeps_existing_file_and_part(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            target = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msix")
            with open(target, "wb") as handle:
                handle.write(b"existing")

            with patch("MSStoreHelper.requests.get", return_value=FakeDownloadResponse()):
                ok, message = StoreAPI.download_file("https://example.test/app.msix", target)

            self.assertFalse(ok)
            self.assertIn("network dropped", message)
            with open(target, "rb") as handle:
                self.assertEqual(handle.read(), b"existing")
            self.assertTrue(os.path.exists(f"{target}.part"))

    def test_signature_failure_is_reported_without_admin_rights(self):
        signature = {
            "Status": "Valid",
            "Signer": "CN=Contoso",
            "Root": "CN=Contoso Root",
            "ChainValid": True,
        }
        with patch("MSStoreHelper.subprocess.run", return_value=FakeRunResult(0, json.dumps(signature), "")):
            ok, message = StoreAPI.verify_package_signature(r"C:\Packages\Contoso.msix")

        self.assertFalse(ok)
        self.assertIn("Contoso", message)

    def test_appx_install_failure_returns_powershell_output(self):
        with patch("MSStoreHelper.subprocess.run", return_value=FakeRunResult(1, "", "0x80073CF3 dependency missing")):
            ok, message = StoreAPI.install_package(r"C:\Packages\Contoso.msix")

        self.assertFalse(ok)
        self.assertIn("0x80073CF3", message)

    def test_intunewin_command_failure_is_raised_without_running_tool(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            downloads = os.path.join(temp_dir, "downloads")
            os.makedirs(downloads)
            package_path = os.path.join(downloads, "Contoso.App_1.0.0.0_x64__test.msixbundle")
            tool_path = os.path.join(temp_dir, "IntuneWinAppUtil.exe")
            with open(package_path, "wb") as handle:
                handle.write(b"package")
            with open(tool_path, "wb") as handle:
                handle.write(b"tool")

            with patch("MSStoreHelper.subprocess.run", return_value=FakeRunResult(87, "", "content prep failed")):
                with self.assertRaisesRegex(RuntimeError, "content prep failed"):
                    StoreAPI.create_intunewin_package(
                        [{"FileName": os.path.basename(package_path), "LocalPath": package_path}],
                        downloads,
                        os.path.join(temp_dir, "out", "Queue.intunewin"),
                        tool_path,
                        "x64",
                    )


if __name__ == "__main__":
    unittest.main()
