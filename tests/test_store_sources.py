#!/usr/bin/env python3

import unittest
from unittest.mock import patch

from MSStoreHelper import StoreAPI
from store_sources import (
    detect_source_health,
    package_lookup_fallbacks,
    request_with_retries,
)


class FakeResponse:
    def __init__(self, status_code=200, text="", json_data=None):
        self.status_code = status_code
        self.text = text
        self._json_data = json_data or {}
        self.url = "https://example.test"

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def json(self):
        return self._json_data


class FakeRunResult:
    returncode = 0
    stdout = "v1.9.0\n"
    stderr = ""


class StoreSourceTests(unittest.TestCase):
    def test_request_with_retries_recovers_from_retryable_status(self):
        responses = [FakeResponse(503), FakeResponse(200, json_data={"ok": True})]

        response, errors = request_with_retries("Test Source", lambda: responses.pop(0), attempts=2)

        self.assertEqual(response.status_code, 200)
        self.assertEqual(errors, ["attempt 1: HTTP 503"])

    def test_detect_source_health_reports_http_and_commands(self):
        def which(command):
            return f"C:\\Tools\\{command}.exe" if command in {"winget", "store"} else None

        statuses = detect_source_health(
            storeedge_request=lambda: FakeResponse(200),
            rgadguard_request=lambda: FakeResponse(200),
            which=which,
            run=lambda *_args, **_kwargs: FakeRunResult(),
        )

        by_key = {status["Key"]: status for status in statuses}
        self.assertTrue(by_key["storeedgefd"]["Available"])
        self.assertTrue(by_key["rg-adguard"]["Available"])
        self.assertTrue(by_key["winget"]["Available"])
        self.assertTrue(by_key["store-cli"]["Available"])
        self.assertEqual(by_key["winget"]["Version"], "v1.9.0")

    def test_package_lookup_fallbacks_include_available_sources(self):
        fallbacks = package_lookup_fallbacks("9N0DX20HK701", [
            {"Key": "winget", "Available": True},
            {"Key": "store-cli", "Available": True},
        ])

        commands = "\n".join(fallback["Command"] for fallback in fallbacks)
        self.assertIn("winget install --source msstore --id 9N0DX20HK701", commands)
        self.assertIn("store install 9N0DX20HK701", commands)

    def test_package_diagnostics_parse_rgadguard_table(self):
        html = """
        <table class="tftable">
          <tr><td><a href="https://cdn.test/app.msixbundle">Contoso.App_1.2.3.4_x64__abc.msixbundle</a></td></tr>
          <tr><td><a href="https://cdn.test/app.BlockMap">Contoso.App_1.2.3.4_x64__abc.BlockMap</a></td></tr>
        </table>
        """
        with patch("MSStoreHelper.requests.post", return_value=FakeResponse(200, text=html)):
            diagnostic = StoreAPI.get_packages_with_diagnostics("9TEST")

        self.assertEqual(diagnostic["Source"], "RG-Adguard package proxy")
        self.assertEqual(diagnostic["Errors"], [])
        self.assertEqual(len(diagnostic["Packages"]), 1)
        self.assertEqual(diagnostic["Packages"][0]["Architecture"], "x64")
        self.assertEqual(diagnostic["Packages"][0]["FileType"], "MSIXBUNDLE")

    def test_package_diagnostics_expose_fallbacks_on_missing_table(self):
        with patch("MSStoreHelper.requests.post", return_value=FakeResponse(200, text="<html></html>")):
            with patch("MSStoreHelper.StoreAPI.detect_source_health", return_value=[{"Key": "winget", "Available": True}]):
                diagnostic = StoreAPI.get_packages_with_diagnostics("9TEST")

        self.assertEqual(diagnostic["Packages"], [])
        self.assertIn("package table", diagnostic["Errors"][0])
        self.assertEqual(diagnostic["Fallbacks"][0]["Source"], "WinGet")


if __name__ == "__main__":
    unittest.main()
