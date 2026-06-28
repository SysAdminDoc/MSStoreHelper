#!/usr/bin/env python3

import json
import os
import tempfile
import unittest
from datetime import datetime, timezone

from MSStoreHelper import StoreAPI


class WingetExportTests(unittest.TestCase):
    def test_build_manifest_uses_msstore_source_and_dedupes_apps(self):
        created_at = datetime(2026, 6, 27, 12, 0, 0, tzinfo=timezone.utc)
        manifest = StoreAPI.build_winget_import_manifest(
            [
                {"Name": "Windows Terminal", "ProductId": "9N0DX20HK701"},
                {"Name": "Terminal Duplicate", "ProductId": "9N0DX20HK701"},
                {"Name": "Calculator", "ProductId": "9WZDNCRFHVN5"},
            ],
            winget_version="v1.28.240",
            created_at=created_at,
        )

        self.assertEqual(manifest["$schema"], "https://aka.ms/winget-packages.schema.2.0.json")
        self.assertEqual(manifest["CreationDate"], "2026-06-27T12:00:00.000-00:00")
        self.assertEqual(manifest["WinGetVersion"], "1.28.240")
        self.assertEqual(manifest["Sources"][0]["SourceDetails"]["Name"], "msstore")
        self.assertEqual(
            manifest["Sources"][0]["Packages"],
            [
                {"PackageIdentifier": "9N0DX20HK701"},
                {"PackageIdentifier": "9WZDNCRFHVN5"},
            ],
        )

    def test_build_manifest_rejects_apps_without_product_ids(self):
        with self.assertRaises(ValueError):
            StoreAPI.build_winget_import_manifest([{"Name": "No ID"}])

    def test_write_manifest_round_trips_json(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            manifest_path = os.path.join(temp_dir, "winget-import.json")

            path, count = StoreAPI.write_winget_import_manifest(
                [{"Name": "App Installer", "ProductId": "9NBLGGH4NNS1"}],
                manifest_path,
                winget_version="1.28.240",
                created_at=datetime(2026, 6, 27, tzinfo=timezone.utc),
            )

            self.assertEqual(path, manifest_path)
            self.assertEqual(count, 1)
            with open(manifest_path, encoding="utf-8") as handle:
                data = json.load(handle)
            self.assertEqual(data["Sources"][0]["Packages"][0]["PackageIdentifier"], "9NBLGGH4NNS1")


if __name__ == "__main__":
    unittest.main()
