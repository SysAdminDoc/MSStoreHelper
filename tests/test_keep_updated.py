#!/usr/bin/env python3

import unittest

from MSStoreHelper import StoreAPI


def package(filename, arch="x64"):
    return {
        "FileName": filename,
        "Architecture": arch,
        "FileType": filename.rsplit(".", 1)[-1].upper(),
        "IsBundle": filename.lower().endswith(("appxbundle", "msixbundle")),
        "IsEncrypted": False,
        "Url": f"https://example.test/{filename}",
    }


class KeepUpdatedTests(unittest.TestCase):
    def test_normalize_installed_appx_versions_keeps_newest_duplicate(self):
        versions = StoreAPI.normalize_installed_appx_versions([
            {"Name": "Microsoft.WindowsTerminal", "Version": "1.0.0.0"},
            {"Name": "microsoft.windowsterminal", "Version": "2.0.0.0"},
            {"DisplayName": "Microsoft.WindowsCalculator", "Version": "11.2405.0.0"},
        ])

        self.assertEqual(versions["microsoft.windowsterminal"], "2.0.0.0")
        self.assertEqual(versions["microsoft.windowscalculator"], "11.2405.0.0")

    def test_select_catalog_update_packages_queues_newer_app_and_needed_dependency(self):
        catalog = {
            "Productivity": {
                "apps": [
                    {"Name": "Windows Terminal", "ProductId": "9N0DX20HK701"},
                ],
            },
        }
        available = [
            package("Microsoft.WindowsTerminal_2.0.0.0_x64__8wekyb3d8bbwe.msixbundle"),
            package("Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64__8wekyb3d8bbwe.appx"),
        ]
        installed = {
            "microsoft.windowsterminal": "1.0.0.0",
            "microsoft.vclibs.140.00.uwpdesktop": "14.0.33500.0",
        }

        updates = StoreAPI.select_catalog_update_packages(
            catalog,
            installed,
            lambda app: available,
            "x64",
        )

        self.assertEqual(
            [item["FileName"] for item in updates],
            [
                "Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64__8wekyb3d8bbwe.appx",
                "Microsoft.WindowsTerminal_2.0.0.0_x64__8wekyb3d8bbwe.msixbundle",
            ],
        )
        self.assertEqual(updates[-1]["UpdateSourceApp"], "Windows Terminal")
        self.assertEqual(updates[-1]["UpdateInstalledIdentity"], "microsoft.windowsterminal")
        self.assertEqual(updates[-1]["UpdateInstalledVersion"], "1.0.0.0")
        self.assertEqual(updates[-1]["UpdateAvailableVersion"], "2.0.0.0")

    def test_select_catalog_update_packages_skips_current_and_uninstalled_apps(self):
        catalog = {
            "Productivity": {
                "apps": [
                    {"Name": "Windows Terminal", "ProductId": "9N0DX20HK701"},
                ],
            },
        }
        available = [
            package("Microsoft.WindowsTerminal_2.0.0.0_x64__8wekyb3d8bbwe.msixbundle"),
        ]

        current = StoreAPI.select_catalog_update_packages(
            catalog,
            {"microsoft.windowsterminal": "2.0.0.0"},
            lambda app: available,
            "x64",
        )
        missing = StoreAPI.select_catalog_update_packages(catalog, {}, lambda app: available, "x64")

        self.assertEqual(current, [])
        self.assertEqual(missing, [])


if __name__ == "__main__":
    unittest.main()
