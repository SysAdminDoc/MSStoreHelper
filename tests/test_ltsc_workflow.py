#!/usr/bin/env python3

import unittest

from MSStoreHelper import APP_CATALOG, LTSC_COMPONENT_REQUIREMENTS, QUICK_FIX_PRESETS, XBOX_CORE_PACKAGE_PINS, StoreAPI


class LtscWorkflowTests(unittest.TestCase):
    def test_ltsc_essentials_preset_contains_core_missing_apps(self):
        preset_name = next(name for name in QUICK_FIX_PRESETS if "LTSC Essentials" in name)

        self.assertEqual(
            QUICK_FIX_PRESETS[preset_name]["apps"],
            ["Windows Terminal", "PowerShell 7", "WSL", "Photos", "Calculator", "Snipping Tool"],
        )

    def test_ltsc_essentials_apps_exist_in_catalog(self):
        catalog_names = {
            app["Name"]
            for category in APP_CATALOG.values()
            for app in category["apps"]
        }
        preset_name = next(name for name in QUICK_FIX_PRESETS if "LTSC Essentials" in name)

        self.assertTrue(set(QUICK_FIX_PRESETS[preset_name]["apps"]).issubset(catalog_names))

    def test_ltsc_requirement_apps_exist_in_catalog(self):
        catalog_names = {
            app["Name"]
            for category in APP_CATALOG.values()
            for app in category["apps"]
        }

        self.assertTrue({item["Name"] for item in LTSC_COMPONENT_REQUIREMENTS}.issubset(catalog_names))

    def test_detect_missing_ltsc_components_skips_installed_identities(self):
        missing = StoreAPI.detect_missing_ltsc_components({
            "microsoft.windowscalculator",
            "microsoft.windows.photos",
            "microsoft.desktopappinstaller",
        })
        missing_names = {app["Name"] for app in missing}

        self.assertNotIn("Calculator", missing_names)
        self.assertNotIn("Photos", missing_names)
        self.assertNotIn("App Installer", missing_names)
        self.assertIn("Windows Terminal", missing_names)

    def test_detect_missing_ltsc_components_returns_empty_when_all_tracked_components_exist(self):
        installed = {
            identity.lower()
            for requirement in LTSC_COMPONENT_REQUIREMENTS
            for identity in requirement["Identities"]
        }

        self.assertEqual(StoreAPI.detect_missing_ltsc_components(installed), [])

    def test_xbox_core_pins_have_known_good_versions(self):
        self.assertEqual(
            {pin["Name"]: pin["KnownGoodVersions"] for pin in XBOX_CORE_PACKAGE_PINS},
            {
                "Xbox Identity": ["12.50.6001.0"],
                "Gaming Services": ["2.51.3002.0"],
            },
        )

    def test_select_pinned_xbox_packages_prefers_exact_known_good_versions(self):
        packages = [
            {"FileName": "Microsoft.GamingServices_99.0.0.0_x64__8wekyb3d8bbwe.msixbundle", "FileType": "MSIXBUNDLE", "Architecture": "x64"},
            {"FileName": "Microsoft.GamingServices_2.51.3002.0_x64__8wekyb3d8bbwe.msixbundle", "FileType": "MSIXBUNDLE", "Architecture": "x64"},
            {"FileName": "Microsoft.XboxIdentityProvider_12.50.6001.0_x64__8wekyb3d8bbwe.appxbundle", "FileType": "APPXBUNDLE", "Architecture": "x64"},
            {"FileName": "Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.appx", "FileType": "APPX", "Architecture": "x64"},
        ]

        selected = StoreAPI.select_pinned_xbox_packages(packages, "x64")
        filenames = [package["FileName"] for package in selected]

        self.assertLess(filenames.index("Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.appx"), filenames.index("Microsoft.GamingServices_2.51.3002.0_x64__8wekyb3d8bbwe.msixbundle"))
        self.assertIn("Microsoft.GamingServices_2.51.3002.0_x64__8wekyb3d8bbwe.msixbundle", filenames)
        self.assertNotIn("Microsoft.GamingServices_99.0.0.0_x64__8wekyb3d8bbwe.msixbundle", filenames)
        self.assertTrue(next(package for package in selected if package.get("XboxCoreName") == "Gaming Services")["PinnedVersionMatched"])

    def test_select_pinned_xbox_packages_falls_back_when_pin_is_unavailable(self):
        packages = [
            {"FileName": "Microsoft.GamingServices_3.0.0.0_x64__8wekyb3d8bbwe.msixbundle", "FileType": "MSIXBUNDLE", "Architecture": "x64"},
            {"FileName": "Microsoft.GamingServices_4.0.0.0_x64__8wekyb3d8bbwe.msixbundle", "FileType": "MSIXBUNDLE", "Architecture": "x64"},
        ]

        selected = StoreAPI.select_pinned_xbox_packages(packages, "x64")

        self.assertEqual(selected[0]["FileName"], "Microsoft.GamingServices_4.0.0.0_x64__8wekyb3d8bbwe.msixbundle")
        self.assertFalse(selected[0]["PinnedVersionMatched"])


if __name__ == "__main__":
    unittest.main()
