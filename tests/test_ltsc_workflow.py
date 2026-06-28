#!/usr/bin/env python3

import unittest

from MSStoreHelper import APP_CATALOG, LTSC_COMPONENT_REQUIREMENTS, QUICK_FIX_PRESETS, StoreAPI


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


if __name__ == "__main__":
    unittest.main()
