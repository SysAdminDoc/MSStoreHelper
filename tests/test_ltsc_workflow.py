#!/usr/bin/env python3

import unittest

from MSStoreHelper import APP_CATALOG, QUICK_FIX_PRESETS


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


if __name__ == "__main__":
    unittest.main()
