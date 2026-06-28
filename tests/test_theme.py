#!/usr/bin/env python3

import json
import os
import tempfile
import unittest

from MSStoreHelper import StoreAPI, Theme


class ThemeTests(unittest.TestCase):
    def test_resolve_mode_honors_system_light_setting(self):
        self.assertEqual(Theme.resolve_mode("System", apps_use_light=True), "Light")
        self.assertEqual(Theme.resolve_mode("System", apps_use_light=False), "Dark")
        self.assertEqual(Theme.resolve_mode("Light", apps_use_light=False), "Light")
        self.assertEqual(Theme.resolve_mode("Dark", apps_use_light=True), "Dark")

    def test_normalize_mode_rejects_unknown_values(self):
        self.assertEqual(Theme.normalize_mode("light"), "Light")
        self.assertEqual(Theme.normalize_mode("unknown"), "System")
        self.assertEqual(Theme.normalize_mode(None), "System")

    def test_accent_from_windows_dword_decodes_abgr(self):
        self.assertEqual(Theme.accent_from_windows_dword(0x00D77800), "#0078d7")

    def test_shift_hex_color_lightens_and_darkens(self):
        self.assertEqual(Theme.shift_hex_color("#000000", 0.5), "#808080")
        self.assertEqual(Theme.shift_hex_color("#808080", -0.5), "#404040")

    def test_profile_round_trips_theme_mode(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            profile_path = os.path.join(temp_dir, "profile.json")
            profile = StoreAPI.default_user_profile()
            profile["ThemeMode"] = "Light"

            StoreAPI.save_user_profile(profile, profile_path)
            loaded = StoreAPI.load_user_profile(profile_path)

            self.assertEqual(loaded["ThemeMode"], "Light")

    def test_profile_invalid_theme_mode_falls_back_to_system(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            profile_path = os.path.join(temp_dir, "profile.json")
            with open(profile_path, "w", encoding="utf-8") as handle:
                json.dump({"ThemeMode": "sepia"}, handle)

            loaded = StoreAPI.load_user_profile(profile_path)

            self.assertEqual(loaded["ThemeMode"], "System")


if __name__ == "__main__":
    unittest.main()
