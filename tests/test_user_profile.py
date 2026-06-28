#!/usr/bin/env python3

import os
import tempfile
import unittest

from MSStoreHelper import StoreAPI


class UserProfileTests(unittest.TestCase):
    def test_add_search_history_dedupes_and_limits_recent_queries(self):
        profile = StoreAPI.default_user_profile()
        for query in ["terminal", "calculator", "terminal", "photos"]:
            StoreAPI.add_search_history(profile, query, max_items=3)

        self.assertEqual(profile["SearchHistory"], ["photos", "terminal", "calculator"])

    def test_add_pinned_favorites_dedupes_by_product_id(self):
        profile = StoreAPI.default_user_profile()
        added = StoreAPI.add_pinned_favorites(profile, [
            {"Name": "Calculator", "ProductId": "9WZDNCRFHVN5", "Icon": "C"},
            {"Name": "Calculator Updated", "ProductId": "9WZDNCRFHVN5", "Icon": "C"},
            {"Name": "Terminal", "ProductId": "9N0DX20HK701", "Icon": "T"},
        ])

        self.assertEqual(added, 2)
        self.assertEqual(len(profile["PinnedFavorites"]), 2)
        self.assertEqual(profile["PinnedFavorites"][0]["ProductId"], "9N0DX20HK701")
        self.assertEqual(profile["PinnedFavorites"][1]["Name"], "Calculator Updated")

    def test_user_profile_round_trips_json(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            profile_path = os.path.join(temp_dir, "profile.json")
            profile = StoreAPI.default_user_profile()
            StoreAPI.add_search_history(profile, "terminal")
            StoreAPI.add_pinned_favorites(profile, [{"Name": "Terminal", "ProductId": "9N0DX20HK701"}])

            StoreAPI.save_user_profile(profile, profile_path)
            loaded = StoreAPI.load_user_profile(profile_path)

            self.assertEqual(loaded["SearchHistory"], ["terminal"])
            self.assertEqual(loaded["PinnedFavorites"][0]["ProductId"], "9N0DX20HK701")

    def test_user_profile_round_trips_store_query_settings(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            profile_path = os.path.join(temp_dir, "profile.json")
            profile = StoreAPI.default_user_profile()
            profile["StoreRing"] = "WIS"
            profile["StoreLanguage"] = "de-DE"
            profile["StoreMarket"] = "DE"

            StoreAPI.save_user_profile(profile, profile_path)
            loaded = StoreAPI.load_user_profile(profile_path)

            self.assertEqual(loaded["StoreRing"], "WIS")
            self.assertEqual(loaded["StoreLanguage"], "de-DE")
            self.assertEqual(loaded["StoreMarket"], "DE")

    def test_invalid_store_query_settings_fall_back_to_defaults(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            profile_path = os.path.join(temp_dir, "profile.json")
            with open(profile_path, "w", encoding="utf-8") as handle:
                handle.write('{"StoreRing":"Canary","StoreLanguage":"english","StoreMarket":"United States"}')

            loaded = StoreAPI.load_user_profile(profile_path)

            self.assertEqual(loaded["StoreRing"], "Retail")
            self.assertEqual(loaded["StoreLanguage"], "en-US")
            self.assertEqual(loaded["StoreMarket"], "US")


if __name__ == "__main__":
    unittest.main()
