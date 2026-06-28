#!/usr/bin/env python3

import unittest

from MSStoreHelper import Theme


class AccessibilityTests(unittest.TestCase):
    def test_theme_text_contrast_meets_wcag_aa_for_core_surfaces(self):
        pairs = [
            (Theme.TEXT_PRIMARY, Theme.BG_DARK),
            (Theme.TEXT_PRIMARY, Theme.BG_CARD),
            (Theme.TEXT_SECONDARY, Theme.BG_DARK),
            (Theme.TEXT_SECONDARY, Theme.BG_CARD),
            (Theme.TEXT_PRIMARY, Theme.BG_INPUT),
        ]

        for mode in ("Dark", "Light"):
            with self.subTest(mode=mode):
                for foreground, background in pairs:
                    ratio = Theme.contrast_ratio(
                        Theme.color_for_mode(foreground, mode),
                        Theme.color_for_mode(background, mode),
                    )
                    self.assertGreaterEqual(ratio, 4.5)

    def test_theme_muted_text_contrast_meets_large_text_threshold(self):
        for mode in ("Dark", "Light"):
            ratio = Theme.contrast_ratio(
                Theme.color_for_mode(Theme.TEXT_MUTED, mode),
                Theme.color_for_mode(Theme.BG_CARD, mode),
            )
            self.assertGreaterEqual(ratio, 3.0)

    def test_outline_action_text_contrast_meets_wcag_aa(self):
        for mode in ("Dark", "Light"):
            with self.subTest(mode=mode):
                for background in (Theme.BG_CARD, Theme.BG_CARD_HOVER):
                    ratio = Theme.contrast_ratio(
                        Theme.color_for_mode(Theme.PRIMARY_OUTLINE_TEXT, mode),
                        Theme.color_for_mode(background, mode),
                    )
                    self.assertGreaterEqual(ratio, 4.5)

    def test_contrast_ratio_is_symmetric(self):
        self.assertAlmostEqual(
            Theme.contrast_ratio("#ffffff", "#000000"),
            Theme.contrast_ratio("#000000", "#ffffff"),
        )


if __name__ == "__main__":
    unittest.main()
