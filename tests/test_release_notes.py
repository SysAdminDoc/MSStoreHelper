#!/usr/bin/env python3

import unittest

from MSStoreHelper import StoreAPI


class ReleaseNotesTests(unittest.TestCase):
    def test_parse_release_notes_from_heading_section(self):
        html = """
        <html><head><title>Contoso App</title></head>
        <body><h2>What's new</h2><div>Fixed LTSC install flow.</div><h2>Ratings</h2></body></html>
        """

        notes = StoreAPI.parse_release_notes_html("contoso", html, "https://apps.microsoft.com/detail/contoso")

        self.assertEqual(notes["Notes"], "Fixed LTSC install flow.")
        self.assertEqual(notes["Source"], "heading")

    def test_parse_release_notes_from_embedded_json_key(self):
        html = '<html><script>window.data = {"releaseNotes":"Added pinned package flow."}</script></html>'

        notes = StoreAPI.parse_release_notes_html("contoso", html, "https://apps.microsoft.com/detail/contoso")

        self.assertEqual(notes["Notes"], "Added pinned package flow.")
        self.assertEqual(notes["Source"], "releaseNotes")

    def test_parse_release_notes_falls_back_to_product_description(self):
        html = """
        <html><head><title>Fallback</title></head>
        <script type="application/ld+json">
        {"@type":"SoftwareApplication","name":"Fallback App","description":"Store description text."}
        </script></html>
        """

        notes = StoreAPI.parse_release_notes_html("contoso", html, "https://apps.microsoft.com/detail/contoso")

        self.assertEqual(notes["Title"], "Fallback App")
        self.assertEqual(notes["Notes"], "Store description text.")
        self.assertEqual(notes["Source"], "product-description")


if __name__ == "__main__":
    unittest.main()
