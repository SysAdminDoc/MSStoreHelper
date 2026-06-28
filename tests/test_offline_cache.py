#!/usr/bin/env python3

import os
import tempfile
import unittest

from MSStoreHelper import StoreAPI


class OfflineCacheTests(unittest.TestCase):
    def test_cache_downloaded_artifact_copies_installable_file(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msixbundle")
            cache_dir = os.path.join(temp_dir, "cache")
            with open(source, "wb") as handle:
                handle.write(b"package")

            ok, message = StoreAPI.cache_downloaded_artifact({
                "FileName": os.path.basename(source),
                "LocalPath": source,
            }, cache_dir)

            self.assertTrue(ok)
            self.assertIn("Cached:", message)
            with open(os.path.join(cache_dir, os.path.basename(source)), "rb") as handle:
                self.assertEqual(handle.read(), b"package")

    def test_cache_downloaded_artifact_reuses_same_size_existing_copy(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.appxbundle")
            cache_dir = os.path.join(temp_dir, "cache")
            os.makedirs(cache_dir)
            destination = os.path.join(cache_dir, os.path.basename(source))

            with open(source, "wb") as handle:
                handle.write(b"package")
            with open(destination, "wb") as handle:
                handle.write(b"cached!")

            ok, message = StoreAPI.cache_downloaded_artifact({
                "FileName": os.path.basename(source),
                "LocalPath": source,
            }, cache_dir)

            self.assertTrue(ok)
            self.assertIn("Already cached:", message)
            with open(destination, "rb") as handle:
                self.assertEqual(handle.read(), b"cached!")

    def test_cache_downloaded_artifact_rejects_non_installable_file(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = os.path.join(temp_dir, "Contoso.App.BlockMap")
            with open(source, "wb") as handle:
                handle.write(b"blockmap")

            ok, message = StoreAPI.cache_downloaded_artifact({
                "FileName": os.path.basename(source),
                "LocalPath": source,
            }, os.path.join(temp_dir, "cache"))

            self.assertFalse(ok)
            self.assertEqual(message, "File type is not cacheable")


if __name__ == "__main__":
    unittest.main()
