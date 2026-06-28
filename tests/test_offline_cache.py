#!/usr/bin/env python3

import json
import os
import tempfile
import unittest
from unittest.mock import patch

from MSStoreHelper import StoreAPI


class FakeDownloadResponse:
    def __init__(self, chunks, content_length=None, error_at=None):
        self.chunks = chunks
        self.headers = {}
        if content_length is not None:
            self.headers["content-length"] = str(content_length)
        self.error_at = error_at

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return False

    def raise_for_status(self):
        return None

    def iter_content(self, chunk_size=8192):
        for index, chunk in enumerate(self.chunks):
            if self.error_at == index:
                raise RuntimeError("network dropped")
            yield chunk


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

    def test_cache_downloaded_artifact_replaces_same_size_hash_mismatch(self):
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
            self.assertIn("Cached:", message)
            with open(destination, "rb") as handle:
                self.assertEqual(handle.read(), b"package")

            manifest_path = os.path.join(cache_dir, "msstorehelper-cache-manifest.json")
            with open(manifest_path, "r", encoding="utf-8") as handle:
                manifest = json.load(handle)
            self.assertIn(os.path.basename(source), manifest["Artifacts"])
            self.assertEqual(manifest["Artifacts"][os.path.basename(source)]["SizeBytes"], 7)

    def test_cache_downloaded_artifact_reuses_valid_manifest_copy(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msix")
            cache_dir = os.path.join(temp_dir, "cache")
            os.makedirs(cache_dir)
            destination = os.path.join(cache_dir, os.path.basename(source))

            with open(source, "wb") as handle:
                handle.write(b"package")
            with open(destination, "wb") as handle:
                handle.write(b"package")

            StoreAPI.write_artifact_manifest({"FileName": os.path.basename(source)}, destination, cache_dir)
            ok, message = StoreAPI.cache_downloaded_artifact({
                "FileName": os.path.basename(source),
                "LocalPath": source,
            }, cache_dir)

            self.assertTrue(ok)
            self.assertIn("Already cached:", message)

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

    def test_download_file_writes_final_file_atomically_and_records_manifest(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            target = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msix")
            package = {"FileName": os.path.basename(target), "Url": "https://example.invalid/app.msix"}

            with patch("MSStoreHelper.requests.get", return_value=FakeDownloadResponse([b"pack", b"age"], content_length=7)):
                ok, message = StoreAPI.download_file(package["Url"], target, package=package)

            self.assertTrue(ok, message)
            self.assertFalse(os.path.exists(f"{target}.part"))
            with open(target, "rb") as handle:
                self.assertEqual(handle.read(), b"package")
            self.assertEqual(package["Sha256"], StoreAPI.file_sha256(target))
            self.assertTrue(os.path.exists(os.path.join(temp_dir, "msstorehelper-cache-manifest.json")))

    def test_download_file_keeps_part_file_on_failure(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            target = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msix")

            with patch("MSStoreHelper.requests.get", return_value=FakeDownloadResponse([b"partial", b"tail"], content_length=11, error_at=1)):
                ok, message = StoreAPI.download_file("https://example.invalid/app.msix", target)

            self.assertFalse(ok)
            self.assertIn("network dropped", message)
            self.assertFalse(os.path.exists(target))
            self.assertTrue(os.path.exists(f"{target}.part"))
            with open(f"{target}.part", "rb") as handle:
                self.assertEqual(handle.read(), b"partial")


if __name__ == "__main__":
    unittest.main()
