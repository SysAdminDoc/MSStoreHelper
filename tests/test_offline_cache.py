#!/usr/bin/env python3

import json
import os
import tempfile
import threading
import unittest
from unittest.mock import patch

from MSStoreHelper import StoreAPI, USER_AGENT
from test_trust_utils import inspect_as_trusted, mark_package_trusted


class FakeDownloadResponse:
    def __init__(
        self,
        chunks,
        content_length=None,
        error_at=None,
        status_code=200,
        content_range=None,
        etag=None,
        last_modified=None,
    ):
        self.chunks = chunks
        self.headers = {}
        if content_length is not None:
            self.headers["content-length"] = str(content_length)
        if content_range is not None:
            self.headers["content-range"] = content_range
        if etag is not None:
            self.headers["etag"] = etag
        if last_modified is not None:
            self.headers["last-modified"] = last_modified
        self.error_at = error_at
        self.status_code = status_code

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

            package = {
                "FileName": os.path.basename(source),
                "LocalPath": source,
            }
            mark_package_trusted(package, source)
            ok, message = StoreAPI.cache_downloaded_artifact(
                package,
                cache_dir,
            )

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

            package = {
                "FileName": os.path.basename(source),
                "LocalPath": source,
            }
            mark_package_trusted(package, source)
            ok, message = StoreAPI.cache_downloaded_artifact(
                package,
                cache_dir,
            )

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

            package = {
                "FileName": os.path.basename(source),
                "LocalPath": source,
            }
            mark_package_trusted(package, source)
            StoreAPI.write_artifact_manifest(package, destination, cache_dir)
            ok, message = StoreAPI.cache_downloaded_artifact(
                package,
                cache_dir,
            )

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

    def test_cache_history_keeps_last_two_versions_and_removes_older_artifact(self):
        with tempfile.TemporaryDirectory() as cache_dir:
            paths = []
            for version in ("1.0.0.0", "2.0.0.0", "3.0.0.0"):
                path = os.path.join(cache_dir, f"Contoso.App_{version}_x64__test.msixbundle")
                with open(path, "wb") as handle:
                    handle.write(version.encode("ascii"))
                package = {"FileName": os.path.basename(path)}
                mark_package_trusted(package, path)
                StoreAPI.write_artifact_manifest(package, path, cache_dir)
                paths.append(path)

            manifest = StoreAPI.load_cache_manifest(cache_dir)
            history = next(iter(manifest["History"].values()))

            self.assertEqual([item["AvailableVersion"] for item in history], ["3.0.0.0", "2.0.0.0"])
            dimensions = history[0]["CacheDimensions"]
            self.assertEqual(
                set(dimensions),
                {
                    "Identity",
                    "Architecture",
                    "PackageType",
                    "Version",
                    "Ring",
                    "Language",
                    "Market",
                    "Source",
                },
            )
            self.assertEqual(dimensions["Architecture"], "x64")
            self.assertEqual(dimensions["PackageType"], "MSIXBUNDLE")
            self.assertEqual(len(history[0]["CacheKey"]), 64)
            self.assertEqual(len(history[0]["CompatibilityKey"]), 64)
            self.assertNotIn(os.path.basename(paths[0]), manifest["Artifacts"])
            self.assertFalse(os.path.exists(paths[0]))
            self.assertTrue(os.path.exists(paths[1]))
            self.assertTrue(os.path.exists(paths[2]))

    def test_rollback_candidates_selects_newest_cached_version_below_current(self):
        with tempfile.TemporaryDirectory() as cache_dir:
            for version in ("1.0.0.0", "2.0.0.0", "3.0.0.0"):
                path = os.path.join(cache_dir, f"Contoso.App_{version}_x64__test.msixbundle")
                with open(path, "wb") as handle:
                    handle.write(version.encode("ascii"))
                package = {"FileName": os.path.basename(path)}
                mark_package_trusted(package, path)
                StoreAPI.write_artifact_manifest(package, path, cache_dir)

            candidates = StoreAPI.rollback_candidates(
                [cache_dir],
                ["Contoso.App"],
                {"contoso.app": "3.0.0.0"},
            )

            self.assertEqual(len(candidates), 1)
            self.assertEqual(candidates[0]["RollbackIdentity"], "contoso.app")
            self.assertEqual(candidates[0]["RollbackVersion"], "2.0.0.0")

    def test_rollback_candidates_use_second_newest_without_current_version(self):
        with tempfile.TemporaryDirectory() as cache_dir:
            for version in ("1.0.0.0", "2.0.0.0"):
                path = os.path.join(cache_dir, f"Contoso.App_{version}_x64__test.msixbundle")
                with open(path, "wb") as handle:
                    handle.write(version.encode("ascii"))
                package = {"FileName": os.path.basename(path)}
                mark_package_trusted(package, path)
                StoreAPI.write_artifact_manifest(package, path, cache_dir)

            candidates = StoreAPI.rollback_candidates([cache_dir], ["Contoso.App"], {})

            self.assertEqual(candidates[0]["RollbackVersion"], "1.0.0.0")

    def test_cache_selection_never_crosses_architecture_or_store_ring(self):
        with tempfile.TemporaryDirectory() as cache_dir:
            variants = (
                ("x64", "Retail", ("1.0.0.0", "2.0.0.0")),
                ("x86", "Retail", ("7.0.0.0", "8.0.0.0")),
                ("x64", "WIS", ("9.0.0.0", "10.0.0.0")),
            )
            for architecture, ring, versions in variants:
                for version in versions:
                    path = os.path.join(
                        cache_dir,
                        (
                            f"Contoso.App_{version}_{architecture}"
                            "__test.msix"
                        ),
                    )
                    with open(path, "wb") as handle:
                        handle.write(
                            f"{architecture}-{ring}-{version}".encode(
                                "ascii"
                            )
                        )
                    package = {
                        "FileName": os.path.basename(path),
                        "StoreQuery": {
                            "ProductId": "9CONTOSO",
                            "Ring": ring,
                            "Language": "en-US",
                            "Market": "US",
                        },
                    }
                    mark_package_trusted(package, path)
                    StoreAPI.write_artifact_manifest(
                        package,
                        path,
                        cache_dir,
                    )

            query = {
                "Ring": "Retail",
                "Language": "en-US",
                "Market": "US",
            }
            rollback = StoreAPI.rollback_candidates(
                [cache_dir],
                ["Contoso.App"],
                {"contoso.app": "3.0.0.0"},
                target_arch="x64",
                store_query=query,
            )
            diffs = StoreAPI.package_diff_candidates(
                [cache_dir],
                ["Contoso.App"],
                target_arch="x64",
                store_query=query,
            )

            self.assertEqual(len(rollback), 1)
            self.assertEqual(rollback[0]["RollbackVersion"], "2.0.0.0")
            self.assertEqual(
                rollback[0]["CacheDimensions"]["Architecture"],
                "x64",
            )
            self.assertEqual(
                rollback[0]["CacheDimensions"]["Ring"],
                "retail",
            )
            self.assertEqual(len(diffs), 1)
            self.assertEqual(diffs[0]["Old"]["AvailableVersion"], "1.0.0.0")
            self.assertEqual(diffs[0]["New"]["AvailableVersion"], "2.0.0.0")
            self.assertEqual(
                diffs[0]["Old"]["CompatibilityKey"],
                diffs[0]["New"]["CompatibilityKey"],
            )

    def test_concurrent_cache_updates_preserve_every_identity(self):
        with tempfile.TemporaryDirectory() as cache_dir:
            packages = []
            for index in range(8):
                path = os.path.join(
                    cache_dir,
                    f"Contoso.App{index}_1.0.0.0_x64__test.msix",
                )
                with open(path, "wb") as handle:
                    handle.write(f"package-{index}".encode("ascii"))
                package = {"FileName": os.path.basename(path)}
                mark_package_trusted(package, path)
                packages.append((package, path))

            failures = []

            def persist(package, path):
                try:
                    StoreAPI.write_artifact_manifest(
                        package,
                        path,
                        cache_dir,
                    )
                except Exception as exc:
                    failures.append(exc)

            threads = [
                threading.Thread(target=persist, args=item)
                for item in packages
            ]
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join()

            manifest = StoreAPI.load_cache_manifest(cache_dir)
            self.assertEqual(failures, [])
            self.assertEqual(len(manifest["Artifacts"]), len(packages))
            self.assertEqual(len(manifest["History"]), len(packages))

    def test_rollback_package_runs_remove_then_add(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            package_path = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msixbundle")
            with open(package_path, "wb") as handle:
                handle.write(b"package")
            package = {"FileName": os.path.basename(package_path)}
            mark_package_trusted(package, package_path)

            with patch("MSStoreHelper.run_command") as run_mock:
                run_mock.return_value.returncode = 0
                run_mock.return_value.stdout = "rollback ok"
                run_mock.return_value.stderr = ""

                ok, message = StoreAPI.rollback_package(
                    "Contoso.App",
                    package_path,
                    package,
                )

            self.assertTrue(ok, message)
            command = run_mock.call_args.args[0][-1]
            environment = run_mock.call_args.kwargs["env"]
            self.assertIn("Remove-AppxPackage", command)
            self.assertIn("Add-AppxPackage", command)
            self.assertNotIn(package_path, command)
            self.assertEqual(
                environment["MSSTOREHELPER_PACKAGE_PATH"],
                os.path.realpath(package_path),
            )
            self.assertEqual(
                environment["MSSTOREHELPER_ROLLBACK_IDENTITY"],
                "Contoso.App",
            )

    def test_download_file_writes_final_file_atomically_and_records_manifest(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            target = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msix")
            package = {"FileName": os.path.basename(target), "Url": "https://example.invalid/app.msix"}

            with patch("MSStoreHelper.requests.get", return_value=FakeDownloadResponse([b"pack", b"age"], content_length=7)):
                with patch.object(StoreAPI, "inspect_package_trust", side_effect=inspect_as_trusted):
                    ok, message = StoreAPI.download_file(package["Url"], target, package=package)

            self.assertTrue(ok, message)
            self.assertFalse(os.path.exists(f"{target}.part"))
            with open(target, "rb") as handle:
                self.assertEqual(handle.read(), b"package")
            self.assertEqual(package["Sha256"], StoreAPI.file_sha256(target))
            self.assertTrue(os.path.exists(os.path.join(temp_dir, "msstorehelper-cache-manifest.json")))

    def test_download_file_resumes_existing_part_with_range_request(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            target = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msix")
            captured = {}

            def interrupted_get(_url, **_kwargs):
                return FakeDownloadResponse(
                    [b"pack", b"age"],
                    content_length=7,
                    error_at=1,
                    etag='"representation-v1"',
                )

            with patch("MSStoreHelper.requests.get", side_effect=interrupted_get):
                ok, _message = StoreAPI.download_file(
                    "https://example.invalid/app.msix",
                    target,
                )
            self.assertFalse(ok)
            self.assertTrue(os.path.exists(f"{target}.part.json"))

            def resumed_get(_url, **kwargs):
                captured["headers"] = kwargs.get("headers")
                return FakeDownloadResponse(
                    [b"age"],
                    content_length=3,
                    status_code=206,
                    content_range="bytes 4-6/7",
                    etag='"representation-v1"',
                )

            with patch("MSStoreHelper.requests.get", side_effect=resumed_get):
                with patch.object(StoreAPI, "inspect_package_trust", side_effect=inspect_as_trusted):
                    ok, message = StoreAPI.download_file("https://example.invalid/app.msix", target)

            self.assertTrue(ok, message)
            self.assertEqual(
                captured["headers"],
                {
                    "User-Agent": USER_AGENT,
                    "Accept-Encoding": "identity",
                    "Range": "bytes=4-",
                    "If-Range": '"representation-v1"',
                },
            )
            self.assertFalse(os.path.exists(f"{target}.part"))
            self.assertFalse(os.path.exists(f"{target}.part.json"))
            with open(target, "rb") as handle:
                self.assertEqual(handle.read(), b"package")

    def test_download_file_reuses_existing_verified_file(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            target = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msix")
            with open(target, "wb") as handle:
                handle.write(b"package")
            package = {
                "FileName": os.path.basename(target),
                "SizeBytes": 7,
                "Sha256": StoreAPI.file_sha256(target),
            }
            mark_package_trusted(package, target)

            with patch("MSStoreHelper.requests.get") as get_mock:
                ok, message = StoreAPI.download_file("https://example.invalid/app.msix", target, package=package)

            self.assertTrue(ok, message)
            self.assertEqual(
                message,
                "Already downloaded; Package trust state: trusted",
            )
            get_mock.assert_not_called()

    def test_download_file_keeps_part_file_on_failure(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            target = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msix")

            with patch("MSStoreHelper.requests.get", return_value=FakeDownloadResponse([b"partial", b"tail"], content_length=11, error_at=1, etag='"partial-v1"')):
                ok, message = StoreAPI.download_file("https://example.invalid/app.msix", target)

            self.assertFalse(ok)
            self.assertIn("network dropped", message)
            self.assertFalse(os.path.exists(target))
            self.assertTrue(os.path.exists(f"{target}.part"))
            self.assertTrue(os.path.exists(f"{target}.part.json"))
            with open(f"{target}.part", "rb") as handle:
                self.assertEqual(handle.read(), b"partial")

    def test_download_state_round_trips_queue_without_widgets(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "download-state.json")
            package = {
                "FileName": "Contoso.App_1.0.0.0_x64__test.msix",
                "Url": "https://example.invalid/app.msix",
                "Architecture": "x64",
                "DownloadStatus": "Partial",
                "LastError": "network dropped",
                "_status_widget": object(),
            }

            StoreAPI.write_download_state([package], temp_dir, state_path)
            loaded = StoreAPI.load_download_state(state_path)
            with open(state_path, "r", encoding="utf-8") as handle:
                persisted = json.load(handle)

            self.assertEqual(persisted["SchemaVersion"], 2)
            self.assertEqual(persisted["Version"], 2)
            self.assertEqual(loaded["OutputPath"], os.path.abspath(temp_dir))
            self.assertEqual(loaded["Queue"][0]["FileName"], package["FileName"])
            self.assertEqual(loaded["Queue"][0]["DownloadStatus"], "Partial")
            self.assertNotIn("_status_widget", loaded["Queue"][0])

            StoreAPI.clear_download_state(state_path)
            self.assertFalse(os.path.exists(state_path))


if __name__ == "__main__":
    unittest.main()
