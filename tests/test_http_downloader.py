#!/usr/bin/env python3

import hashlib
import json
import os
import tempfile
import threading
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from MSStoreHelper import StoreAPI
from http_downloader import (
    DownloadCancelled,
    HttpDownloadError,
    download_http_file,
    validated_content_length,
)
from test_trust_utils import inspect_as_trusted


FILENAME = "Contoso.App_1.0.0.0_x64__test.msix"
URL = "https://cdn.test/package.msix?token=old"
SOURCE_ID = "product:9test|file:" + FILENAME.lower()


class FakeResponse:
    def __init__(
        self,
        chunks=(),
        *,
        status=200,
        content_length=None,
        content_range=None,
        etag=None,
        last_modified=None,
        error_at=None,
        url=URL,
        cancel_event=None,
    ):
        self.chunks = list(chunks)
        self.status_code = status
        self.url = url
        self.history = []
        self.error_at = error_at
        self.cancel_event = cancel_event
        self.headers = {}
        for key, value in (
            ("content-length", content_length),
            ("content-range", content_range),
            ("etag", etag),
            ("last-modified", last_modified),
        ):
            if value is not None:
                self.headers[key] = str(value)

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def iter_content(self, chunk_size=1024 * 1024):
        for index, chunk in enumerate(self.chunks):
            if self.error_at == index:
                raise RuntimeError("connection interrupted")
            yield chunk
            if self.cancel_event is not None:
                self.cancel_event.set()


def download(target, get, **kwargs):
    return download_http_file(
        URL,
        target,
        filename=FILENAME,
        source_identity=SOURCE_ID,
        get=get,
        free_space_reserve_bytes=0,
        **kwargs,
    )


def read_bytes(path):
    with open(path, "rb") as stream:
        return stream.read()


class HttpDownloaderTests(unittest.TestCase):
    def test_interruption_writes_hash_bound_state_and_exact_resume(self):
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            interrupted = FakeResponse(
                [b"pack", b"age"],
                content_length=7,
                etag='"v1"',
                error_at=1,
            )
            with self.assertRaisesRegex(RuntimeError, "interrupted"):
                download(target, lambda *_args, **_kwargs: interrupted)

            state_path = f"{target}.part.json"
            with open(state_path, "r", encoding="utf-8") as stream:
                state = json.load(stream)
            self.assertEqual(state["DownloadedBytes"], 4)
            self.assertEqual(state["ExpectedLength"], 7)
            self.assertEqual(
                state["PartialSha256"],
                hashlib.sha256(b"pack").hexdigest(),
            )
            self.assertNotIn("token=old", json.dumps(state))
            captured = {}

            def resumed_get(_url, **kwargs):
                captured.update(kwargs)
                return FakeResponse(
                    [b"age"],
                    status=206,
                    content_length=3,
                    content_range="bytes 4-6/7",
                    etag='"v1"',
                )

            evidence = download(target, resumed_get)

            self.assertEqual(
                captured["headers"],
                {
                    "Accept-Encoding": "identity",
                    "Range": "bytes=4-",
                    "If-Range": '"v1"',
                },
            )
            self.assertTrue(evidence["Resumed"])
            self.assertEqual(read_bytes(target), b"package")
            self.assertFalse(os.path.exists(state_path))

    def test_legacy_partial_without_state_is_discarded(self):
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            with open(f"{target}.part", "wb") as stream:
                stream.write(b"unbound")
            captured = {}

            def get(_url, **kwargs):
                captured.update(kwargs)
                return FakeResponse([b"fresh"], content_length=5)

            download(target, get)

            self.assertEqual(
                captured["headers"],
                {"Accept-Encoding": "identity"},
            )
            self.assertEqual(read_bytes(target), b"fresh")

    def test_http_200_fallback_replaces_saved_representation(self):
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            interrupted = FakeResponse(
                [b"old", b"-tail"],
                content_length=8,
                etag='"old"',
                error_at=1,
            )
            with self.assertRaises(RuntimeError):
                download(target, lambda *_args, **_kwargs: interrupted)

            replacement = b"replacement"
            evidence = download(
                target,
                lambda *_args, **_kwargs: FakeResponse(
                    [replacement],
                    content_length=len(replacement),
                    etag='"new"',
                ),
            )

            self.assertFalse(evidence["Resumed"])
            self.assertEqual(read_bytes(target), replacement)

    def test_malformed_content_range_is_discarded_and_retried(self):
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            interrupted = FakeResponse(
                [b"part", b"tail"],
                content_length=8,
                etag='"v1"',
                error_at=1,
            )
            with self.assertRaises(RuntimeError):
                download(target, lambda *_args, **_kwargs: interrupted)

            responses = iter([
                FakeResponse(
                    [b"tail"],
                    status=206,
                    content_length=4,
                    content_range="bytes 3-6/8",
                    etag='"v1"',
                ),
                FakeResponse([b"clean"], content_length=5, etag='"v2"'),
            ])
            calls = []

            def get(_url, **kwargs):
                calls.append(kwargs.get("headers"))
                return next(responses)

            download(target, get)

            self.assertEqual(len(calls), 2)
            self.assertIsNotNone(calls[0])
            self.assertEqual(
                calls[1],
                {"Accept-Encoding": "identity"},
            )
            self.assertEqual(read_bytes(target), b"clean")

    def test_exact_416_promotes_verified_complete_partial(self):
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            data = b"complete"
            part_path = f"{target}.part"
            with open(part_path, "wb") as stream:
                stream.write(data)
            with open(f"{target}.part.json", "w", encoding="utf-8") as stream:
                json.dump({
                    "SchemaVersion": 1,
                    "SourceIdentity": SOURCE_ID,
                    "SourceUrl": URL,
                    "EffectiveUrl": URL,
                    "ETag": '"v1"',
                    "LastModified": "",
                    "ExpectedLength": len(data),
                    "DownloadedBytes": len(data),
                    "HashAlgorithm": "sha256",
                    "PartialSha256": hashlib.sha256(data).hexdigest(),
                    "UpdatedAt": "2026-07-29T00:00:00+00:00",
                }, stream)

            evidence = download(
                target,
                lambda *_args, **_kwargs: FakeResponse(
                    status=416,
                    content_range=f"bytes */{len(data)}",
                ),
            )

            self.assertTrue(evidence["Resumed"])
            self.assertEqual(read_bytes(target), data)

    def test_mismatched_416_discards_partial_and_restarts(self):
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            interrupted = FakeResponse(
                [b"part", b"tail"],
                content_length=8,
                etag='"v1"',
                error_at=1,
            )
            with self.assertRaises(RuntimeError):
                download(target, lambda *_args, **_kwargs: interrupted)
            responses = iter([
                FakeResponse(status=416, content_range="bytes */99"),
                FakeResponse([b"fresh"], content_length=5, etag='"v2"'),
            ])

            evidence = download(
                target,
                lambda *_args, **_kwargs: next(responses),
            )

            self.assertFalse(evidence["Resumed"])
            self.assertEqual(read_bytes(target), b"fresh")

    def test_zero_oversized_and_disk_reserve_responses_fail(self):
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            with self.assertRaisesRegex(HttpDownloadError, "zero-byte"):
                download(
                    target,
                    lambda *_args, **_kwargs: FakeResponse(
                        content_length=0,
                    ),
                )
            with self.assertRaisesRegex(HttpDownloadError, "exceeds"):
                download(
                    target,
                    lambda *_args, **_kwargs: FakeResponse(
                        [b"12345"],
                        content_length=5,
                    ),
                    max_bytes=4,
                )
            with patch(
                "http_downloader.shutil.disk_usage",
                return_value=SimpleNamespace(free=5),
            ):
                with self.assertRaisesRegex(HttpDownloadError, "needs"):
                    download_http_file(
                        URL,
                        target,
                        filename=FILENAME,
                        source_identity=SOURCE_ID,
                        get=lambda *_args, **_kwargs: FakeResponse(
                            [b"12"],
                            content_length=2,
                        ),
                        max_bytes=10,
                        free_space_reserve_bytes=4,
                    )

    def test_cancellation_discards_unbound_partial(self):
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            event = threading.Event()
            response = FakeResponse(
                [b"part", b"tail"],
                content_length=8,
                cancel_event=event,
            )
            with self.assertRaises(DownloadCancelled):
                download(
                    target,
                    lambda *_args, **_kwargs: response,
                    cancel_event=event,
                )
            self.assertFalse(os.path.exists(f"{target}.part"))
            self.assertFalse(os.path.exists(f"{target}.part.json"))

    def test_cancellation_preserves_validator_bound_partial(self):
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            event = threading.Event()
            response = FakeResponse(
                [b"part", b"tail"],
                content_length=8,
                etag='"v1"',
                cancel_event=event,
            )
            with self.assertRaises(DownloadCancelled):
                download(
                    target,
                    lambda *_args, **_kwargs: response,
                    cancel_event=event,
                )
            self.assertEqual(read_bytes(f"{target}.part"), b"part")
            with open(
                f"{target}.part.json",
                "r",
                encoding="utf-8",
            ) as stream:
                state = json.load(stream)
            self.assertEqual(state["ETag"], '"v1"')
            self.assertEqual(state["DownloadedBytes"], 4)

    def test_head_length_and_store_url_refresh_are_bounded(self):
        head = FakeResponse(content_length=7)
        self.assertEqual(
            validated_content_length(URL, head, max_bytes=7),
            7,
        )
        with self.assertRaises(HttpDownloadError):
            validated_content_length(
                URL,
                FakeResponse(content_length=8),
                max_bytes=7,
            )

        package = {
            "FileName": FILENAME,
            "Url": URL,
            "StoreQuery": {"ProductId": "9TEST"},
        }
        stale = FakeResponse(status=403)
        fresh_url = "https://cdn.test/package.msix?token=fresh"
        fresh = FakeResponse(
            [b"package"],
            content_length=7,
            etag='"fresh"',
            url=fresh_url,
        )
        with tempfile.TemporaryDirectory() as folder:
            target = os.path.join(folder, FILENAME)
            with patch(
                "MSStoreHelper.requests.get",
                side_effect=[stale, fresh],
            ) as get_mock:
                with patch.object(
                    StoreAPI,
                    "refresh_package_url",
                    return_value=fresh_url,
                ):
                    with patch.object(
                        StoreAPI,
                        "inspect_package_trust",
                        side_effect=inspect_as_trusted,
                    ):
                        ok, message = StoreAPI.download_file(
                            URL,
                            target,
                            package=package,
                            free_space_reserve_bytes=0,
                        )

        self.assertTrue(ok, message)
        self.assertEqual(get_mock.call_args_list[1].args[0], fresh_url)
        self.assertEqual(package["Url"], fresh_url)

    def test_refresh_package_url_uses_persisted_store_query(self):
        package = {
            "FileName": FILENAME,
            "Url": URL,
            "StoreQuery": {
                "ProductId": "9TEST",
                "Ring": "WIS",
                "Language": "de-DE",
                "Market": "DE",
            },
        }
        refreshed_url = "https://cdn.test/refreshed.msix?token=new"
        diagnostic = {
            "Query": package["StoreQuery"].copy(),
            "Packages": [{
                "FileName": FILENAME,
                "Url": refreshed_url,
            }],
        }
        with patch.object(
            StoreAPI,
            "get_packages_with_diagnostics",
            return_value=diagnostic,
        ) as lookup:
            result = StoreAPI.refresh_package_url(package)

        self.assertEqual(result, refreshed_url)
        lookup.assert_called_once_with("9TEST", "WIS", "de-DE", "DE")
        self.assertEqual(package["Url"], refreshed_url)
        self.assertTrue(package["UrlRefreshedAt"].startswith("20"))


if __name__ == "__main__":
    unittest.main()
