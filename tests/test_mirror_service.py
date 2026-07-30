#!/usr/bin/env python3

import concurrent.futures
import io
import json
import os
import tempfile
import threading
import time
import unittest
import urllib.error
import urllib.request
from pathlib import Path

from MSStoreHelper import StoreAPI, run_cli
from mirror_service import (
    MIRROR_AUDIT_FILENAME,
    MirrorConfigurationError,
    validate_network_policy,
)
from test_trust_utils import mark_package_trusted


PACKAGE_FILENAME = (
    "Microsoft.WindowsTerminal_1.0.0.0_"
    "x64__8wekyb3d8bbwe.msixbundle"
)
PACKAGE_BYTES = b"verified-package-payload"


class MirrorServiceTests(unittest.TestCase):
    def _trusted_cache(self, folder):
        path = Path(folder, PACKAGE_FILENAME)
        path.write_bytes(PACKAGE_BYTES)
        package = {
            "FileName": PACKAGE_FILENAME,
            "Url": (
                "https://example.test/package"
                "?sig=secret&token=hidden"
            ),
        }
        mark_package_trusted(package, path)
        StoreAPI.write_artifact_manifest(package, path, folder)
        return path

    def _start_server(self, folder):
        server, index = StoreAPI.create_mirror_server(
            folder,
            "127.0.0.1",
            0,
        )
        thread = threading.Thread(
            target=server.serve_forever,
            daemon=True,
        )
        thread.start()
        return server, index, thread

    def _stop_server(self, server, thread):
        server.shutdown()
        server.server_close()
        thread.join(timeout=5)

    def _request(self, url, *, method="GET", headers=None):
        request = urllib.request.Request(
            url,
            method=method,
            headers=headers or {},
        )
        return urllib.request.urlopen(request, timeout=5)

    def test_network_policy_requires_explicit_safe_lan_mode(self):
        with self.assertRaisesRegex(
            MirrorConfigurationError,
            "explicit LAN mode",
        ):
            validate_network_policy("0.0.0.0")
        with self.assertRaisesRegex(
            MirrorConfigurationError,
            "cleartext-risk",
        ):
            validate_network_policy(
                "0.0.0.0",
                advertised_host="192.0.2.10",
                lan_mode=True,
            )
        with self.assertRaisesRegex(
            MirrorConfigurationError,
            "advertised host",
        ):
            validate_network_policy(
                "0.0.0.0",
                lan_mode=True,
                acknowledge_cleartext=True,
            )

        policy = validate_network_policy(
            "0.0.0.0",
            advertised_host="192.0.2.10",
            lan_mode=True,
            acknowledge_cleartext=True,
        )

        self.assertEqual(policy["BindHost"], "0.0.0.0")
        self.assertEqual(policy["AdvertisedHost"], "192.0.2.10")
        self.assertTrue(policy["LanMode"])
        self.assertFalse(policy["TlsEnabled"])

    def test_server_allowlists_index_and_verified_package_routes(self):
        with tempfile.TemporaryDirectory() as folder:
            package_path = self._trusted_cache(folder)
            Path(folder, "notes-secret.txt").write_text(
                "do not expose",
                encoding="utf-8",
            )
            server, _index, thread = self._start_server(folder)
            port = server.server_address[1]
            base = f"http://127.0.0.1:{port}"
            try:
                with self._request(
                    f"{base}/msstorehelper-mirror-index.json"
                ) as response:
                    index_bytes = response.read()
                    index = json.loads(index_bytes.decode("utf-8"))
                with self._request(
                    f"{base}/packages/{PACKAGE_FILENAME}"
                ) as response:
                    package_bytes = response.read()

                blocked_routes = [
                    "/notes-secret.txt",
                    "/msstorehelper-cache-manifest.json",
                    f"/{PACKAGE_FILENAME}",
                    "/packages/%2e%2e/notes-secret.txt",
                    (
                        f"/packages/{PACKAGE_FILENAME}"
                        "?authorization=secret"
                    ),
                ]
                for route in blocked_routes:
                    with self.subTest(route=route):
                        with self.assertRaises(urllib.error.HTTPError) as error:
                            self._request(f"{base}{route}")
                        self.assertEqual(error.exception.code, 404)
            finally:
                self._stop_server(server, thread)

            serialized = index_bytes.decode("utf-8")
            manifest_name = os.path.join(
                folder,
                "msstorehelper-cache-manifest.json",
            )

        self.assertEqual(package_bytes, PACKAGE_BYTES)
        self.assertEqual(index["PackageCount"], 1)
        self.assertNotIn(folder, serialized)
        self.assertNotIn(manifest_name, serialized)
        self.assertNotIn("sig=secret", serialized)
        self.assertNotIn("token=hidden", serialized)
        self.assertEqual(
            index["Packages"][0]["Url"],
            f"{base}/packages/{PACKAGE_FILENAME}",
        )
        self.assertTrue(package_path.name.endswith(".msixbundle"))

    def test_head_and_single_range_are_bounded(self):
        with tempfile.TemporaryDirectory() as folder:
            self._trusted_cache(folder)
            server, _index, thread = self._start_server(folder)
            port = server.server_address[1]
            url = (
                f"http://127.0.0.1:{port}/packages/"
                f"{PACKAGE_FILENAME}"
            )
            try:
                with self._request(url, method="HEAD") as response:
                    head_body = response.read()
                    head_length = int(response.headers["Content-Length"])
                    accept_ranges = response.headers["Accept-Ranges"]
                with self._request(
                    url,
                    headers={"Range": "bytes=2-7"},
                ) as response:
                    range_body = response.read()
                    content_range = response.headers["Content-Range"]
                    status = response.status
                for invalid in ("bytes=999-", "bytes=0-1,3-4"):
                    with self.subTest(value=invalid):
                        with self.assertRaises(
                            urllib.error.HTTPError
                        ) as error:
                            self._request(
                                url,
                                headers={"Range": invalid},
                            )
                        self.assertEqual(error.exception.code, 416)
            finally:
                self._stop_server(server, thread)

        self.assertEqual(head_body, b"")
        self.assertEqual(head_length, len(PACKAGE_BYTES))
        self.assertEqual(accept_ranges, "bytes")
        self.assertEqual(status, 206)
        self.assertEqual(content_range, f"bytes 2-7/{len(PACKAGE_BYTES)}")
        self.assertEqual(range_body, PACKAGE_BYTES[2:8])

    def test_package_tamper_after_start_is_not_served(self):
        with tempfile.TemporaryDirectory() as folder:
            package_path = self._trusted_cache(folder)
            server, _index, thread = self._start_server(folder)
            port = server.server_address[1]
            package_path.write_bytes(b"tampered")
            try:
                with self.assertRaises(urllib.error.HTTPError) as error:
                    self._request(
                        f"http://127.0.0.1:{port}/packages/"
                        f"{PACKAGE_FILENAME}"
                    )
            finally:
                self._stop_server(server, thread)

        self.assertEqual(error.exception.code, 404)

    def test_bearer_auth_is_header_only_and_audit_is_redacted(self):
        with tempfile.TemporaryDirectory() as folder:
            self._trusted_cache(folder)
            token = "test-token-" + ("x" * 32)
            expires = time.time() + 120
            handler = StoreAPI.mirror_http_handler(
                folder,
                "127.0.0.1",
                8765,
                bearer_token=token,
                token_expires_at=expires,
            )
            from http.server import ThreadingHTTPServer

            server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
            server.daemon_threads = False
            server.block_on_close = True
            port = server.server_address[1]
            thread = threading.Thread(
                target=server.serve_forever,
                daemon=True,
            )
            thread.start()
            url = (
                f"http://127.0.0.1:{port}/packages/"
                f"{PACKAGE_FILENAME}"
            )
            try:
                with self.assertRaises(urllib.error.HTTPError) as missing:
                    self._request(url)
                with self.assertRaises(urllib.error.HTTPError) as rejected:
                    self._request(
                        url,
                        headers={"Authorization": "Bearer wrong"},
                    )
                with self._request(
                    url,
                    headers={"Authorization": f"Bearer {token}"},
                ) as response:
                    body = response.read()
                with self._request(
                    f"http://127.0.0.1:{port}/"
                    "msstorehelper-mirror-index.json",
                    headers={"Authorization": f"Bearer {token}"},
                ) as response:
                    authenticated_index = response.read().decode("utf-8")
            finally:
                self._stop_server(server, thread)

            audit_path = Path(folder, MIRROR_AUDIT_FILENAME)
            audit_text = audit_path.read_text(encoding="utf-8")
            records = [
                json.loads(line)
                for line in audit_text.splitlines()
            ]

        self.assertEqual(missing.exception.code, 401)
        self.assertEqual(rejected.exception.code, 401)
        self.assertEqual(body, PACKAGE_BYTES)
        self.assertNotIn(token, audit_text)
        self.assertNotIn(token, authenticated_index)
        self.assertNotIn("wrong", audit_text)
        self.assertCountEqual(
            [record["Authorization"] for record in records],
            ["missing", "rejected", "accepted", "accepted"],
        )
        self.assertTrue(
            all(record["ClientNetwork"] == "loopback" for record in records)
        )

    def test_explicit_lan_server_generates_short_lived_header_token(self):
        with tempfile.TemporaryDirectory() as folder:
            self._trusted_cache(folder)
            server, index = StoreAPI.create_mirror_server(
                folder,
                "0.0.0.0",
                0,
                advertised_host="127.0.0.1",
                lan_mode=True,
                acknowledge_cleartext=True,
                token_ttl_seconds=1,
            )
            port = server.server_address[1]
            token = server.mirror_bearer_token
            thread = threading.Thread(
                target=server.serve_forever,
                daemon=True,
            )
            thread.start()
            url = (
                f"http://127.0.0.1:{port}/packages/"
                f"{PACKAGE_FILENAME}"
            )
            try:
                with self.assertRaises(urllib.error.HTTPError) as missing:
                    self._request(url)
                with self._request(
                    url,
                    headers={"Authorization": f"Bearer {token}"},
                ) as response:
                    body = response.read()
            finally:
                self._stop_server(server, thread)
            serialized = json.dumps(index)

        self.assertEqual(missing.exception.code, 401)
        self.assertEqual(body, PACKAGE_BYTES)
        self.assertGreaterEqual(len(token), 40)
        self.assertNotIn(token, serialized)
        self.assertTrue(index["Authorization"]["Required"])
        self.assertTrue(index["Authorization"]["ExpiresAt"])

    def test_concurrent_index_writes_leave_one_valid_generation(self):
        with tempfile.TemporaryDirectory() as folder:
            self._trusted_cache(folder)

            def write_index(port):
                return StoreAPI.write_mirror_index(
                    folder,
                    "127.0.0.1",
                    port,
                )

            with concurrent.futures.ThreadPoolExecutor(
                max_workers=8
            ) as executor:
                list(executor.map(write_index, range(8700, 8740)))

            index_path = Path(
                folder,
                "msstorehelper-mirror-index.json",
            )
            payload = json.loads(index_path.read_text(encoding="utf-8"))
            temporary_files = list(
                Path(folder).glob(
                    ".msstorehelper-mirror-index.json.*.tmp"
                )
            )

        self.assertEqual(payload["SchemaVersion"], 2)
        self.assertEqual(payload["PackageCount"], 1)
        self.assertEqual(temporary_files, [])

    def test_cli_rejects_implicit_or_unacknowledged_lan_exposure(self):
        with tempfile.TemporaryDirectory() as folder:
            stdout = io.StringIO()
            stderr = io.StringIO()
            implicit = run_cli(
                ["--mirror", folder, "--host", "0.0.0.0"],
                stdout,
                stderr,
            )
            explicit_without_risk = run_cli(
                [
                    "--mirror",
                    folder,
                    "--host",
                    "0.0.0.0",
                    "--advertise-host",
                    "192.0.2.10",
                    "--lan",
                ],
                io.StringIO(),
                io.StringIO(),
            )

        self.assertEqual(implicit, 2)
        self.assertIn("explicit LAN mode", stderr.getvalue())
        self.assertEqual(explicit_without_risk, 2)


if __name__ == "__main__":
    unittest.main()
