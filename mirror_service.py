#!/usr/bin/env python3
"""Allowlisted HTTP service primitives for trusted Store packages."""

import hashlib
import hmac
import ipaddress
import json
import os
import re
import secrets
import ssl
import threading
import time
import uuid
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlsplit


MIRROR_AUDIT_FILENAME = "msstorehelper-mirror-access.jsonl"
MIRROR_AUDIT_MAX_BYTES = 1024 * 1024
MIRROR_AUDIT_RETENTION = 3
DEFAULT_TOKEN_TTL_SECONDS = 15 * 60
MIN_TOKEN_TTL_SECONDS = 60
MAX_TOKEN_TTL_SECONDS = 60 * 60
_INDEX_WRITE_LOCK = threading.Lock()


class MirrorConfigurationError(ValueError):
    """Raised when mirror exposure would violate the network policy."""


def utc_timestamp(value=None):
    value = value or datetime.now(timezone.utc)
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).isoformat()


def is_loopback_host(host):
    host = str(host or "").strip().strip("[]")
    if host.lower() == "localhost":
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def is_wildcard_host(host):
    return str(host or "").strip().strip("[]") in {
        "",
        "0.0.0.0",
        "::",
    }


def normalize_advertised_host(host):
    host = str(host or "").strip()
    if not host:
        raise MirrorConfigurationError(
            "An advertised mirror host is required"
        )
    unwrapped = host.strip("[]")
    try:
        ipaddress.ip_address(unwrapped)
    except ValueError:
        pass
    else:
        if is_wildcard_host(unwrapped):
            raise MirrorConfigurationError(
                "A wildcard bind address cannot be advertised to clients"
            )
        return unwrapped
    parsed = urlsplit(f"//{host}")
    if (
        parsed.username
        or parsed.password
        or parsed.port is not None
        or parsed.path
        or parsed.query
        or parsed.fragment
    ):
        raise MirrorConfigurationError(
            "Advertised host must not contain a scheme, port, path, or secret"
        )
    normalized = parsed.hostname or ""
    if not re.fullmatch(r"[A-Za-z0-9_.:-]+", normalized):
        raise MirrorConfigurationError("Advertised mirror host is invalid")
    if is_wildcard_host(normalized):
        raise MirrorConfigurationError(
            "A wildcard bind address cannot be advertised to clients"
        )
    return normalized


def mirror_base_url(advertised_host, port, *, tls_enabled):
    host = normalize_advertised_host(advertised_host)
    try:
        ip = ipaddress.ip_address(host)
    except ValueError:
        rendered_host = host
    else:
        rendered_host = f"[{host}]" if ip.version == 6 else host
    scheme = "https" if tls_enabled else "http"
    return f"{scheme}://{rendered_host}:{int(port)}"


def normalize_token_ttl(value):
    try:
        value = int(value)
    except (TypeError, ValueError):
        value = DEFAULT_TOKEN_TTL_SECONDS
    return max(
        MIN_TOKEN_TTL_SECONDS,
        min(MAX_TOKEN_TTL_SECONDS, value),
    )


def validate_network_policy(
    bind_host,
    *,
    advertised_host=None,
    lan_mode=False,
    acknowledge_cleartext=False,
    tls_cert=None,
    tls_key=None,
):
    bind_host = str(bind_host or "").strip()
    if not bind_host:
        raise MirrorConfigurationError("Mirror bind host is required")
    loopback = is_loopback_host(bind_host)
    tls_cert = str(tls_cert or "").strip()
    tls_key = str(tls_key or "").strip()
    if bool(tls_cert) != bool(tls_key):
        raise MirrorConfigurationError(
            "Both TLS certificate and key are required"
        )
    tls_enabled = bool(tls_cert)
    if tls_enabled:
        if not os.path.isfile(tls_cert) or not os.path.isfile(tls_key):
            raise MirrorConfigurationError(
                "TLS certificate or key does not exist"
            )

    if not loopback:
        if not lan_mode:
            raise MirrorConfigurationError(
                "Non-loopback binding requires explicit LAN mode"
            )
        if not tls_enabled and not acknowledge_cleartext:
            raise MirrorConfigurationError(
                "LAN HTTP requires explicit cleartext-risk acknowledgement"
            )

    if advertised_host:
        advertised = normalize_advertised_host(advertised_host)
    elif is_wildcard_host(bind_host):
        raise MirrorConfigurationError(
            "Wildcard binding requires a distinct advertised host"
        )
    else:
        advertised = normalize_advertised_host(bind_host)

    return {
        "BindHost": bind_host,
        "AdvertisedHost": advertised,
        "Loopback": loopback,
        "LanMode": bool(lan_mode),
        "TlsEnabled": tls_enabled,
        "TlsCert": os.path.abspath(tls_cert) if tls_cert else "",
        "TlsKey": os.path.abspath(tls_key) if tls_key else "",
        "CleartextRiskAcknowledged": bool(acknowledge_cleartext),
    }


def create_bearer_token():
    return secrets.token_urlsafe(32)


def atomic_write_json(path, payload):
    path = os.path.abspath(path)
    folder = os.path.dirname(path)
    os.makedirs(folder, exist_ok=True)
    temp_path = os.path.join(
        folder,
        f".{os.path.basename(path)}.{uuid.uuid4().hex}.tmp",
    )
    with _INDEX_WRITE_LOCK:
        try:
            with open(
                temp_path,
                "x",
                encoding="utf-8",
                newline="\n",
            ) as handle:
                json.dump(payload, handle, indent=2, sort_keys=True)
                handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temp_path, path)
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
    return path


def _redact_client_address(value):
    try:
        address = ipaddress.ip_address(str(value))
    except ValueError:
        return "unknown"
    if address.is_loopback:
        return "loopback"
    if address.version == 4:
        network = ipaddress.ip_network(f"{address}/24", strict=False)
    else:
        network = ipaddress.ip_network(f"{address}/64", strict=False)
    return str(network)


class MirrorAuditLog:
    """Thread-safe, bounded JSONL access log without secrets or raw paths."""

    def __init__(
        self,
        path,
        *,
        max_bytes=MIRROR_AUDIT_MAX_BYTES,
        retention=MIRROR_AUDIT_RETENTION,
    ):
        self.path = os.path.abspath(path)
        self.max_bytes = max(4096, int(max_bytes))
        self.retention = max(1, min(10, int(retention)))
        self._lock = threading.Lock()

    def _rotate(self):
        if (
            not os.path.isfile(self.path)
            or os.path.getsize(self.path) < self.max_bytes
        ):
            return
        oldest = f"{self.path}.{self.retention}"
        if os.path.isfile(oldest):
            os.remove(oldest)
        for index in range(self.retention - 1, 0, -1):
            source = f"{self.path}.{index}"
            if os.path.isfile(source):
                os.replace(source, f"{self.path}.{index + 1}")
        os.replace(self.path, f"{self.path}.1")

    def write(
        self,
        *,
        client,
        method,
        route,
        status,
        bytes_sent,
        auth,
        request_id,
    ):
        record = {
            "Timestamp": utc_timestamp(),
            "RequestId": str(request_id),
            "ClientNetwork": _redact_client_address(client),
            "Method": str(method),
            "Route": str(route),
            "Status": int(status),
            "BytesSent": int(bytes_sent),
            "Authorization": str(auth),
        }
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with self._lock:
            self._rotate()
            with open(
                self.path,
                "a",
                encoding="utf-8",
                newline="\n",
            ) as handle:
                handle.write(json.dumps(record, sort_keys=True))
                handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())


def parse_byte_range(value, size):
    if not value:
        return None
    match = re.fullmatch(r"bytes=(\d*)-(\d*)", str(value).strip())
    if not match or "," in str(value):
        raise ValueError("Unsupported Range header")
    start_text, end_text = match.groups()
    if not start_text and not end_text:
        raise ValueError("Empty Range header")
    if not start_text:
        length = int(end_text)
        if length <= 0:
            raise ValueError("Invalid suffix range")
        start = max(0, size - length)
        end = size - 1
    else:
        start = int(start_text)
        end = int(end_text) if end_text else size - 1
        if start >= size or end < start:
            raise ValueError("Unsatisfiable range")
        end = min(end, size - 1)
    return start, end


def _copy_file_range(source, output, start, count):
    source.seek(start)
    remaining = count
    while remaining > 0:
        chunk = source.read(min(1024 * 1024, remaining))
        if not chunk:
            break
        output.write(chunk)
        remaining -= len(chunk)
    return count - remaining


def make_mirror_handler(
    *,
    index_name,
    index_payload,
    package_routes,
    app_version,
    audit_log,
    bearer_token=None,
    token_expires_at=None,
):
    index_bytes = (
        json.dumps(index_payload, indent=2, sort_keys=True) + "\n"
    ).encode("utf-8")
    routes = {}
    for route, package in package_routes.items():
        if isinstance(package, dict):
            routes[str(route)] = {
                "Path": os.path.abspath(package["Path"]),
                "SizeBytes": int(package["SizeBytes"]),
                "Sha256": str(package["Sha256"]).lower(),
            }
        else:
            path = os.path.abspath(package)
            routes[str(route)] = {
                "Path": path,
                "SizeBytes": os.path.getsize(path),
                "Sha256": "",
            }
    index_routes = {"/", f"/{index_name}"}
    token = str(bearer_token or "")
    expires_at = float(token_expires_at or 0)

    class MirrorRequestHandler(BaseHTTPRequestHandler):
        server_version = f"MSStoreHelperMirror/{app_version}"
        protocol_version = "HTTP/1.1"

        def do_GET(self):
            self._serve(head_only=False)

        def do_HEAD(self):
            self._serve(head_only=True)

        def do_POST(self):
            self._reject_method()

        def do_PUT(self):
            self._reject_method()

        def do_DELETE(self):
            self._reject_method()

        def log_message(self, _format, *_args):
            return

        def _audit(
            self,
            request_id,
            route,
            status,
            bytes_sent,
            auth,
        ):
            audit_log.write(
                client=self.client_address[0],
                method=self.command,
                route=route,
                status=status,
                bytes_sent=bytes_sent,
                auth=auth,
                request_id=request_id,
            )

        def _authorization_state(self):
            if not token:
                return True, "not-required"
            supplied = self.headers.get("Authorization", "")
            if time.time() >= expires_at:
                return False, "expired"
            if not supplied:
                return False, "missing"
            prefix = "Bearer "
            if not supplied.startswith(prefix):
                return False, "rejected"
            candidate = supplied[len(prefix):]
            if not hmac.compare_digest(candidate, token):
                return False, "rejected"
            return True, "accepted"

        def _send_common_headers(self):
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-MSStoreHelper-Mirror", app_version)

        def _send_error_response(
            self,
            status,
            *,
            request_id,
            route,
            auth,
            extra_headers=None,
        ):
            body = b""
            self.send_response(status)
            self._send_common_headers()
            for name, value in (extra_headers or {}).items():
                self.send_header(name, value)
            self.send_header("Content-Length", "0")
            self.end_headers()
            self._audit(
                request_id,
                route,
                status,
                len(body),
                auth,
            )

        def _reject_method(self):
            request_id = str(uuid.uuid4())
            self._send_error_response(
                405,
                request_id=request_id,
                route="method-not-allowed",
                auth="not-evaluated",
                extra_headers={"Allow": "GET, HEAD"},
            )

        def _serve(self, *, head_only):
            request_id = str(uuid.uuid4())
            parsed = urlsplit(self.path)
            if parsed.query or parsed.fragment:
                self._send_error_response(
                    404,
                    request_id=request_id,
                    route="not-found",
                    auth="not-evaluated",
                )
                return
            authorized, auth_state = self._authorization_state()
            if not authorized:
                self._send_error_response(
                    401,
                    request_id=request_id,
                    route="authorization",
                    auth=auth_state,
                    extra_headers={
                        "WWW-Authenticate": (
                            'Bearer realm="MSStoreHelper mirror"'
                        ),
                    },
                )
                return
            if parsed.path in index_routes:
                self._serve_index(
                    index_bytes,
                    request_id,
                    auth_state,
                    head_only,
                )
                return
            package = routes.get(parsed.path)
            if package is None:
                self._send_error_response(
                    404,
                    request_id=request_id,
                    route="not-found",
                    auth=auth_state,
                )
                return
            self._serve_package(
                parsed.path,
                package,
                request_id,
                auth_state,
                head_only,
            )

        def _serve_index(
            self,
            payload,
            request_id,
            auth_state,
            head_only,
        ):
            self.send_response(200)
            self._send_common_headers()
            self.send_header(
                "Content-Type",
                "application/json; charset=utf-8",
            )
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            sent = 0
            if not head_only:
                self.wfile.write(payload)
                sent = len(payload)
            self._audit(
                request_id,
                "index",
                200,
                sent,
                auth_state,
            )

        def _serve_package(
            self,
            route,
            package,
            request_id,
            auth_state,
            head_only,
        ):
            path = package["Path"]
            if (
                not os.path.isfile(path)
                or os.path.islink(path)
                or (
                    getattr(os.path, "isjunction", None)
                    and os.path.isjunction(path)
                )
            ):
                self._send_error_response(
                    404,
                    request_id=request_id,
                    route="package",
                    auth=auth_state,
                )
                return
            try:
                source = open(path, "rb")
                stat = os.fstat(source.fileno())
                size = stat.st_size
                digest = hashlib.sha256()
                for chunk in iter(
                    lambda: source.read(1024 * 1024),
                    b"",
                ):
                    digest.update(chunk)
                actual_sha256 = digest.hexdigest()
                expected_sha256 = package["Sha256"]
                if (
                    size != package["SizeBytes"]
                    or (
                        expected_sha256
                        and not hmac.compare_digest(
                            actual_sha256,
                            expected_sha256,
                        )
                    )
                ):
                    source.close()
                    self._send_error_response(
                        404,
                        request_id=request_id,
                        route="package-integrity",
                        auth=auth_state,
                    )
                    return
                byte_range = parse_byte_range(
                    self.headers.get("Range"),
                    size,
                )
            except (OSError, ValueError):
                try:
                    source.close()
                except (NameError, OSError):
                    pass
                self._send_error_response(
                    416,
                    request_id=request_id,
                    route="package",
                    auth=auth_state,
                    extra_headers={
                        "Content-Range": (
                            f"bytes */{os.path.getsize(path)}"
                            if os.path.isfile(path)
                            else "bytes */0"
                        ),
                    },
                )
                return

            if byte_range is None:
                start, end = 0, size - 1
                status = 200
            else:
                start, end = byte_range
                status = 206
            length = max(0, end - start + 1)
            self.send_response(status)
            self._send_common_headers()
            self.send_header(
                "Content-Type",
                "application/octet-stream",
            )
            self.send_header("Accept-Ranges", "bytes")
            self.send_header("Content-Length", str(length))
            self.send_header(
                "ETag",
                f'"sha256-{actual_sha256}"',
            )
            if status == 206:
                self.send_header(
                    "Content-Range",
                    f"bytes {start}-{end}/{size}",
                )
            self.end_headers()
            sent = 0
            if not head_only:
                try:
                    sent = _copy_file_range(
                        source,
                        self.wfile,
                        start,
                        length,
                    )
                except (BrokenPipeError, ConnectionResetError, OSError):
                    sent = 0
            source.close()
            self._audit(
                request_id,
                "package",
                status,
                sent,
                auth_state,
            )

    return MirrorRequestHandler


def wrap_server_tls(server, cert_path, key_path):
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.minimum_version = ssl.TLSVersion.TLSv1_2
    context.load_cert_chain(certfile=cert_path, keyfile=key_path)
    server.socket = context.wrap_socket(
        server.socket,
        server_side=True,
    )
    return server
