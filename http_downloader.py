#!/usr/bin/env python3
"""Representation-bound, disk-bounded HTTP package downloads."""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable
from urllib.parse import urlsplit, urlunsplit

from package_ingress import (
    ensure_path_within_root,
    validate_package_filename,
    validate_package_url,
    validate_response_redirects,
)


RESUME_SCHEMA_VERSION = 1
DEFAULT_MAX_DOWNLOAD_BYTES = 16 * 1024 * 1024 * 1024
DEFAULT_FREE_SPACE_RESERVE_BYTES = 512 * 1024 * 1024
STALE_URL_STATUS_CODES = frozenset({401, 403, 404, 410})
CONTENT_RANGE_RE = re.compile(r"^bytes (\d+)-(\d+)/(\d+)$", re.IGNORECASE)
UNSATISFIED_RANGE_RE = re.compile(r"^bytes \*/(\d+)$", re.IGNORECASE)


class HttpDownloadError(RuntimeError):
    """Raised when a response cannot be written without violating limits."""


class StaleDownloadUrlError(HttpDownloadError):
    def __init__(self, status_code: int):
        self.status_code = int(status_code)
        super().__init__(f"Package URL returned HTTP {self.status_code}")


class DownloadCancelled(HttpDownloadError):
    """Raised after a resumable partial is safely persisted or discarded."""


def _timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


def _redacted_url(url: str) -> str:
    parsed = urlsplit(validate_package_url(url))
    return urlunsplit(
        (parsed.scheme, parsed.netloc, parsed.path, "", "")
    )


def _header(response: Any, name: str) -> str:
    headers = getattr(response, "headers", {}) or {}
    for key, value in headers.items():
        if str(key).lower() == name.lower():
            return str(value).strip()
    return ""


def _positive_header_int(response: Any, name: str) -> int | None:
    value = _header(response, name)
    if not value:
        return None
    try:
        number = int(value)
    except (TypeError, ValueError) as exc:
        raise HttpDownloadError(f"Invalid {name} response header") from exc
    if number < 0:
        raise HttpDownloadError(f"Invalid {name} response header")
    return number


def _file_sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _hash_existing(path: str) -> Any:
    digest = hashlib.sha256()
    if os.path.exists(path):
        with open(path, "rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    return digest


def _atomic_write_json(path: str, payload: dict[str, Any]) -> None:
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(
        f".{destination.name}.{uuid.uuid4().hex}.tmp"
    )
    try:
        with temporary.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(payload, stream, indent=2, sort_keys=True)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, destination)
    finally:
        try:
            temporary.unlink()
        except FileNotFoundError:
            pass


def _remove_partial(part_path: str, state_path: str) -> None:
    for path in (part_path, state_path):
        try:
            os.remove(path)
        except FileNotFoundError:
            pass


def _resume_validator(state: dict[str, Any]) -> tuple[str, str]:
    etag = str(state.get("ETag") or "")
    if etag and not etag.lower().startswith("w/"):
        return "ETag", etag
    last_modified = str(state.get("LastModified") or "")
    if last_modified:
        return "Last-Modified", last_modified
    return "", ""


def _load_verified_resume(
    part_path: str,
    state_path: str,
    *,
    source_identity: str,
    max_bytes: int,
) -> dict[str, Any] | None:
    if not os.path.exists(part_path) and not os.path.exists(state_path):
        return None
    try:
        with open(state_path, "r", encoding="utf-8") as stream:
            state = json.load(stream)
        downloaded = os.path.getsize(part_path)
        expected = int(state["ExpectedLength"])
        if (
            not isinstance(state, dict)
            or state.get("SchemaVersion") != RESUME_SCHEMA_VERSION
            or state.get("SourceIdentity") != source_identity
            or int(state.get("DownloadedBytes", -1)) != downloaded
            or downloaded <= 0
            or downloaded > max_bytes
            or expected < downloaded
            or expected > max_bytes
            or state.get("HashAlgorithm") != "sha256"
            or state.get("PartialSha256") != _file_sha256(part_path)
            or not any(_resume_validator(state))
        ):
            raise ValueError("resume metadata does not match the partial file")
        return state
    except (OSError, TypeError, ValueError, KeyError, json.JSONDecodeError):
        _remove_partial(part_path, state_path)
        return None


def _write_resume_state(
    state_path: str,
    *,
    source_url: str,
    effective_url: str,
    source_identity: str,
    etag: str,
    last_modified: str,
    expected_length: int,
    downloaded_bytes: int,
    partial_sha256: str,
) -> dict[str, Any]:
    state = {
        "SchemaVersion": RESUME_SCHEMA_VERSION,
        "SourceUrl": source_url,
        "EffectiveUrl": effective_url,
        "SourceIdentity": source_identity,
        "ETag": etag,
        "LastModified": last_modified,
        "ExpectedLength": int(expected_length),
        "DownloadedBytes": int(downloaded_bytes),
        "HashAlgorithm": "sha256",
        "PartialSha256": partial_sha256,
        "UpdatedAt": _timestamp(),
    }
    _atomic_write_json(state_path, state)
    return state


def _ensure_capacity(
    folder: str,
    *,
    remaining_bytes: int | None,
    reserve_bytes: int,
    reclaimed_bytes: int = 0,
) -> int:
    free = int(shutil.disk_usage(folder).free) + max(0, int(reclaimed_bytes))
    writable = free - reserve_bytes
    if writable <= 0:
        raise HttpDownloadError(
            "Destination does not have the configured free-space reserve"
        )
    if remaining_bytes is not None and remaining_bytes > writable:
        raise HttpDownloadError(
            f"Download needs {remaining_bytes} bytes but only {writable} "
            "bytes are available above the free-space reserve"
        )
    return writable


def validated_content_length(
    initial_url: str,
    response: Any,
    *,
    max_bytes: int = DEFAULT_MAX_DOWNLOAD_BYTES,
) -> int:
    validate_response_redirects(validate_package_url(initial_url), response)
    response.raise_for_status()
    length = _positive_header_int(response, "Content-Length")
    if length is None:
        raise HttpDownloadError("Response did not declare Content-Length")
    if length == 0:
        raise HttpDownloadError("Response declared a zero-byte package")
    if length > int(max_bytes):
        raise HttpDownloadError(
            f"Response length {length} exceeds the {int(max_bytes)}-byte limit"
        )
    return length


def download_http_file(
    url: str,
    filepath: str,
    *,
    filename: str,
    source_identity: str,
    get: Callable[..., Any],
    progress_callback: Callable[[float], None] | None = None,
    cancel_event: Any = None,
    max_bytes: int = DEFAULT_MAX_DOWNLOAD_BYTES,
    free_space_reserve_bytes: int = DEFAULT_FREE_SPACE_RESERVE_BYTES,
    timeout: float = 60,
    request_headers: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Download one exact representation into ``filepath`` and return evidence."""
    url = validate_package_url(url)
    filename = validate_package_filename(filename)
    max_bytes = int(max_bytes)
    reserve_bytes = int(free_space_reserve_bytes)
    if max_bytes <= 0 or reserve_bytes < 0:
        raise ValueError("download and reserve limits must be non-negative")

    destination_root = os.path.realpath(
        os.path.abspath(os.path.dirname(os.path.abspath(filepath)))
    )
    os.makedirs(destination_root, exist_ok=True)
    filepath = ensure_path_within_root(destination_root, filepath)
    if os.path.basename(filepath) != filename:
        raise HttpDownloadError(
            "Download destination does not match the validated filename"
        )
    part_path = ensure_path_within_root(destination_root, f"{filepath}.part")
    state_path = ensure_path_within_root(
        destination_root,
        f"{filepath}.part.json",
    )

    for request_attempt in range(2):
        resume = _load_verified_resume(
            part_path,
            state_path,
            source_identity=source_identity,
            max_bytes=max_bytes,
        )
        existing = os.path.getsize(part_path) if resume else 0
        headers: dict[str, str] = {
            str(key): str(value)
            for key, value in (request_headers or {}).items()
            if str(key).lower() not in {"range", "if-range"}
        }
        headers["Accept-Encoding"] = "identity"
        if resume:
            _validator_name, validator_value = _resume_validator(resume)
            headers.update({
                "Range": f"bytes={existing}-",
                "If-Range": validator_value,
            })
        if cancel_event is not None and cancel_event.is_set():
            raise DownloadCancelled("Download cancelled before the request started")

        with get(
            url,
            stream=True,
            timeout=timeout,
            headers=headers,
            allow_redirects=True,
        ) as response:
            effective_url = validate_response_redirects(url, response)
            status_code = int(getattr(response, "status_code", 200) or 200)
            if status_code in STALE_URL_STATUS_CODES:
                raise StaleDownloadUrlError(status_code)

            if status_code == 416:
                match = UNSATISFIED_RANGE_RE.fullmatch(
                    _header(response, "Content-Range")
                )
                total = int(match.group(1)) if match else -1
                if (
                    resume
                    and total == existing
                    and int(resume["ExpectedLength"]) == total
                ):
                    os.replace(part_path, filepath)
                    try:
                        os.remove(state_path)
                    except FileNotFoundError:
                        pass
                    return {
                        "SourceUrl": _redacted_url(url),
                        "EffectiveUrl": _redacted_url(effective_url),
                        "SizeBytes": total,
                        "ETag": resume.get("ETag", ""),
                        "LastModified": resume.get("LastModified", ""),
                        "Sha256": resume["PartialSha256"],
                        "Resumed": True,
                    }
                _remove_partial(part_path, state_path)
                if request_attempt == 0:
                    continue
                raise HttpDownloadError(
                    "Server returned an unusable HTTP 416 response"
                )

            response.raise_for_status()
            if status_code not in {200, 206}:
                raise HttpDownloadError(
                    f"Unexpected HTTP {status_code} download response"
                )

            etag = _header(response, "ETag")
            last_modified = _header(response, "Last-Modified")
            content_length = _positive_header_int(response, "Content-Length")
            expected_length = 0
            append = False

            if status_code == 206:
                content_range = CONTENT_RANGE_RE.fullmatch(
                    _header(response, "Content-Range")
                )
                if not resume or not content_range:
                    _remove_partial(part_path, state_path)
                    if request_attempt == 0:
                        continue
                    raise HttpDownloadError(
                        "HTTP 206 response lacked a valid resume request"
                    )
                start, end, total = (
                    int(content_range.group(index))
                    for index in range(1, 4)
                )
                response_validator_name, response_validator = (
                    _resume_validator({
                        "ETag": etag,
                        "LastModified": last_modified,
                    })
                )
                resume_validator_name, resume_validator = _resume_validator(resume)
                range_length = end - start + 1
                if (
                    start != existing
                    or end < start
                    or end >= total
                    or total != int(resume["ExpectedLength"])
                    or content_length == 0
                    or (
                        content_length is not None
                        and content_length != range_length
                    )
                    or response_validator_name != resume_validator_name
                    or response_validator != resume_validator
                ):
                    _remove_partial(part_path, state_path)
                    if request_attempt == 0:
                        continue
                    raise HttpDownloadError(
                        "HTTP 206 response did not match the saved representation"
                    )
                expected_length = total
                append = True
            else:
                if content_length == 0:
                    _remove_partial(part_path, state_path)
                    raise HttpDownloadError(
                        "Response declared a zero-byte package"
                    )
                expected_length = content_length or 0
                if existing:
                    try:
                        os.remove(state_path)
                    except FileNotFoundError:
                        pass

            if expected_length > max_bytes:
                _remove_partial(part_path, state_path)
                raise HttpDownloadError(
                    f"Response length {expected_length} exceeds the "
                    f"{max_bytes}-byte limit"
                )
            initial_bytes = existing if append else 0
            remaining = (
                expected_length - initial_bytes
                if expected_length
                else None
            )
            writable = _ensure_capacity(
                destination_root,
                remaining_bytes=remaining,
                reserve_bytes=reserve_bytes,
                reclaimed_bytes=existing if existing and not append else 0,
            )
            digest = _hash_existing(part_path) if append else hashlib.sha256()
            downloaded = initial_bytes
            mode = "ab" if append else "wb"
            resumable = bool(
                expected_length
                and (etag and not etag.lower().startswith("w/") or last_modified)
            )

            try:
                with open(part_path, mode) as stream:
                    for chunk in response.iter_content(chunk_size=1024 * 1024):
                        if cancel_event is not None and cancel_event.is_set():
                            raise DownloadCancelled(
                                "Download cancelled at a verified chunk boundary"
                            )
                        if not chunk:
                            continue
                        next_size = downloaded + len(chunk)
                        if next_size > max_bytes:
                            raise HttpDownloadError(
                                f"Response exceeded the {max_bytes}-byte limit"
                            )
                        if next_size - initial_bytes > writable:
                            raise HttpDownloadError(
                                "Download exhausted space above the "
                                "configured free-space reserve"
                            )
                        if expected_length and next_size > expected_length:
                            raise HttpDownloadError(
                                "Response body exceeded its declared length"
                            )
                        stream.write(chunk)
                        digest.update(chunk)
                        downloaded = next_size
                        if progress_callback and expected_length:
                            progress_callback(downloaded / expected_length)
                    stream.flush()
                    os.fsync(stream.fileno())
            except Exception:
                if (
                    resumable
                    and 0 < downloaded < expected_length
                    and os.path.exists(part_path)
                    and os.path.getsize(part_path) == downloaded
                ):
                    _write_resume_state(
                        state_path,
                        source_url=_redacted_url(url),
                        effective_url=_redacted_url(effective_url),
                        source_identity=source_identity,
                        etag=etag,
                        last_modified=last_modified,
                        expected_length=expected_length,
                        downloaded_bytes=downloaded,
                        partial_sha256=digest.hexdigest(),
                    )
                else:
                    _remove_partial(part_path, state_path)
                raise

            if downloaded == 0:
                _remove_partial(part_path, state_path)
                raise HttpDownloadError("Response produced a zero-byte package")
            if expected_length and downloaded != expected_length:
                if resumable and downloaded < expected_length:
                    _write_resume_state(
                        state_path,
                        source_url=_redacted_url(url),
                        effective_url=_redacted_url(effective_url),
                        source_identity=source_identity,
                        etag=etag,
                        last_modified=last_modified,
                        expected_length=expected_length,
                        downloaded_bytes=downloaded,
                        partial_sha256=digest.hexdigest(),
                    )
                else:
                    _remove_partial(part_path, state_path)
                raise HttpDownloadError(
                    f"Downloaded {downloaded} bytes; expected "
                    f"{expected_length} bytes"
                )

            os.replace(part_path, filepath)
            try:
                os.remove(state_path)
            except FileNotFoundError:
                pass
            return {
                "SourceUrl": _redacted_url(url),
                "EffectiveUrl": _redacted_url(effective_url),
                "SizeBytes": downloaded,
                "ETag": etag,
                "LastModified": last_modified,
                "Sha256": digest.hexdigest(),
                "Resumed": append,
            }

    raise HttpDownloadError("Download could not establish a safe response")
