#!/usr/bin/env python3
"""Fail-closed recursive redaction and prepared diagnostic ZIP output."""

import hashlib
import ipaddress
import json
import math
import os
import re
import tempfile
import uuid
import zipfile
from datetime import date, datetime
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit


class DiagnosticRedactionError(ValueError):
    """Raised when diagnostic data cannot be serialized safely."""


_OMIT = object()
_SECRET_KEY = re.compile(
    r"(?i)(?:authorization|password|passwd|pwd|secret|token|"
    r"api[_-]?key|access[_-]?key|private[_-]?key|client[_-]?secret|"
    r"credential|cookie)"
)
_SECRET_QUERY_KEY = re.compile(
    r"(?i)^(?:"
    r"sig|signature|token|access_token|id_token|code|key|api[_-]?key|"
    r"password|passwd|secret|credential|x-amz-credential|"
    r"x-amz-signature|x-goog-signature|xamzsignature|"
    r"xgoogsignature|se|sp|sv|sas"
    r")$"
)
_URL_KEY = re.compile(
    r"(?i)(?:^|[_-])(?:url|uri|href)(?:$|[_-])|(?:url|uri|href)$"
)
_PATH_KEY = re.compile(
    r"(?i)(?:^|[_-])(?:"
    r"path|paths|folder|directory|executable|working[_-]?directory"
    r")(?:$|[_-])"
)
_URL_IN_TEXT = re.compile(r"(?i)\bhttps?://[^\s<>\"']+")
_AUTHORIZATION_LINE = re.compile(
    r"(?im)(\b(?:proxy-)?authorization\s*[:=]\s*)[^\r\n]*"
)
_NAMED_SECRET = re.compile(
    r"(?im)(\b(?:api[_-]?key|password|passwd|pwd|secret|token|"
    r"client[_-]?secret|access[_-]?key|credential)\b\s*[:=]\s*)"
    r"(?:\"[^\r\n\"]*\"|'[^\r\n']*'|[^\r\n,;}]+)"
)
_SWITCH_SECRET = re.compile(
    r"(?i)(--?(?:api[_-]?key|password|passwd|secret|token|"
    r"client[_-]?secret)(?:\s+|=))"
    r"(?:\"[^\r\n\"]*\"|'[^\r\n']*'|[^\s,;}]+)"
)
_WINDOWS_PATH = re.compile(
    r"(?i)(?<![A-Za-z0-9])(?:[A-Z]:[\\/]|\\\\[^\\/\s]+[\\/])"
    r"[^\r\n\"'<>|,;]*"
)


def _default_path_tokens():
    return {
        "USERPROFILE": os.environ.get("USERPROFILE"),
        "APPDATA": os.environ.get("APPDATA"),
        "LOCALAPPDATA": os.environ.get("LOCALAPPDATA"),
        "PROGRAMDATA": (
            os.environ.get("PROGRAMDATA")
            or os.environ.get("ProgramData")
        ),
        "WINDIR": os.environ.get("WINDIR"),
        "TEMP": tempfile.gettempdir(),
    }


def _path_token_patterns(path_tokens=None):
    tokens = _default_path_tokens()
    tokens.update(path_tokens or {})
    patterns = []
    for label, value in tokens.items():
        value = str(value or "").strip().rstrip("\\/")
        if not value:
            continue
        variants = {
            value,
            value.replace("\\", "/"),
            value.replace("/", "\\"),
        }
        for variant in variants:
            patterns.append((
                len(variant),
                re.compile(re.escape(variant), re.IGNORECASE),
                f"%{str(label).upper()}%",
            ))
    patterns.sort(key=lambda item: item[0], reverse=True)
    return patterns


def sanitize_url(value):
    value = str(value or "")
    try:
        parsed = urlsplit(value)
        hostname = parsed.hostname
        port = parsed.port
    except ValueError:
        return "[REDACTED_URL]"
    if parsed.scheme.lower() not in {"http", "https"} or not hostname:
        return "[REDACTED_URL]"
    try:
        address = ipaddress.ip_address(hostname)
    except ValueError:
        rendered_host = hostname
    else:
        rendered_host = f"[{hostname}]" if address.version == 6 else hostname
    netloc = rendered_host
    if port is not None:
        netloc = f"{netloc}:{port}"
    safe_query = [
        (key, value)
        for key, value in parse_qsl(
            parsed.query,
            keep_blank_values=True,
        )
        if not _SECRET_QUERY_KEY.fullmatch(key)
    ]
    return urlunsplit((
        parsed.scheme.lower(),
        netloc,
        parsed.path,
        urlencode(safe_query, doseq=True),
        "",
    ))


def _sanitize_url_match(match):
    value = match.group(0)
    trailing = ""
    while value and value[-1] in ".,);]":
        trailing = value[-1] + trailing
        value = value[:-1]
    return sanitize_url(value) + trailing


def redact_text(value, *, path_tokens=None):
    text = str(value or "")
    text = _URL_IN_TEXT.sub(_sanitize_url_match, text)
    text = _AUTHORIZATION_LINE.sub(r"\1[REDACTED]", text)
    text = _SWITCH_SECRET.sub(r"\1[REDACTED]", text)
    text = _NAMED_SECRET.sub(r"\1[REDACTED]", text)
    for _length, pattern, replacement in _path_token_patterns(path_tokens):
        text = pattern.sub(replacement, text)
    text = _WINDOWS_PATH.sub("%LOCAL_PATH%", text)
    return text


def _redact_value(value, *, key=None, path_tokens=None):
    if key is not None and _SECRET_KEY.search(key):
        return _OMIT
    if value is None or isinstance(value, (bool, int)):
        return value
    if isinstance(value, float):
        if not math.isfinite(value):
            raise DiagnosticRedactionError(
                "Non-finite diagnostic numbers are not supported"
            )
        return value
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Path):
        return redact_text(str(value), path_tokens=path_tokens)
    if isinstance(value, str):
        if key is not None and _URL_KEY.search(key):
            return sanitize_url(value)
        return redact_text(value, path_tokens=path_tokens)
    if isinstance(value, (list, tuple)):
        return [
            item
            for item in (
                _redact_value(
                    child,
                    path_tokens=path_tokens,
                )
                for child in value
            )
            if item is not _OMIT
        ]
    if isinstance(value, dict):
        redacted = {}
        for child_key, child_value in value.items():
            if not isinstance(child_key, str):
                raise DiagnosticRedactionError(
                    "Diagnostic object keys must be strings"
                )
            if _SECRET_KEY.search(child_key):
                continue
            if _SECRET_QUERY_KEY.fullmatch(child_key):
                continue
            safe_value = _redact_value(
                child_value,
                key=child_key,
                path_tokens=path_tokens,
            )
            if safe_value is not _OMIT:
                redacted[child_key] = safe_value
        return redacted
    raise DiagnosticRedactionError(
        f"Unsupported diagnostic value type: {type(value).__name__}"
    )


def redact_structure(value, *, path_tokens=None):
    redacted = _redact_value(value, path_tokens=path_tokens)
    if redacted is _OMIT:
        raise DiagnosticRedactionError(
            "Diagnostic root cannot be a secret value"
        )
    return redacted


def _json_bytes(value):
    return (
        json.dumps(
            value,
            indent=2,
            sort_keys=True,
            ensure_ascii=False,
        )
        + "\n"
    ).encode("utf-8")


def prepare_diagnostic_entries(
    *,
    diagnostics,
    source_health,
    queue,
    app_log,
    powershell_transcript,
    repair_manifests,
    operation_history=None,
    path_tokens=None,
):
    structured = {
        "diagnostics.json": diagnostics,
        "source-health.json": source_health,
        "queue.json": queue,
        "repair-manifests.json": repair_manifests,
        "operation-history.json": operation_history or [],
    }
    entries = {}
    for name, value in structured.items():
        entries[name] = _json_bytes(
            redact_structure(value, path_tokens=path_tokens)
        )
    entries["app-log.txt"] = (
        redact_text(app_log, path_tokens=path_tokens) + "\n"
    ).encode("utf-8")
    entries["powershell-transcript.txt"] = (
        redact_text(
            powershell_transcript,
            path_tokens=path_tokens,
        )
        + "\n"
    ).encode("utf-8")
    return dict(sorted(entries.items()))


def bundle_inventory(entries):
    inventory = []
    for name, content in sorted(entries.items()):
        if not isinstance(name, str) or "/" in name or "\\" in name:
            raise DiagnosticRedactionError(
                "Diagnostic entry name is invalid"
            )
        if not isinstance(content, bytes):
            raise DiagnosticRedactionError(
                f"Diagnostic entry is not bytes: {name}"
            )
        inventory.append({
            "Name": name,
            "SizeBytes": len(content),
            "Sha256": hashlib.sha256(content).hexdigest(),
        })
    return inventory


def diagnostic_preview_text(entries):
    lines = ["ZIP inventory:"]
    for item in bundle_inventory(entries):
        lines.append(
            f"- {item['Name']} — {item['SizeBytes']} bytes — "
            f"SHA-256 {item['Sha256']}"
        )
    for name, content in sorted(entries.items()):
        lines.extend([
            "",
            f"===== {name} =====",
            content.decode("utf-8"),
        ])
    return "\n".join(lines).rstrip() + "\n"


def write_prepared_bundle(bundle_path, entries):
    inventory = bundle_inventory(entries)
    bundle_path = os.path.abspath(bundle_path)
    folder = os.path.dirname(bundle_path)
    os.makedirs(folder, exist_ok=True)
    temp_path = os.path.join(
        folder,
        f".{os.path.basename(bundle_path)}.{uuid.uuid4().hex}.tmp",
    )
    try:
        with zipfile.ZipFile(
            temp_path,
            "x",
            compression=zipfile.ZIP_DEFLATED,
        ) as archive:
            for name, content in sorted(entries.items()):
                archive.writestr(name, content)
        with zipfile.ZipFile(temp_path, "r") as archive:
            if sorted(archive.namelist()) != sorted(entries):
                raise DiagnosticRedactionError(
                    "Diagnostic ZIP inventory verification failed"
                )
            for item in inventory:
                content = archive.read(item["Name"])
                if hashlib.sha256(content).hexdigest() != item["Sha256"]:
                    raise DiagnosticRedactionError(
                        "Diagnostic ZIP content verification failed"
                    )
        os.replace(temp_path, bundle_path)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
    return bundle_path
