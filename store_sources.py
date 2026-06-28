#!/usr/bin/env python3
"""Store source health, retry, and fallback helpers."""

import shutil
import subprocess
import re


class StoreSourceError(RuntimeError):
    def __init__(self, source_name, errors):
        self.source_name = source_name
        self.errors = list(errors)
        super().__init__(f"{source_name}: {'; '.join(self.errors)}")


def _clean_output(value):
    text = " ".join(str(value or "").strip().split())
    return text.encode("ascii", "replace").decode("ascii")


def _version_from_output(output):
    text = _clean_output(output)
    match = re.search(r"\bv?\d+(?:\.\d+){1,5}(?:[-+][A-Za-z0-9_.-]+)?", text)
    if match:
        return match.group(0)
    return text[:120]


def _creation_flags():
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def request_with_retries(source_name, request_callable, attempts=2, retry_statuses=None):
    retry_statuses = set(retry_statuses or {408, 429, 500, 502, 503, 504})
    errors = []
    attempts = max(1, int(attempts or 1))

    for attempt in range(1, attempts + 1):
        try:
            response = request_callable()
            status_code = getattr(response, "status_code", None)
            if status_code in retry_statuses and attempt < attempts:
                errors.append(f"attempt {attempt}: HTTP {status_code}")
                continue
            response.raise_for_status()
            return response, errors
        except Exception as exc:
            errors.append(f"attempt {attempt}: {type(exc).__name__}: {exc}")
            if attempt >= attempts:
                raise StoreSourceError(source_name, errors)

    raise StoreSourceError(source_name, errors)


def detect_http_source(key, name, endpoint, request_callable):
    status = {
        "Key": key,
        "Name": name,
        "Kind": "http",
        "Available": False,
        "Detail": "Not checked",
        "Endpoint": endpoint,
    }
    try:
        response = request_callable()
        code = getattr(response, "status_code", None)
        if code is None or 200 <= code < 400:
            status["Available"] = True
            status["Detail"] = f"HTTP {code}" if code is not None else "Responded"
        else:
            status["Detail"] = f"HTTP {code}"
    except Exception as exc:
        status["Detail"] = f"{type(exc).__name__}: {exc}"
    return status


def detect_command_source(key, name, commands, version_args=None, which=shutil.which, run=subprocess.run, timeout=5):
    status = {
        "Key": key,
        "Name": name,
        "Kind": "command",
        "Available": False,
        "Detail": "Command not found",
        "Command": "",
        "Version": "",
    }
    version_args = version_args or ["--version"]

    for command in commands:
        path = which(command)
        if not path:
            continue

        status["Available"] = True
        status["Command"] = command
        status["Detail"] = f"Found at {path}"
        try:
            result = run(
                [path] + list(version_args),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                creationflags=_creation_flags(),
            )
            output = result.stdout or result.stderr
            if output:
                status["Version"] = _version_from_output(output)
            if result.returncode != 0:
                status["Detail"] = f"Found at {path}; version probe exited {result.returncode}"
        except Exception as exc:
            status["Detail"] = f"Found at {path}; version probe failed: {type(exc).__name__}: {exc}"
        return status

    return status


def detect_source_health(storeedge_request, rgadguard_request, which=shutil.which, run=subprocess.run):
    return [
        detect_http_source(
            "storeedgefd",
            "Microsoft Store Search API",
            "https://storeedgefd.dsx.mp.microsoft.com/v9.0/manifestSearch",
            storeedge_request,
        ),
        detect_http_source(
            "rg-adguard",
            "RG-Adguard package proxy",
            "https://store.rg-adguard.net/api/GetFiles",
            rgadguard_request,
        ),
        detect_command_source(
            "winget",
            "WinGet msstore source",
            ["winget"],
            ["--version"],
            which=which,
            run=run,
        ),
        detect_command_source(
            "store-cli",
            "Microsoft Store CLI",
            ["store"],
            ["--help"],
            which=which,
            run=run,
        ),
    ]


def package_lookup_fallbacks(product_id, statuses):
    available = {status.get("Key"): status for status in statuses if status.get("Available")}
    fallbacks = []

    if "winget" in available:
        fallbacks.append({
            "Source": "WinGet",
            "Command": f"winget install --source msstore --id {product_id} --accept-source-agreements --accept-package-agreements",
            "Detail": "Installs from the WinGet msstore source when direct package URLs are unavailable.",
        })

    if "store-cli" in available:
        fallbacks.append({
            "Source": "Store CLI",
            "Command": f"store install {product_id}",
            "Detail": "Requires Microsoft Store to be enabled on the PC.",
        })

    return fallbacks


def source_status_summary(status):
    prefix = "available" if status.get("Available") else "unavailable"
    detail = status.get("Detail") or ""
    version = status.get("Version") or ""
    if version:
        detail = f"{detail}; {version}" if detail else version
    return f"{status.get('Name', 'Source')}: {prefix} ({detail})"
