#!/usr/bin/env python3
"""Fail-closed package filename, URL, and filesystem ingress helpers."""

import ntpath
import os
import re
import unicodedata
from urllib.parse import unquote, urljoin, urlsplit


PACKAGE_EXTENSIONS = frozenset({
    ".appx",
    ".appxbundle",
    ".eappx",
    ".eappxbundle",
    ".emsix",
    ".emsixbundle",
    ".msix",
    ".msixbundle",
})
INSTALLABLE_PACKAGE_EXTENSIONS = frozenset({
    ".appx",
    ".appxbundle",
    ".msix",
    ".msixbundle",
})
NETWORK_PACKAGE_SCHEMES = frozenset({"http", "https"})
WINDOWS_INVALID_FILENAME_CHARACTERS = frozenset('<>:"/\\|?*')
WINDOWS_RESERVED_NAMES = frozenset(
    {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"}
    | {f"COM{index}" for index in range(1, 10)}
    | {f"LPT{index}" for index in range(1, 10)}
    | {f"COM{index}" for index in "¹²³"}
    | {f"LPT{index}" for index in "¹²³"}
)


class PackageIngressError(ValueError):
    """Raised when package-controlled input cannot safely cross a boundary."""


def validate_package_filename(value):
    """Return a normalized, basename-only Windows package filename."""
    if not isinstance(value, str) or not value:
        raise PackageIngressError("Package filename is missing")
    if value != value.strip():
        raise PackageIngressError("Package filename has leading or trailing whitespace")

    filename = unicodedata.normalize("NFC", value)
    if filename in {".", ".."}:
        raise PackageIngressError("Package filename is a traversal segment")
    if ntpath.basename(filename) != filename or ntpath.dirname(filename):
        raise PackageIngressError("Package filename must be a basename")
    if any(character in WINDOWS_INVALID_FILENAME_CHARACTERS for character in filename):
        raise PackageIngressError("Package filename contains a Windows-invalid character")
    if any(not character.isprintable() for character in filename):
        raise PackageIngressError("Package filename contains a control character")
    if filename.endswith((".", " ")):
        raise PackageIngressError("Package filename has an unsafe Windows suffix")
    if len(filename.encode("utf-16-le")) // 2 > 255:
        raise PackageIngressError("Package filename exceeds the Windows component limit")

    device_stem = filename.split(".", 1)[0].rstrip(" .").upper()
    if device_stem in WINDOWS_RESERVED_NAMES:
        raise PackageIngressError("Package filename uses a reserved Windows device name")

    extension = ntpath.splitext(filename)[1].lower()
    if extension not in PACKAGE_EXTENSIONS:
        raise PackageIngressError("Package filename has an unsupported or malformed extension")
    filename_stem = filename[:-len(extension)]
    if not filename_stem or filename_stem in {".", ".."}:
        raise PackageIngressError("Package filename has no valid stem")
    return filename


def validate_package_url(value):
    """Return a package URL after rejecting non-network and ambiguous forms."""
    if not isinstance(value, str) or not value:
        raise PackageIngressError("Package URL is missing")
    if (
        value != value.strip()
        or any(character.isspace() or not character.isprintable() for character in value)
    ):
        raise PackageIngressError("Package URL contains whitespace or control characters")
    if re.search(r"%(?![0-9A-Fa-f]{2})", value):
        raise PackageIngressError("Package URL has malformed percent encoding")
    decoded_url = unquote(value)
    if any(not character.isprintable() for character in decoded_url):
        raise PackageIngressError("Package URL contains an encoded control character")
    if "\\" in value or "\\" in decoded_url:
        raise PackageIngressError("Package URL contains a backslash")

    try:
        parsed = urlsplit(value)
        hostname = parsed.hostname
    except ValueError as exc:
        raise PackageIngressError("Package URL is malformed") from exc
    if parsed.scheme.lower() not in NETWORK_PACKAGE_SCHEMES:
        raise PackageIngressError("Package URL uses an unsafe scheme")
    if not hostname:
        raise PackageIngressError("Package URL has no host")
    if parsed.username is not None or parsed.password is not None:
        raise PackageIngressError("Package URL must not contain credentials")
    if parsed.fragment:
        raise PackageIngressError("Package URL must not contain a fragment")
    try:
        parsed.port
    except ValueError as exc:
        raise PackageIngressError("Package URL has an invalid port") from exc
    return value


def validate_response_redirects(initial_url, response):
    """Validate every observed redirect URL and reject HTTPS downgrades."""
    chain = [validate_package_url(initial_url)]
    for hop in list(getattr(response, "history", None) or []) + [response]:
        hop_url = getattr(hop, "url", None)
        if hop_url:
            chain.append(validate_package_url(hop_url))

        location = getattr(hop, "headers", {}).get("location")
        if location:
            base_url = hop_url or chain[-1]
            chain.append(validate_package_url(urljoin(base_url, location)))

    previous_scheme = urlsplit(chain[0]).scheme.lower()
    for url in chain[1:]:
        scheme = urlsplit(url).scheme.lower()
        if previous_scheme == "https" and scheme != "https":
            raise PackageIngressError("Package redirect downgraded HTTPS")
        previous_scheme = scheme
    return chain[-1]


def ensure_path_within_root(root, path):
    """Return a real absolute path only when it remains beneath root."""
    if not root:
        raise PackageIngressError("Package path root is missing")
    root_path = os.path.realpath(os.path.abspath(os.fspath(root)))
    candidate = os.path.realpath(os.path.abspath(os.fspath(path)))
    try:
        common = os.path.commonpath([root_path, candidate])
    except ValueError as exc:
        raise PackageIngressError("Package path is on a different filesystem root") from exc
    if os.path.normcase(common) != os.path.normcase(root_path) or candidate == root_path:
        raise PackageIngressError("Package path escapes its root")
    return candidate


def package_path(root, filename):
    """Build a confined package path from a validated local filename."""
    safe_filename = validate_package_filename(filename)
    root_path = os.path.realpath(os.path.abspath(os.fspath(root)))
    return ensure_path_within_root(root_path, os.path.join(root_path, safe_filename))


def validate_existing_package_path(
    path,
    *,
    expected_filename=None,
    root=None,
    require_file=False,
):
    """Validate an existing/local package path without trusting its metadata."""
    if not path:
        raise PackageIngressError("Package path is missing")
    absolute_path = os.path.abspath(os.fspath(path))
    filename = validate_package_filename(os.path.basename(absolute_path))
    if expected_filename is not None:
        expected = validate_package_filename(expected_filename)
        if os.path.normcase(filename) != os.path.normcase(expected):
            raise PackageIngressError("Package path does not match its validated filename")
    if root is not None:
        absolute_path = ensure_path_within_root(root, absolute_path)
    else:
        absolute_path = os.path.realpath(absolute_path)
    if os.path.islink(path):
        raise PackageIngressError("Package path must not be a symbolic link")
    if require_file and not os.path.isfile(absolute_path):
        raise PackageIngressError("Package file is missing")
    return absolute_path


def validate_package_record(package, *, require_url=False):
    """Copy package metadata after validating its path-bearing fields."""
    if not isinstance(package, dict):
        raise PackageIngressError("Package metadata must be an object")
    record = package.copy()
    filename = validate_package_filename(record.get("FileName"))
    safe_filename = record.get("SafeFileName")
    if safe_filename is not None:
        safe_filename = validate_package_filename(safe_filename)
        if os.path.normcase(safe_filename) != os.path.normcase(filename):
            raise PackageIngressError("Package filename does not match its safe filename")
    record["FileName"] = filename
    record["SafeFileName"] = filename

    url = record.get("Url")
    if require_url or url:
        record["Url"] = validate_package_url(url)
    return record
