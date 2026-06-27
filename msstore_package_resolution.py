#!/usr/bin/env python3
"""Package selection and install-order helpers for MSStoreHelper."""

import os
import re

INSTALLABLE_EXTENSIONS = {".appx", ".msix", ".appxbundle", ".msixbundle"}

ROLE_ORDER = {
    "net_native_framework": 10,
    "net_native_runtime": 20,
    "vclibs": 30,
    "ui_xaml": 40,
    "windows_app_runtime": 50,
    "app": 100,
}

ROLE_LABELS = {
    "net_native_framework": ".NET Native Framework",
    "net_native_runtime": ".NET Native Runtime",
    "vclibs": "VCLibs",
    "ui_xaml": "UI.Xaml",
    "windows_app_runtime": "Windows App Runtime",
    "app": "App",
}

DEPENDENCY_PREFIXES = (
    ("net_native_framework", ("microsoft.net.native.framework",)),
    ("net_native_runtime", ("microsoft.net.native.runtime",)),
    ("vclibs", ("microsoft.vclibs",)),
    ("ui_xaml", ("microsoft.ui.xaml",)),
    (
        "windows_app_runtime",
        ("microsoft.windowsappruntime", "microsoft.winapp.runtime"),
    ),
)


def package_identity(filename):
    """Return the AppX package identity name from a Store filename."""
    base = os.path.basename(filename)
    stem, _ = os.path.splitext(base)
    return stem.split("_", 1)[0]


def package_role(filename):
    identity = package_identity(filename).lower()
    for role, prefixes in DEPENDENCY_PREFIXES:
        if any(identity.startswith(prefix) for prefix in prefixes):
            return role
    return "app"


def package_role_label(filename):
    return ROLE_LABELS.get(package_role(filename), "App")


def is_dependency_package(package):
    return package_role(package["FileName"]) != "app"


def package_extension(package):
    filename = package["FileName"]
    ext = os.path.splitext(filename)[1].lower()
    if ext:
        return ext
    file_type = package.get("FileType", "").lower()
    return f".{file_type}" if file_type else ""


def is_installable_package(package):
    ext = package_extension(package)
    if ext not in INSTALLABLE_EXTENSIONS:
        return False
    return not package.get("IsEncrypted", False)


def package_architecture(package):
    arch = package.get("Architecture")
    if arch:
        return str(arch).lower()

    lower = package["FileName"].lower()
    if "_arm64_" in lower:
        return "arm64"
    if "_x64_" in lower or "_amd64_" in lower:
        return "x64"
    if "_x86_" in lower:
        return "x86"
    if "_arm_" in lower:
        return "arm"
    return "neutral"


def is_bundle_package(package):
    if package.get("IsBundle"):
        return True
    return package_extension(package) in {".appxbundle", ".msixbundle"}


def is_arch_compatible(package, target_arch):
    arch = package_architecture(package)
    if is_bundle_package(package):
        return True
    return arch in {target_arch.lower(), "neutral"}


def package_version_tuple(filename):
    """Parse the package version from Name_version_arch_publisher.ext."""
    stem, _ = os.path.splitext(os.path.basename(filename))
    parts = stem.split("_")
    for part in parts[1:]:
        if re.fullmatch(r"\d+(?:\.\d+){1,4}", part):
            return tuple(int(piece) for piece in part.split("."))

    matches = re.findall(r"\d+(?:\.\d+){1,4}", stem)
    if not matches:
        return ()
    return tuple(int(piece) for piece in matches[-1].split("."))


def version_tuple_from_text(text):
    match = re.search(r"\d+(?:\.\d+){1,4}", str(text or ""))
    if not match:
        return ()
    return tuple(int(piece) for piece in match.group(0).split("."))


def format_version_tuple(version_tuple):
    return ".".join(str(piece) for piece in version_tuple) if version_tuple else "unknown"


def compare_version_tuples(left, right):
    width = max(len(left), len(right))
    padded_left = tuple(left) + (0,) * (width - len(left))
    padded_right = tuple(right) + (0,) * (width - len(right))
    if padded_left == padded_right:
        return 0
    return 1 if padded_left > padded_right else -1


def installed_version_satisfies_package(package, installed_version):
    available = package_version_tuple(package["FileName"])
    installed = version_tuple_from_text(installed_version)
    if not installed or not available:
        return False
    return compare_version_tuples(installed, available) >= 0


def _candidate_score(package, target_arch, prefer_exact_arch=False):
    arch = package_architecture(package)
    if prefer_exact_arch:
        return (
            1 if arch == target_arch.lower() else 0,
            package_version_tuple(package["FileName"]),
            1 if is_bundle_package(package) else 0,
            1 if arch == "neutral" else 0,
        )

    return (
        package_version_tuple(package["FileName"]),
        1 if is_bundle_package(package) else 0,
        1 if arch == target_arch.lower() else 0,
        1 if arch == "neutral" else 0,
    )


def select_recommended_packages(packages, target_arch, prefer_exact_arch=False):
    """Pick best installable packages, including dependency frameworks."""
    best_by_identity = {}
    for package in packages:
        if not is_installable_package(package):
            continue
        if not is_arch_compatible(package, target_arch):
            continue

        key = package_identity(package["FileName"]).lower()
        current = best_by_identity.get(key)
        if current is None or _candidate_score(package, target_arch, prefer_exact_arch) > _candidate_score(current, target_arch, prefer_exact_arch):
            best_by_identity[key] = package

    return order_packages_for_install(best_by_identity.values(), target_arch)


def order_packages_for_install(packages, target_arch):
    """Install frameworks before app packages while preserving arch preference."""
    return sorted(
        packages,
        key=lambda package: (
            ROLE_ORDER.get(package_role(package["FileName"]), ROLE_ORDER["app"]),
            package_identity(package["FileName"]).lower(),
            package_architecture(package) != target_arch.lower(),
            package_version_tuple(package["FileName"]),
            package["FileName"].lower(),
        ),
    )


def annotate_package(package):
    """Add resolver metadata used by the GUI without changing API shape."""
    role = package_role(package["FileName"])
    package["PackageRole"] = role
    package["PackageRoleLabel"] = ROLE_LABELS.get(role, "App")
    package["InstallOrder"] = ROLE_ORDER.get(role, ROLE_ORDER["app"])
    package["PackageIdentity"] = package_identity(package["FileName"])
    package["AvailableVersion"] = format_version_tuple(package_version_tuple(package["FileName"]))
    return package
