#!/usr/bin/env python3
"""Manifest-first AppX/MSIX inspection and install planning."""

from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import PurePosixPath
from typing import Any


PLAN_SCHEMA_VERSION = 1
INSPECTION_SCHEMA_VERSION = 1
KNOWN_INVENTORY_STATUSES = frozenset({"success", "empty"})
PACKAGE_ARCHITECTURES = frozenset(
    {"x64", "x86", "arm64", "arm", "neutral", "resource"}
)
MAX_MANIFEST_BYTES = 4 * 1024 * 1024
MAX_NESTED_PACKAGE_BYTES = 8 * 1024 * 1024 * 1024


class AppxInspectionError(ValueError):
    """Raised when package structure cannot be inspected safely."""


class InstallPlanError(ValueError):
    """Raised when a queue cannot produce one deterministic install plan."""


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _version_tuple(value: Any) -> tuple[int, ...]:
    match = re.fullmatch(r"\d+(?:\.\d+){1,3}", str(value or "").strip())
    if not match:
        return ()
    return tuple(int(piece) for piece in match.group(0).split("."))


def _compare_versions(left: Any, right: Any) -> int:
    left_tuple = _version_tuple(left)
    right_tuple = _version_tuple(right)
    if not left_tuple or not right_tuple:
        return 0
    width = max(len(left_tuple), len(right_tuple))
    normalized_left = left_tuple + (0,) * (width - len(left_tuple))
    normalized_right = right_tuple + (0,) * (width - len(right_tuple))
    if normalized_left == normalized_right:
        return 0
    return 1 if normalized_left > normalized_right else -1


def _canonical_dicts(values: list[dict[str, Any]]) -> list[dict[str, Any]]:
    unique: dict[str, dict[str, Any]] = {}
    for value in values:
        key = json.dumps(value, sort_keys=True, separators=(",", ":"))
        unique[key] = value
    return [unique[key] for key in sorted(unique)]


def _identity(root: ET.Element, source: str) -> dict[str, str]:
    identities = [
        element
        for element in root.iter()
        if _local_name(element.tag) == "Identity"
    ]
    if len(identities) != 1:
        raise AppxInspectionError(
            f"{source} must contain exactly one package identity"
        )
    attributes = identities[0].attrib
    name = str(attributes.get("Name") or "").strip()
    publisher = str(attributes.get("Publisher") or "").strip()
    version = str(attributes.get("Version") or "").strip()
    architecture = str(
        attributes.get("ProcessorArchitecture")
        or attributes.get("Architecture")
        or "neutral"
    ).strip().lower()
    resource_id = str(attributes.get("ResourceId") or "").strip()
    if (
        not name
        or not publisher
        or not re.fullmatch(r"\d+(?:\.\d+){3}", version)
    ):
        raise AppxInspectionError(f"{source} has an incomplete identity")
    if architecture not in PACKAGE_ARCHITECTURES:
        raise AppxInspectionError(
            f"{source} declares unsupported architecture {architecture!r}"
        )
    return {
        "Name": name,
        "Publisher": publisher,
        "Version": version,
        "ProcessorArchitecture": architecture,
        "ResourceId": resource_id,
    }


def _property_flags(root: ET.Element) -> dict[str, bool]:
    values = {
        "Framework": False,
        "ResourcePackage": False,
        "AllowExecution": True,
    }
    for container in root.iter():
        if _local_name(container.tag) != "Properties":
            continue
        for element in container:
            name = _local_name(element.tag)
            if name not in values:
                continue
            text = str(element.text or "").strip().lower()
            if text in {"true", "false"}:
                values[name] = text == "true"
    return values


def _manifest_details(
    root: ET.Element,
    *,
    source: str,
    manifest_name: str,
) -> dict[str, Any]:
    identity = _identity(root, source)
    package_dependencies: list[dict[str, str]] = []
    main_dependencies: list[dict[str, str]] = []
    target_families: list[dict[str, str]] = []
    capabilities: list[dict[str, str]] = []
    applications: list[dict[str, str]] = []

    for container in root.iter():
        container_name = _local_name(container.tag)
        if container_name == "Dependencies":
            for element in container:
                kind = _local_name(element.tag)
                attributes = {
                    str(key): str(value).strip()
                    for key, value in element.attrib.items()
                    if str(value).strip()
                }
                if kind == "TargetDeviceFamily":
                    target_families.append(attributes)
                elif kind in {
                    "MainPackageDependency",
                    "MainAppPackageDependency",
                }:
                    attributes["Type"] = kind
                    main_dependencies.append(attributes)
                elif kind in {
                    "PackageDependency",
                    "HostRuntimeDependency",
                    "OSPackageDependency",
                }:
                    attributes["Type"] = kind
                    package_dependencies.append(attributes)
        elif container_name == "Capabilities":
            for element in container:
                kind = _local_name(element.tag)
                name = str(element.attrib.get("Name") or "").strip()
                if name:
                    capabilities.append({"Type": kind, "Name": name})
        elif container_name == "Applications":
            for element in container:
                if _local_name(element.tag) != "Application":
                    continue
                applications.append({
                    str(key): str(value).strip()
                    for key, value in element.attrib.items()
                    if str(value).strip()
                })

    target_families = _canonical_dicts(target_families)
    minimum_versions = [
        _version_tuple(item.get("MinVersion"))
        for item in target_families
        if _version_tuple(item.get("MinVersion"))
    ]
    min_os_version = (
        ".".join(str(piece) for piece in max(minimum_versions))
        if minimum_versions
        else ""
    )
    package_dependencies = _canonical_dicts(package_dependencies)
    main_dependencies = _canonical_dicts(main_dependencies)
    capabilities = _canonical_dicts(capabilities)
    return {
        "ManifestName": manifest_name,
        "Identity": identity,
        "Properties": _property_flags(root),
        "MinOSVersion": min_os_version,
        "TargetDeviceFamilies": target_families,
        "PackageDependencies": package_dependencies,
        "MainPackageDependencies": main_dependencies,
        "CapabilitiesDetailed": capabilities,
        "Capabilities": [
            f"{item['Type']}: {item['Name']}"
            for item in capabilities
        ],
        "Dependencies": sorted(
            [
                _format_dependency(item)
                for item in package_dependencies + main_dependencies
            ]
            + [
                _format_target_family(item)
                for item in target_families
            ]
        ),
        "Applications": applications,
    }


def _format_dependency(dependency: dict[str, str]) -> str:
    kind = dependency.get("Type") or "PackageDependency"
    name = dependency.get("Name") or "unknown"
    minimum = dependency.get("MinVersion") or ""
    publisher = dependency.get("Publisher") or ""
    label = f"{kind}: {name}"
    if minimum:
        label += f" >= {minimum}"
    if publisher:
        label += f" ({publisher})"
    return label


def _format_target_family(family: dict[str, str]) -> str:
    label = f"TargetDeviceFamily: {family.get('Name') or 'unknown'}"
    if family.get("MinVersion"):
        label += f" min {family['MinVersion']}"
    if family.get("MaxVersionTested"):
        label += f" tested {family['MaxVersionTested']}"
    return label


def _manifest_entry(
    archive: zipfile.ZipFile,
    candidates: tuple[str, ...],
) -> zipfile.ZipInfo:
    available = {
        info.filename.replace("\\", "/").lower(): info
        for info in archive.infolist()
    }
    for candidate in candidates:
        info = available.get(candidate.lower())
        if info is not None:
            if info.file_size > MAX_MANIFEST_BYTES:
                raise AppxInspectionError(
                    f"{candidate} exceeds the manifest inspection limit"
                )
            return info
    raise AppxInspectionError("AppX/MSIX manifest was not found")


def _parse_manifest(
    archive: zipfile.ZipFile,
    candidates: tuple[str, ...],
    source: str,
) -> tuple[ET.Element, str]:
    info = _manifest_entry(archive, candidates)
    try:
        root = ET.fromstring(archive.read(info))
    except (ET.ParseError, OSError, RuntimeError) as exc:
        raise AppxInspectionError(
            f"{source} contains an unreadable manifest"
        ) from exc
    return root, info.filename.replace("\\", "/")


def _safe_bundle_member(value: str) -> str:
    normalized = str(value or "").replace("\\", "/").strip()
    path = PurePosixPath(normalized)
    if (
        not normalized
        or path.is_absolute()
        or ".." in path.parts
        or ":" in normalized
    ):
        raise AppxInspectionError(
            f"Bundle contains an unsafe package member {value!r}"
        )
    return str(path)


def _nested_archive(
    archive: zipfile.ZipFile,
    info: zipfile.ZipInfo,
) -> zipfile.ZipFile:
    if info.file_size <= 0 or info.file_size > MAX_NESTED_PACKAGE_BYTES:
        raise AppxInspectionError(
            f"Bundle member {info.filename!r} has an invalid size"
        )
    temporary = tempfile.SpooledTemporaryFile(max_size=16 * 1024 * 1024)
    with archive.open(info) as source:
        shutil.copyfileobj(source, temporary, length=1024 * 1024)
    temporary.seek(0)
    try:
        nested = zipfile.ZipFile(temporary)
        nested._msstorehelper_temporary = temporary
        return nested
    except zipfile.BadZipFile as exc:
        temporary.close()
        raise AppxInspectionError(
            f"Bundle member {info.filename!r} is not an AppX/MSIX archive"
        ) from exc


def _inspect_bundle(
    archive: zipfile.ZipFile,
    root: ET.Element,
    *,
    package_path: str,
    manifest_name: str,
) -> dict[str, Any]:
    outer_identity = _identity(root, manifest_name)
    available = {
        info.filename.replace("\\", "/"): info
        for info in archive.infolist()
    }
    descriptors = [
        element
        for element in root.iter()
        if _local_name(element.tag) == "Package"
        and element.attrib.get("FileName")
    ]
    if not descriptors:
        raise AppxInspectionError(
            "Bundle manifest does not declare any inner packages"
        )

    inner_packages = []
    for descriptor in descriptors:
        member_name = _safe_bundle_member(descriptor.attrib["FileName"])
        info = available.get(member_name)
        if info is None:
            raise AppxInspectionError(
                f"Bundle member is missing: {member_name}"
            )
        nested = _nested_archive(archive, info)
        try:
            inner_root, inner_manifest = _parse_manifest(
                nested,
                ("AppxManifest.xml",),
                member_name,
            )
            details = _manifest_details(
                inner_root,
                source=member_name,
                manifest_name=f"{member_name}!/{inner_manifest}",
            )
        finally:
            temporary = getattr(
                nested,
                "_msstorehelper_temporary",
                None,
            )
            nested.close()
            if temporary is not None:
                temporary.close()

        descriptor_arch = str(
            descriptor.attrib.get("Architecture")
            or descriptor.attrib.get("ProcessorArchitecture")
            or "neutral"
        ).strip().lower()
        inner_arch = details["Identity"]["ProcessorArchitecture"]
        if descriptor_arch not in PACKAGE_ARCHITECTURES:
            raise AppxInspectionError(
                f"Bundle member {member_name!r} has an invalid architecture"
            )
        if (
            descriptor_arch not in {"neutral", "resource"}
            and inner_arch not in {"neutral", "resource"}
            and descriptor_arch != inner_arch
        ):
            raise AppxInspectionError(
                f"Bundle member architecture disagrees with {member_name}"
            )
        resources = []
        for child in descriptor.iter():
            if _local_name(child.tag) != "Resource":
                continue
            resources.append({
                str(key): str(value).strip()
                for key, value in child.attrib.items()
                if str(value).strip()
            })
        details.update({
            "BundleFileName": member_name,
            "BundlePackageType": str(
                descriptor.attrib.get("Type") or "application"
            ).strip().lower(),
            "BundleArchitecture": descriptor_arch,
            "BundleResourceId": str(
                descriptor.attrib.get("ResourceId") or ""
            ).strip(),
            "BundleResources": _canonical_dicts(resources),
        })
        inner_packages.append(details)

    package_dependencies = _canonical_dicts([
        dependency
        for package in inner_packages
        for dependency in package["PackageDependencies"]
    ])
    main_dependencies = _canonical_dicts([
        dependency
        for package in inner_packages
        for dependency in package["MainPackageDependencies"]
    ])
    target_families = _canonical_dicts([
        family
        for package in inner_packages
        for family in package["TargetDeviceFamilies"]
    ])
    capabilities = _canonical_dicts([
        capability
        for package in inner_packages
        for capability in package["CapabilitiesDetailed"]
    ])
    minimum_versions = [
        _version_tuple(package["MinOSVersion"])
        for package in inner_packages
        if _version_tuple(package["MinOSVersion"])
    ]
    return {
        "SchemaVersion": INSPECTION_SCHEMA_VERSION,
        "Path": os.path.abspath(package_path),
        "ContainerType": "bundle",
        "ManifestName": manifest_name,
        "Identity": outer_identity,
        "Properties": {
            "Framework": False,
            "ResourcePackage": False,
            "AllowExecution": True,
        },
        "Architectures": sorted({
            package["BundleArchitecture"]
            for package in inner_packages
        }),
        "MinOSVersion": (
            ".".join(str(piece) for piece in max(minimum_versions))
            if minimum_versions
            else ""
        ),
        "TargetDeviceFamilies": target_families,
        "PackageDependencies": package_dependencies,
        "MainPackageDependencies": main_dependencies,
        "CapabilitiesDetailed": capabilities,
        "Capabilities": [
            f"{item['Type']}: {item['Name']}"
            for item in capabilities
        ],
        "Dependencies": sorted(
            [
                _format_dependency(item)
                for item in package_dependencies + main_dependencies
            ]
            + [
                _format_target_family(item)
                for item in target_families
            ]
        ),
        "Applications": [
            application
            for package in inner_packages
            for application in package["Applications"]
        ],
        "InnerPackages": inner_packages,
    }


def inspect_appx_archive(package_path: str) -> dict[str, Any]:
    """Inspect an outer package and every package manifest inside a bundle."""
    package_path = os.path.abspath(package_path)
    try:
        with zipfile.ZipFile(package_path) as archive:
            root, manifest_name = _parse_manifest(
                archive,
                (
                    "AppxManifest.xml",
                    "AppxMetadata/AppxBundleManifest.xml",
                ),
                package_path,
            )
            if _local_name(root.tag) == "Bundle":
                return _inspect_bundle(
                    archive,
                    root,
                    package_path=package_path,
                    manifest_name=manifest_name,
                )
            details = _manifest_details(
                root,
                source=package_path,
                manifest_name=manifest_name,
            )
    except (OSError, zipfile.BadZipFile) as exc:
        raise AppxInspectionError(
            f"Package is not a readable AppX/MSIX archive: {package_path}"
        ) from exc

    architecture = details["Identity"]["ProcessorArchitecture"]
    details.update({
        "SchemaVersion": INSPECTION_SCHEMA_VERSION,
        "Path": package_path,
        "ContainerType": "package",
        "Architectures": [architecture],
        "InnerPackages": [],
    })
    return details


def inspection_role(inspection: dict[str, Any]) -> str:
    """Classify deployment role from signed manifest content."""
    if inspection.get("MainPackageDependencies"):
        return "optional"
    properties = inspection.get("Properties") or {}
    if properties.get("Framework"):
        return "dependency"
    if properties.get("ResourcePackage"):
        return "resource"
    if inspection.get("ContainerType") == "bundle":
        inner = inspection.get("InnerPackages") or []
        application_packages = [
            package
            for package in inner
            if package.get("BundlePackageType") == "application"
        ]
        if application_packages and all(
            package.get("MainPackageDependencies")
            for package in application_packages
        ):
            return "optional"
        if not application_packages:
            return "resource"
    return "app"


def inspection_supports_architecture(
    inspection: dict[str, Any],
    target_architecture: str,
    *,
    dependency: bool = False,
) -> bool:
    target = str(target_architecture or "").strip().lower()
    architectures = {
        str(value).strip().lower()
        for value in inspection.get("Architectures") or []
    }
    if inspection.get("ContainerType") == "bundle":
        application_arches = {
            str(package.get("BundleArchitecture") or "").lower()
            for package in inspection.get("InnerPackages") or []
            if package.get("BundlePackageType") == "application"
        }
        architectures = application_arches or architectures
    accepted = {target, "neutral", "resource"}
    if dependency and target == "x64":
        accepted.add("x86")
    if dependency and target == "arm64":
        accepted.update({"arm", "x86"})
    return bool(architectures & accepted)


def _inventory_records(
    inventory: dict[str, Any] | None,
) -> tuple[list[dict[str, Any]], str]:
    if inventory is None:
        return [], "not-provided"
    status = str(inventory.get("Status") or "unavailable")
    if status not in KNOWN_INVENTORY_STATUSES:
        raise InstallPlanError(
            "Installed package inventory is not authoritative "
            f"(status: {status})"
        )
    records = [
        dict(record)
        for record in inventory.get("Records") or []
        if isinstance(record, dict)
    ]
    if not records:
        for name, version in (inventory.get("Versions") or {}).items():
            records.append({
                "Name": str(name),
                "Version": str(version),
                "Source": "installed",
            })
    return records, status


def _installed_matches(
    records: list[dict[str, Any]],
    identity: dict[str, str],
    architectures: set[str],
) -> list[dict[str, Any]]:
    name = identity["Name"].lower()
    matches = []
    for record in records:
        if str(record.get("Name") or "").strip().lower() != name:
            continue
        publisher = str(record.get("Publisher") or "").strip()
        if publisher and publisher != identity["Publisher"]:
            continue
        architecture = str(
            record.get("Architecture") or ""
        ).strip().lower()
        if (
            architecture
            and architecture not in {"neutral", "resource"}
            and architecture not in architectures
        ):
            continue
        matches.append(record)
    return matches


def _plan_artifact(
    package: dict[str, Any],
    inspection: dict[str, Any],
    role: str,
) -> dict[str, Any]:
    return {
        "FileName": str(
            package.get("FileName")
            or os.path.basename(inspection["Path"])
        ),
        "Path": inspection["Path"],
        "Role": role,
        "Identity": dict(inspection["Identity"]),
        "ContainerType": inspection["ContainerType"],
        "Architectures": list(inspection["Architectures"]),
        "MinOSVersion": inspection.get("MinOSVersion") or "",
        "PackageDependencies": list(
            inspection.get("PackageDependencies") or []
        ),
        "MainPackageDependencies": list(
            inspection.get("MainPackageDependencies") or []
        ),
        "Capabilities": list(inspection.get("Capabilities") or []),
        "InnerPackages": list(inspection.get("InnerPackages") or []),
        "StoreQuery": dict(package.get("StoreQuery") or {}),
        "Action": "install",
    }


def _requirement_satisfied(
    requirement: dict[str, str],
    artifacts: list[dict[str, Any]],
    installed: list[dict[str, Any]],
) -> bool:
    name = str(requirement.get("Name") or "").strip().lower()
    minimum = str(requirement.get("MinVersion") or "").strip()
    publisher = str(requirement.get("Publisher") or "").strip()
    for artifact in artifacts:
        identity = artifact["Identity"]
        if identity["Name"].lower() != name:
            continue
        if publisher and identity["Publisher"] != publisher:
            continue
        if (
            not minimum
            or _compare_versions(identity["Version"], minimum) >= 0
        ):
            return True
    for record in installed:
        if str(record.get("Name") or "").strip().lower() != name:
            continue
        record_publisher = str(record.get("Publisher") or "").strip()
        if publisher and record_publisher and record_publisher != publisher:
            continue
        if (
            not minimum
            or _compare_versions(record.get("Version"), minimum) >= 0
        ):
            return True
    return False


def build_install_plan(
    packages: list[dict[str, Any]],
    *,
    target_architecture: str,
    inventory: dict[str, Any] | None = None,
    target_os_version: str = "",
    allow_downgrade: bool = False,
) -> dict[str, Any]:
    """Build one deterministic, manifest-grounded deployment plan."""
    target_architecture = str(target_architecture or "").strip().lower()
    if target_architecture not in PACKAGE_ARCHITECTURES - {"resource"}:
        raise InstallPlanError(
            f"Unsupported target architecture: {target_architecture}"
        )
    installed, inventory_status = _inventory_records(inventory)
    inspected = []
    seen_paths = set()
    for package in packages:
        raw_path = str(package.get("Path") or "").strip()
        if not raw_path:
            raise InstallPlanError("Install-plan package path is missing")
        path = os.path.abspath(raw_path)
        if path in seen_paths:
            raise InstallPlanError(
                f"Duplicate package path in queue: {path}"
            )
        seen_paths.add(path)
        inspection = (
            dict(package["Inspection"])
            if isinstance(package.get("Inspection"), dict)
            else inspect_appx_archive(path)
        )
        role = inspection_role(inspection)
        if not inspection_supports_architecture(
            inspection,
            target_architecture,
            dependency=role in {"dependency", "resource"},
        ):
            raise InstallPlanError(
                f"{os.path.basename(path)} does not support "
                f"{target_architecture}"
            )
        inspected.append(
            _plan_artifact(package, inspection, role)
        )

    mains = [item for item in inspected if item["Role"] == "app"]
    main_names = {
        item["Identity"]["Name"].lower()
        for item in mains
    }
    if not mains:
        raise InstallPlanError(
            "Install plan requires one main app package or bundle"
        )
    if len(mains) != 1 or len(main_names) != 1:
        names = ", ".join(sorted(
            item["Identity"]["Name"] for item in mains
        ))
        raise InstallPlanError(
            "Queue contains multiple independent main apps "
            f"({names}); split it into one plan per app"
        )
    main = mains[0]
    main_name = main["Identity"]["Name"]

    optional = [
        item for item in inspected if item["Role"] == "optional"
    ]
    for item in optional:
        linked_names = {
            str(dependency.get("Name") or "").strip().lower()
            for dependency in item["MainPackageDependencies"]
        }
        if main_name.lower() not in linked_names:
            raise InstallPlanError(
                f"Optional package {item['Identity']['Name']} is not "
                f"linked to main package {main_name}"
            )

    resources = [
        item for item in inspected if item["Role"] == "resource"
    ]
    dependencies = [
        item for item in inspected if item["Role"] == "dependency"
    ]
    requirements = _canonical_dicts([
        requirement
        for item in [main] + optional + dependencies
        for requirement in item["PackageDependencies"]
        if requirement.get("Type") == "PackageDependency"
    ])
    relevant_dependency_names = {
        str(requirement.get("Name") or "").strip().lower()
        for requirement in requirements
    }
    unused_dependencies = [
        item
        for item in dependencies
        if item["Identity"]["Name"].lower()
        not in relevant_dependency_names
    ]
    dependencies = [
        item
        for item in dependencies
        if item not in unused_dependencies
    ]

    conflicts = []
    warnings = []
    if inventory_status == "not-provided":
        warnings.append(
            "Installed-state conflicts were not evaluated because no "
            "authoritative inventory was supplied."
        )
    if unused_dependencies:
        warnings.append(
            "Ignored unreferenced dependency artifacts: "
            + ", ".join(sorted(
                item["FileName"] for item in unused_dependencies
            ))
        )

    planned_for_requirements = dependencies + resources
    for requirement in requirements:
        if _requirement_satisfied(
            requirement,
            planned_for_requirements,
            installed,
        ):
            continue
        conflicts.append({
            "Code": "missing-dependency",
            "Package": requirement.get("Name") or "unknown",
            "Message": (
                f"Required dependency {requirement.get('Name') or 'unknown'} "
                f"{requirement.get('MinVersion') or ''} is neither queued "
                "nor installed at a sufficient version."
            ).replace("  ", " "),
            "Blocking": True,
        })

    target_os = _version_tuple(target_os_version)
    required_os = _version_tuple(main.get("MinOSVersion"))
    if target_os and required_os and _compare_versions(
        target_os_version,
        main["MinOSVersion"],
    ) < 0:
        conflicts.append({
            "Code": "minimum-os",
            "Package": main_name,
            "Message": (
                f"{main_name} requires Windows {main['MinOSVersion']} "
                f"but the target reports {target_os_version}."
            ),
            "Blocking": True,
        })

    all_packages = dependencies + resources + optional + [main]
    for artifact in all_packages:
        identity = artifact["Identity"]
        architectures = {
            str(value).lower()
            for value in artifact["Architectures"]
        }
        matches = _installed_matches(
            installed,
            identity,
            architectures,
        )
        newest = max(
            matches,
            key=lambda record: _version_tuple(record.get("Version")),
            default=None,
        )
        if newest is None:
            continue
        installed_version = str(newest.get("Version") or "")
        comparison = _compare_versions(
            installed_version,
            identity["Version"],
        )
        artifact["InstalledVersion"] = installed_version
        if comparison >= 0:
            artifact["Action"] = "skip"
        if (
            artifact is main
            and comparison > 0
            and not allow_downgrade
        ):
            conflicts.append({
                "Code": "downgrade",
                "Package": identity["Name"],
                "Message": (
                    f"Installed {identity['Name']} {installed_version} "
                    f"is newer than queued {identity['Version']}."
                ),
                "Blocking": True,
            })

    plan = {
        "SchemaVersion": PLAN_SCHEMA_VERSION,
        "TargetArchitecture": target_architecture,
        "TargetOSVersion": str(target_os_version or ""),
        "InventoryStatus": inventory_status,
        "Main": main,
        "Dependencies": dependencies,
        "ResourcePackages": resources,
        "OptionalPackages": optional,
        "UnusedPackages": unused_dependencies,
        "Requirements": requirements,
        "Conflicts": conflicts,
        "Warnings": warnings,
        "Installable": not any(
            conflict.get("Blocking") for conflict in conflicts
        ),
        "Deployment": {
            "Command": "Add-AppxPackage",
            "MainPath": main["Path"],
            "DependencyPaths": [
                item["Path"]
                for item in dependencies + resources
                if item["Action"] == "install"
            ],
            "ExternalPackagePaths": [
                item["Path"]
                for item in optional
                if item["Action"] == "install"
            ],
        },
    }
    validate_install_plan(plan)
    return plan


def validate_install_plan(plan: dict[str, Any]) -> dict[str, Any]:
    """Validate a round-tripped v1 plan before execution or export."""
    if not isinstance(plan, dict):
        raise InstallPlanError("Install plan must be an object")
    if plan.get("SchemaVersion") != PLAN_SCHEMA_VERSION:
        raise InstallPlanError("Install plan schema is unsupported")
    main = plan.get("Main")
    if not isinstance(main, dict) or not main.get("Path"):
        raise InstallPlanError("Install plan main package is missing")
    deployment = plan.get("Deployment")
    if not isinstance(deployment, dict):
        raise InstallPlanError("Install plan deployment is missing")
    if deployment.get("Command") != "Add-AppxPackage":
        raise InstallPlanError("Install plan command is unsupported")
    if os.path.abspath(str(deployment.get("MainPath") or "")) != os.path.abspath(
        str(main["Path"])
    ):
        raise InstallPlanError("Install plan main path is inconsistent")
    allowed_paths = {
        os.path.abspath(str(item.get("Path") or ""))
        for key in (
            "Dependencies",
            "ResourcePackages",
            "OptionalPackages",
        )
        for item in plan.get(key) or []
        if isinstance(item, dict) and item.get("Path")
    }
    supplied_paths = {
        os.path.abspath(str(path))
        for key in ("DependencyPaths", "ExternalPackagePaths")
        for path in deployment.get(key) or []
    }
    if not supplied_paths <= allowed_paths:
        raise InstallPlanError(
            "Install plan deployment references an unknown package"
        )
    json.loads(json.dumps(plan, sort_keys=True))
    return plan


def render_install_plan(plan: dict[str, Any]) -> str:
    validate_install_plan(plan)
    main = plan["Main"]
    lines = [
        f"Plan schema: {plan['SchemaVersion']}",
        (
            f"Main: {main['Identity']['Name']} "
            f"{main['Identity']['Version']} "
            f"[{','.join(main['Architectures'])}] "
            f"({main['Action']})"
        ),
        f"Target: {plan['TargetArchitecture']}"
        + (
            f" / Windows {plan['TargetOSVersion']}"
            if plan.get("TargetOSVersion")
            else ""
        ),
    ]
    for label, key in (
        ("Dependencies", "Dependencies"),
        ("Resources", "ResourcePackages"),
        ("Optional", "OptionalPackages"),
    ):
        values = plan.get(key) or []
        lines.append(f"{label}: {len(values)}")
        for item in values:
            identity = item["Identity"]
            lines.append(
                f"  - {identity['Name']} {identity['Version']} "
                f"[{','.join(item['Architectures'])}] "
                f"({item['Action']})"
            )
    if plan.get("Conflicts"):
        lines.append("Conflicts:")
        for conflict in plan["Conflicts"]:
            lines.append(f"  - {conflict['Message']}")
    if plan.get("Warnings"):
        lines.append("Warnings:")
        for warning in plan["Warnings"]:
            lines.append(f"  - {warning}")
    lines.append(
        "Result: "
        + ("ready" if plan.get("Installable") else "blocked")
    )
    return "\n".join(lines)
