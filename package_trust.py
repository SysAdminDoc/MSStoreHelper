#!/usr/bin/env python3
"""Package signature, manifest, identity, and promotion trust policy."""

import hashlib
import ntpath
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from urllib.parse import urlsplit, urlunsplit

from package_ingress import (
    INSTALLABLE_PACKAGE_EXTENSIONS,
    validate_existing_package_path,
    validate_package_filename,
)


TRUST_SCHEMA_VERSION = 1
TRUST_STATE_TRUSTED = "trusted"
TRUST_STATE_REVIEW_REQUIRED = "review-required"
TRUST_STATE_REVIEWED = "reviewed"
TRUST_STATE_BLOCKED = "blocked"
AUTOMATION_TRUST_STATES = frozenset({
    TRUST_STATE_TRUSTED,
    TRUST_STATE_REVIEWED,
})
PACKAGE_ARCHITECTURES = frozenset({
    "arm",
    "arm64",
    "neutral",
    "resource",
    "x64",
    "x86",
})
PUBLISHER_ID_ALPHABET = "0123456789abcdefghjkmnpqrstvwxyz"
MAX_MANIFEST_BYTES = 2 * 1024 * 1024


class PackageTrustError(ValueError):
    """Raised when a package cannot be inspected or promoted safely."""


def utc_timestamp(value=None):
    value = value or datetime.now(timezone.utc)
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).isoformat()


def publisher_id_from_subject(publisher):
    """Calculate the Windows package publisher ID from an X.500 subject."""
    if not isinstance(publisher, str) or not publisher.strip():
        raise PackageTrustError("Manifest publisher is missing")
    digest = hashlib.sha256(publisher.encode("utf-16le")).digest()[:8]
    bits = "".join(f"{byte:08b}" for byte in digest) + "0"
    return "".join(
        PUBLISHER_ID_ALPHABET[int(bits[offset:offset + 5], 2)]
        for offset in range(0, 65, 5)
    )


def package_filename_metadata(filename):
    """Parse identity, version, architecture, type, and PFN from a Store name."""
    filename = validate_package_filename(filename)
    stem, extension = ntpath.splitext(filename)
    identity = stem.split("_", 1)[0]
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9.-]{0,127}", identity):
        raise PackageTrustError("Package filename identity is malformed")

    version = ""
    architecture = "neutral"
    for piece in stem.split("_")[1:]:
        if not version and re.fullmatch(r"\d+(?:\.\d+){3}", piece):
            version = piece
        if piece.lower() in PACKAGE_ARCHITECTURES:
            architecture = piece.lower()

    publisher_id = ""
    if "__" in stem:
        publisher_id = stem.rsplit("__", 1)[1].lower()
    if not re.fullmatch(r"[0-9abcdefghjkmnpqrstvwxyz]{13}", publisher_id):
        raise PackageTrustError("Package filename publisher ID is missing or malformed")

    return {
        "FileName": filename,
        "Identity": identity,
        "Version": version,
        "Architecture": architecture,
        "Extension": extension.lower(),
        "IsBundle": extension.lower() in {".appxbundle", ".msixbundle"},
        "PublisherId": publisher_id,
        "PackageFamilyName": f"{identity}_{publisher_id}",
    }


def _xml_local_name(tag):
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _archive_manifest_entry(archive, expected_bundle):
    names = {}
    for info in archive.infolist():
        normalized = info.filename.replace("\\", "/").lower()
        if normalized in names:
            raise PackageTrustError("Package archive contains duplicate manifest paths")
        names[normalized] = info

    single = names.get("appxmanifest.xml")
    bundle = names.get("appxmetadata/appxbundlemanifest.xml")
    if single and bundle:
        raise PackageTrustError("Package archive contains ambiguous manifests")
    selected = bundle if expected_bundle else single
    if selected is None:
        expected = "bundle" if expected_bundle else "package"
        raise PackageTrustError(f"Package archive has no {expected} manifest")
    if selected.file_size <= 0 or selected.file_size > MAX_MANIFEST_BYTES:
        raise PackageTrustError("Package manifest size is invalid")
    return selected


def read_package_manifest(package_path, filename=None):
    """Read bounded identity details from a package or bundle manifest."""
    filename = validate_package_filename(
        filename or os.path.basename(os.path.abspath(package_path))
    )
    package_path = validate_existing_package_path(
        package_path,
        expected_filename=filename,
        require_file=True,
    )
    file_metadata = package_filename_metadata(filename)

    try:
        with zipfile.ZipFile(package_path) as archive:
            manifest_info = _archive_manifest_entry(
                archive,
                file_metadata["IsBundle"],
            )
            root = ET.fromstring(archive.read(manifest_info))
    except (OSError, ET.ParseError, zipfile.BadZipFile) as exc:
        raise PackageTrustError("Package is not a readable signed archive") from exc

    identities = [
        element
        for element in root.iter()
        if _xml_local_name(element.tag) == "Identity"
    ]
    if len(identities) != 1:
        raise PackageTrustError("Package manifest must contain one identity")
    identity = identities[0]

    name = identity.attrib.get("Name", "").strip()
    publisher = identity.attrib.get("Publisher", "").strip()
    version = identity.attrib.get("Version", "").strip()
    architecture = (
        identity.attrib.get("ProcessorArchitecture")
        or identity.attrib.get("Architecture")
        or "neutral"
    ).strip().lower()
    if not name or not publisher or not re.fullmatch(r"\d+(?:\.\d+){3}", version):
        raise PackageTrustError("Package manifest identity is incomplete")
    if architecture not in PACKAGE_ARCHITECTURES:
        raise PackageTrustError("Package manifest architecture is invalid")

    bundle_architectures = set()
    if file_metadata["IsBundle"]:
        for element in root.iter():
            if _xml_local_name(element.tag) != "Package":
                continue
            package_arch = (
                element.attrib.get("Architecture")
                or element.attrib.get("ProcessorArchitecture")
                or ""
            ).strip().lower()
            if package_arch:
                if package_arch not in PACKAGE_ARCHITECTURES:
                    raise PackageTrustError(
                        "Bundle manifest contains an invalid architecture"
                    )
                bundle_architectures.add(package_arch)

    publisher_id = publisher_id_from_subject(publisher)
    return {
        "ManifestPath": manifest_info.filename.replace("\\", "/"),
        "PackageType": "bundle" if file_metadata["IsBundle"] else "package",
        "Identity": name,
        "Publisher": publisher,
        "PublisherId": publisher_id,
        "PackageFamilyName": f"{name}_{publisher_id}",
        "Version": version,
        "Architecture": architecture,
        "BundleArchitectures": sorted(bundle_architectures),
    }


def source_url_summary(url):
    """Remove credentials, queries, and fragments from review provenance."""
    if not url:
        return ""
    try:
        parsed = urlsplit(str(url))
    except ValueError:
        return ""
    if parsed.scheme.lower() not in {"http", "https"} or not parsed.hostname:
        return ""
    host = parsed.hostname
    if parsed.port:
        host = f"{host}:{parsed.port}"
    return urlunsplit((parsed.scheme.lower(), host, parsed.path, "", ""))


def _check(checks, name, passed, detail, *, warning=False):
    checks.append({
        "Name": name,
        "Status": "warning" if warning else ("pass" if passed else "fail"),
        "Detail": str(detail),
    })
    return passed


def normalize_chain_status(value):
    """Normalize PowerShell's scalar-or-array JSON shape."""
    if value is None:
        return []
    values = value if isinstance(value, (list, tuple, set)) else [value]
    return [
        str(item).strip()
        for item in values
        if str(item).strip()
    ]


def evaluate_package_trust(
    package,
    artifact_sha256,
    signature_info,
    manifest,
    *,
    evaluated_at=None,
):
    """Evaluate immutable evidence and return a promotion trust report."""
    if not isinstance(package, dict):
        raise PackageTrustError("Package metadata is missing")
    file_metadata = package_filename_metadata(package.get("FileName"))
    if not isinstance(signature_info, dict):
        signature_info = {}
    if not isinstance(manifest, dict):
        raise PackageTrustError("Package manifest evidence is missing")

    checks = []
    failure_codes = []

    def required(name, condition, detail, code):
        if not _check(checks, name, bool(condition), detail):
            failure_codes.append(code)

    status = str(signature_info.get("Status", "")).strip()
    status_valid = status.lower() in {"valid", "0"}
    required(
        "authenticode-status",
        status_valid,
        status or "missing",
        "signature-status-invalid",
    )
    required(
        "windows-chain",
        signature_info.get("ChainValid") is True,
        "; ".join(normalize_chain_status(
            signature_info.get("ChainStatus")
        )) or "Windows chain result",
        "signature-chain-invalid",
    )

    signer_thumbprint = str(signature_info.get("SignerThumbprint", "")).strip()
    root_thumbprint = str(signature_info.get("RootThumbprint", "")).strip()
    required(
        "signer-certificate",
        bool(re.fullmatch(r"[0-9A-Fa-f]{40,128}", signer_thumbprint)),
        signer_thumbprint or "missing",
        "signer-certificate-missing",
    )
    required(
        "root-certificate",
        bool(re.fullmatch(r"[0-9A-Fa-f]{40,128}", root_thumbprint)),
        root_thumbprint or "missing",
        "root-certificate-missing",
    )

    revocation_state = str(
        signature_info.get("RevocationState", "")
    ).strip().lower()
    required(
        "revocation",
        revocation_state in {"checked", "offline"},
        revocation_state or "missing",
        "revocation-failed",
    )
    if revocation_state == "offline":
        _check(
            checks,
            "revocation-offline",
            True,
            "Online revocation was indeterminate; the Windows chain passed offline",
            warning=True,
        )

    signer = str(signature_info.get("Signer", "")).strip()
    manifest_publisher = str(manifest.get("Publisher", "")).strip()
    required(
        "signed-manifest-publisher",
        bool(signer) and signer == manifest_publisher,
        f"signer={signer or 'missing'}; manifest={manifest_publisher or 'missing'}",
        "publisher-mismatch",
    )

    manifest_identity = str(manifest.get("Identity", "")).strip()
    required(
        "filename-manifest-identity",
        manifest_identity == file_metadata["Identity"],
        f"filename={file_metadata['Identity']}; manifest={manifest_identity or 'missing'}",
        "identity-mismatch",
    )
    manifest_version = str(manifest.get("Version", "")).strip()
    required(
        "filename-manifest-version",
        bool(file_metadata["Version"])
        and manifest_version == file_metadata["Version"],
        f"filename={file_metadata['Version'] or 'missing'}; manifest={manifest_version or 'missing'}",
        "version-mismatch",
    )
    manifest_publisher_id = str(manifest.get("PublisherId", "")).strip().lower()
    required(
        "filename-manifest-publisher-id",
        manifest_publisher_id == file_metadata["PublisherId"],
        f"filename={file_metadata['PublisherId']}; manifest={manifest_publisher_id or 'missing'}",
        "publisher-id-mismatch",
    )

    expected_type = "bundle" if file_metadata["IsBundle"] else "package"
    manifest_type = str(manifest.get("PackageType", "")).strip().lower()
    required(
        "package-type",
        file_metadata["Extension"] in INSTALLABLE_PACKAGE_EXTENSIONS
        and manifest_type == expected_type
        and not package.get("IsEncrypted", False),
        f"extension={file_metadata['Extension']}; manifest={manifest_type or 'missing'}",
        "package-type-mismatch",
    )

    declared_file_type = str(package.get("FileType", "")).strip().lower()
    if declared_file_type:
        required(
            "source-file-type",
            f".{declared_file_type}" == file_metadata["Extension"],
            f"source={declared_file_type}; filename={file_metadata['Extension']}",
            "source-type-mismatch",
        )

    source_architecture = str(
        package.get("Architecture") or file_metadata["Architecture"]
    ).strip().lower()
    manifest_architecture = str(
        manifest.get("Architecture") or "neutral"
    ).strip().lower()
    bundle_architectures = {
        str(value).lower()
        for value in manifest.get("BundleArchitectures", [])
    }
    architecture_valid = source_architecture in PACKAGE_ARCHITECTURES
    if file_metadata["IsBundle"]:
        if source_architecture not in {"neutral", "resource"} and bundle_architectures:
            architecture_valid = (
                architecture_valid
                and source_architecture in bundle_architectures
            )
    else:
        architecture_valid = (
            architecture_valid
            and source_architecture == file_metadata["Architecture"]
            and manifest_architecture == file_metadata["Architecture"]
        )
    required(
        "package-architecture",
        architecture_valid,
        (
            f"source={source_architecture}; filename={file_metadata['Architecture']}; "
            f"manifest={manifest_architecture}; bundle={','.join(sorted(bundle_architectures))}"
        ),
        "architecture-mismatch",
    )

    expected_pfn = str(
        package.get("ExpectedPackageFamilyName")
        or file_metadata["PackageFamilyName"]
    ).strip()
    manifest_pfn = str(manifest.get("PackageFamilyName", "")).strip()
    required(
        "expected-package-family",
        bool(expected_pfn) and manifest_pfn.lower() == expected_pfn.lower(),
        f"expected={expected_pfn or 'missing'}; manifest={manifest_pfn or 'missing'}",
        "package-family-mismatch",
    )

    expected_identity = str(
        package.get("ExpectedPackageIdentity") or ""
    ).strip()
    if expected_identity:
        required(
            "expected-package-identity",
            manifest_identity.lower() == expected_identity.lower(),
            f"expected={expected_identity}; manifest={manifest_identity or 'missing'}",
            "expected-identity-mismatch",
        )

    store_query = package.get("StoreQuery") or {}
    expected_product_id = str(
        package.get("ExpectedProductId")
        or store_query.get("ProductId")
        or ""
    ).strip()
    dependency_binding = package.get("ExpectedDependency") is True
    binding_missing = False
    if not expected_product_id:
        binding_missing = True
        _check(
            checks,
            "expected-product-binding",
            True,
            "No Store product ID was supplied",
            warning=True,
        )
    elif expected_identity:
        _check(
            checks,
            "expected-product-binding",
            True,
            f"{expected_product_id} is bound to {expected_identity}",
        )
    elif dependency_binding:
        _check(
            checks,
            "expected-product-binding",
            True,
            f"{expected_product_id} supplied this dependency",
        )
    else:
        binding_missing = True
        _check(
            checks,
            "expected-product-binding",
            True,
            f"{expected_product_id} has no authoritative identity mapping",
            warning=True,
        )

    required(
        "artifact-hash",
        bool(re.fullmatch(r"[0-9A-Fa-f]{64}", str(artifact_sha256 or ""))),
        str(artifact_sha256 or "missing"),
        "artifact-hash-invalid",
    )

    if failure_codes:
        state = TRUST_STATE_BLOCKED
    elif binding_missing:
        state = TRUST_STATE_REVIEW_REQUIRED
    else:
        state = TRUST_STATE_TRUSTED

    report = {
        "SchemaVersion": TRUST_SCHEMA_VERSION,
        "ReportKind": "package-trust-evaluation",
        "State": state,
        "AutomationAllowed": state in AUTOMATION_TRUST_STATES,
        "ReviewEligible": state == TRUST_STATE_REVIEW_REQUIRED,
        "EvaluatedAt": utc_timestamp(evaluated_at),
        "ArtifactSha256": str(artifact_sha256 or "").lower(),
        "Source": {
            "Url": source_url_summary(package.get("Url")),
            "ProductId": expected_product_id,
        },
        "Expected": {
            "FileName": file_metadata["FileName"],
            "Identity": expected_identity,
            "PackageFamilyName": expected_pfn,
            "Dependency": dependency_binding,
        },
        "Manifest": {
            "ManifestPath": manifest.get("ManifestPath", ""),
            "PackageType": manifest_type,
            "Identity": manifest_identity,
            "Publisher": manifest_publisher,
            "PublisherId": manifest_publisher_id,
            "PackageFamilyName": manifest_pfn,
            "Version": manifest_version,
            "Architecture": manifest_architecture,
            "BundleArchitectures": sorted(bundle_architectures),
        },
        "Signature": {
            "Status": status,
            "Signer": signer,
            "SignerThumbprint": signer_thumbprint,
            "Root": str(signature_info.get("Root", "")).strip(),
            "RootThumbprint": root_thumbprint,
            "ChainValid": signature_info.get("ChainValid") is True,
            "ChainStatus": normalize_chain_status(
                signature_info.get("ChainStatus")
            ),
            "RevocationState": revocation_state,
        },
        "Checks": checks,
        "ReasonCodes": failure_codes + (
            ["product-binding-review-required"] if binding_missing else []
        ),
    }
    return report


def blocked_trust_report(
    package,
    artifact_sha256,
    reason,
    *,
    signature_info=None,
    evaluated_at=None,
):
    """Create a non-reviewable report when inspection itself fails."""
    package = package if isinstance(package, dict) else {}
    signature_info = (
        signature_info if isinstance(signature_info, dict) else {}
    )
    return {
        "SchemaVersion": TRUST_SCHEMA_VERSION,
        "ReportKind": "package-trust-evaluation",
        "State": TRUST_STATE_BLOCKED,
        "AutomationAllowed": False,
        "ReviewEligible": False,
        "EvaluatedAt": utc_timestamp(evaluated_at),
        "ArtifactSha256": str(artifact_sha256 or "").lower(),
        "Source": {
            "Url": source_url_summary(package.get("Url")),
            "ProductId": str(
                package.get("ExpectedProductId")
                or (package.get("StoreQuery") or {}).get("ProductId")
                or ""
            ),
        },
        "Expected": {
            "FileName": str(package.get("FileName") or ""),
            "Identity": str(
                package.get("ExpectedPackageIdentity") or ""
            ),
            "PackageFamilyName": str(
                package.get("ExpectedPackageFamilyName") or ""
            ),
            "Dependency": package.get("ExpectedDependency") is True,
        },
        "Manifest": {},
        "Signature": {
            "Status": str(signature_info.get("Status", "")),
            "Signer": str(signature_info.get("Signer", "")),
            "SignerThumbprint": str(
                signature_info.get("SignerThumbprint", "")
            ),
            "Root": str(signature_info.get("Root", "")),
            "RootThumbprint": str(
                signature_info.get("RootThumbprint", "")
            ),
            "ChainValid": signature_info.get("ChainValid") is True,
            "ChainStatus": normalize_chain_status(
                signature_info.get("ChainStatus") or []
            ),
            "RevocationState": str(
                signature_info.get("RevocationState", "")
            ),
        },
        "Checks": [{
            "Name": "package-inspection",
            "Status": "fail",
            "Detail": str(reason or "Package inspection failed"),
        }],
        "ReasonCodes": ["package-inspection-failed"],
    }


def trust_report_allows_automation(report, artifact_sha256):
    """Return whether a report is current and authorizes automation."""
    if not isinstance(report, dict):
        return False
    if report.get("SchemaVersion") != TRUST_SCHEMA_VERSION:
        return False
    if report.get("ReportKind") != "package-trust-evaluation":
        return False
    if report.get("State") not in AUTOMATION_TRUST_STATES:
        return False
    if report.get("AutomationAllowed") is not True:
        return False
    return str(report.get("ArtifactSha256", "")).lower() == str(
        artifact_sha256 or ""
    ).lower()


def review_trust_report(
    report,
    artifact_sha256,
    *,
    reviewer="local-user",
    reviewed_at=None,
):
    """Promote only a review-eligible, hash-current report."""
    if not isinstance(report, dict):
        raise PackageTrustError("Package has no trust report")
    if report.get("State") != TRUST_STATE_REVIEW_REQUIRED:
        raise PackageTrustError("Package is not awaiting manual review")
    if report.get("ReviewEligible") is not True:
        raise PackageTrustError("Package trust failures are not reviewable")
    if str(report.get("ArtifactSha256", "")).lower() != str(
        artifact_sha256 or ""
    ).lower():
        raise PackageTrustError("Package changed after trust inspection")
    if any(check.get("Status") == "fail" for check in report.get("Checks", [])):
        raise PackageTrustError("Failed trust checks cannot be manually promoted")

    reviewed = {
        key: value.copy() if isinstance(value, dict) else list(value)
        if isinstance(value, list) else value
        for key, value in report.items()
    }
    reviewed["State"] = TRUST_STATE_REVIEWED
    reviewed["AutomationAllowed"] = True
    reviewed["ReviewEligible"] = False
    reviewed["Review"] = {
        "Reviewer": str(reviewer or "local-user"),
        "ReviewedAt": utc_timestamp(reviewed_at),
        "Decision": "promoted",
        "Acknowledged": [
            "Source",
            "Identity",
            "Publisher",
            "Manifest",
        ],
    }
    return reviewed
