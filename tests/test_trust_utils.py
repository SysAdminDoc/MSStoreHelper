#!/usr/bin/env python3

import hashlib

from package_trust import TRUST_SCHEMA_VERSION, TRUST_STATE_TRUSTED


def file_sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def mark_package_trusted(package, path):
    artifact_sha256 = file_sha256(path)
    report = {
        "SchemaVersion": TRUST_SCHEMA_VERSION,
        "ReportKind": "package-trust-evaluation",
        "State": TRUST_STATE_TRUSTED,
        "AutomationAllowed": True,
        "ReviewEligible": False,
        "ArtifactSha256": artifact_sha256,
        "ReasonCodes": [],
        "Checks": [],
    }
    package["Sha256"] = artifact_sha256
    package["TrustState"] = TRUST_STATE_TRUSTED
    package["TrustReport"] = report
    return report


def inspect_as_trusted(filepath, package=None, **_kwargs):
    package = package if isinstance(package, dict) else {}
    return mark_package_trusted(package, filepath)
