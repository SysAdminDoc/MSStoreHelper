#!/usr/bin/env python3

import json
import os
import subprocess
import tempfile
import unittest
import zipfile
from datetime import datetime, timezone
from types import SimpleNamespace
from unittest.mock import patch

from MSStoreHelper import StoreAPI
from package_trust import (
    PackageTrustError,
    TRUST_STATE_BLOCKED,
    TRUST_STATE_REVIEWED,
    TRUST_STATE_REVIEW_REQUIRED,
    TRUST_STATE_TRUSTED,
    evaluate_package_trust,
    publisher_id_from_subject,
    read_package_manifest,
    review_trust_report,
    trust_report_allows_automation,
)


MICROSOFT_PUBLISHER = (
    "CN=Microsoft Corporation, O=Microsoft Corporation, "
    "L=Redmond, S=Washington, C=US"
)
PACKAGE_FILENAME = (
    "Microsoft.WindowsTerminal_1.0.0.0_x64__8wekyb3d8bbwe.msix"
)
ARTIFACT_SHA256 = "a" * 64


def package_metadata(**overrides):
    package = {
        "FileName": PACKAGE_FILENAME,
        "FileType": "MSIX",
        "Architecture": "x64",
        "Url": "https://cdn.test/package.msix?secret=redacted",
        "StoreQuery": {"ProductId": "9N0DX20HK701"},
        "ExpectedProductId": "9N0DX20HK701",
        "ExpectedPackageIdentity": "Microsoft.WindowsTerminal",
        "ExpectedPackageFamilyName": (
            "Microsoft.WindowsTerminal_8wekyb3d8bbwe"
        ),
    }
    package.update(overrides)
    return package


def signature_evidence(**overrides):
    evidence = {
        "Status": "Valid",
        "StatusMessage": "Signature verified",
        "Signer": MICROSOFT_PUBLISHER,
        "SignerThumbprint": "A" * 40,
        "Root": "CN=Microsoft Root Certificate Authority 2011",
        "RootThumbprint": "B" * 40,
        "ChainValid": True,
        "ChainStatus": [],
        "RevocationState": "checked",
    }
    evidence.update(overrides)
    return evidence


def manifest_evidence(**overrides):
    evidence = {
        "ManifestPath": "AppxManifest.xml",
        "PackageType": "package",
        "Identity": "Microsoft.WindowsTerminal",
        "Publisher": MICROSOFT_PUBLISHER,
        "PublisherId": "8wekyb3d8bbwe",
        "PackageFamilyName": (
            "Microsoft.WindowsTerminal_8wekyb3d8bbwe"
        ),
        "Version": "1.0.0.0",
        "Architecture": "x64",
        "BundleArchitectures": [],
    }
    evidence.update(overrides)
    return evidence


def write_appx(
    path,
    *,
    name="Microsoft.WindowsTerminal",
    publisher=MICROSOFT_PUBLISHER,
    version="1.0.0.0",
    architecture="x64",
):
    manifest = f"""<?xml version="1.0" encoding="utf-8"?>
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10">
  <Identity Name="{name}" Publisher="{publisher}" Version="{version}" ProcessorArchitecture="{architecture}" />
</Package>
"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("AppxManifest.xml", manifest)


class PackageTrustTests(unittest.TestCase):
    def test_publisher_id_matches_known_microsoft_package_family(self):
        self.assertEqual(
            publisher_id_from_subject(MICROSOFT_PUBLISHER),
            "8wekyb3d8bbwe",
        )

    def test_valid_signature_manifest_and_product_binding_are_trusted(self):
        report = evaluate_package_trust(
            package_metadata(),
            ARTIFACT_SHA256,
            signature_evidence(),
            manifest_evidence(),
            evaluated_at=datetime(2026, 7, 29, tzinfo=timezone.utc),
        )

        self.assertEqual(report["State"], TRUST_STATE_TRUSTED)
        self.assertTrue(report["AutomationAllowed"])
        self.assertEqual(report["Source"]["Url"], "https://cdn.test/package.msix")
        self.assertEqual(
            report["EvaluatedAt"],
            "2026-07-29T00:00:00+00:00",
        )
        self.assertTrue(
            trust_report_allows_automation(report, ARTIFACT_SHA256)
        )

    def test_failed_chain_and_misleading_root_are_blocked(self):
        report = evaluate_package_trust(
            package_metadata(),
            ARTIFACT_SHA256,
            signature_evidence(
                ChainValid=False,
                ChainStatus=["UntrustedRoot"],
                Root="CN=Not Microsoft Test Root",
            ),
            manifest_evidence(),
        )

        self.assertEqual(report["State"], TRUST_STATE_BLOCKED)
        self.assertIn("signature-chain-invalid", report["ReasonCodes"])
        self.assertFalse(report["AutomationAllowed"])

    def test_microsoft_substring_does_not_replace_exact_publisher_match(self):
        report = evaluate_package_trust(
            package_metadata(),
            ARTIFACT_SHA256,
            signature_evidence(Signer="CN=Definitely Microsoft-ish"),
            manifest_evidence(),
        )

        self.assertEqual(report["State"], TRUST_STATE_BLOCKED)
        self.assertIn("publisher-mismatch", report["ReasonCodes"])

    def test_identity_publisher_id_version_and_architecture_mismatches_block(self):
        cases = [
            (
                manifest_evidence(Identity="Contoso.App"),
                "identity-mismatch",
            ),
            (
                manifest_evidence(PublisherId="notpublisher1"),
                "publisher-id-mismatch",
            ),
            (
                manifest_evidence(Version="2.0.0.0"),
                "version-mismatch",
            ),
            (
                manifest_evidence(Architecture="arm64"),
                "architecture-mismatch",
            ),
            (
                manifest_evidence(PackageType="bundle"),
                "package-type-mismatch",
            ),
        ]

        for manifest, reason in cases:
            with self.subTest(reason=reason):
                report = evaluate_package_trust(
                    package_metadata(),
                    ARTIFACT_SHA256,
                    signature_evidence(),
                    manifest,
                )
                self.assertEqual(report["State"], TRUST_STATE_BLOCKED)
                self.assertIn(reason, report["ReasonCodes"])

    def test_offline_revocation_is_explicit_but_base_chain_can_promote(self):
        report = evaluate_package_trust(
            package_metadata(),
            ARTIFACT_SHA256,
            signature_evidence(
                RevocationState="offline",
                ChainStatus=[
                    "RevocationStatusUnknown",
                    "OfflineRevocation",
                ],
            ),
            manifest_evidence(),
        )

        self.assertEqual(report["State"], TRUST_STATE_TRUSTED)
        self.assertEqual(
            report["Signature"]["RevocationState"],
            "offline",
        )
        self.assertTrue(any(
            check["Name"] == "revocation-offline"
            and check["Status"] == "warning"
            for check in report["Checks"]
        ))

    def test_missing_authoritative_product_mapping_requires_review(self):
        package = package_metadata()
        package.pop("ExpectedPackageIdentity")
        report = evaluate_package_trust(
            package,
            ARTIFACT_SHA256,
            signature_evidence(),
            manifest_evidence(),
        )

        self.assertEqual(report["State"], TRUST_STATE_REVIEW_REQUIRED)
        self.assertTrue(report["ReviewEligible"])
        self.assertFalse(report["AutomationAllowed"])

        reviewed = review_trust_report(
            report,
            ARTIFACT_SHA256,
            reviewer="interactive-user",
            reviewed_at=datetime(2026, 7, 29, 1, 2, 3),
        )
        self.assertEqual(reviewed["State"], TRUST_STATE_REVIEWED)
        self.assertEqual(
            reviewed["Review"]["Acknowledged"],
            ["Source", "Identity", "Publisher", "Manifest"],
        )
        self.assertTrue(
            trust_report_allows_automation(reviewed, ARTIFACT_SHA256)
        )

    def test_failed_checks_cannot_be_manually_promoted(self):
        report = evaluate_package_trust(
            package_metadata(),
            ARTIFACT_SHA256,
            signature_evidence(ChainValid=False),
            manifest_evidence(),
        )

        with self.assertRaises(PackageTrustError):
            review_trust_report(report, ARTIFACT_SHA256)

    def test_changed_artifact_invalidates_report(self):
        report = evaluate_package_trust(
            package_metadata(),
            ARTIFACT_SHA256,
            signature_evidence(),
            manifest_evidence(),
        )

        self.assertFalse(trust_report_allows_automation(report, "b" * 64))

    def test_manifest_reader_extracts_signed_identity_shape(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, PACKAGE_FILENAME)
            write_appx(path)

            manifest = read_package_manifest(path)

        self.assertEqual(manifest["Identity"], "Microsoft.WindowsTerminal")
        self.assertEqual(manifest["PublisherId"], "8wekyb3d8bbwe")
        self.assertEqual(
            manifest["PackageFamilyName"],
            "Microsoft.WindowsTerminal_8wekyb3d8bbwe",
        )
        self.assertEqual(manifest["PackageType"], "package")

    def test_manifest_reader_rejects_tampered_or_ambiguous_archives(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            bad_path = os.path.join(temp_dir, PACKAGE_FILENAME)
            with open(bad_path, "wb") as handle:
                handle.write(b"not a package archive")
            with self.assertRaises(PackageTrustError):
                read_package_manifest(bad_path)

            write_appx(bad_path)
            with zipfile.ZipFile(bad_path, "a") as archive:
                archive.writestr(
                    "AppxMetadata/AppxBundleManifest.xml",
                    "<Bundle />",
                )
            with self.assertRaises(PackageTrustError):
                read_package_manifest(bad_path)

    def test_windows_signature_query_uses_env_and_normalizes_scalar_status(self):
        signature = signature_evidence(
            ChainStatus="OfflineRevocation",
            RevocationState="offline",
        )
        result = SimpleNamespace(
            returncode=0,
            stdout=json.dumps(signature),
            stderr="",
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, PACKAGE_FILENAME)
            write_appx(path)

            with patch(
                "MSStoreHelper.subprocess.run",
                return_value=result,
            ) as run_mock:
                evidence = StoreAPI.query_package_signature(path)

        command = run_mock.call_args.args[0][-1]
        environment = run_mock.call_args.kwargs["env"]
        self.assertEqual(evidence["ChainStatus"], ["OfflineRevocation"])
        self.assertNotIn(path, command)
        self.assertEqual(
            environment["MSSTOREHELPER_PACKAGE_PATH"],
            os.path.realpath(path),
        )
        self.assertIn("X509RevocationMode]::Online", command)
        self.assertIn("X509RevocationFlag]::ExcludeRoot", command)
        self.assertEqual(run_mock.call_args.kwargs["timeout"], 30)

    def test_store_api_inspection_promotes_authoritatively_bound_package(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, PACKAGE_FILENAME)
            write_appx(path)
            package = package_metadata()

            report = StoreAPI.inspect_package_trust(
                path,
                package,
                signature_info=signature_evidence(),
                evaluated_at=datetime(
                    2026,
                    7,
                    29,
                    tzinfo=timezone.utc,
                ),
            )

        self.assertEqual(report["State"], TRUST_STATE_TRUSTED)
        self.assertEqual(package["TrustState"], TRUST_STATE_TRUSTED)
        self.assertEqual(package["TrustReport"], report)
        self.assertEqual(package["Sha256"], report["ArtifactSha256"])

    def test_signature_inspection_timeout_creates_blocked_report(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, PACKAGE_FILENAME)
            write_appx(path)
            package = package_metadata()

            with patch.object(
                StoreAPI,
                "query_package_signature",
                side_effect=subprocess.TimeoutExpired("powershell", 30),
            ):
                report = StoreAPI.inspect_package_trust(path, package)

        self.assertEqual(report["State"], TRUST_STATE_BLOCKED)
        self.assertFalse(report["ReviewEligible"])
        self.assertEqual(
            report["ReasonCodes"],
            ["package-inspection-failed"],
        )
        self.assertEqual(package["TrustState"], TRUST_STATE_BLOCKED)

    def test_quarantine_blocks_automation_until_journaled_review(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, PACKAGE_FILENAME)
            write_appx(path)
            package = package_metadata(
                ExpectedProductId="9UNKNOWN",
                StoreQuery={"ProductId": "9UNKNOWN"},
                LocalPath=path,
            )
            package.pop("ExpectedPackageIdentity")
            report = StoreAPI.inspect_package_trust(
                path,
                package,
                signature_info=signature_evidence(),
            )

            self.assertEqual(
                report["State"],
                TRUST_STATE_REVIEW_REQUIRED,
            )
            self.assertFalse(
                StoreAPI.package_trust_status(
                    package,
                    path,
                    inspect_missing=False,
                )[0]
            )

            with patch("MSStoreHelper.subprocess.run") as run_mock:
                install_ok, _message = StoreAPI.install_package(path, package)
                rollback_ok, _message = StoreAPI.rollback_package(
                    "Microsoft.WindowsTerminal",
                    path,
                    package,
                )
            self.assertFalse(install_ok)
            self.assertFalse(rollback_ok)
            run_mock.assert_not_called()

            cache_ok, _message = StoreAPI.cache_downloaded_artifact(
                package,
                os.path.join(temp_dir, "cache"),
            )
            self.assertFalse(cache_ok)
            with self.assertRaisesRegex(ValueError, "quarantined"):
                StoreAPI.generate_dism_provision_script(
                    [package],
                    temp_dir,
                    "x64",
                    temp_dir,
                )
            with self.assertRaisesRegex(ValueError, "quarantined"):
                StoreAPI.write_appinstaller_export(
                    [package],
                    temp_dir,
                    os.path.join(temp_dir, "Queue.appinstaller"),
                    "x64",
                )
            with self.assertRaisesRegex(ValueError, "quarantined"):
                StoreAPI.prepare_intune_package_source(
                    [package],
                    os.path.join(temp_dir, "staging"),
                    temp_dir,
                    "x64",
                )

            StoreAPI.write_artifact_manifest(package, path, temp_dir)
            manifest = StoreAPI.load_cache_manifest(temp_dir)
            self.assertNotIn(PACKAGE_FILENAME, manifest["Artifacts"])
            self.assertIn(PACKAGE_FILENAME, manifest["Quarantine"])
            self.assertEqual(manifest["History"], {})
            self.assertEqual(
                StoreAPI.build_mirror_index(temp_dir)["PackageCount"],
                0,
            )

            invalid_journal_parent = os.path.join(
                temp_dir,
                "journal-parent-is-a-file",
            )
            with open(
                invalid_journal_parent,
                "w",
                encoding="utf-8",
            ) as handle:
                handle.write("not a directory")
            with self.assertRaises(OSError):
                StoreAPI.review_package_trust(
                    package,
                    path,
                    journal_path=os.path.join(
                        invalid_journal_parent,
                        "trust-review.jsonl",
                    ),
                )
            self.assertEqual(
                package["TrustState"],
                TRUST_STATE_REVIEW_REQUIRED,
            )
            self.assertFalse(package["TrustReport"]["AutomationAllowed"])

            journal_path = os.path.join(
                temp_dir,
                "trust-review.jsonl",
            )
            reviewed = StoreAPI.review_package_trust(
                package,
                path,
                journal_path=journal_path,
                reviewed_at=datetime(
                    2026,
                    7,
                    29,
                    1,
                    2,
                    3,
                    tzinfo=timezone.utc,
                ),
            )

            self.assertEqual(reviewed["State"], TRUST_STATE_REVIEWED)
            self.assertTrue(
                StoreAPI.package_trust_status(
                    package,
                    path,
                    inspect_missing=False,
                )[0]
            )
            self.assertEqual(
                StoreAPI.build_mirror_index(temp_dir)["PackageCount"],
                1,
            )
            manifest = StoreAPI.load_cache_manifest(temp_dir)
            self.assertIn(PACKAGE_FILENAME, manifest["Artifacts"])
            self.assertNotIn(PACKAGE_FILENAME, manifest["Quarantine"])
            with open(journal_path, "r", encoding="utf-8") as handle:
                event = json.loads(handle.readline())

        self.assertEqual(event["Event"], "package-trust-promotion")
        self.assertEqual(event["Source"]["ProductId"], "9UNKNOWN")
        self.assertEqual(
            event["Source"]["Url"],
            "https://cdn.test/package.msix",
        )


if __name__ == "__main__":
    unittest.main()
