#!/usr/bin/env python3

import os
import tempfile
import unittest
import zipfile

from MSStoreHelper import StoreAPI


def write_appx(path, version, capabilities=None, dependencies=None):
    capabilities = capabilities or []
    dependencies = dependencies or []
    cap_xml = "\n".join(f'    <uap:Capability Name="{name}" />' for name in capabilities)
    dep_xml = "\n".join(
        f'    <PackageDependency Name="{name}" Publisher="CN=Microsoft Corporation" MinVersion="{min_version}" />'
        for name, min_version in dependencies
    )
    manifest = f"""<?xml version="1.0" encoding="utf-8"?>
<Package
  xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
  xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10">
  <Identity Name="Contoso.App" Publisher="CN=Contoso" Version="{version}" ProcessorArchitecture="x64" />
  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.19041.0" MaxVersionTested="10.0.26100.0" />
{dep_xml}
  </Dependencies>
  <Capabilities>
{cap_xml}
  </Capabilities>
</Package>
"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("AppxManifest.xml", manifest)


class PackageDiffTests(unittest.TestCase):
    def test_diff_appx_manifests_reports_capability_and_dependency_changes(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            old_path = os.path.join(temp_dir, "Contoso.App_1.0.0.0_x64__test.msix")
            new_path = os.path.join(temp_dir, "Contoso.App_2.0.0.0_x64__test.msix")
            write_appx(
                old_path,
                "1.0.0.0",
                capabilities=["internetClient", "picturesLibrary"],
                dependencies=[("Microsoft.VCLibs.140.00", "14.0.30000.0")],
            )
            write_appx(
                new_path,
                "2.0.0.0",
                capabilities=["internetClient", "documentsLibrary"],
                dependencies=[
                    ("Microsoft.VCLibs.140.00", "14.0.30000.0"),
                    ("Microsoft.UI.Xaml.2.8", "8.2300.0.0"),
                ],
            )

            diff = StoreAPI.diff_appx_manifests(old_path, new_path)
            formatted = StoreAPI.format_package_diff(diff)

            self.assertTrue(diff["VersionChanged"])
            self.assertIn("Capability: documentsLibrary", diff["Capabilities"]["Added"])
            self.assertIn("Capability: picturesLibrary", diff["Capabilities"]["Removed"])
            self.assertIn("PackageDependency: Microsoft.UI.Xaml.2.8 >= 8.2300.0.0 (CN=Microsoft Corporation)", diff["Dependencies"]["Added"])
            self.assertIn("Added: Capability: documentsLibrary", formatted)
            self.assertIn("Removed: Capability: picturesLibrary", formatted)

    def test_package_diff_candidates_use_newest_two_cached_versions(self):
        with tempfile.TemporaryDirectory() as cache_dir:
            for version in ("1.0.0.0", "2.0.0.0", "3.0.0.0"):
                path = os.path.join(cache_dir, f"Contoso.App_{version}_x64__test.msix")
                write_appx(path, version, capabilities=["internetClient"])
                StoreAPI.write_artifact_manifest({"FileName": os.path.basename(path)}, path, cache_dir)

            candidates = StoreAPI.package_diff_candidates([cache_dir], ["Contoso.App"])

            self.assertEqual(len(candidates), 1)
            self.assertEqual(candidates[0]["New"]["AvailableVersion"], "3.0.0.0")
            self.assertEqual(candidates[0]["Old"]["AvailableVersion"], "2.0.0.0")


if __name__ == "__main__":
    unittest.main()
