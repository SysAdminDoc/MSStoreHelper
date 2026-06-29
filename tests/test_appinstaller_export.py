#!/usr/bin/env python3

import os
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

from MSStoreHelper import APPINSTALLER_NS, StoreAPI


def write_fake_appx(path, name, publisher="CN=Contoso", version="1.0.0.0", arch="x64"):
    manifest = f"""<?xml version="1.0" encoding="utf-8"?>
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10">
  <Identity Name="{name}" Publisher="{publisher}" Version="{version}" ProcessorArchitecture="{arch}" />
</Package>
"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("AppxManifest.xml", manifest)


class AppInstallerExportTests(unittest.TestCase):
    def test_write_appinstaller_export_copies_packages_and_generates_manifest(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            downloads = os.path.join(temp_dir, "downloads")
            os.makedirs(downloads)
            app = os.path.join(downloads, "Contoso.App_2.0.0.0_x64__test.msix")
            vclibs = os.path.join(downloads, "Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.appx")
            write_fake_appx(app, "Contoso.App", version="2.0.0.0")
            write_fake_appx(vclibs, "Microsoft.VCLibs.140.00", publisher="CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US", version="14.0.33519.0")

            appinstaller_path = os.path.join(temp_dir, "ContosoQueue.appinstaller")
            result = StoreAPI.write_appinstaller_export(
                [
                    {"FileName": os.path.basename(app), "LocalPath": app},
                    {"FileName": os.path.basename(vclibs), "LocalPath": vclibs},
                ],
                downloads,
                appinstaller_path,
                "x64",
            )

            self.assertEqual(result["PackageCount"], 2)
            self.assertTrue(os.path.exists(os.path.join(result["PackageDir"], os.path.basename(app))))
            self.assertTrue(os.path.exists(os.path.join(result["PackageDir"], os.path.basename(vclibs))))

            root = ET.parse(appinstaller_path).getroot()
            ns = {"ai": APPINSTALLER_NS}
            main = root.find("ai:MainPackage", ns)
            deps = root.find("ai:Dependencies", ns)
            update = root.find("ai:UpdateSettings/ai:OnLaunch", ns)

            self.assertEqual(root.tag, f"{{{APPINSTALLER_NS}}}AppInstaller")
            self.assertEqual(main.attrib["Name"], "Contoso.App")
            self.assertEqual(main.attrib["ProcessorArchitecture"], "x64")
            self.assertEqual(deps.find("ai:Package", ns).attrib["Name"], "Microsoft.VCLibs.140.00")
            self.assertEqual(update.attrib["HoursBetweenUpdateChecks"], "12")
            self.assertIn(".appinstaller", root.attrib["Uri"])
            self.assertIn(os.path.basename(app), main.attrib["Uri"])

    def test_write_appinstaller_export_requires_downloaded_package_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaises(ValueError):
                StoreAPI.write_appinstaller_export(
                    [{"FileName": "Contoso.App_1.0.0.0_x64__test.msix"}],
                    os.path.join(temp_dir, "downloads"),
                    os.path.join(temp_dir, "Queue.appinstaller"),
                    "x64",
                )

    def test_read_appx_identity_rejects_non_appx_archive(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            package = os.path.join(temp_dir, "broken.msix")
            with open(package, "wb") as handle:
                handle.write(b"not a zip")

            with self.assertRaises(ValueError):
                StoreAPI.read_appx_identity(package)


if __name__ == "__main__":
    unittest.main()
