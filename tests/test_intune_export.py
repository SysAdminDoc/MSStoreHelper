#!/usr/bin/env python3

import os
import tempfile
import unittest

from MSStoreHelper import StoreAPI


class IntuneExportTests(unittest.TestCase):
    def test_prepare_intune_package_source_writes_install_and_detection_scripts(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            downloads = os.path.join(temp_dir, "downloads")
            staging = os.path.join(temp_dir, "staging")
            os.makedirs(downloads)
            vclibs = os.path.join(downloads, "Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.appx")
            app = os.path.join(downloads, "Contoso.App_2.1.0.0_x64__test.msixbundle")
            for path in (vclibs, app):
                with open(path, "wb") as handle:
                    handle.write(b"package")

            info = StoreAPI.prepare_intune_package_source(
                [
                    {"FileName": os.path.basename(app), "LocalPath": app},
                    {
                        "FileName": os.path.basename(vclibs),
                        "LocalPath": vclibs,
                        "StoreQuery": {"Ring": "WIF", "Language": "ja-JP", "Market": "JP"},
                    },
                ],
                staging,
                downloads,
                "x64",
                "ContosoQueue",
            )

            self.assertEqual(info["PackageCount"], 2)
            self.assertTrue(os.path.exists(os.path.join(info["PackagesDir"], os.path.basename(app))))
            self.assertTrue(os.path.exists(os.path.join(info["PackagesDir"], os.path.basename(vclibs))))
            with open(os.path.join(info["SourceDir"], "ContosoQueue.ps1"), encoding="utf-8") as handle:
                install_script = handle.read()
            with open(info["DetectionPath"], encoding="utf-8") as handle:
                detection_script = handle.read()
            self.assertLess(install_script.index("Microsoft.VCLibs.140.00"), install_script.index("Contoso.App"))
            self.assertIn("/Add-ProvisionedAppxPackage", install_script)
            self.assertIn("StoreRing = 'WIF'", install_script)
            self.assertIn("StoreLanguage = 'ja-JP'", install_script)
            self.assertIn("Get-AppxProvisionedPackage", detection_script)
            self.assertIn("Contoso.App", detection_script)
            with open(info["GuidePath"], encoding="utf-8") as handle:
                guide = handle.read()
            self.assertIn("Store query: Retail/en-US/US, WIF/ja-JP/JP", guide)

    def test_prepare_intune_package_source_rejects_missing_downloads(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaises(ValueError):
                StoreAPI.prepare_intune_package_source(
                    [{"FileName": "Contoso.App_1.0.0.0_x64__test.msixbundle"}],
                    os.path.join(temp_dir, "staging"),
                    os.path.join(temp_dir, "downloads"),
                    "x64",
                )

    def test_build_intunewinapputil_command_uses_official_switches(self):
        command = StoreAPI.build_intunewinapputil_command(
            r"C:\Tools\IntuneWinAppUtil.exe",
            r"C:\Source",
            r"C:\Source\Install.cmd",
            r"C:\Output",
        )

        self.assertEqual(
            command,
            [
                r"C:\Tools\IntuneWinAppUtil.exe",
                "-c", r"C:\Source",
                "-s", "Install.cmd",
                "-o", r"C:\Output",
                "-q",
            ],
        )


if __name__ == "__main__":
    unittest.main()
