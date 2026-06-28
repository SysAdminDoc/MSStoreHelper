#!/usr/bin/env python3

import os
import tempfile
import unittest

from MSStoreHelper import StoreAPI


class DismExportTests(unittest.TestCase):
    def test_generate_dism_script_orders_dependencies_before_apps(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            packages = [
                {"FileName": "Contoso.App_2.0.0.0_x64__test.msixbundle"},
                {"FileName": "Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.appx"},
                {"FileName": "Contoso.App_2.0.0.0_x64__test.BlockMap"},
            ]

            script = StoreAPI.generate_dism_provision_script(packages, temp_dir, "x64", temp_dir)

            self.assertIn("/Add-ProvisionedAppxPackage", script)
            self.assertIn("/SkipLicense", script)
            self.assertLess(
                script.index("Microsoft.VCLibs.140.00"),
                script.index("Contoso.App_2.0.0.0"),
            )
            self.assertIn("PackagePath = 'Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.appx'", script)
            self.assertNotIn("BlockMap", script)

    def test_generate_dism_script_prefers_local_path_and_escapes_values(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source_dir = os.path.join(temp_dir, "external")
            script_dir = os.path.join(temp_dir, "scripts")
            os.makedirs(source_dir)
            package_path = os.path.join(source_dir, "Contoso.O'Hare_1.0.0.0_neutral__test.msix")
            with open(package_path, "wb") as handle:
                handle.write(b"package")

            script = StoreAPI.generate_dism_provision_script(
                [{"FileName": os.path.basename(package_path), "LocalPath": package_path}],
                os.path.join(temp_dir, "downloads"),
                "x64",
                script_dir,
            )

            escaped_path = os.path.abspath(package_path).replace("'", "''")
            self.assertIn("Contoso.O''Hare_1.0.0.0_neutral__test.msix", script)
            self.assertIn(f"PackagePath = '{escaped_path}'", script)

    def test_generate_dism_script_rejects_empty_installable_queue(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaises(ValueError):
                StoreAPI.generate_dism_provision_script(
                    [{"FileName": "Contoso.App_1.0.0.0_x64__test.BlockMap"}],
                    temp_dir,
                    "x64",
                    temp_dir,
                )


if __name__ == "__main__":
    unittest.main()
