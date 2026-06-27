#!/usr/bin/env python3

import unittest

from msstore_package_resolution import (
    installed_version_satisfies_package,
    is_dependency_package,
    order_packages_for_install,
    package_role,
    package_version_tuple,
    select_recommended_packages,
    signature_info_is_valid_microsoft,
    version_tuple_from_text,
)


def package(filename, arch="x64", is_bundle=False, encrypted=False):
    return {
        "FileName": filename,
        "Architecture": arch,
        "FileType": filename.rsplit(".", 1)[-1].upper(),
        "IsBundle": is_bundle,
        "IsEncrypted": encrypted,
    }


class PackageResolutionTests(unittest.TestCase):
    def test_select_recommended_packages_includes_dependency_chain(self):
        packages = [
            package("Contoso.App_5.0.0.0_x64__8wekyb3d8bbwe.Msix"),
            package("Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64__8wekyb3d8bbwe.Appx"),
            package("Microsoft.NET.Native.Runtime.2.2_2.2.28604.0_x64__8wekyb3d8bbwe.Appx"),
            package("Microsoft.UI.Xaml.2.8_8.2310.30001.0_x64__8wekyb3d8bbwe.Appx"),
            package("Microsoft.NET.Native.Framework.2.2_2.2.29512.0_x64__8wekyb3d8bbwe.Appx"),
        ]

        selected = select_recommended_packages(reversed(packages), "x64")

        self.assertEqual(
            [item["FileName"] for item in selected],
            [
                "Microsoft.NET.Native.Framework.2.2_2.2.29512.0_x64__8wekyb3d8bbwe.Appx",
                "Microsoft.NET.Native.Runtime.2.2_2.2.28604.0_x64__8wekyb3d8bbwe.Appx",
                "Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64__8wekyb3d8bbwe.Appx",
                "Microsoft.UI.Xaml.2.8_8.2310.30001.0_x64__8wekyb3d8bbwe.Appx",
                "Contoso.App_5.0.0.0_x64__8wekyb3d8bbwe.Msix",
            ],
        )

    def test_select_recommended_packages_keeps_newest_compatible_package(self):
        packages = [
            package("Contoso.App_1.0.0.0_x64__8wekyb3d8bbwe.Msix"),
            package("Contoso.App_2.0.0.0_x86__8wekyb3d8bbwe.Msix", arch="x86"),
            package("Contoso.App_3.0.0.0_x64__8wekyb3d8bbwe.Emsix", encrypted=True),
            package("Contoso.App_2.5.0.0_x64__8wekyb3d8bbwe.Msix"),
        ]

        selected = select_recommended_packages(packages, "x64")

        self.assertEqual(len(selected), 1)
        self.assertEqual(selected[0]["FileName"], "Contoso.App_2.5.0.0_x64__8wekyb3d8bbwe.Msix")

    def test_explicit_architecture_override_prefers_exact_arch(self):
        packages = [
            package("Contoso.App_3.0.0.0_neutral__8wekyb3d8bbwe.Msixbundle", arch="neutral", is_bundle=True),
            package("Contoso.App_2.5.0.0_x64__8wekyb3d8bbwe.Msix", arch="x64"),
        ]

        selected = select_recommended_packages(packages, "x64", prefer_exact_arch=True)

        self.assertEqual(len(selected), 1)
        self.assertEqual(selected[0]["FileName"], "Contoso.App_2.5.0.0_x64__8wekyb3d8bbwe.Msix")

    def test_order_packages_for_install_sorts_manual_queue(self):
        packages = [
            package("Contoso.App_5.0.0.0_x64__8wekyb3d8bbwe.Msix"),
            package("Microsoft.UI.Xaml.2.8_8.2310.30001.0_x64__8wekyb3d8bbwe.Appx"),
            package("Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64__8wekyb3d8bbwe.Appx"),
        ]

        ordered = order_packages_for_install(packages, "x64")

        self.assertEqual(
            [package_role(item["FileName"]) for item in ordered],
            ["vclibs", "ui_xaml", "app"],
        )
        self.assertTrue(is_dependency_package(ordered[0]))
        self.assertEqual(package_version_tuple(ordered[-1]["FileName"]), (5, 0, 0, 0))

    def test_installed_version_satisfies_package(self):
        candidate = package("Contoso.App_2.5.0.0_x64__8wekyb3d8bbwe.Msix")

        self.assertTrue(installed_version_satisfies_package(candidate, "2.5.0.0"))
        self.assertTrue(installed_version_satisfies_package(candidate, "3.0.0.0"))
        self.assertFalse(installed_version_satisfies_package(candidate, "2.4.9.0"))

    def test_unknown_available_version_is_not_current(self):
        candidate = package("Contoso.App_x64__8wekyb3d8bbwe.Msix")

        self.assertEqual(package_version_tuple(candidate["FileName"]), ())
        self.assertEqual(version_tuple_from_text("Version: 1.2.3.4"), (1, 2, 3, 4))
        self.assertFalse(installed_version_satisfies_package(candidate, "9.9.9.9"))

    def test_signature_info_requires_valid_microsoft_signature(self):
        self.assertTrue(signature_info_is_valid_microsoft({
            "Status": "Valid",
            "ChainValid": False,
            "Signer": "CN=Microsoft Corporation",
            "Root": "CN=Microsoft Root Certificate Authority 2011",
        }))
        self.assertFalse(signature_info_is_valid_microsoft({
            "Status": "Valid",
            "ChainValid": True,
            "Signer": "CN=Contoso",
            "Root": "CN=Contoso Root",
        }))
        self.assertFalse(signature_info_is_valid_microsoft({
            "Status": "HashMismatch",
            "ChainValid": True,
            "Signer": "CN=Microsoft Corporation",
            "Root": "CN=Microsoft Root Certificate Authority 2011",
        }))


if __name__ == "__main__":
    unittest.main()
