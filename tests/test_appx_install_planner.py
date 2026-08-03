#!/usr/bin/env python3

import io
import json
import os
import tempfile
import unittest
import zipfile

from appx_install_planner import (
    AppxInspectionError,
    InstallPlanError,
    build_install_plan,
    inspect_appx_archive,
    render_install_plan,
    validate_install_plan,
)


PUBLISHER = "CN=Contoso"


def appx_bytes(
    name,
    *,
    version="1.0.0.0",
    architecture="x64",
    framework=False,
    resource=False,
    dependency=None,
    main_dependency=None,
    capability=None,
    min_os="10.0.19041.0",
):
    properties = []
    if framework:
        properties.append("<Framework>true</Framework>")
    if resource:
        properties.append("<ResourcePackage>true</ResourcePackage>")
    property_xml = (
        "<Properties>" + "".join(properties) + "</Properties>"
        if properties
        else ""
    )
    dependencies = [
        (
            '<TargetDeviceFamily Name="Windows.Desktop" '
            f'MinVersion="{min_os}" MaxVersionTested="10.0.26100.0" />'
        )
    ]
    if dependency:
        dependencies.append(
            f'<PackageDependency Name="{dependency[0]}" '
            f'Publisher="{dependency[1]}" MinVersion="{dependency[2]}" />'
        )
    if main_dependency:
        dependencies.append(
            '<uap3:MainPackageDependency '
            f'Name="{main_dependency}" />'
        )
    capability_xml = (
        f"<Capabilities><uap:Capability Name=\"{capability}\" /></Capabilities>"
        if capability
        else ""
    )
    manifest = f"""<?xml version="1.0" encoding="utf-8"?>
<Package
  xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
  xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
  xmlns:uap3="http://schemas.microsoft.com/appx/manifest/uap/windows10/3">
  <Identity Name="{name}" Publisher="{PUBLISHER}" Version="{version}" ProcessorArchitecture="{architecture}" />
  {property_xml}
  <Dependencies>{''.join(dependencies)}</Dependencies>
  <Applications><Application Id="App" Executable="app.exe" EntryPoint="Windows.FullTrustApplication" /></Applications>
  {capability_xml}
</Package>
"""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("AppxManifest.xml", manifest)
    return buffer.getvalue()


def write_appx(path, name, **kwargs):
    with open(path, "wb") as handle:
        handle.write(appx_bytes(name, **kwargs))


def write_bundle(path):
    outer = """<?xml version="1.0" encoding="utf-8"?>
<Bundle xmlns="http://schemas.microsoft.com/appx/2013/bundle">
  <Identity Name="Contoso.App" Publisher="CN=Contoso" Version="2.0.0.0" />
  <Packages>
    <Package Type="application" Architecture="x64" FileName="Contoso.App_x64.msix" />
    <Package Type="resource" Architecture="resource" ResourceId="resources.en" FileName="Contoso.App_resources.msix">
      <Resources><Resource Language="en-US" Scale="200" /></Resources>
    </Package>
  </Packages>
</Bundle>
"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "AppxMetadata/AppxBundleManifest.xml",
            outer,
        )
        archive.writestr(
            "Contoso.App_x64.msix",
            appx_bytes(
                "Contoso.App",
                version="2.0.0.0",
                dependency=(
                    "Contoso.Framework",
                    PUBLISHER,
                    "3.0.0.0",
                ),
                capability="documentsLibrary",
                min_os="10.0.22621.0",
            ),
        )
        archive.writestr(
            "Contoso.App_resources.msix",
            appx_bytes(
                "Contoso.App",
                version="2.0.0.0",
                architecture="resource",
                resource=True,
                min_os="10.0.22621.0",
            ),
        )


class AppxInspectionTests(unittest.TestCase):
    def test_bundle_inspection_includes_inner_manifests(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(
                temp_dir,
                "Contoso.App_2.0.0.0_neutral__test.msixbundle",
            )
            write_bundle(path)

            details = inspect_appx_archive(path)

            self.assertEqual(details["ContainerType"], "bundle")
            self.assertEqual(details["Identity"]["Name"], "Contoso.App")
            self.assertEqual(len(details["InnerPackages"]), 2)
            self.assertEqual(details["MinOSVersion"], "10.0.22621.0")
            self.assertIn(
                "Capability: documentsLibrary",
                details["Capabilities"],
            )
            self.assertEqual(
                details["PackageDependencies"][0]["Name"],
                "Contoso.Framework",
            )
            self.assertEqual(
                details["InnerPackages"][1]["BundleResources"][0][
                    "Language"
                ],
                "en-US",
            )

    def test_bundle_rejects_missing_inner_package(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, "broken.msixbundle")
            manifest = """<Bundle xmlns="http://schemas.microsoft.com/appx/2013/bundle">
  <Identity Name="Contoso.App" Publisher="CN=Contoso" Version="1.0.0.0" />
  <Packages><Package Type="application" Architecture="x64" FileName="missing.msix" /></Packages>
</Bundle>"""
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr(
                    "AppxMetadata/AppxBundleManifest.xml",
                    manifest,
                )

            with self.assertRaises(AppxInspectionError):
                inspect_appx_archive(path)


class InstallPlanTests(unittest.TestCase):
    def _write_plan_packages(self, temp_dir):
        main = os.path.join(temp_dir, "Contoso.App.msix")
        framework = os.path.join(temp_dir, "Contoso.Framework.appx")
        optional = os.path.join(temp_dir, "Contoso.Addon.msix")
        resource = os.path.join(temp_dir, "Contoso.App.resources.appx")
        write_appx(
            main,
            "Contoso.App",
            version="2.0.0.0",
            dependency=(
                "Contoso.Framework",
                PUBLISHER,
                "3.0.0.0",
            ),
        )
        write_appx(
            framework,
            "Contoso.Framework",
            version="3.1.0.0",
            framework=True,
        )
        write_appx(
            optional,
            "Contoso.Addon",
            version="2.0.0.0",
            main_dependency="Contoso.App",
        )
        write_appx(
            resource,
            "Contoso.App",
            version="2.0.0.0",
            architecture="resource",
            resource=True,
        )
        return [main, framework, optional, resource]

    def test_plan_groups_manifest_roles_and_round_trips(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = self._write_plan_packages(temp_dir)
            inventory = {
                "Status": "success",
                "Records": [{
                    "Name": "Contoso.Framework",
                    "Version": "3.1.0.0",
                    "Publisher": PUBLISHER,
                    "Architecture": "x64",
                }],
            }

            plan = build_install_plan(
                [
                    {"Path": path, "FileName": os.path.basename(path)}
                    for path in paths
                ],
                target_architecture="x64",
                target_os_version="10.0.26100.0",
                inventory=inventory,
            )
            round_tripped = json.loads(json.dumps(plan))

            validate_install_plan(round_tripped)
            self.assertTrue(plan["Installable"])
            self.assertEqual(plan["Main"]["Identity"]["Name"], "Contoso.App")
            self.assertEqual(
                plan["Dependencies"][0]["Action"],
                "skip",
            )
            self.assertEqual(
                plan["OptionalPackages"][0]["Identity"]["Name"],
                "Contoso.Addon",
            )
            self.assertEqual(len(plan["ResourcePackages"]), 1)
            self.assertNotIn(
                paths[1],
                plan["Deployment"]["DependencyPaths"],
            )
            self.assertIn(
                "Result: ready",
                render_install_plan(round_tripped),
            )

    def test_plan_rejects_multiple_independent_apps(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            first = os.path.join(temp_dir, "first.msix")
            second = os.path.join(temp_dir, "second.msix")
            write_appx(first, "Contoso.First")
            write_appx(second, "Contoso.Second")

            with self.assertRaisesRegex(
                InstallPlanError,
                "multiple independent main apps",
            ):
                build_install_plan(
                    [{"Path": first}, {"Path": second}],
                    target_architecture="x64",
                )

    def test_plan_blocks_missing_dependency_os_and_downgrade(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            main = os.path.join(temp_dir, "Contoso.App.msix")
            write_appx(
                main,
                "Contoso.App",
                version="2.0.0.0",
                dependency=(
                    "Contoso.Framework",
                    PUBLISHER,
                    "3.0.0.0",
                ),
                min_os="10.0.22621.0",
            )
            inventory = {
                "Status": "success",
                "Records": [{
                    "Name": "Contoso.App",
                    "Version": "3.0.0.0",
                    "Publisher": PUBLISHER,
                    "Architecture": "x64",
                }],
            }

            plan = build_install_plan(
                [{"Path": main}],
                target_architecture="x64",
                target_os_version="10.0.19041.0",
                inventory=inventory,
            )

            self.assertFalse(plan["Installable"])
            self.assertEqual(
                {item["Code"] for item in plan["Conflicts"]},
                {"missing-dependency", "minimum-os", "downgrade"},
            )

    def test_plan_rejects_unknown_inventory(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            main = os.path.join(temp_dir, "Contoso.App.msix")
            write_appx(main, "Contoso.App")

            with self.assertRaisesRegex(
                InstallPlanError,
                "not authoritative",
            ):
                build_install_plan(
                    [{"Path": main}],
                    target_architecture="x64",
                    inventory={"Status": "timed-out"},
                )


if __name__ == "__main__":
    unittest.main()
