#!/usr/bin/env python3

import json
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from command_runner import CommandTimeoutError
from MSStoreHelper import StoreAPI
from windows_capabilities import (
    InventoryDiscoveryError,
    capability_blocking_text,
    inventory_failure_text,
    inventory_is_known,
    probe_windows_capabilities,
    query_appx_inventory,
)


def completed(payload=None, *, returncode=0, stderr=""):
    stdout = json.dumps(payload) if payload is not None else ""
    return SimpleNamespace(
        returncode=returncode,
        stdout=stdout,
        stderr=stderr,
    )


class AppxInventoryTests(unittest.TestCase):
    def test_successful_machine_inventory_preserves_scope_and_sources(self):
        payload = {
            "SchemaVersion": 1,
            "Scope": "machine",
            "Installed": [{
                "Name": "Contoso.App",
                "Version": "2.0.0.0",
                "Source": "installed",
            }],
            "Provisioned": [{
                "Name": "Contoso.Base",
                "Version": "1.0.0.0",
                "Source": "provisioned",
            }],
        }
        runner = Mock(return_value=completed(payload))

        result = query_appx_inventory(
            runner,
            "powershell",
            scope="machine",
            is_admin=True,
        )

        self.assertEqual(result["Status"], "success")
        self.assertTrue(result["Known"])
        self.assertEqual(
            result["Identities"],
            ["contoso.app", "contoso.base"],
        )
        command = runner.call_args.args[0][-1]
        self.assertIn("Get-AppxPackage -AllUsers", command)
        self.assertIn("Get-AppxProvisionedPackage -Online", command)

    def test_empty_inventory_is_authoritative_not_unavailable(self):
        result = query_appx_inventory(
            lambda *_args, **_kwargs: completed({
                "SchemaVersion": 1,
                "Scope": "current-user",
                "Installed": [],
                "Provisioned": [],
            }),
            "powershell",
        )

        self.assertEqual(result["Status"], "empty")
        self.assertTrue(inventory_is_known(result))
        self.assertIn("Confirm the intended user", result["NextAction"])

    def test_machine_inventory_requires_elevation_without_running_command(self):
        runner = Mock()

        result = query_appx_inventory(
            runner,
            "powershell",
            scope="machine",
            is_admin=False,
        )

        self.assertEqual(result["Status"], "denied")
        self.assertFalse(result["Known"])
        self.assertEqual(result["ErrorCode"], "elevation-required")
        runner.assert_not_called()

    def test_timeout_policy_denial_and_unavailable_are_distinct(self):
        timed_out = query_appx_inventory(
            Mock(side_effect=CommandTimeoutError("deadline")),
            "powershell",
        )
        policy = query_appx_inventory(
            Mock(return_value=completed(
                returncode=1,
                stderr="blocked by your administrator",
            )),
            "powershell",
        )
        denied = query_appx_inventory(
            Mock(return_value=completed(
                returncode=1,
                stderr="Access is denied 0x80070005",
            )),
            "powershell",
        )
        unavailable = query_appx_inventory(
            Mock(side_effect=FileNotFoundError("powershell missing")),
            "powershell",
        )

        self.assertEqual(timed_out["Status"], "timed-out")
        self.assertEqual(policy["Status"], "policy-blocked")
        self.assertEqual(denied["Status"], "denied")
        self.assertEqual(unavailable["Status"], "unavailable")
        self.assertTrue(
            all(
                not inventory_is_known(result)
                for result in (
                    timed_out,
                    policy,
                    denied,
                    unavailable,
                )
            )
        )

    def test_invalid_inventory_output_fails_closed(self):
        result = query_appx_inventory(
            Mock(return_value=SimpleNamespace(
                returncode=0,
                stdout="{truncated",
                stderr="",
            )),
            "powershell",
        )

        self.assertEqual(result["Status"], "unavailable")
        self.assertEqual(result["ErrorCode"], "invalid-command-output")

    def test_store_api_projects_versions_without_losing_status(self):
        inventory = {
            "SchemaVersion": 1,
            "Status": "success",
            "Known": True,
            "Scope": "current-user",
            "Records": [
                {
                    "Name": "Contoso.App",
                    "Version": "1.0.0.0",
                },
                {
                    "Name": "contoso.app",
                    "Version": "2.0.0.0",
                },
            ],
            "Identities": ["contoso.app"],
        }
        with patch(
            "MSStoreHelper.query_appx_inventory",
            return_value=inventory,
        ):
            result = StoreAPI.get_installed_appx_versions()

        self.assertEqual(result["Status"], "success")
        self.assertEqual(result["Versions"]["contoso.app"], "2.0.0.0")

    def test_unknown_inventory_cannot_drive_missing_component_detection(self):
        result = {
            "Status": "timed-out",
            "Message": "Inventory timed out.",
            "NextAction": "Retry.",
        }

        with self.assertRaises(InventoryDiscoveryError):
            StoreAPI.detect_missing_ltsc_components(result)
        self.assertIn("Next action: Retry", inventory_failure_text(result))


class WindowsCapabilityTests(unittest.TestCase):
    def _payload(self):
        return {
            "SchemaVersion": 1,
            "OS": {
                "Caption": "Microsoft Windows 11 Enterprise",
                "Edition": "Enterprise",
                "ProductName": "Windows 11 Enterprise",
                "DisplayVersion": "24H2",
                "Build": "26100",
                "UBR": 4652,
                "Architecture": "64-bit",
            },
            "Context": {
                "IsElevated": True,
                "IsSystem": False,
            },
            "Policies": {
                "RemoveWindowsStore": 1,
                "DisableStoreApps": None,
                "EnableAppInstaller": None,
            },
            "Services": [
                {
                    "Name": name,
                    "Exists": True,
                    "State": "Running",
                    "StartMode": "Manual",
                }
                for name in (
                    "AppXSvc",
                    "ClipSVC",
                    "InstallService",
                    "LicenseManager",
                    "wuauserv",
                )
            ],
            "RebootPending": {
                "Pending": True,
                "ComponentBasedServicing": True,
                "WindowsUpdate": False,
                "PendingFileRenameOperations": False,
            },
        }

    @patch("windows_capabilities.platform.system", return_value="Windows")
    def test_preflight_reports_os_context_policy_services_network_and_reboot(
        self,
        _system,
    ):
        report = probe_windows_capabilities(
            Mock(return_value=completed(self._payload())),
            "powershell",
            is_admin=True,
            source_health=[{
                "Key": "rg-adguard",
                "Available": True,
                "Detail": "HTTP 200",
            }],
        )

        self.assertEqual(report["Status"], "policy-blocked")
        self.assertEqual(report["Platform"]["Edition"], "Enterprise")
        self.assertEqual(report["Platform"]["Build"], "26100.4652")
        self.assertEqual(
            report["Context"]["InventoryScope"],
            "machine",
        )
        self.assertEqual(report["Policies"]["Store"], "blocked")
        self.assertEqual(len(report["Services"]), 5)
        self.assertEqual(report["Network"]["Status"], "available")
        self.assertEqual(report["RebootPending"]["State"], "pending")
        self.assertEqual(
            capability_blocking_text(
                report,
                required_sources={"rg-adguard"},
            ),
            "",
        )
        self.assertIn(
            "Windows Store access is disabled",
            capability_blocking_text(
                report,
                respect_policy_codes={"store-policy-blocked"},
            ),
        )

    @patch("windows_capabilities.platform.system", return_value="Windows")
    def test_required_service_and_endpoint_block_with_next_action(
        self,
        _system,
    ):
        payload = self._payload()
        payload["Policies"]["RemoveWindowsStore"] = None
        payload["Services"][0] = {
            "Name": "AppXSvc",
            "Exists": True,
            "State": "Stopped",
            "StartMode": "Disabled",
        }
        report = probe_windows_capabilities(
            Mock(return_value=completed(payload)),
            "powershell",
            is_admin=True,
            source_health=[{
                "Key": "rg-adguard",
                "Available": False,
                "Detail": "HTTP 503",
            }],
        )

        service_block = capability_blocking_text(
            report,
            required_services={"AppXSvc"},
        )
        source_block = capability_blocking_text(
            report,
            required_sources={"rg-adguard"},
        )
        ignored_service = capability_blocking_text(
            report,
            required_services={"ClipSVC"},
        )

        self.assertIn("AppXSvc is disabled", service_block)
        self.assertIn("Next action:", service_block)
        self.assertIn("rg-adguard", source_block)
        self.assertEqual(ignored_service, "")


if __name__ == "__main__":
    unittest.main()
