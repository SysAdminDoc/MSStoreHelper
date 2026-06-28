#!/usr/bin/env python3

import json
import os
import tempfile
import unittest
from unittest.mock import patch

from MSStoreHelper import StoreAPI


class FakePowerShellResult:
    def __init__(self, returncode=0, stdout="", stderr=""):
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr


class StoreRepairTests(unittest.TestCase):
    def test_repair_steps_include_cache_token_and_license_sync(self):
        steps = StoreAPI.get_store_repair_steps()
        descriptions = "\n".join(description for description, _command in steps)
        commands = "\n".join(command for _description, command in steps)

        self.assertIn("Resetting Store cache", descriptions)
        self.assertIn("Rebuilding Store token cache", descriptions)
        self.assertIn("Re-syncing Store licensing", descriptions)
        self.assertIn("wsreset.exe", commands)
        self.assertIn("TokenBroker", commands)
        self.assertIn("ClipSVC", commands)
        self.assertIn("LicenseManager", commands)
        self.assertIn("Microsoft.StorePurchaseApp", commands)
        self.assertIn("Backup-MSStoreHelperPath", commands)

    def test_provisioning_repair_steps_clear_deprovisioned_store_keys(self):
        steps = StoreAPI.get_provisioning_repair_steps()
        descriptions = "\n".join(description for description, _command in steps)
        commands = "\n".join(command for _description, command in steps)

        self.assertIn("Clearing Store deprovision tombstones", descriptions)
        self.assertIn("Re-registering Store apps", descriptions)
        self.assertIn("AppxAllUserStore\\Deprovisioned", commands)
        self.assertIn("Microsoft.WindowsStore", commands)
        self.assertIn("Microsoft.StorePurchaseApp", commands)
        self.assertIn("Microsoft.DesktopAppInstaller", commands)
        self.assertIn("Add-AppxPackage", commands)
        self.assertIn("Backup-MSStoreHelperRegistryPath", commands)

    def test_licensing_reset_steps_restart_services_and_clear_cache(self):
        steps = StoreAPI.get_licensing_reset_steps()
        descriptions = "\n".join(description for description, _command in steps)
        commands = "\n".join(command for _description, command in steps)

        self.assertIn("Stopping licensing services", descriptions)
        self.assertIn("Clearing ClipSVC license cache", descriptions)
        self.assertIn("Starting licensing services", descriptions)
        self.assertIn("ClipSVC", commands)
        self.assertIn("LicenseManager", commands)
        self.assertIn("GenuineTicket", commands)
        self.assertIn("Microsoft.StorePurchaseApp", commands)
        self.assertIn("Backup-MSStoreHelperPath", commands)

    def test_cache_rebuild_steps_scan_backup_and_recreate_cache(self):
        steps = StoreAPI.get_cache_rebuild_steps()
        descriptions = "\n".join(description for description, _command in steps)
        commands = "\n".join(command for _description, command in steps)

        self.assertIn("Scanning Store cache folders", descriptions)
        self.assertIn("Backing up existing Store caches", descriptions)
        self.assertIn("Rebuilding clean Store cache folders", descriptions)
        self.assertIn("LocalCache", commands)
        self.assertIn("INetCache", commands)
        self.assertIn("Backup-MSStoreHelperPath", commands)
        self.assertIn("New-Item", commands)
        self.assertIn("wsreset.exe", commands)

    def test_run_powershell_steps_records_manifest_restore_script_and_output(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch("MSStoreHelper.subprocess.run", return_value=FakePowerShellResult(1, "out text", "err text")) as run_mock:
                results = StoreAPI._run_powershell_steps(
                    [("Failing step", "Write-Error 'bad'")],
                    timeout=5,
                    repair_name="unit-test",
                    backup_root=temp_dir,
                )

            self.assertEqual(len(results), 1)
            self.assertFalse(results[0]["Success"])
            self.assertEqual(results[0]["ReturnCode"], 1)
            self.assertEqual(results[0]["Stdout"], "out text")
            self.assertEqual(results[0]["Stderr"], "err text")
            self.assertTrue(os.path.exists(results[0]["ManifestPath"]))
            self.assertTrue(os.path.exists(results[0]["RestoreScriptPath"]))
            invoked_command = run_mock.call_args.args[0][-1]
            self.assertIn("Backup-MSStoreHelperPath", invoked_command)
            self.assertIn("Microsoft\\.PowerShell\\.Core\\\\Registry::HKEY_LOCAL_MACHINE", invoked_command)
            self.assertIn("Write-Error 'bad'", invoked_command)

            with open(results[0]["ManifestPath"], "r", encoding="utf-8") as handle:
                manifest = json.load(handle)
            self.assertEqual(manifest["RepairName"], "unit-test")
            self.assertEqual(manifest["Results"][0]["Stderr"], "err text")

            with open(results[0]["RestoreScriptPath"], "r", encoding="utf-8") as handle:
                restore_script = handle.read()
            self.assertIn("backup-records.jsonl", restore_script)
            self.assertIn("reg.exe import", restore_script)


if __name__ == "__main__":
    unittest.main()
