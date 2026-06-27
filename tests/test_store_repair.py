#!/usr/bin/env python3

import unittest

from MSStoreHelper import StoreAPI


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

    def test_cache_rebuild_steps_scan_backup_and_recreate_cache(self):
        steps = StoreAPI.get_cache_rebuild_steps()
        descriptions = "\n".join(description for description, _command in steps)
        commands = "\n".join(command for _description, command in steps)

        self.assertIn("Scanning Store cache folders", descriptions)
        self.assertIn("Backing up existing Store caches", descriptions)
        self.assertIn("Rebuilding clean Store cache folders", descriptions)
        self.assertIn("LocalCache", commands)
        self.assertIn("INetCache", commands)
        self.assertIn("Move-Item", commands)
        self.assertIn(".bak-", commands)
        self.assertIn("New-Item", commands)
        self.assertIn("wsreset.exe", commands)


if __name__ == "__main__":
    unittest.main()
