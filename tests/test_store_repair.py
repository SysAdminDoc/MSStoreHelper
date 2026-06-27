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


if __name__ == "__main__":
    unittest.main()
