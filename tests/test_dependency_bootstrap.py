#!/usr/bin/env python3

import unittest

from MSStoreHelper import dependency_setup_message, find_missing_dependencies


class DependencyBootstrapTests(unittest.TestCase):
    def test_find_missing_dependencies_returns_pinned_requirements(self):
        def fake_import(name):
            if name == "requests":
                raise ImportError(name)
            return object()

        self.assertEqual(find_missing_dependencies(fake_import), ["requests==2.32.5"])

    def test_dependency_setup_message_includes_online_and_offline_commands(self):
        message = dependency_setup_message(["customtkinter==5.2.2"])

        self.assertIn("pip install -r requirements.txt", message)
        self.assertIn("pip download -r requirements.txt -d wheelhouse", message)
        self.assertIn("--no-index --find-links wheelhouse", message)


if __name__ == "__main__":
    unittest.main()
