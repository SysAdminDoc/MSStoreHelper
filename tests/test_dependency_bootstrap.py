#!/usr/bin/env python3

import unittest

from MSStoreHelper import (
    dependency_setup_message,
    find_missing_dependencies,
    python_runtime_error,
)


class DependencyBootstrapTests(unittest.TestCase):
    def test_find_missing_dependencies_returns_pinned_requirements(self):
        def fake_import(name):
            if name == "requests":
                raise ImportError(name)
            return object()

        self.assertEqual(find_missing_dependencies(fake_import), ["requests==2.34.2"])

    def test_dependency_setup_message_includes_online_and_offline_commands(self):
        message = dependency_setup_message(["customtkinter==5.2.2"])

        self.assertIn("pip install --require-hashes -r requirements.txt", message)
        self.assertIn("scripts\\build_wheelhouse.py --output wheelhouse --test", message)
        self.assertIn("--no-index --find-links wheelhouse", message)

    def test_runtime_floor_rejects_unsupported_python(self):
        self.assertIn("detected Python 3.10", python_runtime_error((3, 10)))
        self.assertIsNone(python_runtime_error((3, 11)))
        self.assertIsNone(python_runtime_error((3, 14)))


if __name__ == "__main__":
    unittest.main()
