#!/usr/bin/env python3

import re
import tomllib
import unittest
from pathlib import Path

from scripts.build_wheelhouse import default_lock_path, parse_lock
from scripts.lock_dependencies import ARCHITECTURES, PYTHON_TAGS, read_input_pins

ROOT = Path(__file__).resolve().parents[1]


class DependencyLockTests(unittest.TestCase):
    def test_project_floor_and_direct_dependencies_match_lock_input(self):
        project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))[
            "project"
        ]
        self.assertEqual(project["requires-python"], ">=3.11")
        pins = read_input_pins()
        project_dependencies = {
            re.split(r"==", requirement, maxsplit=1)[0].lower(): requirement
            for requirement in project["dependencies"]
        }
        for name, requirement in project_dependencies.items():
            display_name, version = pins[name]
            self.assertEqual(requirement, f"{display_name}=={version}")
        self.assertEqual(project_dependencies["requests"], "requests==2.34.2")

    def test_every_supported_target_has_a_complete_single_artifact_lock(self):
        expected = set(read_input_pins())
        for python_tag in PYTHON_TAGS:
            for architecture in ARCHITECTURES:
                with self.subTest(python_tag=python_tag, architecture=architecture):
                    path = (
                        ROOT
                        / "locks"
                        / f"windows-cp{python_tag}-{architecture}.txt"
                    )
                    parsed = parse_lock(path)
                    self.assertEqual(set(parsed), expected)
                    self.assertTrue(all(len(hashes) == 1 for _, hashes in parsed.values()))

    def test_aggregate_requirements_has_only_exact_hashed_pins(self):
        parsed = parse_lock(ROOT / "requirements.txt")
        self.assertEqual(set(parsed), set(read_input_pins()))
        self.assertTrue(all(hashes for _, hashes in parsed.values()))

    def test_default_target_mapping_is_explicit(self):
        self.assertEqual(
            default_lock_path((3, 11), "AMD64").name,
            "windows-cp311-x64.txt",
        )
        self.assertEqual(
            default_lock_path((3, 14), "ARM64").name,
            "windows-cp314-arm64.txt",
        )
        with self.assertRaises(ValueError):
            default_lock_path((3, 10), "AMD64")


if __name__ == "__main__":
    unittest.main()
