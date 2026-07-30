#!/usr/bin/env python3

import json
import os
import subprocess
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

from state_repository import (
    JsonStateSpec,
    atomic_write_json,
    load_json_state,
    pop_recovery_notices,
    save_json_state,
    update_json_state,
)


def migrate_zero_to_one(payload):
    return {
        "SchemaVersion": 1,
        "Items": list(payload.get("Items") or []),
    }


SPEC = JsonStateSpec(
    name="test state",
    current_version=1,
    default_factory=lambda: {"Items": []},
    migrations={0: migrate_zero_to_one},
    validator=lambda value: isinstance(value.get("Items"), list),
)


class StateRepositoryTests(unittest.TestCase):
    def setUp(self):
        pop_recovery_notices()

    def test_legacy_state_migrates_once_and_is_persisted(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "state.json")
            with open(path, "w", encoding="utf-8") as stream:
                json.dump({"Items": ["legacy"]}, stream)

            loaded = load_json_state(path, SPEC)
            reloaded = load_json_state(path, SPEC)

        self.assertTrue(loaded.migrated)
        self.assertFalse(reloaded.migrated)
        self.assertEqual(reloaded.data["SchemaVersion"], 1)
        self.assertEqual(reloaded.data["Items"], ["legacy"])

    def test_corrupt_state_is_quarantined_with_recovery_notice(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "state.json")
            with open(path, "w", encoding="utf-8") as stream:
                stream.write("{truncated")

            loaded = load_json_state(path, SPEC)
            notices = pop_recovery_notices()

            self.assertEqual(loaded.data["Items"], [])
            self.assertIsNotNone(loaded.recovery)
            self.assertFalse(os.path.exists(path))
            self.assertTrue(os.path.exists(loaded.recovery.quarantine_path))
            self.assertEqual(notices, [loaded.recovery])
            self.assertIn("JSONDecodeError", loaded.recovery.reason)

    def test_newer_or_invalid_schema_is_quarantined_not_overwritten(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "state.json")
            atomic_write_json(
                path,
                {"SchemaVersion": 99, "Items": ["future"]},
            )

            loaded = load_json_state(path, SPEC)

            with open(
                loaded.recovery.quarantine_path,
                "r",
                encoding="utf-8",
            ) as stream:
                quarantined = json.load(stream)
        self.assertEqual(quarantined["Items"], ["future"])
        self.assertEqual(loaded.data["Items"], [])

    def test_failed_replace_preserves_last_valid_generation(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "state.json")
            save_json_state(path, {"Items": ["valid"]}, SPEC)

            with patch(
                "state_repository.os.replace",
                side_effect=OSError("simulated crash"),
            ):
                with self.assertRaisesRegex(OSError, "simulated crash"):
                    save_json_state(path, {"Items": ["new"]}, SPEC)

            loaded = load_json_state(path, SPEC)
            temporary = [
                name
                for name in os.listdir(folder)
                if name.endswith(".tmp")
            ]

        self.assertEqual(loaded.data["Items"], ["valid"])
        self.assertEqual(temporary, [])

    def test_threaded_updates_do_not_lose_generations(self):
        spec = JsonStateSpec(
            name="counter",
            current_version=1,
            default_factory=lambda: {"Count": 0},
            migrations={0: lambda _value: {
                "SchemaVersion": 1,
                "Count": 0,
            }},
            validator=lambda value: isinstance(value.get("Count"), int),
        )
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "counter.json")

            def increment():
                for _index in range(25):
                    update_json_state(
                        path,
                        spec,
                        lambda value: {
                            **value,
                            "Count": value["Count"] + 1,
                        },
                    )

            threads = [
                threading.Thread(target=increment)
                for _index in range(4)
            ]
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join()
            count = load_json_state(path, spec).data["Count"]

        self.assertEqual(count, 100)

    def test_cross_process_updates_use_one_interprocess_lock(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "counter.json")
            code = "\n".join([
                "import sys",
                "from state_repository import JsonStateSpec, update_json_state",
                "path = sys.argv[1]",
                "spec = JsonStateSpec(",
                "    name='counter',",
                "    current_version=1,",
                "    default_factory=lambda: {'Count': 0},",
                "    migrations={0: lambda _value: {'SchemaVersion': 1, 'Count': 0}},",
                "    validator=lambda value: isinstance(value.get('Count'), int),",
                ")",
                "for _index in range(10):",
                "    update_json_state(path, spec, lambda value: {**value, 'Count': value['Count'] + 1})",
            ])
            processes = [
                subprocess.Popen(
                    [sys.executable, "-c", code, path],
                    cwd=os.path.dirname(os.path.dirname(__file__)),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                )
                for _index in range(4)
            ]
            failures = []
            for process in processes:
                stdout, stderr = process.communicate(timeout=30)
                if process.returncode:
                    failures.append((process.returncode, stdout, stderr))
            with open(path, "r", encoding="utf-8") as stream:
                payload = json.load(stream)

        self.assertEqual(failures, [])
        self.assertEqual(payload["Count"], 40)
        self.assertEqual(payload["SchemaVersion"], 1)


if __name__ == "__main__":
    unittest.main()
