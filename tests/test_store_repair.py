#!/usr/bin/env python3

import json
import os
import shutil
import tempfile
import threading
import unittest
import uuid
from pathlib import Path
from unittest.mock import patch

import repair_transaction as repair


WINDOWS_POWERSHELL = (
    r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
)


class StoreRepairPlanTests(unittest.TestCase):
    def test_production_plans_are_exact_and_fail_closed(self):
        with tempfile.TemporaryDirectory() as backup_base:
            for repair_type in (
                "store-repair",
                "provisioning-repair",
                "licensing-reset",
                "cache-rebuild",
            ):
                plan = repair.build_repair_plan(
                    repair_type,
                    backup_base=backup_base,
                )
                rendered = repair.render_repair_plan(plan)

                self.assertEqual(
                    uuid.UUID(plan["OperationId"]).version,
                    4,
                )
                self.assertTrue(plan["ConfirmationToken"])
                self.assertTrue(plan["Impact"])
                self.assertTrue(plan["Permissions"])
                self.assertIn("Administrator:", rendered)
                self.assertIn("Reboot:", rendered)
                self.assertIn("Backups before mutation:", rendered)
                self.assertIn("Preconditions:", rendered)
                self.assertIn("Mutation steps:", rendered)
                self.assertNotIn("SilentlyContinue } catch", rendered)
                for step in plan["Steps"]:
                    self.assertIn("-ErrorAction Stop", step["Command"])

    def test_confirmation_token_is_mandatory(self):
        with tempfile.TemporaryDirectory() as sandbox:
            state = Path(sandbox, "state")
            state.mkdir()
            with tempfile.TemporaryDirectory() as backup_base:
                plan = repair.build_sandbox_repair_plan(
                    sandbox,
                    backup_base=backup_base,
                )
                with self.assertRaisesRegex(
                    repair.RepairTransactionError,
                    "confirmation",
                ):
                    repair.execute_repair_plan(
                        plan,
                        confirmation_token="wrong",
                        powershell_exe=WINDOWS_POWERSHELL,
                        secure_backup=False,
                    )

    def test_legacy_best_effort_runner_is_disabled(self):
        from MSStoreHelper import StoreAPI

        with self.assertRaisesRegex(
            repair.RepairTransactionError,
            "explicitly confirm",
        ):
            StoreAPI._run_powershell_steps([
                ("unsafe", "Write-Output 'should not run'"),
            ])

    def test_operation_lock_is_exclusive_and_uuid_owned(self):
        with tempfile.TemporaryDirectory() as backup_base:
            owner = str(uuid.uuid4())
            with repair.RepairOperationLock(backup_base, owner):
                with self.assertRaises(repair.RepairLockError):
                    with repair.RepairOperationLock(
                        backup_base,
                        str(uuid.uuid4()),
                    ):
                        self.fail("The second operation acquired the lock")
            lock_record = json.loads(
                Path(backup_base, repair.LOCK_FILENAME).read_text(
                    encoding="utf-8"
                )
            )
            self.assertEqual(lock_record["OperationId"], owner)

    def test_retention_prunes_only_transaction_directories(self):
        with tempfile.TemporaryDirectory() as backup_base:
            roots = []
            for index in range(3):
                root = Path(backup_base, f"repair-{index}")
                root.mkdir()
                manifest = {
                    "CompletedAt": f"2026-07-2{index}T00:00:00+00:00",
                }
                Path(root, repair.MANIFEST_FILENAME).write_text(
                    json.dumps(manifest),
                    encoding="utf-8",
                )
                roots.append(root)
            unrelated = Path(backup_base, "unrelated")
            unrelated.mkdir()

            removed = repair._prune_repair_backups(
                backup_base,
                retention_count=1,
            )

            self.assertEqual(len(removed), 2)
            self.assertTrue(roots[-1].is_dir())
            self.assertTrue(unrelated.is_dir())


@unittest.skipUnless(
    os.name == "nt" and os.path.isfile(WINDOWS_POWERSHELL),
    "Real repair transaction tests require Windows PowerShell",
)
class StoreRepairWindowsSandboxTests(unittest.TestCase):
    def _create_sandbox(self, root):
        sandbox = Path(root, "sandbox")
        state = Path(sandbox, "state")
        nested = Path(state, "nested")
        nested.mkdir(parents=True)
        Path(state, "state.txt").write_text(
            "original",
            encoding="utf-8",
        )
        Path(nested, "data.bin").write_bytes(b"\x00\x01original")
        return sandbox, state

    def test_real_repair_and_repeatable_restore_preserve_backup(self):
        with tempfile.TemporaryDirectory() as temp_root:
            sandbox, state = self._create_sandbox(temp_root)
            backup_base = Path(temp_root, "backups")
            backup_base.mkdir()
            plan = repair.build_sandbox_repair_plan(
                sandbox,
                backup_base=backup_base,
                retention_count=3,
            )

            context = repair.execute_repair_plan(
                plan,
                confirmation_token=plan["ConfirmationToken"],
                powershell_exe=WINDOWS_POWERSHELL,
                secure_backup=True,
            )

            self.assertEqual(context["Outcome"], "succeeded")
            self.assertEqual(
                Path(state, "state.txt").read_text(encoding="utf-8"),
                "mutated",
            )
            restore_plan = repair.build_restore_plan(
                context["BackupRoot"],
                backup_base=backup_base,
                allow_sandbox=True,
            )
            self.assertIn(
                "backups are retained",
                repair.render_restore_plan(restore_plan),
            )
            backup_path = Path(
                restore_plan["RestoreTargets"][0]["BackupPath"]
            )
            backup_digest = repair._filesystem_inventory(
                backup_path
            )["Digest"]

            first_restore = repair.execute_restore_plan(
                restore_plan,
                confirmation_token=restore_plan["ConfirmationToken"],
                powershell_exe=WINDOWS_POWERSHELL,
                secure_backup=True,
            )
            self.assertEqual(first_restore["Outcome"], "succeeded")
            self.assertEqual(
                Path(state, "state.txt").read_text(encoding="utf-8"),
                "original",
            )

            Path(state, "state.txt").write_text(
                "changed again",
                encoding="utf-8",
            )
            second_restore = repair.execute_restore_plan(
                restore_plan,
                confirmation_token=restore_plan["ConfirmationToken"],
                powershell_exe=WINDOWS_POWERSHELL,
                secure_backup=True,
            )

            self.assertEqual(second_restore["Outcome"], "succeeded")
            self.assertEqual(
                Path(state, "state.txt").read_text(encoding="utf-8"),
                "original",
            )
            self.assertEqual(
                repair._filesystem_inventory(backup_path)["Digest"],
                backup_digest,
            )
            history = Path(
                restore_plan["RestoreHistoryPath"]
            ).read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(history), 2)

    def test_precondition_failure_stops_before_mutation(self):
        with tempfile.TemporaryDirectory() as temp_root:
            sandbox, state = self._create_sandbox(temp_root)
            backup_base = Path(temp_root, "backups")
            backup_base.mkdir()
            plan = repair.build_sandbox_repair_plan(
                sandbox,
                backup_base=backup_base,
            )
            plan["Preconditions"][0]["Command"] = (
                "throw 'intentional precondition failure'"
            )

            context = repair.execute_repair_plan(
                plan,
                confirmation_token=plan["ConfirmationToken"],
                powershell_exe=WINDOWS_POWERSHELL,
                secure_backup=True,
            )

            self.assertEqual(context["Outcome"], "preflight-failed")
            self.assertFalse(context["MutationStarted"])
            self.assertEqual(
                Path(state, "state.txt").read_text(encoding="utf-8"),
                "original",
            )

    def test_cancellation_is_observed_at_a_safe_checkpoint(self):
        with tempfile.TemporaryDirectory() as temp_root:
            sandbox, state = self._create_sandbox(temp_root)
            backup_base = Path(temp_root, "backups")
            backup_base.mkdir()
            plan = repair.build_sandbox_repair_plan(
                sandbox,
                backup_base=backup_base,
            )
            cancel_event = threading.Event()
            cancel_event.set()

            context = repair.execute_repair_plan(
                plan,
                confirmation_token=plan["ConfirmationToken"],
                powershell_exe=WINDOWS_POWERSHELL,
                cancel_event=cancel_event,
                secure_backup=True,
            )

            self.assertEqual(context["Outcome"], "cancelled")
            self.assertFalse(context["MutationStarted"])
            self.assertEqual(
                Path(state, "state.txt").read_text(encoding="utf-8"),
                "original",
            )

    def test_tampered_backup_is_rejected_before_restore(self):
        with tempfile.TemporaryDirectory() as temp_root:
            sandbox, _state = self._create_sandbox(temp_root)
            backup_base = Path(temp_root, "backups")
            backup_base.mkdir()
            plan = repair.build_sandbox_repair_plan(
                sandbox,
                backup_base=backup_base,
            )
            context = repair.execute_repair_plan(
                plan,
                confirmation_token=plan["ConfirmationToken"],
                powershell_exe=WINDOWS_POWERSHELL,
                secure_backup=True,
            )
            backup_path = Path(
                context["BackupRoot"],
                "files",
                "sandbox-state",
                "state.txt",
            )
            backup_path.write_text("tampered", encoding="utf-8")

            with self.assertRaisesRegex(
                repair.RepairTransactionError,
                "hash",
            ):
                repair.build_restore_plan(
                    context["BackupRoot"],
                    backup_base=backup_base,
                    allow_sandbox=True,
                )

    def test_insufficient_disk_space_fails_before_mutation(self):
        with tempfile.TemporaryDirectory() as temp_root:
            sandbox, state = self._create_sandbox(temp_root)
            backup_base = Path(temp_root, "backups")
            backup_base.mkdir()
            plan = repair.build_sandbox_repair_plan(
                sandbox,
                backup_base=backup_base,
            )
            fake_usage = shutil._ntuple_diskusage(100, 100, 0)

            with patch(
                "repair_transaction.shutil.disk_usage",
                return_value=fake_usage,
            ):
                context = repair.execute_repair_plan(
                    plan,
                    confirmation_token=plan["ConfirmationToken"],
                    powershell_exe=WINDOWS_POWERSHELL,
                    secure_backup=True,
                )

            self.assertEqual(context["Outcome"], "preflight-failed")
            self.assertFalse(context["MutationStarted"])
            self.assertEqual(
                Path(state, "state.txt").read_text(encoding="utf-8"),
                "original",
            )


if __name__ == "__main__":
    unittest.main()
