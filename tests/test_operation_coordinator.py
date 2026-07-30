#!/usr/bin/env python3

import json
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path

from command_runner import (
    CommandCancelledError,
    CommandTimeoutError,
    run_command,
)
from operation_coordinator import (
    OperationConflictError,
    OperationCoordinator,
    OperationJournal,
    OperationState,
)


class OperationCoordinatorTests(unittest.TestCase):
    def test_partial_result_has_correlation_id_exit_code_and_journal(self):
        changes = []
        with tempfile.TemporaryDirectory() as folder:
            journal = OperationJournal(Path(folder, "operations.json"))
            coordinator = OperationCoordinator(
                journal=journal,
                on_change=changes.append,
            )

            def worker(context):
                context.succeeded("first", "downloaded")
                context.failed("second", "trust blocked")

            result = coordinator.run(
                "download",
                worker,
                input_summary={"PackageCount": 2},
            )

            self.assertEqual(result.state, OperationState.PARTIAL)
            self.assertEqual(result.exit_code, 2)
            self.assertRegex(
                result.correlation_id,
                r"^[0-9a-f]{8}-[0-9a-f-]{27}$",
            )
            self.assertEqual(
                [change["State"] for change in changes],
                ["queued", "running", "running", "running", "partial"],
            )
            history = journal.snapshot()
            self.assertEqual(history[0]["CorrelationId"], result.correlation_id)
            self.assertEqual(history[0]["State"], "partial")
            self.assertEqual(history[0]["Counts"]["Failed"], 1)

    def test_conflicting_operation_is_rejected_and_worker_is_non_daemon(self):
        release = threading.Event()
        started = threading.Event()
        coordinator = OperationCoordinator()

        def worker(context):
            started.set()
            release.wait(2)
            context.succeeded("one")

        coordinator.start("download", worker)
        self.assertTrue(started.wait(1))
        self.assertFalse(coordinator.active_thread.daemon)
        with self.assertRaises(OperationConflictError):
            coordinator.start("install", worker)
        release.set()
        result = coordinator.wait(2)
        self.assertEqual(result.state, OperationState.SUCCEEDED)

    def test_cancellation_and_shutdown_finish_at_safe_checkpoint(self):
        coordinator = OperationCoordinator()

        def worker(context):
            while True:
                context.cancellation_checkpoint()
                time.sleep(0.01)

        coordinator.start("repair", worker)
        self.assertTrue(coordinator.shutdown(timeout=2))
        self.assertEqual(
            coordinator.active_result.state,
            OperationState.CANCELLED,
        )
        self.assertEqual(coordinator.active_result.exit_code, 130)

    def test_journal_is_bounded_and_valid_json(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder, "operations.json")
            coordinator = OperationCoordinator(
                journal=OperationJournal(path, limit=2)
            )
            for index in range(3):
                coordinator.run(
                    f"operation-{index}",
                    lambda context, value=index: context.succeeded(str(value)),
                )
            payload = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(payload["SchemaVersion"], 1)
            self.assertEqual(len(payload["Operations"]), 2)
            self.assertEqual(payload["Operations"][-1]["Kind"], "operation-2")


class CommandRunnerTests(unittest.TestCase):
    def test_command_returns_captured_output(self):
        result = run_command(
            [sys.executable, "-c", "print('ready')"],
            timeout=5,
        )
        self.assertEqual(result.returncode, 0)
        self.assertEqual(result.stdout.strip(), "ready")

    def test_timeout_terminates_process(self):
        started = time.monotonic()
        with self.assertRaises(CommandTimeoutError):
            run_command(
                [sys.executable, "-c", "import time; time.sleep(30)"],
                timeout=0.2,
            )
        self.assertLess(time.monotonic() - started, 5)

    def test_cancellation_terminates_process(self):
        cancel = threading.Event()
        timer = threading.Timer(0.2, cancel.set)
        timer.start()
        try:
            with self.assertRaises(CommandCancelledError):
                run_command(
                    [sys.executable, "-c", "import time; time.sleep(30)"],
                    timeout=10,
                    cancel_event=cancel,
                )
        finally:
            timer.cancel()


if __name__ == "__main__":
    unittest.main()
