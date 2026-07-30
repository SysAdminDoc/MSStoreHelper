#!/usr/bin/env python3
"""Subprocess execution with mandatory deadlines and deterministic termination."""

from __future__ import annotations

import subprocess
import threading
import time
from collections.abc import Mapping, Sequence


DEFAULT_COMMAND_TIMEOUT = 90.0
TERMINATION_GRACE_SECONDS = 3.0


class CommandExecutionError(RuntimeError):
    def __init__(self, message: str, *, stdout: str = "", stderr: str = ""):
        super().__init__(message)
        self.stdout = stdout
        self.stderr = stderr


class CommandTimeoutError(CommandExecutionError):
    pass


class CommandCancelledError(CommandExecutionError):
    pass


def _terminate(process: subprocess.Popen[str]) -> tuple[str, str]:
    if process.poll() is None:
        process.terminate()
    try:
        return process.communicate(timeout=TERMINATION_GRACE_SECONDS)
    except subprocess.TimeoutExpired:
        process.kill()
        return process.communicate()


def run_command(
    args: Sequence[str],
    *,
    timeout: float = DEFAULT_COMMAND_TIMEOUT,
    cancel_event: threading.Event | None = None,
    env: Mapping[str, str] | None = None,
    cwd: str | None = None,
    creationflags: int | None = None,
    check: bool = False,
) -> subprocess.CompletedProcess[str]:
    if not args or not all(isinstance(value, str) and value for value in args):
        raise ValueError("command arguments must be non-empty strings")
    timeout = float(timeout)
    if timeout <= 0:
        raise ValueError("command timeout must be positive")
    flags = (
        getattr(subprocess, "CREATE_NO_WINDOW", 0)
        if creationflags is None
        else int(creationflags)
    )
    process = subprocess.Popen(
        list(args),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env=dict(env) if env is not None else None,
        cwd=cwd,
        creationflags=flags,
    )
    deadline = time.monotonic() + timeout
    stdout = ""
    stderr = ""
    while True:
        if cancel_event and cancel_event.is_set():
            stdout, stderr = _terminate(process)
            raise CommandCancelledError(
                "command cancelled and terminated",
                stdout=stdout,
                stderr=stderr,
            )
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            stdout, stderr = _terminate(process)
            raise CommandTimeoutError(
                f"command exceeded its {timeout:g}-second deadline and was terminated",
                stdout=stdout,
                stderr=stderr,
            )
        try:
            stdout, stderr = process.communicate(timeout=min(0.1, remaining))
            break
        except subprocess.TimeoutExpired:
            continue

    completed = subprocess.CompletedProcess(
        list(args),
        process.returncode,
        stdout,
        stderr,
    )
    if check and completed.returncode:
        raise subprocess.CalledProcessError(
            completed.returncode,
            completed.args,
            output=completed.stdout,
            stderr=completed.stderr,
        )
    return completed
