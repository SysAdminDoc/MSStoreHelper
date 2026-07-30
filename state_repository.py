#!/usr/bin/env python3
"""Versioned, crash-safe JSON state with cross-process serialization."""

from __future__ import annotations

import copy
import json
import os
import threading
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable


DEFAULT_LOCK_TIMEOUT_SECONDS = 10.0


class StateRepositoryError(RuntimeError):
    """Raised when state cannot be read or committed safely."""


class StateLockTimeout(StateRepositoryError):
    """Raised when another process retains a state lock past the deadline."""


Migration = Callable[[dict[str, Any]], dict[str, Any]]
Validator = Callable[[dict[str, Any]], bool]


@dataclass(frozen=True)
class JsonStateSpec:
    name: str
    current_version: int
    default_factory: Callable[[], dict[str, Any]]
    migrations: dict[int, Migration] = field(default_factory=dict)
    validator: Validator | None = None

    def default(self) -> dict[str, Any]:
        value = copy.deepcopy(self.default_factory())
        if not isinstance(value, dict):
            raise StateRepositoryError(
                f"{self.name} default factory did not return an object"
            )
        value["SchemaVersion"] = self.current_version
        return value


@dataclass(frozen=True)
class StateRecoveryNotice:
    state_name: str
    original_path: str
    quarantine_path: str
    reason: str
    quarantined_at: str

    def to_dict(self) -> dict[str, str]:
        return {
            "StateName": self.state_name,
            "OriginalPath": self.original_path,
            "QuarantinePath": self.quarantine_path,
            "Reason": self.reason,
            "QuarantinedAt": self.quarantined_at,
        }


@dataclass(frozen=True)
class StateLoadResult:
    data: dict[str, Any]
    recovery: StateRecoveryNotice | None = None
    migrated: bool = False


_PROCESS_LOCKS_GUARD = threading.Lock()
_PROCESS_LOCKS: dict[str, threading.RLock] = {}
_RECOVERY_GUARD = threading.Lock()
_RECOVERY_NOTICES: list[StateRecoveryNotice] = []


def _process_lock(path: str) -> threading.RLock:
    key = os.path.normcase(os.path.abspath(path))
    with _PROCESS_LOCKS_GUARD:
        return _PROCESS_LOCKS.setdefault(key, threading.RLock())


class InterProcessFileLock:
    """Blocking same-process and OS-level exclusive lock for one state file."""

    def __init__(
        self,
        target_path: str | os.PathLike[str],
        *,
        timeout: float = DEFAULT_LOCK_TIMEOUT_SECONDS,
    ):
        self.target_path = os.path.abspath(os.fspath(target_path))
        self.lock_path = f"{self.target_path}.lock"
        self.timeout = max(0.0, float(timeout))
        self._process_lock = _process_lock(self.target_path)
        self._handle = None
        self._os_locked = False

    def __enter__(self):
        if not self._process_lock.acquire(timeout=self.timeout):
            raise StateLockTimeout(
                f"Timed out waiting for state lock: {self.target_path}"
            )
        try:
            os.makedirs(os.path.dirname(self.target_path), exist_ok=True)
            self._handle = open(self.lock_path, "a+b")
            self._handle.seek(0, os.SEEK_END)
            if self._handle.tell() == 0:
                self._handle.write(b"\0")
                self._handle.flush()
                os.fsync(self._handle.fileno())
            deadline = time.monotonic() + self.timeout
            while True:
                self._handle.seek(0)
                try:
                    if os.name == "nt":
                        import msvcrt

                        msvcrt.locking(
                            self._handle.fileno(),
                            msvcrt.LK_NBLCK,
                            1,
                        )
                    else:
                        import fcntl

                        fcntl.flock(
                            self._handle.fileno(),
                            fcntl.LOCK_EX | fcntl.LOCK_NB,
                        )
                    self._os_locked = True
                    return self
                except OSError as exc:
                    if time.monotonic() >= deadline:
                        raise StateLockTimeout(
                            f"Timed out waiting for state lock: "
                            f"{self.target_path}"
                        ) from exc
                    time.sleep(0.025)
        except Exception:
            if self._handle is not None:
                self._handle.close()
                self._handle = None
            self._process_lock.release()
            raise

    def __exit__(self, _exc_type, _exc, _traceback):
        try:
            if self._handle is not None and self._os_locked:
                self._handle.seek(0)
                if os.name == "nt":
                    import msvcrt

                    msvcrt.locking(
                        self._handle.fileno(),
                        msvcrt.LK_UNLCK,
                        1,
                    )
                else:
                    import fcntl

                    fcntl.flock(self._handle.fileno(), fcntl.LOCK_UN)
        finally:
            self._os_locked = False
            if self._handle is not None:
                self._handle.close()
                self._handle = None
            self._process_lock.release()


def _atomic_write_bytes_locked(path: str, content: bytes) -> str:
    path = os.path.abspath(path)
    folder = os.path.dirname(path)
    os.makedirs(folder, exist_ok=True)
    temporary = os.path.join(
        folder,
        f".{os.path.basename(path)}.{uuid.uuid4().hex}.tmp",
    )
    try:
        with open(temporary, "xb") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        try:
            os.remove(temporary)
        except FileNotFoundError:
            pass
    return path


def atomic_write_bytes(
    path: str | os.PathLike[str],
    content: bytes,
    *,
    lock_timeout: float = DEFAULT_LOCK_TIMEOUT_SECONDS,
) -> str:
    path = os.path.abspath(os.fspath(path))
    with InterProcessFileLock(path, timeout=lock_timeout):
        return _atomic_write_bytes_locked(path, bytes(content))


def _json_bytes(payload: Any) -> bytes:
    return (
        json.dumps(
            payload,
            indent=2,
            sort_keys=True,
            ensure_ascii=False,
        )
        + "\n"
    ).encode("utf-8")


def atomic_write_json(
    path: str | os.PathLike[str],
    payload: Any,
    *,
    lock_timeout: float = DEFAULT_LOCK_TIMEOUT_SECONDS,
) -> str:
    path = os.path.abspath(os.fspath(path))
    with InterProcessFileLock(path, timeout=lock_timeout):
        return _atomic_write_bytes_locked(path, _json_bytes(payload))


def append_jsonl(
    path: str | os.PathLike[str],
    payload: Any,
    *,
    lock_timeout: float = DEFAULT_LOCK_TIMEOUT_SECONDS,
) -> str:
    path = os.path.abspath(os.fspath(path))
    with InterProcessFileLock(path, timeout=lock_timeout):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "a", encoding="utf-8", newline="\n") as stream:
            stream.write(
                json.dumps(payload, sort_keys=True, ensure_ascii=False)
            )
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
    return path


def remove_state_file(
    path: str | os.PathLike[str],
    *,
    lock_timeout: float = DEFAULT_LOCK_TIMEOUT_SECONDS,
) -> bool:
    path = os.path.abspath(os.fspath(path))
    with InterProcessFileLock(path, timeout=lock_timeout):
        try:
            os.remove(path)
            return True
        except FileNotFoundError:
            return False


def _state_version(payload: dict[str, Any]) -> int:
    value = payload.get("SchemaVersion", payload.get("Version", 0))
    if isinstance(value, bool):
        raise ValueError("state schema version is invalid")
    return int(value)


def _migrate(payload: dict[str, Any], spec: JsonStateSpec) -> tuple[dict[str, Any], bool]:
    version = _state_version(payload)
    if version > spec.current_version:
        raise ValueError(
            f"schema {version} is newer than supported "
            f"{spec.current_version}"
        )
    migrated = False
    while version < spec.current_version:
        migration = spec.migrations.get(version)
        if migration is None:
            raise ValueError(
                f"no deterministic migration from schema {version}"
            )
        payload = migration(copy.deepcopy(payload))
        if not isinstance(payload, dict):
            raise ValueError(
                f"migration from schema {version} returned invalid state"
            )
        next_version = _state_version(payload)
        if next_version != version + 1:
            raise ValueError(
                f"migration from schema {version} did not advance exactly once"
            )
        version = next_version
        migrated = True
    payload["SchemaVersion"] = spec.current_version
    if spec.validator and not spec.validator(payload):
        raise ValueError("state failed schema validation")
    return payload, migrated


def _quarantine_locked(
    path: str,
    spec: JsonStateSpec,
    reason: str,
) -> StateRecoveryNotice:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    source = Path(path)
    quarantine = source.with_name(
        f"{source.stem}.corrupt-{stamp}-{uuid.uuid4().hex[:8]}"
        f"{source.suffix or '.json'}"
    )
    os.replace(path, quarantine)
    notice = StateRecoveryNotice(
        state_name=spec.name,
        original_path=path,
        quarantine_path=str(quarantine),
        reason=reason,
        quarantined_at=datetime.now(timezone.utc).isoformat(),
    )
    with _RECOVERY_GUARD:
        _RECOVERY_NOTICES.append(notice)
    return notice


def _load_locked(path: str, spec: JsonStateSpec) -> StateLoadResult:
    if not os.path.exists(path):
        return StateLoadResult(spec.default())
    try:
        with open(path, "r", encoding="utf-8") as stream:
            payload = json.load(stream)
        if not isinstance(payload, dict):
            raise ValueError("state root must be an object")
        payload, migrated = _migrate(payload, spec)
        if migrated:
            _atomic_write_bytes_locked(path, _json_bytes(payload))
        return StateLoadResult(payload, migrated=migrated)
    except (OSError, TypeError, ValueError, json.JSONDecodeError) as exc:
        notice = _quarantine_locked(
            path,
            spec,
            f"{type(exc).__name__}: {exc}",
        )
        return StateLoadResult(spec.default(), recovery=notice)


def load_json_state(
    path: str | os.PathLike[str],
    spec: JsonStateSpec,
    *,
    lock_timeout: float = DEFAULT_LOCK_TIMEOUT_SECONDS,
) -> StateLoadResult:
    path = os.path.abspath(os.fspath(path))
    with InterProcessFileLock(path, timeout=lock_timeout):
        return _load_locked(path, spec)


def save_json_state(
    path: str | os.PathLike[str],
    payload: dict[str, Any],
    spec: JsonStateSpec,
    *,
    lock_timeout: float = DEFAULT_LOCK_TIMEOUT_SECONDS,
) -> str:
    path = os.path.abspath(os.fspath(path))
    value = copy.deepcopy(payload)
    value["SchemaVersion"] = spec.current_version
    value, _migrated = _migrate(value, spec)
    with InterProcessFileLock(path, timeout=lock_timeout):
        return _atomic_write_bytes_locked(path, _json_bytes(value))


def update_json_state(
    path: str | os.PathLike[str],
    spec: JsonStateSpec,
    update: Callable[[dict[str, Any]], dict[str, Any] | None],
    *,
    lock_timeout: float = DEFAULT_LOCK_TIMEOUT_SECONDS,
) -> StateLoadResult:
    path = os.path.abspath(os.fspath(path))
    with InterProcessFileLock(path, timeout=lock_timeout):
        loaded = _load_locked(path, spec)
        value = copy.deepcopy(loaded.data)
        replacement = update(value)
        if replacement is not None:
            value = replacement
        if not isinstance(value, dict):
            raise StateRepositoryError(
                f"{spec.name} update did not return an object"
            )
        value["SchemaVersion"] = spec.current_version
        value, _migrated = _migrate(value, spec)
        _atomic_write_bytes_locked(path, _json_bytes(value))
        return StateLoadResult(
            value,
            recovery=loaded.recovery,
            migrated=loaded.migrated,
        )


def pop_recovery_notices() -> list[StateRecoveryNotice]:
    with _RECOVERY_GUARD:
        notices = list(_RECOVERY_NOTICES)
        _RECOVERY_NOTICES.clear()
    return notices
