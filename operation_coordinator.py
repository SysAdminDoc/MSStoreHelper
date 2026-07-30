#!/usr/bin/env python3
"""Typed operation lifecycle, cancellation, and durable result journal."""

from __future__ import annotations

import json
import os
import threading
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Callable


OPERATION_SCHEMA_VERSION = 1
DEFAULT_JOURNAL_LIMIT = 200


def utc_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


class OperationState(str, Enum):
    QUEUED = "queued"
    RUNNING = "running"
    CANCELLING = "cancelling"
    SUCCEEDED = "succeeded"
    PARTIAL = "partial"
    FAILED = "failed"
    CANCELLED = "cancelled"


TERMINAL_STATES = {
    OperationState.SUCCEEDED,
    OperationState.PARTIAL,
    OperationState.FAILED,
    OperationState.CANCELLED,
}
ITEM_STATES = {"succeeded", "skipped", "failed", "cancelled"}
ALLOWED_TRANSITIONS = {
    OperationState.QUEUED: {
        OperationState.RUNNING,
        OperationState.CANCELLING,
        OperationState.FAILED,
        OperationState.CANCELLED,
    },
    OperationState.RUNNING: {
        OperationState.CANCELLING,
        OperationState.SUCCEEDED,
        OperationState.PARTIAL,
        OperationState.FAILED,
        OperationState.CANCELLED,
    },
    OperationState.CANCELLING: {
        OperationState.SUCCEEDED,
        OperationState.PARTIAL,
        OperationState.FAILED,
        OperationState.CANCELLED,
    },
}


class OperationConflictError(RuntimeError):
    """Raised when an exclusive operation is already active."""


class OperationCancelled(RuntimeError):
    """Raised by a worker at a safe cancellation checkpoint."""


@dataclass
class OperationItemResult:
    key: str
    state: str
    message: str = ""
    metadata: dict[str, Any] = field(default_factory=dict)
    updated_at: str = field(default_factory=utc_timestamp)

    def __post_init__(self) -> None:
        if self.state not in ITEM_STATES:
            raise ValueError(f"unsupported operation item state: {self.state}")

    def to_dict(self) -> dict[str, Any]:
        return {
            "Key": self.key,
            "State": self.state,
            "Message": self.message,
            "Metadata": dict(self.metadata),
            "UpdatedAt": self.updated_at,
        }


@dataclass
class OperationResult:
    kind: str
    correlation_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    state: OperationState = OperationState.QUEUED
    created_at: str = field(default_factory=utc_timestamp)
    started_at: str = ""
    completed_at: str = ""
    input_summary: dict[str, Any] = field(default_factory=dict)
    progress: float = 0.0
    progress_message: str = ""
    items: dict[str, OperationItemResult] = field(default_factory=dict)
    error: str = ""

    @property
    def is_terminal(self) -> bool:
        return self.state in TERMINAL_STATES

    @property
    def counts(self) -> dict[str, int]:
        counts = {
            "Total": len(self.items),
            "Succeeded": 0,
            "Skipped": 0,
            "Failed": 0,
            "Cancelled": 0,
        }
        labels = {
            "succeeded": "Succeeded",
            "skipped": "Skipped",
            "failed": "Failed",
            "cancelled": "Cancelled",
        }
        for item in self.items.values():
            counts[labels[item.state]] += 1
        return counts

    @property
    def exit_code(self) -> int:
        return {
            OperationState.SUCCEEDED: 0,
            OperationState.PARTIAL: 2,
            OperationState.FAILED: 1,
            OperationState.CANCELLED: 130,
        }.get(self.state, 1)

    def inferred_terminal_state(self) -> OperationState:
        counts = self.counts
        completed = counts["Succeeded"] + counts["Skipped"]
        if counts["Failed"] and completed:
            return OperationState.PARTIAL
        if counts["Failed"]:
            return OperationState.FAILED
        if counts["Cancelled"]:
            return (
                OperationState.PARTIAL
                if completed
                else OperationState.CANCELLED
            )
        return OperationState.SUCCEEDED

    def to_dict(self) -> dict[str, Any]:
        return {
            "SchemaVersion": OPERATION_SCHEMA_VERSION,
            "CorrelationId": self.correlation_id,
            "Kind": self.kind,
            "State": self.state.value,
            "CreatedAt": self.created_at,
            "StartedAt": self.started_at,
            "CompletedAt": self.completed_at,
            "InputSummary": dict(self.input_summary),
            "Progress": self.progress,
            "ProgressMessage": self.progress_message,
            "Counts": self.counts,
            "Items": [
                self.items[key].to_dict()
                for key in sorted(self.items)
            ],
            "Error": self.error,
            "ExitCode": self.exit_code,
        }


class OperationJournal:
    """Bounded, atomically replaced local operation history."""

    def __init__(self, path: str | os.PathLike[str], *, limit: int = DEFAULT_JOURNAL_LIMIT):
        self.path = Path(path).resolve()
        self.limit = max(1, min(5000, int(limit)))
        self._lock = threading.Lock()

    def _read_unlocked(self) -> list[dict[str, Any]]:
        if not self.path.exists():
            return []
        payload = json.loads(self.path.read_text(encoding="utf-8"))
        if (
            not isinstance(payload, dict)
            or payload.get("SchemaVersion") != OPERATION_SCHEMA_VERSION
            or not isinstance(payload.get("Operations"), list)
        ):
            raise ValueError("operation journal schema is invalid")
        return [
            item for item in payload["Operations"]
            if isinstance(item, dict)
        ]

    def snapshot(self) -> list[dict[str, Any]]:
        with self._lock:
            return self._read_unlocked()

    def append(self, result: OperationResult) -> None:
        if not result.is_terminal:
            raise ValueError("only terminal operation results can be journaled")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._lock:
            operations = self._read_unlocked()
            operations.append(result.to_dict())
            operations = operations[-self.limit :]
            payload = {
                "SchemaVersion": OPERATION_SCHEMA_VERSION,
                "Operations": operations,
            }
            temporary = self.path.with_name(
                f".{self.path.name}.{uuid.uuid4().hex}.tmp"
            )
            try:
                with temporary.open("x", encoding="utf-8", newline="\n") as stream:
                    json.dump(payload, stream, indent=2, sort_keys=True)
                    stream.write("\n")
                    stream.flush()
                    os.fsync(stream.fileno())
                os.replace(temporary, self.path)
            finally:
                try:
                    temporary.unlink()
                except FileNotFoundError:
                    pass


class OperationContext:
    def __init__(self, coordinator: "OperationCoordinator", result: OperationResult):
        self._coordinator = coordinator
        self.result = result

    @property
    def correlation_id(self) -> str:
        return self.result.correlation_id

    @property
    def cancellation_requested(self) -> bool:
        return self._coordinator.cancellation_requested

    @property
    def cancel_event(self) -> threading.Event:
        return self._coordinator._cancel_event

    def cancellation_checkpoint(self) -> None:
        if self.cancellation_requested:
            raise OperationCancelled("operation cancelled at a safe checkpoint")

    def progress(self, value: float, message: str = "") -> None:
        self._coordinator._set_progress(value, message)

    def record(
        self,
        key: str,
        state: str,
        message: str = "",
        metadata: dict[str, Any] | None = None,
    ) -> None:
        self._coordinator._record_item(
            key,
            state,
            message,
            metadata or {},
        )

    def succeeded(self, key: str, message: str = "", **metadata: Any) -> None:
        self.record(key, "succeeded", message, metadata)

    def skipped(self, key: str, message: str = "", **metadata: Any) -> None:
        self.record(key, "skipped", message, metadata)

    def failed(self, key: str, message: str = "", **metadata: Any) -> None:
        self.record(key, "failed", message, metadata)

    def cancelled(self, key: str, message: str = "", **metadata: Any) -> None:
        self.record(key, "cancelled", message, metadata)


OperationWorker = Callable[[OperationContext], OperationState | None]
OperationCallback = Callable[[dict[str, Any]], None]


class OperationCoordinator:
    """Own one exclusive non-daemon operation and its truthful result."""

    def __init__(
        self,
        *,
        journal: OperationJournal | None = None,
        on_change: OperationCallback | None = None,
    ):
        self.journal = journal
        self.on_change = on_change
        self._lock = threading.RLock()
        self._active: OperationResult | None = None
        self._thread: threading.Thread | None = None
        self._cancel_event = threading.Event()

    @property
    def active_result(self) -> OperationResult | None:
        with self._lock:
            return self._active

    @property
    def active_thread(self) -> threading.Thread | None:
        with self._lock:
            return self._thread

    @property
    def cancellation_requested(self) -> bool:
        return self._cancel_event.is_set()

    @property
    def is_active(self) -> bool:
        result = self.active_result
        return bool(result and not result.is_terminal)

    def _emit(self) -> None:
        callback = self.on_change
        result = self._active
        if not callback or not result:
            return
        snapshot = result.to_dict()
        try:
            callback(snapshot)
        except Exception:
            pass

    def _transition(self, state: OperationState, *, error: str = "") -> None:
        with self._lock:
            result = self._active
            if result is None:
                raise RuntimeError("no active operation")
            if state == result.state:
                return
            allowed = ALLOWED_TRANSITIONS.get(result.state, set())
            if state not in allowed:
                raise RuntimeError(
                    f"invalid operation transition: {result.state.value} -> {state.value}"
                )
            result.state = state
            if state == OperationState.RUNNING:
                result.started_at = utc_timestamp()
            if state in TERMINAL_STATES:
                result.completed_at = utc_timestamp()
                result.progress = 1.0
            if error:
                result.error = str(error)
            self._emit()

    def _record_item(
        self,
        key: str,
        state: str,
        message: str,
        metadata: dict[str, Any],
    ) -> None:
        with self._lock:
            if self._active is None:
                raise RuntimeError("no active operation")
            self._active.items[str(key)] = OperationItemResult(
                key=str(key),
                state=state,
                message=str(message),
                metadata=dict(metadata),
            )
            self._emit()

    def _set_progress(self, value: float, message: str) -> None:
        with self._lock:
            if self._active is None:
                raise RuntimeError("no active operation")
            self._active.progress = max(0.0, min(1.0, float(value)))
            self._active.progress_message = str(message)
            self._emit()

    def _reserve(
        self,
        kind: str,
        input_summary: dict[str, Any] | None,
    ) -> OperationResult:
        with self._lock:
            if self._active and not self._active.is_terminal:
                raise OperationConflictError(
                    f"{self._active.kind} operation "
                    f"{self._active.correlation_id} is already {self._active.state.value}"
                )
            self._cancel_event = threading.Event()
            self._active = OperationResult(
                kind=str(kind),
                input_summary=dict(input_summary or {}),
            )
            self._thread = None
            self._emit()
            return self._active

    def _finish(self, result: OperationResult) -> None:
        if self.journal:
            try:
                self.journal.append(result)
            except (OSError, ValueError, json.JSONDecodeError) as exc:
                result.input_summary["JournalError"] = str(exc)
                self._emit()

    def _run_reserved(self, worker: OperationWorker) -> OperationResult:
        result = self.active_result
        if result is None:
            raise RuntimeError("no reserved operation")
        context = OperationContext(self, result)
        try:
            if (
                self.cancellation_requested
                and result.state == OperationState.CANCELLING
            ):
                self._transition(OperationState.CANCELLED)
                self._finish(result)
                return result
            self._transition(OperationState.RUNNING)
            requested_state = worker(context)
            if requested_state is not None and requested_state not in TERMINAL_STATES:
                raise ValueError("worker returned a non-terminal operation state")
            if requested_state is None:
                requested_state = (
                    OperationState.CANCELLED
                    if self.cancellation_requested and not result.items
                    else result.inferred_terminal_state()
                )
            self._transition(requested_state)
        except OperationCancelled as exc:
            self._transition(OperationState.CANCELLED, error=str(exc))
        except Exception as exc:
            try:
                self._record_item("operation", "failed", str(exc), {})
                self._transition(OperationState.FAILED, error=str(exc))
            except RuntimeError:
                result.error = str(exc)
        self._finish(result)
        return result

    def start(
        self,
        kind: str,
        worker: OperationWorker,
        *,
        input_summary: dict[str, Any] | None = None,
    ) -> OperationResult:
        result = self._reserve(kind, input_summary)
        thread = threading.Thread(
            target=self._run_reserved,
            args=(worker,),
            name=f"MSStoreHelper-{kind}-{result.correlation_id[:8]}",
            daemon=False,
        )
        with self._lock:
            self._thread = thread
        thread.start()
        return result

    def run(
        self,
        kind: str,
        worker: OperationWorker,
        *,
        input_summary: dict[str, Any] | None = None,
    ) -> OperationResult:
        self._reserve(kind, input_summary)
        return self._run_reserved(worker)

    def cancel(self) -> bool:
        with self._lock:
            if not self._active or self._active.is_terminal:
                return False
            self._cancel_event.set()
            if self._active.state in {
                OperationState.QUEUED,
                OperationState.RUNNING,
            }:
                self._transition(OperationState.CANCELLING)
            return True

    def wait(self, timeout: float | None = None) -> OperationResult | None:
        thread = self.active_thread
        if thread and thread is not threading.current_thread():
            thread.join(timeout)
        return self.active_result

    def shutdown(self, timeout: float = 30.0) -> bool:
        self.cancel()
        thread = self.active_thread
        if thread and thread is not threading.current_thread():
            thread.join(max(0.0, float(timeout)))
            return not thread.is_alive()
        return True
