#!/usr/bin/env python3
"""Fail-closed, evidence-backed repair and restore transactions."""

import ctypes
import hashlib
import hmac
import json
import os
import re
import secrets
import shutil
import subprocess
import threading
import uuid
from datetime import datetime, timezone

from command_runner import run_command


REPAIR_SCHEMA_VERSION = 1
DEFAULT_REPAIR_RETENTION = 10
MIN_REPAIR_RETENTION = 1
MAX_REPAIR_RETENTION = 50
MIN_BACKUP_FREE_BYTES = 64 * 1024 * 1024
REGISTRY_BACKUP_RESERVE_BYTES = 16 * 1024 * 1024
LOCK_FILENAME = ".repair-operation.lock"
MANIFEST_FILENAME = "repair-manifest.json"
BACKUP_RECORDS_FILENAME = "backup-records.jsonl"
BASELINE_STATE_FILENAME = "baseline-state.json"
RESTORE_HISTORY_FILENAME = "restore-history.jsonl"


class RepairTransactionError(RuntimeError):
    """Raised when a repair cannot proceed without violating its contract."""


class RepairLockError(RepairTransactionError):
    """Raised when another repair or restore owns the operation lock."""


class RepairCancelled(RepairTransactionError):
    """Raised only between transaction steps at a safe checkpoint."""


_PROCESS_LOCK_GUARD = threading.Lock()
_PROCESS_LOCKS = set()


def utc_timestamp(value=None):
    value = value or datetime.now(timezone.utc)
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).isoformat()


def _safe_name(value):
    return (
        re.sub(r"[^A-Za-z0-9_.-]+", "-", str(value or "repair")).strip("-")
        or "repair"
    )


def normalize_retention(value):
    try:
        value = int(value)
    except (TypeError, ValueError):
        value = DEFAULT_REPAIR_RETENTION
    return max(MIN_REPAIR_RETENTION, min(MAX_REPAIR_RETENTION, value))


def _absolute_path(path):
    if not path:
        raise RepairTransactionError("Repair path is missing")
    return os.path.abspath(os.path.expandvars(os.path.expanduser(str(path))))


def _is_link_or_junction(path):
    if os.path.islink(path):
        return True
    isjunction = getattr(os.path, "isjunction", None)
    return bool(isjunction and isjunction(path))


def _validate_real_directory(path, *, create=False):
    path = _absolute_path(path)
    if create:
        os.makedirs(path, exist_ok=True)
    if not os.path.isdir(path):
        raise RepairTransactionError(f"Repair directory does not exist: {path}")
    if _is_link_or_junction(path):
        raise RepairTransactionError(
            f"Repair directory cannot be a link or junction: {path}"
        )
    if os.path.normcase(os.path.realpath(path)) != os.path.normcase(path):
        raise RepairTransactionError(
            f"Repair directory does not resolve to itself: {path}"
        )
    return path


def _path_within(root, path):
    root = _absolute_path(root)
    path = _absolute_path(path)
    try:
        return os.path.commonpath([root, path]) == root
    except ValueError:
        return False


def _atomic_write_json(path, payload):
    path = _absolute_path(path)
    folder = os.path.dirname(path)
    os.makedirs(folder, exist_ok=True)
    temp_path = os.path.join(
        folder,
        f".{os.path.basename(path)}.{uuid.uuid4().hex}.tmp",
    )
    try:
        with open(temp_path, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, indent=2, sort_keys=True)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


def _append_jsonl(path, payload):
    path = _absolute_path(path)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "a", encoding="utf-8", newline="\n") as handle:
        handle.write(json.dumps(payload, sort_keys=True))
        handle.write("\n")
        handle.flush()
        os.fsync(handle.fileno())


def _json_hash(payload):
    serialized = json.dumps(
        payload,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(serialized).hexdigest()


def _sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class RepairOperationLock:
    """Cross-process repair lock with a UUID owner record."""

    def __init__(self, backup_base, operation_id):
        self.backup_base = _validate_real_directory(backup_base, create=True)
        self.operation_id = str(operation_id)
        self.path = os.path.join(self.backup_base, LOCK_FILENAME)
        self.handle = None
        self._process_key = os.path.normcase(self.path)

    def __enter__(self):
        with _PROCESS_LOCK_GUARD:
            if self._process_key in _PROCESS_LOCKS:
                raise RepairLockError(
                    "Another repair or restore is already running"
                )
            _PROCESS_LOCKS.add(self._process_key)

        try:
            self.handle = open(self.path, "a+b")
            self.handle.seek(0)
            if os.name == "nt":
                import msvcrt

                if os.path.getsize(self.path) == 0:
                    self.handle.write(b"\0")
                    self.handle.flush()
                self.handle.seek(0)
                try:
                    msvcrt.locking(
                        self.handle.fileno(),
                        msvcrt.LK_NBLCK,
                        1,
                    )
                except OSError as exc:
                    raise RepairLockError(
                        "Another repair or restore owns the operation lock"
                    ) from exc
            else:
                import fcntl

                try:
                    fcntl.flock(
                        self.handle.fileno(),
                        fcntl.LOCK_EX | fcntl.LOCK_NB,
                    )
                except OSError as exc:
                    raise RepairLockError(
                        "Another repair or restore owns the operation lock"
                    ) from exc

            record = {
                "OperationId": self.operation_id,
                "ProcessId": os.getpid(),
                "AcquiredAt": utc_timestamp(),
            }
            self.handle.seek(0)
            self.handle.truncate()
            self.handle.write(
                (json.dumps(record, sort_keys=True) + "\n").encode("utf-8")
            )
            self.handle.flush()
            os.fsync(self.handle.fileno())
            self.handle.seek(0)
            return self
        except Exception:
            self._release_process_key()
            if self.handle:
                self.handle.close()
                self.handle = None
            raise

    def _release_process_key(self):
        with _PROCESS_LOCK_GUARD:
            _PROCESS_LOCKS.discard(self._process_key)

    def __exit__(self, exc_type, exc, traceback):
        if self.handle:
            try:
                self.handle.seek(0)
                if os.name == "nt":
                    import msvcrt

                    msvcrt.locking(
                        self.handle.fileno(),
                        msvcrt.LK_UNLCK,
                        1,
                    )
                else:
                    import fcntl

                    fcntl.flock(self.handle.fileno(), fcntl.LOCK_UN)
            finally:
                self.handle.close()
                self.handle = None
        self._release_process_key()
        return False


def _package_paths(environ):
    local_app_data = environ.get("LOCALAPPDATA") or os.path.join(
        os.path.expanduser("~"),
        "AppData",
        "Local",
    )
    program_data = environ.get("ProgramData") or environ.get(
        "PROGRAMDATA",
        r"C:\ProgramData",
    )
    return {
        "token_cache": os.path.join(
            local_app_data,
            "Microsoft",
            "TokenBroker",
            "Cache",
        ),
        "store_cache": os.path.join(
            local_app_data,
            "Packages",
            "Microsoft.WindowsStore_8wekyb3d8bbwe",
            "LocalCache",
        ),
        "store_token_cache": os.path.join(
            local_app_data,
            "Packages",
            "Microsoft.WindowsStore_8wekyb3d8bbwe",
            "AC",
            "TokenBroker",
            "Cache",
        ),
        "aad_token_cache": os.path.join(
            local_app_data,
            "Packages",
            "Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy",
            "AC",
            "TokenBroker",
            "Cache",
        ),
        "store_inet_cache": os.path.join(
            local_app_data,
            "Packages",
            "Microsoft.WindowsStore_8wekyb3d8bbwe",
            "AC",
            "INetCache",
        ),
        "purchase_cache": os.path.join(
            local_app_data,
            "Packages",
            "Microsoft.StorePurchaseApp_8wekyb3d8bbwe",
            "LocalCache",
        ),
        "genuine_ticket": os.path.join(
            program_data,
            "Microsoft",
            "Windows",
            "ClipSVC",
            "GenuineTicket",
        ),
        "license_tokens": os.path.join(
            program_data,
            "Microsoft",
            "Windows",
            "ClipSVC",
            "Tokens",
        ),
    }


def _filesystem_target(target_id, path, description):
    return {
        "Id": target_id,
        "Type": "FileSystem",
        "Path": _absolute_path(path),
        "Description": description,
        "RequiredIfPresent": True,
    }


def _registry_target(target_id, path, description):
    return {
        "Id": target_id,
        "Type": "Registry",
        "Path": str(path),
        "Description": description,
        "RequiredIfPresent": True,
    }


def _repair_definitions(environ):
    paths = _package_paths(environ)
    store_packages = [
        "Microsoft.WindowsStore",
        "Microsoft.StorePurchaseApp",
    ]
    provisioning_packages = store_packages + [
        "Microsoft.DesktopAppInstaller",
    ]
    service_check = (
        "$names = @({names}); "
        "$missing = @($names | Where-Object {{ "
        "-not (Get-Service -Name $_ -ErrorAction SilentlyContinue) }}); "
        "if ($missing.Count) {{ throw ('Missing services: ' + "
        "($missing -join ', ')) }}"
    )
    package_check = (
        "$names = @({names}); "
        "$installed = @(Get-AppxPackage -AllUsers -ErrorAction Stop); "
        "$missing = @($names | Where-Object {{ "
        "$_ -notin $installed.Name }}); "
        "if ($missing.Count) {{ throw ('Missing registered packages: ' + "
        "($missing -join ', ')) }}"
    )
    start_services = (
        "$names = @({names}); foreach ($name in $names) {{ "
        "$service = Get-Service -Name $name -ErrorAction Stop; "
        "if ($service.StartType -eq 'Disabled') {{ "
        "throw \"Service $name is disabled by policy\" }}; "
        "if ($service.Status -ne 'Running') {{ "
        "Start-Service -Name $name -ErrorAction Stop }} }}"
    )
    service_postcondition = (
        "$names = @({names}); foreach ($name in $names) {{ "
        "$service = Get-Service -Name $name -ErrorAction Stop; "
        "if ($service.StartType -eq 'Disabled') {{ "
        "throw \"Service $name is disabled\" }} }}"
    )
    package_postcondition = package_check

    store_services = [
        "wuauserv",
        "bits",
        "ClipSVC",
        "LicenseManager",
    ]
    appx_services = ["AppXSVC", "ClipSVC"]
    licensing_services = ["LicenseManager", "ClipSVC"]

    store_environment = {
        "MSSTOREHELPER_REPAIR_PATHS": json.dumps([
            paths["token_cache"],
            paths["store_cache"],
            paths["store_token_cache"],
            paths["aad_token_cache"],
        ]),
    }
    licensing_environment = {
        "MSSTOREHELPER_REPAIR_PATHS": json.dumps([
            paths["genuine_ticket"],
            paths["license_tokens"],
        ]),
    }
    cache_environment = {
        "MSSTOREHELPER_REPAIR_PATHS": json.dumps([
            paths["store_cache"],
            paths["store_inet_cache"],
            paths["purchase_cache"],
        ]),
    }
    clear_paths = (
        "$paths = @($env:MSSTOREHELPER_REPAIR_PATHS | ConvertFrom-Json); "
        "foreach ($path in $paths) { "
        "if (Test-Path -LiteralPath $path) { "
        "Remove-Item -LiteralPath $path -Recurse -Force "
        "-ErrorAction Stop } }"
    )
    recreate_paths = (
        "$paths = @($env:MSSTOREHELPER_REPAIR_PATHS | ConvertFrom-Json); "
        "foreach ($path in $paths) { "
        "New-Item -ItemType Directory -Path $path -Force "
        "-ErrorAction Stop | Out-Null; "
        "if (-not (Test-Path -LiteralPath $path -PathType Container)) { "
        "throw \"Cache directory was not recreated: $path\" } }"
    )

    return {
        "store-repair": {
            "DisplayName": "Microsoft Store repair",
            "RequiresAdmin": True,
            "Reboot": "recommended",
            "Impact": [
                "Stops Store broker processes.",
                "Replaces Store and identity-token cache folders.",
                "Re-registers the existing Store packages.",
                "Runs wsreset.exe after verified backups are complete.",
            ],
            "Permissions": [
                "Administrator access",
                "Read/write access to the current user's Store caches",
                "Control of Windows Update and Store licensing services",
            ],
            "Environment": store_environment,
            "BackupTargets": [
                _filesystem_target(
                    "token-cache",
                    paths["token_cache"],
                    "Windows TokenBroker cache",
                ),
                _filesystem_target(
                    "store-local-cache",
                    paths["store_cache"],
                    "Microsoft Store LocalCache",
                ),
                _filesystem_target(
                    "store-token-cache",
                    paths["store_token_cache"],
                    "Microsoft Store TokenBroker cache",
                ),
                _filesystem_target(
                    "aad-token-cache",
                    paths["aad_token_cache"],
                    "AAD Broker TokenBroker cache",
                ),
            ],
            "Services": store_services,
            "Packages": store_packages,
            "AllUsersPackageInventory": True,
            "Preconditions": [
                {
                    "Id": "required-services",
                    "Description": "Verify required Store services exist",
                    "Command": service_check.format(
                        names=", ".join(
                            f"'{name}'" for name in store_services
                        )
                    ),
                },
                {
                    "Id": "required-packages",
                    "Description": "Verify Store packages are registered",
                    "Command": package_check.format(
                        names=", ".join(
                            f"'{name}'" for name in store_packages
                        )
                    ),
                },
                {
                    "Id": "wsreset",
                    "Description": "Verify wsreset.exe is available",
                    "Command": (
                        "if (-not (Get-Command wsreset.exe "
                        "-ErrorAction SilentlyContinue)) { "
                        "throw 'wsreset.exe is unavailable' }"
                    ),
                },
            ],
            "Steps": [
                {
                    "Id": "start-services",
                    "Description": "Start Store and licensing services",
                    "Command": start_services.format(
                        names=", ".join(
                            f"'{name}'" for name in store_services
                        )
                    ),
                    "Postcondition": service_postcondition.format(
                        names=", ".join(
                            f"'{name}'" for name in store_services
                        )
                    ),
                },
                {
                    "Id": "stop-brokers",
                    "Description": "Stop Store broker processes",
                    "Command": (
                        "$processes = @(Get-Process WinStore.App,"
                        "MicrosoftStore,RuntimeBroker "
                        "-ErrorAction SilentlyContinue); "
                        "if ($processes) { $processes | Stop-Process "
                        "-Force -ErrorAction Stop }"
                    ),
                },
                {
                    "Id": "clear-caches",
                    "Description": "Clear the verified Store cache targets",
                    "Command": clear_paths,
                    "Postcondition": (
                        "$paths = @($env:MSSTOREHELPER_REPAIR_PATHS | "
                        "ConvertFrom-Json); foreach ($path in $paths) { "
                        "if (Test-Path -LiteralPath $path) { "
                        "throw \"Cache target still exists: $path\" } }"
                    ),
                },
                {
                    "Id": "register-packages",
                    "Description": "Re-register existing Store packages",
                    "Command": (
                        "$names = @('Microsoft.WindowsStore', "
                        "'Microsoft.StorePurchaseApp'); "
                        "$packages = @(Get-AppxPackage -AllUsers "
                        "-ErrorAction Stop | Where-Object { "
                        "$_.Name -in $names }); "
                        "foreach ($package in $packages) { "
                        "$manifest = Join-Path $package.InstallLocation "
                        "'AppXManifest.xml'; "
                        "if (-not (Test-Path -LiteralPath $manifest)) { "
                        "throw \"Missing manifest: $manifest\" }; "
                        "Add-AppxPackage -DisableDevelopmentMode "
                        "-Register $manifest -ErrorAction Stop }"
                    ),
                    "Postcondition": package_postcondition.format(
                        names=", ".join(
                            f"'{name}'" for name in store_packages
                        )
                    ),
                },
                {
                    "Id": "wsreset",
                    "Description": "Reset the Store cache service",
                    "Command": (
                        "$process = Start-Process wsreset.exe "
                        "-WindowStyle Hidden -Wait -PassThru "
                        "-ErrorAction Stop; "
                        "if ($process.ExitCode -ne 0) { "
                        "throw \"wsreset.exe exited $($process.ExitCode)\" }"
                    ),
                },
            ],
            "FinalPostconditions": [
                {
                    "Id": "packages-remain-registered",
                    "Description": "Verify Store identities remain registered",
                    "Command": package_postcondition.format(
                        names=", ".join(
                            f"'{name}'" for name in store_packages
                        )
                    ),
                },
            ],
        },
        "provisioning-repair": {
            "DisplayName": "Store provisioning repair",
            "RequiresAdmin": True,
            "Reboot": "recommended",
            "Impact": [
                "Removes Store-related deprovision tombstones.",
                "Re-registers existing Store and App Installer packages.",
                "Does not download or provision new package binaries.",
            ],
            "Permissions": [
                "Administrator access",
                "Read/write access to HKLM AppX provisioning state",
                "All-users AppX inventory access",
            ],
            "Environment": {},
            "BackupTargets": [
                _registry_target(
                    "deprovisioned-root",
                    (
                        r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion"
                        r"\Appx\AppxAllUserStore\Deprovisioned"
                    ),
                    "AppX deprovisioned-package registry state",
                ),
            ],
            "Services": appx_services,
            "Packages": provisioning_packages,
            "AllUsersPackageInventory": True,
            "Preconditions": [
                {
                    "Id": "required-services",
                    "Description": "Verify AppX services exist",
                    "Command": service_check.format(
                        names=", ".join(
                            f"'{name}'" for name in appx_services
                        )
                    ),
                },
                {
                    "Id": "required-packages",
                    "Description": "Verify Store packages are registered",
                    "Command": package_check.format(
                        names=", ".join(
                            f"'{name}'" for name in provisioning_packages
                        )
                    ),
                },
            ],
            "Steps": [
                {
                    "Id": "start-services",
                    "Description": "Start AppX services",
                    "Command": start_services.format(
                        names=", ".join(
                            f"'{name}'" for name in appx_services
                        )
                    ),
                    "Postcondition": service_postcondition.format(
                        names=", ".join(
                            f"'{name}'" for name in appx_services
                        )
                    ),
                },
                {
                    "Id": "clear-tombstones",
                    "Description": "Clear Store deprovision tombstones",
                    "Command": (
                        "$root = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\"
                        "CurrentVersion\\Appx\\AppxAllUserStore\\"
                        "Deprovisioned'; "
                        "$patterns = @('*Microsoft.WindowsStore*', "
                        "'*Microsoft.StorePurchaseApp*', "
                        "'*Microsoft.DesktopAppInstaller*'); "
                        "if (Test-Path -LiteralPath $root) { "
                        "$targets = @(Get-ChildItem -LiteralPath $root "
                        "-ErrorAction Stop | Where-Object { "
                        "$name = $_.PSChildName; "
                        "@($patterns | Where-Object { "
                        "$name -like $_ }).Count -gt 0 }); "
                        "foreach ($target in $targets) { "
                        "Remove-Item -LiteralPath $target.PSPath "
                        "-Recurse -Force -ErrorAction Stop } }"
                    ),
                    "Postcondition": (
                        "$root = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\"
                        "CurrentVersion\\Appx\\AppxAllUserStore\\"
                        "Deprovisioned'; "
                        "$patterns = @('*Microsoft.WindowsStore*', "
                        "'*Microsoft.StorePurchaseApp*', "
                        "'*Microsoft.DesktopAppInstaller*'); "
                        "if (Test-Path -LiteralPath $root) { "
                        "$remaining = @(Get-ChildItem -LiteralPath $root "
                        "-ErrorAction Stop | Where-Object { "
                        "$name = $_.PSChildName; "
                        "@($patterns | Where-Object { "
                        "$name -like $_ }).Count -gt 0 }); "
                        "if ($remaining.Count) { "
                        "throw 'Store deprovision tombstones remain' } }"
                    ),
                },
                {
                    "Id": "register-packages",
                    "Description": "Re-register Store packages for users",
                    "Command": (
                        "$names = @('Microsoft.WindowsStore', "
                        "'Microsoft.StorePurchaseApp', "
                        "'Microsoft.DesktopAppInstaller'); "
                        "$packages = @(Get-AppxPackage -AllUsers "
                        "-ErrorAction Stop | Where-Object { "
                        "$_.Name -in $names }); "
                        "foreach ($package in $packages) { "
                        "$manifest = Join-Path $package.InstallLocation "
                        "'AppXManifest.xml'; "
                        "if (-not (Test-Path -LiteralPath $manifest)) { "
                        "throw \"Missing manifest: $manifest\" }; "
                        "Add-AppxPackage -DisableDevelopmentMode "
                        "-Register $manifest -ErrorAction Stop }"
                    ),
                    "Postcondition": package_postcondition.format(
                        names=", ".join(
                            f"'{name}'" for name in provisioning_packages
                        )
                    ),
                },
            ],
            "FinalPostconditions": [
                {
                    "Id": "packages-remain-registered",
                    "Description": "Verify provisioning identities",
                    "Command": package_postcondition.format(
                        names=", ".join(
                            f"'{name}'" for name in provisioning_packages
                        )
                    ),
                },
            ],
        },
        "licensing-reset": {
            "DisplayName": "Store licensing reset",
            "RequiresAdmin": True,
            "Reboot": "not-required",
            "Impact": [
                "Temporarily stops Store licensing services.",
                "Replaces ClipSVC ticket and token caches.",
                "Re-registers existing Store licensing packages.",
            ],
            "Permissions": [
                "Administrator access",
                "Read/write access to ProgramData ClipSVC state",
                "Control of ClipSVC and LicenseManager",
            ],
            "Environment": licensing_environment,
            "BackupTargets": [
                _filesystem_target(
                    "genuine-ticket",
                    paths["genuine_ticket"],
                    "ClipSVC GenuineTicket cache",
                ),
                _filesystem_target(
                    "license-tokens",
                    paths["license_tokens"],
                    "ClipSVC token cache",
                ),
            ],
            "Services": licensing_services,
            "Packages": store_packages,
            "AllUsersPackageInventory": True,
            "Preconditions": [
                {
                    "Id": "required-services",
                    "Description": "Verify licensing services exist",
                    "Command": service_check.format(
                        names=", ".join(
                            f"'{name}'" for name in licensing_services
                        )
                    ),
                },
                {
                    "Id": "required-packages",
                    "Description": "Verify licensing packages are registered",
                    "Command": package_check.format(
                        names=", ".join(
                            f"'{name}'" for name in store_packages
                        )
                    ),
                },
            ],
            "Steps": [
                {
                    "Id": "stop-services",
                    "Description": "Stop Store licensing services",
                    "Command": (
                        "$names = @('LicenseManager', 'ClipSVC'); "
                        "foreach ($name in $names) { "
                        "$service = Get-Service -Name $name "
                        "-ErrorAction Stop; "
                        "if ($service.Status -ne 'Stopped') { "
                        "Stop-Service -Name $name -Force "
                        "-ErrorAction Stop } }"
                    ),
                    "Postcondition": (
                        "$names = @('LicenseManager', 'ClipSVC'); "
                        "foreach ($name in $names) { "
                        "$service = Get-Service -Name $name "
                        "-ErrorAction Stop; "
                        "if ($service.Status -ne 'Stopped') { "
                        "throw \"Service did not stop: $name\" } }"
                    ),
                },
                {
                    "Id": "clear-license-cache",
                    "Description": "Clear verified ClipSVC cache targets",
                    "Command": clear_paths,
                    "Postcondition": (
                        "$paths = @($env:MSSTOREHELPER_REPAIR_PATHS | "
                        "ConvertFrom-Json); foreach ($path in $paths) { "
                        "if (Test-Path -LiteralPath $path) { "
                        "throw \"Licensing target still exists: $path\" } }"
                    ),
                },
                {
                    "Id": "start-services",
                    "Description": "Restart Store licensing services",
                    "Command": start_services.format(
                        names=", ".join(
                            f"'{name}'" for name in licensing_services
                        )
                    ),
                    "Postcondition": service_postcondition.format(
                        names=", ".join(
                            f"'{name}'" for name in licensing_services
                        )
                    ),
                },
                {
                    "Id": "register-packages",
                    "Description": "Re-register Store licensing packages",
                    "Command": (
                        "$names = @('Microsoft.StorePurchaseApp', "
                        "'Microsoft.WindowsStore'); "
                        "$packages = @(Get-AppxPackage -AllUsers "
                        "-ErrorAction Stop | Where-Object { "
                        "$_.Name -in $names }); "
                        "foreach ($package in $packages) { "
                        "$manifest = Join-Path $package.InstallLocation "
                        "'AppXManifest.xml'; "
                        "if (-not (Test-Path -LiteralPath $manifest)) { "
                        "throw \"Missing manifest: $manifest\" }; "
                        "Add-AppxPackage -DisableDevelopmentMode "
                        "-Register $manifest -ErrorAction Stop }"
                    ),
                    "Postcondition": package_postcondition.format(
                        names=", ".join(
                            f"'{name}'" for name in store_packages
                        )
                    ),
                },
            ],
            "FinalPostconditions": [],
        },
        "cache-rebuild": {
            "DisplayName": "Store cache rebuild",
            "RequiresAdmin": False,
            "Reboot": "not-required",
            "Impact": [
                "Stops Store broker processes for the current user.",
                "Replaces three Store cache folders after verified copies.",
                "Runs wsreset.exe against the rebuilt cache.",
            ],
            "Permissions": [
                "Current-user Store package data access",
                "Permission to stop the current user's Store processes",
            ],
            "Environment": cache_environment,
            "BackupTargets": [
                _filesystem_target(
                    "store-local-cache",
                    paths["store_cache"],
                    "Microsoft Store LocalCache",
                ),
                _filesystem_target(
                    "store-inet-cache",
                    paths["store_inet_cache"],
                    "Microsoft Store INetCache",
                ),
                _filesystem_target(
                    "purchase-cache",
                    paths["purchase_cache"],
                    "Store Purchase App LocalCache",
                ),
            ],
            "Services": [],
            "Packages": ["Microsoft.WindowsStore"],
            "AllUsersPackageInventory": False,
            "Preconditions": [
                {
                    "Id": "wsreset",
                    "Description": "Verify wsreset.exe is available",
                    "Command": (
                        "if (-not (Get-Command wsreset.exe "
                        "-ErrorAction SilentlyContinue)) { "
                        "throw 'wsreset.exe is unavailable' }"
                    ),
                },
            ],
            "Steps": [
                {
                    "Id": "stop-brokers",
                    "Description": "Stop Store cache-owner processes",
                    "Command": (
                        "$processes = @(Get-Process WinStore.App,"
                        "MicrosoftStore,RuntimeBroker "
                        "-ErrorAction SilentlyContinue); "
                        "if ($processes) { $processes | Stop-Process "
                        "-Force -ErrorAction Stop }"
                    ),
                },
                {
                    "Id": "clear-caches",
                    "Description": "Remove verified Store cache targets",
                    "Command": clear_paths,
                    "Postcondition": (
                        "$paths = @($env:MSSTOREHELPER_REPAIR_PATHS | "
                        "ConvertFrom-Json); foreach ($path in $paths) { "
                        "if (Test-Path -LiteralPath $path) { "
                        "throw \"Cache target still exists: $path\" } }"
                    ),
                },
                {
                    "Id": "recreate-caches",
                    "Description": "Create clean Store cache folders",
                    "Command": recreate_paths,
                    "Postcondition": (
                        "$paths = @($env:MSSTOREHELPER_REPAIR_PATHS | "
                        "ConvertFrom-Json); foreach ($path in $paths) { "
                        "if (-not (Test-Path -LiteralPath $path "
                        "-PathType Container)) { "
                        "throw \"Cache directory is missing: $path\" } }"
                    ),
                },
                {
                    "Id": "wsreset",
                    "Description": "Reset the Store cache service",
                    "Command": (
                        "$process = Start-Process wsreset.exe "
                        "-WindowStyle Hidden -Wait -PassThru "
                        "-ErrorAction Stop; "
                        "if ($process.ExitCode -ne 0) { "
                        "throw \"wsreset.exe exited $($process.ExitCode)\" }"
                    ),
                },
            ],
            "FinalPostconditions": [
                {
                    "Id": "cache-folders",
                    "Description": "Verify rebuilt cache folders",
                    "Command": (
                        "$paths = @($env:MSSTOREHELPER_REPAIR_PATHS | "
                        "ConvertFrom-Json); foreach ($path in $paths) { "
                        "if (-not (Test-Path -LiteralPath $path "
                        "-PathType Container)) { "
                        "throw \"Cache directory is missing: $path\" } }"
                    ),
                },
            ],
        },
    }


def build_repair_plan(
    repair_type,
    *,
    backup_base,
    retention_count=DEFAULT_REPAIR_RETENTION,
    environ=None,
    operation_id=None,
    confirmation_token=None,
    created_at=None,
):
    environ = dict(environ or os.environ)
    definitions = _repair_definitions(environ)
    if repair_type not in definitions:
        raise RepairTransactionError(f"Unknown repair type: {repair_type}")
    definition = definitions[repair_type]
    operation_id = str(operation_id or uuid.uuid4())
    try:
        uuid.UUID(operation_id)
    except ValueError as exc:
        raise RepairTransactionError("Repair operation ID is invalid") from exc

    plan = {
        "SchemaVersion": REPAIR_SCHEMA_VERSION,
        "OperationId": operation_id,
        "ConfirmationToken": str(
            confirmation_token or secrets.token_urlsafe(24)
        ),
        "CreatedAt": utc_timestamp(created_at),
        "RepairType": repair_type,
        "DisplayName": definition["DisplayName"],
        "BackupBase": _absolute_path(backup_base),
        "RetentionCount": normalize_retention(retention_count),
        "RequiresAdmin": bool(definition["RequiresAdmin"]),
        "Reboot": definition["Reboot"],
        "Impact": list(definition["Impact"]),
        "Permissions": list(definition["Permissions"]),
        "Environment": dict(definition["Environment"]),
        "BackupTargets": [
            target.copy() for target in definition["BackupTargets"]
        ],
        "Services": list(definition["Services"]),
        "Packages": list(definition["Packages"]),
        "AllUsersPackageInventory": bool(
            definition["AllUsersPackageInventory"]
        ),
        "Preconditions": [
            item.copy() for item in definition["Preconditions"]
        ],
        "Steps": [item.copy() for item in definition["Steps"]],
        "FinalPostconditions": [
            item.copy() for item in definition["FinalPostconditions"]
        ],
    }
    return plan


def build_sandbox_repair_plan(
    sandbox_root,
    *,
    backup_base,
    retention_count=DEFAULT_REPAIR_RETENTION,
    operation_id=None,
    confirmation_token=None,
):
    sandbox_root = _validate_real_directory(sandbox_root)
    operation_id = str(operation_id or uuid.uuid4())
    target = os.path.join(sandbox_root, "state")
    return {
        "SchemaVersion": REPAIR_SCHEMA_VERSION,
        "OperationId": operation_id,
        "ConfirmationToken": str(
            confirmation_token or secrets.token_urlsafe(24)
        ),
        "CreatedAt": utc_timestamp(),
        "RepairType": "sandbox",
        "DisplayName": "Repair transaction sandbox",
        "BackupBase": _absolute_path(backup_base),
        "RetentionCount": normalize_retention(retention_count),
        "RequiresAdmin": False,
        "Reboot": "not-required",
        "Impact": [
            "Replaces only the declared sandbox state directory.",
        ],
        "Permissions": ["Read/write access to the sandbox directory"],
        "Environment": {
            "MSSTOREHELPER_SANDBOX_TARGET": target,
        },
        "BackupTargets": [
            _filesystem_target(
                "sandbox-state",
                target,
                "Sandbox state directory",
            ),
        ],
        "Services": [],
        "Packages": [],
        "AllUsersPackageInventory": False,
        "Preconditions": [
            {
                "Id": "sandbox-state",
                "Description": "Verify sandbox state exists",
                "Command": (
                    "$path = $env:MSSTOREHELPER_SANDBOX_TARGET; "
                    "if (-not (Test-Path -LiteralPath $path "
                    "-PathType Container)) { "
                    "throw 'Sandbox state is missing' }"
                ),
            },
        ],
        "Steps": [
            {
                "Id": "sandbox-mutation",
                "Description": "Replace sandbox state",
                "Command": (
                    "$path = $env:MSSTOREHELPER_SANDBOX_TARGET; "
                    "Remove-Item -LiteralPath $path -Recurse -Force "
                    "-ErrorAction Stop; "
                    "New-Item -ItemType Directory -Path $path "
                    "-ErrorAction Stop | Out-Null; "
                    "Set-Content -LiteralPath "
                    "(Join-Path $path 'state.txt') "
                    "-Value 'mutated' -NoNewline -ErrorAction Stop"
                ),
                "Postcondition": (
                    "$path = Join-Path "
                    "$env:MSSTOREHELPER_SANDBOX_TARGET 'state.txt'; "
                    "if ((Get-Content -LiteralPath $path -Raw "
                    "-ErrorAction Stop) -ne 'mutated') { "
                    "throw 'Sandbox mutation was not observed' }"
                ),
            },
        ],
        "FinalPostconditions": [],
        "SandboxRoot": sandbox_root,
    }


def _plan_without_secret(plan):
    return {
        key: value
        for key, value in plan.items()
        if key != "ConfirmationToken"
    }


def render_repair_plan(plan):
    validate_repair_plan(plan)
    lines = [
        f"Operation: {plan['DisplayName']}",
        f"Operation ID: {plan['OperationId']}",
        (
            "Administrator: required"
            if plan["RequiresAdmin"]
            else "Administrator: not required"
        ),
        f"Reboot: {plan['Reboot']}",
        f"Backup retention: {plan['RetentionCount']} transaction(s)",
        f"Backup base: {plan['BackupBase']}",
        "",
        "Impact:",
    ]
    lines.extend(f"- {item}" for item in plan["Impact"])
    lines.extend(["", "Permissions:"])
    lines.extend(f"- {item}" for item in plan["Permissions"])
    lines.extend(["", "Backups before mutation:"])
    if plan["BackupTargets"]:
        lines.extend(
            (
                f"- [{target['Type']}] {target['Description']}: "
                f"{target['Path']}"
            )
            for target in plan["BackupTargets"]
        )
    else:
        lines.append("- None")
    lines.extend(["", "Preconditions:"])
    for index, item in enumerate(plan["Preconditions"], 1):
        lines.append(f"{index}. {item['Description']}")
        lines.append(f"   {item['Command']}")
    lines.extend(["", "Mutation steps:"])
    for index, item in enumerate(plan["Steps"], 1):
        lines.append(f"{index}. {item['Description']}")
        lines.append(f"   {item['Command']}")
        if item.get("Postcondition"):
            lines.append(f"   Verify: {item['Postcondition']}")
    lines.extend(["", "Final verification:"])
    if plan["FinalPostconditions"]:
        for index, item in enumerate(plan["FinalPostconditions"], 1):
            lines.append(f"{index}. {item['Description']}")
            lines.append(f"   {item['Command']}")
    else:
        lines.append("- Baseline service/package state evidence only")
    return "\n".join(lines)


def validate_repair_plan(plan):
    if not isinstance(plan, dict):
        raise RepairTransactionError("Repair plan is missing")
    if plan.get("SchemaVersion") != REPAIR_SCHEMA_VERSION:
        raise RepairTransactionError("Repair plan schema is unsupported")
    try:
        uuid.UUID(str(plan.get("OperationId", "")))
    except ValueError as exc:
        raise RepairTransactionError("Repair operation ID is invalid") from exc
    if not plan.get("ConfirmationToken"):
        raise RepairTransactionError("Repair confirmation token is missing")
    if not plan.get("DisplayName") or not plan.get("RepairType"):
        raise RepairTransactionError("Repair plan identity is incomplete")
    if not isinstance(plan.get("Steps"), list) or not plan["Steps"]:
        raise RepairTransactionError("Repair plan has no mutation steps")
    backup_base = _absolute_path(plan.get("BackupBase"))
    for target in plan.get("BackupTargets", []):
        if target.get("Type") not in {"FileSystem", "Registry"}:
            raise RepairTransactionError("Repair backup target type is invalid")
        if not target.get("Id") or not target.get("Path"):
            raise RepairTransactionError("Repair backup target is incomplete")
    for collection in (
        "Preconditions",
        "Steps",
        "FinalPostconditions",
    ):
        for item in plan.get(collection, []):
            if not item.get("Id") or not item.get("Command"):
                raise RepairTransactionError(
                    f"Repair {collection} entry is incomplete"
                )
    plan["BackupBase"] = backup_base
    plan["RetentionCount"] = normalize_retention(
        plan.get("RetentionCount")
    )
    return plan


def _confirm_plan(plan, confirmation_token):
    supplied = str(confirmation_token or "")
    expected = str(plan.get("ConfirmationToken") or "")
    if not supplied or not hmac.compare_digest(supplied, expected):
        raise RepairTransactionError(
            "Explicit repair confirmation is required"
        )


def _default_is_admin():
    if os.name != "nt":
        return os.geteuid() == 0
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except (AttributeError, OSError):
        return False


def _run_process(args, *, env=None, timeout=90):
    return run_command(
        args,
        env=env,
        timeout=timeout,
    )


def _run_powershell(
    powershell_exe,
    command,
    *,
    environment=None,
    timeout=90,
):
    process_environment = os.environ.copy()
    process_environment.update({
        str(key): str(value)
        for key, value in (environment or {}).items()
    })
    guarded = "\n".join([
        "$ErrorActionPreference = 'Stop'",
        "Set-StrictMode -Version 3.0",
        command,
    ])
    return _run_process(
        [
            powershell_exe,
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            guarded,
        ],
        env=process_environment,
        timeout=timeout,
    )


def _result_record(
    *,
    phase,
    item_id,
    description,
    success,
    started_at,
    completed_at=None,
    return_code=None,
    stdout="",
    stderr="",
    evidence=None,
):
    return {
        "Phase": phase,
        "ItemId": item_id,
        "Description": description,
        "Success": bool(success),
        "StartedAt": started_at,
        "CompletedAt": completed_at or utc_timestamp(),
        "ReturnCode": return_code,
        "Stdout": str(stdout or "").strip(),
        "Stderr": str(stderr or "").strip(),
        "Evidence": evidence or {},
    }


def _run_command_evidence(
    powershell_exe,
    item,
    *,
    phase,
    environment,
    timeout,
):
    started = utc_timestamp()
    try:
        result = _run_powershell(
            powershell_exe,
            item["Command"],
            environment=environment,
            timeout=timeout,
        )
        return _result_record(
            phase=phase,
            item_id=item["Id"],
            description=item.get("Description", item["Id"]),
            success=result.returncode == 0,
            started_at=started,
            return_code=result.returncode,
            stdout=result.stdout,
            stderr=result.stderr,
            evidence={"Command": item["Command"]},
        )
    except (OSError, subprocess.SubprocessError) as exc:
        return _result_record(
            phase=phase,
            item_id=item["Id"],
            description=item.get("Description", item["Id"]),
            success=False,
            started_at=started,
            stderr=str(exc),
            evidence={"Command": item["Command"]},
        )


def _secure_backup_directory(path, powershell_exe, timeout=30):
    path = _validate_real_directory(path)
    if os.name != "nt":
        os.chmod(path, 0o700)
        mode = os.stat(path).st_mode & 0o777
        if mode != 0o700:
            raise RepairTransactionError(
                f"Backup permissions are not restrictive: {oct(mode)}"
            )
        return {
            "AclVerified": True,
            "Mode": oct(mode),
        }

    command = r'''
$path = $env:MSSTOREHELPER_REPAIR_BACKUP_ROOT
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sid = $identity.User.Value
$allowed = @(
    $sid,
    'S-1-3-4',
    'S-1-5-18',
    'S-1-5-32-544'
)
$grants = @(
    "*${sid}:(OI)(CI)F",
    '*S-1-5-18:(OI)(CI)F',
    '*S-1-5-32-544:(OI)(CI)F'
)
& icacls.exe $path /inheritance:r /grant:r $grants /Q | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "icacls.exe failed with exit code $LASTEXITCODE"
}
$directory = [System.IO.DirectoryInfo]::new($path)
$acl = $directory.GetAccessControl()
$entries = @(
    foreach ($entry in $acl.GetAccessRules(
        $true,
        $true,
        [System.Security.Principal.SecurityIdentifier]
    )) {
        $entrySid = $entry.IdentityReference.Value
        [pscustomobject]@{
            Sid = $entrySid
            Inherited = [bool]$entry.IsInherited
            Type = [string]$entry.AccessControlType
            Rights = [string]$entry.FileSystemRights
        }
    }
)
$unexpected = @(
    $entries | Where-Object {
        $_.Inherited -or
        $_.Type -ne 'Allow' -or
        $_.Sid -notin $allowed
    }
)
if ($unexpected.Count -gt 0) {
    $detail = $entries | ConvertTo-Json -Compress -Depth 4
    throw "Backup ACL contains inherited, denied, or unexpected principals: $detail"
}
if (-not ($entries | Where-Object { $_.Sid -eq $sid })) {
    throw 'Backup ACL does not grant the current user access'
}
[pscustomobject]@{
    AclVerified = $true
    OwnerSid = $acl.GetOwner(
        [System.Security.Principal.SecurityIdentifier]
    ).Value
    AllowedSids = @($entries.Sid | Select-Object -Unique)
} | ConvertTo-Json -Compress -Depth 4
'''
    result = _run_powershell(
        powershell_exe,
        command,
        environment={"MSSTOREHELPER_REPAIR_BACKUP_ROOT": path},
        timeout=timeout,
    )
    if result.returncode != 0:
        raise RepairTransactionError(
            result.stderr.strip()
            or result.stdout.strip()
            or "Backup ACL verification failed"
        )
    try:
        evidence = json.loads(result.stdout.strip())
    except json.JSONDecodeError as exc:
        raise RepairTransactionError(
            "Backup ACL verification returned invalid evidence"
        ) from exc
    if evidence.get("AclVerified") is not True:
        raise RepairTransactionError("Backup ACL was not verified")
    return evidence


def _scan_size(path):
    if not os.path.exists(path):
        return 0
    if _is_link_or_junction(path):
        raise RepairTransactionError(
            f"Backup target cannot be a link or junction: {path}"
        )
    if os.path.isfile(path):
        return os.path.getsize(path)
    total = 0
    for root, directories, files in os.walk(path, followlinks=False):
        for name in directories:
            candidate = os.path.join(root, name)
            if _is_link_or_junction(candidate):
                raise RepairTransactionError(
                    f"Backup target contains a link or junction: {candidate}"
                )
        for name in files:
            candidate = os.path.join(root, name)
            if _is_link_or_junction(candidate):
                raise RepairTransactionError(
                    f"Backup target contains a link: {candidate}"
                )
            total += os.path.getsize(candidate)
    return total


def _filesystem_inventory(path):
    path = _absolute_path(path)
    if not os.path.exists(path):
        return {
            "Present": False,
            "Kind": "missing",
            "Entries": [],
            "Digest": _json_hash([]),
            "SizeBytes": 0,
        }
    if _is_link_or_junction(path):
        raise RepairTransactionError(
            f"Repair state cannot be a link or junction: {path}"
        )

    if os.path.isfile(path):
        entries = [{
            "Path": ".",
            "Kind": "file",
            "SizeBytes": os.path.getsize(path),
            "Sha256": _sha256_file(path),
        }]
        kind = "file"
    else:
        entries = [{"Path": ".", "Kind": "directory"}]
        for root, directories, files in os.walk(path, followlinks=False):
            relative_root = os.path.relpath(root, path)
            for name in sorted(directories, key=str.casefold):
                candidate = os.path.join(root, name)
                if _is_link_or_junction(candidate):
                    raise RepairTransactionError(
                        "Repair state contains a link or junction: "
                        f"{candidate}"
                    )
                relative = os.path.normpath(
                    os.path.join(relative_root, name)
                ).replace("\\", "/")
                entries.append({
                    "Path": relative,
                    "Kind": "directory",
                })
            for name in sorted(files, key=str.casefold):
                candidate = os.path.join(root, name)
                if _is_link_or_junction(candidate):
                    raise RepairTransactionError(
                        f"Repair state contains a link: {candidate}"
                    )
                relative = os.path.normpath(
                    os.path.join(relative_root, name)
                ).replace("\\", "/")
                entries.append({
                    "Path": relative,
                    "Kind": "file",
                    "SizeBytes": os.path.getsize(candidate),
                    "Sha256": _sha256_file(candidate),
                })
        kind = "directory"
    entries.sort(key=lambda item: (item["Path"].casefold(), item["Kind"]))
    size_bytes = sum(
        item.get("SizeBytes", 0)
        for item in entries
        if item["Kind"] == "file"
    )
    return {
        "Present": True,
        "Kind": kind,
        "Entries": entries,
        "Digest": _json_hash(entries),
        "SizeBytes": size_bytes,
    }


def _copy_filesystem_backup(source, destination):
    source = _absolute_path(source)
    destination = _absolute_path(destination)
    if os.path.exists(destination):
        raise RepairTransactionError(
            f"Backup destination already exists: {destination}"
        )
    source_inventory = _filesystem_inventory(source)
    if not source_inventory["Present"]:
        return source_inventory, None
    os.makedirs(os.path.dirname(destination), exist_ok=True)
    if source_inventory["Kind"] == "directory":
        shutil.copytree(source, destination, copy_function=shutil.copy2)
    else:
        shutil.copy2(source, destination)
    backup_inventory = _filesystem_inventory(destination)
    if (
        source_inventory["Digest"] != backup_inventory["Digest"]
        or source_inventory["Kind"] != backup_inventory["Kind"]
    ):
        raise RepairTransactionError(
            f"Backup verification failed for {source}"
        )
    return source_inventory, backup_inventory


def _normalize_registry_path(path):
    path = str(path or "").strip().replace("/", "\\")
    replacements = {
        "HKEY_LOCAL_MACHINE\\": "HKLM\\",
        "HKLM:\\": "HKLM\\",
    }
    for prefix, replacement in replacements.items():
        if path.upper().startswith(prefix.upper()):
            path = replacement + path[len(prefix):]
            break
    if not re.fullmatch(r"HKLM\\[A-Za-z0-9 _.,{}()\\-]+", path):
        raise RepairTransactionError(
            f"Registry backup path is not allowed: {path}"
        )
    return path


def _backup_registry_target(target, destination, timeout):
    registry_path = _normalize_registry_path(target["Path"])
    query = _run_process(
        ["reg.exe", "query", registry_path, "/s"],
        timeout=timeout,
    )
    if query.returncode == 1:
        return {
            "Present": False,
            "RegistryPath": registry_path,
        }
    if query.returncode != 0:
        raise RepairTransactionError(
            query.stderr.strip()
            or query.stdout.strip()
            or f"Registry query failed for {registry_path}"
        )
    result = _run_process(
        ["reg.exe", "export", registry_path, destination, "/y"],
        timeout=timeout,
    )
    if result.returncode != 0:
        raise RepairTransactionError(
            result.stderr.strip()
            or result.stdout.strip()
            or f"Registry export failed for {registry_path}"
        )
    if not os.path.isfile(destination) or os.path.getsize(destination) <= 0:
        raise RepairTransactionError(
            f"Registry backup is empty for {registry_path}"
        )
    return {
        "Present": True,
        "RegistryPath": registry_path,
        "SizeBytes": os.path.getsize(destination),
        "Sha256": _sha256_file(destination),
        "QuerySha256": hashlib.sha256(
            query.stdout.replace("\r\n", "\n").strip().encode("utf-8")
        ).hexdigest(),
    }


def _estimate_backup_requirement(plan):
    estimated = 0
    evidence = []
    for target in plan.get("BackupTargets", []):
        if target["Type"] == "FileSystem":
            size = _scan_size(_absolute_path(target["Path"]))
        else:
            size = REGISTRY_BACKUP_RESERVE_BYTES
        estimated += size
        evidence.append({
            "Id": target["Id"],
            "EstimatedBytes": size,
        })
    required = max(
        MIN_BACKUP_FREE_BYTES,
        (estimated * 2) + (16 * 1024 * 1024),
    )
    return required, evidence


def _create_context(plan):
    backup_base = _validate_real_directory(
        plan["BackupBase"],
        create=True,
    )
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    name = (
        f"{_safe_name(plan['RepairType'])}-{timestamp}-"
        f"{plan['OperationId']}"
    )
    root = os.path.join(backup_base, name)
    os.mkdir(root, mode=0o700)
    root = _validate_real_directory(root)
    context = {
        "SchemaVersion": REPAIR_SCHEMA_VERSION,
        "OperationId": plan["OperationId"],
        "RepairType": plan["RepairType"],
        "RepairName": plan["DisplayName"],
        "StartedAt": utc_timestamp(),
        "CompletedAt": None,
        "Outcome": "preparing",
        "Phase": "preflight",
        "MutationStarted": False,
        "RestoreAvailable": False,
        "BackupRoot": root,
        "ManifestPath": os.path.join(root, MANIFEST_FILENAME),
        "BackupLogPath": os.path.join(root, BACKUP_RECORDS_FILENAME),
        "BaselineStatePath": os.path.join(root, BASELINE_STATE_FILENAME),
        "RestoreHistoryPath": os.path.join(root, RESTORE_HISTORY_FILENAME),
        "Plan": _plan_without_secret(plan),
        "PlanHash": _json_hash(_plan_without_secret(plan)),
        "Results": [],
    }
    _atomic_write_json(context["ManifestPath"], context)
    return context


def _write_context(context):
    _atomic_write_json(context["ManifestPath"], context)


def _add_result(context, result, log_callback=None):
    context["Results"].append(result)
    _write_context(context)
    if log_callback:
        prefix = "PASS" if result.get("Success") else "FAIL"
        log_callback(f"[{prefix}] {result.get('Description')}")
    return result


def _check_cancel(cancel_event, *, mutation_started):
    if cancel_event is not None and cancel_event.is_set():
        suffix = " after mutation" if mutation_started else ""
        raise RepairCancelled(
            f"Repair cancelled at a safe checkpoint{suffix}"
        )


def _capture_state(plan, powershell_exe, timeout):
    spec = {
        "Services": list(plan.get("Services", [])),
        "Packages": list(plan.get("Packages", [])),
        "AllUsers": bool(plan.get("AllUsersPackageInventory")),
    }
    if not spec["Services"] and not spec["Packages"]:
        return {
            "CapturedAt": utc_timestamp(),
            "Services": [],
            "Packages": [],
        }
    command = r'''
$spec = $env:MSSTOREHELPER_REPAIR_STATE_SPEC | ConvertFrom-Json
$services = @(
    foreach ($name in @($spec.Services)) {
        $service = Get-Service -Name $name -ErrorAction SilentlyContinue
        if ($service) {
            [pscustomobject]@{
                Name = $name
                Exists = $true
                Status = [string]$service.Status
                StartType = [string]$service.StartType
            }
        } else {
            [pscustomobject]@{
                Name = $name
                Exists = $false
                Status = ''
                StartType = ''
            }
        }
    }
)
$allPackages = @()
if (@($spec.Packages).Count -gt 0) {
    if ($spec.AllUsers) {
        $allPackages = @(Get-AppxPackage -AllUsers -ErrorAction Stop)
    } else {
        $allPackages = @(Get-AppxPackage -ErrorAction Stop)
    }
}
$packages = @(
    foreach ($name in @($spec.Packages)) {
        $matches = @($allPackages | Where-Object { $_.Name -eq $name })
        [pscustomobject]@{
            Name = $name
            Installed = [bool]($matches.Count -gt 0)
            Versions = @($matches.Version | ForEach-Object { [string]$_ } | Sort-Object -Unique)
            PackageFullNames = @($matches.PackageFullName | Sort-Object -Unique)
        }
    }
)
[pscustomobject]@{
    CapturedAt = (Get-Date).ToUniversalTime().ToString('o')
    Services = $services
    Packages = $packages
} | ConvertTo-Json -Compress -Depth 8
'''
    result = _run_powershell(
        powershell_exe,
        command,
        environment={
            "MSSTOREHELPER_REPAIR_STATE_SPEC": json.dumps(spec),
        },
        timeout=timeout,
    )
    if result.returncode != 0:
        raise RepairTransactionError(
            result.stderr.strip()
            or result.stdout.strip()
            or "Repair baseline state capture failed"
        )
    try:
        state = json.loads(result.stdout.strip())
    except json.JSONDecodeError as exc:
        raise RepairTransactionError(
            "Repair baseline state capture returned invalid evidence"
        ) from exc
    if not isinstance(state, dict):
        raise RepairTransactionError(
            "Repair baseline state capture returned invalid evidence"
        )
    state.setdefault("Services", [])
    state.setdefault("Packages", [])
    return state


def _backup_target(context, target, timeout):
    started = utc_timestamp()
    target_id = _safe_name(target["Id"])
    if target["Type"] == "FileSystem":
        original_path = _absolute_path(target["Path"])
        destination = os.path.join(
            context["BackupRoot"],
            "files",
            target_id,
        )
        source_inventory, backup_inventory = _copy_filesystem_backup(
            original_path,
            destination,
        )
        record = {
            "SchemaVersion": REPAIR_SCHEMA_VERSION,
            "OperationId": context["OperationId"],
            "TargetId": target["Id"],
            "Type": "FileSystem",
            "OriginalPath": original_path,
            "BackupPath": (
                destination if source_inventory["Present"] else ""
            ),
            "Present": source_inventory["Present"],
            "SourceInventory": source_inventory,
            "BackupInventory": backup_inventory or {},
            "RecordedAt": utc_timestamp(),
        }
    else:
        destination_folder = os.path.join(
            context["BackupRoot"],
            "registry",
        )
        os.makedirs(destination_folder, exist_ok=True)
        destination = os.path.join(
            destination_folder,
            f"{target_id}.reg",
        )
        registry_evidence = _backup_registry_target(
            target,
            destination,
            timeout,
        )
        record = {
            "SchemaVersion": REPAIR_SCHEMA_VERSION,
            "OperationId": context["OperationId"],
            "TargetId": target["Id"],
            "Type": "Registry",
            "OriginalPath": registry_evidence["RegistryPath"],
            "BackupPath": (
                destination if registry_evidence["Present"] else ""
            ),
            "Present": registry_evidence["Present"],
            "RegistryEvidence": registry_evidence,
            "RecordedAt": utc_timestamp(),
        }
    _append_jsonl(context["BackupLogPath"], record)
    return _result_record(
        phase="backup",
        item_id=target["Id"],
        description=f"Back up {target['Description']}",
        success=True,
        started_at=started,
        evidence=record,
    )


def _load_jsonl(path):
    records = []
    with open(path, "r", encoding="utf-8") as handle:
        for line_number, line in enumerate(handle, 1):
            if not line.strip():
                continue
            try:
                item = json.loads(line)
            except json.JSONDecodeError as exc:
                raise RepairTransactionError(
                    f"Invalid JSONL record at line {line_number}"
                ) from exc
            if not isinstance(item, dict):
                raise RepairTransactionError(
                    f"Invalid JSONL record at line {line_number}"
                )
            records.append(item)
    return records


def _prune_repair_backups(backup_base, retention_count, protected=None):
    backup_base = _validate_real_directory(backup_base, create=True)
    protected = {
        os.path.normcase(_absolute_path(path))
        for path in (protected or [])
    }
    candidates = []
    for name in os.listdir(backup_base):
        path = os.path.join(backup_base, name)
        if not os.path.isdir(path) or _is_link_or_junction(path):
            continue
        manifest_path = os.path.join(path, MANIFEST_FILENAME)
        if not os.path.isfile(manifest_path):
            continue
        try:
            with open(manifest_path, "r", encoding="utf-8") as handle:
                manifest = json.load(handle)
            timestamp = (
                manifest.get("CompletedAt")
                or manifest.get("StartedAt")
                or ""
            )
        except (OSError, json.JSONDecodeError):
            continue
        candidates.append((str(timestamp), path))
    candidates.sort(key=lambda item: item[0], reverse=True)
    removed = []
    kept_count = 0
    for _timestamp, path in candidates:
        normalized = os.path.normcase(_absolute_path(path))
        if normalized in protected:
            kept_count += 1
            continue
        if kept_count < retention_count:
            kept_count += 1
            continue
        if not _path_within(backup_base, path):
            continue
        shutil.rmtree(path)
        removed.append(path)
    return removed


def execute_repair_plan(
    plan,
    *,
    confirmation_token,
    powershell_exe,
    is_admin=None,
    cancel_event=None,
    log_callback=None,
    progress_callback=None,
    timeout=90,
    secure_backup=True,
):
    plan = validate_repair_plan(plan)
    _confirm_plan(plan, confirmation_token)
    if plan["RequiresAdmin"] and not (
        _default_is_admin() if is_admin is None else bool(is_admin)
    ):
        raise RepairTransactionError(
            "Administrator access is required for this repair"
        )

    backup_base = _validate_real_directory(
        plan["BackupBase"],
        create=True,
    )
    with RepairOperationLock(backup_base, plan["OperationId"]):
        context = _create_context(plan)
        total_units = (
            4
            + len(plan["Preconditions"])
            + len(plan["BackupTargets"])
            + len(plan["Steps"]) * 2
            + len(plan["FinalPostconditions"])
        )
        completed_units = 0

        def advance():
            nonlocal completed_units
            completed_units += 1
            if progress_callback:
                progress_callback(min(1.0, completed_units / total_units))

        try:
            _check_cancel(cancel_event, mutation_started=False)
            acl_started = utc_timestamp()
            if secure_backup:
                acl_evidence = _secure_backup_directory(
                    context["BackupRoot"],
                    powershell_exe,
                )
            else:
                acl_evidence = {"AclVerified": True, "TestOverride": True}
            _add_result(
                context,
                _result_record(
                    phase="preflight",
                    item_id="backup-acl",
                    description="Verify restrictive backup permissions",
                    success=True,
                    started_at=acl_started,
                    evidence=acl_evidence,
                ),
                log_callback,
            )
            advance()

            required_bytes, estimate_evidence = _estimate_backup_requirement(
                plan
            )
            free_bytes = shutil.disk_usage(context["BackupRoot"]).free
            disk_success = free_bytes >= required_bytes
            disk_result = _result_record(
                phase="preflight",
                item_id="backup-space",
                description="Verify backup free space",
                success=disk_success,
                started_at=utc_timestamp(),
                stderr=(
                    ""
                    if disk_success
                    else (
                        f"Backup requires {required_bytes} bytes; "
                        f"{free_bytes} bytes are available"
                    )
                ),
                evidence={
                    "RequiredBytes": required_bytes,
                    "FreeBytes": free_bytes,
                    "Targets": estimate_evidence,
                },
            )
            _add_result(context, disk_result, log_callback)
            advance()
            if not disk_success:
                raise RepairTransactionError(disk_result["Stderr"])

            for item in plan["Preconditions"]:
                _check_cancel(cancel_event, mutation_started=False)
                result = _run_command_evidence(
                    powershell_exe,
                    item,
                    phase="precondition",
                    environment=plan["Environment"],
                    timeout=timeout,
                )
                _add_result(context, result, log_callback)
                advance()
                if not result["Success"]:
                    raise RepairTransactionError(
                        result["Stderr"]
                        or result["Stdout"]
                        or f"Precondition failed: {item['Description']}"
                    )

            baseline_started = utc_timestamp()
            baseline = _capture_state(plan, powershell_exe, timeout)
            _atomic_write_json(context["BaselineStatePath"], baseline)
            baseline_hash = _sha256_file(context["BaselineStatePath"])
            context["BaselineStateSha256"] = baseline_hash
            _add_result(
                context,
                _result_record(
                    phase="preflight",
                    item_id="baseline-state",
                    description="Capture service and package baseline",
                    success=True,
                    started_at=baseline_started,
                    evidence={
                        "Path": context["BaselineStatePath"],
                        "Sha256": baseline_hash,
                        "ServiceCount": len(baseline["Services"]),
                        "PackageCount": len(baseline["Packages"]),
                    },
                ),
                log_callback,
            )
            advance()

            context["Phase"] = "backup"
            _write_context(context)
            for target in plan["BackupTargets"]:
                _check_cancel(cancel_event, mutation_started=False)
                try:
                    result = _backup_target(context, target, timeout)
                except Exception as exc:
                    result = _result_record(
                        phase="backup",
                        item_id=target["Id"],
                        description=f"Back up {target['Description']}",
                        success=False,
                        started_at=utc_timestamp(),
                        stderr=str(exc),
                    )
                _add_result(context, result, log_callback)
                advance()
                if not result["Success"]:
                    raise RepairTransactionError(
                        result["Stderr"]
                        or f"Backup failed: {target['Description']}"
                    )

            context["RestoreAvailable"] = True
            context["Phase"] = "mutation"
            _write_context(context)
            for item in plan["Steps"]:
                _check_cancel(
                    cancel_event,
                    mutation_started=context["MutationStarted"],
                )
                context["MutationStarted"] = True
                _write_context(context)
                result = _run_command_evidence(
                    powershell_exe,
                    item,
                    phase="mutation",
                    environment=plan["Environment"],
                    timeout=timeout,
                )
                _add_result(context, result, log_callback)
                advance()
                if not result["Success"]:
                    raise RepairTransactionError(
                        result["Stderr"]
                        or result["Stdout"]
                        or f"Mutation failed: {item['Description']}"
                    )
                if item.get("Postcondition"):
                    post_item = {
                        "Id": f"{item['Id']}-postcondition",
                        "Description": (
                            f"Verify {item['Description'].lower()}"
                        ),
                        "Command": item["Postcondition"],
                    }
                    post_result = _run_command_evidence(
                        powershell_exe,
                        post_item,
                        phase="postcondition",
                        environment=plan["Environment"],
                        timeout=timeout,
                    )
                    _add_result(context, post_result, log_callback)
                    if not post_result["Success"]:
                        raise RepairTransactionError(
                            post_result["Stderr"]
                            or post_result["Stdout"]
                            or (
                                "Postcondition failed: "
                                f"{item['Description']}"
                            )
                        )
                advance()

            context["Phase"] = "postcondition"
            _write_context(context)
            for item in plan["FinalPostconditions"]:
                result = _run_command_evidence(
                    powershell_exe,
                    item,
                    phase="postcondition",
                    environment=plan["Environment"],
                    timeout=timeout,
                )
                _add_result(context, result, log_callback)
                advance()
                if not result["Success"]:
                    raise RepairTransactionError(
                        result["Stderr"]
                        or result["Stdout"]
                        or (
                            "Final postcondition failed: "
                            f"{item['Description']}"
                        )
                    )

            final_state_started = utc_timestamp()
            final_state = _capture_state(plan, powershell_exe, timeout)
            final_state_path = os.path.join(
                context["BackupRoot"],
                "post-repair-state.json",
            )
            _atomic_write_json(final_state_path, final_state)
            _add_result(
                context,
                _result_record(
                    phase="postcondition",
                    item_id="final-state",
                    description="Capture post-repair service and package state",
                    success=True,
                    started_at=final_state_started,
                    evidence={
                        "Path": final_state_path,
                        "Sha256": _sha256_file(final_state_path),
                    },
                ),
                log_callback,
            )
            advance()
            context["Outcome"] = "succeeded"
        except RepairCancelled as exc:
            context["Outcome"] = (
                "cancelled-restore-available"
                if context["MutationStarted"]
                else "cancelled"
            )
            _add_result(
                context,
                _result_record(
                    phase=context["Phase"],
                    item_id="cancellation",
                    description="Repair cancellation",
                    success=False,
                    started_at=utc_timestamp(),
                    stderr=str(exc),
                    evidence={
                        "SafeCheckpoint": True,
                        "MutationStarted": context["MutationStarted"],
                    },
                ),
                log_callback,
            )
        except Exception as exc:
            if context["MutationStarted"]:
                context["Outcome"] = "failed-restore-available"
            elif context["RestoreAvailable"]:
                context["Outcome"] = "failed-before-mutation"
            else:
                context["Outcome"] = "preflight-failed"
            if not context["Results"] or context["Results"][-1].get(
                "Success"
            ):
                _add_result(
                    context,
                    _result_record(
                        phase=context["Phase"],
                        item_id="transaction",
                        description="Repair transaction",
                        success=False,
                        started_at=utc_timestamp(),
                        stderr=str(exc),
                    ),
                    log_callback,
                )
        finally:
            context["Phase"] = "complete"
            context["CompletedAt"] = utc_timestamp()
            _write_context(context)
            if progress_callback:
                progress_callback(1.0)

        _prune_repair_backups(
            backup_base,
            plan["RetentionCount"],
            protected=[context["BackupRoot"]],
        )
        return context


def _read_json_object(path, description):
    try:
        with open(path, "r", encoding="utf-8") as handle:
            value = json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        raise RepairTransactionError(
            f"{description} could not be read: {path}"
        ) from exc
    if not isinstance(value, dict):
        raise RepairTransactionError(f"{description} is not a JSON object")
    return value


def _load_repair_manifest(backup_root, backup_base):
    backup_base = _validate_real_directory(backup_base)
    backup_root = _validate_real_directory(backup_root)
    if not _path_within(backup_base, backup_root):
        raise RepairTransactionError(
            "Repair backup is outside the configured backup base"
        )
    manifest_path = os.path.join(backup_root, MANIFEST_FILENAME)
    manifest = _read_json_object(manifest_path, "Repair manifest")
    if manifest.get("SchemaVersion") != REPAIR_SCHEMA_VERSION:
        raise RepairTransactionError("Repair manifest schema is unsupported")
    try:
        uuid.UUID(str(manifest.get("OperationId", "")))
    except ValueError as exc:
        raise RepairTransactionError(
            "Repair manifest operation ID is invalid"
        ) from exc
    recorded_root = _absolute_path(manifest.get("BackupRoot"))
    if os.path.normcase(recorded_root) != os.path.normcase(backup_root):
        raise RepairTransactionError(
            "Repair manifest backup root does not match its location"
        )
    source_plan = manifest.get("Plan")
    if not isinstance(source_plan, dict):
        raise RepairTransactionError("Repair manifest plan is missing")
    if _json_hash(source_plan) != manifest.get("PlanHash"):
        raise RepairTransactionError("Repair manifest plan hash is invalid")
    if not manifest.get("RestoreAvailable"):
        raise RepairTransactionError(
            "Repair manifest does not contain a complete backup"
        )
    return manifest_path, manifest


def _allowed_restore_targets(manifest, *, allow_sandbox):
    source_plan = manifest["Plan"]
    repair_type = manifest.get("RepairType")
    if repair_type == "sandbox":
        if not allow_sandbox:
            raise RepairTransactionError(
                "Sandbox restore requires an explicit test override"
            )
        sandbox_root = _validate_real_directory(
            source_plan.get("SandboxRoot")
        )
        expected = [
            _filesystem_target(
                "sandbox-state",
                os.path.join(sandbox_root, "state"),
                "Sandbox state directory",
            )
        ]
    else:
        definitions = _repair_definitions(os.environ)
        if repair_type not in definitions:
            raise RepairTransactionError(
                f"Repair type is no longer supported: {repair_type}"
            )
        expected = definitions[repair_type]["BackupTargets"]

    expected_by_id = {target["Id"]: target for target in expected}
    recorded_targets = source_plan.get("BackupTargets")
    if not isinstance(recorded_targets, list):
        raise RepairTransactionError(
            "Repair manifest backup target list is invalid"
        )
    if len(recorded_targets) != len(expected_by_id):
        raise RepairTransactionError(
            "Repair manifest backup targets no longer match the repair"
        )
    for recorded in recorded_targets:
        expected_target = expected_by_id.get(recorded.get("Id"))
        if expected_target is None:
            raise RepairTransactionError(
                f"Unexpected repair target: {recorded.get('Id')}"
            )
        if recorded.get("Type") != expected_target["Type"]:
            raise RepairTransactionError(
                f"Repair target type changed: {recorded.get('Id')}"
            )
        if recorded["Type"] == "FileSystem":
            actual_path = os.path.normcase(
                _absolute_path(recorded.get("Path"))
            )
            expected_path = os.path.normcase(
                _absolute_path(expected_target["Path"])
            )
        else:
            actual_path = _normalize_registry_path(recorded.get("Path"))
            expected_path = _normalize_registry_path(
                expected_target["Path"]
            )
        if actual_path != expected_path:
            raise RepairTransactionError(
                f"Repair target path changed: {recorded.get('Id')}"
            )
    return expected_by_id


def _validate_backup_record(record, expected, manifest, backup_root):
    if record.get("SchemaVersion") != REPAIR_SCHEMA_VERSION:
        raise RepairTransactionError("Backup record schema is unsupported")
    if record.get("OperationId") != manifest["OperationId"]:
        raise RepairTransactionError("Backup record operation ID is invalid")
    if record.get("Type") != expected["Type"]:
        raise RepairTransactionError(
            f"Backup record type is invalid: {record.get('TargetId')}"
        )
    if expected["Type"] == "FileSystem":
        original = os.path.normcase(
            _absolute_path(record.get("OriginalPath"))
        )
        allowed = os.path.normcase(_absolute_path(expected["Path"]))
    else:
        original = _normalize_registry_path(record.get("OriginalPath"))
        allowed = _normalize_registry_path(expected["Path"])
    if original != allowed:
        raise RepairTransactionError(
            f"Backup record target is invalid: {record.get('TargetId')}"
        )

    present = bool(record.get("Present"))
    backup_path = str(record.get("BackupPath") or "")
    if not present:
        if backup_path:
            raise RepairTransactionError(
                f"Missing target has a backup path: {record['TargetId']}"
            )
        return
    if not backup_path:
        raise RepairTransactionError(
            f"Backup path is missing: {record['TargetId']}"
        )
    backup_path = _absolute_path(backup_path)
    if not _path_within(backup_root, backup_path):
        raise RepairTransactionError(
            f"Backup path escapes its transaction: {record['TargetId']}"
        )
    if expected["Type"] == "FileSystem":
        inventory = _filesystem_inventory(backup_path)
        expected_inventory = record.get("BackupInventory") or {}
        source_inventory = record.get("SourceInventory") or {}
        if (
            not inventory["Present"]
            or inventory["Digest"] != expected_inventory.get("Digest")
            or inventory["Digest"] != source_inventory.get("Digest")
            or inventory["Kind"] != source_inventory.get("Kind")
        ):
            raise RepairTransactionError(
                f"Filesystem backup hash is invalid: {record['TargetId']}"
            )
    else:
        evidence = record.get("RegistryEvidence") or {}
        if (
            not os.path.isfile(backup_path)
            or _is_link_or_junction(backup_path)
            or _sha256_file(backup_path) != evidence.get("Sha256")
        ):
            raise RepairTransactionError(
                f"Registry backup hash is invalid: {record['TargetId']}"
            )


def build_restore_plan(
    backup_root,
    *,
    backup_base,
    confirmation_token=None,
    operation_id=None,
    allow_sandbox=False,
):
    backup_base = _validate_real_directory(backup_base)
    manifest_path, manifest = _load_repair_manifest(
        backup_root,
        backup_base,
    )
    backup_root = _absolute_path(backup_root)
    expected_by_id = _allowed_restore_targets(
        manifest,
        allow_sandbox=allow_sandbox,
    )
    backup_log = _absolute_path(manifest.get("BackupLogPath"))
    if not _path_within(backup_root, backup_log):
        raise RepairTransactionError(
            "Repair backup record log escapes its transaction"
        )
    records = _load_jsonl(backup_log)
    records_by_id = {}
    for record in records:
        target_id = record.get("TargetId")
        if target_id in records_by_id:
            raise RepairTransactionError(
                f"Duplicate backup record: {target_id}"
            )
        expected = expected_by_id.get(target_id)
        if expected is None:
            raise RepairTransactionError(
                f"Unexpected backup record: {target_id}"
            )
        _validate_backup_record(
            record,
            expected,
            manifest,
            backup_root,
        )
        records_by_id[target_id] = record
    if set(records_by_id) != set(expected_by_id):
        raise RepairTransactionError(
            "Repair backup records are incomplete"
        )

    baseline_path = _absolute_path(manifest.get("BaselineStatePath"))
    if not _path_within(backup_root, baseline_path):
        raise RepairTransactionError(
            "Repair baseline state escapes its transaction"
        )
    baseline_hash = manifest.get("BaselineStateSha256")
    if (
        not os.path.isfile(baseline_path)
        or _sha256_file(baseline_path) != baseline_hash
    ):
        raise RepairTransactionError(
            "Repair baseline state hash is invalid"
        )
    baseline = _read_json_object(
        baseline_path,
        "Repair baseline state",
    )
    restore_history_path = _absolute_path(
        manifest["RestoreHistoryPath"]
    )
    if not _path_within(backup_root, restore_history_path):
        raise RepairTransactionError(
            "Repair restore history escapes its transaction"
        )

    operation_id = str(operation_id or uuid.uuid4())
    try:
        uuid.UUID(operation_id)
    except ValueError as exc:
        raise RepairTransactionError(
            "Restore operation ID is invalid"
        ) from exc
    plan = {
        "SchemaVersion": REPAIR_SCHEMA_VERSION,
        "OperationId": operation_id,
        "ConfirmationToken": str(
            confirmation_token or secrets.token_urlsafe(24)
        ),
        "CreatedAt": utc_timestamp(),
        "DisplayName": f"Restore {manifest['RepairName']}",
        "BackupBase": backup_base,
        "BackupRoot": backup_root,
        "SourceManifestPath": manifest_path,
        "SourceOperationId": manifest["OperationId"],
        "SourceOutcome": manifest.get("Outcome"),
        "RepairType": manifest["RepairType"],
        "RequiresAdmin": bool(manifest["Plan"].get("RequiresAdmin")),
        "Reboot": manifest["Plan"].get("Reboot", "unknown"),
        "RestoreTargets": [
            records_by_id[target_id]
            for target_id in expected_by_id
        ],
        "Services": list(manifest["Plan"].get("Services", [])),
        "Packages": list(manifest["Plan"].get("Packages", [])),
        "AllUsersPackageInventory": bool(
            manifest["Plan"].get("AllUsersPackageInventory")
        ),
        "BaselineStatePath": baseline_path,
        "BaselineStateSha256": baseline_hash,
        "BaselineState": baseline,
        "RestoreHistoryPath": restore_history_path,
        "AllowSandbox": bool(allow_sandbox),
    }
    plan["RestorePlanHash"] = _json_hash({
        key: value
        for key, value in plan.items()
        if key not in {"ConfirmationToken", "RestorePlanHash"}
    })
    return plan


def validate_restore_plan(plan):
    if not isinstance(plan, dict):
        raise RepairTransactionError("Restore plan is missing")
    if plan.get("SchemaVersion") != REPAIR_SCHEMA_VERSION:
        raise RepairTransactionError("Restore plan schema is unsupported")
    try:
        uuid.UUID(str(plan.get("OperationId", "")))
        uuid.UUID(str(plan.get("SourceOperationId", "")))
    except ValueError as exc:
        raise RepairTransactionError(
            "Restore plan operation ID is invalid"
        ) from exc
    expected_hash = _json_hash({
        key: value
        for key, value in plan.items()
        if key not in {"ConfirmationToken", "RestorePlanHash"}
    })
    if not hmac.compare_digest(
        str(plan.get("RestorePlanHash") or ""),
        expected_hash,
    ):
        raise RepairTransactionError("Restore plan changed after inspection")
    if not plan.get("ConfirmationToken"):
        raise RepairTransactionError(
            "Restore confirmation token is missing"
        )
    if not isinstance(plan.get("RestoreTargets"), list):
        raise RepairTransactionError("Restore targets are missing")
    backup_base = _validate_real_directory(plan.get("BackupBase"))
    backup_root = _validate_real_directory(plan.get("BackupRoot"))
    if not _path_within(backup_base, backup_root):
        raise RepairTransactionError(
            "Restore source is outside the configured backup base"
        )
    return plan


def render_restore_plan(plan):
    validate_restore_plan(plan)
    lines = [
        f"Operation: {plan['DisplayName']}",
        f"Operation ID: {plan['OperationId']}",
        f"Source operation: {plan['SourceOperationId']}",
        f"Source outcome: {plan['SourceOutcome']}",
        (
            "Administrator: required"
            if plan["RequiresAdmin"]
            else "Administrator: not required"
        ),
        f"Reboot: {plan['Reboot']}",
        f"Backup source: {plan['BackupRoot']}",
        "",
        "Restore actions (backups are retained):",
    ]
    for record in plan["RestoreTargets"]:
        action = "restore captured state" if record["Present"] else (
            "restore original absence"
        )
        lines.append(
            f"- [{record['Type']}] {record['TargetId']}: "
            f"{action} at {record['OriginalPath']}"
        )
    lines.extend([
        "",
        "Verification:",
        "- Verify every filesystem or registry target after restore.",
        "- Restore and verify captured service state.",
        "- Verify exact captured package identities and versions.",
        "- Keep the source backup available for repeated restore.",
    ])
    return "\n".join(lines)


def list_repair_backups(backup_base):
    try:
        backup_base = _validate_real_directory(backup_base)
    except RepairTransactionError:
        return []
    summaries = []
    for name in os.listdir(backup_base):
        root = os.path.join(backup_base, name)
        if not os.path.isdir(root) or _is_link_or_junction(root):
            continue
        manifest_path = os.path.join(root, MANIFEST_FILENAME)
        if not os.path.isfile(manifest_path):
            continue
        try:
            manifest = _read_json_object(
                manifest_path,
                "Repair manifest",
            )
            if (
                manifest.get("SchemaVersion") != REPAIR_SCHEMA_VERSION
                or not manifest.get("RestoreAvailable")
            ):
                continue
            summaries.append({
                "BackupRoot": root,
                "RepairName": manifest.get("RepairName", "Repair"),
                "RepairType": manifest.get("RepairType", ""),
                "Outcome": manifest.get("Outcome", "unknown"),
                "CompletedAt": manifest.get("CompletedAt"),
                "OperationId": manifest.get("OperationId"),
            })
        except (OSError, RepairTransactionError):
            continue
    summaries.sort(
        key=lambda item: str(item.get("CompletedAt") or ""),
        reverse=True,
    )
    return summaries


def _remove_restore_path(path):
    if not os.path.lexists(path):
        return
    if _is_link_or_junction(path):
        raise RepairTransactionError(
            f"Restore path cannot be a link or junction: {path}"
        )
    if os.path.isdir(path):
        shutil.rmtree(path)
    else:
        os.remove(path)


def _restore_filesystem_record(record, restore_operation_id):
    started = utc_timestamp()
    target = _absolute_path(record["OriginalPath"])
    parent = _validate_real_directory(
        os.path.dirname(target),
        create=True,
    )
    suffix = f".msstorehelper-{restore_operation_id}"
    staging = os.path.join(parent, suffix + "-staging")
    rollback = os.path.join(parent, suffix + "-rollback")
    for transient in (staging, rollback):
        if os.path.lexists(transient):
            raise RepairTransactionError(
                f"Restore transient path already exists: {transient}"
            )

    before = _filesystem_inventory(target)
    target_moved = False
    try:
        if record["Present"]:
            backup_path = _absolute_path(record["BackupPath"])
            backup_inventory = _filesystem_inventory(backup_path)
            expected = record["SourceInventory"]
            if (
                backup_inventory["Digest"] != expected.get("Digest")
                or backup_inventory["Kind"] != expected.get("Kind")
            ):
                raise RepairTransactionError(
                    f"Backup changed before restore: {record['TargetId']}"
                )
            if backup_inventory["Kind"] == "directory":
                shutil.copytree(
                    backup_path,
                    staging,
                    copy_function=shutil.copy2,
                )
            else:
                shutil.copy2(backup_path, staging)
            staged_inventory = _filesystem_inventory(staging)
            if staged_inventory["Digest"] != expected["Digest"]:
                raise RepairTransactionError(
                    f"Restore staging verification failed: "
                    f"{record['TargetId']}"
                )
            if os.path.lexists(target):
                if _is_link_or_junction(target):
                    raise RepairTransactionError(
                        f"Restore target became a link: {target}"
                    )
                os.replace(target, rollback)
                target_moved = True
            os.replace(staging, target)
            after = _filesystem_inventory(target)
            if (
                after["Digest"] != expected["Digest"]
                or after["Kind"] != expected["Kind"]
            ):
                raise RepairTransactionError(
                    f"Restored target verification failed: "
                    f"{record['TargetId']}"
                )
        else:
            if os.path.lexists(target):
                if _is_link_or_junction(target):
                    raise RepairTransactionError(
                        f"Restore target became a link: {target}"
                    )
                os.replace(target, rollback)
                target_moved = True
            after = _filesystem_inventory(target)
            if after["Present"]:
                raise RepairTransactionError(
                    f"Original absence was not restored: "
                    f"{record['TargetId']}"
                )
        if target_moved:
            _remove_restore_path(rollback)
    except Exception:
        if os.path.lexists(staging):
            _remove_restore_path(staging)
        if target_moved and os.path.lexists(rollback):
            if os.path.lexists(target):
                _remove_restore_path(target)
            os.replace(rollback, target)
        raise
    return _result_record(
        phase="restore",
        item_id=record["TargetId"],
        description=f"Restore filesystem target {record['TargetId']}",
        success=True,
        started_at=started,
        evidence={
            "Before": before,
            "After": after,
            "BackupRetained": os.path.exists(
                str(record.get("BackupPath") or "")
            ) if record["Present"] else True,
        },
    )


def _registry_query_evidence(path, timeout):
    path = _normalize_registry_path(path)
    query = _run_process(["reg.exe", "query", path, "/s"], timeout=timeout)
    if query.returncode == 1:
        return {"Present": False, "QuerySha256": ""}
    if query.returncode != 0:
        raise RepairTransactionError(
            query.stderr.strip()
            or query.stdout.strip()
            or f"Registry query failed for {path}"
        )
    normalized = query.stdout.replace("\r\n", "\n").strip()
    return {
        "Present": True,
        "QuerySha256": hashlib.sha256(
            normalized.encode("utf-8")
        ).hexdigest(),
    }


def _restore_registry_record(record, timeout):
    started = utc_timestamp()
    target = _normalize_registry_path(record["OriginalPath"])
    before = _registry_query_evidence(target, timeout)
    if record["Present"]:
        backup_path = _absolute_path(record["BackupPath"])
        evidence = record["RegistryEvidence"]
        if _sha256_file(backup_path) != evidence.get("Sha256"):
            raise RepairTransactionError(
                f"Registry backup changed: {record['TargetId']}"
            )
        if before["Present"]:
            deleted = _run_process(
                ["reg.exe", "delete", target, "/f"],
                timeout=timeout,
            )
            if deleted.returncode != 0:
                raise RepairTransactionError(
                    deleted.stderr.strip()
                    or deleted.stdout.strip()
                    or f"Registry delete failed for {target}"
                )
        imported = _run_process(
            ["reg.exe", "import", backup_path],
            timeout=timeout,
        )
        if imported.returncode != 0:
            raise RepairTransactionError(
                imported.stderr.strip()
                or imported.stdout.strip()
                or f"Registry import failed for {target}"
            )
        after = _registry_query_evidence(target, timeout)
        if (
            not after["Present"]
            or after["QuerySha256"] != evidence.get("QuerySha256")
        ):
            raise RepairTransactionError(
                f"Restored registry verification failed: "
                f"{record['TargetId']}"
            )
    else:
        if before["Present"]:
            deleted = _run_process(
                ["reg.exe", "delete", target, "/f"],
                timeout=timeout,
            )
            if deleted.returncode != 0:
                raise RepairTransactionError(
                    deleted.stderr.strip()
                    or deleted.stdout.strip()
                    or f"Registry delete failed for {target}"
                )
        after = _registry_query_evidence(target, timeout)
        if after["Present"]:
            raise RepairTransactionError(
                f"Original registry absence was not restored: "
                f"{record['TargetId']}"
            )
    return _result_record(
        phase="restore",
        item_id=record["TargetId"],
        description=f"Restore registry target {record['TargetId']}",
        success=True,
        started_at=started,
        evidence={
            "Before": before,
            "After": after,
            "BackupRetained": (
                os.path.isfile(record["BackupPath"])
                if record["Present"]
                else True
            ),
        },
    )


def _restore_service_state(baseline, powershell_exe, timeout):
    services = baseline.get("Services") or []
    if not services:
        return
    command = r'''
$services = @($env:MSSTOREHELPER_RESTORE_SERVICES | ConvertFrom-Json)
foreach ($expected in $services) {
    $service = Get-Service -Name $expected.Name -ErrorAction SilentlyContinue
    if (-not $expected.Exists) {
        if ($service) { throw "Unexpected service exists: $($expected.Name)" }
        continue
    }
    if (-not $service) { throw "Required service is missing: $($expected.Name)" }
    $startup = [string]$expected.StartType
    if ($startup -in @('Automatic', 'Manual', 'Disabled')) {
        Set-Service -Name $expected.Name -StartupType $startup -ErrorAction Stop
    }
    $status = [string]$expected.Status
    if ($status -eq 'Running' -and $service.Status -ne 'Running') {
        Start-Service -Name $expected.Name -ErrorAction Stop
        (Get-Service -Name $expected.Name).WaitForStatus(
            [System.ServiceProcess.ServiceControllerStatus]::Running,
            [TimeSpan]::FromSeconds(30)
        )
    } elseif ($status -eq 'Stopped' -and $service.Status -ne 'Stopped') {
        Stop-Service -Name $expected.Name -Force -ErrorAction Stop
        (Get-Service -Name $expected.Name).WaitForStatus(
            [System.ServiceProcess.ServiceControllerStatus]::Stopped,
            [TimeSpan]::FromSeconds(30)
        )
    } elseif ($status -notin @('Running', 'Stopped') -and
              [string]$service.Status -ne $status) {
        throw "Unsupported service state requires review: $($expected.Name)"
    }
}
'''
    result = _run_powershell(
        powershell_exe,
        command,
        environment={
            "MSSTOREHELPER_RESTORE_SERVICES": json.dumps(services),
        },
        timeout=timeout,
    )
    if result.returncode != 0:
        raise RepairTransactionError(
            result.stderr.strip()
            or result.stdout.strip()
            or "Service state restore failed"
        )


def _as_object_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, dict):
        return [value]
    raise RepairTransactionError("Captured Windows state is invalid")


def _state_identity(state):
    services = {}
    for item in _as_object_list(state.get("Services")):
        services[str(item.get("Name"))] = {
            "Exists": bool(item.get("Exists")),
            "Status": str(item.get("Status") or ""),
            "StartType": str(item.get("StartType") or ""),
        }
    packages = {}
    for item in _as_object_list(state.get("Packages")):
        versions = item.get("Versions") or []
        if not isinstance(versions, list):
            versions = [versions]
        full_names = item.get("PackageFullNames") or []
        if not isinstance(full_names, list):
            full_names = [full_names]
        packages[str(item.get("Name"))] = {
            "Installed": bool(item.get("Installed")),
            "Versions": sorted(str(value) for value in versions),
            "PackageFullNames": sorted(
                str(value) for value in full_names
            ),
        }
    return {"Services": services, "Packages": packages}


def execute_restore_plan(
    plan,
    *,
    confirmation_token,
    powershell_exe,
    is_admin=None,
    cancel_event=None,
    log_callback=None,
    progress_callback=None,
    timeout=90,
    secure_backup=True,
):
    plan = validate_restore_plan(plan)
    _confirm_plan(plan, confirmation_token)
    if plan["RequiresAdmin"] and not (
        _default_is_admin() if is_admin is None else bool(is_admin)
    ):
        raise RepairTransactionError(
            "Administrator access is required for this restore"
        )

    refreshed = build_restore_plan(
        plan["BackupRoot"],
        backup_base=plan["BackupBase"],
        operation_id=plan["OperationId"],
        confirmation_token=plan["ConfirmationToken"],
        allow_sandbox=plan["AllowSandbox"],
    )
    refreshed_material = {
        "SourceOperationId": refreshed["SourceOperationId"],
        "RestoreTargets": refreshed["RestoreTargets"],
        "BaselineStateSha256": refreshed["BaselineStateSha256"],
        "Services": refreshed["Services"],
        "Packages": refreshed["Packages"],
    }
    inspected_material = {
        "SourceOperationId": plan["SourceOperationId"],
        "RestoreTargets": plan["RestoreTargets"],
        "BaselineStateSha256": plan["BaselineStateSha256"],
        "Services": plan["Services"],
        "Packages": plan["Packages"],
    }
    if _json_hash(refreshed_material) != _json_hash(inspected_material):
        raise RepairTransactionError(
            "Repair backup changed after the restore plan was inspected"
        )

    context = {
        "SchemaVersion": REPAIR_SCHEMA_VERSION,
        "OperationId": plan["OperationId"],
        "SourceOperationId": plan["SourceOperationId"],
        "StartedAt": utc_timestamp(),
        "CompletedAt": None,
        "Outcome": "preparing",
        "Phase": "preflight",
        "Results": [],
    }
    total_units = len(plan["RestoreTargets"]) + 3
    completed_units = 0

    def advance():
        nonlocal completed_units
        completed_units += 1
        if progress_callback:
            progress_callback(min(1.0, completed_units / total_units))

    def add_result(result):
        context["Results"].append(result)
        if log_callback:
            prefix = "PASS" if result.get("Success") else "FAIL"
            log_callback(f"[{prefix}] {result.get('Description')}")

    with RepairOperationLock(plan["BackupBase"], plan["OperationId"]):
        try:
            _check_cancel(cancel_event, mutation_started=False)
            acl_started = utc_timestamp()
            if secure_backup:
                acl_evidence = _secure_backup_directory(
                    plan["BackupRoot"],
                    powershell_exe,
                )
            else:
                acl_evidence = {
                    "AclVerified": True,
                    "TestOverride": True,
                }
            add_result(_result_record(
                phase="preflight",
                item_id="backup-acl",
                description="Verify restrictive backup permissions",
                success=True,
                started_at=acl_started,
                evidence=acl_evidence,
            ))
            advance()

            context["Phase"] = "restore"
            for record in plan["RestoreTargets"]:
                _check_cancel(cancel_event, mutation_started=True)
                if record["Type"] == "FileSystem":
                    result = _restore_filesystem_record(
                        record,
                        plan["OperationId"],
                    )
                else:
                    result = _restore_registry_record(record, timeout)
                add_result(result)
                advance()

            service_started = utc_timestamp()
            _restore_service_state(
                plan["BaselineState"],
                powershell_exe,
                timeout,
            )
            add_result(_result_record(
                phase="restore",
                item_id="service-state",
                description="Restore captured service state",
                success=True,
                started_at=service_started,
                evidence={
                    "ServiceCount": len(
                        _as_object_list(
                            plan["BaselineState"].get("Services")
                        )
                    ),
                },
            ))
            advance()

            context["Phase"] = "verification"
            state_started = utc_timestamp()
            state_plan = {
                "Services": plan["Services"],
                "Packages": plan["Packages"],
                "AllUsersPackageInventory": (
                    plan["AllUsersPackageInventory"]
                ),
            }
            current_state = _capture_state(
                state_plan,
                powershell_exe,
                timeout,
            )
            expected_state = _state_identity(plan["BaselineState"])
            actual_state = _state_identity(current_state)
            state_matches = expected_state == actual_state
            state_result = _result_record(
                phase="verification",
                item_id="windows-state",
                description="Verify captured services and packages",
                success=state_matches,
                started_at=state_started,
                stderr=(
                    ""
                    if state_matches
                    else "Service or package state differs from the backup"
                ),
                evidence={
                    "Expected": expected_state,
                    "Actual": actual_state,
                },
            )
            add_result(state_result)
            advance()
            if not state_matches:
                raise RepairTransactionError(state_result["Stderr"])
            context["Outcome"] = "succeeded"
        except RepairCancelled as exc:
            context["Outcome"] = "cancelled-partial-restore"
            add_result(_result_record(
                phase=context["Phase"],
                item_id="cancellation",
                description="Restore cancellation",
                success=False,
                started_at=utc_timestamp(),
                stderr=str(exc),
                evidence={"SafeCheckpoint": True},
            ))
        except Exception as exc:
            context["Outcome"] = "failed"
            if not context["Results"] or context["Results"][-1]["Success"]:
                add_result(_result_record(
                    phase=context["Phase"],
                    item_id="restore",
                    description="Restore transaction",
                    success=False,
                    started_at=utc_timestamp(),
                    stderr=str(exc),
                ))
        finally:
            context["Phase"] = "complete"
            context["CompletedAt"] = utc_timestamp()
            _append_jsonl(plan["RestoreHistoryPath"], context)
            if progress_callback:
                progress_callback(1.0)
    return context
