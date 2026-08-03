#!/usr/bin/env python3
"""Fail-closed Windows capability and AppX inventory discovery."""

from __future__ import annotations

import json
import platform
from typing import Any, Callable

from command_runner import CommandTimeoutError


CAPABILITY_SCHEMA_VERSION = 1
INVENTORY_SCHEMA_VERSION = 1
STATUS_SUCCESS = "success"
STATUS_EMPTY = "empty"
STATUS_UNAVAILABLE = "unavailable"
STATUS_DENIED = "denied"
STATUS_TIMED_OUT = "timed-out"
STATUS_POLICY_BLOCKED = "policy-blocked"
KNOWN_INVENTORY_STATUSES = frozenset({STATUS_SUCCESS, STATUS_EMPTY})

_DENIED_MARKERS = (
    "access is denied",
    "access denied",
    "unauthorized",
    "requires elevation",
    "administrator privileges",
    "0x80070005",
)
_POLICY_MARKERS = (
    "blocked by policy",
    "blocked by your administrator",
    "group policy",
    "organization's policy",
    "0x800704ec",
)
_UNAVAILABLE_MARKERS = (
    "is not recognized",
    "could not find",
    "not found",
    "get-appxpackage",
    "get-appxprovisionedpackage",
    "powershell",
)


class InventoryDiscoveryError(RuntimeError):
    """Raised when absence cannot be inferred from an inventory result."""

    def __init__(self, result: dict[str, Any]):
        self.result = dict(result)
        super().__init__(inventory_failure_text(result))


def inventory_is_known(result: dict[str, Any] | None) -> bool:
    return bool(
        isinstance(result, dict)
        and result.get("Status") in KNOWN_INVENTORY_STATUSES
    )


def inventory_failure_text(result: dict[str, Any]) -> str:
    status = str(result.get("Status") or STATUS_UNAVAILABLE)
    message = str(
        result.get("Message")
        or "Windows package inventory is unavailable"
    )
    next_action = str(result.get("NextAction") or "")
    if next_action:
        return f"{message} Next action: {next_action}"
    return f"{message} (status: {status})"


def _status_guidance(status: str, scope: str) -> tuple[str, str]:
    if status == STATUS_EMPTY:
        return (
            f"Windows reported an empty {scope} AppX/MSIX inventory.",
            (
                "Confirm the intended user or machine scope before "
                "treating packages as absent."
            ),
        )
    if status == STATUS_DENIED:
        return (
            f"Windows denied the {scope} AppX/MSIX inventory.",
            (
                "Run MSStoreHelper as Administrator for machine-wide "
                "and provisioned-package inventory."
            ),
        )
    if status == STATUS_TIMED_OUT:
        return (
            f"The {scope} AppX/MSIX inventory timed out.",
            (
                "Retry after AppX or DISM servicing completes; inspect "
                "the diagnostic bundle if the timeout repeats."
            ),
        )
    if status == STATUS_POLICY_BLOCKED:
        return (
            f"Windows policy blocked the {scope} AppX/MSIX inventory.",
            (
                "Review Store and App Installer policy with the device "
                "administrator; package absence was not inferred."
            ),
        )
    if status == STATUS_UNAVAILABLE:
        return (
            f"The {scope} AppX/MSIX inventory is unavailable.",
            (
                "Repair Windows PowerShell and the AppX/DISM cmdlets, "
                "then retry."
            ),
        )
    return (
        f"Windows returned the {scope} AppX/MSIX inventory.",
        "No action is required.",
    )


def _inventory_result(
    status: str,
    scope: str,
    *,
    records: list[dict[str, Any]] | None = None,
    error_code: str = "",
    detail: str = "",
) -> dict[str, Any]:
    message, next_action = _status_guidance(status, scope)
    return {
        "SchemaVersion": INVENTORY_SCHEMA_VERSION,
        "Status": status,
        "Known": status in KNOWN_INVENTORY_STATUSES,
        "Scope": scope,
        "Records": list(records or []),
        "Identities": sorted({
            str(record.get("Name") or "").strip().lower()
            for record in (records or [])
            if str(record.get("Name") or "").strip()
        }),
        "ErrorCode": str(error_code),
        "Detail": str(detail).strip(),
        "Message": message,
        "NextAction": next_action,
    }


def _classify_failure(detail: str) -> str:
    normalized = str(detail or "").lower()
    if any(marker in normalized for marker in _POLICY_MARKERS):
        return STATUS_POLICY_BLOCKED
    if any(marker in normalized for marker in _DENIED_MARKERS):
        return STATUS_DENIED
    if any(marker in normalized for marker in _UNAVAILABLE_MARKERS):
        return STATUS_UNAVAILABLE
    return STATUS_UNAVAILABLE


def _inventory_script(scope: str) -> str:
    installed_command = (
        "Get-AppxPackage -AllUsers -ErrorAction Stop"
        if scope == "machine"
        else "Get-AppxPackage -ErrorAction Stop"
    )
    provisioned_command = (
        "Get-AppxProvisionedPackage -Online -ErrorAction Stop"
        if scope == "machine"
        else "@()"
    )
    return "\n".join([
        "$ErrorActionPreference = 'Stop'",
        f"$installed = @({installed_command} | ForEach-Object {{",
        "    [pscustomobject]@{",
        "        Name = [string]$_.Name",
        "        Version = [string]$_.Version",
        "        Architecture = [string]$_.Architecture",
        "        Publisher = [string]$_.Publisher",
        "        PackageFamilyName = [string]$_.PackageFamilyName",
        "        PackageFullName = [string]$_.PackageFullName",
        "        ResourceId = [string]$_.ResourceId",
        "        IsFramework = [bool]$_.IsFramework",
        "        Source = 'installed'",
        "    }",
        "})",
        f"$provisioned = @({provisioned_command} | ForEach-Object {{",
        "    [pscustomobject]@{",
        "        Name = [string]$_.DisplayName",
        "        Version = [string]$_.Version",
        "        Architecture = [string]$_.Architecture",
        "        PublisherId = [string]$_.PublisherId",
        "        PackageFullName = [string]$_.PackageName",
        "        ResourceId = [string]$_.ResourceId",
        "        Source = 'provisioned'",
        "    }",
        "})",
        "[pscustomobject]@{",
        f"    SchemaVersion = {INVENTORY_SCHEMA_VERSION}",
        f"    Scope = '{scope}'",
        "    Installed = $installed",
        "    Provisioned = $provisioned",
        "} | ConvertTo-Json -Compress -Depth 5",
    ])


def query_appx_inventory(
    runner: Callable[..., Any],
    powershell_exe: str,
    *,
    scope: str = "current-user",
    is_admin: bool = False,
    timeout: float = 60,
) -> dict[str, Any]:
    scope = str(scope or "current-user").strip().lower()
    if scope not in {"current-user", "machine"}:
        raise ValueError("inventory scope must be current-user or machine")
    if scope == "machine" and not is_admin:
        return _inventory_result(
            STATUS_DENIED,
            scope,
            error_code="elevation-required",
            detail="Machine inventory requires an elevated process.",
        )

    try:
        result = runner(
            [
                powershell_exe,
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                _inventory_script(scope),
            ],
            timeout=timeout,
        )
    except CommandTimeoutError as exc:
        return _inventory_result(
            STATUS_TIMED_OUT,
            scope,
            error_code="command-timeout",
            detail=str(exc),
        )
    except (OSError, ValueError) as exc:
        return _inventory_result(
            STATUS_UNAVAILABLE,
            scope,
            error_code="command-unavailable",
            detail=str(exc),
        )

    stdout = str(getattr(result, "stdout", "") or "").strip()
    stderr = str(getattr(result, "stderr", "") or "").strip()
    returncode = int(getattr(result, "returncode", 1))
    if returncode:
        detail = stderr or stdout or f"PowerShell exited {returncode}"
        status = _classify_failure(detail)
        return _inventory_result(
            status,
            scope,
            error_code=f"powershell-exit-{returncode}",
            detail=detail,
        )
    if not stdout:
        return _inventory_result(
            STATUS_UNAVAILABLE,
            scope,
            error_code="empty-command-output",
            detail="PowerShell did not return an inventory envelope.",
        )

    try:
        payload = json.loads(stdout)
        if (
            not isinstance(payload, dict)
            or int(payload.get("SchemaVersion", 0))
            != INVENTORY_SCHEMA_VERSION
            or payload.get("Scope") != scope
        ):
            raise ValueError("inventory envelope schema is invalid")
        installed = payload.get("Installed") or []
        provisioned = payload.get("Provisioned") or []
        if isinstance(installed, dict):
            installed = [installed]
        if isinstance(provisioned, dict):
            provisioned = [provisioned]
        if not isinstance(installed, list) or not isinstance(
            provisioned,
            list,
        ):
            raise ValueError("inventory record arrays are invalid")
        records = []
        for record in installed + provisioned:
            if not isinstance(record, dict):
                raise ValueError("inventory record is invalid")
            name = str(record.get("Name") or "").strip()
            if not name:
                continue
            records.append({
                "Name": name,
                "Version": str(record.get("Version") or "").strip(),
                "Source": str(record.get("Source") or "").strip(),
            })
    except (TypeError, ValueError, json.JSONDecodeError) as exc:
        return _inventory_result(
            STATUS_UNAVAILABLE,
            scope,
            error_code="invalid-command-output",
            detail=str(exc),
        )

    return _inventory_result(
        STATUS_SUCCESS if records else STATUS_EMPTY,
        scope,
        records=records,
    )


def _capability_script() -> str:
    return r"""
$ErrorActionPreference = 'Stop'
function Read-PolicyValue([string]$Path, [string]$Name) {
    try {
        $value = Get-ItemPropertyValue -LiteralPath $Path -Name $Name -ErrorAction Stop
        return [int]$value
    } catch {
        return $null
    }
}
$currentVersion = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
$os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$isElevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$serviceNames = @('AppXSvc', 'ClipSVC', 'InstallService', 'LicenseManager', 'wuauserv')
$services = @($serviceNames | ForEach-Object {
    $name = $_
    $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$name'" -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        [pscustomobject]@{ Name = $name; Exists = $false; State = 'missing'; StartMode = 'unknown' }
    } else {
        [pscustomobject]@{
            Name = $name
            Exists = $true
            State = [string]$service.State
            StartMode = [string]$service.StartMode
        }
    }
})
$componentBasedServicing = Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
$windowsUpdate = Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
$pendingRename = $false
try {
    $pendingRename = $null -ne (Get-ItemPropertyValue -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction Stop)
} catch {}
[pscustomobject]@{
    SchemaVersion = 1
    OS = [pscustomobject]@{
        Caption = [string]$os.Caption
        Edition = [string]$currentVersion.EditionID
        ProductName = [string]$currentVersion.ProductName
        DisplayVersion = [string]$currentVersion.DisplayVersion
        Build = [string]$currentVersion.CurrentBuildNumber
        UBR = [int]$currentVersion.UBR
        Architecture = [string]$os.OSArchitecture
    }
    Context = [pscustomobject]@{
        IsElevated = [bool]$isElevated
        IsSystem = [bool]($identity.User.Value -eq 'S-1-5-18')
        InventoryScopes = @('current-user', 'machine')
    }
    Policies = [pscustomobject]@{
        RemoveWindowsStore = Read-PolicyValue 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore' 'RemoveWindowsStore'
        DisableStoreApps = Read-PolicyValue 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore' 'DisableStoreApps'
        EnableAppInstaller = Read-PolicyValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppInstaller' 'EnableAppInstaller'
    }
    Services = $services
    RebootPending = [pscustomobject]@{
        Pending = [bool]($componentBasedServicing -or $windowsUpdate -or $pendingRename)
        ComponentBasedServicing = [bool]$componentBasedServicing
        WindowsUpdate = [bool]$windowsUpdate
        PendingFileRenameOperations = [bool]$pendingRename
    }
} | ConvertTo-Json -Compress -Depth 7
""".strip()


def _empty_capability_result(
    status: str,
    *,
    detail: str = "",
    source_health: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    return {
        "SchemaVersion": CAPABILITY_SCHEMA_VERSION,
        "Status": status,
        "Platform": {
            "System": platform.system() or "unknown",
            "Edition": "unknown",
            "Build": "unknown",
            "Architecture": platform.machine() or "unknown",
        },
        "Context": {
            "InventoryScope": "current-user",
            "MachineInventoryRequiresElevation": True,
            "IsElevated": False,
            "IsSystem": False,
        },
        "Policies": {
            "Store": "unknown",
            "AppInstaller": "unknown",
        },
        "Services": [],
        "Network": _network_status(source_health),
        "RebootPending": {
            "Pending": None,
            "State": "unknown",
        },
        "Blockers": [],
        "Warnings": [],
        "Detail": str(detail).strip(),
    }


def _network_status(
    source_health: list[dict[str, Any]] | None,
) -> dict[str, Any]:
    endpoints = []
    for status in source_health or []:
        if not isinstance(status, dict):
            continue
        endpoints.append({
            "Key": str(status.get("Key") or "source"),
            "Available": (
                bool(status.get("Available"))
                if "Available" in status
                else None
            ),
            "Detail": str(status.get("Detail") or ""),
        })
    known = [item for item in endpoints if item["Available"] is not None]
    if not known:
        state = "unknown"
    elif any(item["Available"] for item in known):
        state = "available"
    else:
        state = "unavailable"
    return {
        "Status": state,
        "Endpoints": endpoints,
    }


def probe_windows_capabilities(
    runner: Callable[..., Any],
    powershell_exe: str,
    *,
    is_admin: bool,
    source_health: list[dict[str, Any]] | None = None,
    timeout: float = 60,
) -> dict[str, Any]:
    if platform.system() != "Windows":
        result = _empty_capability_result(
            STATUS_UNAVAILABLE,
            detail="MSStoreHelper Windows capabilities require Windows.",
            source_health=source_health,
        )
        result["Blockers"].append({
            "Code": "unsupported-platform",
            "Message": "The current operating system is not Windows.",
            "NextAction": "Run MSStoreHelper on a supported Windows build.",
        })
        return result
    try:
        completed = runner(
            [
                powershell_exe,
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                _capability_script(),
            ],
            timeout=timeout,
        )
    except CommandTimeoutError as exc:
        return _empty_capability_result(
            STATUS_TIMED_OUT,
            detail=str(exc),
            source_health=source_health,
        )
    except (OSError, ValueError) as exc:
        return _empty_capability_result(
            STATUS_UNAVAILABLE,
            detail=str(exc),
            source_health=source_health,
        )

    stdout = str(getattr(completed, "stdout", "") or "").strip()
    stderr = str(getattr(completed, "stderr", "") or "").strip()
    returncode = int(getattr(completed, "returncode", 1))
    if returncode or not stdout:
        status = _classify_failure(stderr or stdout)
        return _empty_capability_result(
            status,
            detail=stderr or stdout or f"PowerShell exited {returncode}",
            source_health=source_health,
        )
    try:
        payload = json.loads(stdout)
        if (
            not isinstance(payload, dict)
            or int(payload.get("SchemaVersion", 0))
            != CAPABILITY_SCHEMA_VERSION
        ):
            raise ValueError("capability envelope schema is invalid")
    except (TypeError, ValueError, json.JSONDecodeError) as exc:
        return _empty_capability_result(
            STATUS_UNAVAILABLE,
            detail=str(exc),
            source_health=source_health,
        )

    os_data = payload.get("OS") or {}
    context = payload.get("Context") or {}
    policies = payload.get("Policies") or {}
    services = payload.get("Services") or []
    reboot = payload.get("RebootPending") or {}
    if isinstance(services, dict):
        services = [services]
    result = _empty_capability_result(
        STATUS_SUCCESS,
        source_health=source_health,
    )
    result["Platform"] = {
        "System": "Windows",
        "Caption": str(os_data.get("Caption") or ""),
        "Edition": str(os_data.get("Edition") or "unknown"),
        "ProductName": str(os_data.get("ProductName") or ""),
        "DisplayVersion": str(os_data.get("DisplayVersion") or ""),
        "Build": (
            f"{os_data.get('Build')}.{os_data.get('UBR')}"
            if os_data.get("Build") not in (None, "")
            else "unknown"
        ),
        "Architecture": str(
            os_data.get("Architecture")
            or platform.machine()
            or "unknown"
        ),
    }
    elevated = bool(context.get("IsElevated", is_admin))
    result["Context"] = {
        "InventoryScope": "machine" if elevated else "current-user",
        "MachineInventoryRequiresElevation": not elevated,
        "IsElevated": elevated,
        "IsSystem": bool(context.get("IsSystem", False)),
    }
    store_blocked = (
        policies.get("RemoveWindowsStore") == 1
        or policies.get("DisableStoreApps") == 1
    )
    app_installer_value = policies.get("EnableAppInstaller")
    app_installer_blocked = app_installer_value == 0
    result["Policies"] = {
        "Store": "blocked" if store_blocked else "allowed-or-unconfigured",
        "AppInstaller": (
            "blocked"
            if app_installer_blocked
            else "allowed-or-unconfigured"
        ),
        "Raw": {
            "RemoveWindowsStore": policies.get("RemoveWindowsStore"),
            "DisableStoreApps": policies.get("DisableStoreApps"),
            "EnableAppInstaller": app_installer_value,
        },
    }
    if store_blocked:
        result["Blockers"].append({
            "Code": "store-policy-blocked",
            "Message": "Windows Store access is disabled by policy.",
            "NextAction": (
                "Use only approved offline workflows or ask the device "
                "administrator to review Windows Store policy."
            ),
        })
    if app_installer_blocked:
        result["Blockers"].append({
            "Code": "app-installer-policy-blocked",
            "Message": "App Installer is disabled by policy.",
            "NextAction": (
                "Use an approved deployment path or ask the device "
                "administrator to review App Installer policy."
            ),
        })

    normalized_services = []
    for service in services:
        if not isinstance(service, dict):
            continue
        normalized = {
            "Name": str(service.get("Name") or ""),
            "Exists": bool(service.get("Exists", False)),
            "State": str(service.get("State") or "unknown").lower(),
            "StartMode": str(
                service.get("StartMode") or "unknown"
            ).lower(),
        }
        normalized_services.append(normalized)
        if not normalized["Exists"]:
            result["Blockers"].append({
                "Code": f"service-missing:{normalized['Name']}",
                "Message": (
                    f"Required Windows service {normalized['Name']} "
                    "is not installed."
                ),
                "NextAction": (
                    "Repair the Windows AppX/Store servicing components "
                    "before remediation."
                ),
            })
        elif normalized["StartMode"] == "disabled":
            result["Blockers"].append({
                "Code": f"service-disabled:{normalized['Name']}",
                "Message": (
                    f"Required Windows service {normalized['Name']} "
                    "is disabled."
                ),
                "NextAction": (
                    "Restore the service startup policy before "
                    "Store/AppX remediation."
                ),
            })
        elif normalized["State"] != "running":
            result["Warnings"].append({
                "Code": f"service-stopped:{normalized['Name']}",
                "Message": (
                    f"Windows service {normalized['Name']} is "
                    f"{normalized['State']}."
                ),
                "NextAction": (
                    "Windows may start the service on demand; if the "
                    "workflow fails, restore its normal startup policy."
                ),
            })
    result["Services"] = normalized_services
    pending = bool(reboot.get("Pending", False))
    result["RebootPending"] = {
        "Pending": pending,
        "State": "pending" if pending else "clear",
        "Signals": {
            "ComponentBasedServicing": bool(
                reboot.get("ComponentBasedServicing", False)
            ),
            "WindowsUpdate": bool(
                reboot.get("WindowsUpdate", False)
            ),
            "PendingFileRenameOperations": bool(
                reboot.get("PendingFileRenameOperations", False)
            ),
        },
    }
    if pending:
        result["Warnings"].append({
            "Code": "reboot-pending",
            "Message": (
                "Windows reports a pending reboot; package and service "
                "state may change after restart."
            ),
            "NextAction": (
                "Restart Windows before repair or deployment when "
                "operationally safe, then rerun the preflight."
            ),
        })
    if result["Network"]["Status"] == "unavailable":
        result["Warnings"].append({
            "Code": "store-sources-unavailable",
            "Message": (
                "No probed Store source endpoint is currently available."
            ),
            "NextAction": (
                "Restore the endpoint or use a reviewed offline cache."
            ),
        })
    if result["Blockers"]:
        policy_only = all(
            str(item.get("Code", "")).endswith("policy-blocked")
            for item in result["Blockers"]
        )
        result["Status"] = (
            STATUS_POLICY_BLOCKED
            if policy_only
            else STATUS_UNAVAILABLE
        )
    return result


def capability_summary(report: dict[str, Any]) -> str:
    platform_info = report.get("Platform") or {}
    context = report.get("Context") or {}
    network = report.get("Network") or {}
    return (
        f"Windows capability status={report.get('Status', 'unknown')}; "
        f"edition={platform_info.get('Edition', 'unknown')}; "
        f"build={platform_info.get('Build', 'unknown')}; "
        f"scope={context.get('InventoryScope', 'unknown')}; "
        f"elevated={bool(context.get('IsElevated'))}; "
        f"network={network.get('Status', 'unknown')}; "
        f"reboot={report.get('RebootPending', {}).get('State', 'unknown')}"
    )


def capability_blocking_text(
    report: dict[str, Any],
    *,
    required_sources: set[str] | None = None,
    required_services: set[str] | None = None,
    respect_policy_codes: set[str] | None = None,
) -> str:
    """Return an exact blocker/action pair or an empty string."""
    status = str(report.get("Status") or STATUS_UNAVAILABLE)
    blockers = [
        item
        for item in (report.get("Blockers") or [])
        if isinstance(item, dict)
    ]
    relevant = []
    respected = set(respect_policy_codes or set())
    services = {
        str(name).strip().lower()
        for name in (required_services or set())
        if str(name).strip()
    }
    for blocker in blockers:
        code = str(blocker.get("Code") or "")
        if code.endswith("policy-blocked") and code not in respected:
            continue
        if code.startswith(("service-missing:", "service-disabled:")):
            service_name = code.split(":", 1)[1].lower()
            if service_name not in services:
                continue
        relevant.append(blocker)
    if (
        status not in {STATUS_SUCCESS, STATUS_POLICY_BLOCKED}
        and not relevant
        and not blockers
    ):
        detail = str(report.get("Detail") or "").strip()
        message = (
            "Windows capability discovery did not complete; "
            "no remediation was started."
        )
        action = (
            "Retry the capability check and save diagnostics if it "
            "continues to fail."
        )
        return " ".join(
            item for item in (message, detail, f"Next action: {action}")
            if item
        )
    if relevant:
        blocker = relevant[0]
        return (
            f"{blocker.get('Message', 'Windows capability is blocked')} "
            f"Next action: {blocker.get('NextAction', 'Review diagnostics.')}"
        )

    required = set(required_sources or set())
    if required:
        endpoints = {
            str(item.get("Key") or ""): item
            for item in (
                (report.get("Network") or {}).get("Endpoints") or []
            )
            if isinstance(item, dict)
        }
        unavailable = [
            key
            for key in sorted(required)
            if not bool((endpoints.get(key) or {}).get("Available"))
        ]
        if unavailable:
            return (
                "Required Store source preflight failed for "
                f"{', '.join(unavailable)}. Next action: restore the "
                "endpoint or use a reviewed offline cache; no bulk "
                "queue was derived."
            )
    return ""
