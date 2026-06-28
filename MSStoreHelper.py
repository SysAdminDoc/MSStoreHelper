#!/usr/bin/env python3
"""
MSStoreHelper - Microsoft Store App Installer for Windows LTSC
A user-friendly tool to download and install Microsoft Store apps
without needing access to the Microsoft Store.
"""

import sys
import subprocess
import os
import platform
import threading
import ctypes
import webbrowser
import json
import shutil
import tempfile
from tkinter import filedialog
from datetime import datetime, timezone
from msstore_package_resolution import (
    annotate_package,
    format_version_tuple,
    installed_version_satisfies_package,
    is_dependency_package,
    is_installable_package,
    order_packages_for_install,
    package_identity,
    package_version_tuple,
    package_role_label,
    select_recommended_packages,
    signature_info_is_valid_microsoft,
)

# ==================== DEPENDENCY AUTO-INSTALL ====================
def install_requirements():
    required = {
        'customtkinter': 'customtkinter',
        'requests': 'requests',
        'bs4': 'beautifulsoup4',
    }
    missing = []
    
    for import_name, pip_name in required.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)
    
    if missing:
        print(f"🔧 Installing: {', '.join(missing)}...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing], 
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("✅ Done! Restarting...")
        os.execv(sys.executable, ['python'] + sys.argv)

try:
    import customtkinter as ctk
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    install_requirements()
    import customtkinter as ctk
    import requests
    from bs4 import BeautifulSoup

# ==================== CONFIGURATION ====================

APP_VERSION = "3.16.0"
APP_NAME = "MSStoreHelper"
API_URL = "https://store.rg-adguard.net/api/GetFiles"
STORE_SEARCH_URL = "https://storeedgefd.dsx.mp.microsoft.com/v9.0/manifestSearch"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
WINGET_IMPORT_SCHEMA = "https://aka.ms/winget-packages.schema.2.0.json"
WINGET_MSSTORE_SOURCE = {
    "Argument": "https://storeedgefd.dsx.mp.microsoft.com/v9.0",
    "Identifier": "StoreEdgeFD",
    "Name": "msstore",
    "Type": "Microsoft.Rest",
}
WINDOWS_DIR = os.environ.get("WINDIR", r"C:\Windows")
WINDOWS_POWERSHELL = os.path.join(WINDOWS_DIR, "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
POWERSHELL_EXE = WINDOWS_POWERSHELL if os.path.exists(WINDOWS_POWERSHELL) else "powershell"
POWERSHELL_SECURITY_MODULE = os.path.join(
    WINDOWS_DIR,
    "System32",
    "WindowsPowerShell",
    "v1.0",
    "Modules",
    "Microsoft.PowerShell.Security",
    "Microsoft.PowerShell.Security.psd1",
)

try:
    DEFAULT_OUTPUT = os.path.join(os.environ['USERPROFILE'], "Downloads", "MSStoreHelper")
except:
    DEFAULT_OUTPUT = os.path.join(os.path.expanduser("~"), "Downloads", "MSStoreHelper")

try:
    IS_ADMIN = ctypes.windll.shell32.IsUserAnAdmin() != 0
except:
    IS_ADMIN = False

def get_architecture():
    arch = platform.machine().lower()
    if 'amd64' in arch or 'x86_64' in arch: return 'x64'
    if 'arm64' in arch: return 'arm64'
    if 'x86' in arch: return 'x86'
    return 'neutral'

SYSTEM_ARCH = get_architecture()

# ==================== COLOR THEME ====================
class Theme:
    # Main colors
    BG_DARK = "#0f0f1a"
    BG_CARD = "#1a1a2e"
    BG_CARD_HOVER = "#252542"
    BG_INPUT = "#16213e"
    
    # Accent colors
    PRIMARY = "#6366f1"
    PRIMARY_HOVER = "#818cf8"
    SUCCESS = "#10b981"
    SUCCESS_HOVER = "#34d399"
    WARNING = "#f59e0b"
    DANGER = "#ef4444"
    DANGER_HOVER = "#f87171"
    INFO = "#06b6d4"
    
    # Text colors
    TEXT_PRIMARY = "#f8fafc"
    TEXT_SECONDARY = "#94a3b8"
    TEXT_MUTED = "#64748b"
    
    # Special
    BORDER = "#2a2a4a"
    BUNDLE_COLOR = "#22d3ee"
    ENCRYPTED_COLOR = "#f87171"
    ARCH_MATCH = "#4ade80"

# ==================== HELPER FUNCTIONS ====================

def format_size(size_bytes):
    if size_bytes is None or size_bytes == 0:
        return "—"
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"

def catalog_apps_by_name():
    apps = {}
    for category in APP_CATALOG.values():
        for app in category["apps"]:
            apps[app["Name"]] = app
    return apps

# ==================== APP CATALOG ====================

APP_CATALOG = {
    "🛠️ Essential Repairs": {
        "description": "Fix common Windows Store issues",
        "apps": [
            {"Name": "Microsoft Store", "ProductId": "9WZDNCRFJBMP", "Description": "The main Store app", "Icon": "🏪"},
            {"Name": "App Installer", "ProductId": "9NBLGGH4NNS1", "Description": "Install apps & WinGet CLI", "Icon": "📦"},
            {"Name": "Xbox Identity", "ProductId": "9WZDNCRD1HKW", "Description": "Xbox sign-in support", "Icon": "🎮"},
        ]
    },
    "⚙️ System Components": {
        "description": "Required runtime libraries",
        "apps": [
            {"Name": "VC++ Runtime", "ProductId": "9WZDNCRFJ3PT", "Description": "Visual C++ 2015-2022", "Icon": "⚙️"},
            {"Name": "HEVC Codec", "ProductId": "9NMZLZ57R3T7", "Description": "H.265 video support", "Icon": "🎬"},
            {"Name": "AV1 Codec", "ProductId": "9MVZQVXJBQ9V", "Description": "AV1 video support", "Icon": "🎬"},
            {"Name": "WebP Images", "ProductId": "9PG2DK419DRG", "Description": "WebP format support", "Icon": "🖼️"},
        ]
    },
    "💻 Productivity": {
        "description": "Essential Windows apps",
        "apps": [
            {"Name": "Windows Terminal", "ProductId": "9N0DX20HK701", "Description": "Modern command line", "Icon": "💻"},
            {"Name": "PowerToys", "ProductId": "XP89DCGQ3K6VLD", "Description": "Power user utilities", "Icon": "🔧"},
            {"Name": "Notepad", "ProductId": "9MSMLRH6LZF3", "Description": "Modern text editor", "Icon": "📝"},
            {"Name": "Calculator", "ProductId": "9WZDNCRFHVN5", "Description": "Windows Calculator", "Icon": "🔢"},
            {"Name": "Snipping Tool", "ProductId": "9MZ95KL8MR0L", "Description": "Screenshot tool", "Icon": "✂️"},
            {"Name": "Photos", "ProductId": "9WZDNCRFJBH4", "Description": "Photo viewer & editor", "Icon": "📷"},
        ]
    },
    "🎮 Gaming": {
        "description": "Xbox and gaming services",
        "apps": [
            {"Name": "Xbox App", "ProductId": "9MV0B5HZVK9Z", "Description": "Xbox for PC", "Icon": "🎮"},
            {"Name": "Xbox Game Bar", "ProductId": "9NZKPSTSNW4P", "Description": "In-game overlay", "Icon": "🎯"},
            {"Name": "Gaming Services", "ProductId": "9MWPM2CQNLHN", "Description": "Core gaming support", "Icon": "🕹️"},
        ]
    },
    "🌐 Browsers": {
        "description": "Web browsers",
        "apps": [
            {"Name": "Firefox", "ProductId": "9NZVDKPMR9RD", "Description": "Mozilla Firefox", "Icon": "🦊"},
            {"Name": "Brave", "ProductId": "9P0HQXFZKMFJ", "Description": "Privacy browser", "Icon": "🦁"},
        ]
    },
    "🛠️ Developer Tools": {
        "description": "For developers",
        "apps": [
            {"Name": "VS Code", "ProductId": "XP9KHM4BK9FZ7Q", "Description": "Code editor", "Icon": "📘"},
            {"Name": "Python 3.12", "ProductId": "9NCVDN91XZQP", "Description": "Python language", "Icon": "🐍"},
            {"Name": "PowerShell 7", "ProductId": "9MZ1SNWT0N5D", "Description": "Modern PowerShell", "Icon": "⚡"},
            {"Name": "WSL", "ProductId": "9P9TQF7MRM4R", "Description": "Linux on Windows", "Icon": "🐧"},
        ]
    },
}

QUICK_FIX_PRESETS = {
    "🧰 LTSC Essentials": {
        "description": "Queue core apps commonly missing on LTSC: Terminal, PowerShell, WSL, Photos, Calculator, and Snipping Tool.",
        "apps": ["Windows Terminal", "PowerShell 7", "WSL", "Photos", "Calculator", "Snipping Tool"]
    },
    "🏪 Repair Store": {
        "description": "Reinstall Microsoft Store and essential components to fix most Store-related issues.",
        "apps": ["Microsoft Store", "App Installer", "VC++ Runtime"]
    },
    "🎮 Gaming Setup": {
        "description": "Install Xbox app, Game Bar, and gaming services for PC gaming.",
        "apps": ["Xbox App", "Xbox Game Bar", "Xbox Identity", "Gaming Services"]
    },
    "🎬 Media Codecs": {
        "description": "Add support for modern video formats (HEVC, AV1) and image formats.",
        "apps": ["HEVC Codec", "AV1 Codec", "WebP Images"]
    },
    "💻 Developer Pack": {
        "description": "Essential tools for developers: Terminal, PowerShell 7, and VS Code.",
        "apps": ["Windows Terminal", "PowerShell 7", "VS Code"]
    },
}

LTSC_COMPONENT_REQUIREMENTS = [
    {"Name": "Microsoft Store", "Identities": ["Microsoft.WindowsStore"]},
    {"Name": "App Installer", "Identities": ["Microsoft.DesktopAppInstaller"]},
    {"Name": "VC++ Runtime", "Identities": ["Microsoft.VCLibs.140.00"]},
    {"Name": "Windows Terminal", "Identities": ["Microsoft.WindowsTerminal"]},
    {"Name": "PowerShell 7", "Identities": ["Microsoft.PowerShell"]},
    {"Name": "WSL", "Identities": ["MicrosoftCorporationII.WindowsSubsystemForLinux"]},
    {"Name": "Photos", "Identities": ["Microsoft.Windows.Photos"]},
    {"Name": "Calculator", "Identities": ["Microsoft.WindowsCalculator"]},
    {"Name": "Snipping Tool", "Identities": ["Microsoft.ScreenSketch"]},
    {"Name": "HEVC Codec", "Identities": ["Microsoft.HEVCVideoExtension"]},
    {"Name": "AV1 Codec", "Identities": ["Microsoft.AV1VideoExtension"]},
    {"Name": "WebP Images", "Identities": ["Microsoft.WebpImageExtension"]},
]

# ==================== BACKEND API ====================

class StoreAPI:
    """Handles all API communications"""
    
    @staticmethod
    def search_store(query, max_results=25):
        """Search Microsoft Store by app name"""
        try:
            payload = {"Query": {"KeyWord": query, "MatchType": "Substring"}}
            headers = {"User-Agent": USER_AGENT, "Content-Type": "application/json"}
            
            resp = requests.post(STORE_SEARCH_URL, json=payload, headers=headers, timeout=15)
            resp.raise_for_status()
            data = resp.json()
            
            results = []
            for item in data.get("Data", [])[:max_results]:
                results.append({
                    "ProductId": item.get("PackageIdentifier", ""),
                    "Name": item.get("PackageName", "Unknown"),
                    "Publisher": item.get("Publisher", "Unknown"),
                })
            return results
            
        except Exception as e:
            print(f"Search error: {e}")
            return []
    
    @staticmethod
    def get_packages(product_id, ring="Retail"):
        """Get downloadable packages for a product"""
        try:
            payload = {"type": "ProductId", "url": product_id, "ring": ring, "lang": "en-US"}
            headers = {"User-Agent": USER_AGENT, "Content-Type": "application/x-www-form-urlencoded"}
            
            resp = requests.post(API_URL, data=payload, headers=headers, timeout=30)
            resp.raise_for_status()
            
            soup = BeautifulSoup(resp.text, "html.parser")
            table = soup.find("table", class_="tftable")
            
            if not table:
                return []

            results = []
            for row in table.find_all("tr"):
                cols = row.find_all("td")
                if not cols:
                    continue
                
                link = cols[0].find("a")
                if not link:
                    continue
                
                name = link.text.strip()
                url = link['href']
                
                if name.endswith(".BlockMap"):
                    continue
                
                arch = "neutral"
                lower = name.lower()
                if "_x64_" in lower: arch = "x64"
                elif "_x86_" in lower: arch = "x86"
                elif "_arm64_" in lower: arch = "arm64"
                elif "_arm_" in lower: arch = "arm"
                
                ext = os.path.splitext(name)[1].lower().replace(".", "").upper()
                is_bundle = "BUNDLE" in ext
                is_encrypted = ext.startswith("E")
                
                package = {
                    "FileName": name, "Url": url, "Architecture": arch,
                    "FileType": ext, "IsBundle": is_bundle, "IsEncrypted": is_encrypted,
                    "SizeBytes": None, "SizeStr": "—"
                }
                results.append(annotate_package(package))
            
            return results
            
        except Exception as e:
            print(f"API error: {e}")
            return []
    
    @staticmethod
    def get_file_size(url):
        try:
            resp = requests.head(url, timeout=10, allow_redirects=True)
            size = resp.headers.get('content-length')
            return int(size) if size else None
        except:
            return None
    
    @staticmethod
    def smart_select(packages, target_arch, prefer_exact_arch=False):
        """Intelligently select the best packages"""
        return select_recommended_packages(packages, target_arch, prefer_exact_arch)

    @staticmethod
    def order_packages_for_install(packages, target_arch):
        return order_packages_for_install(packages, target_arch)
    
    @staticmethod
    def download_file(url, filepath, progress_callback=None):
        try:
            with requests.get(url, stream=True, timeout=60) as r:
                r.raise_for_status()
                total = int(r.headers.get('content-length', 0))
                
                with open(filepath, 'wb') as f:
                    downloaded = 0
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total:
                            progress_callback(downloaded / total)
            return True, "Success"
        except Exception as e:
            return False, str(e)

    @staticmethod
    def is_cacheable_artifact(filename):
        return os.path.splitext(filename)[1].lower() in {".appx", ".msix", ".appxbundle", ".msixbundle"}

    @staticmethod
    def cache_downloaded_artifact(package, cache_path):
        local_path = package.get("LocalPath")
        filename = package.get("FileName", os.path.basename(local_path or ""))
        if not local_path or not os.path.exists(local_path):
            return False, "Downloaded file is missing"
        if not StoreAPI.is_cacheable_artifact(filename):
            return False, "File type is not cacheable"

        os.makedirs(cache_path, exist_ok=True)
        destination = os.path.join(cache_path, filename)
        if os.path.exists(destination) and os.path.getsize(destination) == os.path.getsize(local_path):
            return True, f"Already cached: {destination}"

        shutil.copy2(local_path, destination)
        return True, f"Cached: {destination}"

    @staticmethod
    def _powershell_literal(value):
        return "'" + str(value).replace("'", "''") + "'"

    @staticmethod
    def _portable_script_path(package_path, script_dir):
        absolute_path = os.path.abspath(package_path)
        if not script_dir:
            return absolute_path

        absolute_script_dir = os.path.abspath(script_dir)
        try:
            if os.path.commonpath([absolute_script_dir, absolute_path]).lower() == absolute_script_dir.lower():
                return os.path.relpath(absolute_path, absolute_script_dir)
        except ValueError:
            pass
        return absolute_path

    @staticmethod
    def generate_dism_provision_script(packages, output_path, target_arch=SYSTEM_ARCH, script_dir=None):
        provisionable = []
        for package in packages:
            if not package.get("FileName") or not is_installable_package(package):
                continue

            provisionable.append(annotate_package(package.copy()))

        provisionable = StoreAPI.order_packages_for_install(provisionable, target_arch)
        if not provisionable:
            raise ValueError("No AppX/MSIX packages are available in the queue")

        script_dir = script_dir or output_path
        generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines = [
            f"# Generated by {APP_NAME} v{APP_VERSION} on {generated_at}",
            "# Run from an elevated PowerShell session.",
            "$ErrorActionPreference = 'Stop'",
            "",
            "$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }",
            "",
            "function Resolve-QueuePackagePath {",
            "    param([string]$Path)",
            "    if ([System.IO.Path]::IsPathRooted($Path)) { return $Path }",
            "    return Join-Path -Path $ScriptRoot -ChildPath $Path",
            "}",
            "",
            "$packages = @(",
        ]

        for package in provisionable:
            filename = package["FileName"]
            package_path = package.get("LocalPath") or os.path.join(output_path, filename)
            portable_path = StoreAPI._portable_script_path(package_path, script_dir)
            role = package.get("PackageRoleLabel") or package_role_label(filename)
            lines.append(
                "    [pscustomobject]@{ "
                f"FileName = {StoreAPI._powershell_literal(filename)}; "
                f"Role = {StoreAPI._powershell_literal(role)}; "
                f"PackagePath = {StoreAPI._powershell_literal(portable_path)} "
                "}"
            )

        lines.extend([
            ")",
            "",
            "foreach ($package in $packages) {",
            "    $packagePath = Resolve-QueuePackagePath $package.PackagePath",
            "    if (-not (Test-Path -LiteralPath $packagePath)) {",
            "        throw \"Package not found: $packagePath\"",
            "    }",
            "",
            "    Write-Host (\"Provisioning {0} [{1}]\" -f $package.FileName, $package.Role)",
            "    $arguments = @(",
            "        '/Online',",
            "        '/Add-ProvisionedAppxPackage',",
            "        \"/PackagePath:$packagePath\",",
            "        '/SkipLicense'",
            "    )",
            "    & dism.exe @arguments",
            "    if ($LASTEXITCODE -ne 0) {",
            "        throw \"DISM failed with exit code $LASTEXITCODE for $($package.FileName)\"",
            "    }",
            "}",
            "",
            "Write-Host \"Provisioning script complete: $($packages.Count) package(s).\"",
            "",
        ])
        return "\n".join(lines)

    @staticmethod
    def write_dism_provision_script(packages, output_path, script_path, target_arch=SYSTEM_ARCH):
        script_dir = os.path.dirname(os.path.abspath(script_path))
        script = StoreAPI.generate_dism_provision_script(packages, output_path, target_arch, script_dir)
        os.makedirs(script_dir, exist_ok=True)
        with open(script_path, "w", encoding="utf-8", newline="\r\n") as handle:
            handle.write(script)
        return script_path

    @staticmethod
    def get_winget_version():
        try:
            result = subprocess.run(
                ["winget", "--version"],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            version = (result.stdout or "").strip()
            return version.lstrip("v") if result.returncode == 0 and version else None
        except Exception:
            return None

    @staticmethod
    def _winget_creation_date(created_at=None):
        created_at = created_at or datetime.now(timezone.utc)
        if isinstance(created_at, datetime):
            if created_at.tzinfo is None:
                created_at = created_at.replace(tzinfo=timezone.utc)
            return created_at.astimezone(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "-00:00")
        return str(created_at)

    @staticmethod
    def build_winget_import_manifest(apps, winget_version=None, created_at=None):
        packages = []
        seen = set()
        for app in apps:
            package_id = str(app.get("ProductId") or "").strip()
            if not package_id:
                continue
            key = package_id.lower()
            if key in seen:
                continue
            seen.add(key)
            packages.append({"PackageIdentifier": package_id})

        if not packages:
            raise ValueError("No selected apps have WinGet package identifiers")

        manifest = {
            "$schema": WINGET_IMPORT_SCHEMA,
            "CreationDate": StoreAPI._winget_creation_date(created_at),
            "Sources": [
                {
                    "Packages": packages,
                    "SourceDetails": WINGET_MSSTORE_SOURCE.copy(),
                }
            ],
        }
        if winget_version:
            manifest["WinGetVersion"] = str(winget_version).lstrip("v")
        return manifest

    @staticmethod
    def write_winget_import_manifest(apps, manifest_path, winget_version=None, created_at=None):
        detected_version = winget_version if winget_version is not None else StoreAPI.get_winget_version()
        manifest = StoreAPI.build_winget_import_manifest(apps, detected_version, created_at)
        manifest_dir = os.path.dirname(os.path.abspath(manifest_path))
        os.makedirs(manifest_dir, exist_ok=True)
        with open(manifest_path, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(manifest, handle, indent=2)
            handle.write("\n")
        return manifest_path, len(manifest["Sources"][0]["Packages"])

    @staticmethod
    def _safe_filename_stem(path, default_name="MSStoreHelper-IntuneWin"):
        stem = os.path.splitext(os.path.basename(path or ""))[0]
        safe = "".join(ch if ch.isalnum() or ch in "._-" else "-" for ch in stem).strip("._-")
        return safe or default_name

    @staticmethod
    def _queue_package_source_path(package, output_path):
        filename = package["FileName"]
        return os.path.abspath(package.get("LocalPath") or os.path.join(output_path, filename))

    @staticmethod
    def _intune_package_records(packages, output_path, target_arch=SYSTEM_ARCH):
        provisionable = []
        for package in packages:
            if not package.get("FileName") or not is_installable_package(package):
                continue
            provisionable.append(annotate_package(package.copy()))

        provisionable = StoreAPI.order_packages_for_install(provisionable, target_arch)
        records = []
        seen = set()
        for package in provisionable:
            filename = package["FileName"]
            if filename.lower() in seen:
                continue
            seen.add(filename.lower())

            source_path = StoreAPI._queue_package_source_path(package, output_path)
            if not os.path.exists(source_path):
                raise ValueError(f"Downloaded file is missing: {filename}")

            records.append({
                "FileName": filename,
                "SourcePath": source_path,
                "Identity": package.get("PackageIdentity") or package_identity(filename),
                "Version": package.get("AvailableVersion") or format_version_tuple(package_version_tuple(filename)),
                "Role": package.get("PackageRoleLabel") or package_role_label(filename),
            })

        if not records:
            raise ValueError("No downloaded AppX/MSIX packages are available for Intune packaging")
        return records

    @staticmethod
    def _generate_intune_install_script(records):
        lines = [
            f"# Generated by {APP_NAME} v{APP_VERSION}",
            "$ErrorActionPreference = 'Stop'",
            "$PackageRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Packages'",
            "",
            "$packages = @(",
        ]
        for record in records:
            lines.append(
                "    [pscustomobject]@{ "
                f"FileName = {StoreAPI._powershell_literal(record['FileName'])}; "
                f"Role = {StoreAPI._powershell_literal(record['Role'])} "
                "}"
            )
        lines.extend([
            ")",
            "",
            "foreach ($package in $packages) {",
            "    $packagePath = Join-Path -Path $PackageRoot -ChildPath $package.FileName",
            "    if (-not (Test-Path -LiteralPath $packagePath)) {",
            "        throw \"Package not found: $packagePath\"",
            "    }",
            "",
            "    Write-Host (\"Provisioning {0} [{1}]\" -f $package.FileName, $package.Role)",
            "    $arguments = @(",
            "        '/Online',",
            "        '/Add-ProvisionedAppxPackage',",
            "        \"/PackagePath:$packagePath\",",
            "        '/SkipLicense'",
            "    )",
            "    & dism.exe @arguments",
            "    if ($LASTEXITCODE -ne 0) {",
            "        throw \"DISM failed with exit code $LASTEXITCODE for $($package.FileName)\"",
            "    }",
            "}",
            "",
        ])
        return "\n".join(lines)

    @staticmethod
    def _generate_intune_detection_script(records):
        lines = [
            f"# Generated by {APP_NAME} v{APP_VERSION}",
            "$ErrorActionPreference = 'SilentlyContinue'",
            "",
            "function Test-VersionAtLeast {",
            "    param([string]$Installed, [string]$Required)",
            "    if ([string]::IsNullOrWhiteSpace($Required) -or $Required -eq 'unknown') { return $true }",
            "    try { return ([version]$Installed -ge [version]$Required) } catch { return $true }",
            "}",
            "",
            "$required = @(",
        ]
        for record in records:
            lines.append(
                "    [pscustomobject]@{ "
                f"Name = {StoreAPI._powershell_literal(record['Identity'])}; "
                f"Version = {StoreAPI._powershell_literal(record['Version'])} "
                "}"
            )
        lines.extend([
            ")",
            "",
            "foreach ($package in $required) {",
            "    $provisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $package.Name } | Sort-Object Version -Descending | Select-Object -First 1",
            "    $installed = Get-AppxPackage -AllUsers -Name $package.Name | Sort-Object Version -Descending | Select-Object -First 1",
            "    if (-not $provisioned -and -not $installed) { exit 1 }",
            "    $version = if ($installed) { [string]$installed.Version } else { [string]$provisioned.Version }",
            "    if (-not (Test-VersionAtLeast $version $package.Version)) { exit 1 }",
            "}",
            "",
            "exit 0",
            "",
        ])
        return "\n".join(lines)

    @staticmethod
    def prepare_intune_package_source(packages, staging_root, output_path, target_arch=SYSTEM_ARCH, package_basename="MSStoreHelper-IntuneWin"):
        records = StoreAPI._intune_package_records(packages, output_path, target_arch)
        safe_basename = StoreAPI._safe_filename_stem(package_basename, "MSStoreHelper-IntuneWin")
        staging_root = os.path.abspath(staging_root)
        source_dir = os.path.join(staging_root, safe_basename)
        packages_dir = os.path.join(source_dir, "Packages")

        if os.path.exists(source_dir):
            if os.path.commonpath([staging_root, os.path.abspath(source_dir)]) != staging_root:
                raise ValueError("Invalid Intune staging path")
            shutil.rmtree(source_dir)
        os.makedirs(packages_dir, exist_ok=True)

        for record in records:
            shutil.copy2(record["SourcePath"], os.path.join(packages_dir, record["FileName"]))

        install_script = f"{safe_basename}.ps1"
        setup_file = f"{safe_basename}.cmd"
        detection_script = f"{safe_basename}-Detection.ps1"
        guide_file = f"{safe_basename}-Intune-Commands.txt"

        with open(os.path.join(source_dir, install_script), "w", encoding="utf-8", newline="\r\n") as handle:
            handle.write(StoreAPI._generate_intune_install_script(records))

        with open(os.path.join(source_dir, setup_file), "w", encoding="utf-8", newline="\r\n") as handle:
            handle.write("@echo off\n")
            handle.write(f'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0{install_script}"\n')
            handle.write("exit /b %ERRORLEVEL%\n")

        with open(os.path.join(source_dir, detection_script), "w", encoding="utf-8", newline="\r\n") as handle:
            handle.write(StoreAPI._generate_intune_detection_script(records))

        with open(os.path.join(source_dir, guide_file), "w", encoding="utf-8", newline="\r\n") as handle:
            handle.write(f"Install command: {setup_file}\n")
            handle.write(f"Detection script: {detection_script}\n")
            handle.write("Install behavior: System\n")
            handle.write("Run script as 64-bit process: Yes\n")

        return {
            "SourceDir": source_dir,
            "PackagesDir": packages_dir,
            "SetupFile": setup_file,
            "SetupPath": os.path.join(source_dir, setup_file),
            "DetectionScript": detection_script,
            "DetectionPath": os.path.join(source_dir, detection_script),
            "GuidePath": os.path.join(source_dir, guide_file),
            "PackageCount": len(records),
            "ExpectedIntuneWin": f"{safe_basename}.intunewin",
        }

    @staticmethod
    def find_intunewinapputil():
        candidates = [
            os.environ.get("INTUNEWINAPPUTIL"),
            shutil.which("IntuneWinAppUtil.exe"),
            os.path.join(DEFAULT_OUTPUT, "IntuneWinAppUtil.exe"),
        ]
        for candidate in candidates:
            if candidate and os.path.exists(candidate):
                return os.path.abspath(candidate)
        return None

    @staticmethod
    def build_intunewinapputil_command(tool_path, source_dir, setup_file, output_dir):
        return [
            tool_path,
            "-c", source_dir,
            "-s", os.path.basename(setup_file),
            "-o", output_dir,
            "-q",
        ]

    @staticmethod
    def create_intunewin_package(packages, output_path, intunewin_path, tool_path, target_arch=SYSTEM_ARCH):
        if not tool_path or not os.path.exists(tool_path):
            raise FileNotFoundError("IntuneWinAppUtil.exe was not found")

        if not intunewin_path.lower().endswith(".intunewin"):
            intunewin_path += ".intunewin"

        output_dir = os.path.dirname(os.path.abspath(intunewin_path))
        os.makedirs(output_dir, exist_ok=True)
        package_basename = StoreAPI._safe_filename_stem(intunewin_path)

        with tempfile.TemporaryDirectory(prefix="MSStoreHelper-IntuneWin-") as staging_root:
            source_info = StoreAPI.prepare_intune_package_source(
                packages,
                staging_root,
                output_path,
                target_arch,
                package_basename,
            )
            command = StoreAPI.build_intunewinapputil_command(
                tool_path,
                source_info["SourceDir"],
                source_info["SetupFile"],
                output_dir,
            )
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            if result.returncode != 0:
                error = (result.stderr or result.stdout or "IntuneWinAppUtil failed").strip()
                raise RuntimeError(error)

            generated = os.path.join(output_dir, source_info["ExpectedIntuneWin"])
            if not os.path.exists(generated):
                raise RuntimeError(f"IntuneWinAppUtil did not produce {generated}")

            detection_sidecar = os.path.join(output_dir, f"{package_basename}-Detection.ps1")
            shutil.copy2(source_info["DetectionPath"], detection_sidecar)

        return generated, detection_sidecar, source_info["PackageCount"]

    @staticmethod
    def install_package(filepath):
        try:
            cmd = f'Add-AppxPackage -Path "{filepath}" -ErrorAction Stop 2>&1'
            result = subprocess.run(
                [POWERSHELL_EXE, "-NoProfile", "-Command", cmd],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            if result.returncode != 0:
                error_msg = result.stderr.strip() or result.stdout.strip() or "Unknown error"
                return False, error_msg
            
            return True, "Installed successfully"
        except Exception as e:
            return False, str(e)

    @staticmethod
    def verify_package_signature(filepath):
        try:
            safe_path = filepath.replace("'", "''")
            safe_module = POWERSHELL_SECURITY_MODULE.replace("'", "''")
            cmd = f"""
Import-Module '{safe_module}' -ErrorAction Stop
$path = '{safe_path}'
$sig = Get-AuthenticodeSignature -FilePath $path
$chainOk = $false
$rootSubject = ''
$rootThumbprint = ''
if ($sig.SignerCertificate) {{
    $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
    $chain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
    $chainOk = $chain.Build($sig.SignerCertificate)
    if ($chain.ChainElements.Count -gt 0) {{
        $root = $chain.ChainElements[$chain.ChainElements.Count - 1].Certificate
        $rootSubject = $root.Subject
        $rootThumbprint = $root.Thumbprint
    }}
}}
[pscustomobject]@{{
    Status = "$($sig.Status)"
    StatusMessage = "$($sig.StatusMessage)"
    Signer = if ($sig.SignerCertificate) {{ $sig.SignerCertificate.Subject }} else {{ '' }}
    Root = $rootSubject
    RootThumbprint = $rootThumbprint
    ChainValid = $chainOk
}} | ConvertTo-Json -Compress
"""
            result = subprocess.run(
                [POWERSHELL_EXE, "-NoProfile", "-Command", cmd],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if result.returncode != 0:
                error_msg = result.stderr.strip() or result.stdout.strip() or "Signature check failed"
                return False, error_msg

            signature_info = json.loads(result.stdout.strip() or "{}")
            if signature_info_is_valid_microsoft(signature_info):
                root = signature_info.get("Root", "Microsoft root")
                return True, f"{signature_info.get('Status', 'Valid')} signature via {root}"

            status = signature_info.get("Status", "Unknown")
            signer = signature_info.get("Signer", "unknown signer") or "unknown signer"
            root = signature_info.get("Root", "unknown root") or "unknown root"
            return False, f"{status} signature from {signer} via {root}"
        except Exception as e:
            return False, str(e)

    @staticmethod
    def get_installed_package_version(package_name):
        try:
            safe_name = package_name.replace("'", "''")
            cmd = (
                f"Get-AppxPackage -Name '{safe_name}' | "
                "Sort-Object -Property Version -Descending | "
                "Select-Object -First 1 -ExpandProperty Version"
            )
            result = subprocess.run(
                [POWERSHELL_EXE, "-NoProfile", "-Command", cmd],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if result.returncode != 0:
                return None
            lines = [line.strip() for line in result.stdout.splitlines() if line.strip()]
            return lines[-1] if lines else None
        except Exception:
            return None

    @staticmethod
    def get_installed_appx_identities():
        cmd = (
            "$installed = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name); "
            "$provisioned = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayName); "
            "@($installed + $provisioned | Where-Object { $_ } | Sort-Object -Unique) | ConvertTo-Json -Compress"
        )
        try:
            result = subprocess.run(
                [POWERSHELL_EXE, "-NoProfile", "-Command", cmd],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if result.returncode != 0 or not result.stdout.strip():
                return set()

            payload = json.loads(result.stdout)
            if isinstance(payload, str):
                payload = [payload]
            return {str(identity).lower() for identity in payload if identity}
        except Exception:
            return set()

    @staticmethod
    def detect_missing_ltsc_components(installed_identities=None):
        installed = installed_identities if installed_identities is not None else StoreAPI.get_installed_appx_identities()
        installed = {str(identity).lower() for identity in installed}
        catalog = catalog_apps_by_name()
        missing = []

        for requirement in LTSC_COMPONENT_REQUIREMENTS:
            identities = [identity.lower() for identity in requirement["Identities"]]
            if any(identity in installed for identity in identities):
                continue

            app = catalog.get(requirement["Name"])
            if not app:
                continue
            missing_app = app.copy()
            missing_app["MissingIdentities"] = requirement["Identities"]
            missing.append(missing_app)

        return missing

    @staticmethod
    def should_skip_installed_package(package):
        package_name = package.get("PackageIdentity") or package_identity(package["FileName"])
        installed_version = StoreAPI.get_installed_package_version(package_name)
        if not installed_version:
            return False, installed_version, package_name
        return installed_version_satisfies_package(package, installed_version), installed_version, package_name

    @staticmethod
    def is_noop_install_error(error_msg):
        error_lower = error_msg.lower()
        return "0x80073d06" in error_lower or "higher version" in error_lower
    
    @staticmethod
    def get_store_repair_steps():
        return [
            ("🔧 Starting Windows Update...", 'Start-Service -Name wuauserv -ErrorAction SilentlyContinue'),
            ("🔧 Starting BITS...", 'Start-Service -Name bits -ErrorAction SilentlyContinue'),
            ("🔐 Starting licensing services...", 'Start-Service -Name ClipSVC -ErrorAction SilentlyContinue; Start-Service -Name LicenseManager -ErrorAction SilentlyContinue'),
            ("🧹 Closing Store broker processes...", 'Get-Process WinStore.App,MicrosoftStore,RuntimeBroker -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue'),
            ("🧹 Resetting Store cache...", 'Start-Process wsreset.exe -WindowStyle Hidden -Wait'),
            ("🧹 Rebuilding Store token cache...", r'$paths = @("$env:LOCALAPPDATA\Microsoft\TokenBroker\Cache\*", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache\*", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\AC\TokenBroker\Cache\*", "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Cache\*"); foreach ($path in $paths) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }'),
            ("🔄 Resetting Store package state...", 'if (Get-Command Reset-AppxPackage -ErrorAction SilentlyContinue) { Get-AppxPackage Microsoft.WindowsStore -ErrorAction SilentlyContinue | Reset-AppxPackage -ErrorAction SilentlyContinue; Get-AppxPackage Microsoft.StorePurchaseApp -ErrorAction SilentlyContinue | Reset-AppxPackage -ErrorAction SilentlyContinue }'),
            ("🔄 Re-registering Store packages...", r'@("Microsoft.WindowsStore", "Microsoft.StorePurchaseApp") | ForEach-Object { Get-AppxPackage -AllUsers $_ -ErrorAction SilentlyContinue | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue } }'),
            ("🔐 Re-syncing Store licensing...", r'Start-Service -Name ClipSVC -ErrorAction SilentlyContinue; Start-Service -Name LicenseManager -ErrorAction SilentlyContinue; Get-AppxPackage Microsoft.StorePurchaseApp -ErrorAction SilentlyContinue | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue }'),
            ("🌐 Resetting network...", 'netsh winsock reset 2>$null'),
            ("🌐 Flushing DNS...", 'ipconfig /flushdns 2>$null'),
        ]

    @staticmethod
    def get_provisioning_repair_steps():
        return [
            ("🔧 Starting AppX services...", 'Start-Service -Name AppXSVC -ErrorAction SilentlyContinue; Start-Service -Name ClipSVC -ErrorAction SilentlyContinue'),
            ("🧹 Clearing Store deprovision tombstones...", r'$root = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Deprovisioned"; $patterns = @("*Microsoft.WindowsStore*", "*Microsoft.StorePurchaseApp*", "*Microsoft.DesktopAppInstaller*"); foreach ($pattern in $patterns) { Get-ChildItem $root -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like $pattern } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue }'),
            ("🔄 Re-registering Store apps for existing users...", r'@("Microsoft.WindowsStore", "Microsoft.StorePurchaseApp", "Microsoft.DesktopAppInstaller") | ForEach-Object { Get-AppxPackage -AllUsers $_ -ErrorAction SilentlyContinue | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue } }'),
            ("📋 Checking provisioned Store catalog...", 'Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -in @("Microsoft.WindowsStore", "Microsoft.StorePurchaseApp", "Microsoft.DesktopAppInstaller") } | Out-Null'),
        ]

    @staticmethod
    def get_licensing_reset_steps():
        return [
            ("🔐 Stopping licensing services...", 'Stop-Service -Name LicenseManager -Force -ErrorAction SilentlyContinue; Stop-Service -Name ClipSVC -Force -ErrorAction SilentlyContinue'),
            ("🧹 Clearing ClipSVC license cache...", r'$paths = @("$env:ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\*", "$env:ProgramData\Microsoft\Windows\ClipSVC\Tokens\*"); foreach ($path in $paths) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }'),
            ("🔐 Starting licensing services...", 'Start-Service -Name ClipSVC -ErrorAction SilentlyContinue; Start-Service -Name LicenseManager -ErrorAction SilentlyContinue'),
            ("🔄 Re-registering Store licensing app...", r'@("Microsoft.StorePurchaseApp", "Microsoft.WindowsStore") | ForEach-Object { Get-AppxPackage -AllUsers $_ -ErrorAction SilentlyContinue | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue } }'),
        ]

    @staticmethod
    def get_cache_rebuild_steps():
        return [
            ("🧹 Closing Store cache owners...", 'Get-Process WinStore.App,MicrosoftStore,RuntimeBroker -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue'),
            ("🔎 Scanning Store cache folders...", r'$paths = @("$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\AC\INetCache", "$env:LOCALAPPDATA\Packages\Microsoft.StorePurchaseApp_8wekyb3d8bbwe\LocalCache"); foreach ($path in $paths) { if (Test-Path $path) { Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Length -eq 0 } | Measure-Object | Out-Null } }'),
            ("📦 Backing up existing Store caches...", r'$stamp = Get-Date -Format "yyyyMMdd-HHmmss"; $paths = @("$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\AC\INetCache", "$env:LOCALAPPDATA\Packages\Microsoft.StorePurchaseApp_8wekyb3d8bbwe\LocalCache"); foreach ($path in $paths) { if (Test-Path $path) { Move-Item -LiteralPath $path -Destination "$path.bak-$stamp" -Force -ErrorAction SilentlyContinue } }'),
            ("🔄 Rebuilding clean Store cache folders...", r'$paths = @("$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\AC\INetCache", "$env:LOCALAPPDATA\Packages\Microsoft.StorePurchaseApp_8wekyb3d8bbwe\LocalCache"); foreach ($path in $paths) { New-Item -ItemType Directory -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }'),
            ("🧹 Running wsreset after offline rebuild...", 'Start-Process wsreset.exe -WindowStyle Hidden -Wait'),
        ]

    @staticmethod
    def _run_powershell_steps(steps, log_callback=None, progress_callback=None, timeout=90):
        results = []
        for i, (desc, cmd) in enumerate(steps):
            if log_callback:
                log_callback(desc)
            try:
                result = subprocess.run(
                    [POWERSHELL_EXE, "-NoProfile", "-Command", cmd],
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    timeout=90
                )
                results.append((desc, result.returncode == 0))
            except:
                results.append((desc, False))
            
            if progress_callback:
                progress_callback((i + 1) / len(steps))
        
        return results

    @staticmethod
    def run_repair(log_callback=None, progress_callback=None):
        return StoreAPI._run_powershell_steps(StoreAPI.get_store_repair_steps(), log_callback, progress_callback)

    @staticmethod
    def run_provisioning_repair(log_callback=None, progress_callback=None):
        return StoreAPI._run_powershell_steps(StoreAPI.get_provisioning_repair_steps(), log_callback, progress_callback)

    @staticmethod
    def run_licensing_reset(log_callback=None, progress_callback=None):
        return StoreAPI._run_powershell_steps(StoreAPI.get_licensing_reset_steps(), log_callback, progress_callback)

    @staticmethod
    def run_cache_rebuild(log_callback=None, progress_callback=None):
        return StoreAPI._run_powershell_steps(StoreAPI.get_cache_rebuild_steps(), log_callback, progress_callback)

# ==================== UI COMPONENTS ====================

class ModernCard(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color=Theme.BG_CARD, corner_radius=12, border_width=1, border_color=Theme.BORDER, **kwargs)


class AppTile(ctk.CTkFrame):
    def __init__(self, master, app_data, on_select):
        super().__init__(master, fg_color="transparent")
        
        self.app_data = app_data
        self.on_select = on_select
        self.selected = ctk.BooleanVar(value=False)
        
        self.container = ctk.CTkFrame(self, fg_color=Theme.BG_CARD, corner_radius=10, border_width=1, border_color=Theme.BORDER)
        self.container.pack(fill="x", pady=4, padx=2)
        self.container.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.container, text=app_data.get("Icon", "📦"), font=("Segoe UI Emoji", 24), width=50).grid(row=0, column=0, rowspan=2, padx=(15, 10), pady=12)
        ctk.CTkLabel(self.container, text=app_data["Name"], font=("Segoe UI Semibold", 14), anchor="w").grid(row=0, column=1, sticky="sw", padx=5, pady=(12, 0))
        ctk.CTkLabel(self.container, text=app_data.get("Description", ""), font=("Segoe UI", 11), text_color=Theme.TEXT_SECONDARY, anchor="w").grid(row=1, column=1, sticky="nw", padx=5, pady=(0, 12))
        
        self.chk = ctk.CTkCheckBox(self.container, text="", variable=self.selected, width=24, command=self._toggle, fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER)
        self.chk.grid(row=0, column=2, rowspan=2, padx=15)
        
        self.container.bind("<Enter>", lambda e: self.container.configure(fg_color=Theme.BG_CARD_HOVER))
        self.container.bind("<Leave>", lambda e: self.container.configure(fg_color=Theme.BG_CARD))
    
    def _toggle(self):
        self.on_select(self.app_data, self.selected.get())


class SearchResultTile(ctk.CTkFrame):
    def __init__(self, master, app_data, on_fetch, on_select):
        super().__init__(master, fg_color="transparent")
        
        self.app_data = app_data
        self.selected = ctk.BooleanVar(value=False)
        
        self.container = ctk.CTkFrame(self, fg_color=Theme.BG_CARD, corner_radius=10, border_width=1, border_color=Theme.BORDER)
        self.container.pack(fill="x", pady=4, padx=2)
        self.container.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.container, text="📱", font=("Segoe UI Emoji", 20), width=45).grid(row=0, column=0, rowspan=2, padx=(15, 10), pady=10)
        ctk.CTkLabel(self.container, text=app_data.get("Name", "Unknown"), font=("Segoe UI Semibold", 13), anchor="w").grid(row=0, column=1, sticky="sw", padx=5, pady=(10, 0))
        
        info = f"{app_data.get('Publisher', '')}  •  {app_data.get('ProductId', '')}"
        ctk.CTkLabel(self.container, text=info, font=("Consolas", 10), text_color=Theme.TEXT_MUTED, anchor="w").grid(row=1, column=1, sticky="nw", padx=5, pady=(0, 10))
        
        btn_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        btn_frame.grid(row=0, column=2, rowspan=2, padx=10)
        
        ctk.CTkButton(btn_frame, text="Get Files", width=80, height=30, font=("Segoe UI", 12), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=lambda: on_fetch(app_data)).pack(side="left", padx=3)
        ctk.CTkCheckBox(btn_frame, text="", variable=self.selected, width=24, command=lambda: on_select(app_data, self.selected.get()), fg_color=Theme.PRIMARY).pack(side="left", padx=5)
        
        self.container.bind("<Enter>", lambda e: self.container.configure(fg_color=Theme.BG_CARD_HOVER))
        self.container.bind("<Leave>", lambda e: self.container.configure(fg_color=Theme.BG_CARD))


class PackageRow(ctk.CTkFrame):
    def __init__(self, master, pkg_data, on_toggle, index, target_arch):
        super().__init__(master, fg_color=Theme.BG_CARD if index % 2 == 0 else "transparent", corner_radius=6)
        
        self.pkg_data = pkg_data
        self.selected = ctk.BooleanVar(value=False)
        
        self.grid_columnconfigure(1, weight=1)
        
        self.chk = ctk.CTkCheckBox(self, text="", variable=self.selected, width=24, command=lambda: on_toggle(pkg_data, self.selected.get()), fg_color=Theme.PRIMARY)
        self.chk.grid(row=0, column=0, padx=(12, 8), pady=10)
        
        info_frame = ctk.CTkFrame(self, fg_color="transparent")
        info_frame.grid(row=0, column=1, sticky="ew", padx=5, pady=8)
        info_frame.grid_columnconfigure(0, weight=1)
        
        name_color = Theme.ENCRYPTED_COLOR if pkg_data.get('IsEncrypted') else Theme.TEXT_PRIMARY
        ctk.CTkLabel(info_frame, text=pkg_data['FileName'], font=("Consolas", 11), text_color=name_color, anchor="w", wraplength=500).grid(row=0, column=0, sticky="w")
        
        tags_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
        tags_frame.grid(row=1, column=0, sticky="w", pady=(4, 0))
        
        ftype = pkg_data.get('FileType', '?')
        type_color = Theme.BUNDLE_COLOR if 'BUNDLE' in ftype else Theme.PRIMARY
        if pkg_data.get('IsEncrypted'):
            type_color = Theme.ENCRYPTED_COLOR
        
        ctk.CTkLabel(tags_frame, text=f" {ftype} ", font=("Consolas", 10), fg_color=type_color, corner_radius=4, text_color="#000000" if type_color == Theme.BUNDLE_COLOR else Theme.TEXT_PRIMARY).pack(side="left", padx=(0, 6))
        
        arch = pkg_data.get('Architecture', 'neutral')
        arch_color = Theme.ARCH_MATCH if arch in [target_arch, 'neutral'] else Theme.TEXT_MUTED
        ctk.CTkLabel(tags_frame, text=f" {arch} ", font=("Consolas", 10), text_color=arch_color).pack(side="left", padx=(0, 6))

        if is_dependency_package(pkg_data):
            role_label = pkg_data.get('PackageRoleLabel') or package_role_label(pkg_data['FileName'])
            ctk.CTkLabel(tags_frame, text=role_label, font=("Segoe UI", 10), text_color=Theme.INFO).pack(side="left", padx=(0, 6))
        
        if pkg_data.get('IsEncrypted'):
            ctk.CTkLabel(tags_frame, text="⚠️ Encrypted", font=("Segoe UI", 10), text_color=Theme.WARNING).pack(side="left")
        
        self.size_lbl = ctk.CTkLabel(self, text=pkg_data.get('SizeStr', '—'), font=("Consolas", 11), text_color=Theme.TEXT_SECONDARY, width=80)
        self.size_lbl.grid(row=0, column=2, padx=(5, 15))
    
    def set_selected(self, value):
        self.selected.set(value)
    
    def update_size(self, size_str):
        self.size_lbl.configure(text=size_str)


class QueueItem(ctk.CTkFrame):
    def __init__(self, master, pkg_info):
        super().__init__(master, fg_color=Theme.BG_CARD, corner_radius=8)
        
        self.pkg_info = pkg_info
        self.grid_columnconfigure(0, weight=1)
        
        fname = pkg_info['FileName']
        display = fname[:35] + "..." if len(fname) > 38 else fname
        
        ctk.CTkLabel(self, text=display, font=("Consolas", 10), anchor="w").grid(row=0, column=0, sticky="w", padx=10, pady=(8, 2))
        
        info_frame = ctk.CTkFrame(self, fg_color="transparent")
        info_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        
        ctk.CTkLabel(info_frame, text=pkg_info.get('SizeStr', '—'), font=("Consolas", 10), text_color=Theme.INFO).pack(side="left")
        
        if is_dependency_package(pkg_info):
            role_label = pkg_info.get('PackageRoleLabel') or package_role_label(pkg_info['FileName'])
            ctk.CTkLabel(info_frame, text=role_label, font=("Segoe UI", 10), text_color=Theme.WARNING).pack(side="left", padx=(8, 0))

        self.status_lbl = ctk.CTkLabel(info_frame, text="Waiting", font=("Segoe UI", 10), text_color=Theme.TEXT_MUTED)
        self.status_lbl.pack(side="right")
        
        pkg_info['_status_widget'] = self.status_lbl


# ==================== MAIN APPLICATION ====================

class MSStoreHelperApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title(f"📦 {APP_NAME} v{APP_VERSION}")
        self.geometry("1280x800")
        self.minsize(1000, 600)
        
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("dark-blue")
        self.configure(fg_color=Theme.BG_DARK)
        
        self.selected_apps = []
        self.download_queue = []
        self.current_packages = []
        self.selected_packages = set()
        self.package_rows = []
        self.current_view = "welcome"
        self.arch_options = [f"Auto ({SYSTEM_ARCH})", "x64", "x86", "arm64", "arm", "neutral"]
        self.arch_override_var = ctk.StringVar(value=self.arch_options[0])
        self.package_scroll = None
        self.output_path = DEFAULT_OUTPUT
        self.shared_cache_enabled = ctk.BooleanVar(value=False)
        self.shared_cache_path = os.path.join(DEFAULT_OUTPUT, "SharedCache")
        
        self._build_ui()
        self._show_welcome()

    def _target_arch(self):
        choice = self.arch_override_var.get()
        if choice.startswith("Auto"):
            return SYSTEM_ARCH
        return choice.lower()

    def _has_arch_override(self):
        return not self.arch_override_var.get().startswith("Auto")
    
    def _build_ui(self):
        # HEADER
        self.header = ctk.CTkFrame(self, fg_color=Theme.BG_CARD, height=70, corner_radius=0)
        self.header.pack(fill="x", side="top")
        self.header.pack_propagate(False)
        
        title_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        title_frame.pack(side="left", padx=20, pady=15)
        
        ctk.CTkLabel(title_frame, text="📦", font=("Segoe UI Emoji", 28)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(title_frame, text=APP_NAME, font=("Segoe UI Semibold", 22), text_color=Theme.TEXT_PRIMARY).pack(side="left")
        ctk.CTkLabel(title_frame, text=f"v{APP_VERSION}", font=("Segoe UI", 12), text_color=Theme.TEXT_MUTED).pack(side="left", padx=(10, 0), pady=(8, 0))
        
        info_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        info_frame.pack(side="right", padx=20)
        
        admin_text = "✅ Admin" if IS_ADMIN else "⚠️ Not Admin"
        admin_color = Theme.SUCCESS if IS_ADMIN else Theme.WARNING
        ctk.CTkLabel(info_frame, text=admin_text, font=("Segoe UI", 12), text_color=admin_color).pack(side="right", padx=15)
        ctk.CTkLabel(info_frame, text=f"System: {SYSTEM_ARCH}", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY).pack(side="right", padx=15)
        
        ctk.CTkButton(info_frame, text="❓ Help", width=80, height=32, fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._show_help).pack(side="right", padx=5)
        
        # LOG PANEL (at bottom - pack BEFORE main so it claims bottom space)
        self.log_panel = ctk.CTkFrame(self, fg_color=Theme.BG_CARD, corner_radius=0)
        self._build_log_panel()
        
        # MAIN (fills remaining space)
        self.main = ctk.CTkFrame(self, fg_color="transparent")
        self.main.pack(fill="both", expand=True, padx=15, pady=15)
        self.main.grid_columnconfigure(1, weight=1)
        self.main.grid_rowconfigure(0, weight=1)
        
        # SIDEBAR
        self.sidebar = ctk.CTkFrame(self.main, fg_color=Theme.BG_CARD, width=280, corner_radius=12)
        self.sidebar.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        self.sidebar.grid_propagate(False)
        self._build_sidebar()
        
        # CONTENT
        self.content = ctk.CTkFrame(self.main, fg_color="transparent")
        self.content.grid(row=0, column=1, sticky="nsew")
        
        # QUEUE PANEL
        self.right_panel = ctk.CTkFrame(self.main, fg_color=Theme.BG_CARD, width=300, corner_radius=12)
        self.right_panel.grid(row=0, column=2, sticky="ns", padx=(10, 0))
        self.right_panel.grid_propagate(False)
        self._build_queue_panel()
    
    def _build_sidebar(self):
        # SEARCH
        search_section = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        search_section.pack(fill="x", padx=15, pady=(20, 10))
        
        ctk.CTkLabel(search_section, text="🔍 Find Apps", font=("Segoe UI Semibold", 16), anchor="w").pack(fill="x")
        ctk.CTkLabel(search_section, text="Search by name to find any app", font=("Segoe UI", 11), text_color=Theme.TEXT_MUTED, anchor="w").pack(fill="x", pady=(2, 8))
        
        self.search_entry = ctk.CTkEntry(search_section, placeholder_text="e.g. Spotify, WhatsApp, VLC...", height=40, font=("Segoe UI", 13), fg_color=Theme.BG_INPUT, border_color=Theme.BORDER)
        self.search_entry.pack(fill="x", pady=(0, 8))
        self.search_entry.bind("<Return>", lambda e: self._do_search())
        
        ctk.CTkButton(search_section, text="🔍 Search Store", height=38, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._do_search).pack(fill="x")
        
        ctk.CTkFrame(self.sidebar, fg_color=Theme.BORDER, height=1).pack(fill="x", padx=15, pady=15)
        
        # QUICK FIX
        fix_section = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        fix_section.pack(fill="x", padx=15, pady=(0, 10))
        
        ctk.CTkLabel(fix_section, text="⚡ Quick Actions", font=("Segoe UI Semibold", 16), anchor="w").pack(fill="x")
        ctk.CTkLabel(fix_section, text="One-click solutions for common needs", font=("Segoe UI", 11), text_color=Theme.TEXT_MUTED, anchor="w").pack(fill="x", pady=(2, 8))
        
        self.quickfix_var = ctk.StringVar(value=list(QUICK_FIX_PRESETS.keys())[0])
        ctk.CTkOptionMenu(fix_section, values=list(QUICK_FIX_PRESETS.keys()), variable=self.quickfix_var, height=36, font=("Segoe UI", 12), fg_color=Theme.BG_INPUT, button_color=Theme.PRIMARY, button_hover_color=Theme.PRIMARY_HOVER, command=self._update_quickfix_desc).pack(fill="x", pady=(0, 6))
        
        self.quickfix_desc = ctk.CTkLabel(fix_section, text=QUICK_FIX_PRESETS[self.quickfix_var.get()]["description"], font=("Segoe UI", 11), text_color=Theme.TEXT_SECONDARY, wraplength=240, justify="left", anchor="w")
        self.quickfix_desc.pack(fill="x", pady=(0, 8))
        
        ctk.CTkButton(fix_section, text="⚡ Apply Quick Fix", height=38, font=("Segoe UI Semibold", 13), fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER, command=self._apply_quickfix).pack(fill="x", pady=(0, 6))
        ctk.CTkButton(fix_section, text="🔎 Scan LTSC Gaps", height=34, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._scan_ltsc_gaps).pack(fill="x")
        
        ctk.CTkFrame(self.sidebar, fg_color=Theme.BORDER, height=1).pack(fill="x", padx=15, pady=15)
        
        # CATEGORIES
        cat_section = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        cat_section.pack(fill="both", expand=True, padx=15)
        
        ctk.CTkLabel(cat_section, text="📂 Browse Categories", font=("Segoe UI Semibold", 16), anchor="w").pack(fill="x")
        
        cat_scroll = ctk.CTkScrollableFrame(cat_section, fg_color="transparent")
        cat_scroll.pack(fill="both", expand=True, pady=(8, 0))
        
        for cat_name in APP_CATALOG.keys():
            ctk.CTkButton(cat_scroll, text=cat_name, height=36, font=("Segoe UI", 12), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, hover_color=Theme.BG_CARD_HOVER, anchor="w", command=lambda c=cat_name: self._show_category(c)).pack(fill="x", pady=2)
        
        # REPAIR
        repair_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        repair_frame.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkButton(repair_frame, text="🔧 Repair Store", height=40, font=("Segoe UI Semibold", 13), fg_color=Theme.DANGER, hover_color=Theme.DANGER_HOVER, command=self._run_repair).pack(fill="x", pady=(0, 6))
        ctk.CTkButton(repair_frame, text="👥 Provision Store", height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._run_provisioning_repair).pack(fill="x", pady=(0, 6))
        ctk.CTkButton(repair_frame, text="🔐 Reset Licensing", height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._run_licensing_reset).pack(fill="x", pady=(0, 6))
        ctk.CTkButton(repair_frame, text="🧹 Rebuild Cache", height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._run_cache_rebuild).pack(fill="x")
        ctk.CTkLabel(repair_frame, text="Fix connectivity and new-profile Store registration", font=("Segoe UI", 10), text_color=Theme.TEXT_MUTED, wraplength=220).pack(pady=(4, 0))
    
    def _build_queue_panel(self):
        header_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        header_frame.pack(fill="x", padx=15, pady=(20, 10))
        
        ctk.CTkLabel(header_frame, text="📥 Download Queue", font=("Segoe UI Semibold", 16)).pack(side="left")
        self.queue_count = ctk.CTkLabel(header_frame, text="0 items", font=("Segoe UI", 12), text_color=Theme.TEXT_MUTED)
        self.queue_count.pack(side="right")
        
        self.queue_scroll = ctk.CTkScrollableFrame(self.right_panel, fg_color=Theme.BG_INPUT, corner_radius=8, height=210)
        self.queue_scroll.pack(fill="x", padx=15, pady=(0, 10))
        
        self.queue_empty = ctk.CTkLabel(self.queue_scroll, text="📭\n\nNo files in queue\n\nSearch for apps or browse\ncategories to get started", font=("Segoe UI", 12), text_color=Theme.TEXT_MUTED, justify="center")
        self.queue_empty.pack(expand=True, pady=40)
        
        btn_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        ctk.CTkButton(btn_frame, text="Clear", width=70, height=32, font=("Segoe UI", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._clear_queue).pack(side="left")

        cache_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        cache_frame.pack(fill="x", padx=15, pady=(0, 10))

        ctk.CTkCheckBox(
            cache_frame,
            text="Shared cache",
            variable=self.shared_cache_enabled,
            font=("Segoe UI", 12),
            fg_color=Theme.PRIMARY,
            hover_color=Theme.PRIMARY_HOVER,
            command=self._update_shared_cache_state,
        ).pack(side="left")
        ctk.CTkButton(
            cache_frame,
            text="Browse",
            width=70,
            height=30,
            font=("Segoe UI", 12),
            fg_color="transparent",
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._choose_shared_cache_folder,
        ).pack(side="right")

        self.shared_cache_label = ctk.CTkLabel(
            self.right_panel,
            text=self._format_shared_cache_path(),
            font=("Consolas", 10),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
        )
        self.shared_cache_label.pack(fill="x", padx=15, pady=(0, 10))
        
        progress_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        progress_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        self.progress_label = ctk.CTkLabel(progress_frame, text="Ready", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY)
        self.progress_label.pack(fill="x")
        
        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=8, corner_radius=4, fg_color=Theme.BG_INPUT, progress_color=Theme.PRIMARY)
        self.progress_bar.pack(fill="x", pady=(6, 0))
        self.progress_bar.set(0)
        
        action_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        action_frame.pack(fill="x", padx=15, pady=(0, 20))
        
        ctk.CTkButton(action_frame, text="⬇️ Download All", height=42, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._start_download).pack(fill="x", pady=(0, 8))
        ctk.CTkButton(action_frame, text="🧾 Export DISM Script", height=38, font=("Segoe UI Semibold", 13), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._export_dism_script).pack(fill="x", pady=(0, 8))
        ctk.CTkButton(action_frame, text="📦 Export IntuneWin", height=38, font=("Segoe UI Semibold", 13), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._export_intunewin_package).pack(fill="x", pady=(0, 8))
        ctk.CTkButton(action_frame, text="📦 Install Downloaded", height=42, font=("Segoe UI Semibold", 13), fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER, command=self._start_install).pack(fill="x")
    
    def _build_log_panel(self):
        """Build the collapsible log/console panel"""
        # Toggle bar (always visible)
        self.log_toggle = ctk.CTkFrame(self.log_panel, fg_color=Theme.BG_INPUT, height=36)
        self.log_toggle.pack(fill="x", side="top")
        self.log_toggle.pack_propagate(False)
        
        toggle_inner = ctk.CTkFrame(self.log_toggle, fg_color="transparent")
        toggle_inner.pack(fill="x", padx=15)
        
        self.log_toggle_btn = ctk.CTkButton(
            toggle_inner,
            text="▼ Console Output",
            font=("Segoe UI Semibold", 12),
            fg_color="transparent",
            hover_color=Theme.BG_CARD_HOVER,
            anchor="w",
            command=self._toggle_log_panel
        )
        self.log_toggle_btn.pack(side="left", pady=5)
        
        self.log_status = ctk.CTkLabel(
            toggle_inner,
            text="",
            font=("Consolas", 10),
            text_color=Theme.TEXT_MUTED
        )
        self.log_status.pack(side="left", padx=15)
        
        # Log controls
        log_controls = ctk.CTkFrame(toggle_inner, fg_color="transparent")
        log_controls.pack(side="right")
        
        ctk.CTkButton(
            log_controls,
            text="📋 Copy",
            width=60,
            height=26,
            font=("Segoe UI", 11),
            fg_color="transparent",
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._copy_log
        ).pack(side="left", padx=3)
        
        ctk.CTkButton(
            log_controls,
            text="🗑️ Clear",
            width=60,
            height=26,
            font=("Segoe UI", 11),
            fg_color="transparent",
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._clear_log
        ).pack(side="left", padx=3)
        
        # Start collapsed so primary queue actions stay visible at default size.
        self.log_content = ctk.CTkFrame(self.log_panel, fg_color=Theme.BG_DARK, height=180)
        self.log_content.pack_propagate(False)
        
        self.log_text = ctk.CTkTextbox(
            self.log_content,
            font=("Consolas", 11),
            fg_color=Theme.BG_DARK,
            text_color=Theme.TEXT_SECONDARY,
            wrap="word",
            state="disabled"
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Pack the log panel at bottom BEFORE main content
        self.log_panel.pack(fill="x", side="bottom")
        self.log_expanded = False
        self.log_toggle_btn.configure(text="▲ Console Output")
        
        # Add initial log message
        self._log("INFO", f"MSStoreHelper v{APP_VERSION} initialized")
        self._log("INFO", f"System Architecture: {SYSTEM_ARCH}")
        self._log("INFO", f"Administrator: {'Yes' if IS_ADMIN else 'No'}")
        self._log("INFO", f"Output Directory: {DEFAULT_OUTPUT}")
    
    def _toggle_log_panel(self):
        """Toggle log panel expanded/collapsed"""
        if self.log_expanded:
            self.log_content.pack_forget()
            self.log_toggle_btn.configure(text="▲ Console Output")
            self.log_expanded = False
        else:
            self.log_content.pack(fill="x", side="top")
            self.log_toggle_btn.configure(text="▼ Console Output")
            self.log_expanded = True
            # Scroll to bottom
            self.log_text.see("end")
    
    def _log(self, level, message):
        """Add a message to the log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Color codes for different levels
        level_colors = {
            "INFO": Theme.INFO,
            "SUCCESS": Theme.SUCCESS,
            "WARNING": Theme.WARNING,
            "ERROR": Theme.DANGER,
            "DEBUG": Theme.TEXT_MUTED
        }
        
        level_icons = {
            "INFO": "ℹ️",
            "SUCCESS": "✅",
            "WARNING": "⚠️",
            "ERROR": "❌",
            "DEBUG": "🔍"
        }
        
        icon = level_icons.get(level, "•")
        formatted = f"[{timestamp}] {icon} {message}\n"
        
        self.log_text.configure(state="normal")
        self.log_text.insert("end", formatted)
        self.log_text.configure(state="disabled")
        self.log_text.see("end")
        
        # Update status in toggle bar
        short_msg = message[:50] + "..." if len(message) > 50 else message
        self.log_status.configure(text=short_msg)
    
    def _copy_log(self):
        """Copy log contents to clipboard"""
        self.log_text.configure(state="normal")
        content = self.log_text.get("1.0", "end-1c")
        self.log_text.configure(state="disabled")
        self.clipboard_clear()
        self.clipboard_append(content)
        self._log("INFO", "Log copied to clipboard")
    
    def _clear_log(self):
        """Clear log contents"""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self._log("INFO", "Log cleared")
    
    def _clear_content(self):
        self.package_scroll = None
        for widget in self.content.winfo_children():
            widget.destroy()
    
    def _show_welcome(self):
        self._clear_content()
        self.current_view = "welcome"
        
        center = ctk.CTkFrame(self.content, fg_color="transparent")
        center.place(relx=0.5, rely=0.5, anchor="center")
        
        card = ModernCard(center)
        card.pack(padx=40, pady=40)
        
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(padx=50, pady=40)
        
        ctk.CTkLabel(inner, text="👋", font=("Segoe UI Emoji", 48)).pack(pady=(0, 10))
        ctk.CTkLabel(inner, text="Welcome to MSStoreHelper", font=("Segoe UI Semibold", 24)).pack()
        ctk.CTkLabel(inner, text="Download and install Microsoft Store apps\nwithout needing access to the Store", font=("Segoe UI", 14), text_color=Theme.TEXT_SECONDARY, justify="center").pack(pady=(10, 25))
        
        options = ctk.CTkFrame(inner, fg_color="transparent")
        options.pack()
        
        ctk.CTkButton(options, text="🔍 Search for an App", width=200, height=45, font=("Segoe UI Semibold", 14), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=lambda: self.search_entry.focus()).pack(side="left", padx=10)
        ctk.CTkButton(options, text="📂 Browse Categories", width=200, height=45, font=("Segoe UI Semibold", 14), fg_color="transparent", border_width=2, border_color=Theme.PRIMARY, hover_color=Theme.BG_CARD_HOVER, command=lambda: self._show_category("🛠️ Essential Repairs")).pack(side="left", padx=10)
        
        tips = ctk.CTkFrame(inner, fg_color=Theme.BG_INPUT, corner_radius=8)
        tips.pack(fill="x", pady=(30, 0))
        
        tip_inner = ctk.CTkFrame(tips, fg_color="transparent")
        tip_inner.pack(padx=20, pady=15)
        
        ctk.CTkLabel(tip_inner, text="💡 Tips", font=("Segoe UI Semibold", 13), anchor="w").pack(fill="x")
        ctk.CTkLabel(tip_inner, text="• Use 'Smart Select' to automatically pick the best files\n• Bundles (.msixbundle) work on all architectures\n• Run as Administrator for installation to work\n• Use 'Repair Store' if the Store shows connectivity errors", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY, justify="left", anchor="w").pack(fill="x", pady=(5, 0))
    
    def _show_category(self, category_name):
        self._clear_content()
        self.current_view = "category"
        self.selected_apps.clear()
        
        cat_data = APP_CATALOG.get(category_name, {})
        apps = cat_data.get("apps", [])
        
        header = ctk.CTkFrame(self.content, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(10, 5))

        title_group = ctk.CTkFrame(header, fg_color="transparent")
        title_group.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(title_group, text=category_name, font=("Segoe UI Semibold", 22), anchor="w").pack(fill="x")
        ctk.CTkLabel(title_group, text=cat_data.get("description", ""), font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(2, 0))
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.pack(side="right")
        ctk.CTkButton(actions, text="Export WinGet", width=120, height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._export_winget_manifest).pack(side="left", padx=(0, 8))
        ctk.CTkButton(actions, text="Get Selected Apps", width=135, height=36, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._fetch_selected).pack(side="left")
        
        scroll = ctk.CTkScrollableFrame(self.content, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        for app in apps:
            AppTile(scroll, app, self._on_app_toggle).pack(fill="x")
    
    def _show_search_results(self, results, query):
        self._clear_content()
        self.current_view = "search"
        self.selected_apps.clear()
        
        header = ctk.CTkFrame(self.content, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(10, 5))

        title_group = ctk.CTkFrame(header, fg_color="transparent")
        title_group.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(title_group, text=f'🔍 Results for "{query}"', font=("Segoe UI Semibold", 20), anchor="w").pack(fill="x")
        ctk.CTkLabel(title_group, text=f"{len(results)} apps found", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(2, 0))
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.pack(side="right")
        ctk.CTkButton(actions, text="Export WinGet", width=120, height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._export_winget_manifest).pack(side="left", padx=(0, 8))
        ctk.CTkButton(actions, text="Get Selected Apps", width=135, height=36, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._fetch_selected).pack(side="left")
        
        if not results:
            empty = ctk.CTkFrame(self.content, fg_color="transparent")
            empty.pack(expand=True)
            ctk.CTkLabel(empty, text="😕", font=("Segoe UI Emoji", 48)).pack(pady=(0, 10))
            ctk.CTkLabel(empty, text="No apps found", font=("Segoe UI Semibold", 18)).pack()
            ctk.CTkLabel(empty, text="Try a different search term", font=("Segoe UI", 13), text_color=Theme.TEXT_SECONDARY).pack(pady=(5, 0))
            return
        
        scroll = ctk.CTkScrollableFrame(self.content, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        for app in results:
            SearchResultTile(scroll, app, self._fetch_single_app, self._on_app_toggle).pack(fill="x")
    
    def _show_packages(self, packages, title):
        self._clear_content()
        self.current_view = "packages"
        self.current_packages = packages
        self.selected_packages.clear()
        self.package_rows.clear()
        
        header = ctk.CTkFrame(self.content, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(10, 5))
        
        ctk.CTkLabel(header, text=f"📦 {title}", font=("Segoe UI Semibold", 20)).pack(side="left")
        self.selection_info = ctk.CTkLabel(header, text="0 selected", font=("Segoe UI", 12), text_color=Theme.INFO)
        self.selection_info.pack(side="right", padx=15)
        
        toolbar = ctk.CTkFrame(self.content, fg_color=Theme.BG_CARD, corner_radius=8)
        toolbar.pack(fill="x", padx=10, pady=(5, 10))
        
        tb_inner = ctk.CTkFrame(toolbar, fg_color="transparent")
        tb_inner.pack(fill="x", padx=10, pady=8)
        
        ctk.CTkButton(tb_inner, text="✨ Smart Select", width=120, height=32, font=("Segoe UI", 12), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._smart_select).pack(side="left", padx=(0, 8))
        ctk.CTkButton(tb_inner, text="Select All", width=90, height=32, font=("Segoe UI", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._select_all).pack(side="left", padx=(0, 8))
        ctk.CTkButton(tb_inner, text="Clear", width=70, height=32, font=("Segoe UI", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._select_none).pack(side="left")
        ctk.CTkLabel(tb_inner, text="Arch", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY).pack(side="left", padx=(14, 6))
        ctk.CTkOptionMenu(tb_inner, values=self.arch_options, variable=self.arch_override_var, width=120, height=32, font=("Segoe UI", 12), fg_color=Theme.BG_INPUT, button_color=Theme.PRIMARY, button_hover_color=Theme.PRIMARY_HOVER, command=self._on_arch_override_change).pack(side="left")
        ctk.CTkButton(tb_inner, text="➕ Add to Queue", width=130, height=32, font=("Segoe UI Semibold", 12), fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER, command=self._add_to_queue).pack(side="right")
        
        col_header = ctk.CTkFrame(self.content, fg_color=Theme.BG_INPUT, corner_radius=6)
        col_header.pack(fill="x", padx=10, pady=(0, 5))
        
        ch_inner = ctk.CTkFrame(col_header, fg_color="transparent")
        ch_inner.pack(fill="x", padx=12, pady=8)
        ch_inner.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(ch_inner, text="", width=40).grid(row=0, column=0)
        ctk.CTkLabel(ch_inner, text="File Name", font=("Segoe UI Semibold", 11), anchor="w").grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(ch_inner, text="Size", font=("Segoe UI Semibold", 11), width=80).grid(row=0, column=2, padx=(0, 10))
        
        self.package_scroll = ctk.CTkScrollableFrame(self.content, fg_color="transparent")
        self.package_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._render_package_rows()
        
        self._fetch_sizes_async()

    def _render_package_rows(self):
        if not self.package_scroll:
            return

        for widget in self.package_scroll.winfo_children():
            widget.destroy()

        self.package_rows.clear()
        target_arch = self._target_arch()
        for i, pkg in enumerate(self.current_packages):
            row = PackageRow(self.package_scroll, pkg, self._on_package_toggle, i, target_arch)
            row.set_selected(pkg['FileName'] in self.selected_packages)
            row.pack(fill="x", pady=1)
            self.package_rows.append(row)

        self._update_selection_info()
    
    def _show_help(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Help")
        dialog.geometry("500x500")
        dialog.transient(self)
        dialog.grab_set()
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 500) // 2
        y = self.winfo_y() + (self.winfo_height() - 500) // 2
        dialog.geometry(f"+{x}+{y}")
        
        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=25, pady=25)
        
        ctk.CTkLabel(content, text="❓ How to Use MSStoreHelper", font=("Segoe UI Semibold", 20)).pack(anchor="w")
        
        help_text = """
🔍 Finding Apps
Search by name (e.g., "Spotify") or browse categories.

📦 Getting Downloads  
Click "Get Files" or select apps and click "Get Selected Apps".

📋 WinGet Export
Select apps and click "Export WinGet" to save an import manifest.

📦 IntuneWin Export
Download queued packages, then export an IntuneWin package with a detection script.

✨ Smart Select
Automatically picks the best files - prefers bundles, skips encrypted files, chooses newest versions.

⬇️ Downloading
Add files to queue and click "Download All". Files save to Downloads folder.

📦 Installing
Click "Install Downloaded" after downloading. Requires Administrator.

🔧 Repair Store
Fixes "needs to be online" and similar errors.

💡 Tips
• Bundles (.msixbundle) are usually all you need
• Avoid .eappx files (encrypted, won't install)
• Get dependencies like VCLibs too"""
        
        text_scroll = ctk.CTkScrollableFrame(content, fg_color="transparent")
        text_scroll.pack(fill="both", expand=True, pady=(15, 0))
        
        ctk.CTkLabel(text_scroll, text=help_text, font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY, justify="left", anchor="w", wraplength=430).pack(fill="x")
    
    def _update_status(self, text, color=None):
        self.progress_label.configure(text=text)
        if color:
            self.progress_label.configure(text_color=color)
        self.update_idletasks()
    
    def _update_progress(self, value):
        self.progress_bar.set(value)
        self.update_idletasks()
    
    def _update_quickfix_desc(self, choice):
        self.quickfix_desc.configure(text=QUICK_FIX_PRESETS.get(choice, {}).get("description", ""))

    def _format_shared_cache_path(self):
        if len(self.shared_cache_path) <= 38:
            return self.shared_cache_path
        return f"{self.shared_cache_path[:16]}...{self.shared_cache_path[-19:]}"

    def _update_shared_cache_state(self):
        state = "enabled" if self.shared_cache_enabled.get() else "disabled"
        self._log("INFO", f"Shared offline cache {state}: {self.shared_cache_path}")

    def _choose_shared_cache_folder(self):
        selected = filedialog.askdirectory(
            title="Select shared offline cache folder",
            initialdir=self.shared_cache_path if os.path.exists(self.shared_cache_path) else DEFAULT_OUTPUT,
        )
        if not selected:
            return

        self.shared_cache_path = selected
        self.shared_cache_enabled.set(True)
        self.shared_cache_label.configure(text=self._format_shared_cache_path())
        self._log("INFO", f"Shared offline cache folder: {self.shared_cache_path}")

    def _on_arch_override_change(self, _choice=None):
        target_arch = self._target_arch()
        mode = "override" if self._has_arch_override() else "auto"
        self._log("INFO", f"Target architecture set to {target_arch} ({mode})")
        self._render_package_rows()
    
    def _on_app_toggle(self, app_data, selected):
        if selected:
            if app_data not in self.selected_apps:
                self.selected_apps.append(app_data)
        else:
            if app_data in self.selected_apps:
                self.selected_apps.remove(app_data)
    
    def _on_package_toggle(self, pkg_data, selected):
        fname = pkg_data['FileName']
        if selected:
            self.selected_packages.add(fname)
        else:
            self.selected_packages.discard(fname)
        self._update_selection_info()
    
    def _update_selection_info(self):
        count = len(self.selected_packages)
        total_size = sum(p.get('SizeBytes', 0) or 0 for p in self.current_packages if p['FileName'] in self.selected_packages)
        self.selection_info.configure(text=f"{count} selected ({format_size(total_size)})")
    
    def _do_search(self):
        query = self.search_entry.get().strip()
        if not query:
            return
        self._update_status("🔍 Searching...", Theme.INFO)
        threading.Thread(target=self._search_worker, args=(query,), daemon=True).start()
    
    def _search_worker(self, query):
        self.after(0, lambda: self._log("INFO", f"Searching Microsoft Store for: {query}"))
        results = StoreAPI.search_store(query)
        if results:
            self.after(0, lambda: self._log("SUCCESS", f"Found {len(results)} apps"))
            for r in results[:5]:
                self.after(0, lambda r=r: self._log("DEBUG", f"  • {r['Name']} ({r['ProductId']})"))
            if len(results) > 5:
                self.after(0, lambda: self._log("DEBUG", f"  ... and {len(results) - 5} more"))
        else:
            self.after(0, lambda: self._log("WARNING", "No apps found for this search"))
        self.after(0, lambda: self._update_status("Ready", Theme.TEXT_SECONDARY))
        self.after(0, lambda: self._show_search_results(results, query))
    
    def _fetch_selected(self):
        if not self.selected_apps:
            self._update_status("⚠️ No apps selected", Theme.WARNING)
            return
        self._update_status("📥 Fetching packages...", Theme.INFO)
        threading.Thread(target=self._fetch_selected_worker, daemon=True).start()

    def _export_winget_manifest(self):
        if not self.selected_apps:
            self._update_status("⚠️ No apps selected", Theme.WARNING)
            self._log("WARNING", "No selected apps to export to WinGet")
            return

        initial_dir = self.output_path if os.path.exists(self.output_path) else DEFAULT_OUTPUT
        os.makedirs(initial_dir, exist_ok=True)
        manifest_path = filedialog.asksaveasfilename(
            title="Save WinGet import manifest",
            initialdir=initial_dir,
            initialfile="MSStoreHelper-WinGetImport.json",
            defaultextension=".json",
            filetypes=[("WinGet import manifest", "*.json"), ("All files", "*.*")],
        )
        if not manifest_path:
            return

        try:
            saved_path, count = StoreAPI.write_winget_import_manifest(self.selected_apps, manifest_path)
        except ValueError as exc:
            self._update_status("⚠️ No WinGet IDs selected", Theme.WARNING)
            self._log("WARNING", str(exc))
        except Exception as exc:
            self._update_status("❌ WinGet export failed", Theme.DANGER)
            self._log("ERROR", f"Failed to export WinGet manifest: {exc}")
        else:
            self._update_status("✅ WinGet manifest exported", Theme.SUCCESS)
            self._log("SUCCESS", f"WinGet import manifest saved: {saved_path} ({count} package(s))")

    def _scan_ltsc_gaps(self):
        target_arch = self._target_arch()
        prefer_exact = self._has_arch_override()
        self._update_status("🔎 Scanning LTSC components...", Theme.INFO)
        threading.Thread(
            target=self._scan_ltsc_gaps_worker,
            args=(target_arch, prefer_exact),
            daemon=True,
        ).start()

    def _scan_ltsc_gaps_worker(self, target_arch, prefer_exact):
        try:
            missing_apps = StoreAPI.detect_missing_ltsc_components()
        except Exception as exc:
            self.after(0, lambda: self._update_status("❌ LTSC scan failed", Theme.DANGER))
            self.after(0, lambda e=str(exc): self._log("ERROR", f"LTSC component scan failed: {e}"))
            return

        if not missing_apps:
            self.after(0, lambda: self._update_status("✅ LTSC components present", Theme.SUCCESS))
            self.after(0, lambda: self._log("SUCCESS", "LTSC component scan found no missing tracked components"))
            return

        missing_names = ", ".join(app["Name"] for app in missing_apps)
        self.after(0, lambda names=missing_names: self._log("INFO", f"Missing LTSC components: {names}"))

        queued_count = 0
        for app in missing_apps:
            self.after(0, lambda n=app["Name"]: self._update_status(f"📥 Queueing {n}...", Theme.INFO))
            packages = StoreAPI.get_packages(app["ProductId"])
            if not packages:
                self.after(0, lambda n=app["Name"]: self._log("WARNING", f"No packages found for missing LTSC component: {n}"))
                continue

            recommended = StoreAPI.smart_select(packages, target_arch, prefer_exact)
            for package in recommended:
                if any(queued["FileName"] == package["FileName"] for queued in self.download_queue):
                    continue
                self.download_queue.append(annotate_package(package.copy()))
                queued_count += 1

        self.download_queue = StoreAPI.order_packages_for_install(self.download_queue, target_arch)
        dependency_count = sum(1 for pkg in self.download_queue if is_dependency_package(pkg))

        if queued_count:
            self.after(0, self._update_queue_ui)
            self.after(0, lambda c=queued_count: self._update_status(f"✅ Queued {c} LTSC package(s)", Theme.SUCCESS))
            self.after(0, lambda c=queued_count, d=dependency_count: self._log("SUCCESS", f"Queued {c} package(s) for missing LTSC components; dependencies in queue: {d}"))
        else:
            self.after(0, lambda: self._update_status("⚠️ No LTSC packages queued", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", "LTSC scan found missing components but no downloadable packages were queued"))
    
    def _fetch_selected_worker(self):
        all_packages = []
        names = []
        for app in self.selected_apps:
            names.append(app['Name'])
            self.after(0, lambda n=app['Name']: self._update_status(f"📥 Fetching {n}...", Theme.INFO))
            self.after(0, lambda n=app['Name'], pid=app['ProductId']: self._log("INFO", f"Fetching packages for: {n} ({pid})"))
            packages = StoreAPI.get_packages(app['ProductId'])
            if packages:
                self.after(0, lambda n=app['Name'], c=len(packages): self._log("SUCCESS", f"  Found {c} packages for {n}"))
            else:
                self.after(0, lambda n=app['Name']: self._log("WARNING", f"  No packages found for {n}"))
            all_packages.extend(packages)
        
        if not all_packages:
            self.after(0, lambda: self._update_status("⚠️ No packages found", Theme.WARNING))
            self.after(0, lambda: self._log("ERROR", "No downloadable packages found for any selected app"))
            return
        
        self.after(0, lambda c=len(all_packages): self._log("INFO", f"Total packages available: {c}"))
        title = ", ".join(names[:2]) + ("..." if len(names) > 2 else "")
        self.after(0, lambda: self._update_status("Ready", Theme.TEXT_SECONDARY))
        self.after(0, lambda: self._show_packages(all_packages, title))
    
    def _fetch_single_app(self, app_data):
        self._update_status(f"📥 Fetching {app_data['Name']}...", Theme.INFO)
        threading.Thread(target=self._fetch_single_worker, args=(app_data,), daemon=True).start()
    
    def _fetch_single_worker(self, app_data):
        self.after(0, lambda: self._log("INFO", f"Fetching packages for: {app_data['Name']} ({app_data['ProductId']})"))
        packages = StoreAPI.get_packages(app_data['ProductId'])
        if not packages:
            self.after(0, lambda: self._update_status("⚠️ No packages found", Theme.WARNING))
            self.after(0, lambda: self._log("ERROR", f"No packages found for {app_data['Name']}"))
            return
        
        self.after(0, lambda: self._log("SUCCESS", f"Found {len(packages)} packages"))
        
        # Log package details
        bundles = [p for p in packages if p.get('IsBundle')]
        encrypted = [p for p in packages if p.get('IsEncrypted')]
        self.after(0, lambda: self._log("DEBUG", f"  Bundles: {len(bundles)}, Encrypted: {len(encrypted)}, Single-arch: {len(packages) - len(bundles)}"))
        
        self.after(0, lambda: self._update_status("Ready", Theme.TEXT_SECONDARY))
        self.after(0, lambda: self._show_packages(packages, app_data['Name']))
    
    def _apply_quickfix(self):
        preset = QUICK_FIX_PRESETS.get(self.quickfix_var.get(), {})
        app_names = preset.get("apps", [])
        
        all_apps = []
        for cat in APP_CATALOG.values():
            all_apps.extend(cat.get("apps", []))
        
        self.selected_apps = [a for a in all_apps if a['Name'] in app_names]
        if self.selected_apps:
            self._fetch_selected()
    
    def _fetch_sizes_async(self):
        def worker():
            for i, pkg in enumerate(self.current_packages):
                size = StoreAPI.get_file_size(pkg['Url'])
                pkg['SizeBytes'] = size
                pkg['SizeStr'] = format_size(size)
                if i < len(self.package_rows):
                    self.after(0, lambda idx=i, s=pkg['SizeStr']: self.package_rows[idx].update_size(s))
            self.after(0, self._update_selection_info)
        threading.Thread(target=worker, daemon=True).start()
    
    def _smart_select(self):
        self._log("INFO", f"Running Smart Select on {len(self.current_packages)} packages...")
        target_arch = self._target_arch()
        self._log("DEBUG", f"  Target architecture: {target_arch}")
        
        best = StoreAPI.smart_select(self.current_packages, target_arch, self._has_arch_override())
        best_names = {p['FileName'] for p in best}
        self.selected_packages = best_names
        for row in self.package_rows:
            row.set_selected(row.pkg_data['FileName'] in best_names)
        self._update_selection_info()
        self._update_status(f"✨ Selected {len(best)} recommended files", Theme.SUCCESS)
        
        self._log("SUCCESS", f"Smart Select chose {len(best)} packages:")
        for p in best:
            ftype = "Bundle" if p.get('IsBundle') else p.get('Architecture', 'neutral')
            self._log("DEBUG", f"  • {p['FileName'][:60]}... ({ftype})")
    
    def _select_all(self):
        self.selected_packages = {p['FileName'] for p in self.current_packages}
        for row in self.package_rows:
            row.set_selected(True)
        self._update_selection_info()
    
    def _select_none(self):
        self.selected_packages.clear()
        for row in self.package_rows:
            row.set_selected(False)
        self._update_selection_info()
    
    def _add_to_queue(self):
        if not self.selected_packages:
            self._update_status("⚠️ No files selected", Theme.WARNING)
            self._log("WARNING", "No files selected to add to queue")
            return
        
        count = 0
        for pkg in self.current_packages:
            if pkg['FileName'] in self.selected_packages:
                if not any(q['FileName'] == pkg['FileName'] for q in self.download_queue):
                    self.download_queue.append(annotate_package(pkg.copy()))
                    count += 1

        self.download_queue = StoreAPI.order_packages_for_install(self.download_queue, self._target_arch())
        dependency_count = sum(1 for pkg in self.download_queue if is_dependency_package(pkg))

        self._update_queue_ui()
        self._update_status(f"✅ Added {count} files to queue", Theme.SUCCESS)
        self._log("INFO", f"Added {count} files to download queue (total: {len(self.download_queue)})")
        if dependency_count:
            self._log("INFO", f"Install order resolved: {dependency_count} dependency package(s) before apps")
    
    def _update_queue_ui(self):
        for widget in self.queue_scroll.winfo_children():
            widget.destroy()
        
        if not self.download_queue:
            ctk.CTkLabel(self.queue_scroll, text="📭\n\nNo files in queue", font=("Segoe UI", 12), text_color=Theme.TEXT_MUTED, justify="center").pack(expand=True, pady=40)
        else:
            for pkg in self.download_queue:
                QueueItem(self.queue_scroll, pkg).pack(fill="x", pady=3, padx=5)
        
        self.queue_count.configure(text=f"{len(self.download_queue)} items")
    
    def _clear_queue(self):
        self.download_queue.clear()
        self._update_queue_ui()
        self._update_status("Queue cleared", Theme.TEXT_SECONDARY)
    
    def _start_download(self):
        if not self.download_queue:
            self._update_status("⚠️ Queue is empty", Theme.WARNING)
            return
        threading.Thread(target=self._download_worker, daemon=True).start()

    def _export_dism_script(self):
        if not self.download_queue:
            self._update_status("⚠️ Queue is empty", Theme.WARNING)
            self._log("WARNING", "No queued files to export")
            return

        initial_dir = self.output_path if os.path.exists(self.output_path) else DEFAULT_OUTPUT
        os.makedirs(initial_dir, exist_ok=True)
        script_path = filedialog.asksaveasfilename(
            title="Save DISM provisioning script",
            initialdir=initial_dir,
            initialfile="MSStoreHelper-ProvisionQueue.ps1",
            defaultextension=".ps1",
            filetypes=[("PowerShell script", "*.ps1"), ("All files", "*.*")],
        )
        if not script_path:
            return

        try:
            StoreAPI.write_dism_provision_script(
                self.download_queue,
                self.output_path,
                script_path,
                self._target_arch(),
            )
        except ValueError as exc:
            self._update_status("⚠️ No AppX/MSIX files in queue", Theme.WARNING)
            self._log("WARNING", str(exc))
        except Exception as exc:
            self._update_status("❌ DISM export failed", Theme.DANGER)
            self._log("ERROR", f"Failed to export DISM script: {exc}")
        else:
            self._update_status("✅ DISM script exported", Theme.SUCCESS)
            self._log("SUCCESS", f"DISM provisioning script saved: {script_path}")

    def _export_intunewin_package(self):
        if not self.download_queue:
            self._update_status("⚠️ Queue is empty", Theme.WARNING)
            self._log("WARNING", "No queued files to package for Intune")
            return

        tool_path = StoreAPI.find_intunewinapputil()
        if not tool_path:
            tool_path = filedialog.askopenfilename(
                title="Select IntuneWinAppUtil.exe",
                filetypes=[("IntuneWinAppUtil", "IntuneWinAppUtil.exe"), ("Executable", "*.exe"), ("All files", "*.*")],
            )
            if not tool_path:
                self._update_status("⚠️ IntuneWinAppUtil required", Theme.WARNING)
                self._log("WARNING", "IntuneWin export requires Microsoft's IntuneWinAppUtil.exe")
                return

        initial_dir = self.output_path if os.path.exists(self.output_path) else DEFAULT_OUTPUT
        os.makedirs(initial_dir, exist_ok=True)
        intunewin_path = filedialog.asksaveasfilename(
            title="Save IntuneWin package",
            initialdir=initial_dir,
            initialfile="MSStoreHelper-Queue.intunewin",
            defaultextension=".intunewin",
            filetypes=[("IntuneWin package", "*.intunewin"), ("All files", "*.*")],
        )
        if not intunewin_path:
            return

        threading.Thread(
            target=self._export_intunewin_worker,
            args=(tool_path, intunewin_path),
            daemon=True,
        ).start()

    def _export_intunewin_worker(self, tool_path, intunewin_path):
        self.after(0, lambda: self._update_status("📦 Building IntuneWin package...", Theme.INFO))
        self.after(0, lambda: self._log("INFO", f"Building IntuneWin package: {intunewin_path}"))
        try:
            generated, detection_script, count = StoreAPI.create_intunewin_package(
                self.download_queue,
                self.output_path,
                intunewin_path,
                tool_path,
                self._target_arch(),
            )
        except ValueError as exc:
            self.after(0, lambda e=str(exc): self._update_status("⚠️ Download files first", Theme.WARNING))
            self.after(0, lambda e=str(exc): self._log("WARNING", e))
        except Exception as exc:
            self.after(0, lambda: self._update_status("❌ IntuneWin export failed", Theme.DANGER))
            self.after(0, lambda e=str(exc): self._log("ERROR", f"Failed to build IntuneWin package: {e}"))
        else:
            self.after(0, lambda: self._update_status("✅ IntuneWin package exported", Theme.SUCCESS))
            self.after(0, lambda p=generated, d=detection_script, c=count: self._log("SUCCESS", f"IntuneWin package saved: {p} ({c} package(s)); detection script: {d}"))
    
    def _download_worker(self):
        if not os.path.exists(self.output_path):
            os.makedirs(self.output_path)
            self.after(0, lambda: self._log("INFO", f"Created output directory: {self.output_path}"))
        
        self.after(0, lambda: self._log("INFO", f"Starting download of {len(self.download_queue)} files"))
        
        total = len(self.download_queue)
        success_count = 0
        
        for i, pkg in enumerate(self.download_queue):
            fname = pkg['FileName']
            self.after(0, lambda n=fname: self._update_status(f"⬇️ Downloading {n[:40]}...", Theme.INFO))
            self.after(0, lambda n=fname, idx=i+1, tot=total: self._log("INFO", f"[{idx}/{tot}] Downloading: {n}"))
            
            if '_status_widget' in pkg:
                self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Downloading...", text_color=Theme.INFO))
            
            filepath = os.path.join(self.output_path, fname)
            pkg['LocalPath'] = filepath
            
            def progress_cb(val, idx=i, tot=total):
                self.after(0, lambda v=(idx + val) / tot: self._update_progress(v))
            
            success, error_msg = StoreAPI.download_file(pkg['Url'], filepath, progress_cb)
            
            if '_status_widget' in pkg:
                if success:
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="✅ Done", text_color=Theme.SUCCESS))
                    self.after(0, lambda n=fname: self._log("SUCCESS", f"  Downloaded: {n}"))
                    success_count += 1
                    if self.shared_cache_enabled.get():
                        cache_success, cache_msg = StoreAPI.cache_downloaded_artifact(pkg, self.shared_cache_path)
                        level = "SUCCESS" if cache_success else "WARNING"
                        self.after(0, lambda lvl=level, m=cache_msg: self._log(lvl, f"  Shared cache: {m}"))
                else:
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="❌ Failed", text_color=Theme.DANGER))
                    self.after(0, lambda n=fname, e=error_msg: self._log("ERROR", f"  Failed to download {n}: {e}"))
        
        self.after(0, lambda: self._update_progress(0))
        self.after(0, lambda: self._update_status("✅ Downloads complete!", Theme.SUCCESS))
        self.after(0, lambda: self._log("SUCCESS", f"Download complete: {success_count}/{total} files successful"))
        self.after(0, lambda: self._log("INFO", f"Files saved to: {self.output_path}"))
    
    def _start_install(self):
        if not IS_ADMIN:
            self._update_status("⚠️ Administrator required", Theme.WARNING)
            return
        
        to_install = [p for p in self.download_queue if p.get('LocalPath') and os.path.exists(p.get('LocalPath', ''))]
        to_install = StoreAPI.order_packages_for_install(to_install, self._target_arch())
        if not to_install:
            self._update_status("⚠️ No downloaded files", Theme.WARNING)
            return
        
        threading.Thread(target=self._install_worker, args=(to_install,), daemon=True).start()
    
    def _install_worker(self, packages):
        self.after(0, lambda: self._log("INFO", f"Starting installation of {len(packages)} packages"))
        self.after(0, lambda: self._log("INFO", "Note: Install order matters - dependencies should be installed first"))
        
        success_count = 0
        skipped_count = 0
        total = len(packages)
        
        for i, pkg in enumerate(packages):
            fname = pkg['FileName']
            filepath = pkg['LocalPath']
            package_name = pkg.get("PackageIdentity") or package_identity(fname)
            available_version = pkg.get("AvailableVersion") or format_version_tuple(package_version_tuple(fname))
            
            self.after(0, lambda n=fname: self._update_status(f"📦 Installing {n[:40]}...", Theme.INFO))
            self.after(0, lambda n=fname, idx=i+1, tot=total: self._log("INFO", f"[{idx}/{tot}] Installing: {n}"))
            self.after(0, lambda p=filepath: self._log("DEBUG", f"  Path: {p}"))
            
            if '_status_widget' in pkg:
                self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Installing...", text_color=Theme.INFO))

            should_skip, installed_version, package_name = StoreAPI.should_skip_installed_package(pkg)
            if should_skip:
                skipped_count += 1
                if '_status_widget' in pkg:
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Up to date", text_color=Theme.SUCCESS))
                self.after(0, lambda n=package_name, i=installed_version, a=available_version: self._log("SUCCESS", f"  Skipped {n}: installed {i} >= available {a}"))
                continue

            signature_ok, signature_msg = StoreAPI.verify_package_signature(filepath)
            if not signature_ok:
                if '_status_widget' in pkg:
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Signature blocked", text_color=Theme.DANGER))
                self.after(0, lambda n=fname, m=signature_msg: self._log("ERROR", f"  Signature verification blocked {n}: {m}"))
                continue

            self.after(0, lambda m=signature_msg: self._log("DEBUG", f"  Signature verified: {m}"))
            
            success, error_msg = StoreAPI.install_package(filepath)
            
            if '_status_widget' in pkg:
                if success:
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="✅ Installed", text_color=Theme.SUCCESS))
                    self.after(0, lambda n=fname: self._log("SUCCESS", f"  Successfully installed: {n}"))
                    success_count += 1
                else:
                    if StoreAPI.is_noop_install_error(error_msg):
                        skipped_count += 1
                        self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Already current", text_color=Theme.SUCCESS))
                        self.after(0, lambda n=package_name: self._log("SUCCESS", f"  No-op for {n}: installed version is already newer"))
                        continue

                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="❌ Error", text_color=Theme.DANGER))
                    self.after(0, lambda n=fname: self._log("ERROR", f"  Failed to install: {n}"))
                    
                    # Log detailed error message
                    error_lines = error_msg.split('\n')
                    for line in error_lines:
                        line = line.strip()
                        if line:
                            self.after(0, lambda l=line: self._log("ERROR", f"    {l}"))
                    
                    # Provide helpful hints based on common errors
                    error_lower = error_msg.lower()
                    if "0x80073cf3" in error_lower or "already installed" in error_lower:
                        self.after(0, lambda: self._log("INFO", "    Hint: App may already be installed or needs update"))
                    elif "0x80073d19" in error_lower or "dependency" in error_lower:
                        self.after(0, lambda: self._log("INFO", "    Hint: Missing dependency - install VCLibs and .NET packages first"))
                    elif "0x80073cff" in error_lower or "sideload" in error_lower:
                        self.after(0, lambda: self._log("INFO", "    Hint: Enable Developer Mode or Sideloading in Windows Settings"))
                    elif "0x80073cf9" in error_lower:
                        self.after(0, lambda: self._log("INFO", "    Hint: Package may require a different Windows version"))
                    elif "0x80073d02" in error_lower or "in use" in error_lower:
                        self.after(0, lambda: self._log("INFO", "    Hint: Close the app if it's running and try again"))
                    elif "access" in error_lower or "denied" in error_lower:
                        self.after(0, lambda: self._log("INFO", "    Hint: Run as Administrator"))
                    elif "signature" in error_lower or "certificate" in error_lower:
                        self.after(0, lambda: self._log("INFO", "    Hint: Package signature issue - try a different version"))
        
        self.after(0, lambda: self._update_status("✅ Installation complete!", Theme.SUCCESS))
        
        completed_count = success_count + skipped_count

        if completed_count == total:
            self.after(0, lambda: self._log("SUCCESS", f"Installation complete: {success_count} installed, {skipped_count} skipped/no-op"))
        else:
            self.after(0, lambda: self._log("WARNING", f"Installation complete: {success_count} installed, {skipped_count} skipped/no-op, {total - completed_count} failed"))
            self.after(0, lambda: self._log("INFO", "Tip: Check the errors above. Common fixes:"))
            self.after(0, lambda: self._log("INFO", "  1. Install dependencies (VCLibs, .NET) before main apps"))
            self.after(0, lambda: self._log("INFO", "  2. Enable Developer Mode in Windows Settings"))
            self.after(0, lambda: self._log("INFO", "  3. Try a different package version (older/newer)"))
    
    def _run_repair(self):
        if not IS_ADMIN:
            self._update_status("⚠️ Administrator required", Theme.WARNING)
            return

        self._update_status("🔧 Repairing Store...", Theme.INFO)
        self._log("INFO", "Store repair preset started")
        threading.Thread(target=self._repair_worker, daemon=True).start()

    def _run_provisioning_repair(self):
        if not IS_ADMIN:
            self._update_status("⚠️ Administrator required", Theme.WARNING)
            return

        self._update_status("👥 Repairing Store provisioning...", Theme.INFO)
        self._log("INFO", "Store provisioning repair started")
        threading.Thread(target=self._provisioning_repair_worker, daemon=True).start()

    def _run_licensing_reset(self):
        if not IS_ADMIN:
            self._update_status("⚠️ Administrator required", Theme.WARNING)
            return

        self._update_status("🔐 Resetting Store licensing...", Theme.INFO)
        self._log("INFO", "Store licensing reset started")
        threading.Thread(target=self._licensing_reset_worker, daemon=True).start()

    def _run_cache_rebuild(self):
        self._update_status("🧹 Rebuilding Store cache...", Theme.INFO)
        self._log("INFO", "Store cache scan and offline rebuild started")
        threading.Thread(target=self._cache_rebuild_worker, daemon=True).start()
    
    def _repair_worker(self):
        self.after(0, lambda: self._log("INFO", "Starting Microsoft Store repair..."))
        
        def log_cb(msg):
            self.after(0, lambda m=msg: self._update_status(m, Theme.INFO))
            self.after(0, lambda m=msg: self._log("INFO", m))
        
        def progress_cb(val):
            self.after(0, lambda v=val: self._update_progress(v))
        
        results = StoreAPI.run_repair(log_cb, progress_cb)
        success_count = sum(1 for _, ok in results if ok)
        
        self.after(0, lambda: self._update_progress(0))
        
        # Log results
        self.after(0, lambda: self._log("INFO", "Repair results:"))
        for desc, ok in results:
            if ok:
                self.after(0, lambda d=desc: self._log("SUCCESS", f"  ✓ {d}"))
            else:
                self.after(0, lambda d=desc: self._log("ERROR", f"  ✗ {d}"))
        
        if success_count == len(results):
            self.after(0, lambda: self._update_status("✅ Repair complete! Please restart your PC.", Theme.SUCCESS))
            self.after(0, lambda: self._log("SUCCESS", "Repair complete! Please restart your PC for changes to take effect."))
        else:
            self.after(0, lambda: self._update_status(f"⚠️ Repair done ({success_count}/{len(results)} steps)", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", f"Repair partially complete: {success_count}/{len(results)} steps succeeded"))

    def _provisioning_repair_worker(self):
        def log_cb(msg):
            self.after(0, lambda m=msg: self._update_status(m, Theme.INFO))
            self.after(0, lambda m=msg: self._log("INFO", m))

        def progress_cb(val):
            self.after(0, lambda v=val: self._update_progress(v))

        results = StoreAPI.run_provisioning_repair(log_cb, progress_cb)
        success_count = sum(1 for _, ok in results if ok)

        self.after(0, lambda: self._update_progress(0))
        self.after(0, lambda: self._log("INFO", "Provisioning repair results:"))
        for desc, ok in results:
            if ok:
                self.after(0, lambda d=desc: self._log("SUCCESS", f"  ✓ {d}"))
            else:
                self.after(0, lambda d=desc: self._log("ERROR", f"  ✗ {d}"))

        if success_count == len(results):
            self.after(0, lambda: self._update_status("✅ Provisioning repair complete", Theme.SUCCESS))
            self.after(0, lambda: self._log("SUCCESS", "Provisioning repair complete for Store-related packages."))
        else:
            self.after(0, lambda: self._update_status(f"⚠️ Provisioning repair done ({success_count}/{len(results)} steps)", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", f"Provisioning repair partially complete: {success_count}/{len(results)} steps succeeded"))

    def _licensing_reset_worker(self):
        def log_cb(msg):
            self.after(0, lambda m=msg: self._update_status(m, Theme.INFO))
            self.after(0, lambda m=msg: self._log("INFO", m))

        def progress_cb(val):
            self.after(0, lambda v=val: self._update_progress(v))

        results = StoreAPI.run_licensing_reset(log_cb, progress_cb)
        success_count = sum(1 for _, ok in results if ok)

        self.after(0, lambda: self._update_progress(0))
        self.after(0, lambda: self._log("INFO", "Licensing reset results:"))
        for desc, ok in results:
            if ok:
                self.after(0, lambda d=desc: self._log("SUCCESS", f"  ✓ {d}"))
            else:
                self.after(0, lambda d=desc: self._log("ERROR", f"  ✗ {d}"))

        if success_count == len(results):
            self.after(0, lambda: self._update_status("✅ Licensing reset complete", Theme.SUCCESS))
            self.after(0, lambda: self._log("SUCCESS", "Licensing reset complete. Reopen affected Store apps if needed."))
        else:
            self.after(0, lambda: self._update_status(f"⚠️ Licensing reset done ({success_count}/{len(results)} steps)", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", f"Licensing reset partially complete: {success_count}/{len(results)} steps succeeded"))

    def _cache_rebuild_worker(self):
        def log_cb(msg):
            self.after(0, lambda m=msg: self._update_status(m, Theme.INFO))
            self.after(0, lambda m=msg: self._log("INFO", m))

        def progress_cb(val):
            self.after(0, lambda v=val: self._update_progress(v))

        results = StoreAPI.run_cache_rebuild(log_cb, progress_cb)
        success_count = sum(1 for _, ok in results if ok)

        self.after(0, lambda: self._update_progress(0))
        self.after(0, lambda: self._log("INFO", "Cache rebuild results:"))
        for desc, ok in results:
            if ok:
                self.after(0, lambda d=desc: self._log("SUCCESS", f"  ✓ {d}"))
            else:
                self.after(0, lambda d=desc: self._log("ERROR", f"  ✗ {d}"))

        if success_count == len(results):
            self.after(0, lambda: self._update_status("✅ Cache rebuild complete", Theme.SUCCESS))
            self.after(0, lambda: self._log("SUCCESS", "Store cache rebuild complete. Previous caches were kept as .bak folders."))
        else:
            self.after(0, lambda: self._update_status(f"⚠️ Cache rebuild done ({success_count}/{len(results)} steps)", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", f"Cache rebuild partially complete: {success_count}/{len(results)} steps succeeded"))


if __name__ == "__main__":
    app = MSStoreHelperApp()
    app.mainloop()
