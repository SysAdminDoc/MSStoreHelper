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
import argparse
import threading
import ctypes
import webbrowser
import json
import re
import hashlib
import shutil
import ssl
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from http.server import ThreadingHTTPServer
from urllib.parse import quote, urlsplit
try:
    import winreg
except ImportError:
    winreg = None
from tkinter import filedialog
from datetime import datetime, timezone
from msstore_package_resolution import (
    annotate_package,
    compare_version_tuples,
    format_version_tuple,
    installed_version_satisfies_package,
    is_dependency_package,
    is_arch_compatible,
    is_installable_package,
    order_packages_for_install,
    package_identity,
    package_version_tuple,
    package_role_label,
    select_recommended_packages,
    version_tuple_from_text,
)
from package_ingress import (
    PackageIngressError,
    ensure_path_within_root,
    package_path as confined_package_path,
    validate_existing_package_path,
    validate_package_filename,
    validate_package_record,
    validate_package_url,
    validate_response_redirects,
)
from package_trust import (
    PackageTrustError,
    TRUST_STATE_BLOCKED,
    TRUST_STATE_REVIEW_REQUIRED,
    blocked_trust_report,
    evaluate_package_trust,
    normalize_chain_status,
    package_filename_metadata,
    read_package_manifest,
    review_trust_report,
    trust_report_allows_automation,
)
from diagnostic_bundle import (
    DiagnosticRedactionError,
    diagnostic_preview_text,
    prepare_diagnostic_entries,
    redact_structure,
    redact_text,
    write_prepared_bundle,
)
from repair_transaction import (
    DEFAULT_REPAIR_RETENTION,
    RepairTransactionError,
    build_repair_plan,
    build_restore_plan,
    execute_repair_plan,
    execute_restore_plan,
    list_repair_backups,
    normalize_retention,
    render_repair_plan,
    render_restore_plan,
)
from mirror_service import (
    MIRROR_AUDIT_FILENAME,
    MirrorAuditLog,
    MirrorConfigurationError,
    atomic_write_json,
    create_bearer_token,
    make_mirror_handler,
    mirror_base_url,
    normalize_token_ttl,
    utc_timestamp as mirror_utc_timestamp,
    validate_network_policy,
    wrap_server_tls,
)
from store_sources import (
    StoreSourceError,
    detect_source_health,
    package_lookup_fallbacks,
    request_with_retries,
    source_status_summary,
)

# ==================== DEPENDENCY CHECK ====================
REQUIRED_DEPENDENCIES = {
    "customtkinter": "customtkinter==5.2.2",
    "requests": "requests==2.32.5",
    "bs4": "beautifulsoup4==4.14.3",
}


def find_missing_dependencies(importer=__import__):
    missing = []
    for import_name, requirement in REQUIRED_DEPENDENCIES.items():
        try:
            importer(import_name)
        except ImportError:
            missing.append(requirement)
    return missing


def dependency_setup_message(missing):
    requirements = ", ".join(missing)
    return (
        f"Missing Python dependencies: {requirements}\n"
        "Install pinned dependencies with:\n"
        "  py -3 -m pip install -r requirements.txt\n"
        "For offline installs, prepare a wheelhouse on a connected PC:\n"
        "  py -3 -m pip download -r requirements.txt -d wheelhouse\n"
        "Then install on the target PC with:\n"
        "  py -3 -m pip install --no-index --find-links wheelhouse -r requirements.txt"
    )


missing_dependencies = find_missing_dependencies()
if missing_dependencies:
    print(dependency_setup_message(missing_dependencies), file=sys.stderr)
    raise SystemExit(1)

import customtkinter as ctk
import requests
from bs4 import BeautifulSoup

# ==================== CONFIGURATION ====================

APP_VERSION = "3.36.0"
APP_NAME = "MSStoreHelper"
API_URL = "https://store.rg-adguard.net/api/GetFiles"
STORE_SEARCH_URL = "https://storeedgefd.dsx.mp.microsoft.com/v9.0/manifestSearch"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
CACHE_MANIFEST_NAME = "msstorehelper-cache-manifest.json"
MIRROR_INDEX_NAME = "msstorehelper-mirror-index.json"
CACHE_HISTORY_LIMIT = 2
WINGET_IMPORT_SCHEMA = "https://aka.ms/winget-packages.schema.2.0.json"
WINGET_MSSTORE_SOURCE = {
    "Argument": "https://storeedgefd.dsx.mp.microsoft.com/v9.0",
    "Identifier": "StoreEdgeFD",
    "Name": "msstore",
    "Type": "Microsoft.Rest",
}
THEME_MODE_VALUES = ["System", "Dark", "Light"]
STORE_RING_VALUES = ["Retail", "RP", "WIS", "WIF"]
STORE_LANGUAGE_VALUES = ["en-US", "en-GB", "de-DE", "fr-FR", "es-ES", "it-IT", "ja-JP", "ko-KR", "pt-BR", "zh-CN", "zh-TW"]
STORE_MARKET_VALUES = ["US", "GB", "CA", "AU", "DE", "FR", "ES", "IT", "JP", "KR", "BR", "CN", "TW"]
KEEP_UPDATED_INTERVAL_MS = 6 * 60 * 60 * 1000
KEEP_UPDATED_START_DELAY_MS = 5000
APPINSTALLER_NS = "http://schemas.microsoft.com/appx/appinstaller/2021"
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

APP_DATA_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), APP_NAME)
USER_PROFILE_PATH = os.path.join(APP_DATA_DIR, "profile.json")
REPAIR_BACKUP_DIR = os.path.join(APP_DATA_DIR, "RepairBackups")
DOWNLOAD_STATE_PATH = os.path.join(APP_DATA_DIR, "download-state.json")
TRUST_REVIEW_JOURNAL_PATH = os.path.join(APP_DATA_DIR, "trust-review.jsonl")

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
    DEFAULT_ACCENT = "#0f82f2"
    MODE = "Dark"

    # Graphite workspace surfaces
    BG_DARK = ("#f4f6f8", "#10151a")
    BG_SIDEBAR = ("#eef2f5", "#141a20")
    BG_CARD = ("#ffffff", "#171e24")
    BG_ELEVATED = ("#f8fafc", "#1b232b")
    BG_CARD_HOVER = ("#e7eef5", "#202a33")
    BG_INPUT = ("#edf2f6", "#11181e")
    
    # Accent colors
    PRIMARY = DEFAULT_ACCENT
    PRIMARY_HOVER = ("#086fcf", "#43a1ff")
    PRIMARY_OUTLINE_TEXT = ("#075eab", "#75baff")
    SUCCESS = ("#087a55", "#25b780")
    SUCCESS_HOVER = ("#066846", "#42c996")
    WARNING = ("#985b00", "#f0aa32")
    DANGER = ("#c72c3b", "#ef5965")
    DANGER_HOVER = ("#a92331", "#f27b84")
    INFO = ("#0876a8", "#40b5e5")
    
    # Text colors
    TEXT_PRIMARY = ("#101820", "#f3f6f8")
    TEXT_SECONDARY = ("#3f5263", "#b8c2ca")
    TEXT_MUTED = ("#526678", "#9aa6af")
    
    # Special
    BORDER = ("#c7d0da", "#35414b")
    BORDER_SUBTLE = ("#dce2e8", "#28333c")
    BUNDLE_COLOR = ("#0876a8", "#40b5e5")
    ENCRYPTED_COLOR = ("#c72c3b", "#ef5965")
    ARCH_MATCH = ("#087a55", "#4bd3a0")

    @staticmethod
    def normalize_mode(mode):
        value = str(mode or "System").strip().title()
        return value if value in THEME_MODE_VALUES else "System"

    @staticmethod
    def sanitize_hex_color(color):
        value = str(color or "").strip()
        if re.fullmatch(r"#[0-9a-fA-F]{6}", value):
            return value.lower()
        return None

    @staticmethod
    def shift_hex_color(color, amount):
        value = Theme.sanitize_hex_color(color) or Theme.DEFAULT_ACCENT
        amount = max(-1.0, min(1.0, float(amount)))
        channels = [int(value[i:i + 2], 16) for i in (1, 3, 5)]

        shifted = []
        for channel in channels:
            if amount >= 0:
                shifted.append(round(channel + (255 - channel) * amount))
            else:
                shifted.append(round(channel * (1 + amount)))
        return "#" + "".join(f"{channel:02x}" for channel in shifted)

    @staticmethod
    def color_for_mode(color, mode="Dark"):
        if isinstance(color, (tuple, list)) and len(color) >= 2:
            return color[0] if Theme.resolve_mode(mode, apps_use_light=False) == "Light" else color[1]
        return color

    @staticmethod
    def relative_luminance(color):
        value = Theme.sanitize_hex_color(color) or Theme.DEFAULT_ACCENT
        channels = [int(value[i:i + 2], 16) / 255 for i in (1, 3, 5)]
        linear = []
        for channel in channels:
            if channel <= 0.03928:
                linear.append(channel / 12.92)
            else:
                linear.append(((channel + 0.055) / 1.055) ** 2.4)
        return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]

    @staticmethod
    def contrast_ratio(foreground, background):
        fg = Theme.relative_luminance(foreground)
        bg = Theme.relative_luminance(background)
        lighter = max(fg, bg)
        darker = min(fg, bg)
        return (lighter + 0.05) / (darker + 0.05)

    @staticmethod
    def accent_from_windows_dword(value):
        try:
            raw = int(value)
        except (TypeError, ValueError):
            return None

        red = raw & 0xFF
        green = (raw >> 8) & 0xFF
        blue = (raw >> 16) & 0xFF
        return f"#{red:02x}{green:02x}{blue:02x}"

    @staticmethod
    def _read_registry_dword(path, name):
        if winreg is None:
            return None
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                value, _value_type = winreg.QueryValueEx(key, name)
            return int(value)
        except OSError:
            return None

    @staticmethod
    def windows_apps_use_light_theme():
        value = Theme._read_registry_dword(
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            "AppsUseLightTheme",
        )
        if value is None:
            return False
        return bool(value)

    @staticmethod
    def read_windows_accent_color():
        value = Theme._read_registry_dword(r"Software\Microsoft\Windows\DWM", "AccentColor")
        return Theme.sanitize_hex_color(Theme.accent_from_windows_dword(value))

    @staticmethod
    def resolve_mode(mode, apps_use_light=None):
        normalized = Theme.normalize_mode(mode)
        if normalized == "System":
            if apps_use_light is None:
                apps_use_light = Theme.windows_apps_use_light_theme()
            return "Light" if apps_use_light else "Dark"
        return normalized

    @classmethod
    def configure_accent(cls, accent_color=None):
        accent = cls.sanitize_hex_color(accent_color) or cls.DEFAULT_ACCENT
        cls.PRIMARY = accent
        cls.PRIMARY_HOVER = (
            cls.shift_hex_color(accent, -0.14),
            cls.shift_hex_color(accent, 0.22),
        )
        return accent

    @classmethod
    def set_mode(cls, mode, accent_color=None):
        cls.MODE = cls.resolve_mode(mode)
        cls.configure_accent(accent_color or cls.read_windows_accent_color())
        return cls.MODE

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
    "🐧 WSL Distributions": {
        "description": "Linux distribution packages for WSL sideloading",
        "apps": [
            {"Name": "Ubuntu", "ProductId": "9PDXGNCFSCZV", "Description": "Ubuntu terminal environment for WSL", "Icon": "🐧"},
            {"Name": "Debian", "ProductId": "9MSVKQC78PK6", "Description": "Debian command-line environment for WSL", "Icon": "🐧"},
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
    "🐧 WSL Distros": {
        "description": "Queue Ubuntu and Debian Store distribution packages for offline WSL sideloading.",
        "apps": ["Ubuntu", "Debian"]
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

XBOX_CORE_PACKAGE_PINS = [
    {
        "Name": "Xbox Identity",
        "ProductId": "9WZDNCRD1HKW",
        "Identity": "Microsoft.XboxIdentityProvider",
        "KnownGoodVersions": ["12.50.6001.0"],
    },
    {
        "Name": "Gaming Services",
        "ProductId": "9MWPM2CQNLHN",
        "Identity": "Microsoft.GamingServices",
        "KnownGoodVersions": ["2.51.3002.0"],
    },
]

# ==================== BACKEND API ====================

class StoreAPI:
    """Handles all API communications"""

    @staticmethod
    def default_user_profile():
        return {
            "SearchHistory": [],
            "PinnedFavorites": [],
            "ThemeMode": "System",
            "StoreRing": "Retail",
            "StoreLanguage": "en-US",
            "StoreMarket": "US",
            "KeepUpdatedEnabled": False,
            "KeepUpdatedLastScan": "",
            "RepairRetentionCount": DEFAULT_REPAIR_RETENTION,
        }

    @staticmethod
    def normalize_store_ring(value):
        text = str(value or "Retail").strip()
        if text.lower() == "retail":
            return "Retail"
        upper = text.upper()
        return upper if upper in STORE_RING_VALUES else "Retail"

    @staticmethod
    def normalize_store_language(value):
        text = str(value or "en-US").strip().replace("_", "-")
        parts = text.split("-")
        if len(parts) == 2 and all(part.isalpha() for part in parts):
            language = f"{parts[0].lower()}-{parts[1].upper()}"
            if re.fullmatch(r"[a-z]{2}-[A-Z]{2}", language):
                return language
        return "en-US"

    @staticmethod
    def normalize_store_market(value):
        text = str(value or "US").strip().upper()
        return text if re.fullmatch(r"[A-Z]{2}", text) else "US"

    @staticmethod
    def store_query_settings(ring=None, language=None, market=None):
        return {
            "Ring": StoreAPI.normalize_store_ring(ring),
            "Language": StoreAPI.normalize_store_language(language),
            "Market": StoreAPI.normalize_store_market(market),
        }

    @staticmethod
    def package_query_metadata(product_id, ring=None, language=None, market=None):
        metadata = StoreAPI.store_query_settings(ring, language, market)
        metadata["ProductId"] = str(product_id or "").strip()
        return metadata

    @staticmethod
    def expected_product_identities(product_id):
        product_id = str(product_id or "").strip().lower()
        if not product_id:
            return set()
        apps_by_name = {
            app["Name"].lower(): app
            for app in StoreAPI.catalog_apps()
        }
        expected = set()
        for requirement in LTSC_COMPONENT_REQUIREMENTS:
            app = apps_by_name.get(str(requirement.get("Name", "")).lower())
            if not app or str(app.get("ProductId", "")).lower() != product_id:
                continue
            expected.update(
                str(identity)
                for identity in requirement.get("Identities", [])
                if str(identity).strip()
            )
        for pin in XBOX_CORE_PACKAGE_PINS:
            if (
                str(pin.get("ProductId", "")).lower() == product_id
                and pin.get("Identity")
            ):
                expected.add(str(pin["Identity"]))
        return expected

    @staticmethod
    def attach_expected_trust_metadata(package, product_id=None):
        package = package.copy()
        query = package.get("StoreQuery") or {}
        product_id = str(
            product_id
            or package.get("ExpectedProductId")
            or query.get("ProductId")
            or ""
        ).strip()
        if product_id:
            package["ExpectedProductId"] = product_id

        try:
            filename_metadata = package_filename_metadata(package["FileName"])
        except (KeyError, PackageTrustError):
            return package
        if not package.get("ExpectedPackageFamilyName"):
            package["ExpectedPackageFamilyName"] = (
                filename_metadata["PackageFamilyName"]
            )
        if "ExpectedDependency" not in package:
            package["ExpectedDependency"] = is_dependency_package(package)

        expected_identities = StoreAPI.expected_product_identities(product_id)
        matching_identity = next(
            (
                identity
                for identity in expected_identities
                if identity.lower() == filename_metadata["Identity"].lower()
            ),
            None,
        )
        if matching_identity and not package.get("ExpectedPackageIdentity"):
            package["ExpectedPackageIdentity"] = matching_identity
        return package

    @staticmethod
    def catalog_apps():
        apps = []
        for category_name, category in APP_CATALOG.items():
            for app in category.get("apps", []):
                item = StoreAPI.normalize_favorite_app(app)
                item["Category"] = category_name
                apps.append(item)
        return apps

    @staticmethod
    def catalog_identity_map():
        apps_by_name = {app["Name"].lower(): app for app in StoreAPI.catalog_apps()}
        identities = {}
        for requirement in LTSC_COMPONENT_REQUIREMENTS:
            app = apps_by_name.get(str(requirement.get("Name", "")).lower())
            if not app:
                continue
            for identity in requirement.get("Identities", []):
                identities[str(identity).lower()] = app
        for pin in XBOX_CORE_PACKAGE_PINS:
            app = apps_by_name.get(str(pin.get("Name", "")).lower())
            if app and pin.get("Identity"):
                identities[str(pin["Identity"]).lower()] = app
        return identities

    @staticmethod
    def resolve_cli_app(identifier, searcher=None):
        text = str(identifier or "").strip()
        if not text:
            return None, "App identifier is required"

        text_lower = text.lower()
        for app in StoreAPI.catalog_apps():
            app_name = app.get("Name", "")
            product_id = app.get("ProductId", "")
            if text_lower == app_name.lower():
                resolved = app.copy()
                resolved["ResolvedFrom"] = "catalog-name"
                return resolved, None
            if text_lower == product_id.lower():
                resolved = app.copy()
                resolved["ResolvedFrom"] = "product-id"
                return resolved, None

        identity_app = StoreAPI.catalog_identity_map().get(text_lower)
        if identity_app:
            resolved = identity_app.copy()
            resolved["ResolvedFrom"] = "package-identity"
            resolved["PackageIdentity"] = text
            return resolved, None

        searcher = searcher or StoreAPI.search_store_with_diagnostics
        diagnostic = searcher(text, max_results=5)
        results = diagnostic.get("Results", []) if isinstance(diagnostic, dict) else []
        for result in results:
            if result.get("ProductId"):
                resolved = StoreAPI.normalize_favorite_app(result)
                resolved["ResolvedFrom"] = "store-search"
                return resolved, None

        errors = diagnostic.get("Errors", []) if isinstance(diagnostic, dict) else []
        if errors:
            return None, "; ".join(errors)
        return None, f"No Store app matched '{text}'"

    @staticmethod
    def load_user_profile(path=USER_PROFILE_PATH):
        try:
            if not os.path.exists(path):
                return StoreAPI.default_user_profile()
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
            profile = StoreAPI.default_user_profile()
            profile["SearchHistory"] = [str(item) for item in data.get("SearchHistory", []) if str(item).strip()][:10]
            profile["PinnedFavorites"] = [
                StoreAPI.normalize_favorite_app(item)
                for item in data.get("PinnedFavorites", [])
                if isinstance(item, dict) and item.get("Name") and item.get("ProductId")
            ][:20]
            profile["ThemeMode"] = Theme.normalize_mode(data.get("ThemeMode", "System"))
            profile["StoreRing"] = StoreAPI.normalize_store_ring(data.get("StoreRing", "Retail"))
            profile["StoreLanguage"] = StoreAPI.normalize_store_language(data.get("StoreLanguage", "en-US"))
            profile["StoreMarket"] = StoreAPI.normalize_store_market(data.get("StoreMarket", "US"))
            profile["KeepUpdatedEnabled"] = bool(data.get("KeepUpdatedEnabled", False))
            profile["KeepUpdatedLastScan"] = str(data.get("KeepUpdatedLastScan", "") or "")
            profile["RepairRetentionCount"] = normalize_retention(
                data.get(
                    "RepairRetentionCount",
                    DEFAULT_REPAIR_RETENTION,
                )
            )
            return profile
        except Exception:
            return StoreAPI.default_user_profile()

    @staticmethod
    def save_user_profile(profile, path=USER_PROFILE_PATH):
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(profile, handle, indent=2)
            handle.write("\n")
        return path

    @staticmethod
    def normalize_favorite_app(app_data):
        return {
            "Name": str(app_data.get("Name", "")).strip(),
            "ProductId": str(app_data.get("ProductId", "")).strip(),
            "Publisher": str(app_data.get("Publisher", "")).strip(),
            "Description": str(app_data.get("Description", "")).strip(),
            "Icon": app_data.get("Icon", "📦"),
        }

    @staticmethod
    def add_search_history(profile, query, max_items=10):
        query = str(query or "").strip()
        if not query:
            return profile

        history = [item for item in profile.get("SearchHistory", []) if item.lower() != query.lower()]
        profile["SearchHistory"] = [query] + history[:max_items - 1]
        return profile

    @staticmethod
    def add_pinned_favorites(profile, apps, max_items=20):
        favorites = profile.get("PinnedFavorites", [])
        by_id = {app["ProductId"].lower(): app for app in favorites if app.get("ProductId")}
        ordered_ids = [app["ProductId"].lower() for app in favorites if app.get("ProductId")]

        added = 0
        for app in apps:
            favorite = StoreAPI.normalize_favorite_app(app)
            if not favorite["Name"] or not favorite["ProductId"]:
                continue
            key = favorite["ProductId"].lower()
            if key not in by_id:
                ordered_ids.insert(0, key)
                added += 1
            by_id[key] = favorite

        seen = set()
        profile["PinnedFavorites"] = []
        for key in ordered_ids:
            if key in seen or key not in by_id:
                continue
            seen.add(key)
            profile["PinnedFavorites"].append(by_id[key])
            if len(profile["PinnedFavorites"]) >= max_items:
                break
        return added

    @staticmethod
    def detect_source_health():
        search_payload = {"Query": {"KeyWord": "calculator", "MatchType": "Substring"}}
        search_headers = {"User-Agent": USER_AGENT, "Content-Type": "application/json"}
        rg_payload = {"type": "ProductId", "url": "9N0DX20HK701", "ring": "Retail", "lang": "en-US"}
        rg_headers = {"User-Agent": USER_AGENT, "Content-Type": "application/x-www-form-urlencoded"}
        return detect_source_health(
            storeedge_request=lambda: requests.post(STORE_SEARCH_URL, json=search_payload, headers=search_headers, timeout=8),
            rgadguard_request=lambda: requests.post(API_URL, data=rg_payload, headers=rg_headers, timeout=8),
        )

    @staticmethod
    def _source_diagnostic(source_name, packages=None, results=None, errors=None, fallbacks=None):
        return {
            "Source": source_name,
            "Packages": packages or [],
            "Results": results or [],
            "Errors": errors or [],
            "Fallbacks": fallbacks or [],
        }
    
    @staticmethod
    def search_store(query, max_results=25):
        return StoreAPI.search_store_with_diagnostics(query, max_results)["Results"]

    @staticmethod
    def search_store_with_diagnostics(query, max_results=25):
        """Search Microsoft Store by app name"""
        payload = {"Query": {"KeyWord": query, "MatchType": "Substring"}}
        headers = {"User-Agent": USER_AGENT, "Content-Type": "application/json"}
        try:
            resp, retry_errors = request_with_retries(
                "Microsoft Store Search API",
                lambda: requests.post(STORE_SEARCH_URL, json=payload, headers=headers, timeout=15),
                attempts=2,
            )
            data = resp.json()
            
            results = []
            for item in data.get("Data", [])[:max_results]:
                results.append({
                    "ProductId": item.get("PackageIdentifier", ""),
                    "Name": item.get("PackageName", "Unknown"),
                    "Publisher": item.get("Publisher", "Unknown"),
                })
            return StoreAPI._source_diagnostic("Microsoft Store Search API", results=results, errors=retry_errors)
            
        except StoreSourceError as exc:
            return StoreAPI._source_diagnostic(exc.source_name, errors=exc.errors)
        except Exception as exc:
            return StoreAPI._source_diagnostic("Microsoft Store Search API", errors=[f"{type(exc).__name__}: {exc}"])

    @staticmethod
    def parse_release_notes_html(product_id, html_text, url):
        soup = BeautifulSoup(html_text, "html.parser")
        title = soup.title.get_text(" ", strip=True) if soup.title else product_id
        notes = None
        source = "store-page"

        for script in soup.find_all("script"):
            text = script.get_text(" ", strip=True)
            if not text:
                continue
            for key in ("releaseNotes", "ReleaseNotes", "whatsNew", "WhatsNew", "whatIsNew"):
                marker = f'"{key}"'
                if marker not in text:
                    continue
                start = text.find(marker)
                colon = text.find(":", start)
                if colon == -1:
                    continue
                snippet = text[colon + 1:colon + 1200]
                match = re.search(r'"((?:\\.|[^"\\])*)"', snippet)
                if match and match.group(1).strip():
                    notes = bytes(match.group(1), "utf-8").decode("unicode_escape").strip()
                    source = key
                    break
            if notes:
                break

        if not notes:
            headings = soup.find_all(lambda tag: tag.name in {"h1", "h2", "h3", "h4"} and tag.get_text(" ", strip=True).lower() in {
                "what's new",
                "whats new",
                "what's new in this version",
                "release notes",
                "version notes",
            })
            for heading in headings:
                pieces = []
                for sibling in heading.find_next_siblings():
                    if sibling.name in {"h1", "h2", "h3", "h4"}:
                        break
                    text = sibling.get_text("\n", strip=True)
                    if text:
                        pieces.append(text)
                    if len("\n".join(pieces)) > 1200:
                        break
                if pieces:
                    notes = "\n".join(pieces).strip()
                    source = "heading"
                    break

        if not notes:
            ld_json = soup.find("script", attrs={"type": "application/ld+json"})
            if ld_json:
                try:
                    data = json.loads(ld_json.get_text())
                    notes = data.get("description", "").strip()
                    title = data.get("name", title)
                    source = "product-description"
                except Exception:
                    notes = None

        if not notes:
            notes = "No release notes were published on the Microsoft Store product page."
            source = "empty"

        return {
            "ProductId": product_id,
            "Title": title,
            "Url": url,
            "Notes": notes,
            "Source": source,
        }

    @staticmethod
    def fetch_release_notes(product_id, language="en-US", market="US"):
        settings = StoreAPI.store_query_settings(language=language, market=market)
        url = f"https://apps.microsoft.com/detail/{product_id}?hl={settings['Language']}&gl={settings['Market']}"
        response, _retry_errors = request_with_retries(
            "Microsoft Store product page",
            lambda: requests.get(url, timeout=20, headers={"User-Agent": USER_AGENT}),
            attempts=2,
        )
        notes = StoreAPI.parse_release_notes_html(product_id, response.text, response.url)
        notes["StoreQuery"] = StoreAPI.package_query_metadata(product_id, language=settings["Language"], market=settings["Market"])
        return notes
    
    @staticmethod
    def get_packages(product_id, ring="Retail", language="en-US", market="US"):
        return StoreAPI.get_packages_with_diagnostics(product_id, ring, language, market)["Packages"]

    @staticmethod
    def get_packages_with_diagnostics(product_id, ring="Retail", language="en-US", market="US"):
        """Get downloadable packages for a product"""
        query = StoreAPI.package_query_metadata(product_id, ring, language, market)
        payload = {"type": "ProductId", "url": product_id, "ring": query["Ring"], "lang": query["Language"]}
        headers = {"User-Agent": USER_AGENT, "Content-Type": "application/x-www-form-urlencoded"}
        try:
            resp, retry_errors = request_with_retries(
                "RG-Adguard package proxy",
                lambda: requests.post(API_URL, data=payload, headers=headers, timeout=30),
                attempts=2,
            )
            
            soup = BeautifulSoup(resp.text, "html.parser")
            table = soup.find("table", class_="tftable")
            
            if not table:
                statuses = StoreAPI.detect_source_health()
                diagnostic = StoreAPI._source_diagnostic(
                    "RG-Adguard package proxy",
                    errors=retry_errors + ["RG-Adguard response did not include a package table"],
                    fallbacks=package_lookup_fallbacks(product_id, statuses),
                )
                diagnostic["Query"] = query
                return diagnostic

            results = []
            ingress_errors = []
            for row in table.find_all("tr"):
                cols = row.find_all("td")
                if not cols:
                    continue
                
                link = cols[0].find("a")
                if not link:
                    continue
                if link.text.strip().lower().endswith(".blockmap"):
                    continue
                
                try:
                    ingress = validate_package_record(
                        {
                            "FileName": link.text.strip(),
                            "Url": link.get("href"),
                        },
                        require_url=True,
                    )
                except PackageIngressError as exc:
                    ingress_errors.append(f"Rejected unsafe package metadata: {exc}")
                    continue
                name = ingress["FileName"]
                url = ingress["Url"]
                
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
                    "SizeBytes": None, "SizeStr": "—", "StoreQuery": query.copy(),
                    "SafeFileName": name,
                }
                package = annotate_package(package)
                results.append(
                    StoreAPI.attach_expected_trust_metadata(
                        package,
                        query["ProductId"],
                    )
                )
            
            diagnostic = StoreAPI._source_diagnostic(
                "RG-Adguard package proxy",
                packages=results,
                errors=retry_errors + ingress_errors,
            )
            diagnostic["Query"] = query
            return diagnostic
            
        except StoreSourceError as exc:
            statuses = StoreAPI.detect_source_health()
            diagnostic = StoreAPI._source_diagnostic(
                exc.source_name,
                errors=exc.errors,
                fallbacks=package_lookup_fallbacks(product_id, statuses),
            )
            diagnostic["Query"] = query
            return diagnostic
        except Exception as exc:
            statuses = StoreAPI.detect_source_health()
            diagnostic = StoreAPI._source_diagnostic(
                "RG-Adguard package proxy",
                errors=[f"{type(exc).__name__}: {exc}"],
                fallbacks=package_lookup_fallbacks(product_id, statuses),
            )
            diagnostic["Query"] = query
            return diagnostic
    
    @staticmethod
    def get_file_size(url):
        try:
            url = validate_package_url(url)
            resp = requests.head(url, timeout=10, allow_redirects=True)
            validate_response_redirects(url, resp)
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
    def select_pinned_xbox_packages(packages, target_arch, prefer_exact_arch=False):
        selected = []
        dependencies = [
            package for package in packages
            if is_dependency_package(package) and is_installable_package(package) and is_arch_compatible(package, target_arch)
        ]
        selected.extend(annotate_package(package.copy()) for package in select_recommended_packages(dependencies, target_arch, prefer_exact_arch))

        for pin in XBOX_CORE_PACKAGE_PINS:
            identity = pin["Identity"].lower()
            candidates = [
                package for package in packages
                if package_identity(package["FileName"]).lower() == identity
                and is_installable_package(package)
                and is_arch_compatible(package, target_arch)
            ]
            if not candidates:
                continue

            pinned_versions = set(pin["KnownGoodVersions"])
            pinned_candidates = [
                package for package in candidates
                if format_version_tuple(package_version_tuple(package["FileName"])) in pinned_versions
            ]
            source = pinned_candidates or candidates
            recommended = select_recommended_packages(source, target_arch, prefer_exact_arch)
            if not recommended:
                continue

            package = annotate_package(recommended[0].copy())
            package["XboxCoreName"] = pin["Name"]
            package["PinnedVersions"] = list(pin["KnownGoodVersions"])
            package["PinnedVersionMatched"] = bool(pinned_candidates)
            selected.append(package)

        return order_packages_for_install(selected, target_arch)

    @staticmethod
    def file_sha256(path, chunk_size=1024 * 1024):
        digest = hashlib.sha256()
        with open(path, "rb") as handle:
            for chunk in iter(lambda: handle.read(chunk_size), b""):
                digest.update(chunk)
        return digest.hexdigest()

    @staticmethod
    def query_package_signature(filepath):
        package_path = validate_existing_package_path(
            filepath,
            require_file=True,
        )
        safe_module = POWERSHELL_SECURITY_MODULE.replace("'", "''")
        command = f"""
Import-Module '{safe_module}' -ErrorAction Stop
$path = $env:MSSTOREHELPER_PACKAGE_PATH
if ([string]::IsNullOrWhiteSpace($path)) {{ throw 'Package path is missing' }}
$sig = Get-AuthenticodeSignature -FilePath $path
$baseChainOk = $false
$onlineChainOk = $false
$rootSubject = ''
$rootThumbprint = ''
$baseStatuses = @()
$onlineStatuses = @()
$revocationState = 'failed'
if ($sig.SignerCertificate) {{
    $baseChain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
    $baseChain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
    $baseChain.ChainPolicy.VerificationFlags = [System.Security.Cryptography.X509Certificates.X509VerificationFlags]::NoFlag
    $baseChainOk = $baseChain.Build($sig.SignerCertificate)
    $baseStatuses = @($baseChain.ChainStatus | ForEach-Object {{ $_.Status.ToString() }})
    if ($baseChain.ChainElements.Count -gt 0) {{
        $root = $baseChain.ChainElements[$baseChain.ChainElements.Count - 1].Certificate
        $rootSubject = $root.Subject
        $rootThumbprint = $root.Thumbprint
    }}

    $onlineChain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
    $onlineChain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::Online
    $onlineChain.ChainPolicy.RevocationFlag = [System.Security.Cryptography.X509Certificates.X509RevocationFlag]::ExcludeRoot
    $onlineChain.ChainPolicy.VerificationFlags = [System.Security.Cryptography.X509Certificates.X509VerificationFlags]::NoFlag
    $onlineChain.ChainPolicy.UrlRetrievalTimeout = [TimeSpan]::FromSeconds(5)
    $onlineChainOk = $onlineChain.Build($sig.SignerCertificate)
    $onlineStatuses = @($onlineChain.ChainStatus | ForEach-Object {{ $_.Status.ToString() }})
    if ($onlineChainOk) {{
        $revocationState = 'checked'
    }} elseif ($baseChainOk -and $onlineStatuses.Count -gt 0) {{
        $offlineOnly = $true
        foreach ($status in $onlineStatuses) {{
            if ($status -notin @('RevocationStatusUnknown', 'OfflineRevocation')) {{
                $offlineOnly = $false
            }}
        }}
        if ($offlineOnly) {{ $revocationState = 'offline' }}
    }}
}}
[pscustomobject]@{{
    Status = "$($sig.Status)"
    StatusMessage = "$($sig.StatusMessage)"
    Signer = if ($sig.SignerCertificate) {{ $sig.SignerCertificate.Subject }} else {{ '' }}
    SignerThumbprint = if ($sig.SignerCertificate) {{ $sig.SignerCertificate.Thumbprint }} else {{ '' }}
    Root = $rootSubject
    RootThumbprint = $rootThumbprint
    ChainValid = $baseChainOk
    ChainStatus = @($baseStatuses + $onlineStatuses | Select-Object -Unique)
    RevocationState = $revocationState
}} | ConvertTo-Json -Compress -Depth 4
"""
        environment = os.environ.copy()
        environment["MSSTOREHELPER_PACKAGE_PATH"] = package_path
        result = subprocess.run(
            [POWERSHELL_EXE, "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
            env=environment,
            timeout=30,
        )
        if result.returncode != 0:
            error = (
                result.stderr.strip()
                or result.stdout.strip()
                or "Signature inspection failed"
            )
            raise PackageTrustError(error)
        try:
            signature_info = json.loads(result.stdout.strip() or "{}")
        except json.JSONDecodeError as exc:
            raise PackageTrustError(
                "Signature inspection returned invalid evidence"
            ) from exc
        if not isinstance(signature_info, dict):
            raise PackageTrustError(
                "Signature inspection returned invalid evidence"
            )
        signature_info["ChainStatus"] = normalize_chain_status(
            signature_info.get("ChainStatus")
        )
        return signature_info

    @staticmethod
    def inspect_package_trust(
        filepath,
        package=None,
        *,
        signature_info=None,
        evaluated_at=None,
    ):
        original_package = package
        package = validate_package_record(
            package or {
                "FileName": os.path.basename(os.path.abspath(filepath)),
            },
            require_url=False,
        )
        package = StoreAPI.attach_expected_trust_metadata(package)
        filepath = validate_existing_package_path(
            filepath,
            expected_filename=package["FileName"],
            require_file=True,
        )
        artifact_sha256 = StoreAPI.file_sha256(filepath)

        try:
            if signature_info is None:
                signature_info = StoreAPI.query_package_signature(filepath)
            manifest = read_package_manifest(
                filepath,
                package["FileName"],
            )
            report = evaluate_package_trust(
                package,
                artifact_sha256,
                signature_info,
                manifest,
                evaluated_at=evaluated_at,
            )
        except (
            OSError,
            PackageIngressError,
            PackageTrustError,
            subprocess.SubprocessError,
            TypeError,
        ) as exc:
            report = blocked_trust_report(
                package,
                artifact_sha256,
                str(exc),
                signature_info=signature_info,
                evaluated_at=evaluated_at,
            )

        if isinstance(package, dict):
            package["TrustState"] = report["State"]
            package["TrustReport"] = report
            package["Sha256"] = artifact_sha256
        if isinstance(original_package, dict):
            original_package.update(package)
        return report

    @staticmethod
    def package_trust_status(package, filepath, *, inspect_missing=True):
        package = package if isinstance(package, dict) else {}
        try:
            filepath = validate_existing_package_path(
                filepath,
                expected_filename=package.get("FileName"),
                require_file=True,
            )
            artifact_sha256 = StoreAPI.file_sha256(filepath)
        except (OSError, PackageIngressError, TypeError) as exc:
            return False, str(exc), None

        report = package.get("TrustReport") if isinstance(package, dict) else None
        if trust_report_allows_automation(report, artifact_sha256):
            state = report.get("State", "trusted")
            return True, f"Package trust state: {state}", report

        report_hash_matches = (
            isinstance(report, dict)
            and str(report.get("ArtifactSha256", "")).lower()
            == artifact_sha256.lower()
        )
        if inspect_missing and (
            not isinstance(report, dict) or not report_hash_matches
        ):
            report = StoreAPI.inspect_package_trust(filepath, package)

        state = (
            report.get("State", TRUST_STATE_BLOCKED)
            if isinstance(report, dict)
            else TRUST_STATE_BLOCKED
        )
        reasons = ", ".join(
            report.get("ReasonCodes", [])
            if isinstance(report, dict)
            else []
        )
        if state == TRUST_STATE_REVIEW_REQUIRED:
            message = "Package is quarantined pending interactive trust review"
        else:
            message = "Package trust checks blocked automation"
        if reasons:
            message = f"{message}: {reasons}"
        return False, message, report

    @staticmethod
    def review_package_trust(
        package,
        filepath,
        *,
        journal_path=TRUST_REVIEW_JOURNAL_PATH,
        reviewer="interactive-user",
        reviewed_at=None,
    ):
        filepath = validate_existing_package_path(
            filepath,
            expected_filename=package.get("FileName"),
            require_file=True,
        )
        artifact_sha256 = StoreAPI.file_sha256(filepath)
        reviewed = review_trust_report(
            package.get("TrustReport"),
            artifact_sha256,
            reviewer=reviewer,
            reviewed_at=reviewed_at,
        )
        event = {
            "Event": "package-trust-promotion",
            "RecordedAt": reviewed["Review"]["ReviewedAt"],
            "ArtifactSha256": artifact_sha256,
            "FileName": package["FileName"],
            "Source": reviewed.get("Source", {}),
            "Expected": reviewed.get("Expected", {}),
            "Manifest": reviewed.get("Manifest", {}),
            "Signature": reviewed.get("Signature", {}),
            "Review": reviewed.get("Review", {}),
        }
        os.makedirs(os.path.dirname(os.path.abspath(journal_path)), exist_ok=True)
        with open(journal_path, "a", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(event, sort_keys=True))
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())

        promoted_package = package.copy()
        promoted_package["TrustState"] = reviewed["State"]
        promoted_package["TrustReport"] = reviewed
        promoted_package["Sha256"] = artifact_sha256
        StoreAPI.write_artifact_manifest(
            promoted_package,
            filepath,
            os.path.dirname(filepath),
        )
        package.update(promoted_package)
        return reviewed

    @staticmethod
    def artifact_metadata(package, artifact_path, source_url=None):
        filename = validate_package_filename(
            package.get("FileName") or os.path.basename(artifact_path)
        )
        artifact_path = validate_existing_package_path(
            artifact_path,
            expected_filename=filename,
            require_file=True,
        )
        metadata_url = source_url or package.get("Url", "")
        if metadata_url:
            metadata_url = validate_package_url(metadata_url)
        metadata = {
            "FileName": filename,
            "SafeFileName": filename,
            "Path": artifact_path,
            "SizeBytes": os.path.getsize(artifact_path),
            "Sha256": StoreAPI.file_sha256(artifact_path),
            "Url": metadata_url,
            "PackageIdentity": package_identity(filename),
            "AvailableVersion": format_version_tuple(package_version_tuple(filename)),
        }
        for key in (
            "ExpectedProductId",
            "ExpectedPackageIdentity",
            "ExpectedPackageFamilyName",
            "ExpectedDependency",
            "TrustState",
            "TrustReport",
        ):
            if key in package:
                metadata[key] = package.get(key)
        if package.get("PackageRoleLabel"):
            metadata["PackageRoleLabel"] = package.get("PackageRoleLabel")
        if package.get("StoreQuery"):
            query = package.get("StoreQuery", {})
            metadata["StoreQuery"] = StoreAPI.store_query_settings(
                query.get("Ring"),
                query.get("Language"),
                query.get("Market"),
            )
            if query.get("ProductId"):
                metadata["StoreQuery"]["ProductId"] = str(
                    query["ProductId"]
                ).strip()
        return metadata

    @staticmethod
    def _manifest_path(folder):
        return os.path.join(folder, CACHE_MANIFEST_NAME)

    @staticmethod
    def load_cache_manifest(folder):
        manifest_path = StoreAPI._manifest_path(folder)
        try:
            with open(manifest_path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
            if isinstance(data, dict) and isinstance(data.get("Artifacts"), dict):
                data.setdefault("History", {})
                data.setdefault("Quarantine", {})
                return data
        except Exception:
            pass
        return {
            "Version": 2,
            "Artifacts": {},
            "History": {},
            "Quarantine": {},
        }

    @staticmethod
    def save_cache_manifest(folder, manifest):
        os.makedirs(folder, exist_ok=True)
        manifest.setdefault("Artifacts", {})
        manifest.setdefault("History", {})
        manifest.setdefault("Quarantine", {})
        manifest["Version"] = 2
        manifest["UpdatedAt"] = datetime.now(timezone.utc).isoformat()
        with open(StoreAPI._manifest_path(folder), "w", encoding="utf-8") as handle:
            json.dump(manifest, handle, indent=2, sort_keys=True)

    @staticmethod
    def _metadata_version_key(metadata):
        version = version_tuple_from_text(metadata.get("AvailableVersion"))
        if not version and metadata.get("FileName"):
            version = package_version_tuple(metadata["FileName"])
        padded = tuple(version[:5]) + (0,) * max(0, 5 - len(version))
        return padded[:5]

    @staticmethod
    def _path_is_inside_folder(path, folder):
        try:
            ensure_path_within_root(folder, path)
            return True
        except (OSError, PackageIngressError, TypeError, ValueError):
            return False

    @staticmethod
    def _remove_cache_artifacts(folder, metadata_items):
        for metadata in metadata_items:
            try:
                filename = validate_package_filename(metadata.get("FileName"))
                expected_path = confined_package_path(folder, filename)
                path = metadata.get("Path") or expected_path
                path = validate_existing_package_path(
                    path,
                    expected_filename=filename,
                    root=folder,
                )
                if os.path.normcase(path) != os.path.normcase(expected_path):
                    continue
            except (OSError, PackageIngressError, TypeError, ValueError):
                continue
            try:
                if os.path.exists(path):
                    os.remove(path)
            except OSError:
                pass

    @staticmethod
    def _record_cache_history(manifest, metadata, history_limit=CACHE_HISTORY_LIMIT):
        identity = str(metadata.get("PackageIdentity") or package_identity(metadata.get("FileName", ""))).lower()
        if not identity:
            return []

        history = manifest.setdefault("History", {})
        entries = [
            item for item in history.get(identity, [])
            if item.get("FileName") != metadata.get("FileName")
        ]
        entries.append(metadata.copy())
        entries.sort(
            key=lambda item: (StoreAPI._metadata_version_key(item), item.get("CachedAt", "")),
            reverse=True,
        )
        kept = entries[:history_limit]
        pruned = entries[history_limit:]
        history[identity] = kept

        for item in pruned:
            manifest["Artifacts"].pop(item.get("FileName", ""), None)
        return pruned

    @staticmethod
    def write_artifact_manifest(package, artifact_path, manifest_folder=None, source_url=None):
        folder = os.path.realpath(os.path.abspath(
            manifest_folder or os.path.dirname(os.path.abspath(artifact_path))
        ))
        filename = validate_package_filename(
            package.get("FileName") or os.path.basename(artifact_path)
        )
        artifact_path = validate_existing_package_path(
            artifact_path,
            expected_filename=filename,
            root=folder,
            require_file=True,
        )
        metadata = StoreAPI.artifact_metadata(package, artifact_path, source_url)
        metadata["CachedAt"] = datetime.now(timezone.utc).isoformat()
        manifest = StoreAPI.load_cache_manifest(folder)
        filename = metadata["FileName"]
        trusted = trust_report_allows_automation(
            metadata.get("TrustReport"),
            metadata["Sha256"],
        )
        pruned = []
        if trusted:
            manifest["Quarantine"].pop(filename, None)
            manifest["Artifacts"][filename] = metadata
            pruned = StoreAPI._record_cache_history(manifest, metadata)
        else:
            manifest["Artifacts"].pop(filename, None)
            manifest["Quarantine"][filename] = metadata
            identity = str(
                metadata.get("PackageIdentity")
                or package_identity(filename)
            ).lower()
            if identity in manifest["History"]:
                manifest["History"][identity] = [
                    item
                    for item in manifest["History"][identity]
                    if item.get("FileName") != filename
                ]
                if not manifest["History"][identity]:
                    manifest["History"].pop(identity, None)
        StoreAPI.save_cache_manifest(folder, manifest)
        StoreAPI._remove_cache_artifacts(folder, pruned)
        package["LocalPath"] = artifact_path
        package["SizeBytes"] = metadata["SizeBytes"]
        package["Sha256"] = metadata["Sha256"]
        package["CacheManifest"] = StoreAPI._manifest_path(folder)
        return metadata

    @staticmethod
    def redact_diagnostic_text(text):
        return redact_text(
            text,
            path_tokens={"APP_DATA": APP_DATA_DIR},
        )

    @staticmethod
    def redact_diagnostic_structure(value):
        return redact_structure(
            value,
            path_tokens={"APP_DATA": APP_DATA_DIR},
        )

    @staticmethod
    def diagnostic_queue_metadata(queue):
        allowed_keys = [
            "FileName", "PackageIdentity", "PackageRoleLabel", "Architecture", "FileType",
            "IsBundle", "IsEncrypted", "SizeBytes", "SizeStr", "Sha256", "AvailableVersion",
            "LocalPath", "CacheManifest", "StoreQuery", "DownloadStatus", "LastError",
            "ExpectedProductId", "ExpectedPackageIdentity",
            "ExpectedPackageFamilyName", "ExpectedDependency",
            "TrustState", "TrustReport",
            "UpdateSourceApp", "UpdateInstalledIdentity", "UpdateInstalledVersion",
            "UpdateAvailableVersion",
        ]
        items = []
        for package in queue or []:
            item = {key: package.get(key) for key in allowed_keys if key in package}
            items.append(StoreAPI.redact_diagnostic_structure(item))
        return items

    @staticmethod
    def download_state_queue_metadata(queue):
        allowed_keys = [
            "FileName", "SafeFileName", "Url", "Architecture", "FileType", "IsBundle", "IsEncrypted",
            "SizeBytes", "SizeStr", "Sha256", "AvailableVersion", "PackageRole",
            "PackageRoleLabel", "InstallOrder", "PackageIdentity", "LocalPath",
            "CacheManifest", "StoreQuery", "DownloadStatus", "LastError",
            "ExpectedProductId", "ExpectedPackageIdentity",
            "ExpectedPackageFamilyName", "ExpectedDependency",
            "TrustState", "TrustReport",
            "UpdateSourceApp", "UpdateInstalledIdentity", "UpdateInstalledVersion",
            "UpdateAvailableVersion",
        ]
        items = []
        for package in queue or []:
            try:
                package = validate_package_record(package, require_url=True)
            except PackageIngressError:
                continue
            item = {key: package.get(key) for key in allowed_keys if key in package}
            items.append(item)
        return items

    @staticmethod
    def write_download_state(queue, output_path, path=DOWNLOAD_STATE_PATH):
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        state = {
            "Version": 1,
            "UpdatedAt": datetime.now(timezone.utc).isoformat(),
            "OutputPath": os.path.abspath(output_path),
            "Queue": StoreAPI.download_state_queue_metadata(queue),
        }
        with open(path, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(state, handle, indent=2, sort_keys=True)
            handle.write("\n")
        return path

    @staticmethod
    def load_download_state(path=DOWNLOAD_STATE_PATH):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception:
            return {"Version": 1, "OutputPath": DEFAULT_OUTPUT, "Queue": []}

        queue = []
        for item in data.get("Queue", []):
            try:
                item = validate_package_record(item, require_url=True)
            except PackageIngressError:
                continue
            queue.append(annotate_package(item.copy()))
        return {
            "Version": 1,
            "OutputPath": data.get("OutputPath") or DEFAULT_OUTPUT,
            "Queue": queue,
            "UpdatedAt": data.get("UpdatedAt"),
        }

    @staticmethod
    def clear_download_state(path=DOWNLOAD_STATE_PATH):
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError:
            pass

    @staticmethod
    def collect_recent_repair_manifests(limit=5):
        if not os.path.isdir(REPAIR_BACKUP_DIR):
            return []

        manifests = []
        for root, _dirs, files in os.walk(REPAIR_BACKUP_DIR):
            if "repair-manifest.json" not in files:
                continue
            manifest_path = os.path.join(root, "repair-manifest.json")
            try:
                with open(manifest_path, "r", encoding="utf-8") as handle:
                    data = json.load(handle)
                manifests.append((os.path.getmtime(manifest_path), data))
            except Exception:
                continue

        recent = []
        for _mtime, data in sorted(manifests, reverse=True)[:limit]:
            recent.append(
                StoreAPI.redact_diagnostic_structure(data)
            )
        return recent

    @staticmethod
    def powershell_transcript_from_log(log_text):
        needles = ("powershell", "command:", "stdout:", "stderr:", "repair", "installing", "add-appxpackage", "dism")
        lines = [
            line for line in str(log_text or "").splitlines()
            if any(needle in line.lower() for needle in needles)
        ]
        return "\n".join(lines)

    @staticmethod
    def prepare_diagnostics_bundle(
        app_version,
        system_arch,
        is_admin,
        output_path,
        source_health,
        queue,
        log_text,
    ):
        queue_metadata = StoreAPI.diagnostic_queue_metadata(queue)
        repair_manifests = StoreAPI.collect_recent_repair_manifests()
        diagnostics = {
            "AppName": APP_NAME,
            "AppVersion": app_version,
            "GeneratedAt": datetime.now(timezone.utc).isoformat(),
            "Windows": {
                "Platform": platform.platform(),
                "Release": platform.release(),
                "Version": platform.version(),
                "Machine": platform.machine(),
            },
            "Python": {
                "Version": platform.python_version(),
                "Executable": sys.executable,
            },
            "SystemArchitecture": system_arch,
            "IsAdmin": bool(is_admin),
            "OutputPath": output_path,
            "SourceHealth": source_health or [],
            "QueueCount": len(queue_metadata),
            "RepairManifestCount": len(repair_manifests),
        }
        return prepare_diagnostic_entries(
            diagnostics=diagnostics,
            source_health=source_health or [],
            queue=queue_metadata,
            app_log=log_text,
            powershell_transcript=(
                StoreAPI.powershell_transcript_from_log(log_text)
            ),
            repair_manifests=repair_manifests,
            path_tokens={"APP_DATA": APP_DATA_DIR},
        )

    @staticmethod
    def write_diagnostics_bundle(
        bundle_path,
        app_version,
        system_arch,
        is_admin,
        output_path,
        source_health,
        queue,
        log_text,
    ):
        entries = StoreAPI.prepare_diagnostics_bundle(
            app_version,
            system_arch,
            is_admin,
            output_path,
            source_health,
            queue,
            log_text,
        )
        return write_prepared_bundle(bundle_path, entries)

    @staticmethod
    def cached_artifact_is_valid(path, metadata):
        if not os.path.exists(path) or not isinstance(metadata, dict):
            return False
        try:
            expected_size = int(metadata.get("SizeBytes", -1))
        except (TypeError, ValueError):
            return False
        expected_sha = str(metadata.get("Sha256", "")).lower()
        return (
            expected_size == os.path.getsize(path)
            and bool(expected_sha)
            and expected_sha == StoreAPI.file_sha256(path)
        )
    
    @staticmethod
    def download_file(url, filepath, progress_callback=None, package=None, destination_root=None):
        try:
            url = validate_package_url(url)
            validated_package = (
                validate_package_record(package, require_url=False)
                if package is not None
                else None
            )
            filename = validate_package_filename(
                validated_package.get("FileName")
                if validated_package is not None
                else os.path.basename(os.path.abspath(filepath))
            )
            destination_root = os.path.realpath(os.path.abspath(
                destination_root or os.path.dirname(os.path.abspath(filepath))
            ))
            os.makedirs(destination_root, exist_ok=True)
            expected_path = confined_package_path(destination_root, filename)
            supplied_path = ensure_path_within_root(destination_root, filepath)
            if os.path.normcase(supplied_path) != os.path.normcase(expected_path):
                raise PackageIngressError(
                    "Download destination does not match the validated package filename"
                )
            filepath = expected_path
            part_path = ensure_path_within_root(destination_root, f"{filepath}.part")

            if package is not None:
                package.update(validated_package)
            if package is not None and StoreAPI.cached_artifact_is_valid(filepath, package):
                package["LocalPath"] = filepath
                trust_ok, trust_message, _report = StoreAPI.package_trust_status(
                    package,
                    filepath,
                )
                if trust_ok:
                    return True, f"Already downloaded; {trust_message}"
                return False, trust_message

            existing = os.path.getsize(part_path) if os.path.exists(part_path) else 0
            headers = {"Range": f"bytes={existing}-"} if existing else None
            with requests.get(url, stream=True, timeout=60, headers=headers) as r:
                validate_response_redirects(url, r)
                r.raise_for_status()
                status_code = int(getattr(r, "status_code", 200) or 200)
                if existing and status_code != 206:
                    existing = 0

                content_range = r.headers.get("content-range", "")
                total = 0
                match = re.search(r"/(\d+)$", content_range)
                if match:
                    total = int(match.group(1))
                elif r.headers.get('content-length'):
                    total = int(r.headers.get('content-length', 0)) + existing
                
                mode = "ab" if existing else "wb"
                with open(part_path, mode) as f:
                    downloaded = existing
                    for chunk in r.iter_content(chunk_size=8192):
                        if not chunk:
                            continue
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total:
                            progress_callback(downloaded / total)
                if total and downloaded != total:
                    return False, f"Downloaded {downloaded} bytes; expected {total} bytes"

            os.replace(part_path, filepath)
            trust_package = package or {
                "FileName": filename,
                "Url": url,
            }
            report = StoreAPI.inspect_package_trust(
                filepath,
                trust_package,
            )
            StoreAPI.write_artifact_manifest(
                trust_package,
                filepath,
                source_url=url,
            )
            if report.get("AutomationAllowed") is True:
                return True, (
                    "Success; package trust verified"
                    if report.get("State") == "trusted"
                    else "Success; reviewed package trust verified"
                )
            if report.get("State") == TRUST_STATE_REVIEW_REQUIRED:
                return False, (
                    "Downloaded package is quarantined pending interactive "
                    "trust review"
                )
            reasons = ", ".join(report.get("ReasonCodes", []))
            return False, (
                f"Downloaded package failed trust checks: {reasons}"
                if reasons
                else "Downloaded package failed trust checks"
            )
        except Exception as e:
            return False, str(e)

    @staticmethod
    def is_cacheable_artifact(filename):
        return os.path.splitext(filename)[1].lower() in {".appx", ".msix", ".appxbundle", ".msixbundle"}

    @staticmethod
    def cache_downloaded_artifact(package, cache_path):
        original_package = package
        raw_filename = (
            package.get("FileName", os.path.basename(package.get("LocalPath") or ""))
            if isinstance(package, dict)
            else ""
        )
        if not StoreAPI.is_cacheable_artifact(raw_filename):
            return False, "File type is not cacheable"
        try:
            package = validate_package_record(package, require_url=False)
            local_path = validate_existing_package_path(
                package.get("LocalPath"),
                expected_filename=package["FileName"],
                require_file=True,
            )
        except PackageIngressError as exc:
            return False, str(exc)
        filename = package["FileName"]
        if not os.path.exists(local_path):
            return False, "Downloaded file is missing"
        trust_ok, trust_message, _report = StoreAPI.package_trust_status(
            package,
            local_path,
        )
        if isinstance(original_package, dict):
            original_package.update(package)
        if not trust_ok:
            return False, trust_message
        try:
            os.makedirs(cache_path, exist_ok=True)
            destination = confined_package_path(cache_path, filename)
        except (OSError, PackageIngressError) as exc:
            return False, str(exc)
        manifest = StoreAPI.load_cache_manifest(cache_path)
        existing_metadata = manifest["Artifacts"].get(filename)
        if StoreAPI.cached_artifact_is_valid(destination, existing_metadata):
            return True, f"Already cached: {destination}"

        source_metadata = StoreAPI.artifact_metadata(package, local_path)
        if os.path.exists(destination):
            destination_metadata = {
                "SizeBytes": os.path.getsize(destination),
                "Sha256": StoreAPI.file_sha256(destination),
            }
            if StoreAPI.cached_artifact_is_valid(destination, destination_metadata) and destination_metadata["Sha256"] == source_metadata["Sha256"]:
                manifest["Artifacts"][filename] = {**source_metadata, "Path": os.path.abspath(destination)}
                StoreAPI.save_cache_manifest(cache_path, manifest)
                return True, f"Already cached: {destination}"

        shutil.copy2(local_path, destination)
        StoreAPI.write_artifact_manifest(package, destination, cache_path)
        if isinstance(original_package, dict):
            original_package.update(package)
        return True, f"Cached: {destination}"

    @staticmethod
    def mirror_package_records(
        cache_folder,
        advertised_host="127.0.0.1",
        port=8765,
        *,
        tls_enabled=False,
    ):
        folder = os.path.abspath(cache_folder)
        if not os.path.isdir(folder):
            return []

        manifest = StoreAPI.load_cache_manifest(folder)
        artifacts = manifest.get("Artifacts", {})
        records = []
        seen = set()
        base_url = mirror_base_url(
            advertised_host,
            port,
            tls_enabled=tls_enabled,
        )

        for filename in sorted(os.listdir(folder), key=str.lower):
            try:
                filename = validate_package_filename(filename)
                path = confined_package_path(folder, filename)
                path = validate_existing_package_path(
                    path,
                    expected_filename=filename,
                    root=folder,
                    require_file=True,
                )
            except PackageIngressError:
                continue
            if not StoreAPI.is_cacheable_artifact(filename):
                continue
            metadata = artifacts.get(filename, {}).copy()
            if not metadata or not StoreAPI.cached_artifact_is_valid(path, metadata):
                continue
            trust_ok, _trust_message, _report = StoreAPI.package_trust_status(
                metadata,
                path,
                inspect_missing=False,
            )
            if not trust_ok:
                continue

            key = filename.lower()
            if key in seen:
                continue
            seen.add(key)
            records.append({
                "FileName": filename,
                "Url": f"{base_url}/packages/{quote(filename, safe='')}",
                "SizeBytes": os.path.getsize(path),
                "SizeStr": metadata.get("SizeStr") or format_size(os.path.getsize(path)),
                "Sha256": metadata.get("Sha256") or StoreAPI.file_sha256(path),
                "PackageIdentity": metadata.get("PackageIdentity") or package_identity(filename),
                "AvailableVersion": metadata.get("AvailableVersion") or format_version_tuple(package_version_tuple(filename)),
                "Architecture": metadata.get("Architecture", "neutral"),
                "FileType": metadata.get("FileType", os.path.splitext(filename)[1].lower().lstrip(".").upper()),
                "CachedAt": metadata.get("CachedAt") or metadata.get("DownloadedAt") or "",
            })

        return records

    @staticmethod
    def build_mirror_index(
        cache_folder,
        advertised_host="127.0.0.1",
        port=8765,
        *,
        tls_enabled=False,
        requires_authorization=False,
        token_expires_at=None,
    ):
        records = StoreAPI.mirror_package_records(
            cache_folder,
            advertised_host,
            port,
            tls_enabled=tls_enabled,
        )
        base_url = mirror_base_url(
            advertised_host,
            port,
            tls_enabled=tls_enabled,
        )
        return {
            "SchemaVersion": 2,
            "AppName": APP_NAME,
            "AppVersion": APP_VERSION,
            "GeneratedAt": datetime.now(timezone.utc).isoformat(),
            "BaseUrl": base_url,
            "IndexUrl": f"{base_url}/{MIRROR_INDEX_NAME}",
            "IndexFile": MIRROR_INDEX_NAME,
            "Authorization": {
                "Required": bool(requires_authorization),
                "Scheme": (
                    "Bearer" if requires_authorization else "None"
                ),
                "ExpiresAt": (
                    str(token_expires_at)
                    if requires_authorization
                    else ""
                ),
            },
            "PackageCount": len(records),
            "Packages": records,
        }

    @staticmethod
    def write_mirror_index(
        cache_folder,
        advertised_host="127.0.0.1",
        port=8765,
        *,
        tls_enabled=False,
        requires_authorization=False,
        token_expires_at=None,
        index_name=MIRROR_INDEX_NAME,
    ):
        if index_name != MIRROR_INDEX_NAME:
            raise MirrorConfigurationError(
                "Mirror index filename is fixed by the route allowlist"
            )
        os.makedirs(cache_folder, exist_ok=True)
        index = StoreAPI.build_mirror_index(
            cache_folder,
            advertised_host,
            port,
            tls_enabled=tls_enabled,
            requires_authorization=requires_authorization,
            token_expires_at=token_expires_at,
        )
        path = os.path.join(cache_folder, index_name)
        atomic_write_json(path, index)
        return index

    @staticmethod
    def _mirror_package_routes(cache_folder, index):
        folder = os.path.abspath(cache_folder)
        routes = {}
        for package in index.get("Packages", []):
            filename = validate_package_filename(package["FileName"])
            path = confined_package_path(folder, filename)
            path = validate_existing_package_path(
                path,
                expected_filename=filename,
                root=folder,
                require_file=True,
            )
            route = urlsplit(package["Url"]).path
            routes[route] = {
                "Path": path,
                "SizeBytes": int(package["SizeBytes"]),
                "Sha256": str(package["Sha256"]),
            }
        return routes

    @staticmethod
    def mirror_http_handler(
        cache_folder,
        advertised_host="127.0.0.1",
        port=8765,
        *,
        bearer_token=None,
        token_expires_at=None,
        tls_enabled=False,
        audit_log_path=None,
    ):
        folder = os.path.abspath(cache_folder)
        os.makedirs(folder, exist_ok=True)
        index = StoreAPI.build_mirror_index(
            folder,
            advertised_host,
            port,
            tls_enabled=tls_enabled,
            requires_authorization=bool(bearer_token),
            token_expires_at=token_expires_at,
        )
        routes = StoreAPI._mirror_package_routes(folder, index)
        audit = MirrorAuditLog(
            audit_log_path
            or os.path.join(folder, MIRROR_AUDIT_FILENAME)
        )
        return make_mirror_handler(
            index_name=MIRROR_INDEX_NAME,
            index_payload=index,
            package_routes=routes,
            app_version=APP_VERSION,
            audit_log=audit,
            bearer_token=bearer_token,
            token_expires_at=token_expires_at,
        )

    @staticmethod
    def create_mirror_server(
        cache_folder,
        bind_host="127.0.0.1",
        port=8765,
        *,
        advertised_host=None,
        lan_mode=False,
        acknowledge_cleartext=False,
        tls_cert=None,
        tls_key=None,
        token_ttl_seconds=900,
        bearer_token=None,
        audit_log_path=None,
    ):
        policy = validate_network_policy(
            bind_host,
            advertised_host=advertised_host,
            lan_mode=lan_mode,
            acknowledge_cleartext=acknowledge_cleartext,
            tls_cert=tls_cert,
            tls_key=tls_key,
        )
        folder = os.path.abspath(cache_folder)
        os.makedirs(folder, exist_ok=True)
        audit = MirrorAuditLog(
            audit_log_path
            or os.path.join(folder, MIRROR_AUDIT_FILENAME)
        )
        placeholder = make_mirror_handler(
            index_name=MIRROR_INDEX_NAME,
            index_payload={"Status": "initializing"},
            package_routes={},
            app_version=APP_VERSION,
            audit_log=audit,
        )
        server = ThreadingHTTPServer(
            (policy["BindHost"], int(port)),
            placeholder,
        )
        actual_port = int(server.server_address[1])
        requires_authorization = not policy["Loopback"]
        token = ""
        expires_epoch = 0
        expires_at = ""
        if requires_authorization:
            token = str(bearer_token or create_bearer_token())
            ttl = normalize_token_ttl(token_ttl_seconds)
            expires_epoch = time.time() + ttl
            expires_at = mirror_utc_timestamp(
                datetime.fromtimestamp(
                    expires_epoch,
                    tz=timezone.utc,
                )
            )
        index = StoreAPI.write_mirror_index(
            folder,
            policy["AdvertisedHost"],
            actual_port,
            tls_enabled=policy["TlsEnabled"],
            requires_authorization=requires_authorization,
            token_expires_at=expires_at,
        )
        routes = StoreAPI._mirror_package_routes(folder, index)
        server.RequestHandlerClass = make_mirror_handler(
            index_name=MIRROR_INDEX_NAME,
            index_payload=index,
            package_routes=routes,
            app_version=APP_VERSION,
            audit_log=audit,
            bearer_token=token,
            token_expires_at=expires_epoch,
        )
        if policy["TlsEnabled"]:
            wrap_server_tls(
                server,
                policy["TlsCert"],
                policy["TlsKey"],
            )
        server.mirror_policy = policy
        server.mirror_bearer_token = token
        server.mirror_token_expires_at = expires_at
        server.mirror_audit_path = audit.path
        return server, index

    @staticmethod
    def cache_history_entries(cache_folders, package_identities=None):
        wanted = {
            str(identity).strip().lower()
            for identity in (package_identities or [])
            if str(identity).strip()
        }
        entries = []
        seen = set()

        for folder in cache_folders or []:
            if not folder or not os.path.isdir(folder):
                continue
            manifest = StoreAPI.load_cache_manifest(folder)
            history = manifest.get("History") or {}
            if not history:
                for metadata in manifest.get("Artifacts", {}).values():
                    identity = str(metadata.get("PackageIdentity") or package_identity(metadata.get("FileName", ""))).lower()
                    if identity:
                        history.setdefault(identity, []).append(metadata)

            for identity, items in history.items():
                identity = str(identity).strip().lower()
                if wanted and identity not in wanted:
                    continue
                for item in items or []:
                    if not isinstance(item, dict) or not item.get("FileName"):
                        continue
                    try:
                        metadata = validate_package_record(item, require_url=False)
                        expected_path = confined_package_path(folder, metadata["FileName"])
                        metadata_path = metadata.get("Path") or expected_path
                        metadata_path = validate_existing_package_path(
                            metadata_path,
                            expected_filename=metadata["FileName"],
                            root=folder,
                            require_file=True,
                        )
                        if os.path.normcase(metadata_path) != os.path.normcase(expected_path):
                            continue
                    except PackageIngressError:
                        continue
                    metadata["PackageIdentity"] = str(metadata.get("PackageIdentity") or identity)
                    metadata["Path"] = metadata_path
                    metadata["CacheFolder"] = folder
                    key = (metadata["PackageIdentity"].lower(), metadata["FileName"].lower(), os.path.abspath(metadata["Path"]).lower())
                    trust_ok, _trust_message, _report = (
                        StoreAPI.package_trust_status(
                            metadata,
                            metadata["Path"],
                            inspect_missing=False,
                        )
                    )
                    if (
                        key in seen
                        or not StoreAPI.cached_artifact_is_valid(
                            metadata["Path"],
                            metadata,
                        )
                        or not trust_ok
                    ):
                        continue
                    seen.add(key)
                    entries.append(metadata)
        return entries

    @staticmethod
    def rollback_candidates(cache_folders, package_identities=None, current_versions=None):
        versions = {
            str(identity).strip().lower(): str(version).strip()
            for identity, version in (current_versions or {}).items()
            if str(identity).strip() and str(version).strip()
        }
        grouped = {}
        for entry in StoreAPI.cache_history_entries(cache_folders, package_identities):
            identity = str(entry.get("PackageIdentity", "")).lower()
            grouped.setdefault(identity, []).append(entry)

        candidates = []
        for identity, entries in grouped.items():
            entries.sort(
                key=lambda item: (StoreAPI._metadata_version_key(item), item.get("CachedAt", "")),
                reverse=True,
            )
            current_version = versions.get(identity)
            selected = None

            if current_version:
                current_tuple = version_tuple_from_text(current_version)
                for entry in entries:
                    if compare_version_tuples(StoreAPI._metadata_version_key(entry), current_tuple) < 0:
                        selected = entry
                        break
            elif len(entries) > 1:
                selected = entries[1]

            if not selected:
                continue

            candidate = selected.copy()
            candidate["RollbackIdentity"] = identity
            candidate["RollbackVersion"] = (
                candidate.get("AvailableVersion")
                or format_version_tuple(package_version_tuple(candidate["FileName"]))
            )
            candidate["RollbackCurrentVersion"] = current_version or ""
            candidates.append(candidate)

        candidates.sort(key=lambda item: item.get("RollbackIdentity", ""))
        return candidates

    @staticmethod
    def rollback_package(package_identity_name, artifact_path, package=None):
        try:
            package_path = validate_existing_package_path(
                artifact_path,
                require_file=True,
            )
        except PackageIngressError as exc:
            return False, str(exc)
        identity = str(package_identity_name or "").strip()
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]{0,127}", identity):
            return False, "Rollback package identity is invalid"
        trust_package = package or {
            "FileName": os.path.basename(package_path),
            "ExpectedPackageIdentity": identity,
        }
        trust_ok, trust_message, _report = StoreAPI.package_trust_status(
            trust_package,
            package_path,
        )
        if not trust_ok:
            return False, trust_message

        cmd = "\n".join([
            "$ErrorActionPreference = 'Stop'",
            "$identity = $env:MSSTOREHELPER_ROLLBACK_IDENTITY",
            "$path = $env:MSSTOREHELPER_PACKAGE_PATH",
            "$current = Get-AppxPackage -Name $identity | Sort-Object Version -Descending | Select-Object -First 1",
            "if ($current) {",
            "    try {",
            "        Remove-AppxPackage -Package $current.PackageFullName -PreserveApplicationData -ErrorAction Stop",
            "    } catch {",
            "        if ($_.Exception.Message -match 'PreserveApplicationData' -or $_.Exception.Message -match 'parameter') {",
            "            Remove-AppxPackage -Package $current.PackageFullName -ErrorAction Stop",
            "        } else {",
            "            throw",
            "        }",
            "    }",
            "}",
            "Add-AppxPackage -Path $path -ErrorAction Stop",
            "Write-Output \"Rollback installed $identity from $path\"",
        ])

        try:
            environment = os.environ.copy()
            environment["MSSTOREHELPER_PACKAGE_PATH"] = package_path
            environment["MSSTOREHELPER_ROLLBACK_IDENTITY"] = identity
            result = subprocess.run(
                [POWERSHELL_EXE, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                env=environment,
            )
            if result.returncode == 0:
                return True, result.stdout.strip() or "Rollback installed"
            return False, result.stderr.strip() or result.stdout.strip() or "Rollback failed"
        except Exception as exc:
            return False, str(exc)

    @staticmethod
    def _powershell_literal(value):
        return "'" + str(value).replace("'", "''") + "'"

    @staticmethod
    def _package_store_query(package):
        query = package.get("StoreQuery") or {}
        return StoreAPI.store_query_settings(
            query.get("Ring"),
            query.get("Language"),
            query.get("Market"),
        )

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
            if not package.get("FileName"):
                continue
            if str(package["FileName"]).lower().endswith(".blockmap"):
                continue
            package = validate_package_record(package, require_url=False)
            if is_installable_package(package):
                provisionable.append(annotate_package(package))

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
            package_path = StoreAPI._queue_package_source_path(package, output_path)
            portable_path = StoreAPI._portable_script_path(package_path, script_dir)
            role = package.get("PackageRoleLabel") or package_role_label(filename)
            query = StoreAPI._package_store_query(package)
            lines.append(
                "    [pscustomobject]@{ "
                f"FileName = {StoreAPI._powershell_literal(filename)}; "
                f"Role = {StoreAPI._powershell_literal(role)}; "
                f"PackagePath = {StoreAPI._powershell_literal(portable_path)}; "
                f"StoreRing = {StoreAPI._powershell_literal(query['Ring'])}; "
                f"StoreLanguage = {StoreAPI._powershell_literal(query['Language'])}; "
                f"StoreMarket = {StoreAPI._powershell_literal(query['Market'])} "
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
            "    Write-Host (\"Provisioning {0} [{1}] from {2}/{3}/{4}\" -f $package.FileName, $package.Role, $package.StoreRing, $package.StoreLanguage, $package.StoreMarket)",
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
    def _xml_local_name(tag):
        return tag.rsplit("}", 1)[-1] if "}" in tag else tag

    @staticmethod
    def _read_appx_manifest_root(package_path):
        manifest_names = (
            "AppxManifest.xml",
            "AppxMetadata/AppxBundleManifest.xml",
        )
        try:
            with zipfile.ZipFile(package_path) as archive:
                available = {name.replace("\\", "/"): name for name in archive.namelist()}
                manifest_name = next((available[name] for name in manifest_names if name in available), None)
                if not manifest_name:
                    raise ValueError("AppX manifest was not found")
                return ET.fromstring(archive.read(manifest_name)), manifest_name
        except zipfile.BadZipFile as exc:
            raise ValueError(f"Package is not a readable AppX/MSIX archive: {package_path}") from exc

    @staticmethod
    def _appx_identity_from_root(root, package_path):
        identity = None
        for element in root.iter():
            if StoreAPI._xml_local_name(element.tag) == "Identity":
                identity = element
                break
        if identity is None:
            raise ValueError(f"Package identity was not found: {package_path}")

        name = identity.attrib.get("Name", "").strip()
        publisher = identity.attrib.get("Publisher", "").strip()
        version = identity.attrib.get("Version", "").strip()
        architecture = identity.attrib.get("ProcessorArchitecture", "").strip().lower()
        if not name or not publisher or not version:
            raise ValueError(f"Package identity is incomplete: {package_path}")

        return {
            "Name": name,
            "Publisher": publisher,
            "Version": version,
            "ProcessorArchitecture": architecture or "neutral",
        }

    @staticmethod
    def read_appx_identity(package_path):
        root, _manifest_name = StoreAPI._read_appx_manifest_root(package_path)
        return StoreAPI._appx_identity_from_root(root, package_path)

    @staticmethod
    def _dependency_label(element):
        local_name = StoreAPI._xml_local_name(element.tag)
        attrs = element.attrib
        if local_name == "PackageDependency":
            name = attrs.get("Name", "").strip()
            min_version = attrs.get("MinVersion", "").strip()
            publisher = attrs.get("Publisher", "").strip()
            label = f"{name} >= {min_version}" if min_version else name
            if publisher:
                label = f"{label} ({publisher})"
            return label.strip()
        if local_name == "TargetDeviceFamily":
            name = attrs.get("Name", "").strip()
            minimum = attrs.get("MinVersion", "").strip()
            tested = attrs.get("MaxVersionTested", "").strip()
            pieces = [name]
            if minimum:
                pieces.append(f"min {minimum}")
            if tested:
                pieces.append(f"tested {tested}")
            return " ".join(piece for piece in pieces if piece).strip()
        return ""

    @staticmethod
    def read_appx_manifest_details(package_path):
        root, manifest_name = StoreAPI._read_appx_manifest_root(package_path)
        identity = StoreAPI._appx_identity_from_root(root, package_path)
        capabilities = set()
        dependencies = set()

        for container in root.iter():
            container_name = StoreAPI._xml_local_name(container.tag)
            if container_name == "Capabilities":
                for element in container:
                    local_name = StoreAPI._xml_local_name(element.tag)
                    if local_name in {"Capability", "DeviceCapability", "CustomCapability"}:
                        name = element.attrib.get("Name", "").strip()
                        if name:
                            capabilities.add(f"{local_name}: {name}")
            elif container_name == "Dependencies":
                for element in container:
                    local_name = StoreAPI._xml_local_name(element.tag)
                    if local_name in {"PackageDependency", "TargetDeviceFamily"}:
                        label = StoreAPI._dependency_label(element)
                        if label:
                            dependencies.add(f"{local_name}: {label}")

        return {
            "Path": os.path.abspath(package_path),
            "ManifestName": manifest_name,
            "Identity": identity,
            "Capabilities": sorted(capabilities),
            "Dependencies": sorted(dependencies),
        }

    @staticmethod
    def _set_diff(old_values, new_values):
        old_set = set(old_values or [])
        new_set = set(new_values or [])
        return {
            "Added": sorted(new_set - old_set),
            "Removed": sorted(old_set - new_set),
            "Unchanged": sorted(old_set & new_set),
        }

    @staticmethod
    def diff_appx_manifests(old_package_path, new_package_path):
        old_details = StoreAPI.read_appx_manifest_details(old_package_path)
        new_details = StoreAPI.read_appx_manifest_details(new_package_path)
        return {
            "Old": old_details,
            "New": new_details,
            "IdentityChanged": old_details["Identity"]["Name"] != new_details["Identity"]["Name"],
            "VersionChanged": old_details["Identity"]["Version"] != new_details["Identity"]["Version"],
            "Capabilities": StoreAPI._set_diff(old_details["Capabilities"], new_details["Capabilities"]),
            "Dependencies": StoreAPI._set_diff(old_details["Dependencies"], new_details["Dependencies"]),
        }

    @staticmethod
    def package_diff_candidates(cache_folders, package_identities=None):
        grouped = {}
        for entry in StoreAPI.cache_history_entries(cache_folders, package_identities):
            identity = str(entry.get("PackageIdentity", "")).lower()
            grouped.setdefault(identity, []).append(entry)

        candidates = []
        for identity, entries in grouped.items():
            entries.sort(
                key=lambda item: (StoreAPI._metadata_version_key(item), item.get("CachedAt", "")),
                reverse=True,
            )
            if len(entries) < 2:
                continue
            newest, previous = entries[0], entries[1]
            candidates.append({
                "PackageIdentity": identity,
                "Old": previous,
                "New": newest,
            })
        candidates.sort(key=lambda item: item["PackageIdentity"])
        return candidates

    @staticmethod
    def format_package_diff(diff):
        old_identity = diff["Old"]["Identity"]
        new_identity = diff["New"]["Identity"]
        lines = [
            f"{new_identity['Name']}",
            f"Old: {old_identity['Version']} ({diff['Old']['ManifestName']})",
            f"New: {new_identity['Version']} ({diff['New']['ManifestName']})",
            "",
        ]
        if diff.get("IdentityChanged"):
            lines.append(f"Identity changed: {old_identity['Name']} -> {new_identity['Name']}")
            lines.append("")

        for section in ("Capabilities", "Dependencies"):
            lines.append(section)
            changes = diff[section]
            changed = False
            for label, key in (("Added", "Added"), ("Removed", "Removed")):
                for item in changes[key]:
                    lines.append(f"  {label}: {item}")
                    changed = True
            if not changed:
                lines.append("  No changes")
            if changes["Unchanged"]:
                lines.append(f"  Unchanged: {len(changes['Unchanged'])}")
            lines.append("")
        return "\n".join(lines).strip()

    @staticmethod
    def _file_uri(path):
        return Path(os.path.abspath(path)).as_uri()

    @staticmethod
    def _appinstaller_package_tag(record, main=False):
        if record["IsBundle"]:
            return "MainBundle" if main else "Bundle"
        return "MainPackage" if main else "Package"

    @staticmethod
    def _appinstaller_record(package, output_path, package_dir):
        source_path = StoreAPI._queue_package_source_path(package, output_path)
        if not os.path.exists(source_path):
            raise ValueError(f"Downloaded file is missing: {package.get('FileName')}")

        filename = validate_package_filename(
            package.get("FileName") or os.path.basename(source_path)
        )
        destination = confined_package_path(package_dir, filename)
        if os.path.abspath(source_path).lower() != os.path.abspath(destination).lower():
            shutil.copy2(source_path, destination)

        identity = StoreAPI.read_appx_identity(destination)
        is_bundle = filename.lower().endswith(("appxbundle", "msixbundle"))
        architecture = identity.get("ProcessorArchitecture") or package.get("Architecture") or "neutral"
        return {
            "FileName": filename,
            "Path": destination,
            "Uri": StoreAPI._file_uri(destination),
            "Name": identity["Name"],
            "Publisher": identity["Publisher"],
            "Version": identity["Version"],
            "ProcessorArchitecture": architecture.lower(),
            "IsBundle": is_bundle,
            "Role": package.get("PackageRoleLabel") or package_role_label(filename),
            "StoreQuery": StoreAPI._package_store_query(package),
        }

    @staticmethod
    def generate_appinstaller_manifest(records, appinstaller_path, hours_between_update_checks=12):
        if not records:
            raise ValueError("No AppX/MSIX packages are available for App Installer export")

        apps = [record for record in records if record["Role"] == "App"]
        if not apps:
            raise ValueError("App Installer export requires at least one app package")

        main = apps[0]
        dependencies = [record for record in records if record["Role"] != "App"]
        optional = apps[1:]

        ET.register_namespace("", APPINSTALLER_NS)
        root = ET.Element(
            f"{{{APPINSTALLER_NS}}}AppInstaller",
            {
                "Version": main["Version"],
                "Uri": StoreAPI._file_uri(appinstaller_path),
            },
        )

        def add_package(parent, record, main_package=False):
            attributes = {
                "Name": record["Name"],
                "Publisher": record["Publisher"],
                "Version": record["Version"],
                "Uri": record["Uri"],
            }
            if not record["IsBundle"]:
                attributes["ProcessorArchitecture"] = record["ProcessorArchitecture"]
            ET.SubElement(parent, f"{{{APPINSTALLER_NS}}}{StoreAPI._appinstaller_package_tag(record, main_package)}", attributes)

        add_package(root, main, main_package=True)

        if dependencies:
            deps = ET.SubElement(root, f"{{{APPINSTALLER_NS}}}Dependencies")
            for record in dependencies:
                add_package(deps, record)

        if optional:
            optional_packages = ET.SubElement(root, f"{{{APPINSTALLER_NS}}}OptionalPackages")
            for record in optional:
                add_package(optional_packages, record)

        update_settings = ET.SubElement(root, f"{{{APPINSTALLER_NS}}}UpdateSettings")
        ET.SubElement(
            update_settings,
            f"{{{APPINSTALLER_NS}}}OnLaunch",
            {
                "HoursBetweenUpdateChecks": str(int(hours_between_update_checks)),
                "ShowPrompt": "true",
                "UpdateBlocksActivation": "false",
            },
        )

        return ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8")

    @staticmethod
    def write_appinstaller_export(packages, output_path, appinstaller_path, target_arch=SYSTEM_ARCH):
        installable = []
        for package in packages:
            if not package.get("FileName"):
                continue
            if str(package["FileName"]).lower().endswith(".blockmap"):
                continue
            package = validate_package_record(package, require_url=False)
            if is_installable_package(package):
                installable.append(annotate_package(package))
        installable = StoreAPI.order_packages_for_install(installable, target_arch)
        if not installable:
            raise ValueError("No AppX/MSIX packages are available for App Installer export")

        appinstaller_path = os.path.abspath(appinstaller_path)
        export_dir = os.path.dirname(appinstaller_path)
        os.makedirs(export_dir, exist_ok=True)
        package_dir = ensure_path_within_root(
            export_dir,
            os.path.join(
                export_dir,
                f"{StoreAPI._safe_filename_stem(appinstaller_path, 'MSStoreHelper-AppInstaller')}-Packages",
            ),
        )
        if os.path.isdir(package_dir):
            if os.path.commonpath([export_dir, os.path.abspath(package_dir)]) != export_dir:
                raise ValueError("Invalid App Installer package path")
            shutil.rmtree(package_dir)
        os.makedirs(package_dir, exist_ok=True)

        records = [StoreAPI._appinstaller_record(package, output_path, package_dir) for package in installable]
        manifest = StoreAPI.generate_appinstaller_manifest(records, appinstaller_path)
        with open(appinstaller_path, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(manifest)
            handle.write("\n")

        return {
            "AppInstallerPath": appinstaller_path,
            "PackageDir": package_dir,
            "PackageCount": len(records),
            "MainPackage": next(record["Name"] for record in records if record["Role"] == "App"),
        }

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
        package = validate_package_record(package, require_url=False)
        filename = package["FileName"]
        if package.get("LocalPath"):
            source_path = validate_existing_package_path(
                package["LocalPath"],
                expected_filename=filename,
                require_file=True,
            )
        else:
            source_path = validate_existing_package_path(
                confined_package_path(output_path, filename),
                expected_filename=filename,
                root=output_path,
                require_file=True,
            )
        trust_ok, trust_message, _report = StoreAPI.package_trust_status(
            package,
            source_path,
        )
        if not trust_ok:
            raise ValueError(trust_message)
        return source_path

    @staticmethod
    def _intune_package_records(packages, output_path, target_arch=SYSTEM_ARCH):
        provisionable = []
        for package in packages:
            if not package.get("FileName"):
                continue
            if str(package["FileName"]).lower().endswith(".blockmap"):
                continue
            package = validate_package_record(package, require_url=False)
            if is_installable_package(package):
                provisionable.append(annotate_package(package))

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
                "StoreQuery": StoreAPI._package_store_query(package),
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
            query = record["StoreQuery"]
            lines.append(
                "    [pscustomobject]@{ "
                f"FileName = {StoreAPI._powershell_literal(record['FileName'])}; "
                f"Role = {StoreAPI._powershell_literal(record['Role'])}; "
                f"StoreRing = {StoreAPI._powershell_literal(query['Ring'])}; "
                f"StoreLanguage = {StoreAPI._powershell_literal(query['Language'])}; "
                f"StoreMarket = {StoreAPI._powershell_literal(query['Market'])} "
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
            "    Write-Host (\"Provisioning {0} [{1}] from {2}/{3}/{4}\" -f $package.FileName, $package.Role, $package.StoreRing, $package.StoreLanguage, $package.StoreMarket)",
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
        source_dir = ensure_path_within_root(
            staging_root,
            os.path.join(staging_root, safe_basename),
        )
        packages_dir = ensure_path_within_root(
            source_dir,
            os.path.join(source_dir, "Packages"),
        )

        if os.path.exists(source_dir):
            if os.path.commonpath([staging_root, os.path.abspath(source_dir)]) != staging_root:
                raise ValueError("Invalid Intune staging path")
            shutil.rmtree(source_dir)
        os.makedirs(packages_dir, exist_ok=True)

        for record in records:
            shutil.copy2(
                record["SourcePath"],
                confined_package_path(packages_dir, record["FileName"]),
            )

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
            queries = {
                f"{record['StoreQuery']['Ring']}/{record['StoreQuery']['Language']}/{record['StoreQuery']['Market']}"
                for record in records
            }
            handle.write(f"Store query: {', '.join(sorted(queries))}\n")

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
    def install_package(filepath, package=None):
        try:
            package_path = validate_existing_package_path(
                filepath,
                require_file=True,
            )
            trust_package = package or {
                "FileName": os.path.basename(package_path),
            }
            trust_ok, trust_message, _report = StoreAPI.package_trust_status(
                trust_package,
                package_path,
            )
            if not trust_ok:
                return False, trust_message
            cmd = "\n".join([
                "$ErrorActionPreference = 'Stop'",
                "$path = $env:MSSTOREHELPER_PACKAGE_PATH",
                "if ([string]::IsNullOrWhiteSpace($path)) { throw 'Package path is missing' }",
                "Add-AppxPackage -Path $path -ErrorAction Stop",
            ])
            environment = os.environ.copy()
            environment["MSSTOREHELPER_PACKAGE_PATH"] = package_path
            result = subprocess.run(
                [POWERSHELL_EXE, "-NoProfile", "-Command", cmd],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                env=environment,
            )
            
            if result.returncode != 0:
                error_msg = result.stderr.strip() or result.stdout.strip() or "Unknown error"
                return False, error_msg
            
            return True, "Installed successfully"
        except Exception as e:
            return False, str(e)

    @staticmethod
    def verify_package_signature(filepath, package=None):
        try:
            package_path = validate_existing_package_path(
                filepath,
                require_file=True,
            )
            trust_package = package or {
                "FileName": os.path.basename(package_path),
            }
            trust_ok, trust_message, _report = StoreAPI.package_trust_status(
                trust_package,
                package_path,
            )
            return trust_ok, trust_message
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
    def normalize_installed_appx_versions(records):
        if isinstance(records, dict):
            records = [records]

        versions = {}
        for record in records or []:
            if not isinstance(record, dict):
                continue
            name = record.get("Name") or record.get("DisplayName") or record.get("PackageIdentity")
            version = record.get("Version") or record.get("PackageVersion")
            key = str(name or "").strip().lower()
            version_text = str(version or "").strip()
            if not key or not version_text:
                continue

            current = versions.get(key)
            if not current or compare_version_tuples(
                version_tuple_from_text(version_text),
                version_tuple_from_text(current),
            ) > 0:
                versions[key] = version_text
        return versions

    @staticmethod
    def get_installed_appx_versions():
        cmd = (
            "$installed = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | "
            "Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Version';Expression={[string]$_.Version}}); "
            "$provisioned = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | "
            "Select-Object @{Name='Name';Expression={$_.DisplayName}}, @{Name='Version';Expression={[string]$_.Version}}); "
            "@($installed + $provisioned | Where-Object { $_.Name }) | ConvertTo-Json -Compress"
        )
        try:
            result = subprocess.run(
                [POWERSHELL_EXE, "-NoProfile", "-Command", cmd],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if result.returncode != 0 or not result.stdout.strip():
                return {}

            payload = json.loads(result.stdout)
            if isinstance(payload, dict):
                payload = [payload]
            return StoreAPI.normalize_installed_appx_versions(payload)
        except Exception:
            return {}

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
    def select_catalog_update_packages(catalog, installed_versions, package_lookup, target_arch, prefer_exact_arch=False):
        installed = {
            str(identity).strip().lower(): str(version).strip()
            for identity, version in (installed_versions or {}).items()
            if str(identity).strip() and str(version).strip()
        }
        updates = []
        seen_filenames = set()

        for category in (catalog or {}).values():
            for app in category.get("apps", []):
                packages = package_lookup(app) or []
                recommended = StoreAPI.smart_select(packages, target_arch, prefer_exact_arch)
                update_anchor = None

                for package in recommended:
                    if is_dependency_package(package) or not package.get("FileName"):
                        continue
                    identity = (package.get("PackageIdentity") or package_identity(package["FileName"])).lower()
                    installed_version = installed.get(identity)
                    if installed_version and not installed_version_satisfies_package(package, installed_version):
                        update_anchor = (identity, installed_version, package)
                        break

                if not update_anchor:
                    continue

                update_identity, installed_version, app_package = update_anchor
                available_version = (
                    app_package.get("AvailableVersion")
                    or format_version_tuple(package_version_tuple(app_package["FileName"]))
                )

                for package in recommended:
                    if not package.get("FileName"):
                        continue
                    identity = (package.get("PackageIdentity") or package_identity(package["FileName"])).lower()
                    package_installed_version = installed.get(identity)
                    if package_installed_version and installed_version_satisfies_package(package, package_installed_version):
                        continue

                    filename_key = package["FileName"].lower()
                    if filename_key in seen_filenames:
                        continue

                    update_package = annotate_package(package.copy())
                    update_package["UpdateSourceApp"] = app.get("Name", "")
                    update_package["UpdateInstalledIdentity"] = update_identity
                    update_package["UpdateInstalledVersion"] = installed_version
                    update_package["UpdateAvailableVersion"] = available_version
                    updates.append(update_package)
                    seen_filenames.add(filename_key)

        return order_packages_for_install(updates, target_arch)

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
    def _safe_repair_name(name):
        return re.sub(r"[^A-Za-z0-9_.-]+", "-", str(name or "repair")).strip("-") or "repair"

    @staticmethod
    def create_repair_context(repair_name, backup_root=None):
        safe_name = StoreAPI._safe_repair_name(repair_name)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        root = os.path.join(backup_root or REPAIR_BACKUP_DIR, f"{safe_name}-{stamp}")
        os.makedirs(root, exist_ok=True)
        context = {
            "RepairName": repair_name,
            "StartedAt": datetime.now(timezone.utc).isoformat(),
            "BackupRoot": os.path.abspath(root),
            "ManifestPath": os.path.join(root, "repair-manifest.json"),
            "BackupLogPath": os.path.join(root, "backup-records.jsonl"),
            "RestoreScriptPath": os.path.join(root, "restore.ps1"),
            "Steps": [],
            "Results": [],
        }
        StoreAPI.write_repair_restore_script(context)
        StoreAPI.write_repair_manifest(context)
        return context

    @staticmethod
    def write_repair_manifest(context):
        os.makedirs(context["BackupRoot"], exist_ok=True)
        payload = {
            "RepairName": context.get("RepairName"),
            "StartedAt": context.get("StartedAt"),
            "CompletedAt": context.get("CompletedAt"),
            "BackupRoot": context.get("BackupRoot"),
            "BackupLogPath": context.get("BackupLogPath"),
            "RestoreScriptPath": context.get("RestoreScriptPath"),
            "Steps": context.get("Steps", []),
            "Results": context.get("Results", []),
        }
        with open(context["ManifestPath"], "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)

    @staticmethod
    def write_repair_restore_script(context):
        script = r'''# Generated by MSStoreHelper repair runner.
$ErrorActionPreference = 'Continue'
$BackupRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BackupLog = Join-Path $BackupRoot 'backup-records.jsonl'
if (-not (Test-Path -LiteralPath $BackupLog)) {
    Write-Warning "No backup records were found at $BackupLog"
    return
}
Get-Content -LiteralPath $BackupLog | Where-Object { $_.Trim() } | ForEach-Object {
    $record = $_ | ConvertFrom-Json
    if ($record.Type -eq 'Registry') {
        if (Test-Path -LiteralPath $record.BackupPath) {
            & reg.exe import $record.BackupPath | Out-Host
        }
        return
    }
    if (-not (Test-Path -LiteralPath $record.BackupPath)) {
        Write-Warning "Missing backup: $($record.BackupPath)"
        return
    }
    $parent = Split-Path -Parent $record.OriginalPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    if (Test-Path -LiteralPath $record.OriginalPath) {
        $existing = "$($record.OriginalPath).pre-restore-$(Get-Date -Format yyyyMMdd-HHmmss)"
        Move-Item -LiteralPath $record.OriginalPath -Destination $existing -Force
    }
    Move-Item -LiteralPath $record.BackupPath -Destination $record.OriginalPath -Force
}
'''
        with open(context["RestoreScriptPath"], "w", encoding="utf-8") as handle:
            handle.write(script)

    @staticmethod
    def _repair_powershell_prelude(context):
        if not context:
            return ""

        backup_root = StoreAPI._powershell_literal(context["BackupRoot"])
        backup_log = StoreAPI._powershell_literal(context["BackupLogPath"])
        return rf'''
$MSStoreHelperBackupRoot = {backup_root}
$MSStoreHelperBackupLog = {backup_log}
New-Item -ItemType Directory -Path $MSStoreHelperBackupRoot -Force | Out-Null
function Write-MSStoreHelperBackupRecord {{
    param([hashtable]$Record)
    $Record.Timestamp = (Get-Date).ToUniversalTime().ToString("o")
    ($Record | ConvertTo-Json -Compress) | Add-Content -LiteralPath $MSStoreHelperBackupLog -Encoding UTF8
}}
function Backup-MSStoreHelperPath {{
    param([string]$Path)
    $items = @(Get-Item -Path $Path -Force -ErrorAction SilentlyContinue)
    foreach ($item in $items) {{
        if (-not $item) {{ continue }}
        $safeName = ($item.FullName -replace '[:\\\/]+', '_')
        $destination = Join-Path $MSStoreHelperBackupRoot $safeName
        if (Test-Path -LiteralPath $destination) {{
            $destination = "$destination-$(Get-Date -Format yyyyMMddHHmmssfff)"
        }}
        Move-Item -LiteralPath $item.FullName -Destination $destination -Force -ErrorAction Stop
        Write-MSStoreHelperBackupRecord @{{ Type = "FileSystem"; OriginalPath = $item.FullName; BackupPath = $destination }}
    }}
}}
function Backup-MSStoreHelperRegistryPath {{
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {{ return }}
    $safeName = ($Path -replace '[:\\\/]+', '_') + ".reg"
    $destination = Join-Path $MSStoreHelperBackupRoot $safeName
    $regPath = $Path `
        -replace '^Microsoft\.PowerShell\.Core\\Registry::HKEY_LOCAL_MACHINE\\', 'HKLM\' `
        -replace '^HKEY_LOCAL_MACHINE\\', 'HKLM\' `
        -replace '^HKLM:\\', 'HKLM\'
    & reg.exe export $regPath $destination /y | Out-Null
    if ($LASTEXITCODE -eq 0) {{
        Write-MSStoreHelperBackupRecord @{{ Type = "Registry"; OriginalPath = $Path; BackupPath = $destination }}
    }}
}}
'''
    
    @staticmethod
    def get_store_repair_steps():
        return [
            ("🔧 Starting Windows Update...", 'Start-Service -Name wuauserv -ErrorAction SilentlyContinue'),
            ("🔧 Starting BITS...", 'Start-Service -Name bits -ErrorAction SilentlyContinue'),
            ("🔐 Starting licensing services...", 'Start-Service -Name ClipSVC -ErrorAction SilentlyContinue; Start-Service -Name LicenseManager -ErrorAction SilentlyContinue'),
            ("🧹 Closing Store broker processes...", 'Get-Process WinStore.App,MicrosoftStore,RuntimeBroker -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue'),
            ("🧹 Resetting Store cache...", 'Start-Process wsreset.exe -WindowStyle Hidden -Wait'),
            ("🧹 Rebuilding Store token cache...", r'$paths = @("$env:LOCALAPPDATA\Microsoft\TokenBroker\Cache\*", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache\*", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\AC\TokenBroker\Cache\*", "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Cache\*"); foreach ($path in $paths) { Backup-MSStoreHelperPath -Path $path; Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }'),
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
            ("🧹 Clearing Store deprovision tombstones...", r'$root = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Deprovisioned"; $patterns = @("*Microsoft.WindowsStore*", "*Microsoft.StorePurchaseApp*", "*Microsoft.DesktopAppInstaller*"); foreach ($pattern in $patterns) { Get-ChildItem $root -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like $pattern } | ForEach-Object { Backup-MSStoreHelperRegistryPath -Path $_.PSPath; Remove-Item -LiteralPath $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue } }'),
            ("🔄 Re-registering Store apps for existing users...", r'@("Microsoft.WindowsStore", "Microsoft.StorePurchaseApp", "Microsoft.DesktopAppInstaller") | ForEach-Object { Get-AppxPackage -AllUsers $_ -ErrorAction SilentlyContinue | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue } }'),
            ("📋 Checking provisioned Store catalog...", 'Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -in @("Microsoft.WindowsStore", "Microsoft.StorePurchaseApp", "Microsoft.DesktopAppInstaller") } | Out-Null'),
        ]

    @staticmethod
    def get_licensing_reset_steps():
        return [
            ("🔐 Stopping licensing services...", 'Stop-Service -Name LicenseManager -Force -ErrorAction SilentlyContinue; Stop-Service -Name ClipSVC -Force -ErrorAction SilentlyContinue'),
            ("🧹 Clearing ClipSVC license cache...", r'$paths = @("$env:ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\*", "$env:ProgramData\Microsoft\Windows\ClipSVC\Tokens\*"); foreach ($path in $paths) { Backup-MSStoreHelperPath -Path $path; Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }'),
            ("🔐 Starting licensing services...", 'Start-Service -Name ClipSVC -ErrorAction SilentlyContinue; Start-Service -Name LicenseManager -ErrorAction SilentlyContinue'),
            ("🔄 Re-registering Store licensing app...", r'@("Microsoft.StorePurchaseApp", "Microsoft.WindowsStore") | ForEach-Object { Get-AppxPackage -AllUsers $_ -ErrorAction SilentlyContinue | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue } }'),
        ]

    @staticmethod
    def get_cache_rebuild_steps():
        return [
            ("🧹 Closing Store cache owners...", 'Get-Process WinStore.App,MicrosoftStore,RuntimeBroker -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue'),
            ("🔎 Scanning Store cache folders...", r'$paths = @("$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\AC\INetCache", "$env:LOCALAPPDATA\Packages\Microsoft.StorePurchaseApp_8wekyb3d8bbwe\LocalCache"); foreach ($path in $paths) { if (Test-Path $path) { Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Length -eq 0 } | Measure-Object | Out-Null } }'),
            ("📦 Backing up existing Store caches...", r'$paths = @("$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\AC\INetCache", "$env:LOCALAPPDATA\Packages\Microsoft.StorePurchaseApp_8wekyb3d8bbwe\LocalCache"); foreach ($path in $paths) { if (Test-Path $path) { Backup-MSStoreHelperPath -Path $path } }'),
            ("🔄 Rebuilding clean Store cache folders...", r'$paths = @("$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache", "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\AC\INetCache", "$env:LOCALAPPDATA\Packages\Microsoft.StorePurchaseApp_8wekyb3d8bbwe\LocalCache"); foreach ($path in $paths) { New-Item -ItemType Directory -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }'),
            ("🧹 Running wsreset after offline rebuild...", 'Start-Process wsreset.exe -WindowStyle Hidden -Wait'),
        ]

    @staticmethod
    def _run_powershell_steps(steps, log_callback=None, progress_callback=None, timeout=90, repair_name=None, backup_root=None):
        raise RepairTransactionError(
            "Legacy best-effort repair execution is disabled; "
            "inspect and explicitly confirm a repair transaction instead"
        )
        context = StoreAPI.create_repair_context(repair_name or "repair", backup_root) if repair_name else None
        if context:
            context["Steps"] = [
                {"Description": desc, "Command": cmd}
                for desc, cmd in steps
            ]
            StoreAPI.write_repair_manifest(context)

        results = []
        for i, (desc, cmd) in enumerate(steps):
            if log_callback:
                log_callback(desc)
            try:
                command = StoreAPI._repair_powershell_prelude(context) + "\n" + cmd if context else cmd
                result = subprocess.run(
                    [POWERSHELL_EXE, "-NoProfile", "-Command", command],
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    timeout=timeout
                )
                step_result = {
                    "Description": desc,
                    "Command": cmd,
                    "Success": result.returncode == 0,
                    "ReturnCode": result.returncode,
                    "Stdout": result.stdout.strip(),
                    "Stderr": result.stderr.strip(),
                }
            except Exception as exc:
                step_result = {
                    "Description": desc,
                    "Command": cmd,
                    "Success": False,
                    "ReturnCode": None,
                    "Stdout": "",
                    "Stderr": str(exc),
                }

            if context:
                step_result["BackupRoot"] = context["BackupRoot"]
                step_result["ManifestPath"] = context["ManifestPath"]
                step_result["RestoreScriptPath"] = context["RestoreScriptPath"]
                context["Results"].append(step_result)
                StoreAPI.write_repair_manifest(context)
            results.append(step_result)
            
            if progress_callback:
                progress_callback((i + 1) / len(steps))

        if context:
            context["CompletedAt"] = datetime.now(timezone.utc).isoformat()
            StoreAPI.write_repair_manifest(context)
        
        return results

    @staticmethod
    def run_repair(log_callback=None, progress_callback=None):
        return StoreAPI._run_powershell_steps(StoreAPI.get_store_repair_steps(), log_callback, progress_callback, repair_name="store-repair")

    @staticmethod
    def run_provisioning_repair(log_callback=None, progress_callback=None):
        return StoreAPI._run_powershell_steps(StoreAPI.get_provisioning_repair_steps(), log_callback, progress_callback, repair_name="provisioning-repair")

    @staticmethod
    def run_licensing_reset(log_callback=None, progress_callback=None):
        return StoreAPI._run_powershell_steps(StoreAPI.get_licensing_reset_steps(), log_callback, progress_callback, repair_name="licensing-reset")

    @staticmethod
    def run_cache_rebuild(log_callback=None, progress_callback=None):
        return StoreAPI._run_powershell_steps(StoreAPI.get_cache_rebuild_steps(), log_callback, progress_callback, repair_name="cache-rebuild")

# ==================== UI COMPONENTS ====================

class ModernCard(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color=Theme.BG_CARD, corner_radius=8, border_width=1, border_color=Theme.BORDER_SUBTLE, **kwargs)


class AppTile(ctk.CTkFrame):
    def __init__(self, master, app_data, on_select, on_release_notes):
        super().__init__(master, fg_color="transparent")
        
        self.app_data = app_data
        self.on_select = on_select
        self.on_release_notes = on_release_notes
        self.selected = ctk.BooleanVar(value=False)
        
        self.container = ctk.CTkFrame(self, fg_color=Theme.BG_CARD, corner_radius=6, border_width=1, border_color=Theme.BORDER_SUBTLE)
        self.container.pack(fill="x", pady=3, padx=1)
        self.container.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.container, text=app_data.get("Icon", "📦"), font=("Segoe UI Emoji", 24), width=50).grid(row=0, column=0, rowspan=2, padx=(15, 10), pady=12)
        ctk.CTkLabel(self.container, text=app_data["Name"], font=("Segoe UI Semibold", 14), anchor="w").grid(row=0, column=1, sticky="sw", padx=5, pady=(12, 0))
        ctk.CTkLabel(self.container, text=app_data.get("Description", ""), font=("Segoe UI", 11), text_color=Theme.TEXT_SECONDARY, anchor="w", justify="left", wraplength=200).grid(row=1, column=1, sticky="nw", padx=5, pady=(0, 12))
        
        btn_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        btn_frame.grid(row=0, column=2, rowspan=2, padx=10)

        ctk.CTkButton(btn_frame, text="Notes", width=58, height=28, font=("Segoe UI", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=lambda: self.on_release_notes(self.app_data)).pack(side="left", padx=3)
        self.chk = ctk.CTkCheckBox(btn_frame, text="Select", variable=self.selected, width=58, font=("Segoe UI", 10), command=self._toggle, fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER)
        self.chk.pack(side="left", padx=5)
        
        self.container.bind("<Enter>", lambda e: self.container.configure(fg_color=Theme.BG_CARD_HOVER))
        self.container.bind("<Leave>", lambda e: self.container.configure(fg_color=Theme.BG_CARD))
    
    def _toggle(self):
        self.on_select(self.app_data, self.selected.get())


class SearchResultTile(ctk.CTkFrame):
    def __init__(self, master, app_data, on_fetch, on_select, on_release_notes):
        super().__init__(master, fg_color="transparent")
        
        self.app_data = app_data
        self.selected = ctk.BooleanVar(value=False)
        
        self.container = ctk.CTkFrame(self, fg_color=Theme.BG_CARD, corner_radius=6, border_width=1, border_color=Theme.BORDER_SUBTLE)
        self.container.pack(fill="x", pady=3, padx=1)
        self.container.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.container, text="📱", font=("Segoe UI Emoji", 20), width=45).grid(row=0, column=0, rowspan=2, padx=(15, 10), pady=10)
        ctk.CTkLabel(self.container, text=app_data.get("Name", "Unknown"), font=("Segoe UI Semibold", 13), anchor="w").grid(row=0, column=1, sticky="sw", padx=5, pady=(10, 0))
        
        info = f"{app_data.get('Publisher', '')}  •  {app_data.get('ProductId', '')}"
        ctk.CTkLabel(self.container, text=info, font=("Consolas", 10), text_color=Theme.TEXT_MUTED, anchor="w", justify="left", wraplength=200).grid(row=1, column=1, sticky="nw", padx=5, pady=(0, 10))
        
        btn_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        btn_frame.grid(row=0, column=2, rowspan=2, padx=10)
        
        ctk.CTkButton(btn_frame, text="Get Files", width=80, height=30, font=("Segoe UI", 12), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=lambda: on_fetch(app_data)).pack(side="left", padx=3)
        ctk.CTkButton(btn_frame, text="Notes", width=58, height=30, font=("Segoe UI", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=lambda: on_release_notes(app_data)).pack(side="left", padx=3)
        ctk.CTkCheckBox(btn_frame, text="Select", variable=self.selected, width=58, font=("Segoe UI", 10), command=lambda: on_select(app_data, self.selected.get()), fg_color=Theme.PRIMARY).pack(side="left", padx=5)
        
        self.container.bind("<Enter>", lambda e: self.container.configure(fg_color=Theme.BG_CARD_HOVER))
        self.container.bind("<Leave>", lambda e: self.container.configure(fg_color=Theme.BG_CARD))


class PackageRow(ctk.CTkFrame):
    def __init__(self, master, pkg_data, on_toggle, index, target_arch):
        super().__init__(master, fg_color=Theme.BG_CARD if index % 2 == 0 else "transparent", corner_radius=6)
        
        self.pkg_data = pkg_data
        self.selected = ctk.BooleanVar(value=False)
        
        self.grid_columnconfigure(1, weight=1)
        
        self.chk = ctk.CTkCheckBox(self, text="Select", variable=self.selected, width=58, font=("Segoe UI", 10), command=lambda: on_toggle(pkg_data, self.selected.get()), fg_color=Theme.PRIMARY)
        self.chk.grid(row=0, column=0, padx=(12, 8), pady=10)
        
        info_frame = ctk.CTkFrame(self, fg_color="transparent")
        info_frame.grid(row=0, column=1, sticky="ew", padx=5, pady=8)
        info_frame.grid_columnconfigure(0, weight=1)
        
        name_color = Theme.ENCRYPTED_COLOR if pkg_data.get('IsEncrypted') else Theme.TEXT_PRIMARY
        ctk.CTkLabel(info_frame, text=pkg_data['FileName'], font=("Consolas", 11), text_color=name_color, anchor="w", wraplength=220, justify="left").grid(row=0, column=0, sticky="w")
        
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
    def __init__(self, master, pkg_info, review_callback=None):
        super().__init__(master, fg_color=Theme.BG_ELEVATED, corner_radius=5, border_width=1, border_color=Theme.BORDER_SUBTLE)
        
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

        status = pkg_info.get("DownloadStatus") or "Pending"
        status_colors = {
            "Downloaded": Theme.SUCCESS,
            "Quarantined": Theme.WARNING,
            "TrustBlocked": Theme.DANGER,
            "Partial": Theme.WARNING,
            "Failed": Theme.DANGER,
            "Downloading": Theme.INFO,
            "Pending": Theme.TEXT_MUTED,
        }
        status_text = {
            "Downloaded": "✅ Done",
            "Quarantined": "Review required",
            "TrustBlocked": "Trust blocked",
            "Partial": "Partial",
            "Failed": "❌ Failed",
            "Downloading": "Downloading...",
            "Pending": "Waiting",
        }.get(status, "Waiting")
        self.status_lbl = ctk.CTkLabel(info_frame, text=status_text, font=("Segoe UI", 10), text_color=status_colors.get(status, Theme.TEXT_MUTED))
        self.status_lbl.pack(side="right")

        if (
            review_callback
            and pkg_info.get("TrustState") == TRUST_STATE_REVIEW_REQUIRED
            and pkg_info.get("LocalPath")
        ):
            ctk.CTkButton(
                info_frame,
                text="Review",
                width=58,
                height=24,
                font=("Segoe UI Semibold", 10),
                fg_color="transparent",
                text_color=Theme.WARNING,
                border_width=1,
                border_color=Theme.WARNING,
                hover_color=Theme.BG_CARD_HOVER,
                command=lambda: review_callback(pkg_info),
            ).pack(side="right", padx=(0, 8))
        
        pkg_info['_status_widget'] = self.status_lbl


# ==================== MAIN APPLICATION ====================

class MSStoreHelperApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title(f"📦 {APP_NAME} v{APP_VERSION}")
        self.geometry("1280x800")
        self.minsize(1000, 600)
        
        self.user_profile = StoreAPI.load_user_profile()
        self.theme_mode_var = ctk.StringVar(value=Theme.normalize_mode(self.user_profile.get("ThemeMode", "System")))
        self.store_ring_var = ctk.StringVar(value=StoreAPI.normalize_store_ring(self.user_profile.get("StoreRing", "Retail")))
        self.store_language_var = ctk.StringVar(value=StoreAPI.normalize_store_language(self.user_profile.get("StoreLanguage", "en-US")))
        self.store_market_var = ctk.StringVar(value=StoreAPI.normalize_store_market(self.user_profile.get("StoreMarket", "US")))
        self.keep_updated_var = ctk.BooleanVar(value=bool(self.user_profile.get("KeepUpdatedEnabled", False)))
        Theme.set_mode(self.theme_mode_var.get())
        ctk.set_appearance_mode(Theme.MODE)
        ctk.set_default_color_theme("dark-blue")
        self.configure(fg_color=Theme.BG_DARK)
        
        self.selected_apps = []
        self.current_packages = []
        self.selected_packages = set()
        self.package_rows = []
        self.current_view = "welcome"
        self.source_health = []
        self.arch_options = [f"Auto ({SYSTEM_ARCH})", "x64", "x86", "arm64", "arm", "neutral"]
        self.arch_override_var = ctk.StringVar(value=self.arch_options[0])
        self.package_scroll = None
        download_state = StoreAPI.load_download_state()
        self.output_path = download_state.get("OutputPath") or DEFAULT_OUTPUT
        self.download_queue = download_state.get("Queue", [])
        self.shared_cache_enabled = ctk.BooleanVar(value=False)
        self.shared_cache_path = os.path.join(DEFAULT_OUTPUT, "SharedCache")
        self.keep_updated_after_id = None
        self.keep_updated_running = False
        self.repair_retention_var = ctk.StringVar(
            value=str(
                normalize_retention(
                    self.user_profile.get("RepairRetentionCount")
                )
            )
        )
        self._repair_operation_active = False
        self._repair_cancel_event = None
        self._repair_buttons = []
        
        self._build_ui()
        if self.download_queue:
            self._update_queue_ui()
        self._show_welcome()
        if self.download_queue:
            self._log("INFO", f"Restored {len(self.download_queue)} queued download(s) from previous session")
        if os.environ.get("MSSTOREHELPER_SKIP_SOURCE_HEALTH") != "1":
            threading.Thread(target=self._source_health_worker, daemon=True).start()
        if self.keep_updated_var.get():
            self._schedule_keep_updated_scan(KEEP_UPDATED_START_DELAY_MS)

    def _target_arch(self):
        choice = self.arch_override_var.get()
        if choice.startswith("Auto"):
            return SYSTEM_ARCH
        return choice.lower()

    def _has_arch_override(self):
        return not self.arch_override_var.get().startswith("Auto")
    
    def _build_ui(self):
        # ACTIVITY PANEL (packed first so it consistently owns the bottom edge)
        self.log_panel = ctk.CTkFrame(self, fg_color=Theme.BG_CARD, corner_radius=0)
        self._build_log_panel()
        
        # CONTINUOUS THREE-ZONE WORKSPACE
        self.main = ctk.CTkFrame(self, fg_color=Theme.BG_DARK, corner_radius=0)
        self.main.pack(fill="both", expand=True)
        self.main.grid_columnconfigure(0, minsize=242)
        self.main.grid_columnconfigure(1, weight=1, minsize=0)
        self.main.grid_columnconfigure(2, minsize=324)
        self.main.grid_rowconfigure(0, weight=1)
        
        # NAVIGATION RAIL
        self.sidebar = ctk.CTkFrame(
            self.main,
            fg_color=Theme.BG_SIDEBAR,
            width=242,
            corner_radius=0,
            border_width=1,
            border_color=Theme.BORDER_SUBTLE,
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.pack_propagate(False)
        self._build_sidebar()

        # CENTRAL WORKSPACE
        self.content = ctk.CTkFrame(self.main, fg_color=Theme.BG_DARK, corner_radius=0)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.pack_propagate(False)

        # QUEUE INSPECTOR
        self.right_panel = ctk.CTkFrame(
            self.main,
            fg_color=Theme.BG_SIDEBAR,
            width=324,
            corner_radius=0,
            border_width=1,
            border_color=Theme.BORDER_SUBTLE,
        )
        self.right_panel.grid(row=0, column=2, sticky="nsew")
        self.right_panel.grid_propagate(False)
        self._build_queue_panel()

    def _build_store_query_controls(self, parent):
        self.store_query_button = ctk.CTkButton(
            parent,
            text=self._store_query_button_text(),
            height=28,
            font=("Segoe UI", 10),
            fg_color="transparent",
            text_color=Theme.TEXT_SECONDARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._show_store_query_dialog,
        )
        self.store_query_button.pack(fill="x", pady=(6, 0))
    
    def _build_sidebar(self):
        brand = ctk.CTkFrame(self.sidebar, fg_color="transparent", height=72, corner_radius=0)
        brand.pack(fill="x")
        brand.pack_propagate(False)
        ctk.CTkLabel(
            brand,
            text="▣",
            width=30,
            font=("Segoe UI Symbol", 24),
            text_color=Theme.PRIMARY,
        ).pack(side="left", padx=(16, 8))
        brand_text = ctk.CTkFrame(brand, fg_color="transparent")
        brand_text.pack(side="left", fill="y", pady=(16, 10))
        ctk.CTkLabel(
            brand_text,
            text=APP_NAME,
            font=("Segoe UI Semibold", 16),
            text_color=Theme.TEXT_PRIMARY,
            anchor="w",
        ).pack(fill="x")
        ctk.CTkLabel(
            brand_text,
            text=f"Store package workspace · v{APP_VERSION}",
            font=("Segoe UI", 9),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
        ).pack(fill="x")

        footer = ctk.CTkFrame(self.sidebar, fg_color="transparent", corner_radius=0)
        footer.pack(fill="x", side="bottom", padx=12, pady=12)
        theme_row = ctk.CTkFrame(footer, fg_color="transparent")
        theme_row.pack(fill="x")
        ctk.CTkOptionMenu(
            theme_row,
            values=THEME_MODE_VALUES,
            variable=self.theme_mode_var,
            width=150,
            height=32,
            font=("Segoe UI", 11),
            fg_color=Theme.BG_INPUT,
            text_color=Theme.TEXT_PRIMARY,
            button_color=Theme.PRIMARY,
            button_hover_color=Theme.PRIMARY_HOVER,
            command=self._change_theme_mode,
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        ctk.CTkButton(
            theme_row,
            text="?",
            width=34,
            height=32,
            font=("Segoe UI Semibold", 13),
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._show_help,
        ).pack(side="right")
        admin_text = f"{'Admin' if IS_ADMIN else 'Standard'} · {SYSTEM_ARCH}"
        ctk.CTkLabel(
            footer,
            text=admin_text,
            font=("Segoe UI", 10),
            text_color=Theme.SUCCESS if IS_ADMIN else Theme.TEXT_MUTED,
            anchor="w",
        ).pack(fill="x", pady=(8, 0))

        nav = ctk.CTkScrollableFrame(
            self.sidebar,
            fg_color="transparent",
            corner_radius=0,
            scrollbar_button_color=Theme.BORDER,
            scrollbar_button_hover_color=Theme.TEXT_MUTED,
        )
        nav.pack(fill="both", expand=True, padx=(8, 4), pady=(0, 2))

        self.workspace_button = ctk.CTkButton(
            nav,
            text="⌂   Workspace",
            height=36,
            font=("Segoe UI Semibold", 12),
            fg_color=Theme.BG_CARD_HOVER,
            text_color=Theme.PRIMARY_OUTLINE_TEXT,
            hover_color=Theme.BG_CARD_HOVER,
            anchor="w",
            command=self._show_welcome,
        )
        self.workspace_button.pack(fill="x", padx=4, pady=(0, 14))

        search_section = ctk.CTkFrame(nav, fg_color="transparent")
        search_section.pack(fill="x", padx=4, pady=(0, 12))
        ctk.CTkLabel(
            search_section,
            text="FIND APPS",
            font=("Segoe UI Semibold", 10),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
        ).pack(fill="x", pady=(0, 6))
        self.search_entry = ctk.CTkEntry(
            search_section,
            placeholder_text="Search Microsoft Store",
            height=36,
            font=("Segoe UI", 12),
            fg_color=Theme.BG_INPUT,
            border_color=Theme.BORDER,
        )
        self.search_entry.pack(fill="x", pady=(0, 6))
        self.search_entry.bind("<Return>", lambda e: self._do_search())
        ctk.CTkButton(
            search_section,
            text="Search Store",
            height=34,
            font=("Segoe UI Semibold", 12),
            fg_color=Theme.PRIMARY,
            hover_color=Theme.PRIMARY_HOVER,
            command=self._do_search,
        ).pack(fill="x")
        self.search_history_frame = ctk.CTkFrame(search_section, fg_color="transparent", height=1)
        self._render_search_history()
        self._build_store_query_controls(search_section)

        ctk.CTkFrame(nav, fg_color=Theme.BORDER_SUBTLE, height=1).pack(fill="x", padx=4, pady=(0, 12))

        fix_section = ctk.CTkFrame(nav, fg_color="transparent")
        fix_section.pack(fill="x", padx=4, pady=(0, 12))
        ctk.CTkLabel(
            fix_section,
            text="QUICK ACTIONS",
            font=("Segoe UI Semibold", 10),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
        ).pack(fill="x", pady=(0, 6))
        
        self.quickfix_var = ctk.StringVar(value=list(QUICK_FIX_PRESETS.keys())[0])
        ctk.CTkOptionMenu(
            fix_section,
            values=list(QUICK_FIX_PRESETS.keys()),
            variable=self.quickfix_var,
            height=34,
            font=("Segoe UI", 11),
            fg_color=Theme.BG_INPUT,
            text_color=Theme.TEXT_PRIMARY,
            button_color=Theme.PRIMARY,
            button_hover_color=Theme.PRIMARY_HOVER,
            command=self._update_quickfix_desc,
        ).pack(fill="x", pady=(0, 6))
        
        self.quickfix_desc = ctk.CTkLabel(fix_section, text=QUICK_FIX_PRESETS[self.quickfix_var.get()]["description"], font=("Segoe UI", 10), text_color=Theme.TEXT_SECONDARY, wraplength=200, justify="left", anchor="w")
        self.quickfix_desc.pack(fill="x", pady=(0, 8))
        ctk.CTkButton(fix_section, text="Apply Quick Fix", height=34, font=("Segoe UI Semibold", 12), fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER, command=self._apply_quickfix).pack(fill="x", pady=(0, 6))
        ctk.CTkButton(fix_section, text="Scan LTSC Gaps", height=32, font=("Segoe UI", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._scan_ltsc_gaps).pack(fill="x", pady=(0, 5))
        ctk.CTkButton(fix_section, text="Queue Xbox Core", height=32, font=("Segoe UI", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._queue_xbox_core).pack(fill="x", pady=(0, 6))

        updates_row = ctk.CTkFrame(fix_section, fg_color="transparent")
        updates_row.pack(fill="x")
        ctk.CTkCheckBox(
            updates_row,
            text="Keep updated",
            variable=self.keep_updated_var,
            font=("Segoe UI", 12),
            fg_color=Theme.PRIMARY,
            hover_color=Theme.PRIMARY_HOVER,
            command=self._change_keep_updated_mode,
        ).pack(side="left")
        ctk.CTkButton(
            updates_row,
            text="Check",
            width=64,
            height=30,
            font=("Segoe UI", 12),
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._run_keep_updated_scan,
        ).pack(side="right")

        ctk.CTkFrame(nav, fg_color=Theme.BORDER_SUBTLE, height=1).pack(fill="x", padx=4, pady=(0, 12))

        self.pinned_section = ctk.CTkFrame(nav, fg_color="transparent")
        ctk.CTkLabel(self.pinned_section, text="PINNED APPS", font=("Segoe UI Semibold", 10), text_color=Theme.TEXT_MUTED, anchor="w").pack(fill="x")
        self.pinned_list_frame = ctk.CTkFrame(self.pinned_section, fg_color="transparent", height=1)
        self.pinned_list_frame.pack(fill="x", pady=(6, 0))
        self._render_pinned_favorites()
        
        cat_section = ctk.CTkFrame(nav, fg_color="transparent")
        cat_section.pack(fill="x", padx=4, pady=(0, 12))
        ctk.CTkLabel(cat_section, text="BROWSE", font=("Segoe UI Semibold", 10), text_color=Theme.TEXT_MUTED, anchor="w").pack(fill="x", pady=(0, 5))
        for cat_name in APP_CATALOG.keys():
            ctk.CTkButton(cat_section, text=cat_name, height=32, font=("Segoe UI", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, hover_color=Theme.BG_CARD_HOVER, anchor="w", command=lambda c=cat_name: self._show_category(c)).pack(fill="x", pady=1)

        repair_frame = ctk.CTkFrame(nav, fg_color="transparent")
        repair_frame.pack(fill="x", padx=4, pady=(0, 8))
        ctk.CTkLabel(repair_frame, text="ADMIN TOOLS", font=("Segoe UI Semibold", 10), text_color=Theme.TEXT_MUTED, anchor="w").pack(fill="x", pady=(0, 6))
        repair_store_button = ctk.CTkButton(repair_frame, text="Inspect Store Repair", height=34, font=("Segoe UI Semibold", 11), fg_color=Theme.DANGER, hover_color=Theme.DANGER_HOVER, command=self._run_repair)
        repair_store_button.pack(fill="x", pady=(0, 5))
        provision_button = ctk.CTkButton(repair_frame, text="Inspect Provisioning", height=32, font=("Segoe UI", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._run_provisioning_repair)
        provision_button.pack(fill="x", pady=(0, 5))
        licensing_button = ctk.CTkButton(repair_frame, text="Inspect Licensing Reset", height=32, font=("Segoe UI", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._run_licensing_reset)
        licensing_button.pack(fill="x", pady=(0, 5))
        cache_button = ctk.CTkButton(repair_frame, text="Inspect Cache Rebuild", height=32, font=("Segoe UI", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._run_cache_rebuild)
        cache_button.pack(fill="x", pady=(0, 5))
        restore_button = ctk.CTkButton(repair_frame, text="Restore Backup", height=32, font=("Segoe UI", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._choose_repair_restore)
        restore_button.pack(fill="x", pady=(0, 5))
        self.repair_cancel_button = ctk.CTkButton(repair_frame, text="Cancel at Safe Checkpoint", height=32, font=("Segoe UI", 10), fg_color="transparent", text_color=Theme.TEXT_MUTED, border_width=1, border_color=Theme.BORDER_SUBTLE, hover_color=Theme.BG_CARD_HOVER, state="disabled", command=self._cancel_repair_operation)
        self.repair_cancel_button.pack(fill="x")
        self._repair_buttons = [
            repair_store_button,
            provision_button,
            licensing_button,
            cache_button,
            restore_button,
        ]
    
    def _build_queue_panel(self):
        self.right_panel.grid_rowconfigure(1, weight=1)
        self.right_panel.grid_columnconfigure(0, weight=1)

        header_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent", height=70)
        header_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        header_frame.grid_propagate(False)
        title_group = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_group.pack(side="left", fill="y")
        ctk.CTkLabel(title_group, text="Download queue", font=("Segoe UI Semibold", 17), anchor="w").pack(fill="x", pady=(2, 0))
        ctk.CTkLabel(title_group, text="Packages staged for this run", font=("Segoe UI", 10), text_color=Theme.TEXT_MUTED, anchor="w").pack(fill="x", pady=(2, 0))
        self.queue_count = ctk.CTkLabel(header_frame, text="0 items", font=("Segoe UI", 12), text_color=Theme.TEXT_MUTED)
        self.queue_count.pack(side="right", pady=(6, 0))

        self.queue_scroll = ctk.CTkScrollableFrame(
            self.right_panel,
            fg_color=Theme.BG_INPUT,
            corner_radius=7,
            border_width=1,
            border_color=Theme.BORDER_SUBTLE,
            height=100,
        )
        self.queue_scroll.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 10))
        self.queue_empty = ctk.CTkLabel(
            self.queue_scroll,
            text="○\n\nNo packages queued\n\nFind an app or browse a category\nto stage installation files.",
            font=("Segoe UI", 11),
            text_color=Theme.TEXT_MUTED,
            justify="center",
        )
        self.queue_empty.pack(expand=True, pady=32)

        controls = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        controls.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 14))

        output_header = ctk.CTkFrame(controls, fg_color="transparent")
        output_header.pack(fill="x")
        ctk.CTkLabel(output_header, text="OUTPUT", font=("Segoe UI Semibold", 9), text_color=Theme.TEXT_MUTED, anchor="w").pack(side="left")
        ctk.CTkButton(
            output_header,
            text="Change",
            width=56,
            height=24,
            font=("Segoe UI", 10),
            fg_color="transparent",
            text_color=Theme.PRIMARY_OUTLINE_TEXT,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._choose_output_folder,
        ).pack(side="right")
        self.output_path_label = ctk.CTkLabel(
            controls,
            text=self._format_output_path(),
            font=("Consolas", 9),
            text_color=Theme.TEXT_SECONDARY,
            anchor="w",
        )
        self.output_path_label.pack(fill="x", pady=(1, 7))

        cache_frame = ctk.CTkFrame(controls, fg_color="transparent")
        cache_frame.pack(fill="x", pady=(0, 6))

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
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._choose_shared_cache_folder,
        ).pack(side="right")

        self.shared_cache_label = ctk.CTkLabel(
            controls,
            text=self._format_shared_cache_path(),
            font=("Consolas", 10),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
        )
        self.shared_cache_label.pack(fill="x", pady=(0, 7))
        
        progress_frame = ctk.CTkFrame(controls, fg_color="transparent")
        progress_frame.pack(fill="x", pady=(0, 8))
        
        self.progress_label = ctk.CTkLabel(progress_frame, text="Ready", font=("Segoe UI", 11), text_color=Theme.TEXT_SECONDARY, anchor="w")
        self.progress_label.pack(fill="x")
        
        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=6, corner_radius=3, fg_color=Theme.BG_INPUT, progress_color=Theme.PRIMARY)
        self.progress_bar.pack(fill="x", pady=(4, 0))
        self.progress_bar.set(0)
        
        action_frame = ctk.CTkFrame(controls, fg_color="transparent")
        action_frame.pack(fill="x")
        ctk.CTkButton(action_frame, text="Download all", height=38, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._start_download).pack(fill="x", pady=(0, 6))

        export_grid = ctk.CTkFrame(action_frame, fg_color="transparent")
        export_grid.pack(fill="x", pady=(0, 6))
        export_grid.grid_columnconfigure((0, 1), weight=1)
        export_actions = [
            ("DISM script", self._export_dism_script),
            ("AppInstaller", self._export_appinstaller_manifest),
            ("IntuneWin", self._export_intunewin_package),
            ("Diagnostics", self._export_diagnostics_bundle),
        ]
        for index, (label, command) in enumerate(export_actions):
            row, column = divmod(index, 2)
            ctk.CTkButton(
                export_grid,
                text=label,
                height=29,
                font=("Segoe UI", 10),
                fg_color="transparent",
                text_color=Theme.TEXT_PRIMARY,
                border_width=1,
                border_color=Theme.BORDER,
                hover_color=Theme.BG_CARD_HOVER,
                command=command,
            ).grid(row=row, column=column, sticky="ew", padx=(0, 3) if column == 0 else (3, 0), pady=(0, 5) if row == 0 else 0)

        install_row = ctk.CTkFrame(action_frame, fg_color="transparent")
        install_row.pack(fill="x")
        ctk.CTkButton(install_row, text="Install", width=76, height=34, font=("Segoe UI Semibold", 11), fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER, command=self._start_install).pack(side="left", fill="x", expand=True, padx=(0, 3))
        ctk.CTkButton(install_row, text="Rollback", width=76, height=34, font=("Segoe UI Semibold", 11), fg_color="transparent", text_color=Theme.WARNING, border_width=1, border_color=Theme.WARNING, hover_color=Theme.BG_CARD_HOVER, command=self._start_rollback).pack(side="left", fill="x", expand=True, padx=3)
        ctk.CTkButton(install_row, text="Diff", width=64, height=34, font=("Segoe UI Semibold", 11), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._show_package_diff).pack(side="left", fill="x", expand=True, padx=(3, 0))
    
    def _build_log_panel(self):
        """Build the collapsible log/console panel"""
        # Toggle bar (always visible)
        self.log_toggle = ctk.CTkFrame(self.log_panel, fg_color=Theme.BG_INPUT, height=40, corner_radius=0)
        self.log_toggle.pack(fill="x", side="top")
        self.log_toggle.pack_propagate(False)
        
        toggle_inner = ctk.CTkFrame(self.log_toggle, fg_color="transparent")
        toggle_inner.pack(fill="x", padx=15)
        
        self.log_toggle_btn = ctk.CTkButton(
            toggle_inner,
            text="›  Activity",
            font=("Segoe UI Semibold", 12),
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
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
            text="Copy",
            width=60,
            height=26,
            font=("Segoe UI", 11),
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._copy_log
        ).pack(side="left", padx=3)
        
        ctk.CTkButton(
            log_controls,
            text="Clear",
            width=60,
            height=26,
            font=("Segoe UI", 11),
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
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
        self.log_toggle_btn.configure(text="›  Activity")
        
        # Add initial log message
        self._log("INFO", f"MSStoreHelper v{APP_VERSION} initialized")
        self._log("INFO", f"System Architecture: {SYSTEM_ARCH}")
        self._log("INFO", f"Administrator: {'Yes' if IS_ADMIN else 'No'}")
        self._log("INFO", f"Output Directory: {DEFAULT_OUTPUT}")
        self._log("INFO", f"Theme: {self.theme_mode_var.get()} ({Theme.MODE}) Accent: {Theme.PRIMARY}")
        settings = self._store_query_settings()
        self._log("INFO", f"Store query: ring={settings['Ring']}, language={settings['Language']}, market={settings['Market']}")
    
    def _toggle_log_panel(self):
        """Toggle log panel expanded/collapsed"""
        if self.log_expanded:
            self.log_content.pack_forget()
            self.log_toggle_btn.configure(text="›  Activity")
            self.log_expanded = False
        else:
            self.log_content.pack(fill="x", side="top")
            self.log_toggle_btn.configure(text="⌄  Activity")
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
        content = self._current_log_text()
        self.clipboard_clear()
        self.clipboard_append(content)
        self._log("INFO", "Log copied to clipboard")

    def _current_log_text(self):
        self.log_text.configure(state="normal")
        content = self.log_text.get("1.0", "end-1c")
        self.log_text.configure(state="disabled")
        return content
    
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

    def _set_workspace_navigation(self, active):
        if not hasattr(self, "workspace_button"):
            return
        self.workspace_button.configure(
            fg_color=Theme.BG_CARD_HOVER if active else "transparent",
            text_color=Theme.PRIMARY_OUTLINE_TEXT if active else Theme.TEXT_PRIMARY,
        )

    def _workspace_source_status(self):
        if self.source_health:
            available = sum(1 for status in self.source_health if status.get("Available"))
            total = len(self.source_health)
            if available == total:
                return f"{available}/{total} available", Theme.SUCCESS
            if available:
                return f"{available}/{total} available", Theme.WARNING
            return "Unavailable", Theme.DANGER
        if os.environ.get("MSSTOREHELPER_SKIP_SOURCE_HEALTH") == "1":
            return "Not checked", Theme.TEXT_MUTED
        return "Checking", Theme.INFO

    def _refresh_workspace_source_status(self):
        label = getattr(self, "workspace_source_value", None)
        if label is None or not label.winfo_exists():
            return
        text, color = self._workspace_source_status()
        label.configure(text=text, text_color=color)

    @staticmethod
    def _status_cell(parent, label, value, color):
        cell = ctk.CTkFrame(parent, fg_color="transparent")
        ctk.CTkLabel(
            cell,
            text=label.upper(),
            font=("Segoe UI Semibold", 9),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
        ).pack(fill="x")
        value_label = ctk.CTkLabel(
            cell,
            text=value,
            font=("Segoe UI Semibold", 11),
            text_color=color,
            anchor="w",
        )
        value_label.pack(fill="x", pady=(2, 0))
        return cell, value_label

    def _do_workspace_search(self):
        query = self.workspace_search_entry.get().strip()
        if not query:
            self.workspace_search_entry.focus()
            return
        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, query)
        self._do_search()

    def _show_welcome(self):
        self._clear_content()
        self.current_view = "welcome"
        self._set_workspace_navigation(True)

        workspace = ctk.CTkScrollableFrame(
            self.content,
            fg_color="transparent",
            corner_radius=0,
            scrollbar_button_color=Theme.BORDER,
            scrollbar_button_hover_color=Theme.TEXT_MUTED,
        )
        workspace.pack(fill="both", expand=True, padx=(22, 18), pady=(18, 14))

        header = ctk.CTkFrame(workspace, fg_color="transparent")
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="Workspace",
            font=("Segoe UI Semibold", 25),
            text_color=Theme.TEXT_PRIMARY,
            anchor="w",
        ).pack(fill="x")
        ctk.CTkLabel(
            header,
            text="Find, validate, and stage Microsoft Store packages for this PC.",
            font=("Segoe UI", 12),
            text_color=Theme.TEXT_SECONDARY,
            anchor="w",
        ).pack(fill="x", pady=(3, 0))

        trust = ctk.CTkFrame(
            workspace,
            fg_color=Theme.BG_CARD,
            corner_radius=7,
            border_width=1,
            border_color=Theme.BORDER_SUBTLE,
        )
        trust.pack(fill="x", pady=(16, 18))
        trust.grid_columnconfigure((0, 1, 2), weight=1)
        source_text, source_color = self._workspace_source_status()
        source_cell, self.workspace_source_value = self._status_cell(trust, "Sources", source_text, source_color)
        signature_cell, _ = self._status_cell(trust, "Signature policy", "Verify on install", Theme.INFO)
        admin_cell, _ = self._status_cell(
            trust,
            "Admin",
            "Elevated" if IS_ADMIN else "Standard session",
            Theme.SUCCESS if IS_ADMIN else Theme.WARNING,
        )
        source_cell.grid(row=0, column=0, sticky="ew", padx=(14, 10), pady=11)
        signature_cell.grid(row=0, column=1, sticky="ew", padx=10, pady=11)
        admin_cell.grid(row=0, column=2, sticky="ew", padx=(10, 14), pady=11)

        search_surface = ctk.CTkFrame(workspace, fg_color="transparent")
        search_surface.pack(fill="x")
        ctk.CTkLabel(
            search_surface,
            text="FIND APPS",
            font=("Segoe UI Semibold", 10),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
        ).pack(fill="x", pady=(0, 7))
        search_row = ctk.CTkFrame(search_surface, fg_color="transparent")
        search_row.pack(fill="x")
        self.workspace_search_entry = ctk.CTkEntry(
            search_row,
            placeholder_text="Search Microsoft Store",
            height=42,
            font=("Segoe UI", 13),
            fg_color=Theme.BG_INPUT,
            border_color=Theme.BORDER,
        )
        self.workspace_search_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.workspace_search_entry.bind("<Return>", lambda _event: self._do_workspace_search())
        ctk.CTkButton(
            search_row,
            text="Search",
            width=104,
            height=42,
            font=("Segoe UI Semibold", 12),
            fg_color=Theme.PRIMARY,
            hover_color=Theme.PRIMARY_HOVER,
            command=self._do_workspace_search,
        ).pack(side="right")

        history = self.user_profile.get("SearchHistory", [])[:4]
        if history:
            recent = ctk.CTkFrame(search_surface, fg_color="transparent")
            recent.pack(fill="x", pady=(9, 0))
            ctk.CTkLabel(recent, text="Recent", font=("Segoe UI", 10), text_color=Theme.TEXT_MUTED).pack(side="left", padx=(0, 6))
            for query in history:
                ctk.CTkButton(
                    recent,
                    text=query[:20],
                    width=72,
                    height=26,
                    font=("Segoe UI", 10),
                    fg_color=Theme.BG_ELEVATED,
                    text_color=Theme.TEXT_SECONDARY,
                    border_width=1,
                    border_color=Theme.BORDER_SUBTLE,
                    hover_color=Theme.BG_CARD_HOVER,
                    command=lambda q=query: self._search_from_history(q),
                ).pack(side="left", padx=(0, 5))

        favorites = self.user_profile.get("PinnedFavorites", [])[:4]
        favorites_header = ctk.CTkFrame(workspace, fg_color="transparent")
        favorites_header.pack(fill="x", pady=(22, 8))
        ctk.CTkLabel(
            favorites_header,
            text="Favorite apps",
            font=("Segoe UI Semibold", 15),
            text_color=Theme.TEXT_PRIMARY,
            anchor="w",
        ).pack(side="left")
        favorite_header_actions = ctk.CTkFrame(favorites_header, fg_color="transparent")
        favorite_header_actions.pack(side="right")
        ctk.CTkLabel(
            favorite_header_actions,
            text="Pinned from search or browse",
            font=("Segoe UI", 10),
            text_color=Theme.TEXT_MUTED,
        ).pack(side="left")
        if favorites:
            ctk.CTkButton(
                favorite_header_actions,
                text="Clear",
                width=44,
                height=24,
                font=("Segoe UI", 10),
                fg_color="transparent",
                text_color=Theme.PRIMARY_OUTLINE_TEXT,
                hover_color=Theme.BG_CARD_HOVER,
                command=self._clear_pinned_favorites,
            ).pack(side="left", padx=(7, 0))

        favorite_surface = ctk.CTkFrame(
            workspace,
            fg_color=Theme.BG_CARD,
            corner_radius=7,
            border_width=1,
            border_color=Theme.BORDER_SUBTLE,
        )
        favorite_surface.pack(fill="x")
        if favorites:
            for index, app in enumerate(favorites):
                row = ctk.CTkFrame(favorite_surface, fg_color="transparent")
                row.pack(fill="x", padx=12, pady=(9 if index == 0 else 5, 9 if index == len(favorites) - 1 else 5))
                ctk.CTkLabel(row, text=app.get("Icon", "▣"), width=28, font=("Segoe UI Emoji", 17)).pack(side="left")
                ctk.CTkLabel(row, text=app["Name"], font=("Segoe UI Semibold", 11), text_color=Theme.TEXT_PRIMARY, anchor="w").pack(side="left", fill="x", expand=True, padx=(6, 8))
                ctk.CTkButton(
                    row,
                    text="Get files",
                    width=76,
                    height=28,
                    font=("Segoe UI", 10),
                    fg_color="transparent",
                    text_color=Theme.PRIMARY_OUTLINE_TEXT,
                    border_width=1,
                    border_color=Theme.BORDER,
                    hover_color=Theme.BG_CARD_HOVER,
                    command=lambda a=app: self._fetch_single_app(a),
                ).pack(side="right")
        else:
            ctk.CTkLabel(
                favorite_surface,
                text="No favorites pinned yet. Select apps in search or browse, then choose Pin Selected.",
                font=("Segoe UI", 11),
                text_color=Theme.TEXT_MUTED,
                anchor="w",
                justify="left",
                wraplength=520,
            ).pack(fill="x", padx=14, pady=14)

        onboarding = ctk.CTkFrame(
            workspace,
            fg_color=Theme.BG_ELEVATED,
            corner_radius=7,
            border_width=1,
            border_color=Theme.BORDER_SUBTLE,
        )
        onboarding.pack(fill="x", pady=(18, 2))
        onboarding.grid_columnconfigure(0, weight=1)
        onboarding_copy = ctk.CTkFrame(onboarding, fg_color="transparent")
        onboarding_copy.grid(row=0, column=0, sticky="nsew", padx=(18, 10), pady=18)
        ctk.CTkLabel(onboarding_copy, text="Start with a known workflow", font=("Segoe UI Semibold", 16), text_color=Theme.TEXT_PRIMARY, anchor="w").pack(fill="x")
        ctk.CTkLabel(
            onboarding_copy,
            text="Browse the LTSC essentials catalog or scan this PC for common Store capability gaps.",
            font=("Segoe UI", 11),
            text_color=Theme.TEXT_SECONDARY,
            anchor="w",
            justify="left",
            wraplength=500,
        ).pack(fill="x", pady=(5, 12))
        onboarding_actions = ctk.CTkFrame(onboarding_copy, fg_color="transparent")
        onboarding_actions.pack(fill="x")
        ctk.CTkButton(
            onboarding_actions,
            text="Browse essentials",
            width=126,
            height=34,
            font=("Segoe UI Semibold", 11),
            fg_color=Theme.PRIMARY,
            hover_color=Theme.PRIMARY_HOVER,
            command=lambda: self._show_category("🛠️ Essential Repairs"),
        ).pack(side="left", padx=(0, 7))
        ctk.CTkButton(
            onboarding_actions,
            text="Scan LTSC gaps",
            width=122,
            height=34,
            font=("Segoe UI", 11),
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._scan_ltsc_gaps,
        ).pack(side="left")
        ctk.CTkLabel(
            onboarding,
            text="▣",
            width=82,
            font=("Segoe UI Symbol", 42),
            text_color=Theme.PRIMARY,
        ).grid(row=0, column=1, padx=(0, 16), pady=16)
    
    def _show_category(self, category_name):
        self._clear_content()
        self.current_view = "category"
        self._set_workspace_navigation(False)
        self.selected_apps.clear()
        
        cat_data = APP_CATALOG.get(category_name, {})
        apps = cat_data.get("apps", [])
        
        header = ctk.CTkFrame(self.content, fg_color="transparent")
        header.pack(fill="x", padx=22, pady=(20, 10))

        title_group = ctk.CTkFrame(header, fg_color="transparent")
        title_group.pack(fill="x")
        ctk.CTkLabel(title_group, text=category_name, font=("Segoe UI Semibold", 23), anchor="w", justify="left", wraplength=360).pack(fill="x")
        ctk.CTkLabel(title_group, text=cat_data.get("description", ""), font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY, anchor="w", justify="left", wraplength=360).pack(fill="x", pady=(2, 0))
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(actions, text="Pin Selected", width=105, height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._pin_selected_apps).pack(side="left", padx=(0, 8))
        ctk.CTkButton(actions, text="Export WinGet", width=120, height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._export_winget_manifest).pack(side="left", padx=(0, 8))
        ctk.CTkButton(actions, text="Get Selected Apps", width=135, height=36, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._fetch_selected).pack(side="left")
        
        scroll = ctk.CTkScrollableFrame(self.content, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=18, pady=(0, 14))
        
        for app in apps:
            AppTile(scroll, app, self._on_app_toggle, self._show_release_notes).pack(fill="x")
    
    def _show_search_results(self, results, query):
        self._clear_content()
        self.current_view = "search"
        self._set_workspace_navigation(False)
        self.selected_apps.clear()
        
        header = ctk.CTkFrame(self.content, fg_color="transparent")
        header.pack(fill="x", padx=22, pady=(20, 10))

        title_group = ctk.CTkFrame(header, fg_color="transparent")
        title_group.pack(fill="x")
        ctk.CTkLabel(title_group, text=f'Results for "{query}"', font=("Segoe UI Semibold", 22), anchor="w", justify="left", wraplength=360).pack(fill="x")
        ctk.CTkLabel(title_group, text=f"{len(results)} apps found", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(2, 0))
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(actions, text="Pin Selected", width=105, height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._pin_selected_apps).pack(side="left", padx=(0, 8))
        ctk.CTkButton(actions, text="Export WinGet", width=120, height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._export_winget_manifest).pack(side="left", padx=(0, 8))
        ctk.CTkButton(actions, text="Get Selected Apps", width=135, height=36, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._fetch_selected).pack(side="left")
        
        if not results:
            empty = ctk.CTkFrame(self.content, fg_color="transparent")
            empty.pack(expand=True)
            ctk.CTkLabel(empty, text="○", font=("Segoe UI Symbol", 42), text_color=Theme.TEXT_MUTED).pack(pady=(0, 10))
            ctk.CTkLabel(empty, text="No apps found", font=("Segoe UI Semibold", 18)).pack()
            ctk.CTkLabel(empty, text="Try a broader product name or publisher.", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY).pack(pady=(5, 0))
            return
        
        scroll = ctk.CTkScrollableFrame(self.content, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=18, pady=(0, 14))
        
        for app in results:
            SearchResultTile(scroll, app, self._fetch_single_app, self._on_app_toggle, self._show_release_notes).pack(fill="x")
    
    def _show_packages(self, packages, title):
        self._clear_content()
        self.current_view = "packages"
        self._set_workspace_navigation(False)
        self.current_packages = packages
        self.selected_packages.clear()
        self.package_rows.clear()
        
        header = ctk.CTkFrame(self.content, fg_color="transparent")
        header.pack(fill="x", padx=22, pady=(20, 8))
        
        ctk.CTkLabel(header, text=title, font=("Segoe UI Semibold", 21)).pack(side="left")
        self.selection_info = ctk.CTkLabel(header, text="0 selected", font=("Segoe UI", 12), text_color=Theme.INFO)
        self.selection_info.pack(side="right", padx=15)
        
        toolbar = ctk.CTkFrame(self.content, fg_color=Theme.BG_CARD, corner_radius=7, border_width=1, border_color=Theme.BORDER_SUBTLE)
        toolbar.pack(fill="x", padx=22, pady=(5, 10))
        
        tb_inner = ctk.CTkFrame(toolbar, fg_color="transparent")
        tb_inner.pack(fill="x", padx=10, pady=8)
        tb_inner.grid_columnconfigure(3, weight=1)
        ctk.CTkButton(tb_inner, text="Smart Select", width=112, height=32, font=("Segoe UI", 12), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._smart_select).grid(row=0, column=0, sticky="w", padx=(0, 7))
        ctk.CTkButton(tb_inner, text="Select All", width=84, height=32, font=("Segoe UI", 12), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._select_all).grid(row=0, column=1, sticky="w", padx=(0, 7))
        ctk.CTkButton(tb_inner, text="Clear", width=64, height=32, font=("Segoe UI", 12), fg_color="transparent", text_color=Theme.TEXT_PRIMARY, border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._select_none).grid(row=0, column=2, sticky="w")

        arch_frame = ctk.CTkFrame(tb_inner, fg_color="transparent")
        arch_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ctk.CTkLabel(arch_frame, text="Target arch", font=("Segoe UI", 11), text_color=Theme.TEXT_SECONDARY).pack(side="left", padx=(0, 6))
        ctk.CTkOptionMenu(arch_frame, values=self.arch_options, variable=self.arch_override_var, width=120, height=30, font=("Segoe UI", 11), fg_color=Theme.BG_INPUT, text_color=Theme.TEXT_PRIMARY, button_color=Theme.PRIMARY, button_hover_color=Theme.PRIMARY_HOVER, command=self._on_arch_override_change).pack(side="left")
        ctk.CTkButton(tb_inner, text="Add to queue", width=116, height=32, font=("Segoe UI Semibold", 11), fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER, command=self._add_to_queue).grid(row=1, column=3, sticky="e", pady=(8, 0))
        
        col_header = ctk.CTkFrame(self.content, fg_color=Theme.BG_INPUT, corner_radius=6)
        col_header.pack(fill="x", padx=22, pady=(0, 5))
        
        ch_inner = ctk.CTkFrame(col_header, fg_color="transparent")
        ch_inner.pack(fill="x", padx=12, pady=8)
        ch_inner.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(ch_inner, text="Select", font=("Segoe UI Semibold", 10), width=58).grid(row=0, column=0)
        ctk.CTkLabel(ch_inner, text="File Name", font=("Segoe UI Semibold", 11), anchor="w").grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(ch_inner, text="Size", font=("Segoe UI Semibold", 11), width=80).grid(row=0, column=2, padx=(0, 10))
        
        self.package_scroll = ctk.CTkScrollableFrame(self.content, fg_color="transparent")
        self.package_scroll.pack(fill="both", expand=True, padx=22, pady=(0, 14))
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

📄 App Installer Export
Download queued packages, then export a .appinstaller manifest and package folder.

📝 Release Notes
Click "Notes" on any app row to fetch Store product page notes.

📦 IntuneWin Export
Download queued packages, then export an IntuneWin package with a detection script.

✨ Smart Select
Automatically picks the best files - prefers bundles, skips encrypted files, chooses newest versions.

⬇️ Downloading
Add files to queue and click "Download All". Files save to Downloads folder.
Interrupted downloads resume from .part files when the server supports Range requests.

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

    def _show_release_notes(self, app_data):
        product_id = app_data.get("ProductId")
        if not product_id:
            self._update_status("⚠️ Missing product ID", Theme.WARNING)
            return

        self._update_status(f"📝 Fetching notes for {app_data.get('Name', 'app')}...", Theme.INFO)
        threading.Thread(target=self._release_notes_worker, args=(app_data,), daemon=True).start()

    def _release_notes_worker(self, app_data):
        try:
            settings = self._store_query_settings()
            notes = StoreAPI.fetch_release_notes(app_data["ProductId"], settings["Language"], settings["Market"])
            self.after(0, lambda s=settings, n=app_data.get("Name", "app"): self._log("INFO", f"Release notes query for {n}: language={s['Language']}, market={s['Market']}"))
        except Exception as exc:
            self.after(0, lambda: self._update_status("❌ Release notes failed", Theme.DANGER))
            self.after(0, lambda e=str(exc), n=app_data.get("Name", "app"): self._log("ERROR", f"Failed to fetch release notes for {n}: {e}"))
            return

        self.after(0, lambda: self._update_status("Ready", Theme.TEXT_SECONDARY))
        self.after(0, lambda: self._show_release_notes_dialog(app_data, notes))

    def _show_release_notes_dialog(self, app_data, notes):
        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Release Notes - {app_data.get('Name', 'App')}")
        dialog.geometry("640x520")
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 640) // 2
        y = self.winfo_y() + (self.winfo_height() - 520) // 2
        dialog.geometry(f"+{x}+{y}")

        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=22, pady=22)

        ctk.CTkLabel(content, text=app_data.get("Name", notes.get("Title", "Release Notes")), font=("Segoe UI Semibold", 20), anchor="w").pack(fill="x")
        ctk.CTkLabel(content, text=f"Source: Microsoft Store ({notes.get('Source', 'store-page')})", font=("Segoe UI", 11), text_color=Theme.TEXT_MUTED, anchor="w").pack(fill="x", pady=(2, 12))

        textbox = ctk.CTkTextbox(content, font=("Segoe UI", 12), fg_color=Theme.BG_DARK, text_color=Theme.TEXT_SECONDARY, wrap="word")
        textbox.pack(fill="both", expand=True)
        textbox.insert("1.0", notes.get("Notes", "No release notes found."))
        textbox.configure(state="disabled")

        button_row = ctk.CTkFrame(content, fg_color="transparent")
        button_row.pack(fill="x", pady=(12, 0))
        ctk.CTkButton(button_row, text="Open Store Page", width=130, height=32, font=("Segoe UI", 12), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=lambda: webbrowser.open(notes.get("Url", ""))).pack(side="right")
    
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

    def _save_user_profile(self):
        StoreAPI.save_user_profile(self.user_profile)

    def _store_query_settings(self):
        return StoreAPI.store_query_settings(
            self.store_ring_var.get(),
            self.store_language_var.get(),
            self.store_market_var.get(),
        )

    def _store_query_button_text(self):
        settings = self._store_query_settings()
        return f"Query: {settings['Ring']} / {settings['Language']} / {settings['Market']}"

    def _refresh_store_query_button(self):
        if hasattr(self, "store_query_button"):
            self.store_query_button.configure(text=self._store_query_button_text())

    def _save_store_query_settings(self):
        settings = self._store_query_settings()
        self.store_ring_var.set(settings["Ring"])
        self.store_language_var.set(settings["Language"])
        self.store_market_var.set(settings["Market"])
        self.user_profile["StoreRing"] = settings["Ring"]
        self.user_profile["StoreLanguage"] = settings["Language"]
        self.user_profile["StoreMarket"] = settings["Market"]
        self._save_user_profile()
        return settings

    def _change_store_query_setting(self, _choice=None):
        settings = self._save_store_query_settings()
        self._refresh_store_query_button()
        summary = f"ring={settings['Ring']}, language={settings['Language']}, market={settings['Market']}"
        self._update_status(f"Store query: {summary}", Theme.INFO)
        self._log("INFO", f"Store query settings changed: {summary}")

    def _show_store_query_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Store Query Settings")
        dialog.geometry("420x300")
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 420) // 2
        y = self.winfo_y() + (self.winfo_height() - 300) // 2
        dialog.geometry(f"+{x}+{y}")

        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=22, pady=22)

        ctk.CTkLabel(content, text="Store Query", font=("Segoe UI Semibold", 20), anchor="w").pack(fill="x")
        ctk.CTkLabel(
            content,
            text="Used for package lookup, Store product pages, logs, and exported deployment artifacts.",
            font=("Segoe UI", 12),
            text_color=Theme.TEXT_SECONDARY,
            justify="left",
            wraplength=360,
            anchor="w",
        ).pack(fill="x", pady=(4, 14))

        controls = ctk.CTkFrame(content, fg_color="transparent")
        controls.pack(fill="x")
        control_rows = [
            ("Ring", STORE_RING_VALUES, self.store_ring_var),
            ("Language", STORE_LANGUAGE_VALUES, self.store_language_var),
            ("Market", STORE_MARKET_VALUES, self.store_market_var),
        ]
        for row, (label, values, variable) in enumerate(control_rows):
            controls.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(controls, text=label, font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY, anchor="w").grid(row=row, column=0, sticky="w", padx=(0, 12), pady=6)
            ctk.CTkOptionMenu(
                controls,
                values=values,
                variable=variable,
                width=180,
                height=32,
                font=("Segoe UI", 12),
                fg_color=Theme.BG_INPUT,
                text_color=Theme.TEXT_PRIMARY,
                button_color=Theme.PRIMARY,
                button_hover_color=Theme.PRIMARY_HOVER,
                command=self._change_store_query_setting,
            ).grid(row=row, column=1, sticky="ew", pady=6)

        ctk.CTkButton(
            content,
            text="Close",
            width=90,
            height=32,
            font=("Segoe UI", 12),
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=dialog.destroy,
        ).pack(side="right", pady=(18, 0))

    def _change_theme_mode(self, choice):
        mode = Theme.normalize_mode(choice)
        self.theme_mode_var.set(mode)
        self.user_profile["ThemeMode"] = mode
        Theme.set_mode(mode)
        ctk.set_appearance_mode(Theme.MODE)
        self.configure(fg_color=Theme.BG_DARK)
        self._save_user_profile()
        self._update_status(f"Theme: {mode} ({Theme.MODE})", Theme.INFO)
        self._log("INFO", f"Theme changed to {mode} ({Theme.MODE}) with accent {Theme.PRIMARY}")

    def _render_search_history(self):
        if not hasattr(self, "search_history_frame"):
            return

        for widget in self.search_history_frame.winfo_children():
            widget.destroy()
        self.search_history_frame.pack_forget()

    def _render_pinned_favorites(self):
        if not hasattr(self, "pinned_list_frame"):
            return

        for widget in self.pinned_list_frame.winfo_children():
            widget.destroy()
        self.pinned_section.pack_forget()

    def _search_from_history(self, query):
        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, query)
        self._do_search()

    def _pin_selected_apps(self):
        if not self.selected_apps:
            self._update_status("⚠️ No apps selected", Theme.WARNING)
            self._log("WARNING", "No selected apps to pin")
            return

        added = StoreAPI.add_pinned_favorites(self.user_profile, self.selected_apps)
        self._save_user_profile()
        self._render_pinned_favorites()
        self._update_status(f"⭐ Pinned {len(self.selected_apps)} app(s)", Theme.SUCCESS)
        self._log("SUCCESS", f"Pinned favorites updated ({added} new, {len(self.user_profile.get('PinnedFavorites', []))} total)")

    def _clear_pinned_favorites(self):
        self.user_profile["PinnedFavorites"] = []
        self._save_user_profile()
        self._render_pinned_favorites()
        self._update_status("Pinned apps cleared", Theme.TEXT_SECONDARY)
        self._log("INFO", "Pinned favorites cleared")
        if self.current_view == "welcome":
            self._show_welcome()

    def _format_shared_cache_path(self):
        if len(self.shared_cache_path) <= 38:
            return self.shared_cache_path
        return f"{self.shared_cache_path[:16]}...{self.shared_cache_path[-19:]}"

    def _format_output_path(self):
        if len(self.output_path) <= 42:
            return self.output_path
        return f"{self.output_path[:18]}...{self.output_path[-21:]}"

    def _choose_output_folder(self):
        selected = filedialog.askdirectory(
            title="Select package output folder",
            initialdir=self.output_path if os.path.exists(self.output_path) else DEFAULT_OUTPUT,
        )
        if not selected:
            return
        self.output_path = selected
        self.output_path_label.configure(text=self._format_output_path())
        self._save_download_state()
        self._log("INFO", f"Output directory changed: {self.output_path}")

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

    def _post_ui(self, callback):
        try:
            self.after(0, callback)
        except RuntimeError:
            pass

    def _source_health_worker(self):
        self._post_ui(lambda: self._log("INFO", "Checking Store source availability..."))
        try:
            statuses = StoreAPI.detect_source_health()
        except Exception as exc:
            self._post_ui(lambda e=str(exc): self._log("WARNING", f"Source health check failed: {e}"))
            return

        self.source_health = statuses
        self._post_ui(self._refresh_workspace_source_status)
        for status in statuses:
            level = "SUCCESS" if status.get("Available") else "WARNING"
            self._post_ui(lambda s=status, lvl=level: self._log(lvl, source_status_summary(s)))

    def _log_source_diagnostic(self, diagnostic, item_name=None):
        label = item_name or diagnostic.get("Source", "Store source")
        query = diagnostic.get("Query")
        if query:
            summary = f"ring={query.get('Ring')}, language={query.get('Language')}, market={query.get('Market')}"
            self.after(0, lambda n=label, s=summary: self._log("INFO", f"{n} package query: {s}"))
        for error in diagnostic.get("Errors", []):
            self.after(0, lambda e=error, s=diagnostic.get("Source", "Store source"): self._log("ERROR", f"{s}: {e}"))
        for fallback in diagnostic.get("Fallbacks", []):
            command = fallback.get("Command", "")
            detail = fallback.get("Detail", "")
            self.after(0, lambda f=fallback, c=command, d=detail, n=label: self._log("WARNING", f"{n} fallback via {f.get('Source')}: {c} ({d})"))

    def _get_packages_with_logging(self, app_data):
        settings = self._store_query_settings()
        diagnostic = StoreAPI.get_packages_with_diagnostics(
            app_data["ProductId"],
            settings["Ring"],
            settings["Language"],
            settings["Market"],
        )
        self._log_source_diagnostic(diagnostic, app_data.get("Name"))
        return diagnostic["Packages"]

    def _change_keep_updated_mode(self):
        enabled = bool(self.keep_updated_var.get())
        self.user_profile["KeepUpdatedEnabled"] = enabled
        self._save_user_profile()

        if enabled:
            self._update_status("Keep updated enabled", Theme.INFO)
            self._log("INFO", "Keep updated enabled; scans run while MSStoreHelper is open")
            self._run_keep_updated_scan()
        else:
            if self.keep_updated_after_id:
                try:
                    self.after_cancel(self.keep_updated_after_id)
                except Exception:
                    pass
                self.keep_updated_after_id = None
            self._update_status("Keep updated disabled", Theme.TEXT_SECONDARY)
            self._log("INFO", "Keep updated disabled")

    def _schedule_keep_updated_scan(self, delay_ms=KEEP_UPDATED_INTERVAL_MS):
        if not self.keep_updated_var.get():
            return
        if self.keep_updated_after_id:
            try:
                self.after_cancel(self.keep_updated_after_id)
            except Exception:
                pass
        self.keep_updated_after_id = self.after(
            delay_ms,
            lambda: self._run_keep_updated_scan(scheduled=True),
        )

    def _run_keep_updated_scan(self, scheduled=False):
        if self.keep_updated_running:
            self._log("INFO", "Keep updated scan already running")
            return

        target_arch = self._target_arch()
        prefer_exact = self._has_arch_override()
        self.keep_updated_running = True
        self._update_status("Checking installed Store apps...", Theme.INFO)
        threading.Thread(
            target=self._keep_updated_worker,
            args=(scheduled, target_arch, prefer_exact),
            daemon=True,
        ).start()

    def _finish_keep_updated_scan(self):
        self.keep_updated_running = False
        if self.keep_updated_var.get():
            self._schedule_keep_updated_scan()

    def _keep_updated_worker(self, scheduled=False, target_arch=SYSTEM_ARCH, prefer_exact=False):
        try:
            installed_versions = StoreAPI.get_installed_appx_versions()
            if not installed_versions:
                self.after(0, lambda: self._update_status("No installed Store apps found", Theme.WARNING))
                self.after(0, lambda: self._log("WARNING", "Keep updated scan found no installed AppX/MSIX packages"))
                return

            self.after(0, lambda c=len(installed_versions): self._log("INFO", f"Checking {c} installed AppX/MSIX identities against the catalog"))
            update_packages = StoreAPI.select_catalog_update_packages(
                APP_CATALOG,
                installed_versions,
                self._get_packages_with_logging,
                target_arch,
                prefer_exact,
            )
            self.user_profile["KeepUpdatedLastScan"] = datetime.now(timezone.utc).isoformat()
            self._save_user_profile()

            if not update_packages:
                self.after(0, lambda: self._update_status("Installed catalog apps are current", Theme.SUCCESS))
                self.after(0, lambda: self._log("SUCCESS", "Keep updated scan found no newer catalog packages"))
                return

            queued_count = self._queue_unique_packages(update_packages, target_arch)
            if queued_count:
                app_names = sorted({package.get("UpdateSourceApp", "app") for package in update_packages if package.get("UpdateSourceApp")})
                label = ", ".join(app_names[:4]) + ("..." if len(app_names) > 4 else "")
                self.after(0, self._update_queue_ui)
                self.after(0, lambda c=queued_count: self._update_status(f"Queued {c} update package(s)", Theme.SUCCESS))
                self.after(0, lambda c=queued_count, names=label: self._log("SUCCESS", f"Keep updated queued {c} package(s): {names}"))
            else:
                self.after(0, lambda: self._update_status("Updates already queued", Theme.INFO))
                self.after(0, lambda: self._log("INFO", "Keep updated found packages already present in the queue"))
        except Exception as exc:
            self.after(0, lambda: self._update_status("Keep updated scan failed", Theme.DANGER))
            self.after(0, lambda e=str(exc): self._log("ERROR", f"Keep updated scan failed: {e}"))
        finally:
            self.after(0, self._finish_keep_updated_scan)

    def _queue_unique_packages(self, packages, target_arch):
        queued_names = {
            package.get("FileName", "").lower()
            for package in self.download_queue
            if package.get("FileName")
        }
        queued_count = 0
        for package in packages:
            try:
                package = validate_package_record(package, require_url=True)
            except PackageIngressError as exc:
                self._log("WARNING", f"Rejected unsafe package metadata: {exc}")
                continue
            filename = package["FileName"]
            key = filename.lower()
            if key in queued_names:
                continue
            self.download_queue.append(annotate_package(package))
            queued_names.add(key)
            queued_count += 1

        if queued_count:
            self.download_queue = StoreAPI.order_packages_for_install(self.download_queue, target_arch)
            self._save_download_state()
        return queued_count
    
    def _do_search(self):
        query = self.search_entry.get().strip()
        if not query:
            return
        StoreAPI.add_search_history(self.user_profile, query)
        self._save_user_profile()
        self._render_search_history()
        self._update_status("🔍 Searching...", Theme.INFO)
        threading.Thread(target=self._search_worker, args=(query,), daemon=True).start()
    
    def _search_worker(self, query):
        self.after(0, lambda: self._log("INFO", f"Searching Microsoft Store for: {query}"))
        diagnostic = StoreAPI.search_store_with_diagnostics(query)
        results = diagnostic["Results"]
        self._log_source_diagnostic(diagnostic, "Search")
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

        selected_packages = []
        for app in missing_apps:
            self.after(0, lambda n=app["Name"]: self._update_status(f"📥 Queueing {n}...", Theme.INFO))
            packages = self._get_packages_with_logging(app)
            if not packages:
                self.after(0, lambda n=app["Name"]: self._log("WARNING", f"No packages found for missing LTSC component: {n}"))
                continue

            recommended = StoreAPI.smart_select(packages, target_arch, prefer_exact)
            selected_packages.extend(recommended)

        queued_count = self._queue_unique_packages(selected_packages, target_arch)
        dependency_count = sum(1 for pkg in self.download_queue if is_dependency_package(pkg))

        if queued_count:
            self.after(0, self._update_queue_ui)
            self.after(0, lambda c=queued_count: self._update_status(f"✅ Queued {c} LTSC package(s)", Theme.SUCCESS))
            self.after(0, lambda c=queued_count, d=dependency_count: self._log("SUCCESS", f"Queued {c} package(s) for missing LTSC components; dependencies in queue: {d}"))
        else:
            self.after(0, lambda: self._update_status("⚠️ No LTSC packages queued", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", "LTSC scan found missing components but no downloadable packages were queued"))

    def _queue_xbox_core(self):
        target_arch = self._target_arch()
        prefer_exact = self._has_arch_override()
        self._update_status("🎮 Fetching Xbox core packages...", Theme.INFO)
        threading.Thread(
            target=self._queue_xbox_core_worker,
            args=(target_arch, prefer_exact),
            daemon=True,
        ).start()

    def _queue_xbox_core_worker(self, target_arch, prefer_exact):
        all_packages = []
        for pin in XBOX_CORE_PACKAGE_PINS:
            self.after(0, lambda n=pin["Name"]: self._log("INFO", f"Fetching pinned Xbox core package: {n}"))
            packages = self._get_packages_with_logging(pin)
            if not packages:
                self.after(0, lambda n=pin["Name"]: self._log("WARNING", f"No packages found for Xbox core item: {n}"))
            all_packages.extend(packages)

        selected = StoreAPI.select_pinned_xbox_packages(all_packages, target_arch, prefer_exact)
        existing_names = {
            package.get("FileName", "").lower()
            for package in self.download_queue
            if package.get("FileName")
        }
        queued_count = self._queue_unique_packages(selected, target_arch)
        for package in selected:
            if package.get("FileName", "").lower() in existing_names:
                continue
            if package.get("XboxCoreName"):
                if package.get("PinnedVersionMatched"):
                    self.after(0, lambda p=package: self._log("SUCCESS", f"Using pinned {p['XboxCoreName']} version: {p.get('AvailableVersion', 'unknown')}"))
                else:
                    pins = ", ".join(package.get("PinnedVersions", [])) or "configured pin"
                    self.after(0, lambda p=package, pins=pins: self._log("WARNING", f"Pinned {p['XboxCoreName']} version ({pins}) not available; queued {p.get('AvailableVersion', 'unknown')}"))

        if queued_count:
            self.after(0, self._update_queue_ui)
            self.after(0, lambda c=queued_count: self._update_status(f"✅ Queued {c} Xbox core package(s)", Theme.SUCCESS))
            self.after(0, lambda c=queued_count: self._log("SUCCESS", f"Queued {c} Xbox core package(s) with dependency-first ordering"))
        else:
            self.after(0, lambda: self._update_status("⚠️ No Xbox packages queued", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", "Xbox core queue did not find downloadable packages"))
    
    def _fetch_selected_worker(self):
        all_packages = []
        names = []
        for app in self.selected_apps:
            names.append(app['Name'])
            self.after(0, lambda n=app['Name']: self._update_status(f"📥 Fetching {n}...", Theme.INFO))
            self.after(0, lambda n=app['Name'], pid=app['ProductId']: self._log("INFO", f"Fetching packages for: {n} ({pid})"))
            packages = self._get_packages_with_logging(app)
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
        packages = self._get_packages_with_logging(app_data)
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
                    try:
                        package = validate_package_record(pkg, require_url=True)
                    except PackageIngressError as exc:
                        self._log("WARNING", f"Rejected unsafe package metadata: {exc}")
                        continue
                    self.download_queue.append(annotate_package(package))
                    count += 1

        self.download_queue = StoreAPI.order_packages_for_install(self.download_queue, self._target_arch())
        dependency_count = sum(1 for pkg in self.download_queue if is_dependency_package(pkg))

        self._save_download_state()
        self._update_queue_ui()
        self._update_status(f"✅ Added {count} files to queue", Theme.SUCCESS)
        self._log("INFO", f"Added {count} files to download queue (total: {len(self.download_queue)})")
        if dependency_count:
            self._log("INFO", f"Install order resolved: {dependency_count} dependency package(s) before apps")
    
    def _update_queue_ui(self):
        for widget in self.queue_scroll.winfo_children():
            widget.destroy()
        
        if not self.download_queue:
            ctk.CTkLabel(
                self.queue_scroll,
                text="○\n\nNo packages queued\n\nFind an app or browse a category\nto stage installation files.",
                font=("Segoe UI", 11),
                text_color=Theme.TEXT_MUTED,
                justify="center",
            ).pack(expand=True, pady=32)
        else:
            for pkg in self.download_queue:
                QueueItem(
                    self.queue_scroll,
                    pkg,
                    self._review_package_trust,
                ).pack(fill="x", pady=3, padx=5)
        
        count = len(self.download_queue)
        self.queue_count.configure(text=f"{count} {'item' if count == 1 else 'items'}")

    def _review_package_trust(self, package):
        report = package.get("TrustReport") or {}
        path = package.get("LocalPath")
        if (
            report.get("State") != TRUST_STATE_REVIEW_REQUIRED
            or report.get("ReviewEligible") is not True
            or not path
        ):
            self._update_status("Package is not eligible for review", Theme.WARNING)
            self._log("WARNING", "Only identity-valid packages missing authoritative product binding can be reviewed")
            return

        source = report.get("Source") or {}
        manifest = report.get("Manifest") or {}
        signature = report.get("Signature") or {}
        dialog = ctk.CTkToplevel(self)
        dialog.title("Review quarantined package")
        dialog.geometry("720x570")
        dialog.minsize(620, 500)
        dialog.transient(self)
        dialog.grab_set()

        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=24, pady=22)
        ctk.CTkLabel(
            content,
            text="Review quarantined package",
            font=("Segoe UI Semibold", 20),
            anchor="w",
        ).pack(fill="x")
        ctk.CTkLabel(
            content,
            text=(
                "Windows signature and manifest checks passed, but the Store "
                "product could not be bound to an authoritative package identity."
            ),
            font=("Segoe UI", 11),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=430,
        ).pack(fill="x", pady=(4, 14))

        evidence = "\n".join([
            f"File: {package.get('FileName', '')}",
            f"Source: {source.get('Url') or 'not supplied'}",
            f"Store product: {source.get('ProductId') or 'not supplied'}",
            f"Identity: {manifest.get('Identity') or 'missing'}",
            f"Package family: {manifest.get('PackageFamilyName') or 'missing'}",
            f"Publisher: {manifest.get('Publisher') or 'missing'}",
            f"Manifest: {manifest.get('ManifestPath') or 'missing'}",
            f"Version / architecture: {manifest.get('Version') or 'missing'} / {manifest.get('Architecture') or 'missing'}",
            f"Signature chain: {'valid' if signature.get('ChainValid') else 'invalid'}",
            f"Revocation: {signature.get('RevocationState') or 'unknown'}",
            f"SHA-256: {report.get('ArtifactSha256') or 'missing'}",
        ])
        evidence_box = ctk.CTkTextbox(
            content,
            height=150,
            font=("Consolas", 10),
            fg_color=Theme.BG_INPUT,
            text_color=Theme.TEXT_SECONDARY,
            wrap="char",
        )
        evidence_box.pack(fill="x")
        evidence_box.insert("1.0", evidence)
        evidence_box.configure(state="disabled")

        ctk.CTkLabel(
            content,
            text=(
                "Approving records this evidence in the local trust journal "
                "and enables cache, mirror, export, rollback, and install automation."
            ),
            font=("Segoe UI Semibold", 10),
            text_color=Theme.WARNING,
            anchor="w",
            justify="left",
            wraplength=430,
        ).pack(fill="x", pady=(12, 8))

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.pack(fill="x")

        def approve():
            try:
                StoreAPI.review_package_trust(package, path)
            except Exception as exc:
                self._update_status("Package trust review failed", Theme.DANGER)
                self._log("ERROR", f"Package trust review failed: {exc}")
                return
            package["DownloadStatus"] = "Downloaded"
            package.pop("LastError", None)
            self._save_download_state()
            dialog.destroy()
            self._update_queue_ui()
            self._update_status("Package trust review recorded", Theme.SUCCESS)
            self._log(
                "SUCCESS",
                (
                    f"Promoted reviewed package {package.get('FileName')} "
                    f"and recorded {TRUST_REVIEW_JOURNAL_PATH}"
                ),
            )

        ctk.CTkButton(
            actions,
            text="Cancel",
            width=96,
            height=34,
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=dialog.destroy,
        ).pack(side="left")
        ctk.CTkButton(
            actions,
            text="Approve package",
            width=136,
            height=34,
            font=("Segoe UI Semibold", 11),
            fg_color=Theme.WARNING,
            hover_color=Theme.PRIMARY_HOVER,
            command=approve,
        ).pack(side="left", padx=(8, 0))

    def _save_download_state(self):
        if self.download_queue or os.path.abspath(self.output_path) != os.path.abspath(DEFAULT_OUTPUT):
            StoreAPI.write_download_state(self.download_queue, self.output_path)
        else:
            StoreAPI.clear_download_state()
    
    def _clear_queue(self):
        self.download_queue.clear()
        self._update_queue_ui()
        self._save_download_state()
        self._update_status("Queue cleared", Theme.TEXT_SECONDARY)
    
    def _start_download(self):
        if not self.download_queue:
            self._update_status("⚠️ Queue is empty", Theme.WARNING)
            return
        threading.Thread(target=self._download_worker, daemon=True).start()

    def _export_diagnostics_bundle(self):
        try:
            entries = StoreAPI.prepare_diagnostics_bundle(
                APP_VERSION,
                SYSTEM_ARCH,
                IS_ADMIN,
                self.output_path,
                self.source_health,
                self.download_queue,
                self._current_log_text(),
            )
        except (DiagnosticRedactionError, OSError, ValueError) as exc:
            self._update_status(
                "Diagnostics redaction failed closed",
                Theme.DANGER,
            )
            self._log(
                "ERROR",
                f"Diagnostics preview was not created: {exc}",
            )
            return
        self._show_diagnostics_preview(entries)

    def _show_diagnostics_preview(self, entries):
        preview = diagnostic_preview_text(entries)
        dialog = ctk.CTkToplevel(self)
        dialog.title("Preview redacted diagnostics")
        dialog.geometry("820x700")
        dialog.minsize(620, 500)
        dialog.transient(self)
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(0, weight=1)

        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=22,
            pady=20,
        )
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(
            content,
            text="Preview exact diagnostic ZIP contents",
            font=("Segoe UI Semibold", 20),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            content,
            text=(
                "Nothing has been saved. The ZIP will contain exactly "
                "the inventory and redacted values shown below."
            ),
            font=("Segoe UI", 11),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=560,
        ).grid(row=1, column=0, sticky="ew", pady=(3, 12))
        preview_box = ctk.CTkTextbox(
            content,
            font=("Consolas", 10),
            fg_color=Theme.BG_INPUT,
            text_color=Theme.TEXT_SECONDARY,
            wrap="word",
        )
        preview_box.grid(row=3, column=0, sticky="nsew")
        preview_box.insert("1.0", preview)
        preview_box.configure(state="disabled")

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        actions.grid_columnconfigure(2, weight=1)
        ctk.CTkButton(
            actions,
            text="Close",
            width=92,
            height=34,
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=dialog.destroy,
        ).grid(row=0, column=0, padx=(0, 8))

        def save_exact_bundle():
            initial_dir = (
                self.output_path
                if os.path.exists(self.output_path)
                else DEFAULT_OUTPUT
            )
            os.makedirs(initial_dir, exist_ok=True)
            bundle_path = filedialog.asksaveasfilename(
                parent=dialog,
                title="Save reviewed diagnostics bundle",
                initialdir=initial_dir,
                initialfile=(
                    "MSStoreHelper-Diagnostics-"
                    f"{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip"
                ),
                defaultextension=".zip",
                filetypes=[
                    ("ZIP archive", "*.zip"),
                    ("All files", "*.*"),
                ],
            )
            if not bundle_path:
                return
            try:
                write_prepared_bundle(bundle_path, entries)
            except (DiagnosticRedactionError, OSError, zipfile.BadZipFile) as exc:
                self._update_status(
                    "Diagnostics export failed",
                    Theme.DANGER,
                )
                self._log(
                    "ERROR",
                    f"Failed to export diagnostics bundle: {exc}",
                )
                return
            dialog.destroy()
            self._update_status(
                "Diagnostics exported",
                Theme.SUCCESS,
            )
            self._log(
                "SUCCESS",
                f"Diagnostics bundle saved: {bundle_path}",
            )

        ctk.CTkButton(
            actions,
            text="Save Exact ZIP",
            width=138,
            height=34,
            font=("Segoe UI Semibold", 11),
            fg_color=Theme.PRIMARY,
            hover_color=Theme.PRIMARY_HOVER,
            command=save_exact_bundle,
        ).grid(row=0, column=1)
        self._update_status(
            "Diagnostics ready for review",
            Theme.INFO,
        )
        dialog.after(50, dialog.focus_force)

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

    def _export_appinstaller_manifest(self):
        if not self.download_queue:
            self._update_status("⚠️ Queue is empty", Theme.WARNING)
            self._log("WARNING", "No queued files to export as App Installer")
            return

        initial_dir = self.output_path if os.path.exists(self.output_path) else DEFAULT_OUTPUT
        os.makedirs(initial_dir, exist_ok=True)
        appinstaller_path = filedialog.asksaveasfilename(
            title="Save App Installer manifest",
            initialdir=initial_dir,
            initialfile="MSStoreHelper-Queue.appinstaller",
            defaultextension=".appinstaller",
            filetypes=[("App Installer manifest", "*.appinstaller"), ("All files", "*.*")],
        )
        if not appinstaller_path:
            return

        try:
            result = StoreAPI.write_appinstaller_export(
                self.download_queue,
                self.output_path,
                appinstaller_path,
                self._target_arch(),
            )
        except ValueError as exc:
            self._update_status("⚠️ AppInstaller export skipped", Theme.WARNING)
            self._log("WARNING", str(exc))
        except Exception as exc:
            self._update_status("❌ AppInstaller export failed", Theme.DANGER)
            self._log("ERROR", f"Failed to export App Installer manifest: {exc}")
        else:
            self._update_status("✅ AppInstaller exported", Theme.SUCCESS)
            self._log(
                "SUCCESS",
                f"App Installer manifest saved: {result['AppInstallerPath']} ({result['PackageCount']} package(s)); packages: {result['PackageDir']}",
            )

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
        self._save_download_state()
        
        total = len(self.download_queue)
        success_count = 0
        queue_ui_refresh_needed = False
        
        for i, pkg in enumerate(self.download_queue):
            try:
                validated_package = validate_package_record(pkg, require_url=True)
                pkg.update(validated_package)
                fname = validated_package["FileName"]
                filepath = confined_package_path(self.output_path, fname)
            except PackageIngressError as exc:
                pkg["DownloadStatus"] = "Failed"
                pkg["LastError"] = str(exc)
                pkg.pop("LocalPath", None)
                self._save_download_state()
                if '_status_widget' in pkg:
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Path blocked", text_color=Theme.DANGER))
                self.after(0, lambda e=str(exc): self._log("ERROR", f"Rejected unsafe package: {e}"))
                continue
            self.after(0, lambda n=fname: self._update_status(f"⬇️ Downloading {n[:40]}...", Theme.INFO))
            self.after(0, lambda n=fname, idx=i+1, tot=total: self._log("INFO", f"[{idx}/{tot}] Downloading: {n}"))
            
            if '_status_widget' in pkg:
                self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Downloading...", text_color=Theme.INFO))
            
            pkg['LocalPath'] = filepath
            pkg["DownloadStatus"] = "Downloading"
            pkg.pop("LastError", None)
            self._save_download_state()
            
            def progress_cb(val, idx=i, tot=total):
                self.after(0, lambda v=(idx + val) / tot: self._update_progress(v))
            
            success, error_msg = StoreAPI.download_file(pkg['Url'], filepath, progress_cb, pkg)
            
            if '_status_widget' in pkg:
                if success:
                    pkg["DownloadStatus"] = "Downloaded"
                    pkg.pop("LastError", None)
                    self._save_download_state()
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="✅ Done", text_color=Theme.SUCCESS))
                    self.after(0, lambda n=fname: self._log("SUCCESS", f"  Downloaded: {n}"))
                    success_count += 1
                    if self.shared_cache_enabled.get():
                        cache_success, cache_msg = StoreAPI.cache_downloaded_artifact(pkg, self.shared_cache_path)
                        self._save_download_state()
                        level = "SUCCESS" if cache_success else "WARNING"
                        self.after(0, lambda lvl=level, m=cache_msg: self._log(lvl, f"  Shared cache: {m}"))
                else:
                    if pkg.get("TrustState") == TRUST_STATE_REVIEW_REQUIRED:
                        pkg["DownloadStatus"] = "Quarantined"
                        queue_ui_refresh_needed = True
                    elif pkg.get("TrustState") == TRUST_STATE_BLOCKED:
                        pkg["DownloadStatus"] = "TrustBlocked"
                    else:
                        pkg["DownloadStatus"] = "Partial" if os.path.exists(f"{filepath}.part") else "Failed"
                    pkg["LastError"] = error_msg
                    self._save_download_state()
                    status_text = {
                        "Quarantined": "Review required",
                        "TrustBlocked": "Trust blocked",
                        "Partial": "Partial",
                    }.get(pkg["DownloadStatus"], "❌ Failed")
                    status_color = (
                        Theme.WARNING
                        if pkg["DownloadStatus"] in {"Partial", "Quarantined"}
                        else Theme.DANGER
                    )
                    self.after(0, lambda w=pkg['_status_widget'], t=status_text, c=status_color: w.configure(text=t, text_color=c))
                    self.after(0, lambda n=fname, e=error_msg: self._log("ERROR", f"  Failed to download {n}: {e}"))

        if queue_ui_refresh_needed:
            self.after(0, self._update_queue_ui)
        self.after(0, lambda: self._update_progress(0))
        self.after(0, lambda: self._update_status("✅ Downloads complete!", Theme.SUCCESS))
        self.after(0, lambda: self._log("SUCCESS", f"Download complete: {success_count}/{total} files successful"))
        self.after(0, lambda: self._log("INFO", f"Files saved to: {self.output_path}"))

    def _rollback_cache_folders(self):
        folders = [self.output_path]
        if self.shared_cache_enabled.get():
            folders.append(self.shared_cache_path)
        for package in self.download_queue:
            manifest_path = package.get("CacheManifest")
            if manifest_path:
                folders.append(os.path.dirname(manifest_path))

        unique = []
        seen = set()
        for folder in folders:
            folder = os.path.abspath(folder)
            if folder.lower() in seen:
                continue
            seen.add(folder.lower())
            unique.append(folder)
        return unique

    def _queued_app_identities(self):
        return [
            (package.get("PackageIdentity") or package_identity(package["FileName"]))
            for package in self.download_queue
            if package.get("FileName") and not is_dependency_package(package)
        ]

    def _show_package_diff(self):
        identities = self._queued_app_identities()
        if not identities:
            self._update_status("⚠️ Queue an app first", Theme.WARNING)
            self._log("WARNING", "Package diff needs at least one queued app package identity")
            return

        cache_folders = self._rollback_cache_folders()
        self._update_status("Comparing cached package versions...", Theme.INFO)
        threading.Thread(target=self._package_diff_worker, args=(identities, cache_folders), daemon=True).start()

    def _package_diff_worker(self, identities, cache_folders):
        identity_set = {identity.lower() for identity in identities if identity}
        candidates = StoreAPI.package_diff_candidates(cache_folders, identity_set)
        if not candidates:
            self.after(0, lambda: self._update_status("No package diff available", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", "Package diff needs two valid cached versions for a queued app identity"))
            return

        sections = []
        for candidate in candidates:
            try:
                diff = StoreAPI.diff_appx_manifests(candidate["Old"]["Path"], candidate["New"]["Path"])
                sections.append(StoreAPI.format_package_diff(diff))
            except Exception as exc:
                identity = candidate.get("PackageIdentity", "package")
                sections.append(f"{identity}\nDiff failed: {exc}")

        report = "\n\n" + ("-" * 60) + "\n\n"
        report = report.join(sections)
        self.after(0, lambda: self._update_status("Package diff ready", Theme.SUCCESS))
        self.after(0, lambda c=len(sections): self._log("SUCCESS", f"Package diff generated for {c} cached package pair(s)"))
        self.after(0, lambda text=report: self._show_package_diff_dialog(text))

    def _show_package_diff_dialog(self, text):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Package Diff")
        dialog.geometry("720x560")
        dialog.transient(self)
        dialog.grab_set()

        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=22, pady=22)
        ctk.CTkLabel(content, text="Package Diff", font=("Segoe UI Semibold", 20), anchor="w").pack(fill="x")
        ctk.CTkLabel(content, text="Cached manifest capability and dependency changes", font=("Segoe UI", 11), text_color=Theme.TEXT_MUTED, anchor="w").pack(fill="x", pady=(2, 12))

        textbox = ctk.CTkTextbox(content, font=("Consolas", 11), fg_color=Theme.BG_DARK, text_color=Theme.TEXT_SECONDARY, wrap="word")
        textbox.pack(fill="both", expand=True)
        textbox.insert("1.0", text)
        textbox.configure(state="disabled")

        ctk.CTkButton(
            content,
            text="Close",
            width=90,
            height=32,
            font=("Segoe UI", 12),
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=dialog.destroy,
        ).pack(side="right", pady=(12, 0))

    def _start_rollback(self):
        if not IS_ADMIN:
            self._update_status("⚠️ Administrator required", Theme.WARNING)
            self._log("WARNING", "Rollback requires Administrator rights")
            return

        identities = self._queued_app_identities()
        if not identities:
            self._update_status("⚠️ Queue an app first", Theme.WARNING)
            self._log("WARNING", "Rollback needs at least one queued app package identity")
            return

        cache_folders = self._rollback_cache_folders()
        threading.Thread(target=self._rollback_worker, args=(identities, cache_folders), daemon=True).start()

    def _rollback_worker(self, identities, cache_folders):
        self.after(0, lambda: self._update_status("Finding cached rollback packages...", Theme.INFO))
        identity_set = {identity.lower() for identity in identities if identity}
        installed_versions = StoreAPI.get_installed_appx_versions()
        current_versions = {}
        for package in self.download_queue:
            if not package.get("FileName") or is_dependency_package(package):
                continue
            identity = (package.get("PackageIdentity") or package_identity(package["FileName"])).lower()
            if identity not in identity_set:
                continue
            current_versions[identity] = (
                installed_versions.get(identity)
                or package.get("AvailableVersion")
                or format_version_tuple(package_version_tuple(package["FileName"]))
            )

        candidates = StoreAPI.rollback_candidates(
            cache_folders,
            identity_set,
            current_versions,
        )
        if not candidates:
            self.after(0, lambda: self._update_status("No rollback package found", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", "No valid cached previous version was found for queued app identities"))
            return

        success_count = 0
        for candidate in candidates:
            path = candidate["Path"]
            identity = candidate["RollbackIdentity"]
            version = candidate.get("RollbackVersion", "unknown")
            current = candidate.get("RollbackCurrentVersion") or "unknown"
            self.after(0, lambda i=identity, v=version: self._update_status(f"Rolling back {i} to {v}...", Theme.INFO))
            self.after(0, lambda i=identity, c=current, v=version, p=path: self._log("INFO", f"Rollback candidate for {i}: current={c}, rollback={v}, path={p}"))

            signature_ok, signature_msg = StoreAPI.verify_package_signature(
                path,
                candidate,
            )
            if not signature_ok:
                self.after(0, lambda i=identity, m=signature_msg: self._log("ERROR", f"Rollback signature check blocked {i}: {m}"))
                continue

            success, message = StoreAPI.rollback_package(
                identity,
                path,
                candidate,
            )
            if success:
                success_count += 1
                self.after(0, lambda i=identity, v=version: self._log("SUCCESS", f"Rolled back {i} to {v}"))
            else:
                self.after(0, lambda i=identity, m=message: self._log("ERROR", f"Rollback failed for {i}: {m}"))

        if success_count:
            self.after(0, lambda c=success_count: self._update_status(f"Rollback complete: {c} package(s)", Theme.SUCCESS))
        else:
            self.after(0, lambda: self._update_status("Rollback failed", Theme.DANGER))
    
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
            try:
                filepath = validate_existing_package_path(
                    pkg["LocalPath"],
                    expected_filename=fname,
                    require_file=True,
                )
            except PackageIngressError as exc:
                if '_status_widget' in pkg:
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Path blocked", text_color=Theme.DANGER))
                self.after(0, lambda n=fname, e=str(exc): self._log("ERROR", f"  Blocked unsafe package path for {n}: {e}"))
                continue
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

            signature_ok, signature_msg = StoreAPI.verify_package_signature(
                filepath,
                pkg,
            )
            if not signature_ok:
                if '_status_widget' in pkg:
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Signature blocked", text_color=Theme.DANGER))
                self.after(0, lambda n=fname, m=signature_msg: self._log("ERROR", f"  Signature verification blocked {n}: {m}"))
                continue

            self.after(0, lambda m=signature_msg: self._log("DEBUG", f"  Signature verified: {m}"))
            
            success, error_msg = StoreAPI.install_package(filepath, pkg)
            
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
        self._inspect_repair_plan("store-repair")

    def _run_provisioning_repair(self):
        self._inspect_repair_plan("provisioning-repair")

    def _run_licensing_reset(self):
        self._inspect_repair_plan("licensing-reset")

    def _run_cache_rebuild(self):
        self._inspect_repair_plan("cache-rebuild")

    def _inspect_repair_plan(self, repair_type):
        if self._repair_operation_active:
            self._update_status(
                "A repair or restore is already running",
                Theme.WARNING,
            )
            return
        try:
            plan = build_repair_plan(
                repair_type,
                backup_base=REPAIR_BACKUP_DIR,
                retention_count=self.repair_retention_var.get(),
            )
        except RepairTransactionError as exc:
            self._update_status("Repair plan unavailable", Theme.DANGER)
            self._log("ERROR", f"Repair plan unavailable: {exc}")
            return
        self._show_repair_plan_dialog(plan, restore=False)

    def _show_repair_plan_dialog(self, plan, *, restore):
        dialog = ctk.CTkToplevel(self)
        dialog.title(
            "Inspect restore plan" if restore else "Inspect repair plan"
        )
        dialog.geometry("780x680")
        dialog.minsize(620, 500)
        dialog.transient(self)
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(0, weight=1)

        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=22,
            pady=20,
        )
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(2, weight=1)

        title = (
            "Restore captured Windows state"
            if restore
            else plan["DisplayName"]
        )
        ctk.CTkLabel(
            content,
            text=title,
            font=("Segoe UI Semibold", 20),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        summary = (
            "Review the exact targets and verification contract. "
            "The backup remains available after restore."
            if restore
            else (
                "Nothing has run. Review every precondition, backup, "
                "mutation, permission, and reboot impact first."
            )
        )
        ctk.CTkLabel(
            content,
            text=summary,
            font=("Segoe UI", 11),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=540,
        ).grid(row=1, column=0, sticky="ew", pady=(3, 12))

        plan_box = ctk.CTkTextbox(
            content,
            font=("Consolas", 10),
            fg_color=Theme.BG_INPUT,
            text_color=Theme.TEXT_SECONDARY,
            wrap="word",
        )
        plan_box.grid(row=2, column=0, sticky="nsew")

        def render_plan_text():
            text = (
                render_restore_plan(plan)
                if restore
                else render_repair_plan(plan)
            )
            plan_box.configure(state="normal")
            plan_box.delete("1.0", "end")
            plan_box.insert("1.0", text)
            plan_box.configure(state="disabled")

        render_plan_text()
        controls = ctk.CTkFrame(content, fg_color="transparent")
        controls.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        controls.grid_columnconfigure(0, weight=1)

        if not restore:
            retention_row = ctk.CTkFrame(
                controls,
                fg_color="transparent",
            )
            retention_row.grid(row=0, column=0, sticky="ew")
            ctk.CTkLabel(
                retention_row,
                text="Keep verified backups",
                font=("Segoe UI", 11),
                text_color=Theme.TEXT_SECONDARY,
            ).pack(side="left")

            def change_retention(value):
                plan["RetentionCount"] = normalize_retention(value)
                self.repair_retention_var.set(
                    str(plan["RetentionCount"])
                )
                render_plan_text()

            ctk.CTkOptionMenu(
                retention_row,
                values=["1", "3", "5", "10", "20", "50"],
                variable=self.repair_retention_var,
                width=72,
                height=28,
                command=change_retention,
            ).pack(side="right")

        accepted = ctk.BooleanVar(value=False)
        acknowledgement = ctk.CTkCheckBox(
            controls,
            text=(
                "I reviewed this exact plan and accept the Windows changes."
            ),
            variable=accepted,
            font=("Segoe UI", 11),
            fg_color=Theme.PRIMARY,
            hover_color=Theme.PRIMARY_HOVER,
        )
        acknowledgement.grid(
            row=1,
            column=0,
            sticky="w",
            pady=(10, 8),
        )

        actions = ctk.CTkFrame(controls, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(
            actions,
            text="Close",
            width=92,
            height=34,
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=dialog.destroy,
        ).grid(row=0, column=1, padx=(0, 8))
        confirm_button = ctk.CTkButton(
            actions,
            text="Run Verified Restore" if restore else "Run Verified Repair",
            width=156,
            height=34,
            fg_color=Theme.DANGER,
            hover_color=Theme.DANGER_HOVER,
            state="disabled",
        )
        confirm_button.grid(row=0, column=2)

        def update_confirmation_state():
            confirm_button.configure(
                state="normal" if accepted.get() else "disabled"
            )

        acknowledgement.configure(command=update_confirmation_state)

        def confirm():
            if not accepted.get():
                return
            if plan["RequiresAdmin"] and not IS_ADMIN:
                self._update_status(
                    "Administrator access is required",
                    Theme.WARNING,
                )
                self._log(
                    "WARNING",
                    "Restart MSStoreHelper as Administrator to run this plan.",
                )
                dialog.destroy()
                return
            if not restore:
                self.user_profile["RepairRetentionCount"] = (
                    plan["RetentionCount"]
                )
                self._save_user_profile()
            dialog.destroy()
            self._start_repair_transaction(plan, restore=restore)

        confirm_button.configure(command=confirm)
        dialog.after(50, dialog.focus_force)

    def _choose_repair_restore(self):
        if self._repair_operation_active:
            self._update_status(
                "A repair or restore is already running",
                Theme.WARNING,
            )
            return
        backups = list_repair_backups(REPAIR_BACKUP_DIR)
        if not backups:
            self._update_status(
                "No verified repair backups are available",
                Theme.WARNING,
            )
            self._log(
                "INFO",
                f"No restorable transactions found in {REPAIR_BACKUP_DIR}",
            )
            return

        dialog = ctk.CTkToplevel(self)
        dialog.title("Choose repair backup")
        dialog.geometry("700x520")
        dialog.minsize(620, 460)
        dialog.transient(self)
        dialog.grab_set()
        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=22, pady=20)
        ctk.CTkLabel(
            content,
            text="Choose a verified repair backup",
            font=("Segoe UI Semibold", 20),
            anchor="w",
        ).pack(fill="x")
        ctk.CTkLabel(
            content,
            text=(
                "Selecting a backup opens its exact restore plan. "
                "No state changes occur from this screen."
            ),
            font=("Segoe UI", 11),
            text_color=Theme.TEXT_MUTED,
            anchor="w",
        ).pack(fill="x", pady=(3, 12))
        backup_list = ctk.CTkScrollableFrame(
            content,
            fg_color=Theme.BG_INPUT,
            border_width=1,
            border_color=Theme.BORDER_SUBTLE,
        )
        backup_list.pack(fill="both", expand=True)

        def select_backup(root):
            dialog.destroy()
            self._prepare_repair_restore(root)

        for backup in backups:
            completed = str(backup.get("CompletedAt") or "Unknown date")
            operation = str(backup.get("OperationId") or "")[:8]
            label = (
                f"{backup['RepairName']}\n"
                f"{completed}  •  {backup['Outcome']}  •  {operation}"
            )
            ctk.CTkButton(
                backup_list,
                text=label,
                height=54,
                font=("Segoe UI", 11),
                anchor="w",
                fg_color=Theme.BG_CARD,
                text_color=Theme.TEXT_PRIMARY,
                hover_color=Theme.BG_CARD_HOVER,
                border_width=1,
                border_color=Theme.BORDER_SUBTLE,
                command=lambda root=backup["BackupRoot"]: (
                    select_backup(root)
                ),
            ).pack(fill="x", padx=6, pady=4)

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.pack(fill="x", pady=(10, 0))

        def browse_backup():
            root = filedialog.askdirectory(
                parent=dialog,
                initialdir=REPAIR_BACKUP_DIR,
                title="Choose an MSStoreHelper repair backup",
            )
            if root:
                select_backup(root)

        ctk.CTkButton(
            actions,
            text="Browse another folder",
            width=150,
            height=32,
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=browse_backup,
        ).pack(side="left")
        ctk.CTkButton(
            actions,
            text="Close",
            width=88,
            height=32,
            fg_color="transparent",
            text_color=Theme.TEXT_PRIMARY,
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=dialog.destroy,
        ).pack(side="right")

    def _prepare_repair_restore(self, backup_root):
        try:
            plan = build_restore_plan(
                backup_root,
                backup_base=REPAIR_BACKUP_DIR,
            )
        except RepairTransactionError as exc:
            self._update_status("Backup cannot be restored", Theme.DANGER)
            self._log("ERROR", f"Backup cannot be restored: {exc}")
            return
        self._show_repair_plan_dialog(plan, restore=True)

    def _set_repair_controls_running(self, running):
        self._repair_operation_active = running
        for button in self._repair_buttons:
            button.configure(state="disabled" if running else "normal")
        self.repair_cancel_button.configure(
            state="normal" if running else "disabled",
            text_color=Theme.WARNING if running else Theme.TEXT_MUTED,
            border_color=Theme.WARNING if running else Theme.BORDER_SUBTLE,
        )

    def _start_repair_transaction(self, plan, *, restore):
        if self._repair_operation_active:
            return
        self._repair_cancel_event = threading.Event()
        self._set_repair_controls_running(True)
        action = "restore" if restore else "repair"
        self._update_status(
            f"Running verified {action} transaction…",
            Theme.INFO,
        )
        self._log(
            "INFO",
            (
                f"{plan['DisplayName']} started "
                f"(operation {plan['OperationId']})"
            ),
        )
        threading.Thread(
            target=self._repair_transaction_worker,
            args=(plan, restore),
            daemon=True,
        ).start()

    def _cancel_repair_operation(self):
        if not self._repair_operation_active or not self._repair_cancel_event:
            return
        self._repair_cancel_event.set()
        self.repair_cancel_button.configure(state="disabled")
        self._update_status(
            "Cancellation requested; waiting for a safe checkpoint…",
            Theme.WARNING,
        )
        self._log(
            "WARNING",
            "Cancellation requested; the active step will finish first.",
        )

    def _repair_transaction_worker(self, plan, restore):
        def log_callback(message):
            self.after(
                0,
                lambda value=message: self._log("INFO", value),
            )

        def progress_callback(value):
            self.after(
                0,
                lambda progress=value: self._update_progress(progress),
            )

        try:
            if restore:
                context = execute_restore_plan(
                    plan,
                    confirmation_token=plan["ConfirmationToken"],
                    powershell_exe=POWERSHELL_EXE,
                    is_admin=IS_ADMIN,
                    cancel_event=self._repair_cancel_event,
                    log_callback=log_callback,
                    progress_callback=progress_callback,
                )
            else:
                context = execute_repair_plan(
                    plan,
                    confirmation_token=plan["ConfirmationToken"],
                    powershell_exe=POWERSHELL_EXE,
                    is_admin=IS_ADMIN,
                    cancel_event=self._repair_cancel_event,
                    log_callback=log_callback,
                    progress_callback=progress_callback,
                )
        except Exception as exc:
            context = {
                "Outcome": "failed",
                "Results": [{
                    "Description": "Transaction setup",
                    "Success": False,
                    "Stderr": str(exc),
                }],
            }
        self.after(
            0,
            lambda value=context, is_restore=restore: (
                self._finish_repair_transaction(value, is_restore)
            ),
        )

    def _finish_repair_transaction(self, context, restore):
        self._update_progress(0)
        self._set_repair_controls_running(False)
        self._repair_cancel_event = None
        outcome = context.get("Outcome", "failed")
        action = "Restore" if restore else "Repair"
        backup_root = context.get("BackupRoot")
        if backup_root:
            self._log("INFO", f"Verified backup: {backup_root}")
        for result in context.get("Results", []):
            if result.get("Success"):
                continue
            description = result.get("Description", "Transaction step")
            detail = result.get("Stderr") or result.get("Stdout") or ""
            self._log("ERROR", f"{description}: {detail}")
        if outcome == "succeeded":
            self._update_status(
                f"{action} transaction verified",
                Theme.SUCCESS,
            )
            self._log(
                "SUCCESS",
                f"{action} completed with all postconditions verified.",
            )
        elif outcome.startswith("cancelled"):
            self._update_status(
                f"{action} stopped at a safe checkpoint",
                Theme.WARNING,
            )
            self._log(
                "WARNING",
                f"{action} outcome: {outcome}",
            )
        else:
            self._update_status(
                f"{action} stopped: {outcome}",
                Theme.DANGER,
            )
            self._log(
                "ERROR",
                (
                    f"{action} stopped fail-closed ({outcome}). "
                    "Use Restore Backup if mutation began."
                ),
            )

    def _log_repair_results(self, title, results):
        success_count = sum(1 for result in results if result.get("Success"))
        backup_root = next((result.get("BackupRoot") for result in results if result.get("BackupRoot")), None)
        restore_script = next((result.get("RestoreScriptPath") for result in results if result.get("RestoreScriptPath")), None)

        self._log("INFO", f"{title} results:")
        if backup_root:
            self._log("INFO", f"  Backup manifest: {backup_root}")
        if restore_script:
            self._log("INFO", f"  Restore script: {restore_script}")

        for result in results:
            desc = result.get("Description", "Repair step")
            if result.get("Success"):
                self._log("SUCCESS", f"  ✓ {desc}")
                continue

            self._log("ERROR", f"  ✗ {desc} (exit {result.get('ReturnCode')})")
            command = result.get("Command")
            if command:
                self._log("ERROR", f"    Command: {command}")
            for label in ("Stdout", "Stderr"):
                output = result.get(label, "")
                if not output:
                    continue
                for line in output.splitlines()[:8]:
                    self._log("ERROR", f"    {label.lower()}: {line}")
        return success_count

    def _repair_worker(self):
        self.after(0, lambda: self._log("INFO", "Starting Microsoft Store repair..."))
        
        def log_cb(msg):
            self.after(0, lambda m=msg: self._update_status(m, Theme.INFO))
            self.after(0, lambda m=msg: self._log("INFO", m))
        
        def progress_cb(val):
            self.after(0, lambda v=val: self._update_progress(v))
        
        results = StoreAPI.run_repair(log_cb, progress_cb)
        self.after(0, lambda: self._update_progress(0))
        success_count = sum(1 for result in results if result.get("Success"))
        self.after(0, lambda r=results: self._log_repair_results("Repair", r))
        
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
        self.after(0, lambda: self._update_progress(0))
        success_count = sum(1 for result in results if result.get("Success"))
        self.after(0, lambda r=results: self._log_repair_results("Provisioning repair", r))

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
        self.after(0, lambda: self._update_progress(0))
        success_count = sum(1 for result in results if result.get("Success"))
        self.after(0, lambda r=results: self._log_repair_results("Licensing reset", r))

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
        self.after(0, lambda: self._update_progress(0))
        success_count = sum(1 for result in results if result.get("Success"))
        self.after(0, lambda r=results: self._log_repair_results("Cache rebuild", r))

        if success_count == len(results):
            self.after(0, lambda: self._update_status("✅ Cache rebuild complete", Theme.SUCCESS))
            self.after(0, lambda: self._log("SUCCESS", "Store cache rebuild complete. Previous caches were kept as .bak folders."))
        else:
            self.after(0, lambda: self._update_status(f"⚠️ Cache rebuild done ({success_count}/{len(results)} steps)", Theme.WARNING))
            self.after(0, lambda: self._log("WARNING", f"Cache rebuild partially complete: {success_count}/{len(results)} steps succeeded"))


def build_cli_parser():
    parser = argparse.ArgumentParser(
        prog="MSStoreHelper.py",
        description="Headless Microsoft Store package search, download, and install workflow.",
    )
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--search", metavar="QUERY", help="Search Microsoft Store without opening the GUI.")
    action.add_argument("--download", metavar="APP_OR_PRODUCT_ID", help="Resolve, select, and download Store packages.")
    action.add_argument("--install", metavar="APP_OR_PRODUCT_ID", help="Download and install selected Store packages.")
    action.add_argument("--mirror", metavar="CACHE_FOLDER", help="Serve cached packages over local HTTP without contacting Store services.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help=f"Download folder. Default: {DEFAULT_OUTPUT}")
    parser.add_argument("--arch", choices=["auto", "x64", "x86", "arm64", "arm", "neutral"], default="auto")
    parser.add_argument("--ring", choices=STORE_RING_VALUES, default="Retail")
    parser.add_argument("--language", default="en-US")
    parser.add_argument("--market", default="US")
    parser.add_argument("--host", default="127.0.0.1", help="Mirror bind host. Loopback is the safe default.")
    parser.add_argument("--port", type=int, default=8765, help="Mirror bind port.")
    parser.add_argument("--advertise-host", help="Client-facing mirror host. Required when --host is a wildcard address.")
    parser.add_argument("--lan", action="store_true", help="Explicitly allow a non-loopback mirror with bearer authentication.")
    parser.add_argument("--acknowledge-cleartext-risk", action="store_true", help="Allow authenticated LAN HTTP without TLS.")
    parser.add_argument("--tls-cert", help="PEM certificate for an HTTPS mirror.")
    parser.add_argument("--tls-key", help="PEM private key for an HTTPS mirror.")
    parser.add_argument("--mirror-token-ttl", type=int, default=900, help="LAN bearer-token lifetime in seconds (60-3600).")
    parser.add_argument("--mirror-index-only", action="store_true", help="Write the mirror index and exit instead of serving forever.")
    parser.add_argument("--json", action="store_true", help="Emit a machine-readable JSON summary.")
    return parser


def _cli_print(message, stream):
    print(message, file=stream)


def _cli_emit_summary(summary, as_json, stdout):
    if as_json:
        _cli_print(json.dumps(summary, indent=2), stdout)
        return

    action = summary.get("Action", "")
    app = summary.get("App", {})
    if app:
        _cli_print(f"{action}: {app.get('Name')} ({app.get('ProductId')})", stdout)
    else:
        _cli_print(action, stdout)

    for item in summary.get("Packages", []):
        name = item.get("FileName", "")
        status = item.get("Status", "")
        message = item.get("Message", "")
        if status:
            _cli_print(f"- {status}: {name} {message}".rstrip(), stdout)
        else:
            detail = item.get("Url", message)
            _cli_print(f"- {name} {detail}".rstrip(), stdout)


def _cli_package_record(package, status, message="", path=None):
    trust_report = package.get("TrustReport") or {}
    return {
        "FileName": package.get("FileName", ""),
        "PackageIdentity": package.get("PackageIdentity") or package_identity(package.get("FileName", "")),
        "Version": package.get("AvailableVersion") or format_version_tuple(package_version_tuple(package.get("FileName", ""))),
        "Architecture": package.get("Architecture", "neutral"),
        "FileType": package.get("FileType", ""),
        "Status": status,
        "Message": message,
        "LocalPath": path or package.get("LocalPath", ""),
        "TrustState": package.get("TrustState", ""),
        "TrustReviewEligible": trust_report.get("ReviewEligible", False),
        "TrustReasonCodes": list(trust_report.get("ReasonCodes") or []),
    }


def _cli_set_package_record(records, package, status, message="", path=None):
    updated = _cli_package_record(package, status, message, path)
    filename = updated.get("FileName", "")
    for index, record in enumerate(records):
        if record.get("FileName") == filename:
            records[index] = updated
            return
    records.append(updated)


def _cli_search(query, args, stdout, stderr):
    diagnostic = StoreAPI.search_store_with_diagnostics(query)
    results = diagnostic.get("Results", [])
    summary = {
        "Action": "search",
        "Query": query,
        "Source": diagnostic.get("Source", "Microsoft Store Search API"),
        "Errors": diagnostic.get("Errors", []),
        "Results": results,
    }
    for error in summary["Errors"]:
        _cli_print(f"warning: {error}", stderr)

    if args.json:
        _cli_print(json.dumps(summary, indent=2), stdout)
    else:
        _cli_print(f"Search: {query}", stdout)
        for result in results:
            _cli_print(f"- {result.get('Name', 'Unknown')} ({result.get('ProductId', '')})", stdout)

    return 0 if results else 1


def _cli_download_selected(packages, output_path, stderr):
    records = []
    downloaded = []
    os.makedirs(output_path, exist_ok=True)
    for package in packages:
        try:
            package = validate_package_record(package, require_url=True)
            filename = package["FileName"]
            path = confined_package_path(output_path, filename)
        except PackageIngressError as exc:
            records.append(_cli_package_record(
                package if isinstance(package, dict) else {},
                "failed",
                str(exc),
            ))
            _cli_print(f"download blocked: {exc}", stderr)
            continue
        ok, message = StoreAPI.download_file(package.get("Url", ""), path, package=package)
        if ok:
            package["LocalPath"] = path
            downloaded.append(package)
            records.append(_cli_package_record(package, "downloaded", message, path))
        else:
            _cli_print(f"download failed: {filename}: {message}", stderr)
            status = {
                TRUST_STATE_REVIEW_REQUIRED: "quarantined",
                TRUST_STATE_BLOCKED: "trust-blocked",
            }.get(package.get("TrustState"), "failed")
            records.append(_cli_package_record(package, status, message, path))
    return downloaded, records


def _cli_install_downloaded(packages, records, stderr):
    success_count = 0
    skipped_count = 0
    failed_count = 0

    for package in packages:
        try:
            package = validate_package_record(package, require_url=False)
            local_path = validate_existing_package_path(
                package.get("LocalPath"),
                expected_filename=package["FileName"],
                require_file=True,
            )
        except PackageIngressError as exc:
            failed_count += 1
            _cli_set_package_record(
                records,
                package if isinstance(package, dict) else {},
                "failed",
                str(exc),
                "",
            )
            continue

        should_skip, installed_version, package_name = StoreAPI.should_skip_installed_package(package)
        if should_skip:
            skipped_count += 1
            _cli_set_package_record(records, package, "skipped", f"installed {installed_version} is current", local_path)
            continue

        signature_ok, signature_message = StoreAPI.verify_package_signature(
            local_path,
            package,
        )
        if not signature_ok:
            failed_count += 1
            _cli_print(f"signature blocked: {package.get('FileName')}: {signature_message}", stderr)
            _cli_set_package_record(records, package, "failed", f"signature blocked: {signature_message}", local_path)
            continue

        ok, install_message = StoreAPI.install_package(local_path, package)
        if ok:
            success_count += 1
            _cli_set_package_record(records, package, "installed", install_message, local_path)
            continue

        if StoreAPI.is_noop_install_error(install_message):
            skipped_count += 1
            _cli_set_package_record(records, package, "skipped", f"{package_name}: {install_message}", local_path)
            continue

        failed_count += 1
        _cli_print(f"install failed: {package.get('FileName')}: {install_message}", stderr)
        _cli_set_package_record(records, package, "failed", install_message, local_path)

    return success_count, skipped_count, failed_count


def _cli_package_workflow(args, stdout, stderr):
    identifier = args.download or args.install
    app, resolve_error = StoreAPI.resolve_cli_app(identifier)
    if not app:
        _cli_print(f"error: {resolve_error}", stderr)
        return 1

    target_arch = SYSTEM_ARCH if args.arch == "auto" else args.arch
    diagnostic = StoreAPI.get_packages_with_diagnostics(app["ProductId"], args.ring, args.language, args.market)
    errors = diagnostic.get("Errors", [])
    for error in errors:
        _cli_print(f"warning: {error}", stderr)

    selected = StoreAPI.smart_select(
        diagnostic.get("Packages", []),
        target_arch,
        prefer_exact_arch=args.arch != "auto",
    )
    selected = StoreAPI.order_packages_for_install(selected, target_arch)
    if not selected:
        _cli_print(f"error: no installable packages were selected for {app.get('Name')}", stderr)
        return 1

    downloaded, records = _cli_download_selected(selected, os.path.abspath(args.output), stderr)
    action = "install" if args.install else "download"
    summary = {
        "Action": action,
        "App": app,
        "TargetArchitecture": target_arch,
        "OutputPath": os.path.abspath(args.output),
        "StoreQuery": diagnostic.get("Query", StoreAPI.package_query_metadata(app["ProductId"], args.ring, args.language, args.market)),
        "Errors": errors,
        "Packages": records,
    }

    if len(downloaded) != len(selected):
        _cli_emit_summary(summary, args.json, stdout)
        return 1

    if args.install:
        if not IS_ADMIN:
            _cli_print("error: administrator rights are required for --install", stderr)
            _cli_emit_summary(summary, args.json, stdout)
            return 2
        installed, skipped, failed = _cli_install_downloaded(downloaded, records, stderr)
        summary["Installed"] = installed
        summary["Skipped"] = skipped
        summary["Failed"] = failed
        _cli_emit_summary(summary, args.json, stdout)
        return 0 if failed == 0 else 1

    _cli_emit_summary(summary, args.json, stdout)
    return 0


def _cli_mirror(args, stdout, stderr):
    folder = os.path.abspath(args.mirror)
    try:
        policy = validate_network_policy(
            args.host,
            advertised_host=args.advertise_host,
            lan_mode=args.lan,
            acknowledge_cleartext=args.acknowledge_cleartext_risk,
            tls_cert=args.tls_cert,
            tls_key=args.tls_key,
        )
    except MirrorConfigurationError as exc:
        _cli_print(f"error: {exc}", stderr)
        return 2

    if args.mirror_index_only:
        index = StoreAPI.write_mirror_index(
            folder,
            policy["AdvertisedHost"],
            args.port,
            tls_enabled=policy["TlsEnabled"],
        )
        summary = {
            "Action": "mirror-index",
            "AdvertisedHost": policy["AdvertisedHost"],
            "Port": int(args.port),
            **index,
        }
        _cli_emit_summary(summary, args.json, stdout)
        return 0 if index.get("PackageCount", 0) > 0 else 1

    try:
        server, index = StoreAPI.create_mirror_server(
            folder,
            args.host,
            args.port,
            advertised_host=policy["AdvertisedHost"],
            lan_mode=args.lan,
            acknowledge_cleartext=args.acknowledge_cleartext_risk,
            tls_cert=args.tls_cert,
            tls_key=args.tls_key,
            token_ttl_seconds=args.mirror_token_ttl,
        )
    except (MirrorConfigurationError, OSError, ssl.SSLError) as exc:
        _cli_print(f"error: {exc}", stderr)
        return 2
    actual_port = int(server.server_address[1])
    summary = {
        "Action": "mirror",
        "BindHost": policy["BindHost"],
        "AdvertisedHost": policy["AdvertisedHost"],
        "Port": actual_port,
        "TlsEnabled": policy["TlsEnabled"],
        "LanMode": policy["LanMode"],
        "AuditLog": MIRROR_AUDIT_FILENAME,
        **index,
    }
    if index.get("PackageCount", 0) == 0:
        _cli_print(f"error: no cacheable AppX/MSIX packages found in {folder}", stderr)
        server.server_close()
        _cli_emit_summary(summary, args.json, stdout)
        return 1

    if server.mirror_bearer_token:
        _cli_print(
            (
                "LAN bearer token (shown once; do not place it in URLs "
                f"or logs): {server.mirror_bearer_token}"
            ),
            stderr,
        )
        _cli_print(
            f"token expires: {server.mirror_token_expires_at}",
            stderr,
        )
    _cli_emit_summary(summary, args.json, stdout)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        _cli_print("mirror stopped", stderr)
    finally:
        server.server_close()
    return 0


def run_cli(argv=None, stdout=None, stderr=None):
    stdout = stdout or sys.stdout
    stderr = stderr or sys.stderr
    parser = build_cli_parser()
    args = parser.parse_args(argv)
    if args.search:
        return _cli_search(args.search, args, stdout, stderr)
    if args.mirror:
        return _cli_mirror(args, stdout, stderr)
    return _cli_package_workflow(args, stdout, stderr)


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv:
        return run_cli(argv)

    app = MSStoreHelperApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
