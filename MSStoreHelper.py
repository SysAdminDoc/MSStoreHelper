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
import re
import hashlib
import shutil
import tempfile
import zipfile
try:
    import winreg
except ImportError:
    winreg = None
from tkinter import filedialog
from datetime import datetime, timezone
from msstore_package_resolution import (
    annotate_package,
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
    signature_info_is_valid_microsoft,
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

APP_VERSION = "3.26.0"
APP_NAME = "MSStoreHelper"
API_URL = "https://store.rg-adguard.net/api/GetFiles"
STORE_SEARCH_URL = "https://storeedgefd.dsx.mp.microsoft.com/v9.0/manifestSearch"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
CACHE_MANIFEST_NAME = "msstorehelper-cache-manifest.json"
WINGET_IMPORT_SCHEMA = "https://aka.ms/winget-packages.schema.2.0.json"
WINGET_MSSTORE_SOURCE = {
    "Argument": "https://storeedgefd.dsx.mp.microsoft.com/v9.0",
    "Identifier": "StoreEdgeFD",
    "Name": "msstore",
    "Type": "Microsoft.Rest",
}
THEME_MODE_VALUES = ["System", "Dark", "Light"]
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
    DEFAULT_ACCENT = "#6366f1"
    MODE = "Dark"

    # Main colors
    BG_DARK = ("#f5f7fb", "#0f0f1a")
    BG_CARD = ("#ffffff", "#1a1a2e")
    BG_CARD_HOVER = ("#e8eef8", "#252542")
    BG_INPUT = ("#edf2f8", "#16213e")
    
    # Accent colors
    PRIMARY = DEFAULT_ACCENT
    PRIMARY_HOVER = ("#4f46e5", "#818cf8")
    PRIMARY_OUTLINE_TEXT = ("#4338ca", "#c4b5fd")
    SUCCESS = ("#059669", "#10b981")
    SUCCESS_HOVER = ("#047857", "#34d399")
    WARNING = ("#b45309", "#f59e0b")
    DANGER = ("#dc2626", "#ef4444")
    DANGER_HOVER = ("#b91c1c", "#f87171")
    INFO = ("#0891b2", "#06b6d4")
    
    # Text colors
    TEXT_PRIMARY = ("#0f172a", "#f8fafc")
    TEXT_SECONDARY = ("#475569", "#94a3b8")
    TEXT_MUTED = ("#64748b", "#64748b")
    
    # Special
    BORDER = ("#cbd5e1", "#2a2a4a")
    BUNDLE_COLOR = ("#0891b2", "#22d3ee")
    ENCRYPTED_COLOR = ("#dc2626", "#f87171")
    ARCH_MATCH = ("#16a34a", "#4ade80")

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
        return {"SearchHistory": [], "PinnedFavorites": [], "ThemeMode": "System"}

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
    def fetch_release_notes(product_id):
        url = f"https://apps.microsoft.com/detail/{product_id}?hl=en-US&gl=US"
        response, _retry_errors = request_with_retries(
            "Microsoft Store product page",
            lambda: requests.get(url, timeout=20, headers={"User-Agent": USER_AGENT}),
            attempts=2,
        )
        return StoreAPI.parse_release_notes_html(product_id, response.text, response.url)
    
    @staticmethod
    def get_packages(product_id, ring="Retail"):
        return StoreAPI.get_packages_with_diagnostics(product_id, ring)["Packages"]

    @staticmethod
    def get_packages_with_diagnostics(product_id, ring="Retail"):
        """Get downloadable packages for a product"""
        payload = {"type": "ProductId", "url": product_id, "ring": ring, "lang": "en-US"}
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
                return StoreAPI._source_diagnostic(
                    "RG-Adguard package proxy",
                    errors=retry_errors + ["RG-Adguard response did not include a package table"],
                    fallbacks=package_lookup_fallbacks(product_id, statuses),
                )

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
            
            return StoreAPI._source_diagnostic("RG-Adguard package proxy", packages=results, errors=retry_errors)
            
        except StoreSourceError as exc:
            statuses = StoreAPI.detect_source_health()
            return StoreAPI._source_diagnostic(
                exc.source_name,
                errors=exc.errors,
                fallbacks=package_lookup_fallbacks(product_id, statuses),
            )
        except Exception as exc:
            statuses = StoreAPI.detect_source_health()
            return StoreAPI._source_diagnostic(
                "RG-Adguard package proxy",
                errors=[f"{type(exc).__name__}: {exc}"],
                fallbacks=package_lookup_fallbacks(product_id, statuses),
            )
    
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
    def artifact_metadata(package, artifact_path, source_url=None):
        filename = package.get("FileName") or os.path.basename(artifact_path)
        metadata = {
            "FileName": filename,
            "Path": os.path.abspath(artifact_path),
            "SizeBytes": os.path.getsize(artifact_path),
            "Sha256": StoreAPI.file_sha256(artifact_path),
            "Url": source_url or package.get("Url", ""),
            "PackageIdentity": package_identity(filename),
            "AvailableVersion": format_version_tuple(package_version_tuple(filename)),
        }
        if package.get("PackageRoleLabel"):
            metadata["PackageRoleLabel"] = package.get("PackageRoleLabel")
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
                return data
        except Exception:
            pass
        return {"Version": 1, "Artifacts": {}}

    @staticmethod
    def save_cache_manifest(folder, manifest):
        os.makedirs(folder, exist_ok=True)
        manifest["Version"] = 1
        manifest["UpdatedAt"] = datetime.now(timezone.utc).isoformat()
        with open(StoreAPI._manifest_path(folder), "w", encoding="utf-8") as handle:
            json.dump(manifest, handle, indent=2, sort_keys=True)

    @staticmethod
    def write_artifact_manifest(package, artifact_path, manifest_folder=None, source_url=None):
        folder = manifest_folder or os.path.dirname(os.path.abspath(artifact_path))
        metadata = StoreAPI.artifact_metadata(package, artifact_path, source_url)
        manifest = StoreAPI.load_cache_manifest(folder)
        manifest["Artifacts"][metadata["FileName"]] = metadata
        StoreAPI.save_cache_manifest(folder, manifest)
        package["LocalPath"] = artifact_path
        package["SizeBytes"] = metadata["SizeBytes"]
        package["Sha256"] = metadata["Sha256"]
        package["CacheManifest"] = StoreAPI._manifest_path(folder)
        return metadata

    @staticmethod
    def redact_diagnostic_text(text):
        redacted = str(text or "")
        path_tokens = {
            "USERPROFILE": os.environ.get("USERPROFILE"),
            "APPDATA": os.environ.get("APPDATA"),
            "LOCALAPPDATA": os.environ.get("LOCALAPPDATA"),
            "TEMP": tempfile.gettempdir(),
            "APP_DATA": APP_DATA_DIR,
        }
        for label, path in path_tokens.items():
            if path:
                redacted = redacted.replace(path, f"%{label}%")
                redacted = redacted.replace(path.replace("\\", "/"), f"%{label}%")

        secret_patterns = [
            r"(?i)(authorization\s*[:=]\s*)([^\s;]+)",
            r"(?i)((?:api[_-]?key|password|secret|token)\s*[:=]\s*)([^\s;]+)",
        ]
        for pattern in secret_patterns:
            redacted = re.sub(pattern, r"\1[REDACTED]", redacted)
        return redacted

    @staticmethod
    def diagnostic_queue_metadata(queue):
        allowed_keys = [
            "FileName", "PackageIdentity", "PackageRoleLabel", "Architecture", "FileType",
            "IsBundle", "IsEncrypted", "SizeBytes", "SizeStr", "Sha256", "AvailableVersion",
            "LocalPath", "CacheManifest",
        ]
        items = []
        for package in queue or []:
            item = {key: package.get(key) for key in allowed_keys if key in package}
            for key in ("LocalPath", "CacheManifest"):
                if key in item:
                    item[key] = StoreAPI.redact_diagnostic_text(item[key])
            items.append(item)
        return items

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
            text = StoreAPI.redact_diagnostic_text(json.dumps(data, indent=2))
            try:
                recent.append(json.loads(text))
            except json.JSONDecodeError:
                recent.append({"Raw": text})
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
    def write_diagnostics_bundle(bundle_path, app_version, system_arch, is_admin, output_path, source_health, queue, log_text):
        os.makedirs(os.path.dirname(os.path.abspath(bundle_path)), exist_ok=True)
        redacted_log = StoreAPI.redact_diagnostic_text(log_text)
        redacted_transcript = StoreAPI.redact_diagnostic_text(StoreAPI.powershell_transcript_from_log(log_text))
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
                "Executable": StoreAPI.redact_diagnostic_text(sys.executable),
            },
            "SystemArchitecture": system_arch,
            "IsAdmin": bool(is_admin),
            "OutputPath": StoreAPI.redact_diagnostic_text(output_path),
            "SourceHealth": source_health or [],
            "QueueCount": len(queue_metadata),
            "RepairManifestCount": len(repair_manifests),
        }

        with zipfile.ZipFile(bundle_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("diagnostics.json", json.dumps(diagnostics, indent=2))
            archive.writestr("source-health.json", json.dumps(source_health or [], indent=2))
            archive.writestr("queue.json", json.dumps(queue_metadata, indent=2))
            archive.writestr("app-log.txt", redacted_log)
            archive.writestr("powershell-transcript.txt", redacted_transcript)
            archive.writestr("repair-manifests.json", json.dumps(repair_manifests, indent=2))
        return bundle_path

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
    def download_file(url, filepath, progress_callback=None, package=None):
        part_path = f"{filepath}.part"
        try:
            os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)
            with requests.get(url, stream=True, timeout=60) as r:
                r.raise_for_status()
                total = int(r.headers.get('content-length', 0))
                
                with open(part_path, 'wb') as f:
                    downloaded = 0
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
            if package is not None:
                StoreAPI.write_artifact_manifest(package, filepath, source_url=url)
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
        super().__init__(master, fg_color=Theme.BG_CARD, corner_radius=12, border_width=1, border_color=Theme.BORDER, **kwargs)


class AppTile(ctk.CTkFrame):
    def __init__(self, master, app_data, on_select, on_release_notes):
        super().__init__(master, fg_color="transparent")
        
        self.app_data = app_data
        self.on_select = on_select
        self.on_release_notes = on_release_notes
        self.selected = ctk.BooleanVar(value=False)
        
        self.container = ctk.CTkFrame(self, fg_color=Theme.BG_CARD, corner_radius=10, border_width=1, border_color=Theme.BORDER)
        self.container.pack(fill="x", pady=4, padx=2)
        self.container.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.container, text=app_data.get("Icon", "📦"), font=("Segoe UI Emoji", 24), width=50).grid(row=0, column=0, rowspan=2, padx=(15, 10), pady=12)
        ctk.CTkLabel(self.container, text=app_data["Name"], font=("Segoe UI Semibold", 14), anchor="w").grid(row=0, column=1, sticky="sw", padx=5, pady=(12, 0))
        ctk.CTkLabel(self.container, text=app_data.get("Description", ""), font=("Segoe UI", 11), text_color=Theme.TEXT_SECONDARY, anchor="w").grid(row=1, column=1, sticky="nw", padx=5, pady=(0, 12))
        
        btn_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        btn_frame.grid(row=0, column=2, rowspan=2, padx=10)

        ctk.CTkButton(btn_frame, text="Notes", width=58, height=28, font=("Segoe UI", 11), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=lambda: self.on_release_notes(self.app_data)).pack(side="left", padx=3)
        self.chk = ctk.CTkCheckBox(btn_frame, text="", variable=self.selected, width=24, command=self._toggle, fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER)
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
        ctk.CTkButton(btn_frame, text="Notes", width=58, height=30, font=("Segoe UI", 11), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=lambda: on_release_notes(app_data)).pack(side="left", padx=3)
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
        
        self.user_profile = StoreAPI.load_user_profile()
        self.theme_mode_var = ctk.StringVar(value=Theme.normalize_mode(self.user_profile.get("ThemeMode", "System")))
        Theme.set_mode(self.theme_mode_var.get())
        ctk.set_appearance_mode(Theme.MODE)
        ctk.set_default_color_theme("dark-blue")
        self.configure(fg_color=Theme.BG_DARK)
        
        self.selected_apps = []
        self.download_queue = []
        self.current_packages = []
        self.selected_packages = set()
        self.package_rows = []
        self.current_view = "welcome"
        self.source_health = []
        self.arch_options = [f"Auto ({SYSTEM_ARCH})", "x64", "x86", "arm64", "arm", "neutral"]
        self.arch_override_var = ctk.StringVar(value=self.arch_options[0])
        self.package_scroll = None
        self.output_path = DEFAULT_OUTPUT
        self.shared_cache_enabled = ctk.BooleanVar(value=False)
        self.shared_cache_path = os.path.join(DEFAULT_OUTPUT, "SharedCache")
        
        self._build_ui()
        self._show_welcome()
        if os.environ.get("MSSTOREHELPER_SKIP_SOURCE_HEALTH") != "1":
            threading.Thread(target=self._source_health_worker, daemon=True).start()

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
        
        theme_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
        theme_frame.pack(side="right", padx=(5, 10))
        ctk.CTkLabel(theme_frame, text="Theme", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY).pack(side="left", padx=(0, 6))
        ctk.CTkOptionMenu(
            theme_frame,
            values=THEME_MODE_VALUES,
            variable=self.theme_mode_var,
            width=90,
            height=32,
            font=("Segoe UI", 12),
            fg_color=Theme.BG_INPUT,
            button_color=Theme.PRIMARY,
            button_hover_color=Theme.PRIMARY_HOVER,
            command=self._change_theme_mode,
        ).pack(side="left")

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

        self.search_history_frame = ctk.CTkFrame(search_section, fg_color="transparent", height=1)
        self._render_search_history()
        
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
        ctk.CTkButton(fix_section, text="🔎 Scan LTSC Gaps", height=34, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._scan_ltsc_gaps).pack(fill="x", pady=(0, 6))
        ctk.CTkButton(fix_section, text="🎮 Queue Xbox Core", height=34, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._queue_xbox_core).pack(fill="x")

        ctk.CTkFrame(self.sidebar, fg_color=Theme.BORDER, height=1).pack(fill="x", padx=15, pady=15)

        self.pinned_section = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        ctk.CTkLabel(self.pinned_section, text="⭐ Pinned Apps", font=("Segoe UI Semibold", 15), anchor="w").pack(fill="x")
        self.pinned_list_frame = ctk.CTkFrame(self.pinned_section, fg_color="transparent", height=1)
        self.pinned_list_frame.pack(fill="x", pady=(6, 0))
        self._render_pinned_favorites()
        
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
        ctk.CTkButton(btn_frame, text="Diagnostics", width=110, height=32, font=("Segoe UI", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._export_diagnostics_bundle).pack(side="right")

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
        self._log("INFO", f"Theme: {self.theme_mode_var.get()} ({Theme.MODE}) Accent: {Theme.PRIMARY}")
    
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
    
    def _show_welcome(self):
        self._clear_content()
        self.current_view = "welcome"
        
        center = ctk.CTkFrame(self.content, fg_color="transparent")
        center.place(relx=0.5, rely=0.5, anchor="center")
        
        card = ModernCard(center)
        card.pack(padx=20, pady=30)
        
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(padx=32, pady=32)
        
        ctk.CTkLabel(inner, text="👋", font=("Segoe UI Emoji", 42)).pack(pady=(0, 8))
        ctk.CTkLabel(inner, text="Welcome to MSStoreHelper", font=("Segoe UI Semibold", 22), wraplength=520).pack()
        ctk.CTkLabel(inner, text="Download and install Microsoft Store apps\nwithout needing access to the Store", font=("Segoe UI", 13), text_color=Theme.TEXT_SECONDARY, justify="center", wraplength=520).pack(pady=(10, 22))
        
        options = ctk.CTkFrame(inner, fg_color="transparent")
        options.pack()
        
        ctk.CTkButton(options, text="🔍 Search", width=160, height=42, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=lambda: self.search_entry.focus()).pack(side="left", padx=8)
        ctk.CTkButton(options, text="📂 Categories", width=160, height=42, font=("Segoe UI Semibold", 13), fg_color="transparent", border_width=2, border_color=Theme.PRIMARY, text_color=Theme.PRIMARY_OUTLINE_TEXT, hover_color=Theme.BG_CARD_HOVER, command=lambda: self._show_category("🛠️ Essential Repairs")).pack(side="left", padx=8)
        
        tips = ctk.CTkFrame(inner, fg_color=Theme.BG_INPUT, corner_radius=8)
        tips.pack(fill="x", pady=(30, 0))
        
        tip_inner = ctk.CTkFrame(tips, fg_color="transparent")
        tip_inner.pack(padx=18, pady=14)
        
        ctk.CTkLabel(tip_inner, text="💡 Tips", font=("Segoe UI Semibold", 13), anchor="w").pack(fill="x")
        ctk.CTkLabel(tip_inner, text="• Use 'Smart Select' to automatically pick the best files\n• Bundles (.msixbundle) work on all architectures\n• Run as Administrator for installation to work\n• Use 'Repair Store' if the Store shows connectivity errors", font=("Segoe UI", 11), text_color=Theme.TEXT_SECONDARY, justify="left", anchor="w", wraplength=500).pack(fill="x", pady=(5, 0))
    
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
        ctk.CTkButton(actions, text="Pin Selected", width=105, height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._pin_selected_apps).pack(side="left", padx=(0, 8))
        ctk.CTkButton(actions, text="Export WinGet", width=120, height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._export_winget_manifest).pack(side="left", padx=(0, 8))
        ctk.CTkButton(actions, text="Get Selected Apps", width=135, height=36, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._fetch_selected).pack(side="left")
        
        scroll = ctk.CTkScrollableFrame(self.content, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        for app in apps:
            AppTile(scroll, app, self._on_app_toggle, self._show_release_notes).pack(fill="x")
    
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
        ctk.CTkButton(actions, text="Pin Selected", width=105, height=36, font=("Segoe UI Semibold", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._pin_selected_apps).pack(side="left", padx=(0, 8))
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
            SearchResultTile(scroll, app, self._fetch_single_app, self._on_app_toggle, self._show_release_notes).pack(fill="x")
    
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

📝 Release Notes
Click "Notes" on any app row to fetch Store product page notes.

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

    def _show_release_notes(self, app_data):
        product_id = app_data.get("ProductId")
        if not product_id:
            self._update_status("⚠️ Missing product ID", Theme.WARNING)
            return

        self._update_status(f"📝 Fetching notes for {app_data.get('Name', 'app')}...", Theme.INFO)
        threading.Thread(target=self._release_notes_worker, args=(app_data,), daemon=True).start()

    def _release_notes_worker(self, app_data):
        try:
            notes = StoreAPI.fetch_release_notes(app_data["ProductId"])
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

        history = self.user_profile.get("SearchHistory", [])[:4]
        if not history:
            self.search_history_frame.pack_forget()
            return

        if not self.search_history_frame.winfo_manager():
            self.search_history_frame.pack(fill="x", pady=(8, 0))

        ctk.CTkLabel(self.search_history_frame, text="Recent", font=("Segoe UI", 10), text_color=Theme.TEXT_MUTED, anchor="w").pack(fill="x")
        for query in history:
            ctk.CTkButton(
                self.search_history_frame,
                text=query[:32],
                height=26,
                font=("Segoe UI", 11),
                fg_color="transparent",
                border_width=1,
                border_color=Theme.BORDER,
                hover_color=Theme.BG_CARD_HOVER,
                anchor="w",
                command=lambda q=query: self._search_from_history(q),
            ).pack(fill="x", pady=(4, 0))

    def _render_pinned_favorites(self):
        if not hasattr(self, "pinned_list_frame"):
            return

        for widget in self.pinned_list_frame.winfo_children():
            widget.destroy()

        favorites = self.user_profile.get("PinnedFavorites", [])[:5]
        if not favorites:
            self.pinned_section.pack_forget()
            return

        if not self.pinned_section.winfo_manager():
            self.pinned_section.pack(fill="x", padx=15, pady=(0, 10))

        for app in favorites:
            ctk.CTkButton(
                self.pinned_list_frame,
                text=f"{app.get('Icon', '📦')} {app['Name']}"[:34],
                height=30,
                font=("Segoe UI", 11),
                fg_color="transparent",
                border_width=1,
                border_color=Theme.BORDER,
                hover_color=Theme.BG_CARD_HOVER,
                anchor="w",
                command=lambda a=app: self._fetch_single_app(a),
            ).pack(fill="x", pady=(0, 5))

        ctk.CTkButton(
            self.pinned_list_frame,
            text="Clear Pins",
            height=26,
            font=("Segoe UI", 10),
            fg_color="transparent",
            border_width=1,
            border_color=Theme.BORDER,
            hover_color=Theme.BG_CARD_HOVER,
            command=self._clear_pinned_favorites,
        ).pack(fill="x", pady=(2, 0))

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
        for status in statuses:
            level = "SUCCESS" if status.get("Available") else "WARNING"
            self._post_ui(lambda s=status, lvl=level: self._log(lvl, source_status_summary(s)))

    def _log_source_diagnostic(self, diagnostic, item_name=None):
        label = item_name or diagnostic.get("Source", "Store source")
        for error in diagnostic.get("Errors", []):
            self.after(0, lambda e=error, s=diagnostic.get("Source", "Store source"): self._log("ERROR", f"{s}: {e}"))
        for fallback in diagnostic.get("Fallbacks", []):
            command = fallback.get("Command", "")
            detail = fallback.get("Detail", "")
            self.after(0, lambda f=fallback, c=command, d=detail, n=label: self._log("WARNING", f"{n} fallback via {f.get('Source')}: {c} ({d})"))

    def _get_packages_with_logging(self, app_data):
        diagnostic = StoreAPI.get_packages_with_diagnostics(app_data["ProductId"])
        self._log_source_diagnostic(diagnostic, app_data.get("Name"))
        return diagnostic["Packages"]
    
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

        queued_count = 0
        for app in missing_apps:
            self.after(0, lambda n=app["Name"]: self._update_status(f"📥 Queueing {n}...", Theme.INFO))
            packages = self._get_packages_with_logging(app)
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
        queued_count = 0
        for package in selected:
            if any(queued["FileName"] == package["FileName"] for queued in self.download_queue):
                continue
            self.download_queue.append(annotate_package(package.copy()))
            queued_count += 1

            if package.get("XboxCoreName"):
                if package.get("PinnedVersionMatched"):
                    self.after(0, lambda p=package: self._log("SUCCESS", f"Using pinned {p['XboxCoreName']} version: {p.get('AvailableVersion', 'unknown')}"))
                else:
                    pins = ", ".join(package.get("PinnedVersions", [])) or "configured pin"
                    self.after(0, lambda p=package, pins=pins: self._log("WARNING", f"Pinned {p['XboxCoreName']} version ({pins}) not available; queued {p.get('AvailableVersion', 'unknown')}"))

        self.download_queue = StoreAPI.order_packages_for_install(self.download_queue, target_arch)
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

    def _export_diagnostics_bundle(self):
        initial_dir = self.output_path if os.path.exists(self.output_path) else DEFAULT_OUTPUT
        os.makedirs(initial_dir, exist_ok=True)
        bundle_path = filedialog.asksaveasfilename(
            title="Save diagnostics bundle",
            initialdir=initial_dir,
            initialfile=f"MSStoreHelper-Diagnostics-{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip",
            defaultextension=".zip",
            filetypes=[("ZIP archive", "*.zip"), ("All files", "*.*")],
        )
        if not bundle_path:
            return

        try:
            StoreAPI.write_diagnostics_bundle(
                bundle_path,
                APP_VERSION,
                SYSTEM_ARCH,
                IS_ADMIN,
                self.output_path,
                self.source_health,
                self.download_queue,
                self._current_log_text(),
            )
        except Exception as exc:
            self._update_status("âŒ Diagnostics export failed", Theme.DANGER)
            self._log("ERROR", f"Failed to export diagnostics bundle: {exc}")
        else:
            self._update_status("âœ… Diagnostics exported", Theme.SUCCESS)
            self._log("SUCCESS", f"Diagnostics bundle saved: {bundle_path}")

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
            
            success, error_msg = StoreAPI.download_file(pkg['Url'], filepath, progress_cb, pkg)
            
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


def main():
    app = MSStoreHelperApp()
    app.mainloop()


if __name__ == "__main__":
    main()
