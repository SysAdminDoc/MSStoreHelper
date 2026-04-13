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
import re
import ctypes
import webbrowser
from datetime import datetime

# ==================== DEPENDENCY AUTO-INSTALL ====================
def install_requirements():
    required = {
        'customtkinter': 'customtkinter',
        'requests': 'requests',
        'bs4': 'beautifulsoup4',
        'packaging': 'packaging'
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
    from packaging import version
except ImportError:
    install_requirements()
    import customtkinter as ctk
    import requests
    from bs4 import BeautifulSoup
    from packaging import version

# ==================== CONFIGURATION ====================

APP_VERSION = "3.1.2"
APP_NAME = "MSStoreHelper"
API_URL = "https://store.rg-adguard.net/api/GetFiles"
STORE_SEARCH_URL = "https://storeedgefd.dsx.mp.microsoft.com/v9.0/manifestSearch"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

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

def extract_version(filename):
    match = re.search(r'[\._](\d+(?:\.\d+){1,4})[\._]', filename)
    if match:
        try:
            return version.parse(match.group(1))
        except:
            return version.parse("0")
    return version.parse("0")

def extract_package_name(filename):
    name = re.sub(r'[\._]\d+(?:\.\d+){1,4}[\._].*', '', filename)
    return name

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
                
                results.append({
                    "FileName": name, "Url": url, "Architecture": arch,
                    "FileType": ext, "IsBundle": is_bundle, "IsEncrypted": is_encrypted,
                    "SizeBytes": None, "SizeStr": "—"
                })
            
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
    def smart_select(packages, target_arch):
        """Intelligently select the best packages"""
        selected = []
        seen = {}
        
        sorted_pkgs = sorted(packages, key=lambda p: (
            p.get('IsBundle', False),
            not p.get('IsEncrypted', False),
            extract_version(p['FileName'])
        ), reverse=True)
        
        for pkg in sorted_pkgs:
            if pkg.get('IsEncrypted'):
                continue
            
            base = extract_package_name(pkg['FileName'])
            arch = pkg['Architecture']
            is_bundle = pkg.get('IsBundle', False)
            
            if is_bundle:
                if base not in seen:
                    seen[base] = pkg
                    selected.append(pkg)
            else:
                if arch not in [target_arch, 'neutral']:
                    continue
                key = f"{base}_{arch}"
                if key not in seen and base not in seen:
                    seen[key] = pkg
                    selected.append(pkg)
        
        return selected
    
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
    def install_package(filepath):
        try:
            cmd = f'Add-AppxPackage -Path "{filepath}" -ErrorAction Stop 2>&1'
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", cmd],
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
    def run_repair(log_callback=None, progress_callback=None):
        steps = [
            ("🔧 Resetting Internet settings...", 'RunDll32.exe InetCpl.cpl,ResetIEtoDefaults'),
            ("🔧 Starting Windows Update...", 'Start-Service -Name wuauserv -ErrorAction SilentlyContinue'),
            ("🔧 Starting BITS...", 'Start-Service -Name bits -ErrorAction SilentlyContinue'),
            ("🧹 Clearing Store cache...", 'Start-Process wsreset.exe -WindowStyle Hidden -Wait'),
            ("🧹 Clearing tokens...", r'Remove-Item "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache\*" -Recurse -Force -ErrorAction SilentlyContinue'),
            ("🔄 Re-registering Store...", r'Get-AppxPackage -AllUsers Microsoft.WindowsStore | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue }'),
            ("🌐 Resetting network...", 'netsh winsock reset 2>$null'),
            ("🌐 Flushing DNS...", 'ipconfig /flushdns 2>$null'),
        ]
        
        results = []
        for i, (desc, cmd) in enumerate(steps):
            if log_callback:
                log_callback(desc)
            try:
                if 'RunDll32' in cmd:
                    subprocess.run(cmd, shell=True, timeout=15)
                else:
                    subprocess.run(["powershell", "-NoProfile", "-Command", cmd], creationflags=subprocess.CREATE_NO_WINDOW, timeout=30)
                results.append((desc, True))
            except:
                results.append((desc, False))
            
            if progress_callback:
                progress_callback((i + 1) / len(steps))
        
        return results

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
    def __init__(self, master, pkg_data, on_toggle, index):
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
        arch_color = Theme.ARCH_MATCH if arch in [SYSTEM_ARCH, 'neutral'] else Theme.TEXT_MUTED
        ctk.CTkLabel(tags_frame, text=f" {arch} ", font=("Consolas", 10), text_color=arch_color).pack(side="left", padx=(0, 6))
        
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
        self.output_path = DEFAULT_OUTPUT
        
        self._build_ui()
        self._show_welcome()
    
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
        
        ctk.CTkButton(fix_section, text="⚡ Apply Quick Fix", height=38, font=("Segoe UI Semibold", 13), fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER, command=self._apply_quickfix).pack(fill="x")
        
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
        
        ctk.CTkButton(repair_frame, text="🔧 Repair Store", height=40, font=("Segoe UI Semibold", 13), fg_color=Theme.DANGER, hover_color=Theme.DANGER_HOVER, command=self._run_repair).pack(fill="x")
        ctk.CTkLabel(repair_frame, text="Fix Store connectivity issues", font=("Segoe UI", 10), text_color=Theme.TEXT_MUTED).pack(pady=(4, 0))
    
    def _build_queue_panel(self):
        header_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        header_frame.pack(fill="x", padx=15, pady=(20, 10))
        
        ctk.CTkLabel(header_frame, text="📥 Download Queue", font=("Segoe UI Semibold", 16)).pack(side="left")
        self.queue_count = ctk.CTkLabel(header_frame, text="0 items", font=("Segoe UI", 12), text_color=Theme.TEXT_MUTED)
        self.queue_count.pack(side="right")
        
        self.queue_scroll = ctk.CTkScrollableFrame(self.right_panel, fg_color=Theme.BG_INPUT, corner_radius=8)
        self.queue_scroll.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
        self.queue_empty = ctk.CTkLabel(self.queue_scroll, text="📭\n\nNo files in queue\n\nSearch for apps or browse\ncategories to get started", font=("Segoe UI", 12), text_color=Theme.TEXT_MUTED, justify="center")
        self.queue_empty.pack(expand=True, pady=40)
        
        btn_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        ctk.CTkButton(btn_frame, text="Clear", width=70, height=32, font=("Segoe UI", 12), fg_color="transparent", border_width=1, border_color=Theme.BORDER, hover_color=Theme.BG_CARD_HOVER, command=self._clear_queue).pack(side="left")
        
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
        
        # Log content area - starts visible
        self.log_content = ctk.CTkFrame(self.log_panel, fg_color=Theme.BG_DARK, height=180)
        self.log_content.pack(fill="x", side="top")
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
        self.log_expanded = True
        
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
        
        ctk.CTkLabel(header, text=category_name, font=("Segoe UI Semibold", 22)).pack(side="left")
        ctk.CTkButton(header, text="Get Selected Apps", height=36, font=("Segoe UI Semibold", 13), fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER, command=self._fetch_selected).pack(side="right")
        ctk.CTkLabel(header, text=cat_data.get("description", ""), font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY).pack(side="left", padx=20)
        
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
        
        ctk.CTkLabel(header, text=f'🔍 Results for "{query}"', font=("Segoe UI Semibold", 20)).pack(side="left")
        ctk.CTkLabel(header, text=f"{len(results)} apps found", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY).pack(side="left", padx=15)
        
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
        ctk.CTkButton(tb_inner, text="➕ Add to Queue", width=130, height=32, font=("Segoe UI Semibold", 12), fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER, command=self._add_to_queue).pack(side="right")
        
        col_header = ctk.CTkFrame(self.content, fg_color=Theme.BG_INPUT, corner_radius=6)
        col_header.pack(fill="x", padx=10, pady=(0, 5))
        
        ch_inner = ctk.CTkFrame(col_header, fg_color="transparent")
        ch_inner.pack(fill="x", padx=12, pady=8)
        ch_inner.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(ch_inner, text="", width=40).grid(row=0, column=0)
        ctk.CTkLabel(ch_inner, text="File Name", font=("Segoe UI Semibold", 11), anchor="w").grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(ch_inner, text="Size", font=("Segoe UI Semibold", 11), width=80).grid(row=0, column=2, padx=(0, 10))
        
        scroll = ctk.CTkScrollableFrame(self.content, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        for i, pkg in enumerate(packages):
            row = PackageRow(scroll, pkg, self._on_package_toggle, i)
            row.pack(fill="x", pady=1)
            self.package_rows.append(row)
        
        self._fetch_sizes_async()
    
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
        self._log("DEBUG", f"  Target architecture: {SYSTEM_ARCH}")
        
        best = StoreAPI.smart_select(self.current_packages, SYSTEM_ARCH)
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
                    self.download_queue.append(pkg.copy())
                    count += 1
        
        self._update_queue_ui()
        self._update_status(f"✅ Added {count} files to queue", Theme.SUCCESS)
        self._log("INFO", f"Added {count} files to download queue (total: {len(self.download_queue)})")
    
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
        if not to_install:
            self._update_status("⚠️ No downloaded files", Theme.WARNING)
            return
        
        threading.Thread(target=self._install_worker, args=(to_install,), daemon=True).start()
    
    def _install_worker(self, packages):
        self.after(0, lambda: self._log("INFO", f"Starting installation of {len(packages)} packages"))
        self.after(0, lambda: self._log("INFO", "Note: Install order matters - dependencies should be installed first"))
        
        success_count = 0
        total = len(packages)
        
        for i, pkg in enumerate(packages):
            fname = pkg['FileName']
            filepath = pkg['LocalPath']
            
            self.after(0, lambda n=fname: self._update_status(f"📦 Installing {n[:40]}...", Theme.INFO))
            self.after(0, lambda n=fname, idx=i+1, tot=total: self._log("INFO", f"[{idx}/{tot}] Installing: {n}"))
            self.after(0, lambda p=filepath: self._log("DEBUG", f"  Path: {p}"))
            
            if '_status_widget' in pkg:
                self.after(0, lambda w=pkg['_status_widget']: w.configure(text="Installing...", text_color=Theme.INFO))
            
            success, error_msg = StoreAPI.install_package(filepath)
            
            if '_status_widget' in pkg:
                if success:
                    self.after(0, lambda w=pkg['_status_widget']: w.configure(text="✅ Installed", text_color=Theme.SUCCESS))
                    self.after(0, lambda n=fname: self._log("SUCCESS", f"  Successfully installed: {n}"))
                    success_count += 1
                else:
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
        
        if success_count == total:
            self.after(0, lambda: self._log("SUCCESS", f"Installation complete: All {total} packages installed successfully"))
        else:
            self.after(0, lambda: self._log("WARNING", f"Installation complete: {success_count}/{total} packages installed"))
            self.after(0, lambda: self._log("INFO", "Tip: Check the errors above. Common fixes:"))
            self.after(0, lambda: self._log("INFO", "  1. Install dependencies (VCLibs, .NET) before main apps"))
            self.after(0, lambda: self._log("INFO", "  2. Enable Developer Mode in Windows Settings"))
            self.after(0, lambda: self._log("INFO", "  3. Try a different package version (older/newer)"))
    
    def _run_repair(self):
        if not IS_ADMIN:
            self._update_status("⚠️ Administrator required", Theme.WARNING)
            return
        
        dialog = ctk.CTkToplevel(self)
        dialog.title("Repair Store")
        dialog.geometry("400x200")
        dialog.transient(self)
        dialog.grab_set()
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 400) // 2
        y = self.winfo_y() + (self.winfo_height() - 200) // 2
        dialog.geometry(f"+{x}+{y}")
        
        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=25, pady=25)
        
        ctk.CTkLabel(content, text="🔧 Repair Microsoft Store?", font=("Segoe UI Semibold", 18)).pack()
        ctk.CTkLabel(content, text="This will reset Store settings, clear cache,\nand fix common connectivity issues.\n\nA reboot may be required.", font=("Segoe UI", 12), text_color=Theme.TEXT_SECONDARY, justify="center").pack(pady=15)
        
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack()
        
        def do_repair():
            dialog.destroy()
            self._update_status("🔧 Repairing Store...", Theme.INFO)
            threading.Thread(target=self._repair_worker, daemon=True).start()
        
        ctk.CTkButton(btn_frame, text="Cancel", width=100, fg_color="transparent", border_width=1, border_color=Theme.BORDER, command=dialog.destroy).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="🔧 Repair", width=100, fg_color=Theme.DANGER, hover_color=Theme.DANGER_HOVER, command=do_repair).pack(side="left", padx=10)
    
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


if __name__ == "__main__":
    app = MSStoreHelperApp()
    app.mainloop()
