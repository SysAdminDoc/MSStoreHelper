# Roadmap

GUI tool to download/install Microsoft Store apps without the Store UI. Planned work focuses on dependency resolution, offline package management, and LTSC-specific workflows.

## Planned Features

### Package Resolution
- Automatic VCLibs/.NET Framework dependency resolution and chained install ordering
- Delta-update detection (skip if installed version >= available, report 0x80073D06 as no-op)
- Multi-architecture preference override (force x64 even when ARM64 neutral bundle offered)
- Package signature verification against Microsoft root before install

### Store Repair
- `wsreset` + token-cache rebuild + license re-sync preset for common "server stumbled" errors
- AppX provisioning repair for new user profiles (deprovisioned packages not reappearing)
- Licensing service reset (ClipSVC, LicenseManager) as a separate repair button
- Store cache corruption scan + offline rebuild

### Offline / Fleet
- Cache `.msixbundle`/`.appxbundle` artifacts to a shared folder for air-gapped reinstall
- Generate a DISM `/add-provisioned-appxpackage` script from the download queue
- Export selected apps as a WinGet import manifest for reproducible deployment
- Intune Win32 app packaging helper (emit `.intunewin` with detection script)

### LTSC Workflow
- One-click "LTSC Essentials" preset (Terminal, PowerShell 7, WSL, Photos, Calculator, Snipping Tool)
- Detect missing system components that LTSC ships without and queue them automatically
- Xbox Identity/Gaming Services install path with known-good version pinning (several xbox-app SKUs ship broken)

### UX
- Search history + pinned favorites per user profile
- Per-app release-notes fetch (Microsoft Store product page scrape)
- Dark/light theme toggle matching OS accent color
- Progress persistence across app restart (resumable downloads)

## Competitive Research
- **Winget** (Microsoft) — first-party, Store-backed source but opaque package discovery; use its manifest format for export interop.
- **Scoop / Chocolatey** — community repos, stronger dependency graphs; lesson: surface dependencies explicitly in UI.
- **RG-Adguard Store** — the underlying URL generator; steal its file-filter UX (arch/bundle toggles).
- **UUPDump** — sister-tool UX pattern: script-emit workflow for offline scenarios, worth mirroring for fleet installs.

## Nice-to-Haves
- WSL distro sideload from Microsoft's cloud store (Ubuntu/Debian `.appx` pulls)
- Scheduled "keep updated" mode: re-query all installed Store apps and pull newer bundles
- Rollback to previous `.appx` version (keep last 2 bundles in cache)
- Package diff: show what permissions/dependencies changed between two versions
- CLI twin for headless use via RMM (`MSStoreHelper.py --install Microsoft.WindowsTerminal`)
- Telemetry-free mirror mode: pre-download and serve Store packages from a local HTTP cache for multi-PC clinics

## Open-Source Research (Round 2)

### Related OSS Projects
- https://github.com/K3rhos/Microsoft-Store-Apps-EXE-Downloader — pulls raw EXE/MSIX from Store without Store installed, LTSC/tiny11 friendly
- https://github.com/hexadecimal233/Windows-Store-Downloader — uses store.rg-adguard.net API, paid/restricted/free app support
- https://github.com/kkkgo/LTSC-Add-MicrosoftStore — Store bootstrap for Win10 LTSC 2019
- https://github.com/R-YaTian/LTSC-Add-MicrosoftStore-2021_2024 — Store bootstrap for LTSC 2021/2024
- https://github.com/minihub/LTSC-Add-MicrosoftStore — Win11 24H2 LTSC minimal deps
- https://github.com/Goojoe/LTSC-ADD-Microsoft-Store — latest offline Store installer for LTSC
- https://github.com/ishad0w/microsoft-windows-10-ltsc-2021-microsoft-store — bundled NET.Native / VCLibs / UI.Xaml dep installer

### Features to Borrow
- store.rg-adguard.net API parsing for msixbundle download URLs (hexadecimal233) — avoids official Store auth flow entirely
- Paid/restricted app support via authenticated session token (hexadecimal233)
- Bundled dependency installer (VCLibs, NET.Native.Framework/Runtime, UI.Xaml) so MSIX installs actually complete on fresh LTSC (ishad0w)
- `wsreset -i` fallback to trigger Store self-install when AppxPackage add fails (megakarlach note)
- Minimal-component install option — strip optional Microsoft sub-packages to keep LTSC footprint small (minihub)
- Architecture auto-detect (x64/arm64/x86) before choosing msixbundle (K3rhos)
- Language-ring selector (RP/WIS/WIF/Retail) for Store API queries (K3rhos)

### Patterns & Architectures Worth Studying
- Store.rg-adguard.net API as the canonical Store download proxy — all serious OSS tools use this endpoint, document the dependency and add fallback handling
- MSIX dependency resolution order: runtime framework → UI.Xaml → app package (ishad0w installer sequence)
- Checksum verification of downloaded packages against Store-signed hashes (K3rhos) — prevents MITM on unofficial API path
