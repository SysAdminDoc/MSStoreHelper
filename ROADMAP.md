# Roadmap

GUI tool to download/install Microsoft Store apps without the Store UI. Planned work focuses on dependency resolution, offline package management, and LTSC-specific workflows.

## Planned Features

### UX
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

## Research-Driven Additions

- [ ] P2 - Accessibility, DPI, and keyboardless GUI audit
  Why: Dense CustomTkinter rows use fixed widths/wrap lengths and emoji-heavy labels that can clip or degrade on high DPI/theme changes.
  Evidence: `MSStoreHelper.py:1356`, `MSStoreHelper.py:1415`, `MSStoreHelper.py:1534`, `tests/test_theme.py`, README screenshot.
  Touches: UI components, theme tokens, screenshot capture, GUI smoke checks.
  Acceptance: desktop screenshot pass verifies no clipped action buttons or unreadable contrast at 100/125/150 percent scaling in dark and light modes; controls remain usable without documented keyboard shortcuts.
  Complexity: M

- [ ] P2 - Locale, market, and ring controls
  Why: Package lookup hard-codes `Retail` and `en-US`, limiting non-US admins and Store flight troubleshooting.
  Evidence: `MSStoreHelper.py:561`, K3rhos ring/language selector, Microsoft Store regioned product pages.
  Touches: package fetch UI, user profile settings, Store source adapter, tests for query payloads.
  Acceptance: user can select ring, language, and market; selections persist per profile; query metadata appears in logs and exported deployment artifacts.
  Complexity: M

- [ ] P3 - App Installer manifest export
  Why: Microsoft supports `.appinstaller` files with update settings, offering a lightweight distribution path between one-off installs and full Intune packaging.
  Evidence: Microsoft App Installer docs, `Add-AppxPackage -AppInstallerFile`, existing DISM/WinGet/Intune export flows.
  Touches: export actions, package metadata, README, tests for generated XML.
  Acceptance: queued packages can export a valid `.appinstaller` plus package folder, and Windows can install it with `Add-AppxPackage -AppInstallerFile` on a test machine.
  Complexity: M
