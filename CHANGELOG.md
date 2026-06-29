# Changelog

All notable changes to MSStoreHelper will be documented in this file.

## [v3.28.0] - 2026-06-29

- Added App Installer `.appinstaller` export for downloaded AppX/MSIX queues with copied package folders and OnLaunch update settings.
- Changed the queue panel to expose an AppInstaller export action alongside DISM and IntuneWin exports while preserving the default layout.
- Added tests for AppX identity parsing, generated App Installer XML, package copying, and missing-download errors.

## [v3.27.0] - 2026-06-28

- Added persisted Store query controls for RG-Adguard ring, Store language, and market.
- Changed package lookup, release-note fetches, logs, diagnostics, DISM scripts, and IntuneWin staging output to include Store query metadata.
- Added tests for query payloads, profile persistence, localized release-note URLs, diagnostics metadata, and deployment artifact metadata.

## [v3.26.0] - 2026-06-28

- Added theme contrast helpers and tests for core text, muted text, and outline action states.
- Changed the welcome screen layout to avoid clipped primary actions at 100, 125, and 150 percent scaling in dark and light themes.
- Added a deterministic source-health skip hook for GUI smoke captures.

## [v3.25.0] - 2026-06-28

- Added a mocked integration harness for StoreEdgeFD search, RG-Adguard package parsing, failed downloads, signature failures, Appx install failures, and IntuneWin command failures.
- Changed test coverage to exercise Store and PowerShell workflow failures without network, admin rights, or external tools.

## [v3.24.0] - 2026-06-28

- Added a redacted diagnostics ZIP export with app version, Windows/Python details, source health, queue metadata, app logs, PowerShell transcript lines, and recent repair manifests.
- Changed the queue panel to expose a Diagnostics action for support bundle generation.
- Added tests for diagnostics bundle contents and redaction behavior.

## [v3.23.0] - 2026-06-28

- Added pinned `requirements.txt` and `pyproject.toml` package metadata with a `msstorehelper` entry point.
- Changed dependency bootstrap to fail fast with online and offline wheelhouse install commands instead of running `pip install` at app startup.
- Added tests for dependency detection and setup-message content.

## [v3.22.0] - 2026-06-28

- Added a Store source adapter helper for HTTP retry diagnostics, source health checks, and WinGet/Store CLI fallback hints.
- Changed Store search, package lookup, and release-note fetches to use source-specific retry/error reporting.
- Added tests for source health detection, retry recovery, fallback command generation, and RG-Adguard package parsing.

## [v3.21.0] - 2026-06-28

- Added timestamped repair backup manifests and restore scripts for Store repair, provisioning repair, licensing reset, and cache rebuild actions.
- Changed repair PowerShell execution to capture structured stdout, stderr, exit code, command text, and backup paths per step.
- Added tests for fail-visible repair output and rollback artifact generation.

## [v3.20.0] - 2026-06-28

- Added atomic `.part` downloads with final-file replacement only after the full response is received.
- Added SHA-256 artifact metadata manifests for downloaded and shared-cache packages.
- Added System/Dark/Light theme mode with Windows accent-color detection and profile persistence.
- Changed shared cache reuse to validate hash and size instead of trusting same-size files.

## [v3.19.0] - 2026-06-27

- Added per-app Microsoft Store page release-note fetching from category and search result rows.
- Changed app rows to expose a Notes action with an in-app release-notes dialog.
- Added tests for Store page heading, embedded JSON, and product-description fallback parsing.

## [v3.18.0] - 2026-06-27

- Added per-user search history and pinned favorite apps stored in the user's AppData profile.
- Changed category and search result headers to allow pinning selected apps.
- Added sidebar rendering for recent searches and pinned favorites.
- Added tests for profile JSON persistence, history de-duplication, and favorite de-duplication.

## [v3.17.0] - 2026-06-27

- Added a dedicated Xbox Core queue path for Xbox Identity Provider and Gaming Services.
- Changed Xbox core selection to prefer known-good pinned versions when available and log fallback versions when pins are unavailable.
- Added tests for Xbox core version pins, dependency-first ordering, and fallback selection.

## [v3.16.0] - 2026-06-27

- Added LTSC missing-component detection for tracked Store, runtime, media, productivity, and shell packages.
- Changed Quick Actions to include a Scan LTSC Gaps action that queues Smart Select packages for missing components automatically.
- Added tests for requirement catalog coverage and installed-identity filtering.

## [v3.15.0] - 2026-06-27

- Added a one-click LTSC Essentials quick action for Terminal, PowerShell 7, WSL, Photos, Calculator, and Snipping Tool.
- Changed Quick Actions to make the LTSC-focused preset the default selection.
- Added tests that pin the LTSC Essentials preset contents against the app catalog.

## [v3.14.0] - 2026-06-27

- Added IntuneWin queue export using Microsoft's IntuneWinAppUtil when available or selected.
- Changed queue actions to generate Intune install and detection scripts for downloaded AppX/MSIX packages.
- Added tests for Intune source staging, detection script generation, and content prep command arguments.

## [v3.13.0] - 2026-06-27

- Added WinGet `msstore` import manifest export for selected Store apps.
- Changed category and search result headers to expose a selected-app WinGet export action.
- Added tests for manifest source metadata, package de-duplication, JSON write output, and invalid selections.

## [v3.12.0] - 2026-06-27

- Added DISM provisioning script export for queued AppX/MSIX packages.
- Changed the queue actions to expose a fleet-ready PowerShell script generator.
- Added tests for dependency-first DISM output, local path handling, and non-installable queue rejection.

## [v3.11.0] - 2026-06-27

- Added an optional shared offline cache folder for downloaded AppX/MSIX artifacts.
- Changed successful downloads to mirror cacheable artifacts into the selected shared folder.
- Added tests for cache copy, same-size reuse, and non-installable file rejection.

## [v3.10.0] - 2026-06-27

- Added a Store cache scan and offline rebuild action that backs up cache folders before recreating them.
- Changed repair UI to expose Rebuild Cache as its own action.
- Added tests for Store cache scan, backup, rebuild, and `wsreset` steps.

## [v3.9.0] - 2026-06-27

- Added a dedicated Store licensing reset action for ClipSVC and LicenseManager.
- Changed repair UI to expose Reset Licensing as its own button.
- Added tests for licensing service reset and ClipSVC cache cleanup steps.

## [v3.8.0] - 2026-06-27

- Added a Store provisioning repair action for deprovisioned Store packages and new-profile registration.
- Changed repair UI to expose a dedicated Provision Store button.
- Added tests for AppX deprovision tombstone cleanup and Store package re-registration.

## [v3.7.0] - 2026-06-27

- Added a fuller Store repair preset for `wsreset`, Store token-cache rebuild, and licensing service re-sync.
- Changed Store repair to start immediately with console/status feedback instead of a confirmation dialog.
- Added tests that pin the required repair preset steps.

## [v3.6.0] - 2026-06-27

- Added Microsoft-chain package signature verification before AppX/MSIX installation.
- Changed install flow to block and log packages that fail signature verification.
- Added resolver tests for Microsoft signature metadata validation.

## [v3.5.0] - 2026-06-27

- Added a package-list architecture selector with Auto, x64, x86, ARM64, ARM, and neutral targets.
- Changed Smart Select to prefer exact-architecture packages when an explicit override is selected.
- Added resolver coverage for explicit-architecture package preference.

## [v3.4.0] - 2026-06-27

- Added installed-version detection before AppX installs so current or newer packages are skipped.
- Changed `0x80073D06` higher-version install failures to successful no-op results.
- Added resolver tests for installed-version comparisons and unknown-version handling.

## [v3.3.0] - 2026-06-27

- Added automatic dependency framework selection for VCLibs, .NET Native, UI.Xaml, and Windows App Runtime packages.
- Changed Smart Select and installation to order dependency packages before app packages.
- Added resolver tests for package selection, dependency filtering, and install ordering.

## [v3.2.0] - %Y->- (HEAD -> main, origin/main, origin/HEAD)

- Added: Add screenshot to README
- Changed: Update README.md
- Changed: Update README.md
- Changed: Update README.md
- Added: Add files via upload
- Added: Add files via upload
