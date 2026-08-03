# Roadmap

Only incomplete, research-backed work belongs here.

## Research-Driven Additions

### P0 — Now

### P1 — Next

- [ ] P1 — Build inspected install plans and correct deployment exports
  Why: Separate installs can leave partial dependency state, multi-app queues produce invalid App Installer semantics, bundle diffs miss inner manifests, and rollback ignores architecture.
  Evidence: `MSStoreHelper.py::read_appx_manifest_details`, `generate_appinstaller_manifest`, `install_package`, `_install_worker`, `rollback_candidates`, `diff_appx_manifests`; Microsoft `Add-AppxPackage`, MSIX update, App Installer, and DISM documentation.
  Touches: AppX/bundle inspector, install planner, command runner, AppInstaller/DISM/Intune exporters, rollback/diff, tests.
  Acceptance: Outer and inner manifests produce a per-main-app plan with exact identity, publisher, architecture, min OS, dependencies, optional/resource packages, capabilities, and installed-state conflicts; ambiguous multi-app queues are split or rejected; installs use documented dependency/optional-package batching where supported; inventory failures never imply absence; rollback/diff are architecture-safe; and dry-run output plus generated artifacts round-trip through schema/XML/PowerShell validation.
  Complexity: L

- [ ] P1 — Publish the support and trust contract
  Why: The 2026-07-29 docs overstate Python compatibility, under-explain RG-Adguard/Store API risk, advertise LAN binding without a threat warning, and do not define repair recovery or supported Windows builds.
  Evidence: `README.md`, `CHANGELOG.md`, `CLAUDE.md`, `pyproject.toml`; Microsoft Store policy/troubleshooting and Python lifecycle documentation.
  Touches: README, CHANGELOG, CLAUDE, CLI/help copy, release checklist.
  Acceptance: Documentation states the tested Windows/LTSC/Python matrix, user-vs-machine and elevation semantics, source/fallback support status, trust/revocation modes, offline wheelhouse matrix, mirror exposure model, destructive repair plan/backup/restore flow, state/cache migrations, CLI schema/exit codes, and unsigned artifact verification; the malformed v3.2.0 changelog heading and stale repository counts are corrected; no local SHA is described as a Store-signed hash.
  Complexity: S

- [ ] P1 — Version the CLI contract and persist a local operation journal
  Why: RMM consumers need stable schemas and post-run evidence, while v3.35.0 JSON has no schema version and GUI logs disappear with the process.
  Evidence: `MSStoreHelper.py::build_cli_parser`, `_cli_emit_summary`, `_cli_package_workflow`, `_log`; WinGet JSON/configuration patterns; PDQ and Ninite audit/history features.
  Touches: typed result model, CLI JSON/exit codes, bounded journal, diagnostics, mirror/repair/download/install events, contract tests.
  Acceptance: JSON includes `SchemaVersion`, operation ID, absolute timestamps, source/query, trust outcome, per-package state, warnings/errors, reboot requirement, and recovery artifact paths; exit codes distinguish validation, privilege, partial, trust, network, and internal failure; a bounded local JSONL journal is atomic, redacted, exportable, and telemetry-free; backward-compatibility fixtures lock the v1 contract.
  Complexity: M

- [ ] P1 — Make every major UI surface responsive and accessible
  Why: 2026-07-29 live dark/light/system rendering found clipped or invisible actions, sub-4.5:1 small text, unnamed selection controls, and little keyboard/focus support.
  Evidence: `MSStoreHelper.py::Theme`, `AppTile`, `SearchResultTile`, `PackageRow`, `_build_ui`, `_build_sidebar`, `_show_packages`; `tests/test_accessibility.py`; WCAG 2.2 and Microsoft accessibility guidance.
  Touches: theme tokens, shared controls/layouts, welcome/search/packages/queue/log/help/dialogs, accessibility/render tests.
  Acceptance: Deterministic screenshots pass at 100%, 125%, 150%, and 200% in System/Dark/Light and Windows High Contrast at 1000×600 and 1280×800; normal text is at least 4.5:1; content reflows or scrolls without clipped controls; every control has a visible/automation name, role, state, focus cue, and logical tab order; Enter/Space/Escape and screen-reader workflows work; loading/conflicting actions are disabled; source errors differ from empty states.
  Complexity: L

- [ ] P1 — Add Windows CI and reproducible unsigned release artifacts
  Why: There is no tracked CI, clean-wheel/entry-point smoke, real Windows integration lane, SBOM, or packaged GUI verification.
  Evidence: repository tree and 2026-07-29 test run; PyPA secure installs; CustomTkinter packaging guidance.
  Touches: `.github/workflows/`, test/tool configuration, build scripts, package metadata, release checklist, README.
  Acceptance: Windows CI installs locked dependencies on Python 3.11–3.14, runs unit/integration, lint, focused type, dependency-audit, malicious-input, and wheel tests; a clean environment installs the wheel and launches both `msstorehelper --search` and the GUI; a scheduled/manual Windows lane exercises PowerShell/AppX/DISM repair dry-runs and restore in an isolated test account/VM; release builds produce an unsigned self-contained artifact, wheel, SHA-256 manifest, and SBOM from a clean tag, then foreground-smoke the artifact and close it.
  Complexity: L

### P2 — Later

- [ ] P2 — Extract backend services from the GUI monolith
  Why: The 5,303-line module couples trust-critical backend behavior to CustomTkinter state, making concurrency, CLI, testing, and migration changes harder to isolate.
  Evidence: `MSStoreHelper.py`, existing seams in `msstore_package_resolution.py` and `store_sources.py`.
  Touches: new store client, downloader/cache, state, trust/inspection, command runner, repair, exporter, operation, and GUI view-model modules; imports/tests/package metadata.
  Acceptance: CLI startup does not import CustomTkinter; backend services accept typed immutable inputs and have no Tk globals; GUI binds only through a view model/operation coordinator; PowerShell construction is centralized; persistence is behind repositories/migrations; existing public CLI behavior and all tests remain green after each incremental extraction.
  Complexity: L

- [ ] P2 — Consume verified offline repository indexes
  Why: MSStoreHelper can write shared caches and mirror indexes but cannot browse, import, or install from them as a first-class offline source.
  Evidence: `MSStoreHelper.py::build_mirror_index`, `cache_downloaded_artifact`, `_rollback_cache_folders`; Raven, StoreLib, PDQ local repository/cache patterns.
  Touches: source adapter interface, mirror/index client, trust gate, cache browser, GUI/CLI import, queue URL refresh, tests.
  Acceptance: A versioned index can be opened from disk or an authenticated mirror; schema, size, hash, signature/identity, architecture, and query metadata are verified before queueing; packages group by main app/dependency; missing or stale entries can be refreshed only when online; a fully disconnected machine can browse, plan, install, and export from verified cached artifacts without contacting Store services.
  Complexity: L

- [ ] P2 — Add official Store and WinGet execution backends
  Why: Microsoft’s 2026-02-11 Store release exposes `store install/update`, while WinGet already provides `msstore` policy/error semantics; both are useful when present but cannot be assumed.
  Evidence: `store_sources.py` fallback hints, `MSStoreHelper.py::detect_source_health`; Microsoft Store developer-tools announcement and WinGet documentation.
  Touches: source/execution adapters, capability discovery, CLI/GUI source selection, typed results, diagnostics, tests.
  Acceptance: Capability detection distinguishes missing, policy-blocked, unauthenticated, and usable Store/WinGet backends; users can explicitly choose or order backends; commands use exact product IDs/arguments and return typed results; no backend silently changes ring/market/architecture intent; direct/offline workflows remain available and no Store/App Installer dependency becomes mandatory.
  Complexity: M

- [ ] P2 — Add package and queue workflow controls
  Why: Competitors expose filtering, grouping, per-item retry/remove/cancel, version pins, and operation history; the v3.35.0 package list and queue are all-or-nothing.
  Evidence: `MSStoreHelper.py::_show_packages`, `QueueItem`, `_update_queue_ui`; Raven and UniGetUI; WinGet pinning.
  Touches: package/queue view models, responsive UI, persisted state schema, operation coordinator, CLI.
  Acceptance: Packages filter/sort/group by app, dependency role, architecture, type, encryption, version, and size; queue items support remove, retry, cancel, and safe app-group reorder; users can choose output/cache folders in GUI; estimated bytes/free space and pin/skip reasons are visible; all controls persist through the versioned state model and remain keyboard accessible.
  Complexity: M

- [ ] P2 — Externalize strings and preserve Unicode end to end
  Why: Locale/market controls exist, but UI/log/help strings are embedded in code and release-note decoding corrupts non-ASCII text.
  Evidence: `MSStoreHelper.py::parse_release_notes_html` and embedded UI strings; Raven and UniGetUI localization; Microsoft accessibility/globalization guidance.
  Touches: resource loader/catalog, GUI/CLI/help/log copy, locale selection, release-note parser, layout/render tests.
  Acceptance: User-facing strings use locale resources with stable keys and `en-US` fallback; no `unicode_escape` round-trip is used for already-decoded text; Japanese and accented regression fixtures preserve exact Unicode; pseudo-localization catches clipping; locale/market are separate concepts; untranslated keys and formatting mismatches fail tests.
  Complexity: L

- [ ] P2 — Add deterministic local deployment policy profiles
  Why: Commercial tools derive value from approvals, pins, rings, maintenance windows, detection, and retry evidence without requiring every local utility to become a fleet server.
  Evidence: Patch My PC rings/alerts, PDQ retry/deployment history, Ninite update policies, Intune detection/supersedence, WinGet pin/configuration.
  Touches: versioned profile schema, planner, keep-updated mode, CLI, WinGet/Intune/AppInstaller exports, journal, docs/tests.
  Acceptance: Importable/exportable profiles declare approved product IDs, architecture/market/ring, version pin/ignore rules, maintenance window, retry/reboot policy, trust mode, and dry-run behavior; Keep Updated and CLI consume the same plan; Intune exports include deterministic detection and supersedence metadata; policy changes produce a human-readable diff and never enable paid/restricted acquisition.
  Complexity: L

### P3 — Under Consideration

- [ ] P3 — Add AppxBlockMap delta downloads behind a capability flag
  Why: Raven demonstrates meaningful bandwidth savings and Microsoft’s MSIX update format uses 64 KiB block maps, but CDN/range behavior and safe reconstruction add substantial complexity.
  Evidence: Raven; Microsoft MSIX package-update documentation; `MSStoreHelper.py::download_file` and cache history.
  Touches: AppxBlockMap parser, downloader/range planner, cache, trust gate, progress/journal, fallback tests.
  Acceptance: For a verified prior version and compatible source, the planner validates block-map hashes, downloads only changed ranges, reconstructs to a new temporary artifact, then passes the full normal size/hash/signature/identity gate before atomic promotion; any unsupported encoding, validator change, missing block, or hash mismatch discards reconstruction and falls back to a full download without modifying the prior cache.
  Complexity: XL
