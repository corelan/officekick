# Changelog

## 0.1.2

- Permanent ribbon presence: the manifest now declares a dedicated "SlideKick" tab
  with a task-pane button (VersionOverrides / add-in commands) - no more opening
  Add-ins dropdown every time. The dropdown entry remains as fallback.
- The task pane follows the Office host theme (OfficeTheme.isDarkTheme, with
  body-color luminance fallback and an OfficeThemeChanged subscription); the
  simulator follows the browser (prefers-color-scheme) or `?theme=dark|light`.
- Dual-stack loopback listener (127.0.0.1 + ::1): macOS resolves 'localhost' to
  ::1 first and WKWebView does not fall back to IPv4 on a refusal, which made
  the pane fail with "Load failed" while everything else worked.
- The pane retries the initial config fetch (3x/1s) to survive opening during
  an installer-run server restart.

## 0.1.1

- Windows support: `install.ps1` (per-user trusted catalog via HKCU, scheduled task
  with Startup fallback, robocopy deploy) - same add-in code, server and config as
  macOS; documented in README.
- Localization-safe placeholder detection via `Shape.placeholderFormat`
  (PowerPointApi 1.8) with the shape-name heuristic as fallback - matters for
  Windows Office in non-English UI languages.
- License: no trademark claim over "SlideKick"; "Corelan" is a registered EU
  trademark of Corelan Consulting bv.
- Profiles tab now repopulates connection dropdowns on config load; connections
  referenced by profiles cannot be deleted (409 + UI hint); delete errors surfaced.
- Full edit semantics: renaming connections re-points profile references, renaming
  profiles moves the default star along; type/key/models editable in the same call;
  model lists in the profile editors are sorted alphabetically.

## 0.1.0 - initial release

- Office task-pane add-in for PowerPoint on macOS (Office.js), sideloaded via the
  classic XML manifest.
- Local companion server (Python 3 stdlib, 127.0.0.1 only): serves the UI, stores
  connections/profiles/approval policy, proxies AI calls.
- Provider types: openai, anthropic, openrouter, generic-openai (OpenAI-compatible),
  ollama, llama.cpp (llama-server).
- Connections (Name / Base URL / optional API Key) and profiles (content model +
  other-actions model, each picked as connection -> model list).
- Free-text instruction box (`reword the current slide`, `make 4 slides about X`, ...)
  with a structured slide-ops engine (create/edit slide, replace text, notes,
  apply layout, delete slide).
- Per-action approval policy: auto / always-confirm / deny, editable in Settings;
  delete-slide and apply-layout default to always-confirm.
- Browser simulator (`simulator.html`) for end-to-end testing without PowerPoint.
- Installer (`install.sh`): detects PowerPoint, deploys files + manifest
  (~ update/re-run supported), installs a LaunchAgent for the local server.