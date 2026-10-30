# Changelog

## 0.1.3
- Repository renamed to **officekick** (umbrella): `core/` holds the shared server,
  `slidekick/` the PowerPoint add-in, `wordkick/` the planned Word sibling (**DocKick**,
  same connections/profiles/approval model). Both installers now handle Word: the
  installer detects the Office apps, deploys per-add-in manifests (DocKick installs
  automatically once its manifest template lands in `wordkick/`), installs/updates the
  officekick LaunchAgent / scheduled task, and migrates the old SlideKick-era config
  (`~/.slidekick`), install dir and LaunchAgent label on the first run.

- New ops: `create_table` (PowerPointApi 1.8 addTable; new title slide when slide is
  omitted; position/size in points) and `move_shape` (left/top/width/height in points).
  Both default to `auto` in the approval policy (configurable in Settings like the rest).
- "Design this deck": AI design passes over the deck layout/alignment/balance; the
  deck outline now includes shape geometry (and slide size) so the model can compute
  concrete moves. Microsoft's built-in Designer/Design Ideas cannot be invoked by
  add-ins - the chip and 'design ...' instructions run SlideKick's own AI pass.
- Installer: port probing now ignores lingering CLOSE_WAIT sockets (lsof-listener
  first, then a SO_REUSEADDR bind probe matching the server's own bind); the drifted
  port desync between manifest and server is gone for good.
- Manifest version bumped (1.0.1.0) so Office re-registers the updated pane cleanly;
  the About tab reads its version from the server.
## 0.1.2

- Fixed: the "auto-apply" checkbox now actually works - all-`auto` runs apply
  without the proposals step; runs containing confirm/blocked/invalid actions
  still open the gated preview (with the new Select all button).
- Fixed: a dropped variable declaration in the ops validator crashed every deck-edit
  run (regression introduced with the speaker-notes switch).
- Speaker-notes generation is opt-in (off by default): Settings toggle, prompt rule for
  the model, and client-side skipping of notes actions while the setting is off.
- Conversation history: the pane sends the last exchanges with each request
  (session-only, capped) so follow-ups like "use the 2nd title" work; "New
  conversation" clears it.
- Installer: waits for the previous server's listener to exit before picking the port
  and kills a stubborn stale listener instead of drifting to the next port (which
  desynced the manifest URLs).

- Default transport: plain http://localhost (no certificates, no keychain changes);
  the add-in appears in Home -> Add-ins. `--https` opts in to the ribbon button:
  loopback-only local CA + localhost certificate, trusted in the login keychain
  (ribbon UI only registers with https URLs - verified against Office).
- Permanent ribbon presence via `--https`: OfficeTab(TabHome) button + logo
  (VersionOverrides in the taskpaneappversionoverrides namespace, XSD-validated).
- Dual-stack loopback listener (127.0.0.1 + ::1): macOS resolves 'localhost' to
  ::1 first and WKWebView does not fall back to IPv4 on a refusal, which made
  the pane fail with "Load failed" while everything else worked.
- The pane retries the initial config fetch (3x/1s) to survive opening during
  an installer-run server restart.
- Installer guarantees: stops the previous server before updating, restarts it
  via LaunchAgent/RunAtLoad, and health-checks; clears the Office add-in caches
  on every run (only while PowerPoint is closed); the task pane follows the
  Office host theme (isDarkTheme + luminance fallback + OfficeThemeChanged).

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