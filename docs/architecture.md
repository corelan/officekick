# OfficeKick architecture

Two task-pane add-ins on one technology stack: **Corelan SlideKick** (PowerPoint)
and **Corelan DocKick** (Word), one local companion server, one shared config
(connections, profiles, per-action approval policy). Repo:

```
officekick/
  core/server.py           local companion server (Python stdlib only, loopback)
  slidekick/               PowerPoint add-in: manifest template + task pane
  wordkick/                Word add-in (DocKick): manifest template + task pane
  tools/                   mock AI provider + icon generator
  tests/                   server test suite
  install.sh / install.ps1 macOS / Windows installers (both add-ins)
```

```
        Microsoft PowerPoint for Mac / Windows          Microsoft Word for Mac / Windows
+-------------------------------------------+    +-------------------------------------------+
| SlideKick task pane (Office.js, WebView)  |    | DocKick task pane (Office.js, WebView)    |
|   shared pane UI (slate/theme/history)    |    |   same shared pane UI                     |
|   PowerPoint ops engine + approval gate   |    |   Word ops engine + approval gate         |
|   OfficeDeck -> PowerPoint.run(...)       |    |   WordDeck  -> Word.run(...)              |
+---------------------|---------------------+    +---------------------|---------------------+
                      |     same origin http://127.0.0.1:8765 (or https with --https)
                      +----------------------|---------------------+
                                             v
+--------------------------------------------------------------------------------------+
|  core/server.py (python stdlib, threads, dual-stack loopback)                        |
|    static files per add-in: /slidekick/static/...  /wordkick/static/...  (no-store)  |
|    /api/config  /api/connections  /api/profiles   /api/settings (approvals|debug)    |
|    /api/models  /api/chat  /api/test-connection     /api/debug (daily append log)    |
|    Store: ~/.officekick/config.json (0600) - SHARED by both add-ins                  |
|    provider adapters: openai | anthropic(messages) | openrouter |                    |
|                       generic-openai | ollama(chat) | llamacpp(v1/chat/completions)  |
|    reasoning-strip + hidden clean-JSON retry; daily debug log when enabled           |
+------------------------------------------|-------------------------------------------+
                                           | https/http
                                           v
                        OpenAI / Anthropic / OpenRouter /
                        OpenAI-compatible / Ollama / llama-server
```

## The ops protocol (one JSON schema per host)

Both panes run the same pipeline: free-text instruction -> profile/route ->
`/api/chat` -> reasoning strip (hidden repair round for unclear output) -> JSON ops
extraction -> validation against document context -> **approval gate** ->
Office.js apply. Models may answer conversationally (`{"kind":"reply", ...}`) or
emit an ops list (`slide_ops` / `doc_ops`) - the panes share the gate, the
history (last 6 exchanges, session-only), auto-apply, and the Settings model.

- SlideKick ops: `create_slide`, `edit_slide`, `replace_text`, `set_notes`,
  `create_table`, `move_shape`, `apply_layout`, `delete_slide`.
- DocKick ops: `insert_paragraphs` (with built-in styles), `edit_paragraph`,
  `replace_text`, `create_table`, `insert_comment`, `delete_paragraph`,
  `rewrite_doc`.

Approval policy: per-op `auto` / `confirm` / `deny` in the shared config; the
destructive-heavy ops default to `confirm` (deletes, layout moves, full rewrites),
additive ops default to `auto`. The gate is enforced in the pane; the server
never touches Office documents.

## Key design points

- **One server, many add-ins.** `core/server.py` serves each add-in's static
  tree (`/slidekick/static/...`, `/wordkick/static/...`) plus one API surface.
  Connections/profiles are configured once and usable from both apps.
- **One Deck interface per host.** `applyOps` + validation + index-shifting
  live in each pane; each Deck implementation (`OfficeDeck`/`WordDeck`/the
  simulator) exposes the same methods, so the browser simulator is an honest
  end-to-end test of the whole pipeline.
- **Append-only creation.** Office.js slides always append; Word inserts go to
  the document end in v1 - both report the actual result in the applied-ops log.
- **Capability probes.** Requirement sets (`PowerPointApi`/`WordApi`) are
  probed and missing APIs degrade to clear messages or manual-paste fallbacks
  (e.g. comments below WordApi 1.4).
- **Approval policy lives server-side, enforced client-side** (the server never
  touches the document).
- **Installers** deploy the umbrella layout (`~/Library/Application Support/OfficeKick`
  / `%LOCALAPPDATA%\OfficeKick`), render per-add-in manifests from the templates
  (Word's manifest ships as soon as `wordkick/manifest.template.xml` exists),
  sideload into each app's WEF store, register the auto-start agent, and
  health-check the server; legacy SlideKick-era installs are migrated on the run.
