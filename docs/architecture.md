# OfficeKick architecture

Three task-pane add-ins on one technology stack: **Corelan SlideKick**
(PowerPoint), **Corelan DocKick** (Word) and **Corelan MailKick** (Outlook),
one local companion server, one shared config (connections, profiles,
per-action approval policy). Repo layout:

```
officekick/
  core/server.py           local companion server (Python stdlib only, loopback)
  slidekick/               PowerPoint add-in: manifest template + task pane
  wordkick/                Word add-in (DocKick): manifest template + task pane
  mailkick/                Outlook add-in (MailKick): manifest template + task pane
  docs/                    per-add-in docs + this architecture page
  tools/                   mock AI provider + icon generator
  tests/                   server test suite (kept LOCAL, not in the repo)
  VERSION                  current version (the update check reads this from GitHub)
  install.sh / install.ps1 macOS / Windows installers (all three add-ins)
```

```
     PowerPoint task pane               Word task pane              Outlook task pane
     (Office.js WebView)                (Office.js WebView)         (Office.js WebView)
     SlideKick ops engine               DocKick ops engine          MailKick ops engine
     approval gate + history            same shared gate            same shared gate
     OfficeDeck -> PowerPoint.run       WordDeck -> Word.run        MailDeck -> mailbox API
              |                                  |                            |
              +--------------------- same origin https://127.0.0.1:8765 -----------+
                                             |   (loopback; --https opt-in on macOS)
                                             v
+---------------------------------------------------------------------------------------+
|  core/server.py (python stdlib, threads, dual-stack loopback)                          |
|    static files per add-in, no-store: /slidekick/static/  /wordkick/static/           |
|                                       /mailkick/static/  (legacy /outlook/static alias) |
|    /api/health   /api/update-check (reads VERSION from GitHub; ?latest= test override) |
|    /api/config   /api/settings (approvals, autoApplyApproved, debug, autoUpdateCheck)  |
|    /api/connections(+rename)  /api/profiles(+rename)  /api/test-connection             |
|    /api/models   /api/image   /api/debug (daily append log)                            |
|    /api/sessions(+/id, +/rename/<id>) - per-app sessions, AI-named, resumable          |
|    /api/templates(+/id) - MailKick named templates (mailkick only)                     |
|    Store: ~/.officekick/config.json (0600), ~/.officekick/sessions/*.json              |
|    provider adapters: openai | anthropic(messages) | openrouter |                      |
|                       generic-openai | ollama(chat) | llamacpp(v1/chat/completions)    |
|    reasoning-strip + hidden clean-JSON repair round; daily debug log when enabled      |
+-------------------------------------------|--------------------------------------------+
                                            | https/http
                                            v
                     OpenAI / Anthropic / OpenRouter (incl. :online web models) /
                     any OpenAI-compatible server / Ollama / llama-server
```

## The ops protocol (one JSON schema per host)

All panes run the same pipeline: free-text instruction -> profile/route ->
`/api/chat` -> reasoning strip (hidden repair round for unclear output) -> JSON ops
extraction -> validation against document/email context -> **approval gate** ->
Office.js apply. Models may answer conversationally (`{"kind":"reply",...}`) or
emit an ops list (`slide_ops` / `doc_ops` / bare `{"ops":[...]}`); the panes
share the gate, the sessions, the Inspector log and the Settings model.

- SlideKick ops: `create_slide`, `edit_slide`, `replace_text`, `set_notes`,
  `create_table`, `move_shape`, `create_shape`, `create_image`, `apply_layout`,
  `format_text`, `format_selection`, `move_slide`, `set_background`,
  `add_hyperlink`, `set_alt_text`, `delete_slide`.
- DocKick ops: `insert_paragraphs` (built-in styles), `edit_paragraph`,
  `replace_text`, `create_table`, `insert_comment`, `delete_paragraph`,
  `rewrite_doc`, `insert_image`, `format_text`, `create_template`.
- MailKick ops (never sends): `draft_reply`, `compose_body`, `set_categories`,
  `save_draft`, `new_message`.

Approval policy: per-op `auto` / `confirm` / `deny` in the shared config; the
destructive-heavy ops default to `confirm` (deletes, layout moves, full
rewrites, compose-body rewrites), additive ops default to `auto`. The gate is
enforced in the pane; the server never touches Office documents. The gate
offers Select all, a per-run **auto-apply** checkbox, a global
**Auto apply approved actions** setting (Settings -> Approval policy), a
"don't ask again this session" remember per op type, and a red **X** per
proposed action to remove it from the list (removal needs a second click to
confirm). Applications show green, deletions red.

## Sessions, history and the Previous-requests box

- Every exchange is persisted immediately in a per-app session
  (`~/.officekick/sessions/`), named by the AI (2-4 words), resumable from the
  History tab (also inline-renamable); failures are saved as
  `(failed: <reason>)`.
- The server store is the source of truth: a save keeps the existing title and
  the longer of stored/received exchange lists (it never shrinks a session);
  timestamps (`ts`) are stored per message.
- Resuming loads the most recent exchanges into the model window and fills the
  collapsible **Previous requests** box under the prompt: one +/- box per run,
  timestamp first with the request text below it; expanding shows the full
  request (click it to refill an empty input) and the response summary. While
  a run works the typed request stays visible with the input locked; the input
  clears when the result prints (echoed as "Request: ..." above it).
- The model context per run = the last 6 exchanges of the session plus the
  fresh document/deck/mail context. Resuming a session still titled
  "Session" asks the AI for a better title and stores it.

## Web lookups

When a request needs up-to-date info, the model is instructed to reply with
`{"kind":"web_lookup"}`; the pane then re-asks the model once through the
profile's **other** slot with a live-web system prompt (OpenRouter models get
the `:online` web plugin automatically). Returned facts arrive with their
source URLs and are offered as cites. If the model reports no web access, the
pane answers without web data. No server-side search engine, no API keys
involved.

## Key design points

- **One server, many add-ins.** `core/server.py` serves each add-in's static
  tree plus one API surface. Connections/profiles are configured once and
  usable from every app.
- **One Deck interface per host.** `applyOps` + validation + index-shifting
  live in each pane; each Deck implementation (OfficeDeck/WordDeck/MailDeck/
  the simulator) exposes the same methods, so the browser simulator is an
  honest end-to-end test of the pipeline.
- **Append-only creation.** Office.js slides always append; the panes report
  the actual result (and final index) in the applied-ops log.
- **Capability probes + honesty.** Requirement sets (`PowerPointApi`/`WordApi`)
  are probed and missing APIs degrade to clear messages or manual-paste
  fallbacks (e.g. comments below WordApi 1.4). Models are told up front what
  this host cannot do (e.g. SmartArt/icons/3D are not creatable; MailKick
  cannot move, archive, delete or send mail).
- **Images.** Image-capable models emit a `create_image`/`insert_image` op;
  the pane preflights which configured models can generate (`/api/models?
  capability=image`), fetches the picture via `/api/image` and inserts it,
  with layered failure reporting.
- **MailKick never sends.** `draft_reply`/`new_message` open drafts you review
  and send yourself; `compose_body`/`save_draft` only work while composing.
- **Inspector.** Every request/response (including op batches and failures)
  is logged and inspectable from the pane; failed exchanges render red.
- **Update check: anonymous, token-free, git-free.** The server only reads the
  raw `VERSION` file from GitHub (10-minute cache, numeric compare,
  `?latest=` hook for testing). No token, no automatic git operations:
  updating = manual `git pull` + rerun the installer + reload the task panes.
- **Approval policy lives server-side, enforced client-side** (the server
  never touches any document/mailbox).
- **Installers** deploy the umbrella layout (`~/Library/Application
  Support/OfficeKick` / `%LOCALAPPDATA%\OfficeKick`), render per-add-in
  manifests (list-mode, verified with office-addin-manifest), sideload into
  each app's WEF / trusted catalog, register the auto-start agent and
  health-check the server; config survives updates.
