# SlideKick

An AI assistant add-in for **Microsoft PowerPoint** on macOS and Windows. Type what
you want - `reword the current slide`, `make 4 slides about threat hunting`, `add
speaker notes to this deck` - and SlideKick reads your deck, routes the request to the
model you chose, and applies the changes in PowerPoint **after** the per-action
approval policy allows it.

You pick your own AI provider and models. Nothing goes through a third-party service:
SlideKick talks to the providers *you* configure, from a small local server on your own
machine.

```
+--------------------------+        +---------------------------+      +--------------------+
| PowerPoint (Mac/Windows) |        | SlideKick local server    |      | Your AI providers  |
| add-in (Office.js        | -----> | 127.0.0.1:8765            | ---> | OpenAI / Anthropic |
| task pane)               |        | - serves the task pane UI |      | OpenRouter         |
|                  |        | - stores connections,     |      | OpenAI-compatible  |
| applies slide    |        |   profiles, approval      |      | (LM Studio, vLLM,  |
| operations via   |        |   policy, debug logs      |      |  llama-server, ...)|
| Office.js        |        | - proxies AI calls        |      | Ollama             |
+--------------------------+        +---------------------------+      +--------------------+
```

## Features

- **Free-text instructions** for deck work: reword, restructure, tighten bullets,
  generate new slides from a topic, write speaker notes, design/layout suggestions.
- **Your providers, your models**: six connection types - `openai`, `anthropic`,
  `openrouter`, `generic-openai` (any OpenAI-compatible server), `ollama`, `llamacpp`
  (llama-server) - each with Name / Base URL / optional API key.
- **Profiles** pair a *content model* (slide edits & generation) with an
  *other-actions model* (chat, brainstorming, advice). Model lists are fetched live from
  each provider.
- **Per-action approval policy** (Settings): `auto`, `confirm` (must be ticked before
  Apply) or `deny` per action type. Deleting slides and layout changes require approval
  by default.
- **Reasoning-model friendly**: chain-of-thought output (`<think>`, `<think>`, `<thinking>`,
  `<reason>`, `<reasoning>`, Ollama `thinking` fields, OpenAI-style
  `reasoning_content`) is stripped; unclear output triggers a hidden
  "answer in clean JSON" retry that is never shown.
- **Debug logging** (opt-in): appends AI interactions and deck-edit traces to one
  appended file per day (`~/.slidekick/logs/debug-YYYY-MM-DD.log`).
- **Installer**: detects PowerPoint, deploys the add-in, sideloads/registers the
  manifest (LaunchAgent on macOS, per-user trusted catalog + Scheduled Task on
  Windows) and keeps the server running. Re-run either installer to update.
- **Always visible in the ribbon**: a dedicated **SlideKick** tab with a task-pane
  button (add-in commands), so the add-in is one click away - no Add-ins dropdown
  detour (the dropdown entry still lists it).
- **Theme-aware UI**: light/dark follows the Office host theme
  (`OfficeTheme.isDarkTheme`, color-luminance fallback); in the simulator it follows
  the browser scheme.
- **Simulator**: the exact task pane runs in a normal browser against an in-memory deck,
  for testing without PowerPoint.

## Requirements

- macOS (install.sh) or Windows 10/11 (install.ps1)
- Microsoft PowerPoint from Microsoft 365. Slide editing needs PowerPointApi ~1.8
  (Mac 16.96 / 2504+, Windows 2504+ builds; perpetual/LTSC editions run the pane
  read-only). A current build is standard on Microsoft 365.
- python3 >= 3.8 (Apple Command Line Tools or Homebrew; on Windows python.org
  or the py launcher)
- An AI provider reachable from your machine (cloud API or local Ollama/llama-server).

## Install / update (macOS)

```bash
git clone https://github.com/corelan/slidekick.git
cd slidekick
./install.sh
```

The installer:

1. checks python3, finds `/Applications/Microsoft PowerPoint.app`,
2. deploys files to `~/Library/Application Support/SlideKick/`,
3. renders `src/manifest.template.xml` with the chosen port and copies it to
   `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/slidekick.xml`
   (the documented manual-sideload location; `wef` is created if missing),
4. installs `~/Library/LaunchAgents/com.corelan.slidekick.plist` (starts the server at
   login, restarts it on crash) and starts it now,
5. health-checks `http://127.0.0.1:<port>/api/health`.

Useful flags: `--port N`, `--dev` (serve straight from the repo), `--no-launchagent`,
`status`, `run` (foreground server), `uninstall [--purge]`.

**Updating**: re-run `./install.sh`. It overwrites the deployed files, replaces the
manifest, and reloads the LaunchAgent. Your connections/profiles/policy live in
`~/.slidekick/config.json` and are preserved. Quit and reopen PowerPoint afterwards so
it re-reads the manifest.

**In PowerPoint**: open a deck, then click **SlideKick** in the ribbon (its own tab;
falls back to Home -> Add-ins -> Corelan SlideKick). The task pane opens on the right.

## Install on Windows (PowerShell)

Same add-in, same server, same config format and Settings - only the registration
layer differs (per-user trusted catalog + Scheduled Task instead of Wef folder +
LaunchAgent). No administrator rights required.

```powershell
git clone https://github.com/corelan/slidekick.git
cd slidekick
.\install.ps1
```

The installer:

1. finds python3 (>= 3.8; python.org install or the `py` launcher - the Microsoft
   Store stub is skipped) and locates `POWERPNT.EXE` (App Paths registry or the
   Office16 directories),
2. deploys to `%LOCALAPPDATA%\SlideKick` via robocopy (update = idempotent), config
   lives in `%USERPROFILE%\.slidekick` (same as macOS),
3. renders the manifest and registers its folder as a per-user trusted catalog
   (`HKCU\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs`),
4. registers Scheduled Task `SlideKickServer` (logon trigger, restart-on-crash); if
   task registration is blocked it falls back to a Startup-folder .cmd,
5. starts the server now and health-checks it (loopback only - no firewall prompt).

Useful flags: `-Port N`, `-Dev`, `-NoScheduledTask`, `-Action status`, `-Action run`,
`-Action uninstall [-Purge]`.

**In PowerPoint (Windows)**: quit and reopen PowerPoint, then **Insert -> My Add-ins
(dropdown) -> SHARED FOLDER -> Corelan SlideKick -> Add** once; from then on the
dedicated **SlideKick** ribbon tab is always there.

**Windows notes**:

- Needs a current Microsoft 365 build (WebView2 add-in runtime). Old perpetual builds
  whose add-in runtime is IE11-based cannot load the pane.
- Perpetual/LTSC editions cap below PowerPointApi 1.8: the pane runs read-only there
  (reading and advice work, edits are refused with a clear message).
- Title/body placeholders are detected via `placeholderFormat` on modern builds
  (localization-safe, also used on macOS); older builds fall back to shape names.

## First run

1. **Connections** tab: add a connection.

   | Type | Default Base URL | API key |
   |---|---|---|
   | OpenAI (native) | `https://api.openai.com/v1` | required |
   | Anthropic (native) | `https://api.anthropic.com` | required |
   | OpenRouter | `https://openrouter.ai/api/v1` | required |
   | OpenAI-compatible (generic) | *your server* (`http://localhost:1234/v1`, vLLM, LM Studio...) | optional |
   | Ollama | `http://localhost:11434` | usually empty |
   | llama.cpp (llama-server) | `http://localhost:8080` | usually empty |

   The Base URL is editable for every type (proxies, Azure-style gateways, ...).
   "Test connection" lists the models the provider reports. A connection that is
   referenced by a profile cannot be deleted until the profile stops using it.
   Editing is move-based: renaming a connection re-points the profiles using it;
   renaming a profile keeps the default star with it. Model lists in the profile
   editors are sorted alphabetically.

2. **Profiles** tab: name a profile, then for the *content model* and the
   *other-actions model* pick a connection -> pick a model from the live list (or type a
   custom model id). Mark one profile as default.

3. **Chat / Actions** tab: type an instruction and press Run (Cmd/Ctrl+Enter). Quick
   chips prefill common requests for you.

## Typed instructions, routes and approvals

- The **route** decides which model runs: `auto` sends slide-related requests
  (reword / create / make slides / notes / bullets...) to the **content model** and
  everything else to the **other-actions model**; force one with the Route dropdown.
- With *include deck context* checked, the pane sends a compact outline (slides,
  shape indices, texts, selection) with your request, so instructions like
  "reword the current slide" work without naming slide numbers.
- The model returns a structured action list. SlideKick **validates** every action
  against the outline (indices, shape references), shows the proposal, and applies it
  only after the approval policy is satisfied:

  | Action | What it does | Default policy |
  |---|---|---|
  | create slide | append a new slide (title + bullets + notes) | auto |
  | edit slide | replace title / body / notes of an existing slide | auto |
  | replace text | reword one shape in place (the "reword" path) | auto |
  | set notes | write speaker notes | auto |
  | apply layout | switch a slide's layout (design change, text kept) | confirm |
  | delete slide | remove a slide | confirm |
  | reply | conversational answer - never touches the deck | n/a |

  Change any of these in **Settings**. `confirm` shows a checkbox per action (unchecked
  by default), with a *don't ask again this session* option. `deny` blocks the action
  with a visible notice. The auto-apply checkbox only shortcuts the extra Apply click -
  it never bypasses the policy.

- Notes that cannot be written automatically on older Office builds fall back to a
  copy-paste box in the pane.

## Supported requests in real life

- `reword the current slide` - replaces wording, keeps meaning and slide count
- `make 4 slides about threat hunting` - appends 4 new slides (Office.js can only
  append; drag them into place if needed)
- `tighten the bullets on this slide`, `turn this text into a table suggestion` (reply)
- `add speaker notes to slides 2-5`
- `suggest 3 design improvements for the current slide` - answer via the other model
- `set the layout of slide 3 to title only`

## Security & privacy

- The server binds to **127.0.0.1** only. API keys are stored in
  `~/.slidekick/config.json` (chmod 600) and are never returned by the API nor sent to
  the pane; they are only attached server-side to outbound provider calls.
- Slide texts and prompts leave your machine **only** to the provider(s) you configured,
  directly from the local server.
- Debug logging (off by default) writes slide content and provider payloads to local
  disk - keep it off for sensitive decks.

## Troubleshooting

- **Add-in doesn't show** - quit *all* Office apps and reopen PowerPoint; look for
  the **SlideKick** ribbon tab first, then the Add-ins dropdown:
  - macOS: check `./install.sh status`; confirm
    `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/slidekick.xml`
    exists; the Office cache can be cleared by quitting Office and removing
    `~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/*` with care.
  - Windows: check `.\install.ps1 -Action status`; confirm the trusted-catalog
    registry key exists (`HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\C7B1F2A9-...`)
    and that `Insert -> My Add-ins` shows the SHARED FOLDER section.
- **"Cannot reach the SlideKick server" in the pane**:
  - macOS: `launchctl list | grep slidekick`, then `tail -50 ~/Library/Logs/SlideKick/server.log`,
    or just re-run `./install.sh`.
  - Windows: `Get-ScheduledTask SlideKickServer`, then re-run `.\install.ps1`
    (server log: `%LOCALAPPDATA%\SlideKick\logs\server.log`).
- **401/403 from a provider** - fix the key on the Connections tab (blank key field
  keeps the stored key).
- **"lacks the text-editing APIs"** - update Microsoft 365 (need PowerPointApi 1.8:
  Mac 16.96+, Windows M365 2504+).
- **Model answers arrive as plain text** - the model ignored the JSON contract; with
  debug logging on, the daily debug log (`<userhome>/.slidekick/logs/debug-YYYY-MM-DD.log`)
  shows exactly what it sent. Reasoning-wrapped output is cleaned automatically.
- **Wrong port** - re-run the installer with `-Port`/`--port`; it re-renders the
  manifest and restarts.

## Development

```bash
macOS:  ./install.sh --dev         # serve straight from the repo (edit + reopen pane)
Win:    .\install.ps1 -Dev

npx -y office-addin-manifest validate "<install-dir>/manifest/slidekick.xml"  # official manifest validator
python3 tests/test_slidekick_server.py       # server + adapter test suite (19 tests)
python3 tools/mock_provider.py --flavor openai --port 8901   # mock provider for the UI
open http://127.0.0.1:8765/simulator.html    # full task pane + in-memory deck, no PowerPoint
python3 tools/make_icons.py                  # rebuild icons from logo_corelan.png (sips)
```

Note on the validator: it hard-requires HTTPS and has no localhost exception, so it
reports 3 "insecure URL" errors for the `http://localhost:...` panel/icon URLs.
Office itself accepts `http://localhost` for sideloaded/Uploaded dev add-ins (see
the Microsoft sideloading docs), which is what SlideKick uses; an optional HTTPS
mode with a trusted localhost certificate is planned.

The task pane runs unmodified in the simulator via `sim.js` (same Deck interface the
Office.js adapter implements), so prompt -> JSON ops -> validation -> approvals ->
apply is fully testable in a browser.

## Uninstall

```bash
macOS:  ./install.sh uninstall           # keeps ~/.slidekick config
        ./install.sh uninstall --purge   # also wipes config + debug logs
```

```powershell
Windows: .\install.ps1 -Action uninstall           # keeps config
         .\install.ps1 -Action uninstall -Purge     # also wipes config
```

Then quit and reopen PowerPoint.

## Roadmap

- Insert slides at an arbitrary position (blocked on Office.js support; currently append-only)
- Streaming responses, deck thumbnail preview in the pane
- Theme/style application, image generation via providers that support it
- Optional HTTPS for the local server

## License

Licensed by **Corelan Consulting bv** - free to use with attribution, use at your own
risk, may not be sold or commercialized, modifications must reference the original
author. See [LICENSE](LICENSE).