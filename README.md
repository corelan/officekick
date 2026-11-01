# OfficeKick

AI assistant add-ins for **Microsoft Office** (macOS and Windows) on one shared
technology stack, from Corelan:

- **Corelan SlideKick** - the PowerPoint add-in (details: [docs/slidekick.md](docs/slidekick.md))
- **Corelan DocKick** - the Word add-in (details: [docs/dockick.md](docs/dockick.md))
- **Corelan MailKick** - the Outlook add-in: summarize/translate, draft replies (you send), categories ([docs/mailkick.md](docs/mailkick.md))

Type what you want, in plain language: `reword the current slide`,
`make slides about threat hunting`, `add a table of the key findings`, `leave a
review comment on this paragraph`. The add-in reads your file, sends it with
your request to the AI model **you** configured, shows the proposed changes,
and applies them only when your approval policy allows it.

Every exchange is auto-saved in a resumable local session; the pane's
**Previous-requests box** shows each run with its timestamp (+/- to expand).
MailKick adds **reply + / reply −** one-click buttons for a drafted polite
reply. Nothing goes through a third-party service: the add-ins talk to the providers
*you* configure (OpenAI, Anthropic, OpenRouter, any OpenAI-compatible server,
Ollama, llama.cpp), through one small local server on your own machine.
Connections, model profiles and the approval policy are configured **once** and
shared by every Office app.

## Install / update

macOS:

```bash
git clone https://github.com/corelan/officekick.git
cd officekick
./install.sh
```

Windows (PowerShell, no admin needed):

```powershell
git clone https://github.com/corelan/officekick.git
cd officekick
.\install.ps1
```

Both installers are idempotent (re-run to update): they check `python3` (>= 3.8),
detect the installed Office apps, deploy files to
`~/Library/Application Support/OfficeKick` (macOS) or `%LOCALAPPDATA%\OfficeKick`
(Windows), register the manifests into each Office app, set the auto-start
service (LaunchAgent / Scheduled Task; the server also restarts on crash), start
the server and health-check it.

### How to update OfficeKick

The server itself never runs `git` and needs no tokens. It only *detects*
updates:

- **Check**: Settings -> "Check now" (or the **Check for updates automatically
  at startup** option, on by default). It reads this repo's `VERSION` file from
  GitHub, compares with the running version and shows a plain-text notice.
- **Apply** (macOS):

  ```bash
  cd officekick
  git pull
  ./install.sh     # re-deploys + restarts the server, stays idempotent
  ```

- **Apply** (Windows):

  ```powershell
  cd officekick
  git pull
  .\install.ps1
  ```

- **Then**: quit and reopen the Office app(s) so the task panes load the new
  static files. Your config (`~/.officekick/config.json`) and sessions survive
  updates. There is no automatic `git pull`, ever.

Useful flags: `--port N` / `-Port N`, `--dev` / `-Dev` (serve from the repo),
`--no-launchagent` / `-NoScheduledTask`, `--https` (macOS: local CA),
`--ribbon` (keep the commands/AI group in the manifest; default = list mode),
`status`, `run`, `uninstall [--purge]`.

Config (API keys, profiles, policy) lives in `~/.officekick/config.json`
(`%USERPROFILE%\.officekick`, chmod 600) and survives updates.

## Use it

- **PowerPoint (macOS)**: Home tab -> **Add-ins** (dropdown) -> **Developer
  Add-ins** -> **Corelan SlideKick** (with `--ribbon`: the **AI** group button
  on the Home tab instead, if your build renders it).
- **Word (macOS)**: same route -> **Corelan DocKick**.
- **PowerPoint/Word (Windows)**: Insert -> My Add-ins -> SHARED FOLDER ->
  Corelan SlideKick / Corelan DocKick -> Add (once).
- **Outlook (MailKick)**: the installer places the manifest in Outlook's own
  add-in folder (`~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/mailkick.xml`);
  **quit and reopen Outlook**. If it does not appear: Outlook -> Get Add-ins ->
  **My add-ins -> Custom Addins -> Add a custom add-in -> Add from File** - press
  **Cmd+Shift+G** in the picker and paste that folder/path. Windows: same trusted
  catalog as the other two add-ins. See [docs/mailkick.md](docs/mailkick.md).
- **Ribbon "AI" group (opt-in `--ribbon`)**: `./install.sh --ribbon` keeps the
  manifest's commands - an **AI** group with the add-in icon and an
  "Open <add-in>" button on the PowerPoint/Word **Home** tab (MailKick: on the
  message read/compose tabs). Default is `--no-ribbon` (list-mode manifest):
  the Office-Mac build we tested accepts the command manifest, REPLACES the
  Add-ins-list entry with it, and then renders neither the group nor a list
  item - making the add-in unreachable. Verify your build renders the group
  before staying on this mode.

First run: **Connections** tab - add a provider (type + Base URL + optional API
key, "Test connection" lists its models); **Profiles** tab - pick the *content
model* (edits/generation) and the *other-actions model* (advice/brainstorming)
per profile; then type instructions in the pane and Run (Cmd/Ctrl+Enter).

## Security & privacy

- The local server binds to loopback only (`127.0.0.1` + `::1`); no telemetry,
  no phone-home, server never touches your documents.
- API keys stay in the local config file and are attached to provider calls
  server-side only - never shown in the pane, never returned by the API.
- File content leaves your machine only to the provider(s) you configured.
- Opt-in debug logging writes prompts/document text to a local daily log - keep
  it off for sensitive work.

## Uninstall

```bash
macOS:   ./install.sh uninstall           # keeps config
         ./install.sh uninstall --purge   # also wipes config + logs
Windows: .\install.ps1 -Action uninstall
         .\install.ps1 -Action uninstall -Purge
```

Then quit and reopen the Office apps.

## More

- [docs/slidekick.md](docs/slidekick.md) - the PowerPoint add-in: actions,
  approvals, design passes, simulator, troubleshooting.
- [docs/dockick.md](docs/dockick.md) - the Word add-in: document actions,
  comments fallback, simulator.
- [docs/mailkick.md](docs/mailkick.md) - the Outlook add-in: summarize /
  translate, draft replies + forwards (you send), categories, templates,
  simulator, troubleshooting.
- [docs/architecture.md](docs/architecture.md) - how the stack works.
- [CHANGELOG.md](CHANGELOG.md) | [LICENSE](LICENSE)

Licensed by **Corelan Consulting bv** - free to use with attribution, use at
your own risk, may not be sold or commercialized, modifications must reference
the original author.