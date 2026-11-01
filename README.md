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

macOS and Windows take the SAME unified installer (it auto-detects the
platform):

```bash
git clone https://github.com/corelan/officekick.git
cd officekick
python3 install.py        # Windows: python install.py
```

It shows an interactive menu - Install / Update / Uninstall - lists what it
will change and which OS prompts may appear, and asks whether to autostart the
local server (default yes: a login item via SMAppService on macOS /
a windowless per-user scheduled task on Windows; answering no keeps the server
manual). No admin rights are required on either platform. The add-in transport
is plain `http://localhost:8765` (Office allows localhost over http - no
certificates involved on any path). Non-interactive equivalents:
`python3 install.py --install`, `--update` (= `git pull` + update),
`--uninstall [--purge-config]`, `--dry-run` to preview.

The install is idempotent (re-run to update/repair): it checks `python3`
(>= 3.8), detects the installed Office apps, deploys files to
`~/Library/Application Support/OfficeKick` (macOS) or `%LOCALAPPDATA%\OfficeKick`
(Windows), renders + registers/sideloads the add-in manifests (macOS WEF
containers; Windows per-user trusted folder catalog via HKCU - Outlook on
Windows needs a one-time "Add from File" dialog, the installer prints the
steps), registers autostart, starts the server and health-checks it. The
macOS server wrapper also restarts the server on crash. Platform-specific
legacy installers remain available: `install.sh` (macOS) / `install.ps1`
(Windows).

### How to update OfficeKick

The server itself never runs `git` and needs no tokens. It only *detects*
updates:

- **Check**: Settings -> "Check now" (or the **Check for updates automatically
  at startup** option, on by default). It reads this repo's `VERSION` file from
  GitHub, compares with the running version and shows a plain-text notice.
- **Apply** (both platforms):

  ```bash
  cd officekick
  python3 install.py --update    # Windows: python install.py --update
  ```

  (--update = git pull + re-deploy + restart; equivalently: manual `git pull`,
  then a plain `python3 install.py` run through the menu.)

- **Then**: quit and reopen the Office app(s) so the task panes load the new
  static files. Your config (`~/.officekick/config.json`) and sessions survive
  updates. There is no automatic `git pull`, ever (except through an explicit
  `--update`).

Useful flags: `--port N`, `--ribbon` (keep the commands/AI group in the
manifest; default = list mode), `--install` / `--update` /
`--uninstall [--purge-config]`, `--dry-run`; legacy installers keep their own
flags (`install.sh`: `--dev`, `status`, `run`).

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

## macOS permissions (locally hosted models)

Providers on your LAN (Ollama, llama.cpp, any private-IP base URL) need one
additional macOS consent: **Privacy & Security -> Local Network**.

- **What you'll see**: the server (started by the installer via
  `~/Applications/OfficeKickServer.app`) raises macOS's
  "...wants to access devices on your local network" dialog the first time it
  connects to a LAN address. Accept it once - done.
- **If the dialog never appears**: open
  System Settings > Privacy & Security > **Local Network** and enable the
  **OfficeKickServer** entry (it can be listed as **"python"** - same thing).
  The entry may only appear after one failed attempt.
- **The symptom before granting**: Test connection says
  "Cannot reach 192.168.x.x ... ([Errno 65] No route to host). Is the service
  running?" - while `curl http://<ip>:11434/api/tags` from a Terminal works.
  That mismatch IS the missing Local Network grant (0.9.16+ adds this hint to
  the error message itself). Cloud providers (OpenAI/Anthropic/OpenRouter)
  are not affected - only private-IP addresses are behind this wall.
- **Restarting the server from a permitted session** (`./install.sh run` in a
  Terminal) bypasses the grant because the terminal's permission applies.
- Not applicable on Windows/Linux - macOS-only privacy gate.
- Note: `tccutil reset` cannot reset the LocalNetwork service on current
  macOS builds, and broad resets like `tccutil reset All` wipe ALL your saved
  privacy decisions (every app re-asks) - do not use them to "re-trigger"
  this prompt.

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