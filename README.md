# OfficeKick

AI assistant add-ins for **Microsoft Office**, from Corelan. They run on macOS
and Windows and share one technology stack.

- **Corelan SlideKick** - the PowerPoint add-in ([docs/slidekick.md](docs/slidekick.md))
- **Corelan DocKick** - the Word add-in ([docs/dockick.md](docs/dockick.md))
- **Corelan MailKick** - the Outlook add-in: summarize, translate, draft replies (you send), categories ([docs/mailkick.md](docs/mailkick.md); macOS Outlook = one-time loopback-TLS setup (needs a trusted https pane; works in New and classic Outlook) - [requirements](docs/mailkick.md#requirements-for-outlook-macos---read-before-installing); skipped automatically when they are not met)

## Contents

1. [Dependencies](#dependencies)
2. [How to install](#how-to-install) -
   [Windows](#windows) / [macOS](#macos) / [Linux](#linux)
3. [How to install - the installer itself](#the-installer-itself) - the menu, first run, update, uninstall
4. [Using it](#using-it) - where the panes appear, ports, the server,
   permissions
5. [How it works](#how-it-works) - the flow, security and privacy
6. [References](#references) - the full documentation index

## Dependencies

| Platform | What is required | Where it comes from |
|---|---|---|
| **Windows** | **Python 3** (the installer runs as `py -3`; plain `python` is often just the Store alias) + **Git**, but only if you want a clone of the repo | installed for you by the PowerShell script `tools/setup-windows.ps1` - see [Windows](#windows) below |
| **macOS** | **Python 3** (`python3` - Apple's command-line tools provide one; python.org's installer works too). **Git** only for a repo clone | usually already present; `python3 --version` checks. `git` ships with the command-line tools and with GitHub Desktop |
| **Linux** | see the honest note under [Linux](#linux) - the add-ins need desktop Office | - |

Everything installs **per user** - no administrator rights anywhere. The
Windows PowerShell script installs the latest Python 3 through python.org's
official **Python install manager** (never winget for Python) and Git for
Windows, then offers to clone the repository.

**Windows - one command from the internet** (Windows 10+ ships a real
`curl.exe` in System32, so the Linux pipe habit works; `irm | iex` is the
native form):

```powershell
irm https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | iex
```

**Windows - from a downloaded copy** of `tools/setup-windows.ps1`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup-windows.ps1
```

(The `-ExecutionPolicy Bypass` applies to that one process only - the policy
explanation, the flag list, `curl.exe`-pipe variants, and step-by-step
details: [docs/setup-windows.md](docs/setup-windows.md).)

## How to install

### Windows

1. Install the dependencies with the PowerShell script above (or accept the
   defaults: Python 3 + Git land on your PATH).
2. Make the folder and clone the repository **manually** (the script offers
   this too, but you never have to rely on it). In a normal command prompt
   (cmd):

```bat
mkdir "%USERPROFILE%\Tools" 2>nul
cd /d "%USERPROFILE%\Tools"
git clone https://github.com/corelan/officekick.git
cd officekick
py -3 install.py
```

   In PowerShell the equivalent:

```powershell
mkdir "$env:USERPROFILE\Tools" -Force
cd "$env:USERPROFILE\Tools\officekick"
```

   (A download-zip instead of the clone works the same way - just unpack it
   into the Tools folder and `cd` into the unpacked folder.)
3. Run the installer: `py -3 install.py` in that folder.

The Windows Office route for pane registration lives in [Using it](#using-it).

> **cmd vs PowerShell folder syntax**: cmd uses `%USERPROFILE%` and needs
> `cd /d` to switch drives; PowerShell uses `$env:USERPROFILE`. The examples
> above are labeled per shell.

### macOS

1. Confirm `python3 --version` works (command-line tools provide it).
2. Clone the repository and run the installer:

```bash
git clone https://github.com/corelan/officekick.git
cd officekick
python3 install.py
```

3. macOS specifics: locally hosted models (Ollama on your LAN) need one
   consent - see [macOS permissions](#macos-permissions-locally-hosted-models).

### Linux

The task panes host inside **desktop Microsoft Office**, which exists on
Windows and macOS only; `install.py` has no Linux branch and a Linux box
cannot register or serve the panes. There is no supported Linux install -
running the Office web client against the local server is untested and not
part of this project.

## The installer itself

One installer for macOS and Windows. It detects the platform by itself.

The interactive menu:

1. **Install / repair** - the add-in picker (Word + PowerPoint on by default,
   Outlook off until you enable it)
2. **Update**
3. **Server** (start / restart / stop / status)
4. **Transport options** - transport mode, listener ports, certificate mode
5. **Uninstall**

Details, the full flag list, and the self-healing behaviors:
[docs/installer.md](docs/installer.md).

The essentials:

- Everything talks to the local server. One trusted https listener serves
  every add-in the moment a usable certificate exists; plain http is the
  fallback (and the Windows default until you switch the transport).
  Outlook always needs https URLs in its manifest.
- **Certificates**: the default mode is a **self-signed leaf** (no CA is
  created; the leaf itself gets trusted - explicit consent, uninstall-safe).
  Alternatives, chosen through one question or the Transport menu: the
  classic loopback CA, or bring-your-own (see
  [docs/certificates.md](docs/certificates.md)). On Windows the flow verifies
  the trust out of band and offers the Root fallback Office requires
  (WebView2 leaf-trust is field-verified as rejected).

### Update

```bash
python3 install.py --update    # Windows: py -3 install.py --update
```

This pulls the repo with git and updates the install. Quit your Office apps
**before** running it (on macOS the installer refuses to touch the add-in
registration while an Office app is running), then reopen them so the task
panes load the new files.

If a deck or document ever greets you with an
"Add-in Error / This add-in is no longer available" toast on open, heal it
once with the repair tool - see the troubleshooting section in
[docs/slidekick.md](docs/slidekick.md).

### Uninstall

```bash
python3 install.py --uninstall            # keeps your config
python3 install.py --uninstall --remove-certs --purge-config   # full cleanup incl. certificates + config
# Windows: py -3 install.py --uninstall [--purge-config]
```

Or pick **Uninstall** in the installer menu. It asks two questions - the
certificate store entries and the config folder - and both default to KEEP:
an uninstall never destroys your setup unless you say so. After an uninstall
the certificate mode resets to self-signed leaf. Details:
[docs/installer.md > Uninstall](docs/installer.md).

## Using it

### Where the panes appear

- **PowerPoint (macOS)**: Home tab -> **Add-ins** (dropdown) ->
  **Developer Add-ins** -> **Corelan SlideKick**.
- **Word (macOS)**: same route -> **Corelan DocKick**.
- **PowerPoint / Word (Windows)**: Home -> **Add-ins** -> **Advanced** ->
  the classic "Office Add-ins" dialog -> **SHARED FOLDER** (top) -> the
  add-in -> **Add** (once per application). Note: the newer "Apps" dialog
  (Add-ins -> Manage) only lists store and organization add-ins - it never
  shows local folder catalogs.
- **Outlook (MailKick)**: the manifest lands in Outlook's own add-in folder
  (macOS); quit and reopen Outlook. If MailKick does not appear:
  Get Add-ins > My add-ins > Custom Add-ins > Add a custom add-in >
  Add from File (Cmd+Shift+G). See [docs/mailkick.md](docs/mailkick.md).

### First run

1. **Connections** tab: add a provider (Type + Base URL + optional API key).
   "Test connection" lists its models.
   See [settings.md > Connections](docs/settings.md#settings--connections).
2. **Profiles** tab: pick the two model routes - **content** (edits and
   generation) and **other** (advice and brainstorming). Optional slots:
   **Image model** (picture generation) and **Web search model** (a lookup
   channel the pane can use itself).
   See [settings.md > Profiles](docs/settings.md#settings--profiles).
3. **Settings > Search**: the steps the model's web lookups follow. No
   steps configured = the free keyless DuckDuckGo engine.
   See [settings.md > Search](docs/settings.md#settings--search).
4. Type the request in the **Sessions** tab and press **Run**
   (Cmd/Ctrl+Enter). The run reports every step in the activity log under
   the status line; the approval policy decides what gets applied - set it
   in Settings > Approval policy
   ([settings.md > Approval policy](docs/settings.md#settings--approval-policy)).

### Ports and how to tell which scheme a pane uses

Every pane answers on `http://localhost:8765` and, once a working
certificate exists, on `https://localhost:8766` (both configurable - one
listener per scheme serves all three add-ins; with the transport mode on
`https` the https listener is required, and Outlook always needs https
manifest URLs so MailKick stages only when https works).

The **bottom status line** in a task pane tells the truth per pane: a 🔒
plus the https port means TLS; the plain port in blue (light theme) / cyan
(dark) means http. On disk, check the manifest your Office container holds,
e.g. `grep -o 'http[^"]*taskpane.html' "~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/mailkick.xml"`.

### The local server: starting it manually and automatically

If the pane's footer text (next to the heartbeat dot) turns red with an
OS-specific hint, start the server from your OfficeKick checkout:

```bash
python3 install.py --start      # macOS
py -3 install.py --start        # Windows
```

`--start` brings the server up and health-checks it; it changes nothing
else. Related flags: `--status`, `--stop`. In the interactive menu the same
action is **Server**.

How it starts automatically:

- **macOS - login item.** The installer registers the `OfficeKickServer`
  login item (the app bundle in `~/Applications`). It runs at login with
  no visible window.
- **Windows - scheduled task.** The per-user task `OfficeKickServer` (runs
  at logon, hidden, hardened at every install: no stop-on-idle, no
  stop-on-batteries, no execution time limit). Inspect with
  `schtasks /query /tn OfficeKickServer /v`; the pane's footer hint shows
  `schtasks /run /tn OfficeKickServer` as the quickest restart.

Check whether it is up: open `http://127.0.0.1:8765/api/health`, or
`python3 install.py --status`. The server also keeps itself updated on a
git install (a pull + restart on newer releases; other installs update via
the GitHub zip) - your config, profiles and sessions survive every update.

Config lives in `~/.officekick/config.json` (`%USERPROFILE%\.officekick` on
Windows, chmod 600). API keys never leave your machine except to the provider
you configured.

### macOS permissions (locally hosted models)

Providers on your LAN need one macOS consent:
**System Settings -> Privacy & Security -> Local Network**.

- **When it appears**: the first time the server connects to a LAN address
  (for example an Ollama host). macOS shows the
  "...wants to access devices on your local network" dialog. Accept once.
- **If the dialog never appears**: open the Local Network pane and enable the
  **OfficeKickServer** entry. It can be listed as "python" - same thing. The
  entry may only appear after one failed attempt.
- **The symptom without the grant**: Test connection says
  "Cannot reach 192.168.x.x ... ([Errno 65] No route to host)" while the same
  provider answers from a Terminal. Only private-IP providers hit this; cloud
  providers are not affected.
- Windows and Linux have no such gate.

Do not use `tccutil reset` to re-trigger this: on current macOS builds it
cannot reset the LocalNetwork service, and `tccutil reset All` wipes ALL saved
privacy decisions for every app.

## How it works

1. You type what you want, in plain language.
   Examples: `reword the current slide`, `make slides about threat hunting`,
   `add a table of the key findings`, `leave a review comment on this paragraph`.
2. The add-in reads your document.
3. The request goes to an AI model that **you** configured.
4. The add-in shows the proposed changes.
5. Your approval policy decides what gets applied.

Everything runs on your machine:

- One small local server (Python standard library only) serves the add-ins
  and the AI requests. Loopback only: `127.0.0.1` + `::1`.
- You configure providers once; every Office app shares them.
- Supported providers: OpenAI, Anthropic, OpenRouter, any OpenAI-compatible
  server, Ollama, llama.cpp. Nothing goes through a third-party service.
- Every exchange is auto-saved in a resumable local session. The
  **Previous requests** box shows each run with its timestamp.
- MailKick adds one-click **reply+ / reply-** buttons for a drafted polite reply.

### Security & privacy

- The local server binds to loopback only (`127.0.0.1` + `::1`).
- No telemetry, no phone-home. The server never touches your documents.
- API keys stay in the local config file. They are attached to provider calls
  server-side only; the pane and the API never show or return them.
- File content leaves your machine only to the providers you configured.
- Debug logging is opt-in. It writes prompts and document text to a local
  daily log; keep it off for sensitive work.

## References

- [docs/setup-windows.md](docs/setup-windows.md) - the Windows dependency
  installer (`tools/setup-windows.ps1`: the Python install manager + Git,
  the repo clone into `%USERPROFILE%\Tools`) + the PowerShell execution
  policy how-to.
- [docs/installer.md](docs/installer.md) - the installer menus, transport
  options, certificate modes, uninstall (asks + defaults, flags).
- [docs/certificates.md](docs/certificates.md) - how the local https
  certificates are generated, trusted, renewed, and removed, per platform.
- [docs/settings.md](docs/settings.md) - every setting explained: what it
  means, how to configure it (config, profiles, approvals, search steps).
- [docs/slidekick.md](docs/slidekick.md) - the PowerPoint add-in: actions,
  approvals, design passes, simulator, troubleshooting.
- [docs/dockick.md](docs/dockick.md) - the Word add-in: document actions,
  comments fallback, simulator.
- [docs/mailkick.md](docs/mailkick.md) - the Outlook add-in: summarize,
  translate, draft replies + forwards (you send), categories, templates,
  simulator, troubleshooting.
- [docs/architecture.md](docs/architecture.md) - how the stack works.
- [CHANGELOG.md](CHANGELOG.md) | [LICENSE](LICENSE)

Licensed by **Corelan Consulting bv** - free to use with attribution, use at
your own risk, may not be sold or commercialized, modifications must reference
the original author.
