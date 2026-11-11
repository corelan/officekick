# OfficeKick

AI assistant add-ins for **Microsoft Office**, from Corelan. They run on macOS
and Windows and share one technology stack.

- **Corelan SlideKick** - the PowerPoint add-in ([docs/slidekick.md](docs/slidekick.md))
- **Corelan DocKick** - the Word add-in ([docs/dockick.md](docs/dockick.md)): document actions plus one-click in-place document translation (Translate button, tone-preserving, technical terms kept)
- **Corelan MailKick** - the Outlook add-in (New and classic, Windows/macOS): summarize, translate, draft replies (you send), categories. Needs the https prerequisites met (the installer handles and stages it automatically; details: [docs/mailkick.md](docs/mailkick.md))

## Contents

1. [Dependencies](#dependencies) - how to install Python 3 and Git:
   [Windows](#windows) (script or manual) / [macOS](#macos)
2. [Get the code and run the installer](#get-the-code-and-run-the-installer) -
   [Windows](#get-the-code-and-run-the-installer---windows) /
   [macOS](#get-the-code-and-run-the-installer---macos) - what the installer
   does, updates, uninstall
3. [Activate the add-ins](#activate-the-add-ins) - step by step per product:
   [Word](#word) / [PowerPoint](#powerpoint) / [Outlook](#outlook), Windows
   and macOS separately - plus [first run](#first-run-getting-started)
4. [Using it](#using-it) - ports, the server, macOS permissions
5. [How it works](#how-it-works) - the flow, security and privacy
6. [References](#references) - the full documentation index

## Dependencies

Two things must exist before `install.py` runs: **Python 3** (the installer
runs as `py -3` on Windows) and **Git**. Git is technically optional - but it
is the way to get the code AND the way the installer keeps your install
current (`py -3 install.py --update` = pull + redeploy + restart; also one
click from the pane on git installs). Use git unless you have a specific
reason not to: without it, updates re-download a zip instead of a `git pull`.

The add-ins also use the `cryptography` package for the https certificates -
the installer checks for it at START and installs it itself if needed
(`pip --user`); you never install it by hand.

**OfficeKick runs on Windows and macOS only.** Its panes host inside desktop
Microsoft Office, which does not exist on Linux; there is no Linux build and
no Linux support anywhere in the project. All commands below are labeled
with the prompt they belong to:
**PowerShell** (the blue Windows shell - type `powershell` in the Start
menu), **Command prompt** (the black Windows `cmd.exe` window), and
**Terminal** (the macOS shell).

### Windows

**By script (recommended).** Three equivalent command variants follow, so
your setup matches one of them exactly. All install the same two things,
per user, without admin rights: Python 3 through python.org's official
**Python Install Manager** (never winget for Python) and Git for Windows
(winget, with a GitHub-release fallback).

The prompt matters - Windows has TWO different command prompts:

- **PowerShell** = the blue window. Open it: Start menu, type `powershell`,
  press Enter.
- **Command Prompt** (cmd) = the black window with `C:\Users\...>` lines.
  Open it: Start menu, type `cmd`, press Enter.

**1. Run in a POWERSHELL window** (this line does NOT exist in the Command
Prompt - typing it there fails with "'irm' is not recognized"):

```powershell
irm https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | iex
```

**2. Run in a COMMAND PROMPT window** (Windows 10/11 include `curl.exe`;
this form is native to the Command Prompt and does not depend on the
execution policy):

```bat
curl.exe -fsSL https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | powershell -NoProfile -ExecutionPolicy Bypass -Command -
```

**3. From a downloaded copy** of `tools/setup-windows.ps1` (saved from the
repository or a zip) - works in EITHER prompt, because the command line
starts its own PowerShell:

```bat
rem Command Prompt:
cd /d C:\where\you\saved\it
powershell -NoProfile -ExecutionPolicy Bypass -File setup-windows.ps1
```

```powershell
# PowerShell (equivalent):
cd C:\where\you\saved\it
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup-windows.ps1
```

The script stops there: it is a **dependencies** script - it does not clone
the repository and it does not run the installer. Full details, flags, and
the execution-policy explanation: [docs/setup-windows.md](docs/setup-windows.md).

**Alternative: no script at all** (four commands, in a POWERSHELL window):

```powershell
# 1. the official Python Install Manager, then the runtime
Add-AppxPackage -AppInstallerFile https://www.python.org/ftp/python/pymanager/pymanager.appinstaller
py install default
# 2. Git (winget, or the installer from git-scm.com/download/win)
winget install -e --id Git.Git --silent --accept-package-agreements --accept-source-agreements
```

(`Add-AppxPackage` and `winget` are PowerShell tools; the Command Prompt has
no `winget` - keep these four lines in the blue window.)

Then VERIFY the two commands are available. In EITHER prompt (Command
Prompt or PowerShell) - but in a NEW window, because the PATH entries only
appear in windows opened after the install:

```bat
py -3 --version
git --version
```

### macOS

All macOS commands run in the **Terminal** app (open it: Spotlight, type
`Terminal`, press Enter).

Option 1 - the built-in path: **Xcode command line tools** provide both
`git` and a usable `python3`:

```bash
xcode-select --install          # opens a dialog; wait for the install
python3 --version && git --version
```

Option 2 - [Homebrew](https://brew.sh) (recommended if you use it for
anything else):

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
brew install python3 git
```

## Get the code and run the installer

Once the dependencies exist, the install is the same shape on every OS:
make a folder you like, clone, run `install.py`. The installer is
interactive (a menu - details in [docs/installer.md](docs/installer.md));
flag-driven variants like `--install` never pause at the end.

### Get the code and run the installer - Windows

Pick **one** of the two prompts (cmd or PowerShell) and run the whole block:

cmd:

```bat
mkdir "%USERPROFILE%\Tools" 2>nul
cd /d "%USERPROFILE%\Tools"
git clone https://github.com/corelan/officekick.git
cd officekick
py -3 install.py
```

PowerShell:

```powershell
mkdir "$env:USERPROFILE\Tools" -Force
cd "$env:USERPROFILE\Tools"
git clone https://github.com/corelan/officekick.git
cd officekick
py -3 install.py
```

(`cmd` uses `%USERPROFILE%` and needs `cd /d`; PowerShell uses
`$env:USERPROFILE`. A download-zip instead of the clone works the same way:
unpack it into the Tools folder and `cd` into the unpacked folder - but you
lose `git pull` updates, which is a real loss.)

**Quit every Office app first** (Word/PowerPoint/Outlook) - the installer
refuses to touch the add-in registration while they run. The installer walks
six phases (prepare → deploy → certificates → register → autostart → start
the server): the certificate phase shows a consent screen and may open ONE
more OS dialog asking you to ACCEPT the locally generated certificate -
that is expected, accept it (why:
[docs/certificates.md](docs/certificates.md)).

After the run, activate the add-ins in the Office apps -
[the next section](#activate-the-add-ins).

### Get the code and run the installer - macOS

```bash
mkdir -p "$HOME/Tools" && cd "$HOME/Tools"
git clone https://github.com/corelan/officekick.git
cd officekick
python3 install.py
```

macOS specifics: locally hosted models (Ollama on your LAN) need one
consent - see [macOS permissions](#macos-permissions-locally-hosted-models).
MailKick needs the https prerequisites met before the installer offers it
([detail](docs/mailkick.md#requirements-for-outlook-macos---read-before-installing)).

### Update / uninstall

```bash
# macOS, in Terminal:
python3 install.py --update
```

```bat
rem Windows - Command Prompt or PowerShell both work:
py -3 install.py --update
```

`--update` = git pull + redeploy + restart; the update runs to completion
even while your Office apps are open - the files are copied and the server
restarts on the new build in any case, and only the Office-touching
re-registration steps (the WEF manifest purge + sideload on macOS, the cache
flush) are skipped with a note in the output: **quit and reopen each Office
app afterwards** so it loads the new pane files (a later install with all
Office apps quit re-runs the skipped re-registration if you ever need it).
With git installed, the **pane's check-for-updates button does the whole
flow in one click** ("Update now - pull + redeploy + restart"); it asks you
to reload the pane when it finishes. Heal an
"add-in is no longer available" toast once with the repair tool
([docs/slidekick.md](docs/slidekick.md), troubleshooting).

```bash
# macOS, in Terminal:
python3 install.py --uninstall                                   # keeps your config
python3 install.py --uninstall --remove-certs --purge-config     # also removes the certificates and moves the config folder aside to a backup
```

```bat
rem Windows - Command Prompt or PowerShell:
py -3 install.py --uninstall
```

Details: [docs/installer.md](docs/installer.md#uninstall).

## Activate the add-ins

One rule comes first, on every platform:

> **Quit and reopen the Office app after every install or update.** The
> registrations (the registry dev-sideload on Windows, the WEF folder on
> macOS) are read at app LAUNCH - an app that was running during the install
> does not see the add-ins until it restarts. The pane's About tab shows the
> running **Pane build** version; use it to confirm a reload took effect.

Once registered, the add-in is permanently in the list - what you do per
document is **open its pane**:

> **Every NEW document starts without the pane.** Office only shows the pane
> after you open the add-in (the list entry / the ribbon button). Click it
> once per document and the task pane opens beside the content.
>
> **A document REMEMBERS the pane once it has been saved with it open.** The
> open pane is part of the saved file: reopening a saved document brings the
> pane back by itself - no clicking again. So the rhythm after the install
> is: new document -> one click on the add-in -> the pane is there for as
> long as you work -> save -> it comes back next time.

Once the pane is open, set it up once (the next two sections walk through
the activation steps per app, then the one-time configuration).

### Word

**Windows**

1. Open Word and look at the **Home tab → Add-ins** (the dropdown next to
   the button). The install's per-add-in launch-time sideload normally puts
   *Corelan DocKick* there already - if it is listed, you are done.
2. If it is not: use the classic dialog route. `Insertion / Mes compléments`
   → the classic **Office Add-ins** dialog (use the "Advanced" link next to
   Manage - NOT the modern "Apps" browse view, which never lists folder
   catalogs on any build) → tab **SHARED FOLDER / DOSSIERS PARTAGÉS** at the
   top → **Corelan DocKick** → **Add** (once per application; it stays in
   the list afterwards).
3. The Shared Folder tab lists what the Trust Center catalog enumerates.
   The installer pre-registers that catalog in the registry
   (`HKCU\...\WEF\TrustedCatalogs` with `Url` + `flag` = 1). If the tab is
   still empty, add the catalog through the Trust Center once - and on the
   newest Windows 11 Word builds **the catalog UI rejects local drive
   paths** ("Avez-vous inclus "https://"?" - it wants a URL/network form):
   enter the loopback **UNC** form instead
   (the admin share `C$` is on by default on clean installs):

   ```
   \\localhost\c$\Users\<user>\AppData\Local\OfficeKick\manifest
   ```

   Add it through the Word options dialog, in the labels YOUR Word shows
   (the UI language follows the Office install):

   - English Word: File → Options → **Trust Center** → **Trust Center
     Settings…** → **Trusted Add-in Catalogs** → paste that path →
     **Add catalog** → tick **"Show in Menu"** → OK → restart Word.
   - French Word: Fichier → Options → **Centre de gestion de la
     confidentialité** → **Paramètres du Centre de gestion de la
     confidentialité…** → **Catalogues de compléments approuvés** → paste
     that path → **Ajouter un catalogue** → tick **"Afficher dans le
     menu"** → OK → restart Word.

   The full diagnosis order (registry checks first):
   [docs/setup-windows.md > If the add-in does not appear in Word](docs/setup-windows.md#if-the-add-in-does-not-appear-in-word-windows).

**macOS**

1. Open a document.
2. Home tab → **Add-ins** (dropdown) → **Developer Add-ins** →
   **Corelan DocKick**. The pane appears; the icon stays in the dropdown
   every session.

### PowerPoint

**Windows**

1. Open PowerPoint → **Home tab → Add-ins** (dropdown) - the launch-time
   sideload normally lists *Corelan SlideKick* already.
2. Otherwise: Home → Add-ins → **Advanced** (the link next to Manage) → the
   classic **Office Add-ins** dialog → **SHARED FOLDER** at the top →
   **Corelan SlideKick** → **Add**. (The Trust Center catalog route is the
   same one described under [Word (Windows)](#word) above - including the
   UNC-path quirk on the newest Windows 11 builds.)

**macOS**

1. Open a deck.
2. Home tab → **Add-ins** (dropdown) → **Developer Add-ins** →
   **Corelan SlideKick**.
   (16.0.19725+ builds also gain a ribbon button from a commands manifest;
   [docs/slidekick.md](docs/slidekick.md) has the caveat for builds that
   hide ribbon-commands add-ins.)

### Outlook

**MailKick requires https everywhere** - on every platform, in every
Outlook mode (New or classic). Outlook refuses http add-in URLs, so the
installer stages MailKick **only when the https side is ready**: a working
localhost certificate. The **self-signed leaf is the default on every
platform**; the installer asks for your consent and creates/trusts it, and
the loopback CA or bring-your-own stay selectable alternatives
([docs/certificates.md](docs/certificates.md), the macOS requirements in
[docs/mailkick.md](docs/mailkick.md#requirements-for-outlook-macos---read-before-installing)).
Word/PowerPoint fall back to plain http when no certificate exists; ONLY
Outlook does not.

**Windows**

1. The installer stages MailKick only when the https prerequisites are met
   (`cryptography` present - the installer auto-installs it - and your
   trusted certificate). If the picker listed Outlook as blocked, re-run
   `py -3 install.py` after reading the printed reason.
2. Quit and reopen Outlook. Where the add-in appears, by Outlook mode:

   - **Classic Outlook** (the traditional desktop app): the dev-sideload
     registration makes MailKick a button on the item ribbons (reading a
     message / composing); the add-in management dialog is
     **Home > All Apps > Get Add-ins**.
   - **New Outlook** (the default on a clean Windows 11 install; web-based):
     the desktop WEF sideload does not reach this mode. Install MailKick's
     manifest ONCE as a custom add-in, from your BROWSER (Microsoft's
     documented route - custom XML sideloads go through the web dialog, not
     the new in-app Store):
     1. Sign in to outlook.live.com (the same account the desktop New
        Outlook uses) and open `https://aka.ms/olksideload` -
        the **Add-Ins for Outlook** window opens.
     2. **My add-ins** → **Custom Addins** → **Add a custom add-in** →
        **Add from file** → pick
        `%LOCALAPPDATA%\OfficeKick\manifest\mailkick.xml` → **Install**.
     3. Custom add-ins carry across your Outlook versions and devices, so
        the desktop New Outlook picks MailKick up after a restart.
     4. Find it while using Outlook: on the **Home** ribbon / left app bar
        it hides behind **More apps → Add apps** (that page is the STORE -
        MailKick is local, so it only shows after the step above); while
        reading or composing a message, installed add-ins open through the
        **Apps** button on the message's action bar or the add-in's own
        ribbon button. The classic **Add from URL** option is retired:
        download/copy the XML and use **Add from file**.

**macOS**

1. Quit and reopen Outlook; the manifest lands in Outlook's own add-in
   folder (`~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/`).
2. If MailKick does not appear: **Get Add-ins → My add-ins → Custom
   Add-ins → Add a custom add-in → Add from File (Cmd+Shift+G)** and pick
   the staged `mailkick.xml`.
3. Outlook works in New and classic; the one-time loopback-TLS setup and
   its checks live in
   [docs/mailkick.md](docs/mailkick.md#requirements-for-outlook-macos---read-before-installing).

### First run (getting started)

The pane is open - set it up once, in four steps:

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

That is it - no repeated setup. The provider/profile choices apply to every
Office app on the machine (one shared local server). The two things to
remember about the pane itself: it is **per document** (open the add-in once
per new document), and it **comes back on its own** for documents you save
with the pane open ([Activate the add-ins](#activate-the-add-ins)).

## Using it

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
# macOS, in Terminal:
python3 install.py --start
```

```bat
rem Windows - Command Prompt or PowerShell:
py -3 install.py --start
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
- Windows has no such gate.

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

- [docs/setup-windows.md](docs/setup-windows.md) - the Windows **dependency**
  script (`tools/setup-windows.ps1`: the Python Install Manager + Git; no
  clone, no installer run) + the execution-policy how-to + the Windows
  add-in troubleshooting order.
- [docs/installer.md](docs/installer.md) - the installer menus, transport
  options, certificate modes, uninstall (asks + defaults, flags).
- [docs/certificates.md](docs/certificates.md) - how the local https
  certificates are generated, trusted, renewed, and removed, per platform.
- [docs/settings.md](docs/settings.md) - every setting explained: what it
  means, how to configure it (connections, profiles, approvals, search
  steps, the Style tab: custom style prompt + the image style).
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
