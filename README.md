# OfficeKick

AI assistant add-ins for **Microsoft Office**, from Corelan. They run on macOS
and Windows and share one technology stack.

- **Corelan SlideKick** - the PowerPoint add-in ([docs/slidekick.md](docs/slidekick.md))
- **Corelan DocKick** - the Word add-in ([docs/dockick.md](docs/dockick.md))
- **Corelan MailKick** - the Outlook add-in: summarize, translate, draft replies (you send), categories ([docs/mailkick.md](docs/mailkick.md))

## How it works

1. You type what you want, in plain language.
   Examples: `reword the current slide`, `make slides about threat hunting`,
   `add a table of the key findings`, `leave a review comment on this paragraph`.
2. The add-in reads your file.
3. The request goes to an AI model that **you** configured.
4. The add-in shows the proposed changes.
5. Your approval policy decides what gets applied.

Everything runs on your machine:

- One small local server serves the add-ins and the AI requests.
- You configure providers once; every Office app shares them.
- Supported providers: OpenAI, Anthropic, OpenRouter, any OpenAI-compatible
  server, Ollama, llama.cpp. Nothing goes through a third-party service.
- Every exchange is auto-saved in a resumable local session. The
  **Previous requests** box shows each run with its timestamp.
- MailKick adds one-click **reply+ / reply-** buttons for a drafted polite reply.

## Install

One installer for macOS and Windows. It detects the platform by itself.

```bash
git clone https://github.com/corelan/officekick.git
cd officekick
python3 install.py        # Windows: python install.py
```

The installer shows a menu:

1. **Install** (also repairs an existing install)
2. **Update** (pulls the repo with git, then updates)
3. **Uninstall**

What the installer does:

- Checks `python3` (3.8 or newer).
- Finds your Office apps.
- Copies the files to `~/Library/Application Support/OfficeKick` (macOS) or
  `%LOCALAPPDATA%\OfficeKick` (Windows).
- Registers the add-ins with each Office app (macOS: WEF sideload; Windows:
  a per-user trusted folder catalog. Outlook on Windows needs the one-time
  "Add from File" dialog - the installer prints the steps).
- Starts the local server on `http://localhost:8765` and health-checks it.
- Asks whether to start the server at login. Default is **yes**.
  - macOS: a login item ("OfficeKickServer" via SMAppService).
  - Windows: a per-user scheduled task (windowless `pythonw`).

Good to know:

- No admin rights are needed.
- The add-ins talk to the server over plain http. Office allows `localhost`
  over http, so there are no certificates on any path.
- The macOS server wrapper restarts the server if it crashes.
- Re-running the installer is safe. It repairs and updates in place.

## Update

```bash
python3 install.py --update    # Windows: python install.py --update
```

This pulls the repo with git and updates the install. Then quit and reopen
your Office apps so the task panes load the new files.

The pane also checks for updates itself. See **Settings -> Check now** (or the
"Check for updates automatically at startup" option). Your config, profiles and
sessions survive every update.

Config lives in `~/.officekick/config.json` (`%USERPROFILE%\.officekick` on
Windows, chmod 600). API keys never leave your machine except to the provider
you configured.

## Use it

- **PowerPoint (macOS)**: Home tab -> **Add-ins** (dropdown) ->
  **Developer Add-ins** -> **Corelan SlideKick**.
- **Word (macOS)**: same route -> **Corelan DocKick**.
- **PowerPoint / Word (Windows)**: Insert -> My Add-ins -> **SHARED FOLDER** ->
  the add-in -> Add (once).
- **Outlook (MailKick)**: the installer places the manifest in Outlook's own
  add-in folder. Quit and reopen Outlook. If MailKick does not appear:
  Get Add-ins -> My add-ins -> Custom Addins -> Add a custom add-in ->
  Add from File (press Cmd+Shift+G for the path). See
  [docs/mailkick.md](docs/mailkick.md).

First run:

1. **Connections** tab: add a provider (type + Base URL + optional API key).
   "Test connection" lists its models.
2. **Profiles** tab: pick the content model (edits and generation) and the
   other-actions model (advice and brainstorming).
3. Type an instruction in the pane and hit **Run** (Cmd/Ctrl+Enter).

### Ribbon (opt-in)

`python3 install.py --ribbon` keeps the manifest's AI commands: an **AI** group
with the add-in icon on the Home tab (MailKick: message read/compose tabs).

- Default is the list-mode manifest (no ribbon commands).
- Reason: our tested Office-Mac build accepts a command manifest but replaces
  the Add-ins-list entry with it and then renders neither the group nor the
  list item. That leaves the add-in unreachable.
- If your build shows the AI group, you can stay on `--ribbon`.

## macOS permissions (locally hosted models)

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

## Security & privacy

- The local server binds to loopback only (`127.0.0.1` + `::1`).
- No telemetry, no phone-home. The server never touches your documents.
- API keys stay in the local config file. They are attached to provider calls
  server-side only; the pane and the API never show or return them.
- File content leaves your machine only to the providers you configured.
- Debug logging is opt-in. It writes prompts and document text to a local
  daily log; keep it off for sensitive work.

## Uninstall

```bash
python3 install.py --uninstall            # keeps your config
python3 install.py --uninstall --purge-config   # also wipes config
```

Or pick **Uninstall** in the installer menu. Then quit and reopen your Office
apps. The installer keeps `~/.officekick` (API keys, profiles, sessions) unless
you pass `--purge-config` or answer yes on the config question.

## More

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