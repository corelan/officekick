# The installer (install.py)

One installer for macOS and Windows. Run it from your OfficeKick checkout:

```
python3 install.py        # macOS
py -3 install.py          # Windows (not "python" - that is often the Store alias)
```

Full tool chatter lands in `install.log` next to install.py; the console
shows the step lines only. No admin rights are needed anywhere. Re-running is
safe - it repairs and updates in place.

## The main menu

```
 1) Install / repair        (Word + PowerPoint on; Outlook off by default)
 2) Update                  (git pull + refresh; keeps your add-ins)
 3) Server                  (start / restart / stop / status)
    ---
 4) Transport options       (transport: <mode> | ports: ... | certs: <mode>)
    ---
 5) Uninstall               (removes the add-ins + this server)
    ---
 0) exit
```

Arrow keys (or the digit keys) navigate; the Transport row recomputes its
values on every redraw. Every action exists as a CLI flag too: `--install`,
`--update`, `--start`/`--restart`, `--status`, `--stop`, `--uninstall`,
`--remove-certs`, `--purge-config`, `--transport <auto|http|https>`,
`--cert-check`, `--mailkick`, `--dry-run`. Automation runs never prompt -
they take the stored defaults.

## Install / repair (1)

The add-in picker selects what gets installed. Word + PowerPoint are on by
default; Outlook is off until you toggle it.

- The picker shows the **https prerequisites** for Outlook (mode-aware: on
  Windows they matter only when the https transport is selected). Outlook is
  locked while the prerequisites are unmet - the missing pieces are named
  (e.g. the `cryptography` package on Windows).
- Enabling Outlook explains that the installer asks for the certificate mode
  **before the certificates step**, and that the self-signed leaf is the
  default answer.

Phases (questions are asked *before* the phase line that acts on them):

1. **preparing** - frees the ports (the stored http port and the https port;
   a foreign listener on the https port aborts with instructions).
2. **deploying** - copies the files (macOS: `~/Library/Application
   Support/OfficeKick`; Windows: `%LOCALAPPDATA%\OfficeKick`) and registers
   the add-ins (WEF sideload / trusted folder catalog). If an installed
   add-in does not appear in Word/PowerPoint, work through the checks in
   [setup-windows.md](setup-windows.md#if-the-add-in-does-not-appear-in-word-windows)
   (relaunch first - the sideload is read at app launch - then the Trust
   Center catalog, where the newest Windows 11 Word builds require the
   loopback UNC form for folder paths).
3. **certificates** - acts on the answer from the certificate question
   (see [certificates.md](certificates.md)). The consent screen explains
   up front that after you answer yes, YOUR SYSTEM opens one more dialog
   asking you to ACCEPT the leaf certificate; the installer prints a
   heads-up again right before that dialog appears (macOS: keychain
   dialog, click Trust; Windows: the Root-store certificate dialog,
   click Yes).
4. **registering with Office** - re-renders the manifests. When Office apps
   are running, the WEF re-registration (manifest purge + per-app sideload)
   is skipped with an explicit note - the files are copied and the manifests
   rendered in any case; quit + reopen each app to pick up the new build, and
   a later install with all Office apps quit re-runs the skipped
   re-registration. With no Office app running, nothing changes.
5. **autostart** - asks before doing anything (macOS login item /
   Windows scheduled task).
6. **starting the server** - starts + health-checks both listeners.

## Transport options (4)

Controls how the panes reach the local server:

- **1 auto** - https when a working certificate exists, else plain http.
- **2 plain http** - certificates ignored; Outlook cannot use this.
- **3 https always** - a local certificate is created/renewed (leaf, CA, or
  bring-your-own).
- **4/5** - the http and https listener ports (validated, must differ).
- **6 certificate mode** - the [certificate mode
  submenu](certificates.md): self-signed leaf (default, no CA) / loopback CA
  / bring-your-own, the current mode marked. Switching stores the mode; with
  a working pair it offers an immediate conversion behind a consent screen
  (the server restarts on the new certificate).

The menu's cursor starts on the current mode; a bare Enter re-selects it.
When **nothing is installed yet**, switching the transport only *stores* the
mode - no deploy, no WEF staging, no server start until an install. When the
server is installed, the change renders the manifests and restarts the
server immediately.

Settings persistence: the config file (`~/.officekick/config.json`) is the
single source of truth for these settings - the setters write the file first
and mirror to a running server. On any divergence the file wins (the server
re-reads at its next restart).

## Update (2)

`git pull` + an in-place refresh. The update completes even with Office apps
open (files copy + the server restarts); the WEF re-registration and cache
flush are the only steps skipped while apps run - quit and reopen each Office
app afterwards to load the new build. Your config,
profiles, sessions and the selected certificate mode survive every update.

## Server (3)

start / restart / stop / status for the local server only - the same actions
as `--start` / `--restart` / `--stop` / `--status`. The status line shows the
version, both listeners, and the installed add-ins.

## Uninstall (5)

Asks two questions, **both defaulting to KEEP** - an uninstall never
destroys your setup by default:

1. **Certificates**: removes the local certificates OfficeKick installed
   (trust entry - macOS keychain / Windows Trusted People + Root fallback -
   plus the files in `~/.officekick/certs`). A bring-your-own certificate's
   own trust entry is never touched; only the file copies are deleted.
2. **Config folder**: wipes `~/.officekick` (config.json, API keys,
   profiles, sessions). Keeping it means the next install reuses everything.

After an uninstall the certificate mode resets to **self-signed leaf**, the
add-ins and server are removed, and the Office apps should be quit and
reopened to drop the pane registrations.

Flags: `--uninstall [--remove-certs] [--purge-config]` for non-interactive
runs.

## Self-healing and guards

- A missing deploy folder is re-deployed before a server start.
- The https port is probed in the preparing phase; a foreign listener aborts
  with instructions (free it or change the port in Transport options).
- Office running = the manifest/cache parts are skipped with a note (they
  re-run on demand); server-only deploys and restarts stay available.
- Dry-runs (`--dry-run`) change nothing and answer every question with the
  stored defaults.
