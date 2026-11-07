# Windows setup - dependencies (Python 3 + Git)

This page is **step 1 of the Windows install**: putting the prerequisites on
the machine. OfficeKick's add-ins are installed later by `install.py` from a
git clone - the README's **"Get the code and run the installer"** section
carries those steps, and this script **neither clones nor runs the
installer**.

## What you need

- **Python 3** - the installer runs as `py -3`. Do not rely on plain
  `python`: on a clean Windows box it is often a Microsoft Store alias that
  opens the shop instead of an interpreter.
- **Git** - gets you the code and keeps it current (`install.py --update`
  pulls, redeploys and restarts). Git is technically optional but strongly
  recommended: without it, updates re-download a zip instead of a `git pull`.

Everything installs **per user** - no administrator rights needed anywhere.

## Installing the dependencies with the script (recommended)

`tools/setup-windows.ps1` installs both:

- **The Python install manager** - python.org's official tool, downloaded
  from python.org itself (never via winget) with `Add-AppxPackage
  -AppInstallerFile
  https://www.python.org/ftp/python/pymanager/pymanager.appinstaller`
  (per user, no admin; a direct python.org FTP-index MSIX download is the
  fallback), then `py install default` brings the newest Python runtime (the
  manager errors on `py install latest` - `latest` is not a tag; use
  `default`). Skipped when a Python 3.x is already present (`py -3`/`python`).
- **Git for Windows** - via winget (`Git.Git`; winget is used for Git only,
  not Python), or the latest git-for-windows release asset from GitHub.
  Skipped when `git` exists.

The script installs these two things and nothing else - no repo clone, no
installer run.

**Option A - straight from the internet, no download.** One command, and it
MUST be typed in a POWERSHELL window (the blue one; `irm` and `iex` do not
exist in the black Command Prompt window - there they fail with "'irm' is
not recognized"). Open PowerShell first: Start menu, type `powershell`,
Enter. `irm` downloads the script, `iex` runs it:

```powershell
irm https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | iex
```

The literal-curl variant - THE COMMAND PROMPT (cmd) form (Windows 10+ ships
a genuine `curl.exe` in System32, so the download-and-run works there
without touching the execution policy; in PowerShell always type
`curl.exe` - plain `curl` is a DIFFERENT PowerShell alias, not this tool,
on Windows PowerShell 5.1):

```powershell
curl.exe -fsSL https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | powershell -NoProfile -ExecutionPolicy Bypass -Command -
```

**Option B - you have downloaded the script** (e.g. saved
`tools/setup-windows.ps1` from the repository or a zip) - the long form
below works in EITHER window (cmd or PowerShell), because it launches its
own PowerShell:

```powershell
cd <where you saved it>
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup-windows.ps1
```

Flags: `-SkipPython`, `-SkipGit`. Flags through a pipe need the scriptblock
form (a plain pipe cannot carry parameters) - run in POWERSHELL:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1))) -SkipGit"
```

## Manual install (no script)

All of the next two blocks run in a POWERSHELL window (`Add-AppxPackage`
and `winget` are PowerShell tools; the Command Prompt has no `winget`).

Python (the official install manager, then the runtime):

```powershell
Add-AppxPackage -AppInstallerFile https://www.python.org/ftp/python/pymanager/pymanager.appinstaller
py install default
```

Git (either route works):

```powershell
winget install -e --id Git.Git --silent --accept-package-agreements --accept-source-agreements
```

or the 64-bit installer from
[git-scm.com/download/win](https://git-scm.com/download/win), installed
per-user.

Verify both - in EITHER window (Command Prompt or PowerShell), opened NEW
after the installs (the PATH entries only appear in new windows):

```bat
py -3 --version
git --version
```

## Things worth knowing about the one-liners

- No administrator prompt is needed anywhere in this flow (everything
  installs per user); an elevated terminal works too and adds nothing.
- The `irm | iex` form (and the curl pipe) runs the code without saving it
  to disk first - use it with URLs you trust (these point at the OfficeKick
  repository's main branch). To inspect before running, save the file:
  `irm <url> -OutFile setup-windows.ps1`, read it, then run the saved file
  as in option B.
- Older Windows PowerShell builds negotiate ancient TLS by default; if the
  one-liner fails to download, use the hardened form:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12'; irm https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | iex"
```

- **If `irm` "did not work": check which prompt you are in first.** `irm`
  and `iex` are PowerShell commands - in a CMD window they fail with
  "'irm' is not recognized". Either start PowerShell first, or use the
  `curl.exe` pipe form above (which is cmd-native). To make a failing
  one-liner SELF-DIAGNOSING, run this in PowerShell instead of piping on:

```powershell
try { irm https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | iex } catch { "IRM FAILED: $($_.Exception.Message)" }
```

  - "The underlying connection was closed / an unexpected error occurred on
    a send" = old TLS defaults -> use the hardened `Tls12` form above.
  - Name resolution / proxy errors = the machine blocks
    `raw.githubusercontent.com` -> use the downloaded-file route (option B).

## The PowerShell execution policy, properly explained

Windows PowerShell refuses to run `.ps1` script FILES unless the local
execution policy allows it (typing commands into the console is always
allowed; running saved scripts is what the policy gates). Three scopes
matter:

| Scope | Command | Persists | Needs admin |
|---|---|---|---|
| `Process` (recommended) | `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` | only that PowerShell window | no |
| `CurrentUser` | `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` | yes, for your user only | no |
| `LocalMachine` | `Set-ExecutionPolicy -Scope LocalMachine ...` | yes, machine-wide | yes |

- `Bypass` runs everything with no signature checks - fine for a one-off
  script you have read.
- `RemoteSigned` is the reasonable STANDING setting for a user: local
  scripts (like `setup-windows.ps1`) run freely, scripts downloaded from
  the internet must be signed. To set it persistently for your user:

  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
  ```

  (It shows a confirmation question; answer `Y`.)

- Scripts saved from a browser/chat carry a Windows "downloaded from the
  internet" mark (Zone.Identifier). Windows blocks those even under
  `RemoteSigned` until you unblock them:

  ```powershell
  Unblock-File .\setup-windows.ps1
  ```

  (or right-click the file -> Properties -> "Unblock" check box.)

Check the effective policy any time with `Get-ExecutionPolicy -List`.
The OfficeKick installer needs no PowerShell scripts at runtime - this
setup script is the only `.ps1` in the workflow - so `Process`-scope
Bypass for this one run is the smallest, cleanest choice.

## Step 2 is manual - the folder and the clone

The dependency script stops here; the clone is a five-second action you do
yourself (that keeps the script purely about the machine). The full
walkthrough lives in the README: **Get the code and run the installer**.
Short version, in cmd:

```bat
mkdir "%USERPROFILE%\Tools" 2>nul
cd /d "%USERPROFILE%\Tools"
git clone https://github.com/corelan/officekick.git
cd officekick
py -3 install.py
```

(PowerShell uses `$env:USERPROFILE` instead of `%USERPROFILE%`; `cd /d` is
cmd-specific. Verify any time: `py -3 --version`, `git --version`, and
after installing: `http://127.0.0.1:8765/api/health`.)

## If the add-in does not appear in Word (Windows)

The installer registers two layers; check them in this order:

1. **Quit and reopen the app.** The per-add-in launch-time sideload
   (registry: `HKCU\...\Wef\Developer`) is read at every app LAUNCH - a
   Word instance that was running during the install knows nothing of the
   add-ins until it restarts. Verify the layer first:

   ```bat
   rem Command Prompt (cmd):
   reg query "HKCU\Software\Microsoft\Office\16.0\Wef\Developer" /s
   ```

   Expected: `RefreshAddins` = 1 plus one entry per installed add-in, each
   pointing at `%LOCALAPPDATA%\OfficeKick\manifest\<addin>.xml`.

2. **The Trust Center catalog** (the classic dialog's Shared Folder /
   *Dossiers partages* tab) - field testing on a very new Windows 11 Word
   build found this quirk: the catalog UI **rejects local drive paths**
   with *"Desole... L'emplacement que vous avez entre est incorrect.
   Avez-vous inclus "https://"?"*. The working form is the loopback
   **UNC** path (the admin share C$ is on by default on clean Windows 11
   installs; no setup needed):

   ```
   \\localhost\c$\Users\<user>\AppData\Local\OfficeKick\manifest
   ```

   Word -> Fichier -> Options -> Centre de gestion de la confidentialite ->
   Parametres... -> **Catalogues de complements approuves** -> paste the
   UNC path -> **Ajouter un catalogue** -> tick **"Afficher dans le
   menu"** -> OK -> restart Word. The installer pre-registers the same
   catalog in the registry (`...\WEF\TrustedCatalogs` with `flag` = 1;
   older installs wrote the value under a name Office ignores - delete the
   `{C7B1F2A9-...}` subkey once or let the UI entry take over, both store
   the same shape).

3. The modern **Add-ins -> Manage "Apps"** dialog never lists local folder
   catalogs on any build - it is a dead end for this use case.
