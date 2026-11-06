# Windows setup: dependencies, the repository, and PowerShell's execution policy

`tools/setup-windows.ps1` installs the prerequisites for `install.py` on
Windows, then offers to clone the OfficeKick repository into
`%USERPROFILE%\Tools\officekick`. It installs **per user** where possible -
no administrator rights are needed - and touches nothing besides the
package managers' install locations, the PATH, and the tools folder.

## Prerequisites in plain words

- Running the OfficeKick installer (`install.py`) needs **Python 3** on the
  machine - on Windows that means the `py -3` command works. Do NOT rely on
  `python`: on a clean Windows box it is often a Store alias that opens the
  shop instead of an interpreter.
- **Git** (the command-line utility) is required only if you want the
  repository cloned (`git clone ...` into `%USERPROFILE%\Tools`) - which is
  the recommended durable setup (updates become a `git pull`).
- Both are installed by the one script below: it fetches the **Python
  install manager** from python.org and runs `py install default` - the manager's tag for the latest RECOMMENDED release (winget is
  deliberately not used for Python), and installs Git for Windows (winget,
  with a github release fallback).

## Installing the dependencies with the script

**Option A - straight from the internet, no download** (one command in a
PowerShell window; the Linux `curl | bash` habit works here too - `irm`
downloads, `iex` runs):

```powershell
irm https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | iex
```

The literal-curl variant (Windows 10+ ships a genuine `curl.exe` in
System32; type `curl.exe`, not the `curl` PowerShell alias, which is a
different thing in Windows PowerShell 5.1):

```powershell
curl.exe -fsSL https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | powershell -NoProfile -ExecutionPolicy Bypass -Command -
```

**Option B - you have downloaded the script** (e.g. saved
`tools/setup-windows.ps1` from the repository or a zip):

```powershell
cd <where you saved it>
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup-windows.ps1
```

What the script does, in detail:

- **The Python install manager** - python.org's official tool, downloaded
  from python.org itself (never via winget) with PowerShell's
  `Add-AppxPackage -AppInstallerFile
  https://www.python.org/ftp/python/pymanager/pymanager.appinstaller`
  (per user, no admin; a direct python.org FTP-index MSIX download is the
  fallback), then `py install default` brings the newest Python runtime (the manager errors on `py install latest` - 'latest' is not a tag; use `default`).
  Skipped when a Python 3.x is already installed (`py -3` / `python`).
- **Git for Windows** - via winget (`Git.Git`; winget is used for Git only,
  not Python), or the latest git-for-windows release asset from GitHub.
  Skipped when `git` exists.
- **The repo clone** - offered interactively (answer the `y/N` prompt);
  an existing clone is updated with `git pull --ff-only`.
- Flags: `-Yes` (unattended), `-NoClone`, `-SkipPython`, `-SkipGit`,
  `-ToolsDir <dir>` (default `%USERPROFILE%\Tools`), `-CloneUrl <url>`.

Things worth knowing:

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

- Passing flags through a pipe does not work directly (`iex` cannot take
  the script's parameters) - use the script-block form...:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1))) -Yes -NoClone"
```

- **If `irm` "did not work": check which prompt you are in first.**
  `irm` and `iex` are PowerShell commands - in a CMD window they fail with
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

- If you are already inside a PowerShell window, the session-scoped
  equivalent of the `-ExecutionPolicy Bypass` flag is:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\setup-windows.ps1
```

## The execution policy, properly explained

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

## If the add-in does not appear in Word (Windows)

The installer registers two layers; check them in this order:

1. **Quit and reopen the app.** The per-add-in launch-time sideload
   (registry: `HKCU\...\Wef\Developer`) is read at every app LAUNCH - a
   Word instance that was running during the install knows nothing of the
   add-ins until it restarts. Verify the layer first:

   ```bat
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

## After the script - the folder and the clone, manually

The script's clone step is a convenience; you never have to rely on it. The
folder path uses **different syntax per prompt kind**:

- In a **command prompt (cmd)** - `%USERPROFILE%` and `cd /d`:

```bat
mkdir "%USERPROFILE%\Tools" 2>nul
cd /d "%USERPROFILE%\Tools"
git clone https://github.com/corelan/officekick.git
cd officekick
py -3 install.py
```

- In **PowerShell** - `$env:USERPROFILE`:

```powershell
cd "$env:USERPROFILE\Tools\officekick"      # if the script cloned it
# or manually:
mkdir "$env:USERPROFILE\Tools" -Force
cd "$env:USERPROFILE\Tools"
git clone https://github.com/corelan/officekick.git
cd officekick
py -3 install.py
```

(`cd $env:USERPROFILE\...` in a CMD window does nothing useful - cmd does
not expand `$env:` variables. That is why both forms are spelled out.)

The installer opens its menu (1 Install / repair -> the add-in picker ->
the certificate question; see [installer.md](installer.md)). With https
selected, the certificate trust consent screen explains the one OS dialog
(certificate modes: [certificates.md](certificates.md)).

Verify any time: `py -3 --version`, `git --version`, and after installing
`http://127.0.0.1:8765/api/health`.
