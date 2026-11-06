# Windows setup: dependencies, the repository, and PowerShell's execution policy

`tools/setup-windows.ps1` installs the prerequisites for `install.py` on
Windows, then offers to clone the OfficeKick repository into
`%USERPROFILE%\Tools\officekick`. It installs **per user** where possible -
no administrator rights are needed - and touches nothing besides the
package managers' install locations, the PATH, and the tools folder.

What it installs:

- **Python 3** - the latest release winget offers (currently 3.12+; the
  script tries the newest ids first), or the python.org amd64 installer
  silently when winget is not available. Skipped when a Python 3.x is
  already installed (`py -3` / `python`).
- **Git for Windows** - via winget (`Git.Git`), or the latest
  git-for-windows release asset from GitHub. Skipped when `git` exists.
- **The repo clone** - offered interactively (answer the `y/N` prompt);
  an existing clone is updated with `git pull --ff-only`.

Flags: `-Yes` (unattended), `-NoClone`, `-SkipPython`, `-SkipGit`,
`-ToolsDir <dir>` (default `%USERPROFILE%\Tools`), `-CloneUrl <url>`.

---

## Running the script

Open PowerShell (`Win+X` -> "Terminal" or "Windows PowerShell"), then:

```powershell
cd <where you saved the script>
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup-windows.ps1
```

That one-liner is the RECOMMENDED way: **Bypass applies to that process
only** and expires when the shell closes - nothing about your system's
security posture changes.

If you are already in a PowerShell window, the equivalent session-scoped
change is:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\setup-windows.ps1
```

Both forms need no administrator.

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

## After the script

```powershell
cd $env:USERPROFILE\Tools\officekick
py -3 install.py
```

The installer opens its menu (1 Install / repair -> the add-in picker ->
the certificate question; see [installer.md](installer.md)). With https
selected, the certificate trust consent screen explains the one OS dialog
(certificate modes: [certificates.md](certificates.md)).

Verify any time: `py -3 --version`, `git --version`, and after installing
`http://127.0.0.1:8765/api/health`.
