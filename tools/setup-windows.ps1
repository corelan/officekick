# OfficeKick - Windows dependency installer (PowerShell)
#
# Installs the prerequisites for install.py, then (optionally) clones the
# OfficeKick repository into %USERPROFILE%\Tools.
#
#   - Python 3 (latest available through winget: 3.x, user scope, no admin)
#   - Git for Windows (latest release)
#
# HOW TO RUN:
#   from the internet (no download needed):
#     powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/corelan/officekick/main/tools/setup-windows.ps1 | iex"
#   from a saved file:
#     powershell -NoProfile -ExecutionPolicy Bypass -File .\setup-windows.ps1
#   (see docs/setup-windows.md for the execution-policy explanation and
#    how to pass -flags through the pipe)
#
# OPTIONS:
#   -NoClone                skip the repository clone offer
#   -Yes                    answer YES to interactive prompts (unattended)
#   -ToolsDir <dir>         tools folder (default: %USERPROFILE%\Tools)
#   -CloneUrl <url>         repository URL (default: officekick on GitHub)
#   -SkipPython / -SkipGit  do not attempt that install
#
# Everything installs PER-USER where possible - no administrator rights
# are needed. The script changes nothing outside the tools folder, the
# package managers' own install locations, and the PATH.

#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$NoClone,
    [switch]$Yes,
    [switch]$SkipPython,
    [switch]$SkipGit,
    [string]$ToolsDir,
    [string]$CloneUrl = "https://github.com/corelan/officekick.git"
)

$ErrorActionPreference = "Stop"
$BOOTSTRAP_VERSION = "0.9.272"

if (-not $ToolsDir) {
    $home0 = $env:USERPROFILE
    if (-not $home0) { $home0 = $env:HOME }
    $ToolsDir = Join-Path $home0 "Tools"
}
if ([string]::IsNullOrWhiteSpace($ToolsDir)) {
    Fail "the tools folder resolved EMPTY - pass it explicitly: -ToolsDir %USERPROFILE%\Tools"
}

function Info([string]$m) { Write-Host "[info] $m" }
Info "OfficeKick Windows bootstrap v$BOOTSTRAP_VERSION"
Info "tools folder: $ToolsDir"

function Step([string]$m) { Write-Host "[..]   $m" }
function Ok([string]$m)   { Write-Host "[ok]   $m" }
function Warn([string]$m) { Write-Host "[warn] $m" }
function Fail([string]$m) { Write-Host "[fail] $m" -ForegroundColor Red; exit 1 }

function Have-Command([string]$name) {
    return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

# make freshly installed tools visible to THIS session (registry-derived
# PATH for per-user Python installs when the installer did not refresh it)
function Refresh-SessionPath {
    $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $user    = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = ($machine + ";" + $user + ";" + $env:Path)
}

function Get-PythonVersion {
    Refresh-SessionPath
    foreach ($pair in @(@("py", "-3"), @("python"))) {
        $exe = $pair[0]
        $arg = $pair[1]
        try {
            $out = if ($arg) { & $exe $arg --version 2>&1 } else { & $exe --version 2>&1 }
        } catch { continue }
        $s = ($out | Out-String).Trim()
        if ($s -match "^Python 3\.(\d+)\.") { return $s; }
    }
    return $null
}

function Confirm([string]$question) {
    if ($Yes) { return $true; }
    $a = Read-Host "$question [y/N]"
    return ($a -match "^[Yy]")
}

# ----------------------------------------------------------------------
# Python 3 - via the official PYTHON INSTALL MANAGER from python.org
# (never winget for Python). The manager MSIX installs per user; the
# latest Python runtime comes from `py install default` (the manager's
# tag for the latest RECOMMENDED release; bundled runtimes may already
# satisfy it). FIELD-VERIFIED: `py install latest` is not a tag in current
# manager builds - it errors with 'Specify at least one tag to install,
# or default'.
# ----------------------------------------------------------------------
if (-not $SkipPython) {
    $have = Get-PythonVersion
    if ($have) {
        Ok "Python found: $have (nothing to install)"
    } else {
        $installed = $false
        if (Have-Command "py") {
            Step "installing the latest Python via the Python install manager..."
            py install default
            if ($LASTEXITCODE -eq 0) { $installed = $true }
            else { Warn "the existing py command could not install runtimes (it may be the classic launcher) - installing the manager from python.org" }
        }
        if (-not $installed) {
            Step "installing the Python install manager (python.org official MSIX, per user)..."
            try {
                Add-AppxPackage -AppInstallerFile "https://www.python.org/ftp/python/pymanager/pymanager.appinstaller"
            } catch {
                Warn "the appinstaller route failed: $($_.Exception.Message)"
                Step "downloading the manager MSIX directly from the python.org index..."
                try {
                    $ftp = (Invoke-WebRequest -UseBasicParsing -TimeoutSec 30 `
                        "https://www.python.org/ftp/python/pymanager/").Content
                    if ($ftp -isnot [string]) { $ftp = [string]($ftp -join "`n") }
                    $cands = @()
                    foreach ($m in [regex]::Matches($ftp, "python-manager-(\d+)\.(\d+)\.msix")) {
                        $cands += [pscustomobject]@{
                            ver = [version]($m.Groups[1].Value + "." + $m.Groups[2].Value)
                            url = "https://www.python.org/ftp/python/pymanager/" + $m.Value
                        }
                    }
                    $latest = $cands | Sort-Object ver -Descending | Select-Object -First 1
                    if (-not $latest) { Fail "could not resolve a Python install manager MSIX on python.org" }
                    Info "downloading $($latest.url)"
                    $tmpM = Join-Path $env:TEMP "python-manager.msix"
                    Invoke-WebRequest -UseBasicParsing -Uri $latest.url -OutFile $tmpM -TimeoutSec 900
                    Add-AppxPackage -Path $tmpM
                    Remove-Item $tmpM -ErrorAction SilentlyContinue
                } catch {
                    Fail "could not install the Python install manager - install it manually from https://www.python.org/downloads/ (the Python install manager section) and re-run this script"
                }
            }
            Refresh-SessionPath
            Step "installing the latest Python via the install manager..."
            py install default
            if ($LASTEXITCODE -ne 0) {
                Step "trying the explicit major-version tag..."
                py install 3
                if ($LASTEXITCODE -ne 0) {
                    Fail "py install default failed - list what the manager offers with: py list --online   (then: py install <tag>), and re-run this script"
                }
            }
        }
        Refresh-SessionPath
        $have = Get-PythonVersion
        if ($have) { Ok "Python installed: $have" }
        else { Fail "Python is not available yet - open a NEW terminal (PATH refresh) and re-run this script" }
    }
} else {
    Info "skipping Python (-SkipPython)"
}

# ----------------------------------------------------------------------
# Git for Windows
# ----------------------------------------------------------------------
if (-not $SkipGit) {
    Refresh-SessionPath
    $haveGit = $false
    if (Have-Command "git") {
        $gv = (& git --version) | Out-String
        Ok "Git found: $($gv.Trim()) (nothing to install)"
        $haveGit = $true
    }
    if (-not $haveGit) {
        if (Have-Command "winget") {
            Step "installing Git for Windows via winget..."
            winget install -e --id Git.Git --silent `
                --accept-package-agreements --accept-source-agreements
            if ($LASTEXITCODE -ne 0) { Warn "winget Git install failed - falling back" }
        }
        Refresh-SessionPath
        if (-not (Have-Command "git")) {
            Step "fetching the latest Git for Windows release..."
            $rel = Invoke-RestMethod -UseBasicParsing -TimeoutSec 30 `
                "https://api.github.com/repos/git-for-windows/git/releases/latest"
            $asset = $rel.assets | Where-Object { $_.name -match "^Git-.*-64-bit\.exe$" } | Select-Object -First 1
            if (-not $asset) { Fail "could not resolve a Git installer asset - install Git from https://git-scm.com/download/win and re-run" }
            Info "downloading $($asset.name)"
            $tmpGit = Join-Path $env:TEMP "git-setup.exe"
            Invoke-WebRequest -UseBasicParsing -Uri $asset.browser_download_url -OutFile $tmpGit -TimeoutSec 900
            Step "running the Git installer (silent, per-user defaults)..."
            Start-Process -FilePath $tmpGit -ArgumentList "/VERYSILENT", "/NORESTART", `
                "/CURRENTUSER" -Wait
            Remove-Item $tmpGit -ErrorAction SilentlyContinue
            Refresh-SessionPath
        }
        if (Have-Command "git") {
            $gv = (& git --version) | Out-String
            Ok "Git installed: $($gv.Trim())"
        } else {
            Fail "Git did not install - install it from https://git-scm.com/download/win and re-run this script"
        }
    }
} else {
    Info "skipping Git (-SkipGit)"
}

# ----------------------------------------------------------------------
# The OfficeKick repository
# ----------------------------------------------------------------------
Refresh-SessionPath
$repoDir = Join-Path $ToolsDir "officekick"
$offerClone = -not $NoClone
if ($offerClone -and -not $Yes) {
    $offerClone = Confirm ("Clone the OfficeKick repository into [" + $ToolsDir + "]?")
}
if ($offerClone -and [string]::IsNullOrWhiteSpace($ToolsDir)) {
    Fail "internal: the tools folder is empty - pass -ToolsDir <dir> and re-run"
}
if ($offerClone) {
    if (-not (Test-Path $ToolsDir)) {
        Step "creating $ToolsDir"
        New-Item -ItemType Directory -Path $ToolsDir -Force | Out-Null
    }
    if (Test-Path $repoDir) {
        Step "$repoDir exists - updating (git pull --ff-only)..."
        Push-Location $repoDir
        try {
            git -c http.sslBackend=schannel pull --ff-only
            if ($LASTEXITCODE -eq 0) { Ok "updated" }
            else { Warn "git pull failed (exit $LASTEXITCODE) - the existing clone is kept as-is" }
        } finally { Pop-Location }
    } else {
        Step "cloning $CloneUrl into $repoDir ..."
        git clone $CloneUrl $repoDir
        if ($LASTEXITCODE -eq 0) { Ok "cloned" }
        else { Fail "the clone failed - check the URL/credentials, then re-run with -NoClone" }
    }
} else {
    Info "skipping the repository clone"
}

# ----------------------------------------------------------------------
# Next steps
# ----------------------------------------------------------------------
Info ""
Info "Next steps:"
if (Test-Path $repoDir) {
    Info "  cd $repoDir"
    Info "  py -3 install.py            (menu: 1 Install / repair)"
    Info "  documentation:  docs\installer.md, docs\setup-windows.md"
} elseif (Have-Command "git") {
    Info "  clone the repo, then:  py -3 install.py"
} else {
    Info "  install Python 3 + Git, then clone the repository"
}
Ok "done."
