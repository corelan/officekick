# OfficeKick - Windows dependency installer (PowerShell)
#
# Installs the prerequisites for install.py, then (optionally) clones the
# OfficeKick repository into %USERPROFILE%\Tools.
#
#   - Python 3 (latest available through winget: 3.x, user scope, no admin)
#   - Git for Windows (latest release)
#
# HOW TO RUN:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\setup-windows.ps1
#   (see docs/setup-windows.md for the execution-policy explanation)
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

if (-not $ToolsDir) {
    $home0 = $env:USERPROFILE
    if (-not $home0) { $home0 = $env:HOME }
    $ToolsDir = Join-Path $home0 "Tools"
}

function Info([string]$m) { Write-Host "[info] $m" }
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
# Python 3
# ----------------------------------------------------------------------
if (-not $SkipPython) {
    $have = Get-PythonVersion
    if ($have) {
        Ok "Python found: $have (nothing to install)"
    } else {
        $installed = $false
        if (Have-Command "winget") {
            foreach ($id in @("Python.Python.3.14", "Python.Python.3.13")) {
                Step "installing Python 3 via winget (id $id, user scope)..."
                winget install -e --id $id --scope user --silent `
                    --accept-package-agreements --accept-source-agreements
                if ($LASTEXITCODE -eq 0) { $installed = $true; break; }
                Warn "winget id $id unavailable - trying the next one"
            }
        } else {
            Warn "winget is not available - using the python.org direct installer"
        }
        if (-not $installed) {
            # fallback: python.org latest 3.x amd64 installer (per-user, PATH)
            Step "fetching the latest Python 3 installer from python.org..."
            $exe = $null
            try {
                $page = (Invoke-WebRequest -UseBasicParsing -TimeoutSec 30 `
                    "https://www.python.org/downloads/windows/").Content
                if ($page -isnot [string]) { $page = [string]($page -join "`n") }
                $cands = @()
                foreach ($m in [regex]::Matches($page, "https://www\.python\.org/ftp/python/3\.(\d+)\.(\d+)/python-3\.\1\.\2-amd64\.exe")) {
                    $cands += [pscustomobject]@{ ver = [version]("3." + $m.Groups[1].Value + "." + $m.Groups[2].Value); url = $m.Value }
                }
                $cands = $cands | Sort-Object ver -Descending
                $cands = $cands | Where-Object { $_.ver.Major -eq 3 -and $_.ver.Minor -ge 11 }
                if ($cands.Count -gt 0) { $exe = $cands[0].url }
            } catch { Warn "could not read python.org: $($_.Exception.Message)" }
            if (-not $exe) {
                Warn "could not resolve the latest python.org installer - trying a known-good mirror"
                $exe = "https://www.python.org/ftp/python/3.14.0/python-3.14.0-amd64.exe"
            }
            Info "downloading $exe"
            $tmpPy = Join-Path $env:TEMP "python3-setup.exe"
            Invoke-WebRequest -UseBasicParsing -Uri $exe -OutFile $tmpPy -TimeoutSec 600
            Step "running the installer (silent, per-user, pre-pends PATH)..."
            Start-Process -FilePath $tmpPy -ArgumentList "/quiet", `
                "InstallAllUsers=0", "PrependPath=1", "CompileAll=0", "Include_test=0" `
                -Wait
            Remove-Item $tmpPy -ErrorAction SilentlyContinue
            $installed = $true
        }
        Refresh-SessionPath
        $have = Get-PythonVersion
        if ($have) { Ok "Python installed: $have" }
        else { Fail "Python did not install - install it from https://www.python.org/downloads/windows/ and re-run this script" }
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
    $offerClone = Confirm "Clone the OfficeKick repository into $ToolsDir?"
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
