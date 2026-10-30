<#
.SYNOPSIS
  OfficeKick installer for Windows - SlideKick (PowerPoint) and DocKick (Word).
  The add-in code, task pane and local server are identical to macOS; only
  registration differs: a per-user trusted catalog (registry) instead of the
  Mac Wef folder, and a Scheduled Task instead of a LaunchAgent.

.USAGE
  .\install.ps1                          install or update (idempotent)
  .\install.ps1 -Port 8765               fixed port
  .\install.ps1 -Dev                     serve straight from this repo
  .\install.ps1 -NoScheduledTask         skip auto-start registration
  .\install.ps1 -Action status           show install status
  .\install.ps1 -Action run              run the server in the foreground
  .\install.ps1 -Action uninstall -Purge remove; -Purge also wipes config

.NOTES
  Works under Windows PowerShell 5.1 and PowerShell 7. Per-user only - no
  administrator rights required. The add-in appears in PowerPoint under
  Insert -> My Add-ins -> SHARED FOLDER after restarting PowerPoint.
#>
param(
    [int]$Port = 0,
    [switch]$Dev,
    [switch]$NoScheduledTask,
    [switch]$Purge,
    [ValidateSet("install", "uninstall", "status", "run", "help")]
    [string]$Action = "install"
)

$ErrorActionPreference = "Stop"

$App           = "OfficeKick"
$TaskName      = "OfficeKickServer"
$StartupBat    = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup\officekick-server.cmd"
$RepoDir       = $PSScriptRoot
$SrcDir = Join-Path $RepoDir "slidekick"
$InstallDir    = Join-Path $env:LOCALAPPDATA $App
$ManifestDir   = Join-Path $InstallDir "manifest"
$ConfigDir     = Join-Path $env:USERPROFILE ".officekick"
$LogDir        = Join-Path $InstallDir "logs"
$ManifestName  = "slidekick.xml"
$CatalogGUID   = "C7B1F2A9-86D3-4E58-9A4B-2F6D5E0C3B81"
$CatalogBase   = "HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs"
$DefaultPort   = 8765
$Version       = "0.1.0"
$verFile       = Join-Path $RepoDir "VERSION"
if (Test-Path $verFile) { $Version = (Get-Content $verFile -Raw).Trim() }

function Say([string]$m)  { Write-Host $m }
function Ok([string]$m)   { Write-Host "  [ok]   $m" -ForegroundColor Green }
function Warn([string]$m) { Write-Host "  [warn] $m" -ForegroundColor Yellow }
function Die([string]$m)  { Write-Host "  [fail] $m" -ForegroundColor Red; exit 1 }

function Find-PowerPoint {
    $pp = $null
    foreach ($reg in @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe")) {
        if (Test-Path $reg) {
            $v = (Get-ItemProperty $reg -ErrorAction SilentlyContinue).'(default)'
            if ($v) {
                $v = $v.Trim('"')
                if (Test-Path $v) { $pp = $v; break }
            }
        }
    }
    if (-not $pp) {
        foreach ($p in @(
            $(Join-Path $env:ProgramFiles "Microsoft Office\root\Office16\POWERPNT.EXE"),
            $(Join-Path (${env:ProgramFiles(x86)}) "Microsoft Office\root\Office16\POWERPNT.EXE"),
            $(Join-Path $env:ProgramFiles "Microsoft Office\Office16\POWERPNT.EXE"))) {
            if ($p -and (Test-Path $p)) { $pp = $p; break }
        }
    }
    return $pp
}

function Find-Python {
    $candidates = @()
    foreach ($cmd in @("python.exe", "py.exe", "python3.exe")) {
        try { $c = (Get-Command $cmd -ErrorAction Stop).Source; if ($c) { $candidates += $c } } catch {}
    }
    foreach ($exe in $candidates) {
        if ($exe -like "*Microsoft*WindowsApps*") { continue }  # MS Store alias stub
        try {
            $out = & $exe -c 'import sys; print("%d.%d" % sys.version_info[:2])' 2>$null
            if ($LASTEXITCODE -eq 0 -and $out -match '^3\.(\d+)') {
                if ([int]$Matches[1] -ge 8) { return $exe }
            }
        } catch {}
    }
    Die "python3 (>= 3.8) not found - install from python.org or: winget install Python.Python.3.12"
}

function Test-PortFree([int]$p) {
    $listener = $null
    try {
        $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $p)
        $listener.Start()
        return $true
    } catch {
        return $false
    } finally {
        if ($listener) { $listener.Stop() }
    }
}

function HealthCheck([int]$p, [int]$tries) {
    for ($i = 0; $i -lt $tries; $i++) {
        try {
            $r = Invoke-WebRequest -Uri "http://127.0.0.1:$p/api/health" -UseBasicParsing -TimeoutSec 2
            if ($r.Content -like "*OfficeKick*") { return $r.Content }
        } catch {
            Start-Sleep -Milliseconds 500
        }
    }
    return $null
}

function Pick-Port {
    if ($Port -gt 0) { return $Port }
    $h = HealthCheck $DefaultPort 1
    if ($h) { return $DefaultPort }  # our own server already owns it (update path)
    if (Test-PortFree $DefaultPort) { return $DefaultPort }
    for ($p = $DefaultPort + 1; $p -le $DefaultPort + 19; $p++) {
        if (Test-PortFree $p) { return $p }
    }
    Die "no free port found in range $DefaultPort..$($DefaultPort + 19)"
}

function Stop-Server {
    $procs = Get-CimInstance Win32_Process -Filter "Name like 'py%'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -and $_.CommandLine -match "[Oo]fficeKick\\server\.py|[Ss]lideKick\\slidekick_server\.py" }
    foreach ($proc in $procs) {
        try { Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop } catch {}
    }
    if ($procs) { Ok "server process stopped" }
}

function Register-Autostart([string]$py, [string]$serverPy, [int]$p) {
    if ($NoScheduledTask) { Warn "skipping auto-start (-NoScheduledTask)"; return }
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        $action = New-ScheduledTaskAction -Execute $py -Argument ('"' + $serverPy + '" --port ' + $p)
        $trigger = New-ScheduledTaskTrigger -AtLogOn
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) `
            -ExecutionTimeLimit (New-TimeSpan -Days 3650)
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
            -Settings $settings -Description "OfficeKick local AI server" -ErrorAction Stop | Out-Null
        Ok "scheduled task '$TaskName' registered (starts at logon, restarts on crash)"
    } catch {
        # fallback: Startup folder (no restart-on-crash, but always permitted)
        $bat = "@echo off`r`nstart `"SlideKick`" /min `"$py`" `"$serverPy`" --port $p`r`n"
        Set-Content -Path $StartupBat -Value $bat -Encoding ASCII
        Warn "could not register scheduled task ($($_.Exception.Message)); using Startup folder .cmd instead"
    }
}

if ($Action -eq "help") {
    Say "OfficeKick Windows installer:"
    Say ""
    Say "  install (default)       install or update the add-in + server"
    Say "  -Port N                 port for the local server (default $DefaultPort)"
    Say "  -Dev                    serve files directly from this repo (developer mode)"
    Say "  -NoScheduledTask        do not register the logon auto-start"
    Say "  status                  show current install status"
    Say "  run                     run the server in the foreground"
    Say "  uninstall [-Purge]      remove everything (config too with -Purge)"
    exit 0
}

if ($Action -eq "uninstall") {
    Say "OfficeKick - uninstall (SlideKick + DocKick)"
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop; Ok "scheduled task removed" } catch { Warn "no scheduled task" }
    if (Test-Path $StartupBat) { Remove-Item $StartupBat -Force; Ok "Startup .cmd removed" }
    if (Test-Path "$CatalogBase\$CatalogGUID") { Remove-Item "$CatalogBase\$CatalogGUID" -Recurse -Force; Ok "trusted catalog removed" }
    Stop-Server
    if (-not $Dev -and (Test-Path $InstallDir)) { Remove-Item $InstallDir -Recurse -Force; Ok "$InstallDir removed" }
    if ($Purge) { Remove-Item $ConfigDir -Recurse -Force -ErrorAction SilentlyContinue; Ok "config wiped ($ConfigDir)" }
    else { Ok "config kept at $ConfigDir (use -Purge to wipe)" }
    Warn "quit and reopen PowerPoint to finish removing the add-in"
    exit 0
}

if ($Action -eq "status") {
    $h = HealthCheck $DefaultPort 1
    if ($h) { Ok "server is running: $h" } else { Warn "server NOT running (default port $DefaultPort)" }
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($task) { Ok "scheduled task present ($($task.State))" } else { Warn "scheduled task absent" }
    if (Test-Path $StartupBat) { Ok "Startup .cmd present" }
    if (Test-Path "$CatalogBase\$CatalogGUID") { Ok "trusted catalog registered (manifests in $ManifestDir)" } else { Warn "trusted catalog not registered" }
    if (Test-Path (Join-Path $ManifestDir $ManifestName)) { Ok "manifest present: $(Join-Path $ManifestDir $ManifestName)" }
    if (Test-Path (Join-Path $ConfigDir "config.json")) { Ok "config present: $(Join-Path $ConfigDir "config.json")" } else { Warn "no config yet" }
    exit 0
}

if ($Action -eq "run") {
    $py = Find-Python
    if ($Dev) { $serverPy = Join-Path $SrcDir "server.py" }
    else {
        $serverPy = Join-Path $InstallDir "server.py"
        if (-not (Test-Path $serverPy)) { Die "not installed yet (run .\install.ps1 first)" }
    }
    $p = Pick-Port
    & $py $serverPy --port $p
    exit $LASTEXITCODE
}

# ------------------------------------------------------------------ install
Say "OfficeKick $Version - install/update (Windows): SlideKick + DocKick"
Stop-Server   # stop a previously installed server so updates take effect
$py = Find-Python
Ok "python: $py (& $py --version)"

$pp = Find-PowerPoint
if ($pp) {
    $verInfo = (Get-Item $pp).VersionInfo.ProductVersion
    Ok "PowerPoint found: $pp (version $verInfo)"
} else {
    Warn "Microsoft PowerPoint (powerpnt.exe) not found - installing anyway"
}

$Port = Pick-Port
Ok "port: $Port"

if ($Dev) {
    $serverPy  = Join-Path $SrcDir "server.py"
    $staticDir = Join-Path $SrcDir "static"
    Ok "dev mode: serving from $SrcDir (no copy)"
} else {
    foreach ($d in @($InstallDir, $ConfigDir, $LogDir, $ManifestDir)) {
        if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    }
    robocopy $SrcDir "$InstallDir\slidekick" /MIR /XD "__pycache__" ".DS_Store" /NFL /NDL /NJH /NJS | Out-Null
    if ($LASTEXITCODE -gt 7) { Die "robocopy failed with exit code $LASTEXITCODE" }
    Copy-Item (Join-Path $RepoDir "core\server.py") (Join-Path $InstallDir "server.py") -Force
    # DocKick (Word) subtree: deployed when its manifest template exists
    if (Test-Path (Join-Path $RepoDir "wordkick\manifest.template.xml")) {
      robocopy (Join-Path $RepoDir "wordkick") "$InstallDir\wordkick" /MIR /XD "__pycache__" ".DS_Store" /NFL /NDL /NJH /NJS | Out-Null
    }
    $serverPy  = Join-Path $InstallDir "server.py"
    $staticDir = $InstallDir
    Ok "files deployed to $InstallDir"
}

& $py -m py_compile $serverPy
if ($LASTEXITCODE -ne 0) { Die "server source failed to compile" }

# ---- manifest ----
$tpl = Join-Path $SrcDir "manifest.template.xml"
$xml = (Get-Content $tpl -Raw) -replace "__PORT__", "$Port" -replace "__VERSION__", "$Version.0"
$manifestPath = Join-Path $ManifestDir $ManifestName
Set-Content -Path $manifestPath -Value $xml -Encoding UTF8
try {
    [xml](Get-Content $manifestPath -Raw) > $null
    Ok "manifest rendered + valid XML"
} catch {
    Die "generated manifest is not valid XML: $($_.Exception.Message)"
}

# ---- trusted catalog (per-user, no admin) ----
if (-not (Test-Path $CatalogBase)) { New-Item -Path $CatalogBase -Force | Out-Null }
New-Item -Path (Join-Path $CatalogBase $CatalogGUID) -Force | Out-Null
Set-ItemProperty -Path (Join-Path $CatalogBase $CatalogGUID) -Name "Url"   -Value $ManifestDir
Set-ItemProperty -Path (Join-Path $CatalogBase $CatalogGUID) -Name "Flags" -Value 1 -Type DWord
Ok "manifest catalog registered: $ManifestDir (Insert -> My Add-ins -> SHARED FOLDER after restart)"

# ---- clear the Office add-in cache (only while PowerPoint is closed) ----
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Warn "PowerPoint is running - NOT clearing the Office add-in caches."
    Warn "quit PowerPoint, re-run .\install.ps1 (it clears them), then reopen"
} else {
    $wefCache = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
    if (Test-Path $wefCache) {
        Remove-Item (Join-Path $wefCache "*") -Recurse -Force -ErrorAction SilentlyContinue
        Ok "cleared Office add-in cache: $wefCache"
    }
}

# ---- auto-start + server ----
Register-Autostart $py $serverPy $Port

$h = HealthCheck $Port 20
if (-not $h) {
    Start-Process -FilePath $py -ArgumentList ('"' + $serverPy + '" --port ' + $Port) -WindowStyle Hidden
    $h = HealthCheck $Port 20
}
if ($h) { Ok "server restarted + healthy: http://127.0.0.1:$Port  (config: $ConfigDir\config.json)" }
else { Die "server did not become healthy on port $Port - check $(Join-Path $LogDir "server.log") and the task scheduler" }

Say ""
Say "Installed:"
Say "  add-in manifest : $manifestPath  (trusted catalog: SHARED FOLDER)"
Say "  app files       : $InstallDir"
Say "  config (keys)   : $ConfigDir\config.json"
Say "  logs            : $LogDir"
Say "  simulator       : http://127.0.0.1:$Port/simulator.html"
Say ""
Say "First use in PowerPoint:"
Say "  1. quit and reopen PowerPoint (required when it was already running)"
Say "  2. Insert -> My Add-ins (dropdown) -> SHARED FOLDER -> Corelan SlideKick -> Add"
Say "  3. Connections tab: add a provider; Profiles tab: pick your models"
if (-not $pp) { Warn "open the add-in once PowerPoint is installed" }