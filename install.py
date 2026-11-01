#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OfficeKick unified installer - SlideKick (PowerPoint) + DocKick (Word) +
MailKick (Outlook). Works on macOS and Windows; auto-detects the platform.

    python3 install.py                interactive menu: Install / Update / Uninstall
    python3 install.py --install      non-interactive install (default answers)
    python3 install.py --update       git pull (if this is a checkout) + update
    python3 install.py --uninstall [--purge-config]
    python3 install.py --dry-run      with any mode: print the actions, run none
    python3 install.py --port N       server port (default 8765)
    python3 install.py --ribbon       opt-in: render the AI ribbon commands
                                      (experimental on this Office build)

On Windows use: py -3 install.py (the py launcher; plain `python` is often
the Microsoft Store alias that does not run scripts).

Design notes:
- the add-in transport is plain http://localhost:<port>: Office allows localhost
  over http, so there is NO certificate, NO trusted CA, NO keychain step;
- no admin rights are needed on either platform (everything is per-user);
- autostart: macOS = login item (the "OfficeKickServer" app bundle, registered
  with an in-bundle SMAppService helper; System Events fallback), Windows =
  a per-user scheduled task at logon (HKCU Run-key fallback);
- expected popups are listed BEFORE anything happens;
- uninstall keeps ~/.officekick (config, profiles, sessions) unless --purge-config.
"""
import argparse
import json
import os
import platform
import re
import shutil
import socket
import subprocess
import sys
import time
import xml.dom.minidom as minidom

APP_NAME = "OfficeKick"
DEFAULT_PORT = 8765
OURS = (  # (addin dir, display name, manifest filename)
    ("slidekick", "SlideKick", "slidekick.xml"),
    ("wordkick",  "DocKick",   "dockick.xml"),
    ("mailkick",  "MailKick",  "mailkick.xml"),
)
GUIDS = (
    "c7b1f2a9-86d3-4e58-9a4b-2f6d5e0c3b81",
    "ce1e8f09-db1c-4598-8f5b-6a65b7802d59",
    "8f4d3b62-91c7-4a2e-bb5d-3c70e6a92f18",
)
CATALOG_KEY = "{C7B1F2A9-86D3-4E58-9A4B-2F6D5E0C3B81}"
CATALOG_BASE = r"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
CATALOG_BASE_15 = r"Software\Microsoft\Office\15.0\WEF\TrustedCatalogs"
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
TASK_NAME = "OfficeKickServer"
MAC_APP_REL = ("Applications", "OfficeKickServer.app")

REPO_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = DEFAULT_PORT
DRY_RUN = False
RIBBON = False


def is_windows():
    return sys.platform.startswith("win")


def is_mac():
    return sys.platform == "darwin"


def out(msg):
    print(msg)


def ok(msg):
    out(("  [ok]   " if is_windows() else "   \x1b[32mv\x1b[0m ") + msg)


def warn(msg):
    out(("  [warn] " if is_windows() else "   \x1b[33m!\x1b[0m  ") + msg)


def step(msg):
    out(("  [..]   " if is_windows() else "   \x1b[36m~\x1b[0m  ") + msg)


def die(msg, code=1):
    out(("  [FAIL] " if is_windows() else "   \x1b[31mx\x1b[0m ") + msg)
    sys.exit(code)


def hp(*parts):
    return os.path.join(os.path.expanduser("~"), *(p for p in parts if p))


def repo_path(*parts):
    return os.path.join(REPO_DIR, *parts)


def read_version():
    try:
        with open(repo_path("VERSION"), "r", encoding="utf-8") as fh:
            return fh.read().strip() or "0.9.x"
    except OSError:
        return "0.9.x"


VERSION = read_version()


def install_dir():
    if is_windows():
        base = os.environ.get("LOCALAPPDATA") or hp("AppData", "Local")
        return os.path.join(base, "OfficeKick")
    return hp("Library", "Application Support", APP_NAME)


def config_dir():
    return hp(".officekick")


def config_file():
    return os.path.join(config_dir(), "config.json")


def manifest_dir():
    return os.path.join(install_dir(), "manifest")


def server_py():
    return os.path.join(install_dir(), "server.py")


def app_bundle():
    return hp(*MAC_APP_REL)


def run_quiet(cmd, **kw):
    """subprocess.run that never raises on a missing binary; returns (rc, output)."""
    try:
        p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, **kw)
        return p.returncode, (p.stdout or b"").decode("utf-8", "replace").strip()
    except (OSError, ValueError):
        return 127, ""


def ask_yes(prompt, default_yes=True):
    if DRY_RUN:
        return default_yes
    try:
        if not sys.stdin.isatty():
            return default_yes   # piped/automated runs take the default silently
    except (OSError, ValueError, AttributeError):
        return default_yes
    suffix = "[Y/n]" if default_yes else "[y/N]"
    while True:
        try:
            ans = input("  > %s %s: " % (prompt, suffix)).strip().lower()
        except EOFError:
            return default_yes
        except KeyboardInterrupt:
            out("\n  aborted")
            raise SystemExit(130)
        if not ans:
            return default_yes
        if ans in ("y", "yes"):
            return True
        if ans in ("n", "no"):
            return False
        out("  please answer y or n")


# ------------------------------------------------------------------ health #
def health(port, timeout=1.0):
    import urllib.request
    try:
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open("http://127.0.0.1:%d/api/health" % port,
                         timeout=timeout) as resp:
            doc = json.loads(resp.read().decode("utf-8", "replace"))
        if doc.get("app") == APP_NAME and doc.get("ok") is True:
            return doc
    except Exception:
        pass
    return None


def port_listening(port, timeout=0.4):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(timeout)
    try:
        return s.connect_ex(("127.0.0.1", port)) == 0
    finally:
        s.close()


def resolve_port():
    """Pick the server port: reuse ours; bump past foreign listeners."""
    global PORT
    p = PORT
    for _ in range(12):
        h = health(p)
        if h:
            if h.get("app") == APP_NAME:
                PORT = p
                ok("server already running on %d (it gets restarted with the new build)" % p)
                return p
            step("port %d busy with something else - moving on" % p)
            p += 1
            continue
        if port_listening(p):
            step("port %d busy (not answering the OfficeKick health) - moving on" % p)
            p += 1
            continue
        PORT = p
        return p
    die("could not find a usable port starting at %d" % PORT)


# ---------------------------------------------------------- server control #
def stop_server():
    step("stopping any running OfficeKick server")
    if DRY_RUN:
        return
    if is_mac():
        for pat in ("Application Support/OfficeKick/server.py",
                    "OfficeKickServer.app",
                    "SlideKick/slidekick_server.py"):
            run_quiet(["pkill", "-f", pat])
    else:
        rc, nout = run_quiet(["netstat", "-ano", "-p", "tcp"])
        pids = set()
        for line in nout.splitlines():
            parts = [x for x in line.split() if x]
            if len(parts) >= 5 and parts[3].upper() == "LISTENING" \
                    and parts[1].split(":")[-1] == str(PORT):
                pids.add(parts[4])
        for pid in pids:
            run_quiet(["taskkill", "/PID", pid, "/F"])
        run_quiet(["schtasks", "/end", "/tn", TASK_NAME])
    deadline = time.time() + 6
    while time.time() < deadline and port_listening(PORT):
        time.sleep(0.4)


def start_server():
    if DRY_RUN:
        out("  [..]  [dry-run] would start the server on port %d" % PORT)
        return True
    if is_mac():
        shim = os.path.join(app_bundle(), "Contents", "MacOS",
                            "officekick-server")
        deadline = time.time() + 15
        attempt = 0
        while time.time() < deadline and not health(PORT):
            attempt += 1
            if attempt <= 2:
                rc, o = run_quiet(["open", "-g", app_bundle()])
                if rc != 0:
                    warn("open -g failed (rc %d): %s" % (rc, o or ""))
            else:
                # direct launch of the wrapper, detached - same mechanics as
                # the login item, minus LaunchServices
                with open(os.devnull, "ab", 0) as dn:
                    subprocess.Popen([shim], stdout=dn, stderr=dn,
                                     start_new_session=True, close_fds=True)
            time.sleep(3)
        h = health(PORT)
        if h:
            ok("server healthy: http://127.0.0.1:%d (version %s)%s" %
               (PORT, h.get("version"), "" if attempt <= 2 else " (direct launch)"))
            return True
        warn("server did not answer a health check on port %d" % PORT)
        return False
    args = [pythonw_path(), server_py(), "--config", config_file(),
            "--static-dir", install_dir(), "--port", str(PORT)]
    flags = 0
    if hasattr(subprocess, "DETACHED_PROCESS"):
        flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    with open(os.devnull, "ab", 0) as dn:
        subprocess.Popen(args, cwd=install_dir(), stdout=dn, stderr=dn,
                         creationflags=flags, close_fds=True)
    deadline = time.time() + 12
    while time.time() < deadline:
        h = health(PORT)
        if h:
            ok("server healthy: http://127.0.0.1:%d (version %s)" % (PORT, h.get("version")))
            return True
        time.sleep(0.5)
    warn("server did not answer a health check on port %d" % PORT)
    return False


def pythonw_path():
    """the windowless python interpreter on Windows."""
    if not is_windows():
        return sys.executable
    if "pythonw" in os.path.basename(sys.executable).lower():
        return sys.executable
    cand = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    return cand if os.path.isfile(cand) else sys.executable


# ------------------------------------------------------------------ deploy #
def deploy_files():
    dst = install_dir()
    out("")
    out("  Deploying files")
    out("    destination: " + dst)
    if DRY_RUN:
        n = len([a for a in OURS if os.path.isdir(repo_path(a[0]))])
        out("  [..]  [dry-run] would copy core/server.py + %d add-in folder(s)" % n)
        return
    shutil.rmtree(dst, ignore_errors=True)
    os.makedirs(manifest_dir(), exist_ok=True)
    os.makedirs(config_dir(), exist_ok=True)
    shutil.copyfile(repo_path("core", "server.py"), server_py())
    ignore = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store")
    deployed = []
    for (folder, name, _mf) in OURS:
        if os.path.isdir(repo_path(folder)):
            shutil.copytree(repo_path(folder), os.path.join(dst, folder),
                            ignore=ignore, dirs_exist_ok=True)
            deployed.append(name)
    cfgjson = os.path.join(config_dir(), "config.json")
    if os.path.exists(cfgjson):
        try:
            os.chmod(cfgjson, 0o600)
        except OSError:
            pass
    ok((" + ".join(deployed) if deployed else "server") + " deployed")


# --------------------------------------------------------------- manifests #
def office_apps_present():
    if is_mac():
        return {
            "slidekick": os.path.isdir("/Applications/Microsoft PowerPoint.app"),
            "wordkick": os.path.isdir("/Applications/Microsoft Word.app"),
            "mailkick": os.path.isdir("/Applications/Microsoft Outlook.app"),
        }
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pfx = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    for base in (os.path.join(pf, "Microsoft Office", "root", "Office16"),
                 os.path.join(pfx, "Microsoft Office", "root", "Office16")):
        if os.path.isdir(base):
            return {
                "slidekick": os.path.isfile(os.path.join(base, "POWERPNT.EXE")),
                "wordkick": os.path.isfile(os.path.join(base, "WINWORD.EXE")),
                "mailkick": os.path.isfile(os.path.join(base, "OUTLOOK.EXE")),
            }
    return {"slidekick": False, "wordkick": False, "mailkick": False}


def app_label(key):
    return {"slidekick": "PowerPoint", "wordkick": "Word", "mailkick": "Outlook"}[key]


def render_manifest(addin_dir, port):
    with open(repo_path(addin_dir, "manifest.template.xml"), "r", encoding="utf-8") as fh:
        t = fh.read()
    if RIBBON:
        t = re.sub(r"(?m)^[ \t]*__COMMANDS_(?:BEGIN|END)__[ \t]*\r?\n?", "", t)
    else:
        t = re.sub(r"(?s)^[ \t]*__COMMANDS_BEGIN__[ \t]*\r?\n.*?^__COMMANDS_END__[ \t]*\r?\n?",
                   "", t, flags=re.M)
    t = t.replace("__SCHEME__", "http").replace("__PORT__", str(port))
    try:
        minidom.parseString(t)
    except Exception as e:
        die("rendered manifest for %s is not valid XML: %s" % (addin_dir, e))
    return t


def render_and_sideload():
    out("")
    step("rendering add-in manifests (http://localhost:%d, %s)" %
         (PORT, "ribbon commands" if RIBBON else "list-mode"))
    if is_mac() and not DRY_RUN:
        purge_stale_wef_manifests()   # before sideloading, or this eats the fresh copies
    apps = office_apps_present()
    rendered = []
    for (folder, name, mfname) in OURS:
        if not os.path.isfile(repo_path(folder, "manifest.template.xml")):
            warn("%s not built in this checkout - skipped" % name)
            continue
        xml = render_manifest(folder, PORT)
        mpath = os.path.join(manifest_dir(), mfname)
        if DRY_RUN:
            out("  [..]  [dry-run] would render %s -> %s" % (name, mpath))
        else:
            with open(mpath, "w", encoding="utf-8") as fh:
                fh.write(xml)
            if not is_mac():
                out("  [ok] %s manifest written: %s" % (name, mpath))
        rendered.append(name)
        if is_mac():
            sideload_mac(folder, mfname, apps.get(folder))
    if rendered:
        ok("manifests rendered: " + ", ".join(rendered))
    if DRY_RUN:
        out("  [..]  [dry-run] sideloading skipped")
        return
    if not is_mac():
        sideload_windows_catalog()
        out("  note: Outlook on Windows does not use folder catalogs - load MailKick once:")
        out("        Outlook -> Get Add-ins (or the ellipsis overflow) -> My add-ins ->")
        out("        Custom Add-ins -> Add a custom add-in -> Add from File ->")
        out("        " + os.path.join(manifest_dir(), "mailkick.xml"))


def sideload_mac(folder, mfname, present):
    cand = []
    for cid in {"slidekick": ["com.microsoft.Powerpoint", "com.microsoft.powerpoint"],
                "wordkick": ["com.microsoft.word"],
                "mailkick": ["com.microsoft.Outlook", "com.microsoft.outlook"]}[folder]:
        cand.append(hp("Library", "Containers", cid, "Data", "Documents", "wef"))
    target = None
    for p in cand:
        if os.path.isdir(p):
            target = p
            break
    if target is None:
        target = cand[0]
    if DRY_RUN:
        out("  [..]  [dry-run] would sideload %s -> %s" % (mfname, target))
        return
    try:
        os.makedirs(target, exist_ok=True)
        shutil.copyfile(os.path.join(manifest_dir(), mfname), os.path.join(target, mfname))
        out("  [ok] %s sideloaded%s: %s" % (
            folder, "" if present is not False
            else " (Office app not detected - installed anyway)", target))
    except OSError as e:
        warn("could not sideload %s: %s" % (mfname, e))


def purge_stale_wef_manifests():
    base_dir = hp("Library", "Containers")
    if not os.path.isdir(base_dir):
        return 0
    removed = 0
    for cid in sorted(os.listdir(base_dir)):
        if not cid.startswith("com.microsoft."):
            continue
        wef_dir = os.path.join(base_dir, cid, "Data", "Documents", "wef")
        if not os.path.isdir(wef_dir):
            continue
        for f in sorted(os.listdir(wef_dir)):
            if not f.endswith(".xml"):
                continue
            p = os.path.join(wef_dir, f)
            try:
                with open(p, "r", encoding="utf-8", errors="replace") as fh:
                    data = fh.read(200000)
            except OSError:
                continue
            if any(g.lower() in data.lower() for g in GUIDS):
                try:
                    os.remove(p)
                    removed += 1
                except OSError:
                    pass
    if removed:
        ok("cleaned %d stale OfficeKick manifest(s) from Office containers" % removed)
    return removed


def sideload_windows_catalog():
    import winreg
    mdir = manifest_dir()
    os.makedirs(mdir, exist_ok=True)
    key = CATALOG_BASE + "\\" + CATALOG_KEY
    try:
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, "Url", 0, winreg.REG_SZ, mdir)
            winreg.SetValueEx(k, "flags", 0, winreg.REG_DWORD, 1)
        ok("trusted folder catalog registered (HKCU %s)" % key)
    except OSError as e:
        warn("could not register the trusted catalog: %s" % e)


# ------------------------------------------------------------ office caches #
def clear_office_caches():
    """Forget stale Office add-in registrations + webview caches (safe ONLY
    while the Office apps are not running: the stores rebuild at launch)."""
    if is_windows():
        return
    if DRY_RUN:
        out("  [..]  [dry-run] would clear the Office add-in caches (Webkit/WEF stores)")
        return
    base_dir = hp("Library", "Containers")
    if not os.path.isdir(base_dir):
        return
    for cid in sorted(os.listdir(base_dir)):
        if not (cid.startswith("com.microsoft.") or cid.startswith("UFB")):
            continue
        data = os.path.join(base_dir, cid, "Data", "Library")
        caches = os.path.join(data, "Caches")
        if os.path.isdir(caches):
            for sub in ("Microsoft", "com.microsoft.Powerpoint", "com.microsoft.ctrlstrcaches",
                        "com.microsoft.Word", "WebKit"):
                shutil.rmtree(os.path.join(caches, sub), ignore_errors=True)
        wef_store = os.path.join(data, "Application Support", "Microsoft", "Office", "16.0", "Wef")
        if os.path.isdir(wef_store):
            shutil.rmtree(wef_store, ignore_errors=True)
    ok("cleared Office add-in caches + WEF registration stores (rebuilt at launch)")


def office_app_running():
    if is_mac():
        rc, o = run_quiet(["bash", "-c",
                           "pgrep -f 'Microsoft (PowerPoint|Word|Outlook).app' >/dev/null 2>&1 && echo yes"])
        return rc == 0 and "yes" in o
    return False


# -------------------------------------------------------- autostart: macOS #
def build_app_bundle():
    a = app_bundle()
    out("")
    out("  Local server app: " + a)
    exe_dir = os.path.join(a, "Contents", "MacOS")
    if DRY_RUN:
        out("  [..]  [dry-run] would (re)write the bundle (http shim, no certificate flags),")
        out("        compile the SMAppService registrar and ad-hoc sign it")
        return None
    os.makedirs(exe_dir, exist_ok=True)
    os.makedirs(os.path.join(a, "Contents", "Resources"), exist_ok=True)
    with open(os.path.join(a, "Contents", "Info.plist"), "w", encoding="utf-8") as fh:
        fh.write(
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" '
            '"http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
            '<plist version="1.0">\n<dict>\n'
            '  <key>CFBundleIdentifier</key><string>com.corelan.officekick.server</string>\n'
            '  <key>CFBundleName</key><string>OfficeKickServer</string>\n'
            '  <key>CFBundleDisplayName</key><string>OfficeKick Server</string>\n'
            '  <key>CFBundleExecutable</key><string>officekick-server</string>\n'
            '  <key>CFBundlePackageType</key><string>APPL</string>\n'
            '  <key>CFBundleIconFile</key><string>OfficeKickAdmin</string>\n'
            '  <key>NSLocalNetworkUsageDescription</key>\n'
            '  <string>OfficeKick reaches AI providers you run on your own network '
            '(Ollama, llama.cpp) to answer requests from your Office add-ins.</string>\n'
            '  <key>LSUIElement</key><true/>\n'
            '</dict>\n</plist>\n')
    # wrapper shim: child execution (Local Network attribution) + KeepAlive
    def q(s):
        return "'" + s.replace("'", "'\\''") + "'"
    shim = "#!/bin/bash\n"
    shim += "# OfficeKickServer wrapper (com.corelan.officekick.server).\n"
    shim += "# With arguments: runs them as a CHILD of this in-bundle process.\n"
    shim += "# Without arguments (login item / open): launches the installed server.\n"
    shim += "if [[ $# -eq 0 ]]; then\n"
    shim += "  set -- %s \\\n" % q(sys.executable)
    shim += "    %s --config %s --static-dir %s --port %s\n" % (
        q(server_py()), q(config_file()), q(install_dir()), PORT)
    shim += "fi\n"
    shim += "# KeepAlive: respawn only after an abnormal exit; a graceful stop (SIGTERM\n"
    shim += "# to this wrapper = exit code 0 from the child) ends the loop\n"
    shim += "stopping=0\n"
    shim += "trap 'stopping=1' TERM INT\n"
    shim += 'child=""\n'
    shim += 'cleanup() { [[ -n "$child" ]] && kill "$child" 2>/dev/null; exit 0; }\n'
    shim += "trap cleanup TERM INT\n"
    shim += "until [[ $stopping == 1 ]]; do\n"
    shim += '  "$@" &\n'
    shim += "  child=$!\n"
    shim += '  wait "$child"\n'
    shim += "  rc=$?\n"
    shim += '  child=""\n'
    shim += "  [[ $rc -eq 0 ]] && exit 0\n"
    shim += '  echo "officekick-server wrapper: server died (rc=$rc), restarting in 3s" >&2\n'
    shim += "  sleep 3\n"
    shim += "done\n"
    shim += "exit 0\n"
    with open(os.path.join(exe_dir, "officekick-server"), "w", encoding="utf-8") as fh:
        fh.write(shim)
    os.chmod(os.path.join(exe_dir, "officekick-server"), 0o755)
    helper_src = repo_path("tools", "officekick-register-login-item.m")
    helper = os.path.join(exe_dir, "officekick-register-login-item")
    if os.path.isfile(helper_src) and shutil.which("clang"):
        rc, o = run_quiet(["clang", "-framework", "Foundation", "-framework",
                           "ServiceManagement", "-o", helper, helper_src])
        if rc == 0 and os.path.isfile(helper):
            os.chmod(helper, 0o755)
        else:
            warn("login-item registrar helper compile failed - System Events fallback stays")
    else:
        try:
            os.remove(helper)
        except OSError:
            pass
    rc, o = run_quiet(["codesign", "--force", "--deep", "-s", "-", a])
    if rc != 0:
        warn("ad-hoc code signing failed: %s" % (o or ""))
    ok("server app bundle ready")
    return helper


def mac_login_item_register():
    helper = os.path.join(app_bundle(), "Contents", "MacOS",
                          "officekick-register-login-item")
    if os.path.isfile(helper) and os.access(helper, os.X_OK):
        rc, _ = run_quiet([helper, "--check"])
        if rc == 0:
            ok("login item already registered (SMAppService; shows as OfficeKickServer)")
            return True
        run_quiet(["osascript", "-e",
                   'tell application "System Events" to delete '
                   '(every login item whose name is "OfficeKickServer")'])
        rc, o = run_quiet([helper, "register"])
        if rc == 0:
            ok("login item registered via SMAppService "
               "(shows as OfficeKickServer; runs in your session)")
            return True
        warn("SMAppService registration refused (%s) - falling back" % (o or "unknown"))
    # legacy fallback: macOS attributes its one-time background-item
    # notification to "System Events.app" here (harmless, and rare: only when
    # the helper could not be built on this machine)
    run_quiet(["osascript", "-e",
               'tell application "System Events" to delete '
               '(every login item whose name is "OfficeKickServer")'])
    rc, o = run_quiet(["osascript", "-e",
                       'tell application "System Events" to make login item at end '
                       'with properties {path:"%s", hidden:true}' % app_bundle()])
    if rc == 0:
        ok("login item registered (System Events fallback; starts at login, in your session)")
        return True
    warn("login item registration failed - start the server manually with:")
    warn("  %s %s --config %s --static-dir %s --port %d" %
         (sys.executable, server_py(), config_file(), install_dir(), PORT))
    return False


def autostart_mac():
    build_app_bundle()
    out("")
    step("registering autostart (macOS login item)")
    if DRY_RUN:
        out("  [..]  [dry-run] would register the OfficeKickServer login item")
        return True
    return mac_login_item_register()


def mac_login_item_present():
    helper = os.path.join(app_bundle(), "Contents", "MacOS",
                          "officekick-register-login-item")
    if os.path.isfile(helper) and os.access(helper, os.X_OK):
        rc, _ = run_quiet([helper, "--check"])
        if rc == 0:
            return True
    rc, o = run_quiet(["osascript", "-e",
                       'tell application "System Events" to '
                       '(get name of every login item) contains "OfficeKickServer"'])
    return rc == 0 and o.strip().lower() == "true"


# ------------------------------------------------------ autostart: windows #
def windows_task_cmd():
    return '"%s" "%s" --config "%s" --static-dir "%s" --port %d' % (
        pythonw_path(), server_py(), config_file(), install_dir(), PORT)


def autostart_windows():
    out("")
    step("registering autostart (per-user scheduled task at logon)")
    if DRY_RUN:
        out("  [..]  [dry-run] would create task '%s': %s" % (TASK_NAME, windows_task_cmd()))
        return True
    tr = windows_task_cmd()
    rc, o = run_quiet(["schtasks", "/create", "/f", "/sc", "onlogon",
                       "/tn", TASK_NAME, "/tr", tr])
    if rc == 0:
        ok("scheduled task created (starts the server at logon): %s" % TASK_NAME)
        return True
    warn("schtasks failed (%s) - using the HKCU Run key instead" % (o or "rc %d" % rc))
    try:
        import winreg
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, TASK_NAME, 0, winreg.REG_SZ, tr)
        ok("Run-key autostart written (HKCU \\...\\Run\\%s)" % TASK_NAME)
        return True
    except OSError as e:
        warn("autostart could not be registered: %s" % e)
        return False


def windows_autostart_present():
    if not is_windows():
        return False
    rc, _ = run_quiet(["schtasks", "/query", "/tn", TASK_NAME])
    if rc == 0:
        return True
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_QUERY_VALUE) as k:
            winreg.QueryValueEx(k, TASK_NAME)
        return True
    except OSError:
        return False


# -------------------------------------------------------------- uninstall #
def do_uninstall(purge):
    out("")
    out("  Uninstalling OfficeKick")
    stop_server()
    if is_mac():
        if DRY_RUN:
            out("  [..]  [dry-run] would unregister the login item and remove %s" % app_bundle())
        else:
            helper = os.path.join(app_bundle(), "Contents", "MacOS",
                                  "officekick-register-login-item")
            if os.path.isfile(helper) and os.access(helper, os.X_OK):
                rc, o = run_quiet([helper, "--unregister"])
                if rc == 0:
                    ok("login item unregistered (SMAppService)")
                else:
                    warn("login item unregister failed (%s)" % (o or "?"))
            run_quiet(["osascript", "-e",
                       'tell application "System Events" to delete '
                       '(every login item whose name is "OfficeKickServer")'])
            for label in ("com.corelan.officekick", "com.corelan.slidekick"):
                lp = hp("Library", "LaunchAgents", label + ".plist")
                run_quiet(["launchctl", "unload", lp])
                if os.path.isfile(lp):
                    try:
                        os.remove(lp)
                        ok("LaunchAgent removed: %s" % label)
                    except OSError:
                        pass
            if os.path.isdir(app_bundle()):
                shutil.rmtree(app_bundle(), ignore_errors=True)
                ok("removed " + app_bundle())
            purge_stale_wef_manifests()
    else:
        if DRY_RUN:
            out("  [..]  [dry-run] would remove task '%s', the Run key and the trusted catalog" % TASK_NAME)
        else:
            run_quiet(["schtasks", "/delete", "/f", "/tn", TASK_NAME])
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.DeleteValue(k, TASK_NAME)
                ok("Run-key autostart removed")
            except OSError:
                pass
            for base in (CATALOG_BASE, CATALOG_BASE_15):
                try:
                    import winreg
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, base + "\\" + CATALOG_KEY)
                    ok("trusted folder catalog removed")
                except OSError:
                    pass
    if office_app_running():
        warn("Office apps are running - skipping the cache flush (quit Office, then "
             "delete leftover caches by re-running this uninstall)")
    else:
        clear_office_caches()
    inst = install_dir()
    if DRY_RUN:
        out("  [..]  [dry-run] would remove " + inst)
    else:
        if os.path.isdir(inst):
            shutil.rmtree(inst, ignore_errors=True)
            ok("removed " + inst)
    cfg = config_dir()
    if purge is None:
        purge = (not DRY_RUN) and ask_yes(
            "Also remove the config (%s - stored API keys, profiles, sessions)?" % cfg, False)
    if purge:
        if DRY_RUN:
            out("  [..]  [dry-run] would wipe " + cfg)
        else:
            shutil.rmtree(cfg, ignore_errors=True)
            ok("config wiped: " + cfg)
    else:
        out("  [ok] config kept at %s (keys, profiles and sessions intact)" % cfg)
    out("")
    warn("quit and reopen PowerPoint / Word / Outlook to finish removing the add-ins")


# ---------------------------------------------------------------- install #
def preflight(mode):
    out("")
    out("OfficeKick %s - %s (SlideKick + DocKick + MailKick)" % (VERSION, mode))
    out("  platform  : %s (%s) | python %s" % (
        platform.system(), platform.release(), platform.python_version()))
    out("  transport : http://localhost:<port> - Office allows localhost over http,")
    out("              so this installer uses NO certificate and NO https at all.")
    out("  admin     : not required - every step here is per-user.")
    if is_mac():
        out("  popups    : macOS may show -")
        out("    * a one-time 'OfficeKickServer can run in the background' note when the")
        out("      autostart login item is first registered - that IS this add-in's own")
        out("      local server app (correctly named), nothing external is installed;")
        out("    * a Local Network prompt the first time a model on your LAN is used -")
        out("      click Allow so the panes can reach e.g. Ollama on another machine.")
    else:
        out("  popups    : Windows may show -")
        out("    * Windows Defender Firewall may ask once whether Python may listen -")
        out("      click Allow (the add-ins talk to 127.0.0.1 only, but Allow keeps")
        out("      LAN models like Ollama working later);")
        out("    * no SmartScreen: this installer is a plain python script, nothing is")
        out("      downloaded; the add-ins come from the local trusted catalog.")


def ask_autostart():
    out("")
    out("  About autostart: the add-ins talk to a small LOCAL server (python) that")
    out("  serves http://127.0.0.1:%d - the Office task panes and the AI requests." % PORT)
    out("  Without autostart the panes only work while that server is running.")
    if is_mac():
        return ask_yes("Add autostart to the login items (a hidden 'OfficeKickServer' app"
                       " starts it at every login - no window, idle when unused)?", True)
    return ask_yes("Add an autostart scheduled task (a windowless pythonw starts the server"
                   " at every logon)?", True)


def install(mode_name="Install", interactive=True):
    preflight(mode_name)
    stop_server()          # first: free our canonical port (a previous, possibly
    resolve_port()         # https-only install answers no http health check)
    if interactive and not DRY_RUN:
        want = ask_autostart()
    else:
        present = mac_login_item_present() if is_mac() else windows_autostart_present()
        if present:
            ok("existing autostart found - it will be kept/re-registered")
            want = True
        else:
            out("  [..]  no autostart found - it will be added (default)")
            want = True
    stop_server()
    deploy_files()
    if office_app_running():
        warn("Office apps are running - skipping the add-in cache clear"
             " (quit Office, re-run this install, reopen the apps to fully flush"
             " an old registration)")
    else:
        clear_office_caches()
    render_and_sideload()
    if want:
        if is_mac():
            autostart_mac()
        else:
            autostart_windows()
    else:
        out("  [ok] no autostart - start the server manually when needed:")
        out("       %s %s --config %s --static-dir %s --port %d" %
            (sys.executable, server_py(), config_file(), install_dir(), PORT))
    start_server()
    done_note(mode_name)


def done_note(mode_name):
    apps = office_apps_present()
    verb = "updated for" if mode_name == "Update" else "installed for"
    out("")
    for (folder, name, _mf) in OURS:
        if os.path.isdir(repo_path(folder)) and apps.get(folder):
            ok("%-9s %s %s" % (name, verb, app_label(folder)))
    out("")
    out("  Quit and REOPEN your Office apps so the task panes load the new build")
    out("  (About tab -> 'Pane build' should say %s)." % read_version())
    out("  Simulator (no Office needed): http://127.0.0.1:%d/slidekick/static/simulator.html" % PORT)
    out("  Config: %s - API keys stay local (chmod 600)." % config_dir())


def git_pull():
    if not os.path.isdir(repo_path(".git")):
        step("not a git checkout - skipping pull")
        return
    out("")
    step("git pull ...")
    rc, o = run_quiet(["git", "-C", REPO_DIR, "pull", "--ff-only"])
    if rc == 0:
        ok("repo updated (git pull --ff-only)")
        for line in (o or "").splitlines()[-2:]:
            out("    " + line)
    else:
        warn("git pull failed - continuing with the current checkout%s" %
             ((": " + o) if o else ""))


# ------------------------------------------------------------------- flows #
def main():
    global PORT, DRY_RUN, RIBBON
    ap = argparse.ArgumentParser(
        description="OfficeKick cross-platform installer (macOS + Windows)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Examples:  python3 install.py  (Windows: py -3 install.py)  |  "
               "python3 install.py --update  |  python3 install.py --uninstall --purge-config")
    ap.add_argument("--install", action="store_true", help="interactive install now")
    ap.add_argument("--update", action="store_true", help="git pull + update (non-interactive)")
    ap.add_argument("--uninstall", action="store_true", help="remove the add-ins + server")
    ap.add_argument("--purge-config", action="store_true",
                    help="with --uninstall: also wipe ~/.officekick")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT, help="server port (default %d)" % DEFAULT_PORT)
    ap.add_argument("--ribbon", action="store_true",
                    help="opt-in: render the AI ribbon commands (experimental on this Office build)")
    ap.add_argument("--dry-run", action="store_true", help="print the actions, run none")
    args = ap.parse_args()
    PORT = args.port if 1024 <= args.port <= 65535 else DEFAULT_PORT
    DRY_RUN = args.dry_run
    RIBBON = args.ribbon

    if not (is_mac() or is_windows()):
        die("this installer supports macOS and Windows only (detected: %s)." % sys.platform)

    if args.uninstall:
        do_uninstall(True if args.purge_config else None)
        return
    if args.update:
        git_pull()
        install("Update", interactive=False)
        return
    if args.install:
        install("Install", interactive=False)
        return
    while True:
        out("")
        out("OfficeKick %s - installer (auto-detected: %s)" % (VERSION, platform.system()))
        out("  1) Install (or repair the current install)")
        out("  2) Update the current install (also pulls the repo with git)")
        out("  3) Uninstall")
        out("  q) quit")
        try:
            if not sys.stdin.isatty():
                out("  (non-interactive: no choice given - pass --install / --update / --uninstall)")
                return
            choice = input("  > choose [1/2/3, default 1]: ").strip().lower()
        except EOFError:
            return
        if choice in ("", "1"):
            install("Install", interactive=True)
            return
        if choice == "2":
            git_pull()
            install("Update", interactive=False)
            return
        if choice == "3":
            do_uninstall(None)
            return
        if choice in ("q", "quit", "exit"):
            return
        out("  please choose 1, 2, 3 or q")


if __name__ == "__main__":
    main()