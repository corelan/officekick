#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OfficeKick unified installer - SlideKick (PowerPoint) + DocKick (Word) +
MailKick (Outlook). Works on macOS and Windows; auto-detects the platform.

    python3 install.py                menu: Install / Update / Server / Transport / Uninstall
    python3 install.py --install      non-interactive install (Word + PowerPoint;
                                      Outlook only with --mailkick / --mailkick-cert-key
                                      or when Outlook was previously installed)
    python3 install.py --update       git pull (if this is a checkout) + update
    python3 install.py --uninstall [--remove-certs] [--purge-config]
    python3 install.py --start | --stop | --restart | --status
    python3 install.py --dry-run      with any mode: print the actions, run none
    python3 install.py --port N       server port (default 8765)
    python3 install.py --ribbon       opt-in: render the AI ribbon commands
                                      (experimental on this Office build)
    python3 install.py --mailkick     non-interactive opt-in: install Outlook and
                                      create the local self-signed certificate (leaf mode)

Installer output is structured: short step lines on the console, the full
detail trail (file lists, tool output, decisions) in install.log next to
install.py. Word + PowerPoint are enabled by default; Outlook is off by
default - enabling Outlook needs the https prerequisites and then asks for
a certificate mode: a self-signed leaf (default, no CA), a loopback CA +
signed certificate, or your own trusted pair. On Windows no certificate is
needed (Outlook rides the plain-http localhost listener there).

On Windows use: py -3 install.py (the py launcher; plain `python` is often
the Microsoft Store alias that does not run scripts).

Design notes:
- the add-in transport is http://localhost:<port> (https://localhost:<port+1>
  wherever a working loopback certificate exists): Office allows localhost
  over http, so there is NO certificate, NO trusted CA, NO keychain step -
  except Outlook on macOS, which validates every manifest URL as https and
  therefore REQUIRES the loopback TLS component;
- no admin rights are needed on either platform (everything is per-user);
- autostart: macOS = login item (the "OfficeKickServer" app bundle, registered
  with an in-bundle SMAppService helper; System Events fallback), Windows =
  a per-user scheduled task at logon (HKCU Run-key fallback);
- uninstall keeps ~/.officekick (config, profiles, sessions, certificates)
  unless --purge-config / --remove-certs (or the same answers in the interactive
  uninstall). Both removal asks default to KEEP.
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
SKIP_MAILKICK = False
MAILKICK_CERT = ""
MAILKICK_KEY = ""
MAILKICK_CHOICE = None     # set per-platform in main(): default ON on Windows, OFF on macOS


def is_windows():
    return sys.platform.startswith("win")


def is_mac():
    return sys.platform == "darwin"


LOG_HANDLE = None         # install.log (repo root) - full detail trail
QUIET = False             # when True: out/step print only into install.log
PHASE_DIRTY = False       # a phase line is pending on the console
ADDINS = set()            # set of add-in folder names for THIS run
FIRST_RUN = set()         # add-ins freshly added THIS run (never installed before)
INTERACTIVE = False       # set per run: the flow may ask questions
AUTOSTART_WANT = False
HTTPS_PORT = None         # resolved from the persisted settings (defaults to PORT + 1)


def _log(msg, ts=False):
    """install.log-only line (never on the console)."""
    try:
        if LOG_HANDLE is None:
            return
        line = msg.rstrip("\n")
        if ts:
            line = time.strftime("%H:%M:%S ") + line
        LOG_HANDLE.write(line + "\n")
        LOG_HANDLE.flush()
    except Exception:
        pass


def log_open(mode):
    global LOG_HANDLE
    try:
        LOG_HANDLE = open(repo_path("install.log"), "a", encoding="utf-8")
    except OSError:
        LOG_HANDLE = None
        return
    _log("")
    _log("=" * 76)
    _log("%s  |  %s  |  python %s  |  %s %s" % (
        time.strftime("%Y-%m-%d %H:%M:%S"), mode,
        platform.python_version(), platform.system(), platform.release()), ts=False)
    try:
        _log("argv: %s" % (" ".join(sys.argv),))
    except Exception:
        pass
    _log("=" * 76)


def out(msg):
    if not QUIET:
        print(msg)
    _log(msg, ts=True)


def _tagline(tag, msg):
    """One ALIGNED status line - the message column is identical for every
    tag on every platform (the old tags had 1/2/3 different space paddings)."""
    if is_windows():
        return "  %-7s %s" % ("[" + tag + "]", msg)
    m, color = {"ok": ("v", "32"), "warn": ("!", "33"),
                "step": ("~", "36"), "note": ("i", "36")}.get(tag, ("~", "36"))
    return "   \x1b[%sm%s\x1b[0m  %s" % (color, m, msg)


def ok(msg):
    _con = _tagline("ok", msg)
    if not QUIET:
        print(_con)
    _log(msg, ts=True)


def warn(msg):
    # warnings always reach the console, even in a quiet phase; when a phase
    # line is pending (dots without the status yet), close it first
    global PHASE_DIRTY
    if PHASE_DIRTY:
        print("")
        PHASE_DIRTY = False
    print(_tagline("warn", msg))
    _log(msg, ts=True)


def step(msg):
    _con = _tagline("step", msg)
    if not QUIET:
        print(_con)
    _log(msg, ts=True)


def die(msg, code=1):
    print(("  [FAIL] " if is_windows() else "   \x1b[31mx\x1b[0m ") + msg)
    _log(msg, ts=True)
    try:
        if LOG_HANDLE:
            LOG_HANDLE.close()
    except Exception:
        pass
    sys.exit(code)


def phase(idx, total, label, body):
    """One console line per phase: '  [3/6] label ..... status'. The body runs
    with QUIET=True (tool chatter goes to install.log); warn() still shows.
    body() returns the short status string."""
    sys.stdout.write("  [%d/%d] %s " % (idx, total, label))
    pad = 46 - len(label)
    if pad > 3:
        sys.stdout.write("." * pad + " ")
        sys.stdout.flush()
    global QUIET, PHASE_DIRTY
    QUIET = True
    PHASE_DIRTY = True
    try:
        status = body()
    finally:
        QUIET = False
        PHASE_DIRTY = False
    print(status or "done")
    _log("phase %d/%d %s -> %s" % (idx, total, label, status), ts=True)
    return status


def note(msg):
    _log(msg, ts=True)


# ----------------------------------------------------------- ui toolkit --- #
class MenuAbort(Exception):
    """The user aborted a sub-flow (add-in selection 'q') - the interactive
    main menu catches it and re-shows the menu instead of exiting."""


ANSI = {
    "bold": "\x1b[1m", "dim": "\x1b[2m", "rev": "\x1b[7m",
    "red": "\x1b[31m", "green": "\x1b[32m", "yellow": "\x1b[33m",
    "cyan": "\x1b[36m", "magenta": "\x1b[35m",
    "reset": "\x1b[0m",
}


def _paint(txt, *styles):
    """ANSI-painted text on a real terminal; plain text when piped. Standard
    colors (cyan/red/green/yellow/magenta) + bold/dim/reverse all read fine
    on dark AND light terminal palettes (reverse inverts the user's own
    foreground/background, so the cursor row works on every theme)."""
    if styles and ("rev" in styles):
        # reverse covers both themes on its own - no extra color on top
        styles = ("rev",)
    if sys.stdout.isatty() and styles:
        return "".join(ANSI[s] for s in styles) + txt + ANSI["reset"]
    return txt


def _uitty():
    try:
        return sys.stdin.isatty() and sys.stdout.isatty() and not DRY_RUN
    except Exception:
        return False


def _read_key():
    """One keypress: 'up' | 'down' | 'enter' | 'space' | a character.
    None = the terminal does not support raw reads (caller falls back to the
    typed numeric prompt)."""
    try:
        if is_windows():
            import msvcrt
            ch = msvcrt.getwch()
            if ch in ("\x00", "\xe0"):
                ch2 = msvcrt.getwch()
                return {"H": "up", "P": "down"}.get(ch2, ch2)
            if ch in ("\r", "\n"):
                return "enter"
            if ch == "\x03":
                raise KeyboardInterrupt
            return ch
        import termios, tty
        fd = sys.stdin.fileno()
        old = termios.tcgetattr(fd)
        try:
            tty.setcbreak(fd)
            ch = sys.stdin.read(1)
            if ch == "\x1b":
                nxt = sys.stdin.read(1)
                if nxt == "[":
                    return {"A": "up", "B": "down"}.get(sys.stdin.read(1), None)
                return "esc"
            if ch in ("\r", "\n"):
                return "enter"
            if ch == "\x03":
                raise KeyboardInterrupt
            return ch
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old)
    except KeyboardInterrupt:
        raise
    except Exception:
        return None


def menu_pick(rows, start=0, hint=""):
    """Arrow-driven menu (tty only; returns None otherwise, and the caller
    prints the options + reads a line). rows = (action, label, style) where
    a None action + '@SEP' label is a visual separator. Arrow up/down move
    (skipping separators), Enter selects the cursor row, a typed key selects
    the row with that action, 'q'/Esc select the '0' row when present."""
    if not _uitty():
        return None
    if is_windows():
        os.system("")   # enables ANSI/VT processing on modern consoles
    sels = [i for i, r in enumerate(rows) if r[1] != "@SEP"]
    if not sels:
        return None
    idx = sels[min(start, len(sels) - 1)]

    def draw(first):
        if not first:
            sys.stdout.write("\x1b[%dA\r\x1b[J" % (len(rows) + 2))
        for i, (act, label, style) in enumerate(rows):
            if label == "@SEP":
                print("")
                continue
            body = ("%s) %s" % (act, label)) if act else ("    " + label)
            if i == idx:
                print(_paint(" > " + body, "rev"))
            else:
                print("   " + ((_paint(body, style)) if style else body))
        print("")
        print("   " + ((_paint(hint, "dim")) if hint else ""))

    draw(True)
    while True:
        k = _read_key()
        if k is None or k == "":
            return None          # raw reads gone / EOF (Ctrl-D)
        if k == "up":
            j = idx
            while True:
                j = (j - 1) % len(rows)
                if rows[j][1] != "@SEP":
                    idx = j
                    break
            draw(False)
        elif k == "down":
            j = idx
            while True:
                j = (j + 1) % len(rows)
                if rows[j][1] != "@SEP":
                    idx = j
                    break
            draw(False)
        elif k == "enter":
            return idx
        elif k in ("esc", "q", "Q"):
            for i, (act, _lbl, _st) in enumerate(rows):
                if act == "0":
                    return i
            return None
        elif isinstance(k, str) and k.strip():
            for i, (act, _lbl, _st) in enumerate(rows):
                if act and act == k:
                    return i
            draw(False)      # unknown key: ignore, keep the arrow menu


def toggle_pick(rows, hint="", start=0, pre_lines=None):
    """Arrow-driven toggle list for select_addins. rows = (act, checked,
    label, extra, style): a toggle row has checked True/False; a plain action
    row has checked None (act '4' = proceed/install, act '0' = back).
    Returns ('toggle', index) | ('done', None) | ('back', None); None = no
    raw input available (typed fallback). Every draw REPAINTS the whole
    region (cursor home + clear below + pre_lines + rows) - relative erases
    would duplicate rows once something prints between interactions."""
    if not _uitty():
        return None
    if is_windows():
        os.system("")
    idx = min(start, len(rows) - 1)

    def draw():
        sys.stdout.write("\x1b[H\x1b[J")
        for ln in (pre_lines or []):
            print(ln)
        for i, (act, checked, label, extra, style) in enumerate(rows):
            if label == "@SEP":                    # visual separator
                print("")
                continue
            if checked is None:                    # plain action row
                body = "%s) %s" % (act, label)
                if i == idx:
                    print(_paint(" > " + body, "rev"))
                else:
                    print("   " + ((_paint(body, style)) if style else body))
                continue
            box = "[x]" if checked else "[ ]"
            if i == idx:
                print(_paint(" > %s %s%s" % (box, label, (extra or "")), "rev"))
            else:
                line = "   " + _paint(box, "green" if checked else "dim") + " " + label
                if extra:
                    line += " " + ((_paint(extra, style or "dim")) if sys.stdout.isatty() else extra)
                print(line)
        print("")
        print("   " + ((_paint(hint, "dim")) if hint else ""))
        sys.stdout.flush()

    def act_row(kind):
        j = next((i for i, r in enumerate(rows)
                  if r[0] and r[1] is None and r[0].lower() == kind), None)
        return j

    draw()
    while True:
        k = _read_key()
        if k is None or k == "":
            return None                  # raw reads gone / EOF (Ctrl-D)
        if k == "up":
            j = idx
            while True:
                j = (j - 1) % len(rows)
                if rows[j][2] != "@SEP":
                    break
            idx = j
            draw()
        elif k == "down":
            j = idx
            while True:
                j = (j + 1) % len(rows)
                if rows[j][2] != "@SEP":
                    break
            idx = j
            draw()
        elif k in ("space", " ", "s"):
            if rows[idx][1] is not None:
                return ("toggle", idx)
            draw()
        elif isinstance(k, str) and k.strip() and \
                any(r[0] and r[0].lower() == k for r in rows):
            j = act_row(k)
            if j is not None:                      # action row: 4 = proceed, 0 = back
                return ("done", None) if rows[j][0] == "4" else ("back", None)
            j = next((i for i, r in enumerate(rows)
                      if r[0] == k and r[1] is not None), None)
            if j is None:
                draw()
                continue
            return ("toggle", j)
        elif k == "enter":
            act, checked = rows[idx][0], rows[idx][1]
            if checked is None:
                if act == "4":
                    return ("done", None)
                return ("back", None)
            return ("toggle", idx)
        elif k in ("q", "Q", "esc", "\x1b"):
            return ("back", None)
        else:
            draw()


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

MAILKICK_OPT = False      # --mailkick: non-interactive opt-in for Outlook + the CA path
REPO_URL = "https://github.com/corelan/officekick"


def detect_installed():
    """Add-ins already present from a previous install (deployed folder or
    rendered manifest). Update flows keep whatever is installed."""
    inst = install_dir()
    found = set()
    for (folder, _name, mf) in OURS:
        if os.path.isdir(os.path.join(inst, folder)) or \
                os.path.isfile(os.path.join(manifest_dir(), mf)):
            found.add(folder)
    return found


def https_feasible():
    """Would setup of the https (loopback TLS) component succeed on this box?
    Returns (feasible, [missing-reasons], [techniques-used]). Outlook on
    macOS needs it; on Windows Outlook rides plain http, so no TLS is
    required there."""
    if is_mac():
        reasons = []
        techniques = []
        ov = shutil.which("openssl")
        if ov is None:
            reasons.append("openssl not found - it generates the local"
                           " certificate (leaf, or the CA-signed pair)")
        else:
            _rc, over = run_quiet(["openssl", "version"])
            techniques.append("%s (%s) - generates the self-signed certificate"
                              " (leaf, or CA + signed leaf)" % (
                                  os.path.basename(ov), (over or "").strip() or "openssl"))
        if shutil.which("security") is None:
            reasons.append("the macOS 'security' tool is missing - it imports "
                           "+ trusts the certificate in your keychains")
        else:
            techniques.append("the macOS 'security' tool - imports the CA into"
                              " your login keychain and marks it trusted")
        try:
            import ssl  # noqa: F401
            techniques.append("python ssl - serves the https://localhost:%d"
                              " pane listener for Outlook" % (HTTPS_PORT if HTTPS_PORT else PORT + 1))
        except Exception as exc:
            reasons.append("python ssl module unavailable: %s" % exc)
        return (len(reasons) == 0, reasons, techniques)
    # Windows: TLS only comes into play when the stored transport is https
    if _current_transport() != "https":
        return (True, [], [])   # plain http - no TLS component for Outlook
    reasons, techniques = [], []
    try:
        import cryptography  # noqa: F401
        _ver = getattr(cryptography, "__version__", "")
        techniques.append("python cryptography%s - generates the self-signed"
                          " leaf (installer auto-installs it if missing)" %
                          ((" " + _ver) if _ver else ""))
    except ImportError:
        reasons.append("the 'cryptography' package is missing - the startup"
                       " install failed (see install.log); re-run install.py or run:"
                       " python -m pip install --user cryptography")
    try:
        import ssl  # noqa: F401
        techniques.append("python ssl - serves the https://localhost:%d pane"
                          " listener" % (HTTPS_PORT if HTTPS_PORT else PORT + 1))
    except Exception as exc:
        reasons.append("python ssl module unavailable: %s" % exc)
    return (len(reasons) == 0, reasons, techniques)


_CFG_CACHE = {"t": 0.0, "d": None}


_RUNTIME_DEPS_DONE = False


def ensure_runtime_deps():
    """Install python dependencies AT STARTUP - the user rule: the installer
    never defers a dependency it can fetch itself (the add-in picker must
    not lock Outlook over a package the installer can install in seconds).
    Idempotent; a no-op when the imports already work. (0.9.275)"""
    global _RUNTIME_DEPS_DONE
    if _RUNTIME_DEPS_DONE or DRY_RUN or not is_windows():
        return True
    _RUNTIME_DEPS_DONE = True
    try:
        import cryptography  # noqa: F401
        return True
    except ImportError:
        pass
    step("installing the python 'cryptography' package (needed for the"
         " local https certificates; pip --user, no admin)...")
    import importlib
    # pip itself may be absent from a managed runtime - bootstrap via ensurepip
    rc, o = run_quiet([sys.executable, "-m", "pip", "--version"], timeout=30)
    if rc != 0:
        step("pip is missing from this Python - bootstrapping it (ensurepip)...")
        rc, o = run_quiet([sys.executable, "-m", "ensurepip", "--upgrade",
                           "--default-pip"], timeout=120)
        _log("ensurepip: rc %d out=%s" % (rc, (o or "")[-200:]), ts=True)
        if rc != 0:
            warn("pip could not be bootstrapped (see install.log) - run:"
                 " python -m ensurepip --upgrade")
    # wheel-only FIRST: a source build needs a Rust toolchain and fails late
    # (Python 3.15-era runtimes may not have cryptography wheels yet)
    ok_done = False
    last_out = ""
    for pip_args in (["--user", "--only-binary=:all:"], ["--only-binary=:all:"],
                     ["--user"], []):
        rc, o = run_quiet([sys.executable, "-m", "pip", "install"] +
                          pip_args + ["cryptography"], timeout=600)
        last_out = (o or "")
        _log("startup pip install %scryptography: rc %d out=%s"
             % ((" ".join(pip_args) + " ") if pip_args else "", rc,
                last_out[-220:]), ts=True)
        if rc != 0:
            # surface WHY on the console: no wheel / no pip / TLS / proxy
            tail = " ".join(last_out.split())[-180:]
            if tail:
                warn("pip: %s" % tail)
        if rc == 0:
            importlib.invalidate_caches()
            try:
                import cryptography  # noqa: F401
                ok_done = True
                break
            except ImportError:
                continue
    if ok_done:
        ok("the 'cryptography' package is installed")
    else:
        warn("could not install 'cryptography' at startup - Outlook stays"
             " locked; run: python -m pip install --user cryptography"
             " (the reason is in install.log)")
        _log("startup dependency install failed (cryptography, Windows)")
    return ok_done


def _config_settings(max_age=2.0):
    """Settings from the running server first; keys the server does not
    know (older builds) fall back to the config file; {} when neither is
    reachable. Cached ~2 s: the menus call this four times per draw, and on
    some Windows boxes a refused-or-filtered loopback connect waits the full
    socket timeout - four of those per redraw read as a slow menu (0.9.274)."""
    import time as _time
    cached = _CFG_CACHE["d"]
    if cached is not None and (_time.time() - _CFG_CACHE["t"]) < max_age:
        return cached
    srv = None
    try:
        import urllib.request
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open("http://127.0.0.1:%d/api/config" % PORT, timeout=0.6) as resp:
            d = json.loads(resp.read().decode("utf-8", "replace")) or {}
            srv = {k: d.get(k) for k in ("transportMode", "transportHttpPort",
                                         "transportHttpsPort", "certMode", "certHost")}
    except Exception:
        pass
    if srv is not None:
        try:
            with open(config_file(), "r", encoding="utf-8") as fh:
                d2 = json.load(fh) or {}
        except Exception:
            d2 = {}
        # The config FILE is the source of truth for installer-owned settings;
        # a running server mirrors it. When they disagree, the file wins and
        # the divergence is logged (the server re-reads at its next restart).
        out = {}
        for k in ("transportMode", "transportHttpPort",
                  "transportHttpsPort", "certMode", "certHost"):
            file_v = d2.get(k)
            out[k] = file_v if file_v is not None else srv.get(k)
            if srv.get(k) not in (None, out[k]) and file_v is not None:
                _log("settings divergence on %s: server=%r file=%r - the file"
                     " wins (the server re-reads at its next restart)" % (k, srv.get(k), out[k]),
                     ts=True)
        _CFG_CACHE["t"] = _time.time()
        _CFG_CACHE["d"] = out
        return out
    try:
        with open(config_file(), "r", encoding="utf-8") as fh:
            d = json.load(fh) or {}
        out = {k: d.get(k) for k in ("transportMode", "transportHttpPort",
                                     "transportHttpsPort", "certMode", "certHost")}
        _CFG_CACHE["t"] = _time.time()
        _CFG_CACHE["d"] = out
        return out
    except Exception:
        out = {}
        _CFG_CACHE["t"] = _time.time()
        _CFG_CACHE["d"] = out
        return out


def _valid_port(v):
    try:
        return 1024 <= int(v) <= 65535
    except (TypeError, ValueError):
        return False


def _current_transport():
    """The persisted transport mode (auto/http/https); 'auto' when unknown."""
    mode = _config_settings().get("transportMode")
    return mode if mode in ("auto", "http", "https") else "auto"


def _current_ports():
    """(http_port, https_port) from the persisted settings; the classic
    8765/8766 when unknown."""
    d = _config_settings()
    hp = d.get("transportHttpPort")
    sp = d.get("transportHttpsPort")
    hp = int(hp) if _valid_port(hp) else 8765
    sp = int(sp) if _valid_port(sp) and int(sp) != hp else hp + 1
    return hp, sp


def _persist_file(pairs):
    """Merge key/value pairs into the config FILE - the single source of
    truth for installer-owned settings. The running server mirrors them via
    _api_set; if the two ever disagree, the file wins (read side)."""
    cfg = {}
    try:
        with open(config_file(), "r", encoding="utf-8") as fh:
            cfg = json.load(fh) or {}
    except Exception:
        cfg = {}
    for k, v in pairs.items():
        cfg[k] = v
    _cdir = os.path.dirname(config_file())
    if _cdir and not os.path.isdir(_cdir):
        os.makedirs(_cdir, exist_ok=True)   # a fresh machine: no ~/.officekick yet
    _CFG_CACHE["d"] = None   # a write invalidates the read cache
    with open(config_file(), "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2)
    try:
        _protect_secret(config_file())
    except OSError:
        pass
    return True


def _api_set(body):
    """Persist settings through the running server and VERIFY it actually
    stored them (older builds silently ignore unknown keys) - returns True
    only when confirmed; the caller falls back to the config file; when the
    server is down, writes the config file directly."""
    try:
        import urllib.request
        req = urllib.request.Request(
            "http://127.0.0.1:%d/api/settings" % PORT,
            data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json"}, method="PUT")
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open(req, timeout=3) as resp:
            resp.read()
        _CFG_CACHE["d"] = None   # verify against the post-PUT state
        got = _config_settings()
        for k, v in body.items():
            if got.get(k) != v:
                note("the running server did not store %s (has %r) - config-file fallback"
                     % (k, got.get(k)))
                return False
        note("settings persisted via the running server: %s" % json.dumps(body))
        return True
    except Exception as exc:
        note("persisting %s via the server failed: %s" % (json.dumps(body), exc))
        return False


def _set_transport(mode, http_port=None, https_port=None):
    """Persist the transport mode + optional listener ports: through the
    running server (settings PUT) so both sides agree, or straight into the
    config file when the server is down."""
    if mode not in ("auto", "http", "https"):
        return False
    body = {"transportMode": mode}
    if http_port is not None:
        body["transportHttpPort"] = int(http_port)
    if https_port is not None:
        body["transportHttpsPort"] = int(https_port)
    _persist_file(body)
    _api_set(body)          # mirror to a RUNNING server (best effort)
    note("transport persisted (config file): %s" % json.dumps(body))
    return True
def _protect_secret(path, writable=True):
    """chmod 0600 on macOS/POSIX. On Windows os.chmod is a stub (only the
    read-only bit), so the real protection is NTFS ACLs via icacls: strip
    inheritance and grant ONLY the current user (full access for a private
    key / config, read-only would break writes)."""
    try:
        if is_windows():
            user = os.environ.get("USERNAME") or "SYSTEM"
            perm = "F" if writable else "R"
            rc1, _o = run_quiet(["icacls", path, "/inheritance:r"])
            rc2, _o = run_quiet(["icacls", path, "/grant:r", "%s:%s" % (user, perm)])
            if rc1 == 0 and rc2 == 0:
                _log("protected (NTFS ACL, user-only) %s" % path)
                return True
            warn("icacls could not harden %s (rc %d/%d) - check it manually"
                 % (path, rc1, rc2))
            return False
        os.chmod(path, 0o600 if writable else 0o400)
        return True
    except OSError as exc:
        warn("could not protect %s: %s" % (path, exc))
        return False


def clear_screen():
    """Wipe the terminal and put the cursor at the top-left (tty only; a
    no-op for piped/automation runs and windowless scheduled tasks)."""
    try:
        if sys.stdin.isatty() and sys.stdout.isatty():
            if is_windows():
                os.system("")   # enable ANSI/VT processing on modern consoles
            sys.stdout.write("\x1b[2J\x1b[H")
            sys.stdout.flush()
    except Exception:
        pass


def _https_state():
    """(active, description) for the https://localhost:<port+1> listener in
    the info header above menus. Never raises."""
    port_tls = HTTPS_PORT if HTTPS_PORT else PORT + 1
    try:
        import ssl
        import urllib.request
        ctx = ssl._create_unverified_context()   # the loopback pair is ours; no CA dance here
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}),
                                             urllib.request.HTTPSHandler(context=ctx))
        with opener.open("https://localhost:%d/api/health" % port_tls, timeout=1.2) as resp:
            d = json.loads(resp.read().decode("utf-8", "replace"))
        if d.get("ok") is True:
            return True, "https listener active on https://localhost:%d (version %s)" % (
                port_tls, d.get("version"))
        return False, "https listener on port %d answers but is unhealthy" % port_tls
    except Exception:
        return False, "https listener not running on port %d (transport mode: %s)" % (
            port_tls, _current_transport())


def banner(mode=None):
    clear_screen()
    head = "%s v%s  -  installer" % (APP_NAME, VERSION)
    print("")
    print(_paint("-" * 74, "cyan"))
    print("  " + _paint(head, "bold"))
    print("  " + _paint("Created by Corelan Consulting bv", "cyan"))
    print("  " + _paint(REPO_URL, "dim"))
    print(_paint("-" * 74, "cyan"))
    h = health(PORT)
    inst = detect_installed()
    tls_on, tls_desc = _https_state()
    print("  detected : %s %s | python %s" % (
        platform.system(), platform.release(), platform.python_version()))
    print("  install  : %s" % install_dir())
    if h:
        print("  server   : running (version %s)" % h.get("version"))
        print("        http  : http://127.0.0.1:%d (always on - heartbeat for install/update/restart)" % PORT)
        print("        https : %s" % tls_desc)
    else:
        print("  server   : not running (start it here or: install.py --start)")
    print("  add-ins  : %s" % (", ".join(name for (folder, name, _mf) in OURS
                                        if folder in inst) or "not installed yet"))
    if mode:
        note("banner shown (mode=%s, server=%s, add-ins=%s)" % (mode, h is not None, ",".join(sorted(inst)) or "-"))


# ------------------------------------------------------ add-in selection ----
def select_addins(interactive):
    """Decide which add-ins this run deploys. Word + PowerPoint default ON;
    Outlook default OFF. Returns the set (also stored in ADDINS)."""
    global ADDINS, FIRST_RUN
    base = {"slidekick", "wordkick"}
    prev = detect_installed()
    FIRST_RUN = set()
    if not interactive or DRY_RUN:
        addoutlook = ((MAILKICK_OPT or bool(MAILKICK_CERT or MAILKICK_KEY)
                       or "mailkick" in prev) and not SKIP_MAILKICK)
        ADDINS = base | ({"mailkick"} if addoutlook else set())
        note("add-ins selected (non-interactive): %s%s" % (
            ", ".join(sorted(ADDINS)),
            " (outlook: previously installed)" if "mailkick" in prev and not (MAILKICK_OPT or MAILKICK_CERT or MAILKICK_KEY) else ""))
        ok("installing: " + (", ".join(app_label(a) for a in sorted(ADDINS))))
        FIRST_RUN = set(ADDINS) - prev
        return set(ADDINS)
    rows = [("wordkick",  "Word (DocKick)",       "wordkick"  in prev),
            ("slidekick", "PowerPoint (SlideKick)", "slidekick" in prev),
            ("mailkick",  "Outlook (MailKick)",     "mailkick"  in prev)]
    state = {}
    for (folder, _label, was) in rows:
        state[folder] = (True if folder != "mailkick" else False) if not was else True
    feas, feas_why, feas_how = https_feasible()
    clear_screen()

    def _prereq_lines():
        """The https status block (list of lines, no painting side effects)."""
        out_lines = []
        if not is_mac() and _current_transport() != "https":
            out_lines.append(_paint("  https prerequisites for Outlook: ", "bold")
                             + "not needed")
            out_lines.append("")
            out_lines.append("  Windows rides the plain-http listener.")
            out_lines.append("  No certificate is created or trusted.")
            return out_lines
        # Windows + https transport: the same prerequisite rendering as mac
        if feas:
            out_lines.append("  " + _paint("https prerequisites for Outlook: MET",
                                           "green"))
        else:
            out_lines.append("  " + _paint("https prerequisites for Outlook: NOT MET",
                                           "red") + _paint(" - Outlook is locked",
                                                           "bold"))
        out_lines.append("")
        if feas:
            for t in feas_how:
                if "security tool" in t or "'security'" in t:
                    continue   # mac-only technique
                out_lines.append("    - " + t.replace(
                    " (installer auto-installs it if missing)", ""))
        else:
            for rln in feas_why:
                out_lines.append("    - " + rln)
        out_lines.append("")
        out_lines.append(_paint("  Enabling Outlook asks for the certificate mode"
                                " before install.", "dim"))
        out_lines.append(_paint("  Self-signed leaf is the default."
                                " Change it in Transport options.", "dim"))
        out_lines.append("")
        return out_lines

    msg = []   # picker feedback lines (repainted with the screen, never duplicated)
    arrow_done = False
    drew_picker = False
    cur = 0
    if _uitty():
        order = [f for (f, _l, _w) in rows]   # picker index -> folder

        def _rowdata():
            r = []
            # OPTIONS (toggles) carry the digit keys the hints advertise
            # (1/2/3); the ACTION lives in its own group; navigation back in
            # its own group - information (https status) is a block, not a row
            for (n, (folder, label, _was)) in zip(("1", "2", "3"), rows):
                r.append((n, state[folder], label, "", None))
            r.append((None, None, "@SEP", "", None))
            r.append(("4", None, "Install the selected components", "", "cyan"))
            r.append((None, None, "@SEP", "", None))
            r.append(("0", None, "back (nothing is installed yet)", "", "magenta"))
            return r
        while True:
            pre = ["",
                   "  Add-ins - Word + PowerPoint on by default, Outlook is not:",
                   ""] + _prereq_lines() + ([""] if msg else []) + msg
            r = toggle_pick(_rowdata(),
                            "(up/down = move, space/Enter or 1-3 = toggle, 4 = install, 0 = back)",
                            start=cur, pre_lines=pre)
            drew_picker = True
            if r is None:
                break            # raw reads vanished - typed loop takes over
            kind, val = r
            cur = val if kind == "toggle" else cur
            _log("add-in picker: %s %s" % (kind, val), ts=True)
            if kind == "back":
                raise MenuAbort("add-in selection cancelled")
            if kind == "done":
                arrow_done = True
                break
            folder = order[val]
            if folder == "mailkick" and not state[folder] and not feas:
                msg = ["   !  Outlook cannot be enabled on this machine:"] + \
                      ["        - " + rr for rr in feas_why]
                note("outlook enable refused: %s" % "; ".join(feas_why))
                continue
            if folder == "mailkick" and not state[folder] and is_mac():
                print("")
                print("  " + _paint("Outlook needs the local https certificate setup.", "bold"))
                print("    1) self-signed leaf  - DEFAULT; no CA is created:")
                print("                          the LEAF itself is imported + trusted")
                print("                          in your keychain")
                print("    2) loopback CA       - the CA is imported + trusted;")
                print("                          renewals stay silent afterwards")
                print("    3) bring-your-own    - your own already-trusted PEM pair")
                print("  the mode is asked right before the certificates step;")
                print("  change it later in Transport options.")
                go = ask_yes("Enable Outlook (choose the certificate mode during"
                             " install)", False)
                note("outlook CA warning answer: %s" % ("yes" if go else "no"))
                if not go:
                    state[folder] = False
                    msg = ["   - Outlook certificate setup declined - MailKick stays off."]
                    continue
                msg = ["   - Outlook enabled.",
                       "   - The certificate question comes before step 3 of 6.",
                       "   - Self-signed leaf is the default answer."]
                state[folder] = not state[folder]
                continue
            state[folder] = not state[folder]
            msg = []
    while arrow_done is False:
        if not drew_picker:
            print("")
            print("  Add-ins (Word + PowerPoint are enabled by default, Outlook is not):")
            print("")
            for pl in _prereq_lines():
                print(pl)
            print("")
            drew_picker = True
        for (folder, label, _was) in rows:
            box = "[x]" if state[folder] else "[ ]"
            print("    %d) %-24s %s" % (rows.index((folder, label, _was)) + 1, label, box))
        print("")
        print("    4) Install the selected components")
        print("")
        print("    0) back (nothing is installed yet)")
        print("")
        choice = input("  > toggle 1/2/3, 4 = install the selected, 0 = back: ").strip().lower()
        _log("add-in selection input: %r" % choice, ts=True)
        if choice in ("q", "quit", "abort", "0", "back"):
            print("  add-in selection cancelled - back to the menu.")
            note("add-in selection cancelled by user")
            raise MenuAbort("add-in selection cancelled (typed)")
        if choice == "" or choice in ("4", "install"):
            break
        if choice in ("1", "2", "3"):
            folder = rows[int(choice) - 1][0]
            if folder == "mailkick" and not state[folder] and not feas:
                warn("Outlook cannot be enabled on this machine:" if is_mac() else
                     "Outlook cannot be enabled here:")
                for r in feas_why:
                    print("      - " + r)
                note("outlook enable refused: %s" % "; ".join(feas_why))
                continue
            if folder == "mailkick" and not state[folder] and is_mac():
                print("")
                print("  -------------------------------------------------------------")
                print("   OUTLOOK (MailKick) - one-time certificate setup")
                print("   Enabling this will CREATE a SELF-SIGNED certificate and a")
                print("   local CA for localhost and import them into your keychains.")
                print("   macOS will ask to trust the CA - accept to let the pane")
                print("   load. Nothing else is installed; removed again by uninstall.")
                print("  -------------------------------------------------------------")
                print("")
                go = ask_yes("Continue with the certificate setup for Outlook", False)
                note("outlook CA warning answer: %s" % ("yes" if go else "no"))
                if not go:
                    state[folder] = False
                    continue
            state[folder] = not state[folder]
        else:
            print("  choose 1, 2, 3, empty to continue or q to abort")
    ADDINS = {f for (f, _l, _w) in rows if state[f]}
    if not ADDINS:
        print("  nothing selected - back to the menu.")
        note("add-in selection: nothing selected")
        raise MenuAbort("add-in selection empty")
    ok("installing: " + (", ".join(app_label(a) for a in sorted(ADDINS))))
    _log("add-ins selected: %s" % ", ".join(sorted(ADDINS)), ts=True)
    FIRST_RUN = set(ADDINS) - prev
    return set(ADDINS)


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
    """Colored question on its own left-aligned line, printed BEFORE the
    action it belongs to is taken."""
    if DRY_RUN:
        note("question (dry-run): %s -> default %s" % (prompt, "yes" if default_yes else "no"))
        return default_yes
    try:
        if not sys.stdin.isatty():
            note("question (non-tty): %s -> default %s" % (prompt, "yes" if default_yes else "no"))
            return default_yes   # piped/automated runs take the default silently
    except (OSError, ValueError, AttributeError):
        return default_yes
    suffix = "[Y/n]" if default_yes else "[y/N]"
    while True:
        try:
            ans = input("  " + _paint("? " + prompt + " ", "cyan") + _paint(suffix, "dim") + ": ").strip().lower()
        except EOFError:
            note("question EOF: %s -> default %s" % (prompt, "yes" if default_yes else "no"))
            return default_yes
        except KeyboardInterrupt:
            out("\n  aborted")
            raise SystemExit(130)
        if not ans:
            note("question: %s -> default %s" % (prompt, "yes" if default_yes else "no"))
            return default_yes
        if ans in ("y", "yes"):
            note("question: %s -> yes" % prompt)
            return True
        if ans in ("n", "no"):
            note("question: %s -> no" % prompt)
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


def _hard_task_settings(name):
    """schtasks default settings (stop-on-idle, stop-on-batteries, 72h
    execution limit) end long-running servers silently; every install
    re-applies the hard settings so the server runs until shutdown."""
    if DRY_RUN:
        return
    helper = (
        "$d = Get-ScheduledTask -TaskName '%s' -ErrorAction Stop; "
        "$d.Settings.DisallowStartIfOnBatteries = $false; "
        "$d.Settings.StopIfGoingOnBatteries = $false; "
        "$d.Settings.ExecutionTimeLimit = 'PT0S'; "
        "if ($d.Settings.IdleSettings) { $d.Settings.IdleSettings.StopOnIdleEnd = $false; }; "
        "Set-ScheduledTask -InputObject $d | Out-Null" % name)
    ps = shutil.which("powershell") or "powershell"
    rc, o = run_quiet([ps, "-NoProfile", "-Command", helper])
    if rc == 0:
        ok("task hardened: no idle/battery stop, no execution time limit (%s)" % name)
    else:
        warn("task hardening failed (%s)" % (o or "rc %d" % rc))


def start_only():
    """--start: bring the local server up (no install steps) and health-check"""
    out("")
    out("  Starting the OfficeKick server")
    install_dir()   # just for the message below
    if health(PORT):
        h = health(PORT)
        ok("already running: http://127.0.0.1:%d (version %s)"
           % (PORT, h.get("version")))
        rv = read_version()
        hv = (h.get("version") or "").strip()
        if rv and hv and hv != rv:
            warn("the running build (%s) differs from this checkout (%s) - "
                 "run install.py --update" % (hv, rv))
        check_pane_errors()
        return
    stop_server()
    if start_server():
        return
    die("could not start the server - run install.py --install")

def start_server():
    if DRY_RUN:
        step("[dry-run] would start the server on port %d" % PORT)
        return True
    if not os.path.isdir(install_dir()):
        # self-heal: the deployed folder is missing (uninstalled or wiped) -
        # deploy the SERVER first, otherwise Popen cwd raises WinError 267
        warn("the deployed server folder is missing - deploying the server now")
        deploy_files(set())   # server only
        ok("server deployed to " + install_dir())
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


# ------------------------------------------------- MailKick/Outlook choice #
def tls_cert_dir():
    return os.path.join(config_dir(), "certs")


def tls_active():
    """True when a WORKING loopback certificate exists right now (used to
    decide what the installer should TELL the user, 0.9.205)."""
    _crt = os.path.join(tls_cert_dir(), "server.crt")
    _key = os.path.join(tls_cert_dir(), "server.key")
    return (os.path.isfile(_crt) and os.path.isfile(_key)
            and validate_existing_tls(_crt) is None)


def _sim_base():
    return "https://localhost:%d" % (HTTPS_PORT if HTTPS_PORT else PORT + 1) if tls_active() else "http://127.0.0.1:%d" % PORT


def pane_listener_line():
    """One line describing where the panes are served, honest about TLS."""
    if tls_active():
        return (
            "  https://localhost:%d  <- task panes + AI requests (TLS, user-trusted CA)\n"
            "  https://localhost:%d  <- same statics, TLS (Outlook)\"s requirement)"
            % (HTTPS_PORT if HTTPS_PORT else PORT + 1),
            "  http://127.0.0.1:%d   <- same statics, plain http" % (PORT,)
        )
    return "  http://127.0.0.1:%d   <- task panes + AI requests (no certificate)" % PORT

def validate_user_cert(crt, key):
    """Check a user-supplied MailKick TLS pair. Returns (ok, problems[])."""
    problems = []
    if not os.path.isfile(crt): problems.append("certificate file not found: %s" % crt)
    if not os.path.isfile(key): problems.append("key file not found: %s" % key)
    if problems: return False, problems
    rc, o = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-checkend", "0"])
    if rc != 0: problems.append("the certificate has expired or is not yet valid")
    rc, o = run_quiet(["openssl", "pkey", "-in", key, "-noout"])
    if rc != 0: problems.append("the private key cannot be read (wrong file or passphrase-protected? the key must be unencrypted PEM)")
    san = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-text"])[1]
    if "DNS:localhost" not in san and "IP Address:127.0.0.1" not in san:
        problems.append("the certificate must contain subjectAltName DNS:localhost (and ideally IP:127.0.0.1) - the pane is loaded from https://localhost")
    if "TLS Web Server Authentication" not in san and "serverAuth" not in san:
        warn("note: the certificate lacks the serverAuth extended key usage (usually harmless for localhost, but stricter clients may care)")
    rc, o = run_quiet(["security", "verify-cert", "-c", crt, "-p", "ssl"])
    if rc != 0:
        problems.append("the certificate chain is NOT trusted by this Mac (security verify-cert failed: %s)" %
                        (o or "untrusted").strip()[:140] +
                        " - import the issuing CA into a keychain first (Keychain Access or security add-trusted-cert)")
    return len(problems) == 0, problems


def pick_mailkick_route():
    """Non-interactive TLS route fallback for MailKick (automation): a
    user-supplied certificate (--mailkick-cert/key) validates -> 'custom';
    anything else -> 'auto' (reuse-first applies). The interactive
    leaf/CA/BYO question lives in _ask_cert_setup()."""
    if MAILKICK_CERT or MAILKICK_KEY:
        good2, problems = validate_user_cert(MAILKICK_CERT, MAILKICK_KEY)
        if good2:
            return "custom"
        for pr in problems:
            warn("provided MailKick certificate: " + pr)
        return "skip"
    return "auto"


def validate_existing_tls(crt, host=None):
    """Check the cert ALREADY on disk: unexpired, SAN (localhost, or a BYO
    hostname), trusted. None = usable; a short reason string = regenerate.
    Windows reads the file with cryptography (no openssl there) and checks
    the CURRENT USER Trusted People store by thumbprint."""
    host = host or _current_cert_host()
    rc, _o = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-checkend", "0"])
    if rc != 0:
        cert = _load_cert_from_pem(crt)
        if cert is None:
            return "the certificate could not be read"
        from datetime import datetime, timezone
        end = getattr(cert, "not_valid_after_utc", None)
        if end is None:
            end = cert.not_valid_after.replace(tzinfo=timezone.utc)
        if end <= datetime.now(timezone.utc):
            return "expired or not yet valid"
        try:
            san = cert.extensions.get_extension_for_class(
                __import__("cryptography").x509.SubjectAlternativeName).value
            dns = san.get_values_for_type(__import__("cryptography").x509.DNSName)
            ips = san.get_values_for_type(__import__("cryptography").x509.IPAddress)
            if host != "localhost":
                if host not in dns:
                    return "the certificate does not cover '%s' (no DNS SAN for it)" % host
            elif "localhost" not in dns and str(__import__("ipaddress").ip_address("127.0.0.1")) not in [str(i) for i in ips]:
                return "no localhost subjectAltName"
        except Exception:
            pass
    else:
        rc, o = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-text"])
        if host != "localhost":
            if ("DNS:%s" % host) not in o:
                return "the certificate does not cover '%s' (no DNS SAN for it)" % host
        elif "DNS:localhost" not in o and "IP Address:127.0.0.1" not in o:
            return "no localhost subjectAltName"
    if is_mac():
        rc, o = run_quiet(["security", "verify-cert", "-c", crt, "-p", "ssl"])
        if rc != 0:
            return "its chain is not trusted by this Mac (" + (o or "untrusted").strip()[:120] + ")"
        return None
    _sha256, sha1 = _leaf_fingerprints(crt)
    rc, o = run_quiet(["certutil", "-user", "-store", "TrustedPeople"])
    if rc != 0 or not sha1 or sha1 not in (o or "").replace(" ", "").replace(":", "").upper():
        return "the leaf is not trusted in the CURRENT USER Trusted People store"
    return None


CERT_FACTS = {}            # loopback TLS info, printed after the certificates phase


def _load_cert_from_pem(crt):
    """Read a PEM certificate with the 'cryptography' package (the WINDOWS
    leaf generator writes the file; openssl does not exist there)."""
    try:
        from cryptography import x509
        with open(crt, "rb") as fh:
            return x509.load_pem_x509_certificate(fh.read())
    except Exception:
        return None


def _cert_facts(crt):
    """Name + expiry of one PEM certificate ({} = unreadable)."""
    rc, o = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-subject", "-enddate"])
    if rc != 0:
        # WINDOWS: openssl is absent - read the leaf with cryptography
        try:
            from datetime import datetime, timezone
            cert = _load_cert_from_pem(crt)
            if cert is None:
                return {}
            cn = "?"
            for attr in cert.subject:
                if attr.oid.dotted_string == "2.5.4.3":   # CN
                    cn = attr.value
            end = getattr(cert, "not_valid_after_utc", None)
            if end is None:
                end = cert.not_valid_after.replace(tzinfo=timezone.utc)
            expires = end.strftime("%b %d %H:%M:%S %Y GMT")
            days = (end - datetime.now(timezone.utc)).days
            return {"cn": cn, "expires": expires, "days": days}
        except Exception:
            return {}
    cn = "?"
    m = re.search(r"CN ?= ?([^\n,]+)", o)
    if m:
        cn = m.group(1).strip()
    expires = "?"
    m = re.search(r"notAfter=(.+)", o)
    if m:
        expires = m.group(1).strip()
    days = None
    try:
        from datetime import datetime, timezone
        end = datetime.strptime(re.sub(r"\s+", " ", expires),
                                "%b %d %H:%M:%S %Y %Z").replace(tzinfo=timezone.utc)
        days = (end - datetime.now(timezone.utc)).days
    except Exception:
        pass
    return {"cn": cn, "expires": expires, "days": days}


def _fact_line(facts):
    if not facts:
        return "unreadable"
    more = (", %d days left" % facts["days"]) if isinstance(facts.get("days"), int) else ""
    return "CN=%s, expires %s%s" % (facts.get("cn", "?"), facts.get("expires", "?"), more)


def _print_cert_facts():
    """Console block after the certificates phase: names, expiry, file paths."""
    if not CERT_FACTS:
        return
    print("")
    print("  " + _paint("Loopback TLS (%s)" % CERT_FACTS["mode"], "bold"))
    print("    CA    : %s" % CERT_FACTS["ca_path"])
    print("            %s" % CERT_FACTS["ca_line"])
    print("    cert  : %s" % CERT_FACTS["leaf_path"])
    print("            %s" % CERT_FACTS["leaf_line"])
    print("    keys  : %s" % CERT_FACTS["key_line"])


def _pair_matches(crt, key):
    """True when the private key belongs to the certificate (modulus match)."""
    rc1, m1 = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-modulus"])
    rc2, m2 = run_quiet(["openssl", "rsa", "-in", key, "-noout", "-modulus"])
    if rc1 != 0 or rc2 != 0:
        return False
    return m1.strip() == m2.strip().replace("+", "").replace("*", "")


def _ca_healthy(ca_crt):
    """The loopback CA still exists, is not expired and is still trusted."""
    if not (os.path.isfile(ca_crt)):
        return False
    rc, _o = run_quiet(["openssl", "x509", "-in", ca_crt, "-noout", "-checkend", str(86400 * 60)])
    if rc != 0:
        return False
    rc, _o = run_quiet(["security", "verify-cert", "-c", ca_crt, "-p", "ssl"])
    return rc == 0


def _ca_sign(cert_dir, csr, crt):
    """Sign the leaf with the loopback CA; tolerate a re-sign of the same
    subject (unique_subject=no + serial bump + db rotation on refusal)."""
    attr = os.path.join(cert_dir, "index.txt.attr")
    with open(attr, "w") as fh:
        fh.write("unique_subject = no\n")
    cmd = ["openssl", "ca", "-batch", "-config", os.path.join(cert_dir, "ca.cnf"),
           "-startdate", "20180101000000Z", "-enddate", "20380101000000Z",
           "-in", csr, "-out", crt]
    rc, o = run_quiet(cmd)
    if rc != 0 and "already a certificate" in (o or ""):
        _log("ca db refuses a re-sign - rotating index.txt and bumping the serial, then retrying")
        idx = os.path.join(cert_dir, "index.txt")
        if os.path.isfile(idx):
            shutil.move(idx, idx + (".bak-%d" % (time.time_ns() % 1000000)))
        srl = os.path.join(cert_dir, "serial")
        try:
            snum = int(open(srl).read().strip(), 16)
        except Exception:
            snum = int(time.time())
        with open(srl, "w") as fh:
            fh.write("%04X" % (snum + 1))
        rc, o = run_quiet(cmd)
    return rc, o


def ensure_tls_certs():
    """Make sure a USABLE loopback pair exists for the https listener:
    reuse a valid one; reissue the leaf from the SAME already-trusted CA when
    only the leaf is bad/missing/expiring (< 30 days); otherwise generate a
    fresh CA + leaf + trust import. Returns (crt, key) or None; all facts go
    to CERT_FACTS for the console block after the certificates phase."""
    global CERT_FACTS
    CERT_FACTS = {}
    if not is_mac():
        return None
    if not shutil.which("openssl"):
        warn("TLS check skipped: openssl not found")
        return None
    cert_dir = os.path.join(config_dir(), "certs")
    crt = os.path.join(cert_dir, "server.crt")
    key = os.path.join(cert_dir, "server.key")
    ca_crt = os.path.join(cert_dir, "ca.crt")
    ca_key = os.path.join(cert_dir, "ca.key")
    os.makedirs(cert_dir, exist_ok=True)
    if is_windows():
        _protect_secret(cert_dir)          # NTFS ACL: user-only
    else:
        try:
            os.chmod(cert_dir, 0o700)
        except OSError:
            pass

    def _record(mode, extra_note=""):
        leaf = _cert_facts(crt)
        ca = _cert_facts(ca_crt)
        CERT_FACTS.update({
            "mode": mode,
            "ca_path": ca_crt,
            "ca_line": _fact_line(ca),
            "leaf_path": crt,
            "leaf_line": _fact_line(leaf),
            "key_line": "server.key + ca.key in %s (private, chmod 0600)" % cert_dir,
        })
        _log("tls %s: leaf[%s] ca[%s] files in %s%s" % (
            mode, CERT_FACTS["leaf_line"], CERT_FACTS["ca_line"], cert_dir,
            ((" - " + extra_note) if extra_note else "")))

    leaf_reason = None
    if os.path.isfile(crt) and os.path.isfile(key):
        leaf_reason = validate_existing_tls(crt)
        if leaf_reason is None and not _pair_matches(crt, key):
            leaf_reason = "the private key does not match the certificate"
        if leaf_reason is None:
            _record("reused")
            _log("reusing the existing loopback certificate - still valid and"
                 " trusted, key matches; nothing was generated or re-imported")
            return (crt, key)
        warn("the existing loopback certificate is not usable (%s)" % leaf_reason)
        note("check result for %s: %s" % (crt, leaf_reason))
    else:
        _log("no loopback certificate on disk yet (%s)" % cert_dir)

    # leaf-only problem: reissue the leaf from the SAME already-trusted CA -
    # never re-trust to fix a chain problem (no keychain prompt this way)
    if _ca_healthy(ca_crt) and os.path.isfile(ca_key):
        days = (_cert_facts(ca_crt).get("days") or "?")
        _log("the CA is still valid + trusted - reissuing only the leaf from it"
             " (CA %d days left; leaf will be re-signed, no re-trust needed)" % days)
        if DRY_RUN:
            note("[dry-run] would reissue the loopback leaf from the existing CA")
            return None
        ext = os.path.join(cert_dir, "localhost.ext")
        with open(ext, "w") as fh:
            fh.write("subjectAltName=DNS:localhost,IP:127.0.0.1,IP:0:0:0:0:0:0:0:1\n"
                     "keyUsage = digitalSignature, keyEncipherment\n"
                     "extendedKeyUsage = serverAuth\n")
        cnf = os.path.join(cert_dir, "ca.cnf")
        db = os.path.join(cert_dir, "index.txt")
        serial = os.path.join(cert_dir, "serial")
        open(db, "a").close()
        if not os.path.exists(serial):
            open(serial, "w").write("1000")
        with open(cnf, "w") as fh:
            fh.write("""[ ca ]
default_ca = CA_default
[ CA_default ]
database = %s
serial = %s
new_certs_dir = %s
certificate = %s
private_key = %s
default_md = sha256
default_days = 3650
policy = policy_any
copy_extensions = copy
[ policy_any ]
commonName = supplied
[ req ]
distinguished_name = req_dn
[ req_dn ]
""" % (db, serial, cert_dir, ca_crt, ca_key))
        rc, o = run_quiet(["openssl", "req", "-newkey", "rsa:2048", "-nodes", "-keyout", key,
                           "-out", os.path.join(cert_dir, "server.csr"),
                           "-subj", "/CN=localhost", "-addext",
                           "subjectAltName=DNS:localhost,IP:127.0.0.1"])
        if rc == 0:
            rc, o = _ca_sign(cert_dir, os.path.join(cert_dir, "server.csr"), crt)
        else:
            o = "leaf key/req generation failed"
        if rc != 0:
            warn("leaf reissue failed - falling back to a full regeneration (%s)"
                 % (o or "openssl error")[:160])
            _log("leaf reissue failure detail: %s" % (o or ""))
        else:
            try:
                os.remove(os.path.join(cert_dir, "server.csr"))
                os.remove(ext)
            except OSError:
                pass
            _protect_secret(key)
            _record("reissued from the existing CA (no new CA, no re-trust)")
            if validate_existing_tls(crt) is None:
                return (crt, key)
            warn("the reissued leaf did not validate - regenerating the full pair")

    if DRY_RUN:
        note("[dry-run] would generate a new loopback CA + localhost certificate"
             " under %s and import the CA into the login keychain" % cert_dir)
        return None
    ext = os.path.join(cert_dir, "localhost.ext")
    with open(ext, "w") as fh:
        fh.write("subjectAltName=DNS:localhost,IP:127.0.0.1,IP:0:0:0:0:0:0:0:1\n"
                 "keyUsage = digitalSignature, keyEncipherment\n"
                 "extendedKeyUsage = serverAuth\n")
    cnf = os.path.join(cert_dir, "ca.cnf")
    db = os.path.join(cert_dir, "index.txt")
    serial = os.path.join(cert_dir, "serial")
    open(db, "a").close()
    if not os.path.exists(serial):
        open(serial, "w").write("1000")
    with open(cnf, "w") as fh:
        fh.write("""[ ca ]
default_ca = CA_default
[ CA_default ]
database = %s
serial = %s
new_certs_dir = %s
certificate = %s
private_key = %s
default_md = sha256
default_days = 3650
policy = policy_any
copy_extensions = copy
[ policy_any ]
commonName = supplied
[ req ]
distinguished_name = req_dn
[ req_dn ]
""" % (db, serial, cert_dir, ca_crt, ca_key))
    cmds = [
        ["openssl", "req", "-x509", "-newkey", "rsa:2048", "-sha256", "-days", "3650",
         "-nodes", "-keyout", ca_key, "-out", ca_crt,
         "-subj", "/CN=OfficeKick Loopback CA"],
        ["openssl", "req", "-newkey", "rsa:2048", "-nodes", "-keyout", key,
         "-out", os.path.join(cert_dir, "server.csr"),
         "-subj", "/CN=localhost", "-addext",
         "subjectAltName=DNS:localhost,IP:127.0.0.1"],
    ]
    for cmd in cmds:
        rc, o = run_quiet(cmd)
        if rc != 0:
            warn("TLS cert generation failed (%s)" % (o or cmd[0]))
            _log("generation failure detail: %s" % (o or ""))
            return None
    rc, o = _ca_sign(cert_dir, os.path.join(cert_dir, "server.csr"), crt)
    if rc != 0:
        warn("TLS cert signing failed (%s)" % (o or "openssl error")[:160])
        _log("signing failure detail: %s" % (o or ""))
        CERT_FACTS = {}
        return None
    try:
        os.remove(os.path.join(cert_dir, "server.csr"))
        os.remove(ext)
    except OSError:
        pass
    _protect_secret(key)
    keychain = os.path.expanduser("~/Library/Keychains/login.keychain-db")
    step("macOS opens its own dialog now - ACCEPT the loopback CA"
         " (click Trust) to make https work")
    rc, o = run_quiet(["security", "add-trusted-cert", "-r", "trustRoot",
                       "-k", keychain, ca_crt])
    if rc == 0:
        ok("loopback CA trusted in your login keychain (imported: %s)" % ca_crt)
    else:
        warn("could not auto-trust the loopback CA - run once yourself, then restart the server:")
        print("       security add-trusted-cert -r trustRoot -k $HOME/Library/Keychains/login.keychain-db " + ca_crt)
    _record("generated (fresh CA + leaf)")
    return (crt, key)



# ----------------------------------------------- self-signed leaf mode ---- #
CERT_META_NAME = "cert_meta.json"


def _leaf_paths():
    cdir = tls_cert_dir()
    return (cdir,
            os.path.join(cdir, "server.crt"),
            os.path.join(cdir, "server.key"),
            os.path.join(cdir, CERT_META_NAME))


def _leaf_meta_path():
    """The certificate metadata (fingerprints, mode) lives in the CONFIG dir -
    it must survive the certificates folder so an uninstall can still find
    the sha1 of a previously trusted leaf even after its files are gone."""
    return os.path.join(config_dir(), CERT_META_NAME)


def _leaf_meta_load():
    p = _leaf_meta_path()
    try:
        with open(p, "r", encoding="utf-8") as fh:
            return json.load(fh) or {}
    except Exception:
        pass
    try:   # pre-0.9.247 location (next to the certificates)
        with open(_leaf_paths()[3], "r", encoding="utf-8") as fh:
            meta = json.load(fh) or {}
        if meta:
            try:
                with open(p, "w", encoding="utf-8") as fh:
                    json.dump(meta, fh, indent=2)
                _protect_secret(p)
            except OSError:
                pass
        return meta
    except Exception:
        return {}


def _leaf_meta_save(meta):
    try:
        with open(_leaf_meta_path(), "w", encoding="utf-8") as fh:
            json.dump(meta, fh, indent=2)
        _protect_secret(_leaf_meta_path())
    except OSError as exc:
        warn("could not write the certificate metadata: %s" % exc)


def _leaf_fingerprints(crt):
    """(sha256, sha1) of the certificate, used for trust-store identification.
    Uses openssl's own fingerprint output - the previous DER-bytes route broke
    on every binary/encoding roundtrip and produced the sha1 of an empty
    string, which made trust removal delete NOTHING (found by a real
    install -> uninstall audit)."""
    import re as _re
    if crt is None or not (crt and os.path.isfile(crt)):
        return None, None
    sha256 = sha1 = None
    try:
        rc, o = run_quiet(["openssl", "x509", "-in", crt, "-noout",
                           "-fingerprint", "-sha256"])
        if rc == 0:
            m = _re.search(r"=([0-9A-Fa-f:]+)", o or "")
            if m:
                sha256 = m.group(1).replace(":", "").upper()
        rc, o = run_quiet(["openssl", "x509", "-in", crt, "-noout",
                           "-fingerprint", "-sha1"])
        if rc == 0:
            m = _re.search(r"=([0-9A-Fa-f:]+)", o or "")
            if m:
                sha1 = m.group(1).replace(":", "").upper()
    except Exception:
        pass
    if sha1 is None or sha256 is None:
        # WINDOWS: openssl is absent - the leaf generator wrote this file
        try:
            from cryptography.hazmat.primitives import hashes
            cert = _load_cert_from_pem(crt)
            if cert is not None:
                if sha256 is None:
                    sha256 = cert.fingerprint(hashes.SHA256()).hex().upper()
                if sha1 is None:
                    sha1 = cert.fingerprint(hashes.SHA1()).hex().upper()
        except Exception:
            pass
    return sha256, sha1


def _leaf_install_id():
    meta = _leaf_meta_load()
    iid = meta.get("install_id")
    if not iid:
        import uuid
        iid = uuid.uuid4().hex[:10]
    return iid


def _generate_leaf_cert(cdir, crt, key, days, install_id):
    """Generate a self-signed non-CA leaf: CN=localhost (+ OU=officekick-<id>
    for identification), SAN DNS:localhost + IP:127.0.0.1 + IP:::1, critical
    basicConstraints CA:FALSE, digitalSignature key usage, serverAuth EKU,
    SHA-256, RSA 3072. Returns True on success.
    macOS: the openssl binary (as for the CA mode). Windows: the cryptography
    package (user-installable; no keychain/store tricks needed to write PEM)."""
    if is_windows():
        try:
            import ipaddress
            import traceback
            import cryptography
        except ImportError:
            # install the dependency OURSELVES (user-scope pip, no admin) -
            # standing behavior from the mac waves: the installer never sends
            # the user away for a dependency it can fetch
            warn("the 'cryptography' package is missing on Windows - installing"
                 " it now (pip --user, no admin needed)")
            import importlib
            installed = False
            for pip_args in (["--user"], []):
                rc, o = run_quiet([sys.executable, "-m", "pip", "install"] +
                                  pip_args + ["cryptography"], timeout=600)
                _log("pip install %scryptography: rc %d out=%s"
                     % ((" ".join(pip_args) + " ") if pip_args else "", rc,
                        (o or "")[-220:]), ts=True)
                if rc == 0:
                    importlib.invalidate_caches()
                    try:
                        import cryptography  # noqa
                        installed = True
                        break
                    except ImportError:
                        continue
            if installed:
                ok("installed the 'cryptography' package")
            else:
                warn("could not auto-install 'cryptography' - check pip/network, then re-run:")
                print("        python -m pip install --user cryptography")
                _log("leaf generation unavailable: 'cryptography' install failed (Windows)")
                return False
        try:
            from cryptography import x509
            from cryptography.hazmat.primitives import hashes, serialization
            from cryptography.hazmat.primitives.asymmetric import rsa
            from cryptography.x509.oid import NameOID, ExtendedKeyUsageOID
            import datetime
            key_obj = rsa.generate_private_key(public_exponent=65537, key_size=3072)
            name = x509.Name([
                x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
                x509.NameAttribute(NameOID.ORGANIZATION_NAME, "OfficeKick"),
                x509.NameAttribute(NameOID.ORGANIZATIONAL_UNIT_NAME,
                                   "officekick-%s" % install_id),
            ])
            now = datetime.datetime.now(datetime.timezone.utc)
            builder = (x509.CertificateBuilder()
                       .subject_name(name).issuer_name(name)
                       .public_key(key_obj.public_key())
                       .serial_number(x509.random_serial_number())
                       .not_valid_before(now - datetime.timedelta(days=1))
                       .not_valid_after(now + datetime.timedelta(days=int(days)))
                       .add_extension(x509.BasicConstraints(ca=False, path_length=None),
                                      critical=True)
                       .add_extension(x509.KeyUsage(digital_signature=True,
                                                    content_commitment=False,
                                                    key_encipherment=False,
                                                    data_encipherment=False,
                                                    key_agreement=False,
                                                    key_cert_sign=False,
                                                    crl_sign=False,
                                                    encipher_only=False,
                                                    decipher_only=False),
                                      critical=False)
                       .add_extension(x509.ExtendedKeyUsage(
                           [ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
                       .add_extension(x509.SubjectAlternativeName([
                           x509.DNSName("localhost"),
                           x509.IPAddress(ipaddress.ip_address("127.0.0.1")),
                           x509.IPAddress(ipaddress.ip_address("::1")),
                       ]), critical=False)
                       .add_extension(x509.SubjectKeyIdentifier.from_public_key(
                           key_obj.public_key()), critical=False)
                       .sign(key_obj, hashes.SHA256()))
            with open(key, "wb") as fh:
                fh.write(key_obj.private_bytes(serialization.Encoding.PEM,
                                               serialization.PrivateFormat.PKCS8,
                                               serialization.NoEncryption()))
            with open(crt, "wb") as fh:
                fh.write(builder.public_bytes(serialization.Encoding.PEM))
            return True
        except Exception as exc:
            warn("leaf certificate generation failed: %s" % exc)
            _log("leaf generation exception: %s" % traceback.format_exc())
            return False
    cmd = ["openssl", "req", "-x509", "-newkey", "rsa:3072", "-sha256",
           "-nodes", "-keyout", key, "-out", crt,
           "-days", str(int(days)),
           "-subj", "/CN=localhost/O=OfficeKick/OU=officekick-%s" % install_id,
           "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1,IP:::1",
           "-addext", "basicConstraints=critical,CA:FALSE",
           "-addext", "keyUsage=digitalSignature",
           "-addext", "extendedKeyUsage=serverAuth"]
    rc, o = run_quiet(cmd)
    if rc != 0:
        warn("leaf certificate generation failed (%s)" % (o or "openssl error")[:160])
        _log("leaf generation failure detail: %s" % (o or ""))
        return False
    return True


def _win_import_trusted_people(crt):
    """Import + VERIFY a certificate in the CURRENT USER Trusted People
    store. PowerShell first (Import-Certificate - reliable and it RETURNS the
    thumbprint), certutil as the fallback. Returns (ok, detail)."""
    _sha256, sha1 = _leaf_fingerprints(crt)
    ps = ("$p='%s'; $c = Import-Certificate -FilePath $p "
          "-CertStoreLocation Cert:\\CurrentUser\\TrustedPeople; "
          "if ($null -eq $c) { exit 3 }; Write-Output $c.Thumbprint" % crt)
    rc, o = run_quiet(["powershell", "-NoProfile", "-ExecutionPolicy",
                       "Bypass", "-Command", ps], timeout=60)
    if rc == 0 and (o or "").strip().upper() == (sha1 or "").upper():
        return True, "imported by PowerShell (thumbprint verified)"
    warn("PowerShell Import-Certificate did not confirm (rc %d, %s) - trying certutil"
         % (rc, (o or "")[:60]))
    rc, o = run_quiet(["certutil", "-user", "-addstore", "TrustedPeople", crt])
    rc2, o2 = run_quiet(["certutil", "-user", "-store", "TrustedPeople"], timeout=60)
    listing = (o2 or "").replace(" ", "").upper()
    if rc2 == 0 and sha1 and sha1.upper() in listing:
        return True, "imported by certutil (thumbprint verified in the store)"
    return False, "PowerShell rc %d + certutil rc %d - the leaf is NOT listed" % (rc, rc2)


def _win_schannel_ok(port):
    """The decisive Windows check: a REAL https request through the WINDOWS
    TLS stack (SChannel). 200 = the OS accepts the leaf. python's ssl probe
    uses OpenSSL with its own trust store, which says nothing about what
    Office's WebView2 will do."""
    ps = ("try { $r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 6 "
          "-Uri 'https://localhost:%d/api/health'; exit [int]$r.StatusCode } "
          "catch { exit -1 }" % port)
    rc, o = run_quiet(["powershell", "-NoProfile", "-ExecutionPolicy",
                       "Bypass", "-Command", ps], timeout=30)
    if rc == 200:
        return True, "the Windows TLS handler accepts the leaf (test request: 200)"
    return False, "the Windows TLS handler REJECTS the leaf (test request rc %s)" % rc


def _install_leaf_trust(crt):
    """Platform trust for ONE self-signed leaf: macOS = user keychain
    per-certificate SSL trust (no CA involved); Windows = the CURRENT USER
    Trusted People store (a WebView2-verified import, with an EXPLICITLY
    consented Root store fallback; no admin needed)."""
    if is_mac():
        keychain = os.path.expanduser("~/Library/Keychains/login.keychain-db")
        step("macOS opens its own dialog now - ACCEPT the leaf certificate"
             " (click Trust) to make https work")
        rc, o = run_quiet(["security", "add-trusted-cert", "-r", "trustRoot",
                           "-p", "ssl", "-k", keychain, crt])
        if rc == 0:
            ok("self-signed leaf trusted for TLS in your login keychain (user scope)")
            return True
        warn("could not trust the leaf certificate (%s)" % (o or "security error")[:160])
        _log("leaf trust failure detail: %s" % (o or ""))
        return False
    got, detail = _win_import_trusted_people(crt)
    if got:
        ok("self-signed leaf added to the CURRENT USER Trusted People store (%s)" % detail)
        tls_ok, tls_detail = _win_schannel_ok(HTTPS_PORT if HTTPS_PORT else PORT + 1)
        if tls_ok:
            ok(tls_detail)
        else:
            warn("%s - try the Root-store option below" % tls_detail)
        note("Windows does not prompt for this import - it is silent by design.")
        print("")
        print("  " + _paint("Where to see it:", "bold"))
        print("    Win+R -> certmgr.msc > Trusted People > Certificates")
        print("    (CN=localhost). The simplified control-panel applet can")
        print("    hide this store; certmgr.msc always shows it. Or check:")
        print("        certutil -user -store TrustedPeople")
        # WebView2 (Office's webview) does not always honour a leaf-only
        # trust: offer the contained workaround explicitly (the leaf is
        # CA:FALSE - a Root entry trusts exactly this one certificate;
        # Windows shows its own confirmation dialog for Root adds).
        go = ask_yes("Also trust the leaf in the CurrentUser ROOT store?"
                     " (recommended if the add-in reports a certificate error;"
                     " Windows will show its own dialog)", False)
        if go:
            step("Windows opens its own certificate dialog now - ACCEPT it"
                 " (click Yes) to add the leaf certificate to the Root store")
            rcr, orr = run_quiet(["certutil", "-user", "-addstore", "Root", crt])
            if rcr == 0:
                ok("leaf also registered in the CurrentUser ROOT store")
                meta = _leaf_meta_load()
                meta["rootThumbprint"] = (sha1 if 'sha1' in dir() else
                                          _leaf_fingerprints(crt)[1])
                _leaf_meta_save(meta)
            else:
                warn("Root-store import failed (%s)" % (orr or "certutil error")[:160])
        else:
            note("if the pane still reports a certificate error, re-run install.py"
                 " and answer YES to the Root-store option, or switch to the CA mode"
                 " (Transport options > certificate mode).")
        return True
    warn("could not import + verify the leaf in Trusted People (%s)"
         % detail[:160])
    _log("leaf trust failure detail: %s" % detail)
    return False


def _trust_settings_forget_sha1(sha1):
    """Remove a USER-domain trust-settings record identified by its sha1 hash
    (the cert FILE may no longer exist): export the plist, drop the hash key,
    re-import. Used for our own entries only - hashes we stored ourselves."""
    import plistlib
    try:
        p = os.path.join(config_dir(), "trust_export_user.plist")
        rc, _o = run_quiet(["security", "trust-settings-export", p])
        if rc != 0 or not os.path.isfile(p):
            return False
        pl = plistlib.load(open(p, "rb"))
        tl = pl.get("trustList", {})
        key40 = [k for k in tl if k.upper() == str(sha1).upper()]
        if not key40:
            return True    # nothing to forget
        for k in key40:
            tl.pop(k)
        plistlib.dump(pl, open(p, "wb"), fmt=plistlib.FMT_XML)
        rc, _o = run_quiet(["security", "trust-settings-import", p])
        try:
            os.remove(p)
        except OSError:
            pass
        _log("trust-settings hash entry %s removed (re-import)" % key40[0], ts=True)
        return rc == 0
    except Exception as exc:
        warn("could not rebuild the user trust settings: %s" % exc)
        return False


def _remove_leaf_trust(crt):
    """Remove OUR leaf trust (and only ours - identified by fingerprint hash),
    during a replacement or uninstallation. Never touches other entries."""
    try:
        sha256, sha1 = _leaf_fingerprints(crt)
    except Exception:
        sha256, sha1 = None, None
    meta = _leaf_meta_load()
    if sha1 is None:
        sha1 = meta.get("sha1")
    if is_mac():
        # remove-trusted-cert takes the cert FILE when present (NO -k flag:
        # that flag does not exist - the old call failed every time and left
        # the trust entry in place); the keychain item goes by sha1.
        kc = "login.keychain"
        if crt and os.path.isfile(crt):
            rc, _o = run_quiet(["security", "remove-trusted-cert", crt])
            _log("trust settings removed via the cert file (rc %d)" % rc, ts=True)
        elif sha1:
            _trust_settings_forget_sha1(sha1)
        oki = False
        if sha1 and str(sha1) not in ("",) and sha1.upper() != "DA39A3EE5E6B4B0D3255BFEF95601890AFD80709":
            rc, o = run_quiet(["security", "delete-certificate", "-Z", sha1, kc])
            oki = (rc == 0)
            _log("keychain item delete (sha1=%s): rc %d %s" % (sha1, rc, o or ""),
                 ts=True)
        else:
            warn("no usable fingerprint for the leaf trust entry - check the"
                 " keychain manually (the stored metadata was empty)")
        _log("leaf trust/certificate removal attempted (sha1=%s)" % sha1, ts=True)
        return True
    if sha1:
        rc, o = run_quiet(["certutil", "-user", "-delstore", "TrustedPeople", sha1])
        _log("leaf removal from TrustedPeople (sha1=%s): rc %d %s" % (sha1, rc, o or ""))
        meta = _leaf_meta_load()
        if meta.get("rootThumbprint"):
            rcr, orr = run_quiet(["certutil", "-user", "-delstore", "Root",
                                  meta["rootThumbprint"]])
            _log("leaf removal from Root (thumbprint=%s): rc %d %s"
                 % (meta["rootThumbprint"], rcr, orr or ""))
        return rc == 0
    return False


def _leaf_consent(facts, crt):
    """The exact consent block requested for leaf mode. Returns True only
    after an explicit yes; declining never touches the trust store."""
    sha256, _sha1 = _leaf_fingerprints(crt)
    print("")
    print("  " + _paint("Trust Local HTTPS Certificate", "bold"))
    print("")
    print("  This application uses a local HTTPS server to provide its Microsoft Office Add-in.")
    print("  To establish a secure connection, the application has generated a unique,")
    print("  self-signed server certificate for this computer.")
    print("  This certificate is not a Certificate Authority and cannot be used to issue")
    print("  trusted certificates for other websites.")
    print("  To allow Microsoft Office to load the Add-in, you must explicitly trust this")
    print("  certificate for local HTTPS connections.")
    print("  The certificate and private key are generated and stored locally.")
    print("  The private key will not be transmitted or shared.")
    print("  After you answer yes below, YOUR SYSTEM opens one more dialog asking")
    print("  you to ACCEPT adding the leaf certificate - accept it. macOS shows a")
    print("  keychain dialog (click Trust); Windows shows its certificate dialog")
    print("  (click Yes) when the Root-store option is used.")
    print("")
    print("  " + _paint("certificate : %s" % (CERT_FACTS.get("leaf_path", crt)), "bold"))
    print("  " + _paint("SHA-256 fingerprint: %s" % _colons(sha256 or "?"), "cyan"))
    print("  " + _paint("expires: %s" % _fact_line(facts), "cyan"))
    print("")
    # https transport NEEDS this trust - the default follows the need (still
    # an explicit, cancellable screen; declining leaves the store untouched
    # and https stays off until the leaf is trusted)
    need = _current_transport() == "https"
    return ask_yes("Do you want to trust this certificate?"
                   + (" (Enter = trust; https needs it)" if need else ""), need)


def _colons(s):
    if not s or len(s) < 5:
        return s
    return ":".join(s[i:i + 2] for i in range(0, len(s), 2))


def ensure_leaf_certs(days=825, consent_ask=True):
    """Leaf certificate mode: detect existing generated leaf, reuse a valid
    one (no regeneration on every launch), regenerate when expired/invalid -
    with RENEWED CONSENT before re-trusting a replacement. Returns
    (crt, key) or None; CERT_FACTS carries the details."""
    global CERT_FACTS
    CERT_FACTS = {}
    cdir, crt, key, meta_path = _leaf_paths()
    os.makedirs(cdir, exist_ok=True)
    install_id = _leaf_install_id()

    def _record(mode, consent_note=""):
        f = _cert_facts(crt)
        sha256, sha1 = _leaf_fingerprints(crt)
        CERT_FACTS.update({"mode": mode, "leaf_path": crt,
                           "leaf_line": _fact_line(f),
                           "ca_path": cdir,
                           "ca_line": "no CA in this mode - the leaf itself is the trusted end point",
                           "key_line": "server.key + cert_meta.json in %s (private, protected)" % cdir})
        meta = _leaf_meta_load()
        meta.update({"app": "OfficeKick", "mode": "self_signed_leaf",
                     "install_id": install_id, "cn": "localhost",
                     "sha256": sha256, "sha1": sha1,
                     "not_after": f.get("expires"), "days": f.get("days"),
                     "created": meta.get("created") or time.strftime("%Y-%m-%d"),
                     "platform": "windows" if is_windows() else "mac"})
        if consent_note:
            meta["note"] = consent_note
        _leaf_meta_save(meta)

    reason = None
    if os.path.isfile(crt) and os.path.isfile(key):
        reason = validate_existing_tls(crt)
        if reason is None and not _pair_matches(crt, key):
            reason = "the private key does not match the certificate"
    meta = _leaf_meta_load()
    if reason is None and isinstance(meta.get("days"), int) and meta["days"] < 30:
        reason = "expires in %d days" % meta["days"]
    if reason is None and os.path.isfile(crt) and os.path.isfile(key) \
            and meta.get("mode") == "self_signed_leaf":
        # REUSE REQUIRES THE FILES: trusting the metadata alone produced a
        # "reused" pair of nonexistent files and pushed the install into the
        # CA-mode fallback (found in a live audit)
        _record("reused" + (" (%s)" % ("trusted" if meta.get("trusted") else "NOT trusted yet")))
        _log("leaf mode: reusing the existing certificate (still valid; not"
             " regenerated on every launch)")
        return (crt, key)
    if reason:
        warn("the existing leaf certificate must be replaced: %s" % reason)

    if DRY_RUN:
        note("[dry-run] would generate a self-signed leaf (validity %d days)" % days)
        return None
    if not _generate_leaf_cert(cdir, crt, key, days, install_id):
        return None
    _protect_secret(key)
    _record("generated (self-signed leaf, no CA)")

    if not consent_ask:
        note("leaf generated; trust consent will be asked before it is used")
        return (crt, key)
    f = _cert_facts(crt)
    if _leaf_consent(f, crt):
        if _install_leaf_trust(crt):
            old = meta.get("sha1") if meta else None
            m2 = _leaf_meta_load()
            m2["trusted"] = True
            if m2.get("note"):
                m2.pop("note")
            _leaf_meta_save(m2)
            CERT_FACTS["mode"] = "trusted self-signed leaf (renewed consent given)" if old else                                  "trusted self-signed leaf"
            return (crt, key)
        warn("the certificate was generated but NOT trusted - https will not pass"
             " validation until it is trusted (re-run install.py)")
        return None
    warn("you declined the trust - the certificate store was NOT touched and")
    warn("https WILL NOT WORK with this leaf. Re-run: python install.py --install")
    warn("(then answer y at the trust question; the uninstall removes it again)")
    _leaf_meta_load() and _leaf_meta_save(dict(_leaf_meta_load(),
                                               trusted=False))
    CERT_FACTS["mode"] = "generated but NOT trusted (declined)"
    return None


def _current_cert_host():
    """The hostname the https manifests use (the certHost setting)."""
    d = _config_settings()
    h = str(d.get("certHost") or "localhost")
    import re as _re
    return h.lower() if _re.fullmatch(r"(?i)[a-z0-9._-]{1,253}", h) else "localhost"


def _set_cert_host(host):
    if not re.fullmatch(r"(?i)[a-z0-9._-]{1,253}", host or ""):
        return False
    _persist_file({"certHost": host})
    _api_set({"certHost": host})    # mirror to a running server
    return True
def _validate_byo(crt, key, host):
    """(ok, [problems], facts) for a bring-your-own pair + hostname."""
    problems = []
    if not (os.path.isfile(crt) and os.path.isfile(key)):
        return False, ["certificate or key file not found"], {}
    f = _cert_facts(crt)
    if not f:
        problems.append("the certificate could not be parsed (openssl)")
    rc, o = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-text"])
    if rc == 0 and ("DNS:%s" % host) not in o and ("IP Address:%s" % host) not in o             and host != "localhost":
        problems.append("the certificate SAN does not cover '%s'" % host)
    if rc == 0 and "CA:TRUE" in o:
        problems.append("this is a CA certificate, not a server certificate")
    if not _pair_matches(crt, key):
        problems.append("the private key does not match the certificate")
    reason = validate_existing_tls(crt, host)
    if reason:
        if "not trusted" in reason or "chain is not trusted" in reason:
            # BYO: the user attests the certificate is already trusted; the
            # local verify result is reported, not a hard rejection
            note("verify-cert could not confirm trust locally (%s) - in BYO"
                 " mode you guarantee the trust; if Office refuses the pane,"
                 " the certificate is not trusted after all (leaf mode is the"
                 " alternative)" % reason)
        else:
            problems.append("not usable: %s" % reason)
    return (not problems), problems, f


def _ask_cert_setup():
    """THE certificate flow - one deterministic path per install run:
    1. existing usable pair -> keep it (updates never re-ask);
    2. a previously stored certMode -> run that mode directly;
    3. otherwise -> the three-way question: self-signed leaf / local CA +
       CA-signed certificate (mac) / bring-your-own (with an optional
       hostname for the manifests)."""
    host = _current_cert_host()
    tls_dir = tls_cert_dir()
    crt = os.path.join(tls_dir, "server.crt")
    key = os.path.join(tls_dir, "server.key")
    usable = os.path.isfile(crt) and validate_existing_tls(crt, host) is None \
        and _pair_matches(crt, key)
    chosen = _config_settings().get("certMode") \
        if _config_settings().get("certMode") in ("ca", "self_signed_leaf", "byo") else None
    if not usable and chosen == "byo" and os.path.isfile(crt) and os.path.isfile(key):
        # BYO re-runs keep the certificate even when the local verify-trust
        # check is inconclusive (the user attests trust)
        usable, _pr, _f = _validate_byo(crt, key, host)
    if usable:
        note("certificate setup skipped - the existing %s pair validates"
             " (mode %s stays)" % (host, chosen or "ca"))
        return None
    win = is_windows()
    byo_no = 3 if not win else 2
    ca_no = 2 if not win else None
    default_no = 1
    if chosen == "ca" and ca_no:
        default_no = ca_no
    elif chosen == "byo":
        default_no = byo_no
    if chosen:
        # a stored mode only auto-runs while a usable pair exists (updates);
        # with the certificates gone (e.g. the uninstall removed them) the
        # question comes back, PRE-SELECTING the stored mode
        print("")
        note("no usable certificate pair on disk although mode '%s' is stored -"
             " it stays PRE-SELECTED below (Enter re-runs it)" % chosen)
    print("")
    print("  " + _paint("Certificate setup for the local https server", "bold")
          + _paint("  (before step 3 of 6 - certificates)", "dim"))
    print("    1) self-signed leaf certificate, DIRECTLY trusted (no CA is created;")
    print("       the LEAF itself is imported into your keychain; your consent is")
    print("       asked again whenever it must be re-trusted)")
    if not win:
        print("    2) local CA + CA-signed certificate (the CA is imported + trusted")
        print("       in your keychain; renewals are silent after that)")
    print("    %d) bring your OWN already-trusted certificate + key (PEM files)" % byo_no)
    while True:
        ans = input("  %s" % _paint("? choice [default %s]: " % default_no, "cyan")).strip() \
            or str(default_no)
        _log("certificate setup choice: %r" % ans, ts=True)
        if ans == "1":
            _set_cert_mode("self_signed_leaf")
            ok("certificate mode: self-signed leaf (no CA)")
            return ensure_leaf_certs()
        if not win and ans == "2":
            _set_cert_mode("ca")
            ok("certificate mode: local CA + CA-signed certificate")
            return ensure_tls_certs()
        if ans == str(byo_no):
            _set_cert_mode("byo")
            return _byo_setup()
        print("  please answer 1, %s" % ((", 2 or 3" if not win else "or 2")))


def _byo_setup():
    """Bring-your-own: hostname prompt (the manifests + validation use it) +
    the PEM pair; the certificate must already be trusted by the user (we do
    not touch any trust store in this mode)."""
    default_host = _current_cert_host()
    print("")
    print("  " + _paint("Bring your own certificate", "bold"))
    host = input("  " + _paint("? hostname to use in the manifests"
                               " (Enter = %s): " % default_host, "cyan")).strip().lower()
    if not host:
        host = default_host
    if not re.fullmatch(r"(?i)[a-z0-9._-]{1,253}", host):
        warn("'%s' is not a usable DNS name - staying on %s" % (host, default_host))
        host = default_host
    _set_cert_host(host)
    ok("manifest hostname: %s%s" % (host,
       "  (make sure it resolves to 127.0.0.1, e.g. an /etc/hosts entry -"
       " the server stays loopback-only)" if host != "localhost" else ""))
    for attempt in range(3):
        try:
            crt_path = input("  " + _paint("? path to your TLS certificate (PEM): ",
                                           "cyan")).strip()
            key_path = input("  " + _paint("? path to the matching private key (PEM, unencrypted): ",
                                           "cyan")).strip()
        except EOFError:
            return None
        good, problems, facts = _validate_byo(crt_path, key_path, host)
        if good:
            cdir, crt, key, _m = _leaf_paths()
            # copy unless the source IS the target (a missing target copies);
            # samefile needs both files present
            if not (os.path.exists(crt) and os.path.samefile(crt_path, crt)):
                shutil.copyfile(crt_path, crt)
            if not (os.path.exists(key) and os.path.samefile(key_path, key)):
                shutil.copyfile(key_path, key)
            _protect_secret(key)
            global CERT_FACTS
            CERT_FACTS.update({"mode": "your own certificate (BYO, host %s)" % host,
                               "leaf_path": crt, "leaf_line": _fact_line(facts),
                               "ca_path": cdir, "ca_line": "n/a in BYO mode",
                               "key_line": "server.key in %s (private, protected)" % cdir})
            meta = _leaf_meta_load()
            sha256, sha1 = _leaf_fingerprints(crt)
            meta.update({"app": "OfficeKick", "mode": "byo", "certHost": host,
                         "sha256": sha256, "sha1": sha1,
                         "not_after": facts.get("expires")})
            _leaf_meta_save(meta)
            ok("BYO certificate installed for %s (we do not touch the trust"
               " store - the certificate is already trusted by you)" % host)
            return (crt, key)
        for pr in problems:
            warn(pr)
        print("       (attempt %d/3)" % (attempt + 1))
    warn("giving up - keep the current setup or re-run install.py")
    return None


def cert_diagnostics():
    """Run the local certificate checks and print an explicit verdict -
    including the line that store trust does not equal Office compatibility."""
    cdir, crt, key, _meta = _leaf_paths()
    print("")
    print("  ----- certificate diagnostics -----")
    checks = []
    checks.append(("files exist", os.path.isfile(crt) and os.path.isfile(key),
                   "%s + %s" % (crt, key)))
    rc, o = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-text"])
    host = _current_cert_host()
    if host != "localhost":
        san_ok = ("DNS:%s" % host) in (o or "")
        checks.append(("SAN covers '%s'" % host, rc == 0 and san_ok, ""))
    else:
        san_ok = ("localhost" in (o or "") and "127.0.0.1" in (o or ""))
        checks.append(("SAN localhost + 127.0.0.1", rc == 0 and san_ok, ""))
    ca_signed = rc == 0 and "Issuer:" in (o or "") and \
        ("Issuer: CN=OfficeKick" in (o or "") or "issuer" in (o or "").lower())
    bc_bad = rc == 0 and "CA:TRUE" in (o or "")
    bc_ok = not bc_bad
    checks.append(("basicConstraints (leaf, not a CA)", rc == 0 and bc_ok,
                   "" if not (rc == 0 and "basic constraints" not in (o or "").lower()) else
                   "no explicit basicConstraints (older CA-mode leaf, still valid)"))
    eku_ok = "TLS Web Server Authentication" in (o or "")
    if not eku_ok and ca_signed:
        checks.append(("EKU serverAuth", True,
                       "informational: created by an older CA-mode leaf (0.9.239);"
                       " delete the cert pair to regenerate with full extensions"))
    else:
        checks.append(("EKU serverAuth", eku_ok, ""))
    pair_ok = _pair_matches(crt, key) if os.path.isfile(crt) else False
    checks.append(("private key matches certificate", bool(pair_ok), ""))
    f = _cert_facts(crt)
    v_ok = isinstance(f.get("days"), int) and f["days"] >= 0
    checks.append(("validity window", v_ok, _fact_line(f)))
    tls_on, tls_desc = _https_state()
    checks.append(("localhost TLS handshake", tls_on, tls_desc))
    trust_ok = False
    if is_mac() and os.path.isfile(crt):
        trust_ok = run_quiet(["security", "verify-cert", "-c", crt, "-p", "ssl"])[0] == 0
    elif is_windows():
        trust_ok = run_quiet(["certutil", "-user", "-verify", "-store", "TrustedPeople", crt])[0] == 0
    checks.append(("platform trust accepts the certificate", bool(trust_ok),
                   "store trust is NOT the same as Office compatibility"))
    for (label, okk, extra) in checks:
        print("    %s %-38s %s" % (_paint("v" if okk else "x", "green" if okk else "red"),
                                   label, extra or ""))
    print("")
    print("  " + _paint("Store trust is NOT Office compatibility:", "bold"))
    print("  to verify in Office: quit and reopen the app, open the task pane")
    print("  (Insert > Add-ins > My Add-ins), confirm the pane loads without a")
    print("  certificate warning, and check its About tab for the pane build.")
    print("  - macOS: WKWebView uses the per-certificate SSL trust in the user")
    print("    keychain (verify-cert v covers the same evaluation).")
    print("  - Windows: WebView2 honors the Current User Trusted People store for")
    print("    explicitly trusted certificates - VERIFY this in Office before")
    print("    relying on it; if it refuses, report the failure and stay on the")
    print("    CA mode instead of weakening the trust model.")
    metaj = _leaf_meta_load()
    if metaj:
        print("  installed-by-OfficeKick identity: OU=officekick-%s | sha256=%s"
              % (metaj.get("install_id"), _colons(metaj.get("sha256") or "?")))
    return all(okk for (_l, okk, _e) in checks)


def _get_cert_mode():
    """The certificate mode; the DEFAULT is the self-signed leaf (no CA)."""
    d = _config_settings()
    return d.get("certMode") if d.get("certMode") in ("ca", "self_signed_leaf", "byo") \
        else "self_signed_leaf"


def _set_cert_mode(mode):
    if mode not in ("ca", "self_signed_leaf", "byo"):
        return False
    _persist_file({"certMode": mode})
    _api_set({"certMode": mode})    # mirror to a running server
    return True
def deploy_files(addins=None):
    dst = install_dir()
    if addins is None:
        addins = ADDINS or set()
    if DRY_RUN:
        n = len([a for a in OURS if a[0] in addins and os.path.isdir(repo_path(a[0]))])
        return "[dry-run] server + %d add-in folder(s) -> %s" % (n, dst)
    shutil.rmtree(dst, ignore_errors=True)
    os.makedirs(manifest_dir(), exist_ok=True)
    os.makedirs(config_dir(), exist_ok=True)
    shutil.copyfile(repo_path("core", "server.py"), server_py())
    ignore = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store")
    deployed = []
    for (folder, name, _mf) in OURS:
        if folder not in addins:
            note("not deploying %s (not selected this run)" % name)
            continue
        if not os.path.isdir(repo_path(folder)):
            warn("%s not found in this checkout - skipped" % name)
            continue
        shutil.copytree(repo_path(folder), os.path.join(dst, folder),
                        ignore=ignore, dirs_exist_ok=True)
        deployed.append(name)
    cfgjson = os.path.join(config_dir(), "config.json")
    if os.path.exists(cfgjson):
        try:
            _protect_secret(cfgjson)
        except OSError:
            pass
    record_repo_info()
    _log("deployed to %s: %s" % (dst, ", ".join(deployed) or "server only"), ts=True)
    return (" + ".join(deployed) if deployed else "server") + " -> " + dst


def record_repo_info():
    """Remember where this checkout lives (git clone or zip copy) in the
    config so the server can offer a real git-pull update check. Falls back
    gracefully: no .git and no git binary = the server uses its GitHub
    raw-VERSION check."""
    import json
    cfgjson = config_file()
    try:
        with open(cfgjson, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        data = {}
    gitdir = os.path.join(REPO_DIR, ".git")
    is_repo = os.path.isdir(gitdir) or os.path.isfile(gitdir)   # worktree file counts
    if is_repo and DRY_RUN:
        step("[dry-run] would record repoPath=%s" % REPO_DIR)
        return
    data["repoPath"] = REPO_DIR if is_repo else None
    url = None
    if is_repo:
        rc, o = run_quiet(["git", "-C", REPO_DIR, "remote", "get-url", "origin"])
        if rc == 0 and o.strip():
            url = o.strip().splitlines()[0]
    data["repoUrl"] = url
    try:
        with open(cfgjson, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=1)
        if os.name != "nt":
            _protect_secret(cfgjson)
        if is_repo:
            ok("repo recorded for update checks: %s%s" % (REPO_DIR, (" (" + url + ")") if url else ""))
        else:
            warn("no git checkout at %s - update checks use the GitHub VERSION file" % REPO_DIR)
    except OSError as e:
        warn("could not record the repo path: %s" % e)


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


def render_manifest(addin_dir, port, https_port=None, host=None):
    """https_port: when a working loopback certificate exists, ALL add-in
    panes are served over https://<host>:<https port> (the shared TLS
    listener) and their manifests point there; None = plain http.
    host = the hostname the https manifests use (the certHost setting; by
    default localhost, or the BYO certificate's own DNS name)."""
    with open(repo_path(addin_dir, "manifest.template.xml"), "r", encoding="utf-8") as fh:
        t = fh.read()
    if RIBBON:
        t = re.sub(r"(?m)^[ \t]*__COMMANDS_(?:BEGIN|END)__[ \t]*\r?\n?", "", t)
    else:
        t = re.sub(r"(?s)^[ \t]*__COMMANDS_BEGIN__[ \t]*\r?\n.*?^__COMMANDS_END__[ \t]*\r?\n?",
                   "", t, flags=re.M)
    # https_port set = the caller validated a working loopback certificate
    # and EVERY pane rides the shared TLS listener on port+1 (0.9.204:
    # Word/PowerPoint included; MailKick requires that listener).
    scheme, pane_port = ("https", https_port) if https_port else ("http", port)
    host = (host or "localhost")
    if scheme == "https" and host != "localhost":
        # BYO certificate covering a custom name: EVERY https URL in the
        # manifest rides that hostname (before the scheme/port fill)
        t = t.replace("__SCHEME__://localhost", "https://%s" % host)
        t = t.replace("__SCHEME__://127.0.0.1", "https://%s" % host)
    t = t.replace("__SCHEME__", scheme).replace("__PORT__", str(pane_port))
    if scheme == "https":
        t = t.replace("http://localhost:%d" % port, "https://localhost:%d" % pane_port)
        t = t.replace("http://127.0.0.1:%d" % port, "https://127.0.0.1:%d" % pane_port)
    try:
        minidom.parseString(t)
    except Exception as e:
        die("rendered manifest for %s is not valid XML: %s" % (addin_dir, e))
    return t


def render_and_sideload():
    # 0.9.204: a certificate that is already present AND working (validated
    # by validate_existing_tls: expiry, localhost SAN, trusted chain) is
    # activated for EVERY pane - the manifests for Word/PowerPoint/MailKick
    # then point at the shared TLS listener on port+1. Only when no usable
    # certificate exists do the add-ins stay on plain http (except MailKick,
    # which requires the TLS prereqs and is skipped without them on macOS).
    tls_pair = None
    _crt = os.path.join(tls_cert_dir(), "server.crt")
    _key = os.path.join(tls_cert_dir(), "server.key")
    if os.path.isfile(_crt) and os.path.isfile(_key):
        reason = validate_existing_tls(_crt)
        if reason is None:
            tls_pair = (_crt, _key)
        else:
            warn("loopback certificate present but not usable (%s) - re-run" % reason)
            print("        install.py --install to regenerate and re-trust it.")
            _log("tls pair present but unusable: %s" % reason)
    transport = _current_transport()
    if transport == "http" and is_mac() and "mailkick" in (ADDINS or set()):
        warn("transport mode is plain http - Outlook cannot use it; switching"
             " MailKick off the list for this run")
        renderable = ((ADDINS or set()) - {"mailkick"})
        note("transport http: mailkick removed from this render (%s)" % ",".join(sorted(renderable)))
        ADDINS.clear()
        ADDINS.update(renderable)
    if transport == "https":
        cm = _get_cert_mode()
        if not tls_pair and (is_mac() or cm == "self_signed_leaf"):
            warn("transport mode https - the certificate check did not pass;"
                 " creating/renewing the certificate now")
            # route by the STORED MODE - the old platform route generated a CA
            # even in leaf mode (found in a live audit)
            if cm == "self_signed_leaf":
                pair = ensure_leaf_certs()
            elif cm == "ca":
                pair = ensure_tls_certs()
            else:
                pair = None    # bring-your-own cannot be generated automatically
            reason2 = None
            if pair:
                reason2 = validate_existing_tls(pair[0])
                if reason2 is None and not _pair_matches(*pair):
                    reason2 = "key does not match"
            if pair and reason2 is None:
                tls_pair = pair
            else:
                warn("https mode but no usable certificate - the manifests stay plain http"
                     " (the https listener will start once a pair validates)")
        elif not tls_pair and not is_mac():
            warn("transport mode https on macOS requires a certificate mode; on Windows"
                 " set the certificate mode to 'self_signed_leaf' first")
    step("rendering add-in manifests (%s, %s)" %
         (("https://localhost:%d via the loopback certificate" % (HTTPS_PORT if HTTPS_PORT else PORT + 1)) if tls_pair
          else ("http://localhost:%d" % PORT),
          "ribbon commands" if RIBBON else "list-mode"))
    if is_mac() and not DRY_RUN:
        if purge_stale_wef_manifests() < 0:   # Office running - dangling registrations win
            die("quit Office completely (Cmd+Q), then re-run install.py --update/--install")
    apps = office_apps_present()
    rendered = []
    addins = ADDINS or set()
    for (folder, name, mfname) in OURS:
        want = folder in addins and not (folder == "mailkick" and MAILKICK_CHOICE == "skip")
        if not want:
            # selected OUT this run (or the MailKick route failed): drop a
            # previous Outlook registration so the pane disappears
            if folder == "mailkick" and is_mac() and not DRY_RUN:
                wef = wef_dir()
                if wef and os.path.isfile(os.path.join(wef, mfname)):
                    try:
                        os.remove(os.path.join(wef, mfname))
                        _log("removed previous Outlook registration: %s" % mfname)
                    except OSError:
                        pass
            if folder == "mailkick" and not is_mac():
                mpath = os.path.join(manifest_dir(), mfname)
                if os.path.isfile(mpath):
                    try:
                        os.remove(mpath)
                        _log("removed previous Windows manifest: %s" % mpath)
                    except OSError:
                        pass
            continue
        if not os.path.isfile(repo_path(folder, "manifest.template.xml")):
            warn("%s not built in this checkout - skipped" % name)
            continue
        if folder == "mailkick" and is_mac() and tls_pair is None:
            warn("MailKick: no usable loopback certificate - skipped for Outlook")
            continue
        xml = render_manifest(folder, PORT,
                              https_port=(HTTPS_PORT if HTTPS_PORT else PORT + 1) if tls_pair else None,
                              host=_current_cert_host())
        mver = (re.search(r"<Version>([^<]+)</Version>", xml) or [None, "?"])[1]
        mpath = os.path.join(manifest_dir(), mfname)
        if DRY_RUN:
            _log("[dry-run] would render %s -> %s" % (name, mpath))
        else:
            with open(mpath, "w", encoding="utf-8") as fh:
                fh.write(xml)
            _log("rendered %s -> %s (identity version %s)" % (name, mpath, mver))
        rendered.append(name + " (id v" + mver + ")")
        if is_mac() and not DRY_RUN:
            sideload_mac(folder, mfname, apps.get(folder))
    if rendered:
        if is_mac():
            _log("mac tip: a saved deck throwing the 'Add-in Error' toast on open"
                 " carries an unresolvable store stamp - heal it with"
                 " python3 tools/repair_paneref.py <file.pptx|.docx|folder>")
        return "registered: " + ", ".join(rendered)
    return "nothing registered (details: install.log)"


def wef_dir(folder="mailkick"):
    """The macOS WEF sideload folder for one of our add-ins (first existing
    container variant); None when none exists."""
    cands = {"slidekick": ["com.microsoft.Powerpoint", "com.microsoft.powerpoint"],
             "wordkick": ["com.microsoft.word"],
             "mailkick": ["com.microsoft.Outlook", "com.microsoft.outlook"]}[folder]
    for cid in cands:
        d = os.path.join(hp("Library", "Containers"), cid, "Data", "Documents", "wef")
        if os.path.isdir(d):
            return d
    return None


def relax_windows_config_acl():
    """ssh-created config dirs can carry restrictive, non-inherited ACLs;
    hand the current user full control over the whole tree (per-user data)."""
    import subprocess as sp2
    d = config_dir()
    if not os.path.isdir(d):
        return
    user = "%s\\%s" % (os.environ.get("COMPUTERNAME", ""), os.environ.get("USERNAME", ""))
    rc, o = run_quiet(["icacls", d, "/grant", user + ":(OI)(CI)F", "/T", "/C"])
    ok("config ACL relaxed: %s (%s)" % (d, "icacls rc %d" % rc))


def sideload_windows_developer():
    """
    Per-add-in dev sideload on Windows: the new "Apps" dialog on current
    builds (16.0.19x) ignores Trust Center folder catalogs, but every Office
    app reloads the add-ins registered under HKCU ...\\Wef\\Developer at
    launch. Register each add-in there (default value = manifest path) and
    set RefreshAddins so the host re-reads on the next start.
    """
    import winreg
    base = r"Software\Microsoft\Office\16.0\Wef\Developer"
    try:
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, "RefreshAddins", 0, winreg.REG_DWORD, 1)
        n = 0
        for (folder, _name, mfname) in OURS:
            mpath = os.path.join(manifest_dir(), mfname)
            if not os.path.isfile(mpath):
                continue
            aid = id_of_manifest(mpath)
            if not aid:
                warn("no <Id> in %s - dev sideload skipped" % mfname)
                continue
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base + "\\" + aid, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, "", 0, winreg.REG_SZ, mpath)
            n += 1
        ok("dev sideload registered (HKCU %s): %d add-in(s) load at launch" % (base, n))
    except OSError as e:
        warn("dev sideload could not be written: %s" % e)


def id_of_manifest(path):
    try:
        with open(path, encoding="utf-8") as fh:
            m = re.search(r"<Id[^>]*>\s*([0-9A-Za-z\-]+)\s*</Id>", fh.read())
        return m.group(1).strip("{}") if m else None
    except OSError:
        return None


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
        step("[dry-run] would sideload %s -> %s" % (mfname, target))
        return
    try:
        os.makedirs(target, exist_ok=True)
        shutil.copyfile(os.path.join(manifest_dir(), mfname), os.path.join(target, mfname))
        out("  [ok]       %s sideloaded%s: %s" % (
            folder, "" if present is not False
            else " (Office app not detected - installed anyway)", target))
    except OSError as e:
        warn("could not sideload %s: %s" % (mfname, e))


def office_apps_running():
    import subprocess as _sp
    try:
        r = _sp.run(["pgrep", "-x", "Microsoft PowerPoint"], stdout=_sp.PIPE)
        ppt = r.returncode == 0
        r = _sp.run(["pgrep", "-x", "Microsoft Word"], stdout=_sp.PIPE)
        word = r.returncode == 0
        r = _sp.run(["pgrep", "-x", "Microsoft Outlook"], stdout=_sp.PIPE)
        outlook = r.returncode == 0
        return ppt, word, outlook
    except OSError:
        return (False, False, False)


def purge_stale_wef_manifests():
    # removing WEF manifests while a host is running leaves the host with a
    # dangling registration - documents that had the pane open then throw
    # "This add-in is no longer available" on every open until the add-in is
    # refreshed. Refuse loudly instead.
    if DRY_RUN:
        step("[dry-run] would purge stale WEF manifests (real files untouched)")
        return 0
    ppt, word, outlook = office_apps_running()
    running = [n for n, gone in (("PowerPoint", ppt), ("Word", word), ("Outlook", outlook)) if gone]
    if running:
        warn("Office is running (%s) - NOT purging/sideloading WEF manifests. "
             "Quit Office completely (Cmd+Q) and run the install again." % ", ".join(running))
        return -1
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
            # the value Office reads is `flag` (SINGULAR) - `flags` produced a
            # registry entry Office ignored and a French Office telling the
            # user to add/enable catalogs manually (field report 0.9.277)
            winreg.SetValueEx(k, "flag", 0, winreg.REG_DWORD, 1)
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
        step("[dry-run] would clear the Office add-in caches (Webkit/WEF stores)")
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
    note("caches cleared; install flows re-stage the manifests to the WEF dirs"
         " immediately after (render_and_sideload), so no manual re-add step is"
         " needed; MailKick rides the https://localhost:%d loopback listener and"
         " works in New Outlook too." % (HTTPS_PORT if HTTPS_PORT else PORT + 1))


def office_app_running():
    if is_mac():
        # only the MAIN app binaries count: Office keeps extension/widget
        # processes alive inside the .app bundles (e.g. Outlook's
        # CalendarWidgetExtension.appex) even after every app is closed, and
        # their command lines contain the bundle name too
        rc, o = run_quiet(["bash", "-c",
                           "pgrep -f 'Microsoft (PowerPoint|Word|Outlook)\\.app/Contents/MacOS/"
                           "Microsoft (PowerPoint|Word|Outlook)' >/dev/null 2>&1 && echo yes"])
        return rc == 0 and "yes" in o
    return False


# -------------------------------------------------------- autostart: macOS #
def build_app_bundle():
    a = app_bundle()
    out("")
    out("  Local server app: " + a)
    exe_dir = os.path.join(a, "Contents", "MacOS")
    if DRY_RUN:
        out("  [..]      [dry-run] would (re)write the bundle (http shim, no certificate flags),")
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
        step("[dry-run] would register the OfficeKickServer login item")
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
        out("  [..]      [dry-run] would create task '%s': %s" % (TASK_NAME, windows_task_cmd()))
        return True
    tr = windows_task_cmd()
    rc, o = run_quiet(["schtasks", "/create", "/f", "/sc", "onlogon",
                       "/tn", TASK_NAME, "/tr", tr])
    if rc == 0:
        # the 0.9.108-era task died on its own: "Stop the task if it runs
        # longer than 3 days" + idle-stop kill long-running servers. The
        # server must run until the machine shuts down.
        _hard_task_settings(TASK_NAME)
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
# -------------------------------------------------------------- uninstall #
def _mac_keychain_shas(cn):
    """SHA1 hashes of every login-keychain item under a common name.
    security accepts the keychain NAME, not a path: the full path form
    silently returns NOTHING (found when a CA item survived a removal)."""
    shas = []
    for kc in ("login.keychain", None):
        cmd = ["security", "find-certificate", "-a", "-p", "-c", cn]
        if kc:
            cmd.append(kc)
        _rc, pems = run_quiet(cmd)
        for blob in re.findall(r"-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----",
                               pems or "", re.S):
            b64 = "".join(l for l in blob.splitlines() if "-----" not in l)
            try:
                h = hashlib.sha1(base64.b64decode(b64)).hexdigest()
                if h not in shas:
                    shas.append(h)
            except Exception:
                pass
    return shas


def _remove_ca_trust(ca_crt=None):
    """Remove OUR loopback CA trust: user-scope trust settings plus the
    matching login-keychain items (identified by sha1 while the CA file is
    present, otherwise by our common name). Never touches other CAs."""
    if not is_mac():
        return True
    known = None
    if ca_crt and os.path.isfile(ca_crt):
        _rc, o = run_quiet(["openssl", "x509", "-in", ca_crt, "-noout",
                            "-fingerprint", "-sha1"])
        if _rc == 0:
            known = [o.strip().split("=")[-1].replace(":", "").lower()]
    if known:
        rc, o = run_quiet(["security", "remove-trusted-cert", ca_crt])
        _log("loopback CA trust settings removed (sha1=%s, rc %d)" % (known[0], rc), ts=True)
    gone = 0
    for h in _mac_keychain_shas("OfficeKick Loopback CA"):
        if known is None or h in known:
            rc, o = run_quiet(["security", "delete-certificate", "-Z", h, "login.keychain"])
            _log("CA keychain item delete (%s): rc %d %s" % (h[:12], rc, o or ""), ts=True)
            gone += 1
    _log("loopback CA keychain items removed: %d" % gone, ts=True)
    return True


# ---------------------------------------------------------------- uninstall #
def do_uninstall(purge, rm_certs=None):
    """Remove the add-ins + server. Both data asks default to KEEP so an
    uninstall never silently destroys the user's setup:
      - certificates: ask to remove the local certificates we installed
        (trust entry in the keychain/TrustedPeople + the files on disk);
        a BYO certificate's own trust entry is never touched
      - config: ask to wipe ~/.officekick (config.json, API keys, profiles,
        sessions); keeping it means install.py sees the old setup again on
        the next install
    --purge-config / --remove-certs force the removals non-interactively;
    interactive runs ask with the default on 'keep'."""
    out("")
    out("  Uninstalling OfficeKick")
    meta = _leaf_meta_load()
    cdir = tls_cert_dir()
    crt_p = _leaf_paths()[1]
    ca_p = os.path.join(cdir, "ca.crt")
    byo = str(meta.get("mode") or "") == "byo"
    ca_present = os.path.isfile(ca_p)
    leaf_present = os.path.isfile(crt_p)
    anything_certs = ca_present or leaf_present

    def _describe_certs():
        if ca_present and leaf_present:
            return "the loopback CA + localhost certificate"
        if ca_present:
            return "the loopback CA"
        if byo and leaf_present:
            return "the certificate FILES we copied (%s)" % crt_p
        return "the leaf certificate"

    def _describe_trust():
        if byo:
            return "(your own trust entry stays in place)"
        return "(keychain/TrustedPeople trust entry + files)" if (is_mac() or is_windows()) \
            else "(files)"

    stop_server()

    # Questions BEFORE the actions they belong to (default = keep everything)
    if rm_certs is None:
        if DRY_RUN or (not sys.stdin.isatty()):
            rm_certs = False
            if anything_certs and DRY_RUN:
                out("  [..]      [dry-run] would ask to remove %s %s - default: KEEP"
                    % (_describe_certs(), _describe_trust()))
        elif anything_certs:
            print("")
            out("  certificates in place: %s" % os.path.join(cdir, ""))
            out("    %s" % _paint(_describe_certs(), "cyan"))
            print("")
            rm_certs = ask_yes("Also remove %s %s" % (_describe_certs(), _describe_trust()), False)
            _log("uninstall cert-removal answer: %s" % ("yes" if rm_certs else "no (kept)"), ts=True)
            print("")
        else:
            rm_certs = False
    if purge is None:
        if DRY_RUN or (not sys.stdin.isatty()):
            purge = False
            if DRY_RUN:
                step("[dry-run] would ask to wipe the config folder - default: KEEP")
        else:
            purge = ask_yes("Also wipe the config folder (%s - config.json, stored"
                            " API keys, profiles, sessions%s)?" % (
                                config_dir(),
                                "; includes the certificate + log files" if anything_certs else ""),
                            False)
            _log("uninstall config-wipe answer: %s" % ("yes" if purge else "no (kept)"), ts=True)
    if rm_certs and purge:
        note("the config wipe covers the certificate files; the trust removal still runs")

    # --- platform unregistration (always) ---
    if is_mac():
        if DRY_RUN:
            step("[dry-run] would unregister the login item and remove the server app")
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
            step("[dry-run] would remove the task, the Run key and the trusted catalog")
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
        step("[dry-run] would remove " + inst)
    else:
        if os.path.isdir(inst):
            shutil.rmtree(inst, ignore_errors=True)
            ok("removed " + inst)

    # --- certificates (only when the ask/flag said so) ---
    if rm_certs and os.path.isdir(cdir):
        if byo:
            warn("your own certificate stays TRUSTED in the trust store - we never"
                 " touch trust entries we did not create (remove it yourself if you"
                 " no longer want it)")
        else:
            meta_sha = str(meta.get("sha1") or "")
            meta_mode = str(meta.get("mode") or "")
            leaf_trusted_before = os.path.isfile(_leaf_meta_path()) or \
                (meta_sha and meta_mode == "self_signed_leaf")
            if ca_present:
                _remove_ca_trust(ca_p)
                ok("loopback CA removed from the user trust store")
            elif os.path.isdir(cdir) or leaf_trusted_before:
                # CA file already gone (manual cleanup) - sweep our CA by its
                # common name; unrelated CAs are never touched
                _remove_ca_trust(None)
            if leaf_present:
                _remove_leaf_trust(crt_p)
                ok("leaf certificate removed from the user trust store")
            elif meta_mode == "self_signed_leaf" and meta_sha:
                # leaf files already gone but its trust entry may linger -
                # _remove_leaf_trust falls back to the stored sha1
                _remove_leaf_trust(crt_p)
                ok("leaf trust entry removed via the stored fingerprint")
            try:
                if os.path.isfile(_leaf_meta_path()):
                    os.remove(_leaf_meta_path())   # the identified entry is gone
            except OSError:
                pass
        if DRY_RUN:
            step("[dry-run] would delete " + cdir)
        else:
            shutil.rmtree(cdir, ignore_errors=True)
            if os.path.isdir(cdir):
                warn("could not fully remove " + cdir)
            else:
                ok("certificate + key files deleted: " + cdir)
    elif not rm_certs and anything_certs:
        out("  [ok]       certificates + keys kept at %s (trust untouched)" % cdir)

    # --- config (wipe or keep) ---
    if purge:
        if DRY_RUN:
            step("[dry-run] would wipe " + config_dir())
        else:
            shutil.rmtree(config_dir(), ignore_errors=True)
            ok("config wiped: " + config_dir())
    else:
        out("  [ok]       config kept at %s (config.json, keys, profiles and sessions intact)"
            % config_dir())
        if _set_cert_mode("self_signed_leaf"):
            out("  [ok]       certificate mode reset to self-signed leaf (the default)")
        else:
            warn("could not reset the certificate mode to self-signed leaf"
                 " - check the config file")
    out("")
    warn("quit and reopen PowerPoint / Word / Outlook to finish removing the add-ins")
    pause_after("uninstall")


# ---------------------------------------------------------------- install #
def preflight(mode):
    """Console: nothing (the banner carries status). The transport explanation
    and the popup essays live in install.log only."""
    note("preflight (%s); transport: panes -> http://localhost:%d (Office allows"
         " localhost over http%s).%s" % (
             mode, PORT,
             "; plus the https://localhost:%d loopback TLS listener wherever a"
             " working certificate exists" % (HTTPS_PORT if HTTPS_PORT else PORT + 1),
             " Outlook on macOS validates every manifest URL as https, so"
             " MailKick must ride the TLS component." if is_mac() else ""))
    if is_mac():
        note("popups you may see: (1) one-time background-item note when the"
             " autostart login item is first registered - that IS the local"
             " server app; (2) a Local Network prompt the first time a model"
             " on your LAN is used - click Allow so panes can reach e.g."
             " Ollama on another machine; (3) a certificate-trust confirm when"
             " the local certificate is imported into the keychain.")
    else:
        note("popups you may see: a one-time Windows Defender Firewall ask"
             " about Python listening - Allow keeps LAN models (Ollama)"
             " reachable; no SmartScreen: nothing is downloaded; the add-ins"
             " come from the local trusted catalog.")


def ask_autostart():
    if is_mac():
        return ask_yes("Add autostart to the login items (hidden 'OfficeKickServer'"
                       " app, starts the server at every login)?", True)
    return ask_yes("Add an autostart scheduled task (windowless pythonw at logon)?", True)


def install(mode_name="Install", interactive=True):
    global INTERACTIVE, AUTOSTART_WANT, MAILKICK_CHOICE
    INTERACTIVE = interactive
    AUTOSTART_WANT = None
    banner(mode_name)
    preflight(mode_name)
    # add-in selection FIRST (console: one ok-line; details in the log)
    select_addins(interactive)
    total = 6
    phase(1, total, "preparing", prepare_phase)
    phase(2, total, "deploying", deploy_files)
    # questions BEFORE the actions they belong to: the TLS route asks now
    # (the certificate phase acts on the answer), the autostart ask sits
    # right before its own phase
    if "mailkick" in (ADDINS or set()) and is_mac() and MAILKICK_CHOICE is None:
        if DRY_RUN or (not sys.stdin.isatty()):
            MAILKICK_CHOICE = pick_mailkick_route()   # automation: reuse-first
        else:
            r = _ask_cert_setup()        # three-way: leaf / CA / BYO (asked BEFORE the phase)
            _crt = os.path.join(tls_cert_dir(), "server.crt")
            _key = os.path.join(tls_cert_dir(), "server.key")
            if r or (os.path.isfile(_crt) and os.path.isfile(_key)
                     and _pair_matches(_crt, _key)):
                MAILKICK_CHOICE = "auto"  # pair ready; the certificates phase validates it
            else:
                MAILKICK_CHOICE = "skip"  # BYO given up -> Outlook not registered
            _print_cert_facts()
    phase(3, total, "certificates", tls_phase)
    _print_cert_facts()
    phase(4, total, "registering with Office", registration_phase)
    if INTERACTIVE and not DRY_RUN:
        AUTOSTART_WANT = ask_autostart()
    phase(5, total, "autostart", autostart_phase)
    phase(6, total, "starting the server", server_phase)
    done_note(mode_name)
    pause_after(mode_name)


def _drain_stdin():
    """Discard pending terminal input (stale line-buffered leftovers like a
    just-typed answer's newline) so a key-press pause waits for a genuine
    press instead of releasing instantly on old bytes."""
    try:
        if is_windows():
            import msvcrt
            while msvcrt.kbhit():
                msvcrt.getwch()
        else:
            import termios
            termios.tcflush(sys.stdin.fileno(), termios.TCIFLUSH)
    except Exception:
        pass


def pause_after(label):
    """Hold the screen after a flow until a key press (any key) - a real
    terminal wait, no timer and no auto-advance. Pending buffered input is
    flushed first, so a stale newline cannot release it. Piped/automation
    runs and scheduled tasks just continue."""
    _log("flow ended: %s" % label)
    try:
        if not (sys.stdin.isatty() and sys.stdout.isatty()):
            return
        _drain_stdin()
        sys.stdout.write("  " + _paint("(press any key to continue - results above;"
                                       " details in install.log - Ctrl-C quits)", "dim"))
        sys.stdout.flush()
        try:
            k = _read_key()
            _log("pause released (key %r)" % (k,))
        except EOFError:
            _log("pause released (EOF)")
        except KeyboardInterrupt:
            _log("pause released (Ctrl-C -> quit)")
            print("")
            raise SystemExit(130)
        print("")
    except Exception:
        pass


def prepare_phase():
    stop_server()          # first: free our canonical port (a previous,
    resolve_port()         # https-only install answers no http health check)
    return "port %d free, previous server stopped" % PORT


def tls_phase():
    """Certificate work: the route was asked before this phase line; here we
    only act on it. The cert check runs whenever TLS is in use (existing pair
    present) - not tied to the Outlook toggle - and reuses / reissues /
    regenerates as needed."""
    global CERT_FACTS, MAILKICK_CHOICE
    CERT_FACTS = {}
    if _get_cert_mode() == "byo":
        host = _current_cert_host()
        crt = os.path.join(tls_cert_dir(), "server.crt")
        key = os.path.join(tls_cert_dir(), "server.key")
        if os.path.isfile(crt) and os.path.isfile(key):
            good, problems, f = _validate_byo(crt, key, host)
            if good:
                CERT_FACTS.update({"mode": "your own certificate (BYO, host %s)" % host,
                                   "leaf_path": crt, "leaf_line": _fact_line(f),
                                   "ca_path": tls_cert_dir(), "ca_line": "n/a in BYO mode",
                                   "key_line": "server.key (private, protected)"})
                return "your own certificate for %s (validated)" % host
            for pr in problems:
                warn(pr)
            return "BYO certificate unusable - re-run install.py (BYO setup); see install.log"
        return "BYO certificate missing on disk - re-run the BYO setup"
    if not is_mac() and _current_transport() != "https":
        if "mailkick" in ADDINS:
            return "not needed on Windows (plain-http localhost)"
        return "no TLS on Windows (plain http)"
    # Windows + https transport: the same leaf/CA flow as mac
    # (the leaf generator self-installs the 'cryptography' package there)
        # generator self-installs the 'cryptography' package there)
    crt = os.path.join(tls_cert_dir(), "server.crt")
    key = os.path.join(tls_cert_dir(), "server.key")
    have_pair = os.path.isfile(crt) and os.path.isfile(key)
    if _get_cert_mode() == "self_signed_leaf":
        pair = ensure_leaf_certs()
        if pair:
            return "self-signed leaf ready (%s)" % CERT_FACTS["mode"]
        return "leaf certificate NOT trusted/ready - https will not validate (see install.log)"
    if "mailkick" in ADDINS:
        if MAILKICK_CHOICE not in ("auto", "custom", "skip"):
            MAILKICK_CHOICE = pick_mailkick_route()   # non-interactive fallback
        if MAILKICK_CHOICE == "custom":
            cdir = tls_cert_dir()
            os.makedirs(cdir, exist_ok=True)
            shutil.copyfile(MAILKICK_CERT, os.path.join(cdir, "server.crt"))
            shutil.copyfile(MAILKICK_KEY, os.path.join(cdir, "server.key"))
            _protect_secret(os.path.join(cdir, "server.key"))
            f = _cert_facts(crt)
            CERT_FACTS.update({"mode": "your own certificate (PEM, validated + copied)",
                               "ca_path": os.path.join(cdir, "ca.crt"),
                               "ca_line": "yours (unchanged by this install, if present)",
                               "leaf_path": crt, "leaf_line": _fact_line(f),
                               "key_line": "server.key in %s (private, chmod 0600)" % cdir})
            note("custom certificate copied to %s (leaf %s; key %s)"
                 % (crt, CERT_FACTS["leaf_line"], CERT_FACTS["key_line"]))
            return "your own certificate (copied, validated)"
        if MAILKICK_CHOICE == "skip":
            return "skipped (Outlook will not be registered)"
        pair = ensure_tls_certs()
        if pair:
            return "loopback CA + localhost certificate ready (%s)" % CERT_FACTS["mode"]
        return "NOT ready - Outlook registration will be skipped (see install.log)"
    if have_pair:
        pair = ensure_tls_certs()   # https still in use for the other panes
        if pair:
            return "loopback TLS re-checked (%s; used by Word/PowerPoint)" % CERT_FACTS["mode"]
        warn("https is no longer usable - the add-ins fall back to plain http"
             " (re-run install.py --install and enable Outlook to repair TLS)")
        return "existing TLS unusable - plain http fallback (see install.log)"
    if _current_transport() == "https" and is_mac():
        pair = ensure_tls_certs()
        if pair:
            return "loopback TLS created (%s; https transport mode)" % CERT_FACTS["mode"]
        warn("https transport mode on, but the certificate could not be created"
             " - the manifests stay plain http")
        return "https requested - certificate setup failed (plain http fallback)"
    return "none needed - plain http (no Outlook selected, no TLS in use)"


def registration_phase():
    """Windows ACL fix + cache flush + manifest render + sideload."""
    if is_windows():
        relax_windows_config_acl()
    if office_app_running():
        warn("Office apps are running - skipping the add-in cache clear"
             " (quit Office, re-run this install, reopen the apps to fully"
             " flush an old registration)")
    else:
        clear_office_caches()
    ret = render_and_sideload()
    if is_windows() and not DRY_RUN:
        # 0.9.278: the Windows registrations had become DEAD CODE during the
        # installer restructuring - both functions were defined but never
        # called (field: a fresh box deployed the manifests yet had NO
        # Wef\Developer key and NO TrustedCatalogs entry, so the Word
        # add-in list stayed empty no matter what). Run both, every
        # Windows registration, unconditionally: the dev-sideload entries
        # load the panes at launch; the Trust Center catalog enables the
        # classic dialog's Shared Folder tab.
        sideload_windows_developer()
        sideload_windows_catalog()
    return ret


def autostart_phase():
    """Register the autostart with the answer given BEFORE this phase line,
    or the non-interactive default (register/keep)."""
    global AUTOSTART_WANT
    if DRY_RUN:
        AUTOSTART_WANT = True
        return "[dry-run] would register"
    want = AUTOSTART_WANT
    if want is None:
        present = mac_login_item_present() if is_mac() else windows_autostart_present()
        want = True
        note("autostart (non-interactive): existing=%s -> kept/re-registered"
             % ("yes" if present else "no"))
    if not want:
        note("autostart declined - manual start: %s %s --config %s --static-dir %s --port %d"
             % (sys.executable, server_py(), config_file(), install_dir(), PORT))
        return "not added (manual start command: install.log)"
    if is_mac():
        autostart_mac()
        return "login item registered"
    autostart_windows()
    return "scheduled task registered"


def server_phase():
    """Start the server (quiet) and report the resulting health."""
    start_server()
    h = health(PORT)
    if h:
        return "running on http://127.0.0.1:%d (version %s)" % (PORT, h.get("version"))
    return "(not answering yet - check install.log)"


def done_note(mode_name):
    apps = office_apps_present()
    verb = "updated for" if mode_name == "Update" else "installed for"
    out("")
    for (folder, name, _mf) in OURS:
        if folder in (ADDINS or set()) and apps.get(folder):
            ok("%s %s %s" % (name, verb, app_label(folder)))
    out("")
    out("  Quit and REOPEN your Office apps. When a task pane is open, its")
    out("  About tab shows 'Pane build: %s' - to open one use Insert >" % read_version())
    out("  Add-ins > My Add-ins (or Add from File when offered); the per-app")
    out("  steps live in docs/<dockick|slidekick|mailkick>.md. Log: install.log")
    out("  Simulator (no Office needed): %s/slidekick/static/simulator.html" % _sim_base())
    _first_run_notes()
    check_pane_errors()


def _first_run_notes():
    """After installing an add-in for the FIRST time, show what to expect and
    how to surface the pane (Outlook needs one manual step on some builds)."""
    if DRY_RUN or not (FIRST_RUN and set(FIRST_RUN) & set(ADDINS or set())):
        return
    apps = office_apps_present()
    names = [app_label(f) for f in sorted(FIRST_RUN) if f in (ADDINS or set())]
    out("")
    print("  " + _paint("First-run notes (%s):" % ", ".join(names), "bold"))
    if is_mac():
        print("  - macOS may ask ONCE to allow the 'OfficeKickServer' background item")
        print("    and (when a model on your LAN is used) a Local Network prompt -")
        print("    click Allow; both belong to this local server, nothing else is installed.")
    else:
        print("  - Windows may ask ONCE whether Python may listen (firewall) -")
        print("    click Allow; the add-ins talk to 127.0.0.1 only, but Allow keeps")
        print("    LAN models like Ollama working later.")
    if "mailkick" in (FIRST_RUN and (ADDINS or set())) and apps.get("mailkick"):
        mpath = os.path.join(manifest_dir(), "mailkick.xml")
        if is_mac():
            print("  - Outlook: the pane is staged in Outlook's WEF folder already;")
            print("    if it does not appear after a restart, add it once via")
            print("    Get Add-ins > My add-ins > Custom Add-ins > Add from File (Cmd+Shift+G):")
        else:
            print("  - Outlook: add the pane once via Get Add-ins > My add-ins >")
            print("    Custom Add-ins > Add a custom add-in > Add from File:")
        print("        %s" % mpath)
        print("    (the plain-http/localhost listener works; the https one covers New Outlook)")
    note("first-run notes shown for: %s" % ", ".join(sorted(FIRST_RUN)))


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


def check_pane_errors():
    """Point the user at the pane error log when panes crashed today (0.9.182):
    every uncaught task-pane error is appended to
    <config>/logs/pane-errors-YYYY-MM-DD.log by the server - the task pane's
    webview console is invisible from the desktop, so this is the paper trail."""
    import json as _json
    try:
        logdir = os.path.join(config_dir(), "logs")
        if not os.path.isdir(logdir):
            return
        import datetime
        cand = os.path.join(logdir, "pane-errors-" + datetime.datetime.now().strftime("%Y-%m-%d") + ".log")
        if not os.path.isfile(cand):
            return
        with open(cand, "r", encoding="utf-8", errors="replace") as fh:
            rows = [x for x in fh.read().splitlines() if x.strip()]
        if not rows:
            return
        try:
            last = _json.loads(rows[-1])
            laststr = "%s  %s: %s" % (last.get("ts", ""), last.get("pane", "?"),
                                      str(last.get("message", ""))[:110])
        except Exception:
            laststr = rows[-1][:130]
        warn("panes crashed today - %d entr%s in the pane error log" %
             (len(rows), "y" if len(rows) == 1 else "ies"))
        out("       %s" % cand)
        out("       last: %s" % laststr)
    except Exception:
        pass


# ------------------------------------------------------------------- flows #
def _cert_mode_menu():
    """Switch the certificate MODE (self-signed leaf / loopback CA /
    bring-your-own) used for the next https certificate setup. A stored mode
    only takes effect while no working pair blocks the question: with a pair
    on disk the user may switch immediately (existing pair kept until the
    regeneration consent)."""

    def rows():
        cur = _get_cert_mode()

        def _mark(m):
            return "  <-- current" if cur == m else ""
        have = os.path.isfile(_leaf_paths()[1])
        return [
            ("1", "self-signed leaf   (no CA created - the LEAF itself is imported"
                  " + trusted in your keychain)%s" % _mark("self_signed_leaf"), None),
            ("2", "loopback CA        (the CA is created + trusted in your keychain;"
                  " silent renewals afterwards)%s" % _mark("ca"), None),
            ("3", "bring-your-own     (your own already-trusted PEM pair; the"
                  " hostname for the manifests is configurable)%s" % _mark("byo"), None),
            (None, "@SEP", None),
            ("0", "back (mode stays: %s%s)" % (cur,
                ", certificate pair on disk" if have else ", no certificates on disk"),
             "magenta"),
        ]

    while True:
        clear_screen()
        r = rows()
        cm = _get_cert_mode()
        mode_row = {"self_signed_leaf": 0, "ca": 1, "byo": 2}.get(cm, 0)
        sel = menu_pick(r, mode_row, "(1-3 = set the certificate mode, 0 = back)")
        if sel is None:
            # typed fallback
            print("")
            for (act, label, style) in r:
                if label == "@" "SEP":
                    print("")
                    continue
                print("  %s) %s" % (act, label) if act else "    %s" % label)
            print("")
            try:
                if not sys.stdin.isatty():
                    print("  (non-interactive: PUT /api/settings {certMode: ...} or re-run install.py)")
                    return
                choice = input("  > choose [1/2/3, 0] (%s is current; Enter = keep): "
                               % {"self_signed_leaf": "1", "ca": "2", "byo": "3"}.get(
                                   _get_cert_mode(), "0")).strip().lower()
            except EOFError:
                return
            if choice in ("0", "q", "back"):
                return
            if choice == "":
                choice = {"self_signed_leaf": "1", "ca": "2", "byo": "3"}.get(
                    _get_cert_mode(), "0")   # Enter = keep the current mode
            sel2 = next((i for i, (a, l, s) in enumerate(r) if a == choice), None)
            if sel2 is None:
                print("  please choose 1, 2, 3 or 0")
                continue
            sel = sel2
        act = r[sel][0]
        if not act or act == "0":
            return
        new_mode = {"1": "self_signed_leaf", "2": "ca", "3": "byo"}[act]
        cur = _get_cert_mode()
        if new_mode == cur:
            note("certificate mode already %s" % cur)
            continue
        _set_cert_mode(new_mode)
        ok("certificate mode stored: %s" % new_mode)
        crt0 = _leaf_paths()[1]
        have_pair = os.path.isfile(crt0) and _pair_matches(
            crt0, os.path.join(tls_cert_dir(), "server.key"))
        if have_pair:
            go = ask_yes("Switch the certificates to '%s' NOW" % new_mode, False)
            if go:
                shutil.rmtree(tls_cert_dir(), ignore_errors=True)
                got = _ask_cert_setup()
                if got:
                    try:
                        _reapply_transport(_current_transport(), "cert-mode-switch")
                        ok("server restarted on the new certificate setup")
                    except Exception as exc:
                        warn("restart after the certificate switch failed: %s"
                             " (re-run install.py)" % exc)
            else:
                note("the existing pair stays in use until the next install"
                     " (re-running install.py converts it)")
        else:
            note("no certificates on disk yet - the mode applies when the installer"
                 " builds them (run install.py)")


def transport_menu():
    cur = _current_transport()
    hp, sp = _current_ports()

    def rows():
        return [
            ("1", "auto    (https when a working certificate exists - the classic default)", None),
            ("2", "plain http only (certificates ignored; Outlook not usable with this)", None),
            ("3", "https always (https for ALL panes; a local self-signed certificate"
                  " is created/renewed - leaf, CA or your own)", None),
            (None, "@SEP", None),
            ("4", "set the http port   (current: %d)" % hp, None),
            ("5", "set the https port  (current: %d)" % sp, None),
            ("6", "certificate mode   (current: %s - leaf no-CA / CA / bring-your-own)" % _get_cert_mode(), None),
            (None, "@SEP", None),
            ("0", "back (mode stays: %s, ports http %d / https %d)" % (cur, hp, sp), "magenta"),
        ]

    while True:
        clear_screen()
        cur = _current_transport()
        hp, sp = _current_ports()
        h = health(PORT)
        print("")
        print("  ----- transport -----")
        print("  current mode: %s" % cur)
        print("  listener ports: http %d / https %d   (server: %s)" % (
            hp, sp, ("running v%s" % h.get("version")) if h else "not running"))
        r = rows()
        # the cursor AND the typed bare-Enter default sit on the CURRENT mode
        mode_row = {"auto": 0, "http": 1, "https": 2}.get(cur, 0)
        sel = menu_pick(r, mode_row, "(up/down + Enter, 1-6 direct, 0 = back)")
        if sel is None:
            print("")
            for (act, label, _st) in r:
                if label == "@SEP":
                    print("")
                    continue
                print("    %s" % label if not act else "    %s) %s" % (act, label))
            print("")
            try:
                if not sys.stdin.isatty():
                    print("  (non-interactive: pass --transport auto|http|https)")
                    return
                choice = input("  > choose [1/2/3/4/5/6, 0; default 0]: ").strip().lower()
            except EOFError:
                return
            if choice in ("", "0", "q", "back"):
                return
            mode_choice = {"auto": "1", "http": "2", "https": "3"}.get(choice)
            if choice == "":
                choice = {"auto": "1", "http": "2", "https": "3"}.get(
                    cur, "0")   # bare Enter = keep the current (pre-selected) mode
            sel = {"1": 0, "2": 1, "3": 2, "4": 3, "5": 4, "6": 5}.get(choice)
            if sel is None:
                print("  please choose 1, 2, 3, 4, 5, 6 or 0")
                continue
        act = r[sel][0]
        if act == "0" or not act:
            return
        if act == "4" or act == "5":
            what = "http" if act == "4" else "https"
            ccur = hp if act == "4" else sp
            other = sp if act == "4" else hp
            print("")
            while True:
                ans = input("  " + _paint("? %s port (1024-65535, current %d): " % (what, ccur), "cyan")).strip()
                if not ans:
                    break
                if not _valid_port(ans):
                    print("  please enter a number in 1024-65535")
                    continue
                nport = int(ans)
                if nport == other:
                    warn("the https port must differ from the http port (%d)" % other)
                    continue
                nmode = _current_transport()
                if _set_transport(nmode, http_port=(nport if act == "4" else hp),
                                  https_port=(nport if act == "5" else sp)):
                    global HTTPS_PORT
                    _hp, _sp = _current_ports()
                    HTTPS_PORT = _sp
                    ok("%s port persisted: %d - re-rendering the manifests and restarting the server" % (what, nport))
                    _reapply_transport(nmode, act)
                    break
                continue
            continue
        if act == "6":
            _cert_mode_menu()
            continue
        mode = {"1": "auto", "2": "http", "3": "https"}[act]
        _reapply_transport(mode, act)
        cur = _current_transport()


def _reapply_transport(mode, label=None):
    """Store the mode (+ optional ports). When the server component is
    installed, the manifests are re-rendered and the server restarted so the
    change applies right away. When NOTHING is installed yet, ONLY the mode
    is stored - no deploy, no WEF staging, no server start (it applies at the
    next install)."""
    ok("transport mode -> %s" % mode)
    _set_transport(mode)
    installed = os.path.isfile(server_py()) and os.path.isdir(install_dir())
    if not installed:
        note("nothing is installed yet - the %s mode is stored and applies when"
             " the add-ins are installed" % mode)
        return
    if mode == "https" and is_mac():
        crt0 = os.path.join(tls_cert_dir(), "server.crt")
        if INTERACTIVE and not (os.path.isfile(crt0) and validate_existing_tls(crt0) is None
                                and _pair_matches(crt0, os.path.join(tls_cert_dir(), "server.key"))):
            _ask_cert_setup()    # the three-way question (leaf / CA / BYO)
        else:
            tls_phase()          # create/renew the pair when it is missing
        _print_cert_facts()
    if office_app_running():
        warn("Office apps are running - skipping the cache flush (the manifest"
             " render still runs)")
    else:
        clear_office_caches()
    render_and_sideload()
    stop_server()
    start_only()
    ok("done - the panes now load over %s (manifests re-rendered)" % mode)


def do_status():
    h = health(PORT)
    if h:
        print("  running  : http://127.0.0.1:%d (version %s)" % (PORT, h.get("version")))
        tls_on, tls_desc = _https_state()
        print("  https    : %s" % tls_desc)
        inst = detect_installed()
        print("  add-ins  : %s" % (", ".join(name for (folder, name, _mf) in OURS if folder in inst)
                                or "not installed"))
        check_pane_errors()
        return 0
    print("  not running on port %d (try: install.py --start)" % PORT)
    return 1


def server_menu():
    rows = [
        ("1", "start", None),
        ("2", "restart", None),
        ("3", "stop", None),
        ("4", "status", None),
        (None, "@SEP", None),
        ("0", "back", "magenta"),
    ]

    def body():
        clear_screen()
        h = health(PORT)
        print("")
        print("  ----- server -----")
        if h:
            print("  running v%s on http://127.0.0.1:%d" % (h.get("version"), PORT))
            tls_on, tls_desc = _https_state()
            print("  https: %s" % tls_desc)
        else:
            print("  not running")
        return rows

    while True:
        r = body()
        sel = menu_pick(r, 0, "(up/down + Enter, 1-4 = direct, q = back)")
        if sel is None:
            # typed fallback
            print("")
            for (act, label, _st) in r:
                if label == "@SEP":
                    print("")
                    continue
                print("    %s) %s" % (act, label))
            print("")
            try:
                if not sys.stdin.isatty():
                    print("  (non-interactive: pass --start/--stop/--restart/--status)")
                    return
                choice = input("  > choose [1/2/3/4, 0; default 0]: ").strip().lower()
            except EOFError:
                return
            _log("server menu input: %r" % choice, ts=True)
            if choice in ("", "0", "q", "x", "back"):
                return
            selrows = {"1": 0, "2": 1, "3": 2, "4": 3, "5": 4}
            sel = selrows.get(choice)
            if sel is None:
                print("  please choose 1, 2, 3, 4, 5 or 0")
                continue
        act = r[sel][0]
        _log("server menu selection: %s" % act, ts=True)
        print("")
        if act == "1":
            banner("start")
            start_only()
            pause_after("start")
        elif act == "2":
            banner("restart")
            stop_server()
            start_only()
            pause_after("restart")
        elif act == "3":
            banner("stop")
            stop_server()
            ok("server stopped (port %d)" % PORT)
            pause_after("stop")
        elif act == "4":
            do_status()
            pause_after("status")
        else:
            return


def main():
    global PORT, DRY_RUN, RIBBON
    ap = argparse.ArgumentParser(
        description="OfficeKick cross-platform installer (macOS + Windows)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Examples:  python3 install.py  (Windows: py -3 install.py)  |  "
               "python3 install.py --update  |  python3 install.py --uninstall --purge-config  |  "
               "python3 install.py --install --mailkick")
    ap.add_argument("--install", action="store_true",
                    help="non-interactive install: Word + PowerPoint (Outlook only with --mailkick / --mailkick-cert+--mailkick-key, or when Outlook was previously installed)")
    ap.add_argument("--mailkick", action="store_true",
                    help="with --install: also install Outlook (creates the local self-signed certificate in leaf mode when none exists)")
    ap.add_argument("--update", action="store_true", help="git pull + update (non-interactive)")
    ap.add_argument("--uninstall", action="store_true", help="remove the add-ins + server")
    ap.add_argument("--start", action="store_true", help="start the local server only (no install changes)")
    ap.add_argument("--restart", action="store_true", help="stop the local server, then start it again")
    ap.add_argument("--skip-mailkick", action="store_true",
                    help="do not install MailKick in Outlook this run (Word/PowerPoint install normally)")
    ap.add_argument("--mailkick-cert", default="", metavar="PEM",
                    help="use YOUR OWN already-trusted TLS certificate for MailKick's https listener (PEM, SAN must cover localhost)")
    ap.add_argument("--mailkick-key", default="", metavar="PEM",
                    help="matching unencrypted private key for --mailkick-cert (PEM)")
    ap.add_argument("--stop", action="store_true", help="stop the local server")
    ap.add_argument("--status", action="store_true", help="show whether the local server is running")
    ap.add_argument("--remove-certs", action="store_true",
                    help="with --uninstall: also remove the local certificates OfficeKick installed (trust entry + files); otherwise they are kept")
    ap.add_argument("--purge-config", action="store_true",
                    help="with --uninstall: also wipe ~/.officekick")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT, help="server port (default %d)" % DEFAULT_PORT)
    ap.add_argument("--ribbon", action="store_true",
                    help="force-include the AI ribbon commands")
    ap.add_argument("--no-ribbon", action="store_true",
                    help="Windows: leave the commands out (list-mode only)")
    ap.add_argument("--dry-run", action="store_true", help="print the actions, run none")
    ap.add_argument("--transport", default=None, choices=["auto", "http", "https"], metavar="MODE",
                    help="persist the pane transport: auto (https when a valid certificate exists) | http (plain http, certificates ignored) | https (always https; creates the certificate) - updates keep the stored mode")
    ap.add_argument("--cert-check", action="store_true",
                    help="run the local certificate diagnostics (SAN, CA:FALSE, key match, validity, handshake, platform trust)")
    args = ap.parse_args()
    global SKIP_MAILKICK, MAILKICK_CERT, MAILKICK_KEY, MAILKICK_OPT
    SKIP_MAILKICK = args.skip_mailkick
    MAILKICK_CERT = args.mailkick_cert
    MAILKICK_KEY = args.mailkick_key
    MAILKICK_OPT = args.mailkick
    PORT = args.port if 1024 <= args.port <= 65535 else DEFAULT_PORT
    DRY_RUN = args.dry_run
    if args.port == DEFAULT_PORT:
        # the persisted listener ports (explicit --port still wins)
        global HTTPS_PORT
        _hp, _sp = _current_ports()
        if _hp != PORT:
            _hp = _hp if health(PORT) is None else PORT   # a running server keeps its port
        if _valid_port(_hp):
            PORT = _hp
        HTTPS_PORT = _sp if _valid_port(_sp) and _sp != PORT else PORT + 1
    # ribbon commands: opt-in on both platforms. macOS: 16.114 renders
    # nothing from a commands manifest (0.9.29). Windows 16.0.19x (new Apps
    # dialog era): a commands manifest vanishes the add-in from BOTH the
    # Apps dialog and the classic dialog (proven on the VM, 2026-11-02) -
    # list-mode + the Shared Folder dialog = the working path.
    RIBBON = bool(args.ribbon)   # opt-in on BOTH platforms now (see 0.9.107 revert note below)

    if args.transport:
        ok("transport mode -> %s (persisted; updates keep it)" % args.transport)
        _set_transport(args.transport)

    if not (is_mac() or is_windows()):
        die("this installer supports macOS and Windows only (detected: %s)." % sys.platform)

    ensure_runtime_deps()   # dependencies at STARTUP, never mid-flow (0.9.275)

    if args.install:
        log_open("install (non-interactive)")
    elif args.update:
        log_open("update")
    elif args.uninstall:
        log_open("uninstall")
    elif args.dry_run:
        log_open("dry-run")

    if args.cert_check:
        banner("certificate diagnostics")
        cert_diagnostics()
        if health(PORT):
            pause_after("cert-check")
        return

    if args.transport and not (args.install or args.update or args.uninstall):
        banner("transport %s" % args.transport)
        if not is_mac() and args.transport == "https":
            warn("https is a macOS feature - Windows stays plain http")
        else:
            _reapply_transport(args.transport, args.transport)
            pause_after("transport")
        return
    if args.start:
        start_only()
        return
    if args.restart:
        stop_server()
        start_only()
        return
    if args.stop:
        stop_server()
        ok("server stopped (port %d)" % PORT)
        return
    if args.status:
        raise SystemExit(do_status())
    if args.uninstall:
        banner("uninstall")
        do_uninstall(True if args.purge_config else None,
                     True if args.remove_certs else None)
        return
    if args.update:
        git_pull()
        install("Update", interactive=False)
        return
    if args.install:
        install("Install", interactive=False)
        return

    # ---------------- interactive menu ---------------- #
    log_open("menu")
    banner("menu")
    _rows = [
        ("1", "Install / repair    (Word + PowerPoint on; Outlook off by default)", None),
        ("2", "Update              (git pull + refresh; keeps your add-ins)", None),
        ("3", "Server              (start / restart / stop / status)", None),
        (None, "@SEP", None),
        ("4", "Transport options   (transport: %s | ports: http %d / https %d | certs: %s)" % (
            _current_transport(), _current_ports()[0], _current_ports()[1], _get_cert_mode()), None),
        (None, "@SEP", None),
        ("5", "Uninstall          (removes the add-ins + this server)",
         "red"),
        (None, "@SEP", None),
        ("0", "exit", "magenta"),
    ]

    def _ask_menu_typed():
        while True:
            print("")
            for (act, label, style) in _rows:
                if label == "@SEP":
                    print("")
                    continue
                print("  %s) %s" % (act, label))
            print("")
            try:
                if not sys.stdin.isatty():
                    print("  (non-interactive: no choice given - pass --install / --update / --uninstall)")
                    _log("menu: non-interactive, nothing chosen")
                    return None
                choice = input("  > choose [1/2/3/4/5, 0; default 1]: ").strip().lower()
            except EOFError:
                return None
            _log("menu input: %r" % choice, ts=True)
            if choice == "":
                return 0
            sel = next((i for i, (act, _l, _s) in enumerate(_rows) if act == choice), None)
            if sel is not None:
                return sel
            print("  please choose 1, 2, 3, 4 or 0")

    while True:
        print("")
        sel = menu_pick(_rows, 0, "(up/down + Enter = select, 1-5 = direct, q = exit)")
        if sel is None:
            sel = _ask_menu_typed()
            if sel is None:
                return
        # the transport/cert row must show the CURRENT stored values, not the
        # ones from when the menu was built (a switch inside the menu otherwise
        # leaves the stale wording behind)
        _ti = next(i for i, r in enumerate(_rows) if r[0] == "4")
        _rows[_ti] = (
            "4", "Transport options   (transport: %s | ports: http %d / https %d | certs: %s)" % (
                _current_transport(), _current_ports()[0], _current_ports()[1], _get_cert_mode()),
            None)
        act = _rows[sel][0]
        print("")
        _log("menu selection: %s" % act, ts=True)
        try:
            if act == "1":
                install("Install", interactive=True)
            elif act == "2":
                git_pull()
                install("Update", interactive=False)
            elif act == "3":
                server_menu()
            elif act == "4":
                transport_menu()
            elif act == "5":
                banner("uninstall")
                do_uninstall(None)
            else:
                return   # act == "0"
        except MenuAbort:
            pass          # sub-flow aborted - show the menu again
        banner("menu")    # flows loop back to the menu instead of quitting
        continue


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("")
        print("  aborted (Ctrl-C)")
        raise SystemExit(130)