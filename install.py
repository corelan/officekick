#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OfficeKick unified installer - SlideKick (PowerPoint) + DocKick (Word) +
MailKick (Outlook). Works on macOS and Windows; auto-detects the platform.

    python3 install.py                menu: Install / Update / Server / Uninstall
    python3 install.py --install      non-interactive install (Word + PowerPoint;
                                      Outlook only with --mailkick / --mailkick-cert-key
                                      or when Outlook was previously installed)
    python3 install.py --update       git pull (if this is a checkout) + update
    python3 install.py --uninstall [--purge-config]
    python3 install.py --start | --stop | --restart | --status
    python3 install.py --dry-run      with any mode: print the actions, run none
    python3 install.py --port N       server port (default 8765)
    python3 install.py --ribbon       opt-in: render the AI ribbon commands
                                      (experimental on this Office build)
    python3 install.py --mailkick     non-interactive opt-in: install Outlook and
                                      create the self-signed loopback CA + certificate

Installer output is structured: short step lines on the console, the full
detail trail (file lists, tool output, decisions) in install.log next to
install.py. Word + PowerPoint are enabled by default; Outlook is off by
default - enabling Outlook needs the https prerequisites and will create a
self-signed loopback CA + localhost certificate (clearly warned about in the
menu). On Windows no certificate is needed (Outlook rides the plain-http
localhost listener there).

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


def ok(msg):
    _con = ("  [ok]   " if is_windows() else "   \x1b[32mv\x1b[0m ") + msg
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
    print(("  [warn] " if is_windows() else "   \x1b[33m!\x1b[0m  ") + msg)
    _log(msg, ts=True)


def step(msg):
    _con = ("  [..]   " if is_windows() else "   \x1b[36m~\x1b[0m  ") + msg
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
            reasons.append("openssl not found - it generates the loopback CA +"
                           " localhost certificate")
        else:
            _rc, over = run_quiet(["openssl", "version"])
            techniques.append("%s (%s) - generates the self-signed loopback CA"
                              " + localhost certificate" % (
                                  os.path.basename(ov), (over or "").strip() or "openssl"))
        if shutil.which("security") is None:
            reasons.append("the macOS 'security' tool is missing - it imports "
                           "+ trusts the CA in your keychains")
        else:
            techniques.append("the macOS 'security' tool - imports the CA into"
                              " your login keychain and marks it trusted")
        try:
            import ssl  # noqa: F401
            techniques.append("python ssl - serves the https://localhost:%d"
                              " pane listener for Outlook" % (PORT + 1))
        except Exception as exc:
            reasons.append("python ssl module unavailable: %s" % exc)
        return (len(reasons) == 0, reasons, techniques)
    return (True, [], [])   # Windows: no TLS component needed for Outlook


def _current_transport():
    """The persisted transport mode (auto/http/https): from the running
    server first, then from the config file; auto when unknown."""
    try:
        import urllib.request
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open("http://127.0.0.1:%d/api/config" % PORT, timeout=1.5) as resp:
            mode = (json.loads(resp.read().decode("utf-8", "replace")) or {}).get("transportMode")
        if mode in ("auto", "http", "https"):
            return mode
    except Exception:
        pass
    try:
        with open(config_file(), "r", encoding="utf-8") as fh:
            mode = (json.load(fh) or {}).get("transportMode")
        if mode in ("auto", "http", "https"):
            return mode
    except Exception:
        pass
    return "auto"


def _set_transport(mode):
    """Persist the transport mode: through the running server (settings PUT)
    so both sides agree, or straight into the config file when it is down."""
    if mode not in ("auto", "http", "https"):
        return False
    if health(PORT):
        try:
            import urllib.request
            req = urllib.request.Request(
                "http://127.0.0.1:%d/api/settings" % PORT,
                data=json.dumps({"transportMode": mode}).encode("utf-8"),
                headers={"Content-Type": "application/json"}, method="PUT")
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with opener.open(req, timeout=3) as resp:
                resp.read()
            note("transportMode set to %s (via the running server)" % mode)
            return True
        except Exception as exc:
            warn("could not store transportMode via the server (%s) - writing the config file" % exc)
    try:
        cfg = {}
        try:
            with open(config_file(), "r", encoding="utf-8") as fh:
                cfg = json.load(fh) or {}
        except Exception:
            cfg = {}
        cfg["transportMode"] = mode
        with open(config_file(), "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, indent=2)
        try:
            os.chmod(config_file(), 0o600)
        except OSError:
            pass
        note("transportMode set to %s (config file written)" % mode)
        return True
    except Exception as exc:
        warn("could not write %s: %s" % (config_file(), exc))
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
    if h:
        srv = "running on 127.0.0.1:%d (version %s)" % (PORT, h.get("version"))
    else:
        srv = "not running (start it here or: install.py --start)"
    print("  detected : %s %s | python %s" % (
        platform.system(), platform.release(), platform.python_version()))
    print("  install  : %s" % install_dir())
    print("  server   : %s" % srv)
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
        if not is_mac():
            out_lines.append(_paint("  https prereqs for Outlook: ", "bold") + "not needed -")
            out_lines.append("  Outlook rides the plain-http localhost listener on")
            out_lines.append("  Windows; no certificate is created or trusted.")
            return out_lines
        if feas:
            out_lines.append("  " + _paint("https prerequisites for Outlook: MET", "green"))
            for t in feas_how:
                out_lines.append("    - " + t)
        else:
            out_lines.append("  " + _paint("https prerequisites for Outlook: NOT MET", "red")
                             + _paint(" - Outlook is locked", "bold"))
            for rln in feas_why:
                out_lines.append("    - " + rln)
        out_lines.append(_paint("  Enabling Outlook creates a self-signed loopback CA +"
                                " localhost certificate (asked to confirm first).", "dim"))
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
                   "  Add-ins (Word + PowerPoint are enabled by default, Outlook is not):",
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
                go = ask_yes("Outlook needs a SELF-SIGNED loopback CA + localhost"
                             " certificate (created + imported into your keychains"
                             " during install) - continue", False)
                note("outlook CA warning answer: %s" % ("yes" if go else "no"))
                if not go:
                    state[folder] = False
                    msg = ["   - Outlook certificate setup declined - MailKick stays off."]
                    continue
                msg = ["   - Outlook enabled - the self-signed loopback CA + certificate"
                       " will be created during the certificate step."]
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
    return "https://localhost:%d" % (PORT + 1) if tls_active() else "http://127.0.0.1:%d" % PORT


def pane_listener_line():
    """One line describing where the panes are served, honest about TLS."""
    if tls_active():
        return (
            "  https://localhost:%d  <- task panes + AI requests (TLS, user-trusted CA)\n"
            "  http://127.0.0.1:%d   <- same statics, plain http" % (PORT + 1, PORT)
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
    """Which TLS route for Outlook (macOS only; called only when MailKick is
    selected). 'auto' = generate a self-signed loopback CA + localhost
    certificate (the user was already warned about this in the add-in menu),
    'custom' = the user's own already-trusted PEM pair, 'skip' = none.
    Called BEFORE the certificates phase line so the questions precede the
    action."""
    good, _how, techniques = https_feasible()
    del good
    global MAILKICK_CERT, MAILKICK_KEY
    if MAILKICK_CERT or MAILKICK_KEY:
        good2, problems = validate_user_cert(MAILKICK_CERT, MAILKICK_KEY)
        if good2:
            return "custom"
        for pr in problems:
            warn("provided MailKick certificate: " + pr)
        return "skip"
    if DRY_RUN or (not sys.stdin.isatty()):
        return "auto"    # automation opt-in (--mailkick): reuse-first applies
    print("")
    print("  " + _paint("Outlook pane transport", "bold") + " - the self-signed")
    print("  loopback CA + certificate (%s) or your own trusted pair." %
          (techniques[0].split(" - ")[0] if techniques else "openssl"))
    MAILKICK_CHOICE_TMP = ask_yes("Use the self-signed loopback CA (recommended)", True)
    if MAILKICK_CHOICE_TMP:
        return "auto"
    for attempt in range(3):
        crt = input("  " + _paint("? path to your TLS certificate (PEM): ", "cyan")).strip()
        key = input("  " + _paint("? path to the matching private key (PEM, unencrypted): ", "cyan")).strip()
        good2, problems = validate_user_cert(crt, key)
        if good2:
            MAILKICK_CERT, MAILKICK_KEY = crt, key
            return "custom"
        for pr in problems:
            warn(pr)
        print("       (attempt %d/3)" % (attempt + 1))
    print("  giving up - Outlook will not be registered.")
    return "skip"


# ------------------------------------------------------------ loopback TLS #
def validate_existing_tls(crt):
    """Check the cert ALREADY on disk: unexpired, localhost SAN, trusted.
    None = usable; a short reason string = regenerate."""
    rc, _o = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-checkend", "0"])
    if rc != 0:
        return "expired or not yet valid"
    rc, o = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-text"])
    if "DNS:localhost" not in o and "IP Address:127.0.0.1" not in o:
        return "no localhost subjectAltName"
    rc, o = run_quiet(["security", "verify-cert", "-c", crt, "-p", "ssl"])
    if rc != 0:
        return "its chain is not trusted by this Mac (" + (o or "untrusted").strip()[:120] + ")"
    return None


CERT_FACTS = {}            # loopback TLS info, printed after the certificates phase


def _cert_facts(crt):
    """Name + expiry of one PEM certificate ({} = unreadable)."""
    rc, o = run_quiet(["openssl", "x509", "-in", crt, "-noout", "-subject", "-enddate"])
    if rc != 0:
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
            try:
                os.chmod(key, 0o600)
            except OSError:
                pass
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
    try:
        os.chmod(key, 0o600)
    except OSError:
        pass
    keychain = os.path.expanduser("~/Library/Keychains/login.keychain-db")
    rc, o = run_quiet(["security", "add-trusted-cert", "-r", "trustRoot",
                       "-k", keychain, ca_crt])
    if rc == 0:
        ok("loopback CA trusted in your login keychain (imported: %s)" % ca_crt)
    else:
        warn("could not auto-trust the loopback CA - run once yourself, then restart the server:")
        print("       security add-trusted-cert -r trustRoot -k $HOME/Library/Keychains/login.keychain-db " + ca_crt)
    _record("generated (fresh CA + leaf)")
    return (crt, key)


# ------------------------------------------------------------------ deploy #
def deploy_files():
    dst = install_dir()
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
            os.chmod(cfgjson, 0o600)
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
        out("  [..]  [dry-run] would record repoPath=%s" % REPO_DIR)
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
            os.chmod(cfgjson, 0o600)
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


def render_manifest(addin_dir, port, https_port=None):
    """https_port: when a working loopback certificate exists, ALL add-in
    panes are served over https://localhost:<port+1> (same TLS listener as
    MailKick) and their manifests point there; None = plain http."""
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
        if not tls_pair and is_mac():
            warn("transport mode https - the certificate check did not pass;"
                 " creating/renewing the loopback pair now")
            pair = ensure_tls_certs()
            reason2 = None
            if pair:
                reason2 = validate_existing_tls(pair[0])
            if pair and reason2 is None:
                tls_pair = pair
            else:
                warn("https mode but no usable certificate - the manifests stay plain http"
                     " (the https listener will start once a pair validates)")
        elif not tls_pair and not is_mac():
            warn("transport mode https is a macOS feature - Windows stays plain http")
    step("rendering add-in manifests (%s, %s)" %
         (("https://localhost:%d via the loopback certificate" % (PORT + 1)) if tls_pair
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
        xml = render_manifest(folder, PORT, https_port=(PORT + 1) if tls_pair else None)
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
    if not is_mac():
        sideload_windows_catalog()
        sideload_windows_developer()
        out("  note: Outlook on Windows does not use folder catalogs - load MailKick once:")
        out("        Outlook -> Get Add-ins (or the ellipsis overflow) -> My add-ins ->")
        out("        Custom Add-ins -> Add a custom add-in -> Add from File ->")
        out("        " + os.path.join(manifest_dir(), "mailkick.xml"))
    else:
        out("  mac tip: if a saved deck throws the 'Add-in Error' toast on open,")
        out("           the deck's pane reference carries the unresolvable store stamp.")
        out("           Heal: python3 tools/repair_paneref.py <file.pptx | .docx | folder>")
        out("           (backs the originals up to orig/ next to the deck(s), then patches in place)")


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
    note("caches cleared; install flows re-stage the manifests to the WEF dirs"
         " immediately after (render_and_sideload), so no manual re-add step is"
         " needed; MailKick rides the https://localhost:%d loopback listener and"
         " works in New Outlook too." % (PORT + 1))


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
        print("")
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
    pause_after("uninstall")


# ---------------------------------------------------------------- install #
def preflight(mode):
    """Console: nothing (the banner carries status). The transport explanation
    and the popup essays live in install.log only."""
    note("preflight (%s); transport: panes -> http://localhost:%d (Office allows"
         " localhost over http%s).%s" % (
             mode, PORT,
             "; plus the https://localhost:%d loopback TLS listener wherever a"
             " working certificate exists" % (PORT + 1),
             " Outlook on macOS validates every manifest URL as https, so"
             " MailKick must ride the TLS component." if is_mac() else ""))
    if is_mac():
        note("popups you may see: (1) one-time background-item note when the"
             " autostart login item is first registered - that IS the local"
             " server app; (2) a Local Network prompt the first time a model"
             " on your LAN is used - click Allow so panes can reach e.g."
             " Ollama on another machine; (3) the CA-trust confirm when the"
             " loopback CA is imported.")
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
        MAILKICK_CHOICE = pick_mailkick_route()
    phase(3, total, "certificates", tls_phase)
    _print_cert_facts()
    phase(4, total, "registering with Office", registration_phase)
    if INTERACTIVE and not DRY_RUN:
        AUTOSTART_WANT = ask_autostart()
    phase(5, total, "autostart", autostart_phase)
    phase(6, total, "starting the server", server_phase)
    done_note(mode_name)
    pause_after(mode_name)


def pause_after(label):
    """Hold the screen after a flow so the result can be read on a real
    terminal (piped/automation runs and scheduled tasks just continue)."""
    _log("flow ended: %s" % label)
    try:
        if not (sys.stdin.isatty() and sys.stdout.isatty()):
            return
        try:
            input("  " + _paint("? Press Enter to continue (results above; details"
                                " in install.log - Ctrl-C quits)", "dim"))
            _log("pause released (Enter)")
        except EOFError:
            _log("pause released (EOF)")
        except KeyboardInterrupt:
            _log("pause released (Ctrl-C -> quit)")
            print("")
            raise SystemExit(130)
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
    if not is_mac():
        if "mailkick" in ADDINS:
            return "not needed on Windows (plain-http localhost)"
        return "no TLS on Windows (plain http)"
    crt = os.path.join(tls_cert_dir(), "server.crt")
    key = os.path.join(tls_cert_dir(), "server.key")
    have_pair = os.path.isfile(crt) and os.path.isfile(key)
    if "mailkick" in ADDINS:
        if MAILKICK_CHOICE not in ("auto", "custom", "skip"):
            MAILKICK_CHOICE = pick_mailkick_route()   # non-interactive fallback
        if MAILKICK_CHOICE == "custom":
            cdir = tls_cert_dir()
            os.makedirs(cdir, exist_ok=True)
            shutil.copyfile(MAILKICK_CERT, os.path.join(cdir, "server.crt"))
            shutil.copyfile(MAILKICK_KEY, os.path.join(cdir, "server.key"))
            os.chmod(os.path.join(cdir, "server.key"), 0o600)
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
    return render_and_sideload()


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
def _reapply_transport(mode, label):
    """Store the mode, re-render the manifests accordingly, restart the
    server so the listeners match, and report."""
    ok("transport mode -> %s" % mode)
    _set_transport(mode)
    if mode == "https" and is_mac():
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


def transport_menu():
    cur = _current_transport()
    rows = [
        ("1", "auto    (https when a working certificate exists - the classic default)", None),
        ("2", "plain http only (certificates ignored; Outlook not usable with this)", None),
        ("3", "https always (creates/renews the loopback CA + certificate for ALL panes)", None),
        (None, "@SEP", None),
        ("0", "back (mode stays: %s)" % cur, "magenta"),
    ]

    def body():
        clear_screen()
        h = health(PORT)
        print("")
        print("  ----- transport -----")
        print("  current mode: %s" % cur)
        if h:
            print("  server: running v%s on port %d" % (h.get("version"), PORT))
        return rows

    while True:
        r = body()
        sel = menu_pick(rows, 0, "(up/down + Enter, 1-3 = set, 0 = back)")
        if sel is None:
            print("")
            for (act, label, _st) in rows:
                if label == "@SEP":
                    print("")
                    continue
                print("    %s" % label if not act else "    %s) %s" % (act, label))
            print("")
            try:
                if not sys.stdin.isatty():
                    print("  (non-interactive: pass --transport auto|http|https)")
                    return
                choice = input("  > choose [1/2/3, 0; default 0]: ").strip().lower()
            except EOFError:
                return
            if choice in ("", "0", "q", "back"):
                return
            order = {"1": "auto", "2": "http", "3": "https"}
            if choice not in order:
                print("  please choose 1, 2, 3 or 0")
                continue
            _reapply_transport(order[choice], choice)
            cur = _current_transport()
            continue
        act = rows[sel][0]
        if act == "0" or not act:
            return
        mode = {"1": "auto", "2": "http", "3": "https"}[act]
        _reapply_transport(mode, act)
        cur = _current_transport()


def do_status():
    h = health(PORT)
    if h:
        print("  running  : http://127.0.0.1:%d (version %s)" % (PORT, h.get("version")))
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
        print("  " + ("running v%s on port %d" % (h.get("version"), PORT) if h
                         else "not running"))
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
            selrows = {"1": 0, "2": 1, "3": 2, "4": 3}
            sel = selrows.get(choice)
            if sel is None:
                print("  please choose 1, 2, 3, 4 or 0")
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
                    help="with --install: also install Outlook (creates the self-signed loopback CA + localhost certificate)")
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
    args = ap.parse_args()
    global SKIP_MAILKICK, MAILKICK_CERT, MAILKICK_KEY, MAILKICK_OPT
    SKIP_MAILKICK = args.skip_mailkick
    MAILKICK_CERT = args.mailkick_cert
    MAILKICK_KEY = args.mailkick_key
    MAILKICK_OPT = args.mailkick
    PORT = args.port if 1024 <= args.port <= 65535 else DEFAULT_PORT
    DRY_RUN = args.dry_run
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

    if args.install:
        log_open("install (non-interactive)")
    elif args.update:
        log_open("update")
    elif args.uninstall:
        log_open("uninstall")
    elif args.dry_run:
        log_open("dry-run")

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
        do_uninstall(True if args.purge_config else None)
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
        ("5", "Transport (%s)  (http / https / auto)" % _current_transport(), None),
        (None, "@SEP", None),
        ("4", "Uninstall", "red"),
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
                choice = input("  > choose [1/2/3/4, 0; default 1]: ").strip().lower()
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
        sel = menu_pick(_rows, 0, "(up/down + Enter = select, 1-4 = direct, q = exit)")
        if sel is None:
            sel = _ask_menu_typed()
            if sel is None:
                return
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
            elif act == "5":
                transport_menu()
            elif act == "4":
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