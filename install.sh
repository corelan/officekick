#!/usr/bin/env bash
#
# SlideKick installer for macOS.
#
#   ./install.sh                 install or update (same script, idempotent)
#   ./install.sh --port 8765     use a specific port
#   ./install.sh --dev           serve straight from this repo (no file copy)
#   ./install.sh --no-launchagent  skip the auto-start LaunchAgent
#   ./install.sh run             run the server in the foreground (no install)
#   ./install.sh status          show install status
#   ./install.sh uninstall [--purge]   remove the add-in (--purge also wipes config)
#
# What "install" does:
#   * checks python3 (>= 3.8) and locates Microsoft PowerPoint
#   * copies src/ into ~/Library/Application Support/SlideKick (unless --dev)
#   * renders the add-in manifest with the chosen port and sideloads it into
#     ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
#   * installs a LaunchAgent so the local server starts at login (KeepAlive)
#   * starts the server and health-checks it
#
set -euo pipefail

APP_NAME="SlideKick"
LAUNCH_LABEL="com.corelan.slidekick"
REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" && pwd)"
SRC_DIR="$REPO_DIR/src"
INSTALL_DIR="$HOME/Library/Application Support/$APP_NAME"
CONFIG_DIR="$HOME/.slidekick"
LOG_DIR="$HOME/Library/Logs/$APP_NAME"
PLIST="$HOME/Library/LaunchAgents/$LAUNCH_LABEL.plist"
MANIFEST_NAME="slidekick.xml"
GUID="c7b1f2a9-86d3-4e58-9a4b-2f6d5e0c3b81"
DEFAULT_PORT=8765
VERSION="$(cat "$REPO_DIR/VERSION" 2>/dev/null || echo 0.1.2)"
CERT_DIR="$CONFIG_DIR/certs"

PORT=""
DEV=0
NO_AGENT=0
PURGE=0
HTTPS=0   # default: plain http://localhost, no keychain changes, no ribbon commands
SUB="install"

say()  { printf '%s\n' "$*"; }
ok()   { printf '  [ok]   %s\n' "$*"; }
warn() { printf '  [warn] %s\n' "$*"; }
die()  { printf '  [fail] %s\n' "$*" >&2; exit 1; }

while [[ $# -gt 0 ]]; do
  case "$1" in
    --port) PORT="${2:-}"; shift 2 ;;
    --dev) DEV=1; shift ;;
    --no-launchagent) NO_AGENT=1; shift ;;
    --https) HTTPS=1; shift ;;
    --purge) PURGE=1; shift ;;
    install|update|uninstall|status|run|help) SUB="$1"; shift ;;
    -h|--help) SUB="help"; shift ;;
    *) die "unknown argument: $1 (see $0 --help)" ;;
  esac
done

usage() {
  say "SlideKick installer"
  say ""
  say "  install (default)       install or update the add-in + server"
  say "  --port N                port for the local server (default $DEFAULT_PORT)"
  say "  --dev                   serve files directly from this repo (developer mode)"
  say "  --no-launchagent        do not install the login auto-start"
  say "  --https                 serve https + add the ribbon button (trusts a local CA in the login keychain)"
  say "  status                  show current install status"
  say "  run                     run the server in the foreground"
  say "  uninstall [--purge]     remove everything (config too with --purge)"
  exit 0
}

[[ "$SUB" == "help" ]] && usage

# --------------------------------------------------------------------- #
# environment checks
# --------------------------------------------------------------------- #
PY="$(command -v python3 || true)"
if [[ -z "$PY" ]]; then
  for cand in /usr/bin/python3 /usr/local/bin/python3 /opt/homebrew/bin/python3; do
    [[ -x "$cand" ]] && PY="$cand" && break
  done
fi
[[ -z "$PY" ]] && die "python3 not found - install Apple Command Line Tools or Homebrew python"
"$PY" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 8) else 1)' \
  || die "python3 is too old (need >= 3.8): $($PY --version 2>&1)"

find_powerpoint() {
  local p
  for p in "/Applications/Microsoft PowerPoint.app" "$HOME/Applications/Microsoft PowerPoint.app"; do
    if [[ -d "$p" ]]; then
      PP_APP="$p"
      PP_VERSION="$("$PY" - "$p/Contents/Info.plist" <<'PYEOF' 2>/dev/null || echo ""
import plistlib, sys
try:
    with open(sys.argv[1], "rb") as fh:
        print(plistlib.load(fh).get("CFBundleShortVersionString", ""))
except Exception:
    print("")
PYEOF
)"
      return 0
    fi
  done
  return 1
}

resolve_containers() {
  CONTAINER_PATHS=()
  local base="$HOME/Library/Containers" d
  if [[ -d "$base" ]]; then
    for d in "$base""/com.microsoft.Powerpoint" "$base""/com.microsoft.powerpoint" \
             "$base""/com.microsoft.PowerPoint"; do
      [[ -d "$d" ]] && CONTAINER_PATHS+=("$d")
    done
  fi
  if [[ ${#CONTAINER_PATHS[@]} -eq 0 ]]; then
    CONTAINER_PATHS+=("$HOME/Library/Containers/com.microsoft.Powerpoint")
    warn "PowerPoint container not found; will create $(basename "${CONTAINER_PATHS[0]}")"
  fi
}

port_busy() { "$PY" - "$1" <<'PYEOF'
import socket, sys
s = socket.socket()
try:
    s.bind(("127.0.0.1", int(sys.argv[1])))
    print("free")
except OSError:
    print("busy")
finally:
    s.close()
PYEOF
}

health() { # health PORT -> prints server JSON on success; nonzero rc if unreachable
  if [[ "$HTTPS" == "1" && -f "$CERT_DIR/slidekick-ca.pem" ]]; then
    curl -fsS --max-time 2 --cacert "$CERT_DIR/slidekick-ca.pem" "https://127.0.0.1:$1/api/health"
  else
    curl -fsS --max-time 2 "http://127.0.0.1:$1/api/health"
  fi
}

pick_port() {
  local p
  # explicit choice
  if [[ -n "$PORT" ]]; then return 0; fi
  # if our own server already runs on the default port, keep it (update path)
  local h
  h="$(health "$DEFAULT_PORT" 2>/dev/null || true)"
  if [[ "$h" == *SlideKick* ]]; then PORT="$DEFAULT_PORT"; return 0; fi
  [[ "$(port_busy "$DEFAULT_PORT")" == "free" ]] && PORT="$DEFAULT_PORT" && return 0
  for p in $(seq $((DEFAULT_PORT + 1)) $((DEFAULT_PORT + 19))); do
    [[ "$(port_busy "$p")" == "free" ]] && PORT="$p" && return 0
  done
  die "no free port found in range $DEFAULT_PORT..$((DEFAULT_PORT + 19))"
}

# --------------------------------------------------------------------- #
# local TLS certificate (loopback CA + localhost leaf, user-keychain trust)
# --------------------------------------------------------------------- #
ensure_local_certs() {
  SCHEME="http"
  if [[ "$HTTPS" != "1" ]]; then
    ok "http://localhost mode (default): no certificates, no keychain changes"
    ok "ribbon button not included in this mode - add it with: ./install.sh --https"
    return 0
  fi
  if ! command -v openssl >/dev/null 2>&1; then
    warn "openssl not found - falling back to plain HTTP (ribbon button will not appear)"
    HTTPS=0
    return 0
  fi

  local CA_PEM="$CERT_DIR/slidekick-ca.pem" CA_KEY="$CERT_DIR/slidekick-ca.key"
  local LEAF_PEM="$CERT_DIR/localhost.pem" LEAF_KEY="$CERT_DIR/localhost.key"

  mkdir -p "$CERT_DIR"
  chmod 700 "$CERT_DIR" 2>/dev/null || true

  if [[ ! -f "$CA_PEM" || ! -f "$CA_KEY" || ! -f "$LEAF_PEM" || ! -f "$LEAF_KEY" ]]; then
    say "  ... generating local loopback TLS certificates"
    cat > "$CERT_DIR/ca.cnf" <<CAEOF
[req]
distinguished_name = dn
x509_extensions = v3_ca
prompt = no
[dn]
CN = SlideKick Local CA (localhost-only, machine-generated)
[v3_ca]
basicConstraints = critical, CA:TRUE
keyUsage = critical, keyCertSign, cRLSign
subjectKeyIdentifier = hash
CAEOF
    cat > "$CERT_DIR/leaf.cnf" <<LEAFEOF
[req]
distinguished_name = dn
req_extensions = v3_req
prompt = no
[dn]
CN = localhost
[v3_req]
basicConstraints = critical, CA:FALSE
keyUsage = critical, digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth
subjectAltName = DNS:localhost, IP:127.0.0.1, IP:::1
LEAFEOF
    openssl req -x509 -newkey rsa:2048 -nodes -sha256 -days 3650 \
      -keyout "$CA_KEY" -out "$CA_PEM" -config "$CERT_DIR/ca.cnf"
    openssl req -new -newkey rsa:2048 -nodes -sha256 \
      -keyout "$LEAF_KEY" -out "$CERT_DIR/localhost.csr" -config "$CERT_DIR/leaf.cnf"
    openssl x509 -req -sha256 -days 825 \
      -in "$CERT_DIR/localhost.csr" -CA "$CA_PEM" -CAkey "$CA_KEY" -CAcreateserial \
      -out "$LEAF_PEM" -extfile "$CERT_DIR/leaf.cnf" -extensions v3_req
    rm -f "$CERT_DIR/localhost.csr"
    chmod 600 "$CA_KEY" "$LEAF_KEY"
    ok "certificates generated in $CERT_DIR"
  fi

  # trust the CA for this user (login keychain; no admin needed)
  local LOGIN_KC="$HOME/Library/Keychains/login.keychain-db"
  if ! security find-certificate -c "SlideKick Local CA" "$LOGIN_KC" >/dev/null 2>&1; then
    if security add-trusted-cert -r trustRoot -k "$LOGIN_KC" "$CA_PEM" 2>/dev/null; then
      ok "local CA trusted (login keychain) - the pane and ribbon can use https://localhost"
    else
      warn "could not auto-trust the local CA. Manual step: open Keychain Access,"
      warn "  import $CA_PEM and set 'Always Trust' for SSL - then reopen the pane"
    fi
  else
    ok "local CA already trusted"
  fi
  SCHEME="https"
}


# --------------------------------------------------------------------- #
# Office add-in cache maintenance
# --------------------------------------------------------------------- #
clear_office_caches() {
  # Only safe with PowerPoint closed: the add-in commands registration cache
  # (and page caches) refresh while the app is not running.
  if pgrep -f "Microsoft PowerPoint.app" >/dev/null 2>&1; then
    warn "PowerPoint is running - NOT clearing the Office add-in caches."
    warn "quit PowerPoint, re-run ./install.sh (it clears them), then reopen"
    return 0
  fi
  local container cdir
  for container in "${CONTAINER_PATHS[@]}"; do
    cdir="$container/Data/Library/Caches"
    if [[ -d "$cdir" ]]; then
      rm -rf "$cdir/Microsoft" "$cdir/com.microsoft.Powerpoint" \
             "$cdir/com.microsoft.ctrlstrcaches" "$cdir/WebKit" 2>/dev/null || true
      ok "cleared Office add-in caches: $cdir"
    fi
    # the WEF registration + resource store (manifest copies, downloaded icons):
    # rebuilt from the source manifest at launch; removing it forces a clean
    # re-download so the Add-ins menu picks up current icons
    local wefstore="$container/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
    if [[ -d "$wefstore" ]]; then
      rm -rf "$wefstore" 2>/dev/null || true
      ok "cleared WEF registration store: $wefstore"
    fi
  done
}

# --------------------------------------------------------------------- #
# install / update
# --------------------------------------------------------------------- #
do_install() {
  say "SlideKick $VERSION - install/update"
  PP_APP=""; PP_VERSION=""
  if find_powerpoint; then
    ok "PowerPoint found: $PP_APP (version ${PP_VERSION:-unknown})"
  else
    warn "Microsoft PowerPoint not found in /Applications or ~/Applications - installing anyway"
  fi

  resolve_containers
  # stop a previously installed server first so the canonical port frees up
  say "  ... stopping previously installed local server (if any)"
  launchctl unload "$PLIST" 2>/dev/null || true
  pick_port
  ok "port: $PORT"

  if [[ "$DEV" == "1" ]]; then
    SERVER_PY="$SRC_DIR/slidekick_server.py"
    STATIC_DIR="$SRC_DIR/static"
    ok "dev mode: serving from $SRC_DIR (no copy)"
  else
    mkdir -p "$INSTALL_DIR" "$CONFIG_DIR" "$LOG_DIR"
    say "  ... copying files to $INSTALL_DIR"
    rm -rf "$INSTALL_DIR"
    mkdir -p "$INSTALL_DIR" "$LOG_DIR"
    if command -v rsync >/dev/null 2>&1; then
      rsync -a --exclude '.DS_Store' --exclude '__pycache__' "$SRC_DIR/" "$INSTALL_DIR/"
    else
      cp -R "$SRC_DIR/." "$INSTALL_DIR/"
      find "$INSTALL_DIR" -name '__pycache__' -type d -exec rm -rf {} + 2>/dev/null || true
    fi
    SERVER_PY="$INSTALL_DIR/slidekick_server.py"
    STATIC_DIR="$INSTALL_DIR/static"
    ok "files deployed"
  fi

  "$PY" -m py_compile "$SERVER_PY" || die "server source failed to compile"
  ensure_local_certs

  # ---- manifest ----
  mkdir -p "$INSTALL_DIR/manifest"
  local mtmp
  mtmp="$INSTALL_DIR/manifest/$MANIFEST_NAME"
  "${SED:-sed}" \
    -e "s/__SCHEME__/$SCHEME/" \
    -e "s/__PORT__/$PORT/g" \
    "$SRC_DIR/manifest.template.xml" > "$mtmp"
  if [[ "$HTTPS" != "1" ]]; then
    # ribbon commands require https - drop them in the default http mode so the
    # add-in stays visible in Home -> Add-ins (commands mode removes the entry)
    "${SED:-sed}" -i '' -e '/__COMMANDS_BEGIN__/,/__COMMANDS_END__/d' "$mtmp"
  else
    "${SED:-sed}" -i '' -e '/__COMMANDS_BEGIN__/d' -e '/__COMMANDS_END__/d' "$mtmp"
  fi
  if command -v xmllint >/dev/null 2>&1; then
    xmllint --noout "$mtmp" || die "generated manifest is not valid XML"
    ok "manifest valid XML"
  else
    warn "xmllint not available - skipping XML validation"
  fi

  local container wef installed_any=0
  for container in "${CONTAINER_PATHS[@]}"; do
    wef="$container/Data/Documents/wef"
    mkdir -p "$wef"
    # remove stale copies of this add-in (matching the GUID) before updating
    local f
    for f in "$wef"/*.xml; do
      [[ -f "$f" ]] || continue
      if grep -q "$GUID" "$f" 2>/dev/null; then rm -f "$f"; fi
    done
    cp "$mtmp" "$wef/$MANIFEST_NAME"
    ok "manifest sideloaded: $wef/$MANIFEST_NAME"
    installed_any=1
  done
  [[ "$installed_any" == "1" ]] || die "could not sideload manifest into any PowerPoint container"

  clear_office_caches

  # ---- LaunchAgent ----
  if [[ "$NO_AGENT" == "1" ]]; then
    warn "skipping LaunchAgent (--no-launchagent)"
    if [[ "$(port_busy "$PORT")" == "busy" ]]; then
      :
    else
      warn "server not running - start it manually:  \"$PY\" \"$SERVER_PY\" --port $PORT"
    fi
  else
    say "  ... writing LaunchAgent $PLIST"
  "$PY" - "$PLIST" "$LAUNCH_LABEL" "$PY" "$SERVER_PY" "$PORT" "$STATIC_DIR" "$CERT_DIR" "$LOG_DIR" "$SCHEME" <<'PYPLIST'
import os, plistlib, sys

plist_path, label, py_bin, server_py, port, static_dir, cert_dir, log_dir, scheme = sys.argv[1:10]
args = [py_bin, server_py, "--port", port, "--static-dir", static_dir]
if scheme == "https":
    args += ["--cert", os.path.join(cert_dir, "localhost.pem"),
             "--key", os.path.join(cert_dir, "localhost.key")]
with open(plist_path, "wb") as fh:
    plistlib.dump({
        "Label": label,
        "ProgramArguments": args,
        "RunAtLoad": True,
        "KeepAlive": True,
        "StandardOutPath": os.path.join(log_dir, "server.log"),
        "StandardErrorPath": os.path.join(log_dir, "server.log"),
    }, fh)
PYPLIST
    launchctl unload "$PLIST" 2>/dev/null || true
    launchctl load -w "$PLIST"
  say "  ... restarting local server via LaunchAgent (RunAtLoad + KeepAlive)"
    ok "LaunchAgent installed (starts at login, restarts on crash)"
  fi

  # ---- health check ----
  local tries=0 h=""
  while [[ $tries -lt 20 ]]; do
    h="$(health "$PORT" 2>/dev/null || true)"
    if [[ "$h" == *'"ok": true'* ]]; then break; fi
    sleep 0.5
    tries=$((tries + 1))
  done
  if [[ "$h" == *'"ok": true'* ]]; then
    ok "server healthy: $SCHEME://127.0.0.1:$PORT - config: $CONFIG_DIR/config.json"
  else
    die "server did not become healthy on port $PORT - check $LOG_DIR/server.log"
  fi

  say ""
  say "Installed:"
  say "  add-in manifest : ${CONTAINER_PATHS[0]}/Data/Documents/wef/$MANIFEST_NAME"
  say "  app files       : $INSTALL_DIR"
  say "  config (keys)   : $CONFIG_DIR/config.json   (chmod 600)"
  say "  server log      : $LOG_DIR/server.log"
  say "  simulator       : http://127.0.0.1:$PORT/simulator.html   (test without PowerPoint)"
  say ""
  say "First use in PowerPoint:"
  if [[ "$HTTPS" == "1" ]]; then
    say "  1. quit and reopen PowerPoint (required when it was already running)"
    say "  2. open/create a deck"
    say "  3. the Corelan SlideKick button is in the Home tab (AI actions group)"
    say "     (fallback: Home -> Add-ins -> Developer Add-ins -> Corelan SlideKick)"
    say "  4. Connections tab: add a provider; Profiles tab: pick your models"
  else
    say "  1. quit and reopen PowerPoint (required when it was already running)"
    say "  2. open/create a deck"
    say "  3. Home tab -> Add-ins (dropdown) -> Developer Add-ins -> Corelan SlideKick"
    say "  4. Connections tab: add a provider; Profiles tab: pick your models"
  fi
  [[ -n "$PP_VERSION" ]] && [[ "$PP_APP" && "$DEV" == "0" ]] && true
}

# --------------------------------------------------------------------- #
# uninstall
# --------------------------------------------------------------------- #
do_uninstall() {
  say "SlideKick - uninstall"
  launchctl unload "$PLIST" 2>/dev/null && ok "LaunchAgent unloaded" || warn "no LaunchAgent"
  [[ -f "$PLIST" ]] && rm -f "$PLIST" && ok "LaunchAgent removed"
  resolve_containers
  local container wef f removed=0
  for container in "${CONTAINER_PATHS[@]}"; do
    wef="$container/Data/Documents/wef"
    if [[ -d "$wef" ]]; then
      for f in "$wef"/*.xml; do
        [[ -f "$f" ]] || continue
        if grep -q "$GUID" "$f" 2>/dev/null; then rm -f "$f"; removed=1; fi
      done
    fi
  done
  [[ "$removed" == "1" ]] && ok "manifest removed from PowerPoint container(s)"
  if pgrep -f "slidekick_server.py" >/dev/null 2>&1; then
    pkill -f "slidekick_server.py" 2>/dev/null && ok "server stopped" || warn "could not stop server process"
  fi
  if [[ -d "$INSTALL_DIR" && "$DEV" == "0" ]]; then
    rm -rf "$INSTALL_DIR" && ok "$INSTALL_DIR removed"
  fi
  if [[ "$PURGE" == "1" ]]; then
    rm -rf "$CONFIG_DIR" && ok "config wiped ($CONFIG_DIR)"
  else
    ok "config kept at $CONFIG_DIR (use --purge to wipe)"
  fi
  warn "quit and reopen PowerPoint to finish removing the add-in"
}

# --------------------------------------------------------------------- #
# status
# --------------------------------------------------------------------- #
do_status() {
  resolve_containers
  local h
  h="$(health "$PORT" 2>/dev/null || health "$DEFAULT_PORT" 2>/dev/null || true)"
  if [[ "$h" == *SlideKick* ]]; then
    ok "server is running: $h"
  else
    warn "server NOT running (default port $DEFAULT_PORT)"
  fi
  if launchctl print "gui/$(id -u)/$LAUNCH_LABEL" >/dev/null 2>&1; then
    ok "LaunchAgent present"
  else
    warn "LaunchAgent absent"
  fi
  local container wef f found=""
  for container in "${CONTAINER_PATHS[@]}"; do
    wef="$container/Data/Documents/wef/$MANIFEST_NAME"
    [[ -f "$wef" ]] && found="$wef"
  done
  [[ -n "$found" ]] && ok "manifest sideloaded: $found" || warn "manifest not sideloaded"
  [[ -f "$CONFIG_DIR/config.json" ]] && ok "config present: $CONFIG_DIR/config.json" || warn "no config yet"
  [[ -d "$INSTALL_DIR" ]] && ok "app dir present: $INSTALL_DIR" || warn "app dir missing (or dev mode)"
}

# --------------------------------------------------------------------- #
# run in foreground
# --------------------------------------------------------------------- #
run_foreground() {
  if [[ "$DEV" == "1" ]]; then
    SERVER_PY="$SRC_DIR/slidekick_server.py"
    STATIC_DIR="$SRC_DIR/static"
  else
    SERVER_PY="$INSTALL_DIR/slidekick_server.py"
    STATIC_DIR="$INSTALL_DIR/static"
    [[ -f "$SERVER_PY" ]] || die "not installed yet (run $0 install first)"
  fi
  pick_port
  exec "$PY" "$SERVER_PY" --port "$PORT" --static-dir "$STATIC_DIR"
}

case "$SUB" in
  install|update) do_install ;;
  uninstall) do_uninstall ;;
  status) do_status ;;
  run) run_foreground ;;
  *) usage ;;
esac