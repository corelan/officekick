#!/usr/bin/env bash
#
# OfficeKick installer for macOS (SlideKick for PowerPoint + DocKick for Word).
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

APP_NAME="OfficeKick"
LAUNCH_LABEL="com.corelan.officekick"
REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" && pwd)"
SRC_DIR="$REPO_DIR/slidekick"       # PowerPoint add-in (SlideKick)
WORDKICK_DIR="$REPO_DIR/wordkick"   # Word add-in (DocKick)
SERVER_REPO="$REPO_DIR/core/server.py"
INSTALL_DIR="$HOME/Library/Application Support/$APP_NAME"
CONFIG_DIR="$HOME/.officekick"
LOG_DIR="$HOME/Library/Logs/$APP_NAME"
PLIST="$HOME/Library/LaunchAgents/$LAUNCH_LABEL.plist"
MANIFEST_NAME="slidekick.xml"
WORD_MANIFEST_NAME="dockick.xml"
GUID="c7b1f2a9-86d3-4e58-9a4b-2f6d5e0c3b81"
WORD_GUID="ce1e8f09-db1c-4598-8f5b-6a65b7802d59"
DEFAULT_PORT=8765
VERSION="$(cat "$REPO_DIR/VERSION" 2>/dev/null || echo 0.1.3)"
CERT_DIR="$CONFIG_DIR/certs"

# --- legacy single-add-in layout (SlideKick-only era): migrated on install ---
OLD_LAUNCH_LABEL="com.corelan.slidekick"
OLD_INSTALL_DIR="$HOME/Library/Application Support/SlideKick"
OLD_PLIST="$HOME/Library/LaunchAgents/$OLD_LAUNCH_LABEL.plist"
OLD_CONFIG_DIR="$HOME/.slidekick"

PORT=""
DEV=0
NO_AGENT=0
PURGE=0
HTTPS=0   # default: plain http://localhost, no keychain changes, no ribbon commands
# scheme persistence: plain ./install.sh KEEPS the mode the running LaunchAgent
# uses, so a later default reinstall does not silently undo an earlier --https.
MODE_FORCED=0
for a in "$@"; do
  [[ "$a" == "--https" || "$a" == "--http" || "$a" == "--purge" || "$a" == "uninstall" ]] && MODE_FORCED=1 && break
done
detect_scheme_from_agent() {
  local plist="$HOME/Library/LaunchAgents/${LAUNCH_LABEL:-com.corelan.officekick}.plist"
  if [[ -f "$plist" ]]; then
    if grep -q -- "--cert" "$plist"; then
      HTTPS=1
    else
      HTTPS=0
    fi
    return 0
  fi
  return 1
}
if [[ "$MODE_FORCED" != "1" ]] && detect_scheme_from_agent; then
  echo "[scheme] keeping the running install's transport: $([[ $HTTPS == 1 ]] && echo https || echo http) (use --https/--http to switch)"
fi
SUB="install"

say()   { printf '%s\n' "$*"; }
ok()    { printf '   ✅ %s\n' "$*"; }
warn()  { printf '   ⚠️  %s\n' "$*"; }
die()   { printf '   ❌ %s\n' "$*" >&2; exit 1; }
step()  { printf '   🔄 %s\n' "$*"; }
title() { printf '%s\n' ""; printf '%s\n' "$*"; }

migrate_legacy() {
  # SlideKick-only era -> OfficeKick layout: adopt config + certs, remove the
  # old LaunchAgent and installed files (config is migrated, never deleted)
  if [[ -f "$OLD_CONFIG_DIR/config.json" ]] && grep -q "version" "$OLD_CONFIG_DIR/config.json" 2>/dev/null; then
    mkdir -p "$CONFIG_DIR"
    if [[ ! -f "$CONFIG_DIR/config.json" ]]; then
      cp "$OLD_CONFIG_DIR/config.json" "$CONFIG_DIR/config.json"
      chmod 600 "$CONFIG_DIR/config.json"
      ok "migrated config: $OLD_CONFIG_DIR/config.json -> $CONFIG_DIR/config.json"
    fi
    if [[ -d "$OLD_CONFIG_DIR/certs" ]]; then
      # only seed an EMPTY certs dir (never resurrect certificates the user deleted)
      if [[ ! -e "$CONFIG_DIR/certs/slidekick-ca.pem" && ! -e "$CONFIG_DIR/certs/localhost.pem" ]]; then
        mkdir -p "$CONFIG_DIR/certs"
        if cp -R "$OLD_CONFIG_DIR/certs/." "$CONFIG_DIR/certs/" 2>/dev/null; then
          ok "migrated certificates -> $CONFIG_DIR/certs"
        fi
      fi
    fi
  fi
  if [[ -f "$OLD_PLIST" ]]; then
    launchctl unload "$OLD_PLIST" 2>/dev/null || true
    rm -f "$OLD_PLIST"
    ok "legacy LaunchAgent removed ($OLD_LAUNCH_LABEL)"
  fi
  if [[ -d "$OLD_INSTALL_DIR" ]]; then
    rm -rf "$OLD_INSTALL_DIR"
    ok "legacy install dir removed: $OLD_INSTALL_DIR"
  fi
  pgrep -f "SlideKick/slidekick_server.py" >/dev/null 2>&1 && {
    pkill -f "SlideKick/slidekick_server.py" 2>/dev/null || true
    ok "legacy server process stopped"
  }
  return 0
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --port) PORT="${2:-}"; shift 2 ;;
    --dev) DEV=1; shift ;;
    --no-launchagent) NO_AGENT=1; shift ;;
    --https) HTTPS=1; shift ;;
    --http) HTTPS=0; shift ;;
    --purge) PURGE=1; shift ;;
    install|update|uninstall|status|run|help) SUB="$1"; shift ;;
    -h|--help) SUB="help"; shift ;;
    *) die "unknown argument: $1 (see $0 --help)" ;;
  esac
done

usage() {
  say "OfficeKick installer (SlideKick for PowerPoint + DocKick for Word)"
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

# union of all Office app containers we own add-ins for (PowerPoint + Word),
# with APFS case-variant dedupe by inode (all spellings = one real folder)
all_owned_containers() {
  ALL_OWNED=()
  local base="$HOME/Library/Containers" d g ino seen="|"
  local globs=(
    "com.microsoft.[Pp]owerpoint" "com.microsoft.powerpoint" "com.microsoft.PowerPoint"
    "com.microsoft.[Ww]ord" "com.microsoft.word" "com.microsoft.Word"
  )
  if [[ -d "$base" ]]; then
    for g in "${globs[@]}"; do
      for d in "$base"/$g; do
        [[ -d "$d" ]] || continue
        ino="$(stat -f '%i' "$d" 2>/dev/null)" || continue
        case "$seen" in
          *"|$ino|"*) ;;
          *) seen="$seen$ino|"; ALL_OWNED+=("$d") ;;
        esac
      done
    done
  fi
}

# containers for ONE app's glob list (PowerPoint-only era helper, kept generic)
resolve_containers_for() {
  local out_name="$1"; shift
  local base="$HOME/Library/Containers" g d ino seen="|"
  local result=()
  for g in "$@"; do
    for d in "$base"/$g; do
      [[ -d "$d" ]] || continue
      ino="$(stat -f '%i' "$d" 2>/dev/null)" || continue
      case "$seen" in
        *"|$ino|"*) ;;
        *) seen="$seen$ino|"; result+=("$d") ;;
      esac
    done
  done
  printf '%s\n' "${result[@]}"
}

# compat wrapper for older call sites (PowerPoint + Word union)
resolve_containers() {
  CONTAINER_PATHS=()
  all_owned_containers
  CONTAINER_PATHS=(${ALL_OWNED[@]+"${ALL_OWNED[@]}"})
  if [[ ${#CONTAINER_PATHS[@]} -eq 0 ]]; then
    CONTAINER_PATHS=("$HOME/Library/Containers/com.microsoft.Powerpoint")
    warn "Office containers not found; will create PowerPoint container on demand"
  fi
}

port_busy() {
  # authoritative check first: is anything LISTENING?
  if lsof -nP -iTCP:"$1" -sTCP:LISTEN 2>/dev/null | grep -q LISTEN; then
    echo busy
    return 0
  fi
  # no listener: lingering CLOSE_WAIT/FIN_WAIT sockets from the previous
  # server can still fail a plain bind - probe with SO_REUSEADDR, which is
  # exactly what the SlideKick server itself uses, so it matches reality
  "$PY" - "$1" <<'PYEOF'
import socket, sys
s = socket.socket()
s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
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
  # tolerate both schemes: a TLS server never answers plain http probes
  if [[ -f "$CERT_DIR/slidekick-ca.pem" ]]; then
    curl -fsS --max-time 2 --cacert "$CERT_DIR/slidekick-ca.pem" "https://127.0.0.1:$1/api/health" 2>/dev/null && return 0
    curl -fsSk --max-time 2 "https://127.0.0.1:$1/api/health" 2>/dev/null && return 0
  fi
  curl -fsS --max-time 2 "http://127.0.0.1:$1/api/health" 2>/dev/null
}

pick_port() {
  local p waited=0
  # explicit choice
  if [[ -n "$PORT" ]]; then return 0; fi
  # the previous server is SIGTERMed by launchctl unload above; wait for its
  # listener to actually disappear (a still-bound socket means the process is
  # exiting) - otherwise we would drift to PORT+1 while the manifest keeps PORT
  while [[ "$(port_busy "$DEFAULT_PORT")" != "free" && "$waited" -lt 25 ]]; do
    sleep 0.2; waited=$((waited + 1))
  done
  # if our own server already runs on the default port, keep it (update path)
  local h
  h="$(health "$DEFAULT_PORT" 2>/dev/null || true)"
  if [[ "$h" == *SlideKick* ]]; then PORT="$DEFAULT_PORT"; return 0; fi
  # stubborn listener (busy but not our healthy server): kill it instead of
  # drifting to PORT+1, which would desync the manifest URLs
  if [[ "$(port_busy "$DEFAULT_PORT")" != "free" ]]; then
    local stale
    stale="$(lsof -tnP -iTCP:"$DEFAULT_PORT" -sTCP:LISTEN 2>/dev/null | head -1)"
    if [[ -n "$stale" ]]; then
      warn "killing stale listener on port $DEFAULT_PORT (pid $stale)"
      kill "$stale" 2>/dev/null || true
      sleep 1
    fi
  fi
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
    ok "transport: plain http://localhost (default: no certificates, no keychain changes)"
    say "        (no ribbon button in this mode - opt in with: ./install.sh --https 🔐)"
    return 0
  fi
  warn "https/ribbon mode: on current Office-Mac builds (16.114 verified) the add-in registers"
  warn "but is NOT shown in the Add-ins list while commands are enabled - if it disappears,"
  warn "switch back with: ./install.sh --http  (the pane works in both transports)"
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
CN = OfficeKick Local CA (localhost-only, machine-generated)
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
  # - name-only checks are NOT enough: a regenerated CA keeps the name but not
  #   the keypair, and Office then rejects every https URL (add-in vanishes).
  local LOGIN_KC="$HOME/Library/Keychains/login.keychain-db"
  local KC_FP FILE_FP
  KC_FP="$(security find-certificate -c "OfficeKick Local CA" -p "$LOGIN_KC" 2>/dev/null \
    | openssl x509 -fingerprint -sha256 -noout 2>/dev/null | sed 's/^.*=//')"
  FILE_FP="$(openssl x509 -fingerprint -sha256 -noout -in "$CA_PEM" 2>/dev/null | sed 's/^.*=//')"
  if [[ -n "$KC_FP" && "$KC_FP" == "$FILE_FP" ]]; then
    ok "local CA already trusted (fingerprint match)"
  else
    # remove the old-branded CA and any stale copy under the new name
    security delete-certificate -c "SlideKick Local CA" "$LOGIN_KC" 2>/dev/null || true
    if [[ -n "$KC_FP" && "$KC_FP" != "$FILE_FP" ]]; then
      warn "stale local CA in the keychain (regenerated on disk) - replacing it"
      security delete-certificate -c "OfficeKick Local CA" "$LOGIN_KC" 2>/dev/null || true
    fi
    step "trusting the local CA (login keychain) - approve the mac security prompt"
    if security add-trusted-cert -r trustRoot -k "$LOGIN_KC" "$CA_PEM" 2>/dev/null; then
      ok "local CA trusted (login keychain) - the pane and ribbon can use https://localhost"
    else
      warn "could not auto-trust the local CA. Manual step: open Keychain Access,"
      warn "  import $CA_PEM and set 'Always Trust' for SSL - then reopen the pane"
    fi
  fi
  SCHEME="https"
}


# --------------------------------------------------------------------- #
# Office add-in cache maintenance
# --------------------------------------------------------------------- #
clear_office_caches() {
  # Only safe with the Office apps closed: the add-in commands registration
  # cache (and page caches) refresh while the apps are not running.
  if pgrep -f "Microsoft (PowerPoint|Word).app" >/dev/null 2>&1; then
    warn "Office apps are running - skipping the add-in cache clear."
    warn "quit PowerPoint/Word, re-run ./install.sh (it clears them), then reopen"
    return 0
  fi
  all_owned_containers
  local container cdir
  for container in ${ALL_OWNED[@]+"${ALL_OWNED[@]}"}; do
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
  say ""
  say "🚀  OfficeKick $VERSION - install/update (SlideKick + DocKick)"
  say ""
  PP_APP=""; PP_VERSION=""
  if find_powerpoint; then
    ok "PowerPoint found: $PP_APP (version ${PP_VERSION:-unknown})"
  else
    warn "Microsoft PowerPoint not found in /Applications or ~/Applications - installing anyway"
  fi
  W_APP=""; W_VERSION=""
  if [[ -d "/Applications/Microsoft Word.app" ]]; then
    W_APP="/Applications/Microsoft Word.app"
    W_VERSION="$(plutil -extract CFBundleShortVersionString raw "$W_APP/Contents/Info.plist" 2>/dev/null || true)"
    ok "Word found: $W_APP (version ${W_VERSION:-unknown})"
  elif [[ -d "$HOME/Applications/Microsoft Word.app" ]]; then
    W_APP="$HOME/Applications/Microsoft Word.app"
    ok "Word found: $W_APP"
  else
    warn "Microsoft Word not found - DocKick will be skipped (installed with the next run when present)"
  fi

  migrate_legacy
  all_owned_containers
  # stop previously installed servers first so the canonical port frees up
  step "stopping previously installed local server (if any)"
  launchctl unload "$PLIST" 2>/dev/null || true
  launchctl unload "$OLD_PLIST" 2>/dev/null || true
  pick_port
  ok "port $PORT is free and reserved for the server"

  if [[ "$DEV" == "1" ]]; then
    SERVER_PY="$SERVER_REPO"
    STATIC_DIR="$REPO_DIR"
    ok "dev mode: serving from $REPO_DIR (no copy)"
  else
    title "📦 Deploying files"
    say "    destination: $INSTALL_DIR"
    rm -rf "$INSTALL_DIR"
    mkdir -p "$INSTALL_DIR" "$CONFIG_DIR" "$LOG_DIR"
    # umbrella layout: the server sits at the root, one folder per add-in
    cp "$SERVER_REPO" "$INSTALL_DIR/server.py"
    if command -v rsync >/dev/null 2>&1; then
      rsync -a --exclude '.DS_Store' --exclude '__pycache__' "$SRC_DIR/" "$INSTALL_DIR/slidekick/"
    else
      mkdir -p "$INSTALL_DIR/slidekick"
      cp -R "$SRC_DIR/." "$INSTALL_DIR/slidekick/"
    fi
    find "$INSTALL_DIR" -name '__pycache__' -type d -exec rm -rf {} + 2>/dev/null || true
    if [[ -f "$WORDKICK_DIR/manifest.template.xml" ]]; then
      mkdir -p "$INSTALL_DIR/wordkick"
      if command -v rsync >/dev/null 2>&1; then
        rsync -a --exclude '.DS_Store' --exclude '__pycache__' "$WORDKICK_DIR/" "$INSTALL_DIR/wordkick/"
      else
        cp -R "$WORDKICK_DIR/." "$INSTALL_DIR/wordkick/"
      fi
      ok "files deployed: SlideKick + DocKick"
    else
      ok "files deployed: SlideKick (DocKick not built yet)"
    fi
    SERVER_PY="$INSTALL_DIR/server.py"
    STATIC_DIR="$INSTALL_DIR"
  fi

  "$PY" -m py_compile "$SERVER_PY" || die "server source failed to compile"
  ensure_local_certs

  # ---- manifests: one per add-in, sideloaded into the matching apps ----
  mkdir -p "$INSTALL_DIR/manifest"

  # render_manifest <repo_addin_dir> <out_manifest>
  render_manifest() {
    local addin_dir="$1" mtmp="$2"
    "${SED:-sed}" \
      -e "s/__SCHEME__/$SCHEME/" \
      -e "s/__PORT__/$PORT/g" \
      "$addin_dir/manifest.template.xml" > "$mtmp" || return 1
    # ALWAYS ship the list-mode manifest: Office-Mac (16.114 tested) accepts the
  # add-in + pane via the Add-ins dropdown, but never renders the classic
  # taskpaneappversionoverrides ribbon - commands mode would REMOVE the list
  # entry and silently show nothing. The --https ribbon attempt stays opt-in.
  "${SED:-sed}" -i '' -e '/__COMMANDS_BEGIN__/,/__COMMANDS_END__/d' "$mtmp"
    if command -v xmllint >/dev/null 2>&1; then
      xmllint --noout "$mtmp" || die "generated manifest $mtmp is not valid XML"
      ok "manifest rendered + valid XML: $(basename "$mtmp")"
    else
      warn "xmllint not available - skipping XML validation"
    fi
    return 0
  }

  local container wef installed_any=0 f
  # --- SlideKick (PowerPoint) ---
  render_manifest "$SRC_DIR" "$INSTALL_DIR/manifest/$MANIFEST_NAME"
  local pp_containers=()
  while IFS= read -r line; do pp_containers+=("$line"); done < <(
    resolve_containers_for "com.microsoft.[Pp]owerpoint" "com.microsoft.powerpoint" "com.microsoft.PowerPoint")
  for container in "${pp_containers[@]}"; do
    wef="$container/Data/Documents/wef"
    mkdir -p "$wef"
    for f in "$wef"/*.xml; do
      [[ -f "$f" ]] || continue
      if grep -q "$GUID" "$f" 2>/dev/null; then rm -f "$f"; fi
    done
    cp "$INSTALL_DIR/manifest/$MANIFEST_NAME" "$wef/$MANIFEST_NAME"
    ok "manifest sideloaded (SlideKick): $wef/$MANIFEST_NAME"
    installed_any=1
  done
  [[ "$installed_any" == "1" ]] || die "could not sideload SlideKick manifest into any PowerPoint container"

  # --- DocKick (Word): only once its manifest template exists ---
  if [[ -n "$W_APP" ]]; then
    if [[ -f "$WORDKICK_DIR/manifest.template.xml" ]]; then
      render_manifest "$WORDKICK_DIR" "$INSTALL_DIR/manifest/$WORD_MANIFEST_NAME"
      local w_containers=()
      while IFS= read -r line; do w_containers+=("$line"); done < <(
        resolve_containers_for "com.microsoft.[Ww]ord" "com.microsoft.word" "com.microsoft.Word")
      for container in "${w_containers[@]}"; do
        wef="$container/Data/Documents/wef"
        mkdir -p "$wef"
        for f in "$wef"/*.xml; do
          [[ -f "$f" ]] || continue
          if grep -q "$WORD_GUID" "$f" 2>/dev/null; then rm -f "$f"; fi
        done
        cp "$INSTALL_DIR/manifest/$WORD_MANIFEST_NAME" "$wef/$WORD_MANIFEST_NAME"
        ok "manifest sideloaded (DocKick): $wef/$WORD_MANIFEST_NAME"
      done
    else
      warn "DocKick (Word add-in) not built yet - skipping Word (it installs automatically once wordkick/manifest.template.xml exists)"
    fi
  fi

  clear_office_caches

  # ---- LaunchAgent ----
  # remove the legacy SlideKick label so only the OfficeKick agent remains
  if launchctl print "gui/$(id -u)/$OLD_LAUNCH_LABEL" >/dev/null 2>&1; then
    launchctl remove "$OLD_LAUNCH_LABEL" 2>/dev/null && ok "legacy LaunchAgent label removed"
  fi
  if [[ "$NO_AGENT" == "1" ]]; then
    warn "skipping LaunchAgent (--no-launchagent)"
    if [[ "$(port_busy "$PORT")" == "busy" ]]; then
      :
    else
      warn "server not running - start it manually:  \"$PY\" \"$SERVER_PY\" --port $PORT"
    fi
  else
    title "🚀 Server autostart"
    say "    LaunchAgent: $PLIST"
  "$PY" - "$PLIST" "$LAUNCH_LABEL" "$PY" "$SERVER_PY" "$PORT" "$STATIC_DIR" "$CERT_DIR" "$LOG_DIR" "$SCHEME" <<'PYPLIST'
import os, plistlib, sys

plist_path, label, py_bin, server_py, port, static_dir, cert_dir, log_dir, scheme = sys.argv[1:10]
args = [py_bin, server_py, "--port", port,
            "--config", os.path.expanduser("~/.officekick/config.json"),
            "--static-dir", static_dir]
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
  step "restarting local server via LaunchAgent (RunAtLoad + KeepAlive)"
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
    title "🩺 Health check"
    ok "server healthy: $SCHEME://127.0.0.1:$PORT"
    ok "config (keys, chmod 600): $CONFIG_DIR/config.json"
  else
    die "server did not become healthy on port $PORT - check $LOG_DIR/server.log"
  fi

  title "📦 Installed"
  ok "add-in manifests : $INSTALL_DIR/manifest/ (copied into each Office app's WEF folder)"
  ok "app files        : $INSTALL_DIR"
  ok "server log       : $LOG_DIR/server.log"
  ok "simulator (PPT)  : $SCHEME://127.0.0.1:$PORT/slidekick/static/simulator.html  (test without Office)"
  title "👋  First use - PowerPoint:"
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
  if [[ -n "$W_APP" ]]; then
    if [[ -f "$WORDKICK_DIR/manifest.template.xml" ]]; then
      title "👋  First use - Word:"
      if [[ "$HTTPS" == "1" ]]; then
        say "  3. the Corelan DocKick button is in the Home tab"
      else
        say "  3. Home tab -> Add-ins (dropdown) -> Developer Add-ins -> Corelan DocKick"
      fi
      say "     (profiles/connections are shared with SlideKick)"
    fi
  fi
  say ""
  say "✨  All set."
  say ""
}

# --------------------------------------------------------------------- #
# uninstall
# --------------------------------------------------------------------- #
do_uninstall() {
  say "OfficeKick - uninstall (SlideKick + DocKick)"
  launchctl unload "$PLIST" 2>/dev/null && ok "LaunchAgent unloaded" || warn "no LaunchAgent"
  [[ -f "$PLIST" ]] && rm -f "$PLIST" && ok "LaunchAgent removed"
  # legacy SlideKick-era agent
  if [[ -f "$OLD_PLIST" ]]; then
    launchctl unload "$OLD_PLIST" 2>/dev/null || true
    rm -f "$OLD_PLIST"
    ok "legacy LaunchAgent removed"
  fi
  all_owned_containers
  local container wef f removed=0
  for container in ${ALL_OWNED[@]+"${ALL_OWNED[@]}"}; do
    [[ -n "$container" && -d "$container/Data/Documents/wef" ]] || continue
    wef="$container/Data/Documents/wef"
    for f in "$wef"/*.xml; do
      [[ -f "$f" ]] || continue
      if grep -q "$GUID" "$f" 2>/dev/null || grep -q "$WORD_GUID" "$f" 2>/dev/null; then
        rm -f "$f"; removed=1
      fi
    done
  done
  [[ "$removed" == "1" ]] && ok "manifests removed from Office container(s)"
  if pgrep -f "$INSTALL_DIR/server.py" >/dev/null 2>&1; then
    pkill -f "OfficeKick/server.py" 2>/dev/null && ok "server stopped" || warn "could not stop server process"
  fi
  pgrep -f "SlideKick/slidekick_server.py" >/dev/null 2>&1 && pkill -f "SlideKick/slidekick_server.py" 2>/dev/null || true
  if [[ -d "$INSTALL_DIR" && "$DEV" == "0" ]]; then
    rm -rf "$INSTALL_DIR" && ok "$INSTALL_DIR removed"
  fi
  [[ -d "$OLD_INSTALL_DIR" ]] && rm -rf "$OLD_INSTALL_DIR" && ok "legacy dir removed" || true
  if [[ "$PURGE" == "1" ]]; then
    rm -rf "$CONFIG_DIR" && ok "config wiped ($CONFIG_DIR)"
    rm -rf "$OLD_CONFIG_DIR" && ok "legacy config wiped ($OLD_CONFIG_DIR)"
  else
    ok "config kept at $CONFIG_DIR (use --purge to wipe)"
  fi
  warn "quit and reopen PowerPoint/Word to finish removing the add-ins"
}

# --------------------------------------------------------------------- #
# status
# --------------------------------------------------------------------- #
do_status() {
  resolve_containers
  local h
  h="$(health "$PORT" 2>/dev/null || health "$DEFAULT_PORT" 2>/dev/null || true)"
  if [[ "$h" == *OfficeKick* ]]; then
    ok "server is running: $h"
  else
    warn "server NOT running (default port $DEFAULT_PORT)"
  fi
  if launchctl print "gui/$(id -u)/$LAUNCH_LABEL" >/dev/null 2>&1; then
    ok "LaunchAgent present"
  else
    warn "LaunchAgent absent"
  fi
  [[ -f "$OLD_PLIST" ]] && warn "legacy LaunchAgent still present: $OLD_PLIST (re-run ./install.sh to migrate)"
  local container wef found_sk="" found_dk=""
  for container in "${CONTAINER_PATHS[@]}"; do
    wef="$container/Data/Documents/wef/$MANIFEST_NAME"
    [[ -f "$wef" ]] && found_sk="$wef"
    wef="$container/Data/Documents/wef/$WORD_MANIFEST_NAME"
    [[ -f "$wef" ]] && found_dk="$wef"
  done
  [[ -n "$found_sk" ]] && ok "SlideKick manifest sideloaded: $found_sk" || warn "SlideKick manifest not sideloaded"
  if [[ -n "$found_dk" ]]; then
    ok "DocKick manifest sideloaded: $found_dk"
  elif [[ -f "$WORDKICK_DIR/manifest.template.xml" ]]; then
    warn "DocKick manifest not sideloaded"
  else
    ok "DocKick: not built yet (wordkick/ has no manifest template)"
  fi
  [[ -f "$CONFIG_DIR/config.json" ]] && ok "config present: $CONFIG_DIR/config.json" || warn "no config yet"
  [[ -f "$OLD_CONFIG_DIR/config.json" && ! -f "$CONFIG_DIR/config.json" ]] && warn "legacy config found but not migrated (run ./install.sh)"
  [[ -d "$INSTALL_DIR" ]] && ok "app dir present: $INSTALL_DIR" || warn "app dir missing (or dev mode)"
}

# --------------------------------------------------------------------- #
# run in foreground
# --------------------------------------------------------------------- #
run_foreground() {
  if [[ "$DEV" == "1" ]]; then
    SERVER_PY="$SERVER_REPO"
    STATIC_DIR="$REPO_DIR"
  else
    SERVER_PY="$INSTALL_DIR/server.py"
    STATIC_DIR="$INSTALL_DIR"
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