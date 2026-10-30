#!/usr/bin/env python3
"""
SlideKick - local companion server for the SlideKick PowerPoint add-in.

Responsibilities:
  * Serve the task-pane UI (static files, no-cache).
  * Store platform connections, profiles and the per-action approval policy
    in a local config file (API keys never leave this machine's config).
  * Proxy AI chat + model-list requests to the configured provider
    (openai / anthropic / openrouter / generic-openai / ollama / llamacpp).

Runs on 127.0.0.1 only. Python 3.8+ standard library, no third-party deps.
"""

import argparse
import json
import os
import socket
import ssl
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

APP = "SlideKick"
VERSION = "0.1.2"

CONNECTION_TYPES = (
    "openai",
    "anthropic",
    "openrouter",
    "generic-openai",
    "ollama",
    "llamacpp",
)

DEFAULT_BASES = {
    "openai": "https://api.openai.com/v1",
    "anthropic": "https://api.anthropic.com",
    "openrouter": "https://openrouter.ai/api/v1",
    "generic-openai": "",
    "ollama": "http://localhost:11434",
    "llamacpp": "http://localhost:8080",
}

# Per-action approval policy for deck-mutating ops returned by the model.
# Modes: auto     -> apply without asking
#        confirm  -> user must tick the op before Apply
#        deny     -> op is blocked entirely (shown as blocked, never applied)
DEFAULT_APPROVALS = {
    "create_slide": "auto",
    "edit_slide": "auto",
    "replace_text": "auto",
    "set_notes": "auto",
    "apply_layout": "confirm",
    "delete_slide": "confirm",
}
APPROVAL_MODES = ("auto", "confirm", "deny")
OP_TYPES = tuple(DEFAULT_APPROVALS.keys())

ANTHROPIC_FALLBACK_MODELS = (
    "claude-sonnet-4-5",
    "claude-haiku-4-5",
    "claude-opus-4-1",
    "claude-3-5-haiku-latest",
)

MAX_BODY_BYTES = 2 * 1024 * 1024


class ApiError(Exception):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status
        self.message = message


# --------------------------------------------------------------------------- #
# Config store
# --------------------------------------------------------------------------- #
class Store:
    def __init__(self, path):
        self.path = Path(path).expanduser()
        self.path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
        try:
            os.chmod(self.path.parent, 0o700)
        except OSError:
            pass
        self._lock = threading.Lock()
        self.data = None
        self.force_debug = False  # CLI --debug override

    def load(self):
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text(encoding="utf-8"))
            except Exception as exc:  # unreadable config: refuse to clobber
                raise SystemExit(
                    "Cannot read config file %s: %s" % (self.path, exc)
                )
        else:
            self.data = {}
        self.data.setdefault("version", 1)
        self.data.setdefault("connections", [])
        self.data.setdefault("profiles", [])
        self.data.setdefault("defaultProfile", None)
        approvals = self.data.setdefault("approvals", dict(DEFAULT_APPROVALS))
        for op in OP_TYPES:
            approvals.setdefault(op, DEFAULT_APPROVALS[op])
        self.data.setdefault("debug", False)
        self.save()

    def save(self):
        with self._lock:
            tmp = self.path.with_suffix(".tmp")
            tmp.write_text(
                json.dumps(self.data, indent=2) + "\n", encoding="utf-8"
            )
            os.chmod(tmp, 0o600)
            tmp.replace(self.path)
            os.chmod(self.path, 0o600)

    def debug_enabled(self):
        return bool(self.data.get("debug") or self.force_debug)

    def set_debug(self, enabled):
        self.data["debug"] = bool(enabled)
        self.save()

    # ------------------------------------------------------------------ #
    def find_connection(self, name):
        for c in self.data["connections"]:
            if c["name"] == name:
                return c
        return None

    def find_profile(self, name):
        for p in self.data["profiles"]:
            if p["name"] == name:
                return p
        return None

    def upsert_connection(self, name, ctype, base_url, api_key):
        existing = self.find_connection(name)
        if existing is None:
            conn = {"name": name, "type": ctype, "baseUrl": base_url, "apiKey": None}
            self.data["connections"].append(conn)
        else:
            conn = existing
            conn["type"] = ctype
            conn["baseUrl"] = base_url
        if api_key:  # only overwrite when the caller actually sent one
            conn["apiKey"] = api_key
        # keep a side-channel mapping name->type? not needed
        self.save()
        return conn

    def delete_connection(self, name):
        users = [
            p["name"] for p in self.data["profiles"]
            if (p.get("content") or {}).get("connection") == name
            or (p.get("other") or {}).get("connection") == name
        ]
        if users:
            raise ApiError(
                409,
                "Connection '%s' is used by profile(s): %s - remove it from "
                "those profiles first." % (name, ", ".join(users)),
            )
        conns = self.data["connections"]
        for i, c in enumerate(conns):
            if c["name"] == name:
                del conns[i]
                if self.data.get("defaultProfile") and not self.find_profile(
                    self.data["defaultProfile"]
                ):
                    self.data["defaultProfile"] = None
                self.save()
                return True
        return False

    def upsert_profile(self, name, content, other):
        prof = {
            "name": name,
            "content": {"connection": content[0], "model": content[1]},
            "other": {"connection": other[0], "model": other[1]},
        }
        for i, p in enumerate(self.data["profiles"]):
            if p["name"] == name:
                self.data["profiles"][i] = prof
                break
        else:
            self.data["profiles"].append(prof)
        if not self.data.get("defaultProfile"):
            self.data["defaultProfile"] = name
        self.save()
        return prof

    def delete_profile(self, name):
        profs = self.data["profiles"]
        for i, p in enumerate(profs):
            if p["name"] == name:
                del profs[i]
                if self.data.get("defaultProfile") == name:
                    self.data["defaultProfile"] = (
                        profs[0]["name"] if profs else None
                    )
                self.save()
                return True
        return False

    def rename_connection(self, src, dst, updates=None):
        """Rename a connection (type/base/key editable in the same call);
        profile references are re-pointed so nothing orphans."""
        updates = updates or {}
        conn = self.find_connection(src)
        if conn is None:
            raise ApiError(404, "Unknown connection: %s" % src)
        if src != dst and self.find_connection(dst):
            raise ApiError(409, "Connection '%s' already exists" % dst)
        if not valid_name(dst):
            raise ApiError(400, "New connection name must be 1-64 chars, no leading/trailing spaces or '/'")
        t = updates.get("type") or conn["type"]
        if t not in CONNECTION_TYPES:
            raise ApiError(400, "Unknown connection type: %r" % (t,))
        base = (updates.get("baseUrl") or "").strip() or conn["baseUrl"]
        if not base:
            raise ApiError(400, "Base URL is required for connection type %s" % t)
        key = updates.get("apiKey")
        if key:
            conn["apiKey"] = key
        conn["type"] = t
        conn["baseUrl"] = base.rstrip("/")
        conn["name"] = dst
        for prof in self.data["profiles"]:
            for slot in ("content", "other"):
                ref = prof.get(slot) or {}
                if ref.get("connection") == src:
                    ref["connection"] = dst
        self.save()
        return conn

    def rename_profile(self, src, dst, updates=None):
        """Rename a profile (models editable in the same call);
        the default-profile pointer follows the rename."""
        updates = updates or {}
        prof = self.find_profile(src)
        if prof is None:
            raise ApiError(404, "Unknown profile: %s" % src)
        if not valid_name(dst):
            raise ApiError(400, "New profile name must be 1-64 chars")
        if src != dst and self.find_profile(dst):
            raise ApiError(409, "Profile '%s' already exists" % dst)
        prof["name"] = dst
        if self.data.get("defaultProfile") == src:
            self.data["defaultProfile"] = dst
        for slot in ("content", "other"):
            ref = updates.get(slot) or {}
            if ref.get("connection") is not None:
                if not self.find_connection(ref["connection"]):
                    raise ApiError(400, "Profile %s references unknown connection: %r" % (slot, ref["connection"]))
                model = ref.get("model")
                if not isinstance(model, str) or not model.strip():
                    raise ApiError(400, "Profile %s has no model selected" % slot)
                prof[slot] = {"connection": ref["connection"], "model": model.strip()}
        self.save()
        return prof

    def set_approvals(self, approvals):
        clean = dict(self.data.get("approvals") or DEFAULT_APPROVALS)
        for op, mode in (approvals or {}).items():
            if op not in OP_TYPES:
                raise ApiError(400, "Unknown action type: %s" % op)
            if mode not in APPROVAL_MODES:
                raise ApiError(
                    400,
                    "Invalid approval mode %r for %s (use auto|confirm|deny)"
                    % (mode, op),
                )
            clean[op] = mode
        self.data["approvals"] = clean
        self.save()

    def public(self):
        d = self.data
        return {
            "version": VERSION,
            "configPath": str(self.path),
            "connections": [
                {
                    "name": c["name"],
                    "type": c["type"],
                    "baseUrl": c["baseUrl"],
                    "hasKey": bool(c.get("apiKey")),
                }
                for c in d["connections"]
            ],
            "profiles": d["profiles"],
            "defaultProfile": d.get("defaultProfile"),
            "approvals": d.get("approvals") or dict(DEFAULT_APPROVALS),
            "debug": bool(d.get("debug")),
        }


# --------------------------------------------------------------------------- #
# Debug logging (daily append file, opt-in)
# --------------------------------------------------------------------------- #
DBG_LOCK = threading.Lock()


def dbg_line(store, source, category, payload):
    """Append one JSON line to ~/.slidekick/logs/debug-YYYY-MM-DD.log.

    One file per day, appended. Only writes when the user turned debugging on
    (Settings toggle in the task pane, or --debug on the command line)."""
    try:
        if store is None or not store.debug_enabled():
            return
        import datetime

        logdir = store.path.parent / "logs"
        logdir.mkdir(mode=0o700, parents=True, exist_ok=True)
        try:
            os.chmod(logdir, 0o700)
        except OSError:
            pass
        now = datetime.datetime.now()
        path = logdir / ("debug-%s.log" % now.strftime("%Y-%m-%d"))
        entry = {
            "ts": now.strftime("%Y-%m-%d %H:%M:%S"),
            "source": source,
            "category": category,
            "data": payload,
        }
        line = json.dumps(entry, ensure_ascii=False, default=str)
        if len(line) > 400000:
            line = line[:400000] + " ...<truncated>"
        with DBG_LOCK:
            with open(path, "a", encoding="utf-8") as fh:
                fh.write(line + "\n")
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Provider plumbing
# --------------------------------------------------------------------------- #
def _host_of(url):
    try:
        return urllib.parse.urlsplit(url).netloc or url
    except Exception:
        return url


def _snippet(text, limit=300):
    text = (text or "").strip().replace("\n", " ")
    return text[:limit] + ("..." if len(text) > limit else "")


def http_request(url, method="GET", headers=None, payload=None, timeout=60):
    """Returns parsed JSON (dict/list) or raw text. Raises ApiError on failure."""
    headers = dict(headers or {})
    data = None
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers.setdefault("Content-Type", "application/json")
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read(MAX_BODY_BYTES * 4)
            charset = resp.headers.get_content_charset() or "utf-8"
            text = raw.decode(charset, "replace")
    except urllib.error.HTTPError as exc:
        body = ""
        try:
            body = exc.read(8192).decode("utf-8", "replace")
        except Exception:
            pass
        raise ApiError(
            502,
            "Provider returned HTTP %s for %s %s: %s"
            % (exc.code, method, _host_of(url), _snippet(body)),
        )
    except urllib.error.URLError as exc:
        raise ApiError(
            502,
            "Cannot reach %s (%s). Is the service running?"
            % (_host_of(url), getattr(exc, "reason", exc)),
        )
    except (socket.timeout, TimeoutError):
        raise ApiError(502, "Timeout talking to %s (>%ss)" % (_host_of(url), timeout))
    try:
        return json.loads(text)
    except Exception:
        return text


def normalize_base(conn):
    base = (conn.get("baseUrl") or "").strip().rstrip("/")
    if not base:
        base = DEFAULT_BASES.get(conn["type"], "")
    if not base:
        raise ApiError(
            400,
            "Connection %r has no Base URL (required for type %s)"
            % (conn["name"], conn["type"]),
        )
    return base


def conn_headers(conn):
    t = conn["type"]
    key = conn.get("apiKey") or ""
    h = {"User-Agent": "%s/%s" % (APP, VERSION)}
    if t in ("openai", "generic-openai", "openrouter", "llamacpp"):
        if key:
            h["Authorization"] = "Bearer " + key
        if t == "openrouter":
            h["HTTP-Referer"] = "https://github.com/corelan/slidekick"
            h["X-Title"] = APP
    elif t == "anthropic":
        if key:
            h["x-api-key"] = key
        h["anthropic-version"] = "2023-06-01"
    return h


def _parse_openai_models(data):
    models = []
    if isinstance(data, dict) and isinstance(data.get("data"), list):
        for m in data["data"]:
            if isinstance(m, dict) and m.get("id"):
                models.append(str(m["id"]))
    elif isinstance(data, list):
        for m in data:
            if isinstance(m, dict) and m.get("id"):
                models.append(str(m["id"]))
    return models


def _dedupe(seq):
    seen = set()
    out = []
    for item in seq:
        if item and item not in seen:
            seen.add(item)
            out.append(item)
    return out


def provider_list_models(store, conn_name, store_for_log=True):
    conn = store.find_connection(conn_name)
    if conn is None:
        raise ApiError(404, "Unknown connection: %s" % conn_name)
    return provider_list_models_conn(conn, store if store_for_log else None)


def provider_list_models_conn(conn, store=None):
    """List models for a connection object (saved or ad-hoc for testing)."""
    t = conn["type"]
    base = normalize_base(conn)
    headers = conn_headers(conn)

    models, source = [], "provider"
    if t == "ollama":
        data = http_request(base + "/api/tags", headers=headers, timeout=30)
        if isinstance(data, dict):
            for m in data.get("models") or []:
                models.append(str(m.get("name") or m.get("model") or ""))
    elif t == "anthropic":
        try:
            data = http_request(
                base + "/v1/models?limit=200", headers=headers, timeout=30
            )
            if isinstance(data, dict):
                for m in data.get("data") or []:
                    models.append(str(m.get("id") or ""))
        except ApiError:
            models = list(ANTHROPIC_FALLBACK_MODELS)
            source = "fallback"
    elif t == "llamacpp":
        for url in (base + "/v1/models", base + "/models"):
            try:
                data = http_request(url, headers=headers, timeout=30)
                models = _parse_openai_models(data)
                if models:
                    break
            except ApiError:
                continue
        if not models:
            try:
                props = http_request(base + "/props", headers=headers, timeout=15)
                if isinstance(props, dict):
                    mid = props.get("model_id") or (props.get("metadata") or {}).get("model_id")
                    if mid:
                        models = [str(mid)]
            except ApiError:
                pass
    else:  # openai / generic-openai / openrouter
        urls = [base + "/models"]
        if t == "generic-openai" and not base.endswith("/v1"):
            urls.append(base + "/v1/models")
        last_error = None
        for url in urls:
            try:
                models = _parse_openai_models(
                    http_request(url, headers=headers, timeout=30)
                )
                if models:
                    return {"connection": conn["name"], "source": "provider", "models": models}
            except ApiError as exc:
                last_error = exc
        if last_error is not None:
            raise last_error

    models = _dedupe(models)
    if not models:
        raise ApiError(
            502,
            "Provider %s (%s) returned no models. Check the Base URL/credentials."
            % (conn["name"], t),
        )
    dbg_line(store, "server", "models.list", {
        "connection": conn["name"], "source": source, "count": len(models),
        "models": models[:50],
    })
    return {"connection": conn["name"], "source": source, "models": models}


# Chain-of-thought tag pairs stripped from model output before parsing:
# think / thinking / reason / reasoning - balanced or truncated.
_THINK_TAGS = ("think", "thinking", "reason", "reasoning")
_THINK_RES = None  # compiled lazily


def _strip_thinking(text):
    """Strip chain-of-thought wrappers so only the real answer remains.

    Tag regexes are built with chr() composition + re.escape, so the raw
    opener/closer marker sequence never appears verbatim in this file.
    A truncated block (no closer) strips everything from the opener on:
    an empty result then triggers the client hidden-repair round."""
    global _THINK_RES
    if _THINK_RES is None:
        import re

        lt, gt = chr(60), chr(62)
        _THINK_RES = []
        for tag in _THINK_TAGS:
            opener = lt + tag + gt
            closer = lt + "/" + tag + gt
            _THINK_RES.append(re.compile(
                re.escape(opener) + r".*?" + re.escape(closer) + r"\s*",
                re.DOTALL | re.IGNORECASE))
            _THINK_RES.append(re.compile(
                re.escape(opener) + r".*$",
                re.DOTALL | re.IGNORECASE))
    out = str(text or "")
    for rx in _THINK_RES:
        out = rx.sub("", out)
    return out.strip()


def _openai_choice_text(data):
    if not isinstance(data, dict):
        return str(data)
    choices = data.get("choices") or []
    if not choices:
        return ""
    msg = choices[0].get("message") or {}
    content = msg.get("content")
    if isinstance(content, list):  # some OpenAI-compatible servers return parts
        parts = []
        for part in content:
            if isinstance(part, dict):
                parts.append(str(part.get("text") or ""))
            else:
                parts.append(str(part))
        content = "\n".join(p for p in parts if p)
    text = str(content or "")
    reasoning = msg.get("reasoning_content") or msg.get("reasoning")
    if reasoning and not text.strip():
        text = ""  # reasoning-only response without a final answer
    return text


def provider_chat(store, conn_name, model, messages, temperature, max_tokens, timeout):
    conn = store.find_connection(conn_name)
    if conn is None:
        raise ApiError(404, "Unknown connection: %s" % conn_name)
    t = conn["type"]
    base = normalize_base(conn)
    headers = conn_headers(conn)

    msgs = []
    for m in messages or []:
        role = str(m.get("role") or "user")
        if role not in ("system", "user", "assistant"):
            raise ApiError(400, "Invalid message role: %r" % role)
        msgs.append({"role": role, "content": str(m.get("content") or "")})
    if not msgs:
        raise ApiError(400, "No messages provided")

    started = time.time()
    dbg_line(store, "server", "provider.request", {
        "connection": conn_name, "model": model,
        "temperature": temperature, "max_tokens": max_tokens,
        "messages": msgs,
    })

    if t == "anthropic":
        system = "\n\n".join(m["content"] for m in msgs if m["role"] == "system").strip()
        rest = []
        for m in msgs:
            if m["role"] == "system":
                continue
            # Anthropic requires alternating roles; merge consecutive same-role.
            if rest and rest[-1]["role"] == m["role"]:
                rest[-1]["content"] += "\n\n" + m["content"]
            else:
                rest.append(dict(m))
        if not rest or rest[0]["role"] != "user":
            rest.insert(0, {"role": "user", "content": "(no user message)"})
        payload = {
            "model": model,
            "max_tokens": int(max_tokens),
            "temperature": float(temperature),
            "messages": rest,
        }
        if system:
            payload["system"] = system
        data = http_request(
            base + "/v1/messages", method="POST", headers=headers,
            payload=payload, timeout=timeout,
        )
        text = ""
        if isinstance(data, dict):
            for block in data.get("content") or []:
                if isinstance(block, dict) and block.get("type") == "text":
                    text += block.get("text") or ""
    elif t == "ollama":
        payload = {
            "model": model,
            "messages": msgs,
            "stream": False,
            "options": {
                "temperature": float(temperature),
                "num_predict": int(max_tokens),
            },
        }
        data = http_request(
            base + "/api/chat", method="POST", headers=headers,
            payload=payload, timeout=timeout,
        )
        text = ""
        if isinstance(data, dict):
            message = data.get("message") or {}
            # reasoning-capable ollama models put chain-of-thought in a
            # separate "thinking" field - it is never part of the answer
            text = str(message.get("content") or "")
    elif t == "llamacpp":
        payload = {
            "model": model,
            "messages": msgs,
            "stream": False,
            "temperature": float(temperature),
            "max_tokens": int(max_tokens),
        }
        data = http_request(
            base + "/v1/chat/completions", method="POST", headers=headers,
            payload=payload, timeout=timeout,
        )
        text = _openai_choice_text(data)
    else:  # openai / generic-openai / openrouter
        payload = {
            "model": model,
            "messages": msgs,
            "temperature": float(temperature),
            "max_tokens": int(max_tokens),
            "stream": False,
        }
        urls = [base + "/chat/completions"]
        if t == "generic-openai" and not base.endswith("/v1"):
            urls.append(base + "/v1/chat/completions")
        data, last_error = None, None
        for url in urls:
            try:
                data = http_request(
                    url, method="POST", headers=headers, payload=payload,
                    timeout=timeout,
                )
                text = _openai_choice_text(data)
                if text:
                    break
            except ApiError as exc:
                last_error = exc
        if not text and last_error is not None:
            raise last_error

    text = _strip_thinking(text or "")
    dbg_line(store, "server", "provider.response", {
        "connection": conn_name, "model": model, "raw_text": text,
        "elapsedMs": int((time.time() - started) * 1000),
    })
    return {"connection": conn_name, "model": model, "text": text}


# --------------------------------------------------------------------------- #
# Request validation helpers
# --------------------------------------------------------------------------- #
def valid_name(name):
    return (
        isinstance(name, str)
        and 0 < len(name.strip()) <= 64
        and name.strip() == name
        and "/" not in name
        and "\x00" not in name
    )


def validate_connection_body(body):
    name = body.get("name")
    ctype = body.get("type")
    if not valid_name(name):
        raise ApiError(400, "Connection name must be 1-64 chars, no leading/trailing spaces or '/'")
    if ctype not in CONNECTION_TYPES:
        raise ApiError(400, "Unknown connection type: %r" % (ctype,))
    base = (body.get("baseUrl") or "").strip()
    if not base:
        base = DEFAULT_BASES.get(ctype, "")
    if not base:
        raise ApiError(400, "Base URL is required for connection type %s" % ctype)
    if not base.startswith(("http://", "https://")):
        raise ApiError(400, "Base URL must start with http:// or https://")
    api_key = body.get("apiKey")
    if api_key is not None and not isinstance(api_key, str):
        raise ApiError(400, "apiKey must be a string")
    return name.strip(), ctype, base.strip().rstrip("/"), None if not api_key else api_key


def validate_profile_body(store, body):
    name = body.get("name")
    if not valid_name(name):
        raise ApiError(400, "Profile name must be 1-64 chars")
    refs = []
    for slot in ("content", "other"):
        ref = body.get(slot) or {}
        conn, model = ref.get("connection"), ref.get("model")
        if not store.find_connection(conn):
            raise ApiError(
                400, "Profile %s references unknown connection: %r" % (slot, conn)
            )
        if not isinstance(model, str) or not model.strip():
            raise ApiError(400, "Profile %s has no model selected" % slot)
        refs.append((conn, model.strip()))
    return name.strip(), refs[0], refs[1]


# --------------------------------------------------------------------------- #
# HTTP server
# --------------------------------------------------------------------------- #
STATIC_DOCS = {
    "/": "taskpane.html",
    "/taskpane": "taskpane.html",
}


class SlideKickHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    # filled by build_server()
    store = None
    static_dir = None
    verbose = False
    serving_port = 0

    def log_message(self, fmt, *args):
        if self.verbose:
            sys.stderr.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), fmt % args))

    # ---------------------- plumbing ---------------------- #
    def _send_json(self, obj, status=200):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _error(self, status, message):
        self._send_json({"error": message}, status)

    def _body_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        if length > MAX_BODY_BYTES:
            raise ApiError(413, "Request body too large")
        raw = self.rfile.read(length)
        try:
            return json.loads(raw.decode("utf-8"))
        except Exception as exc:
            raise ApiError(400, "Invalid JSON body: %s" % exc)

    def _static(self, path):
        sdir = Path(self.static_dir).resolve()
        rel = urllib.parse.urlsplit(path).path
        if rel in STATIC_DOCS:
            rel = STATIC_DOCS[rel]
        rel = rel.lstrip("/")
        target = (sdir / rel).resolve()
        is_simulator = target.name == "simulator.html"
        if not str(target).startswith(str(sdir)) or (not is_simulator and not target.is_file()):
            self._error(404, "Not found: %s" % rel)
            return
        suffix = target.suffix.lower()
        ctype = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".png": "image/png",
            ".svg": "image/svg+xml",
            ".ico": "image/x-icon",
            ".json": "application/json; charset=utf-8",
        }.get(suffix, "application/octet-stream")
        if target.name == "simulator.html":
            # simulator = taskpane markup with sim.js instead of Office.js CDN;
            # generated from taskpane.html so the two never drift apart
            src = (sdir / "taskpane.html").read_text(encoding="utf-8")
            src = src.replace(
                '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>',
                '<script src="sim.js"></script>',
            )
            src = src.replace("<title>SlideKick</title>",
                              "<title>SlideKick - simulator</title>")
            data = src.encode("utf-8")
        else:
            data = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("X-Content-Type-Options", "nosniff")
        if suffix in (".html", ".css", ".js"):
            self.send_header("Cache-Control", "no-store")
        else:
            self.send_header("Cache-Control", "max-age=3600")
        self.end_headers()
        self.wfile.write(data)

    # ---------------------- routing ---------------------- #
    def do_GET(self):
        try:
            path = urllib.parse.urlsplit(self.path).path
            query = urllib.parse.parse_qs(urllib.parse.urlsplit(self.path).query)
            if path == "/api/health":
                self._send_json(
                    {
                        "ok": True,
                        "app": APP,
                        "version": VERSION,
                        "port": self.serving_port,
                        "configPath": str(self.store.path),
                        "debug": self.store.debug_enabled(),
                    }
                )
            elif path == "/api/config":
                pub = self.store.public()
                pub["port"] = self.serving_port
                pub["connectionTypes"] = list(CONNECTION_TYPES)
                pub["defaultBases"] = DEFAULT_BASES
                pub["approvalModes"] = list(APPROVAL_MODES)
                pub["opTypes"] = list(OP_TYPES)
                self._send_json(pub)
            elif path == "/api/models":
                conn = (query.get("connection") or [""])[0]
                self._send_json(provider_list_models(self.store, conn))
            elif path.startswith("/api/"):
                self._error(404, "Unknown API endpoint: %s" % path)
            else:
                self._static(path)
        except ApiError as exc:
            self._error(exc.status, exc.message)
        except BrokenPipeError:
            pass
        except Exception as exc:  # pragma: no cover
            import traceback

            traceback.print_exc()
            self._error(500, "Internal error: %s" % exc)

    def do_POST(self):
        try:
            path = urllib.parse.urlsplit(self.path).path
            body = self._body_json()
            if path == "/api/connections":
                name, ctype, base, key = validate_connection_body(body)
                self.store.upsert_connection(name, ctype, base, key)
                self._send_json({"ok": True, "name": name})
            elif path == "/api/profiles":
                name, content, other = validate_profile_body(self.store, body)
                self.store.upsert_profile(name, content, other)
                self._send_json({"ok": True, "name": name})
            elif path == "/api/connections/rename":
                self.store.rename_connection(
                    str(body.get("from") or ""), str(body.get("to") or ""), body)
                self._send_json({"ok": True, "name": body.get("to")})
            elif path == "/api/profiles/rename":
                self.store.rename_profile(
                    str(body.get("from") or ""), str(body.get("to") or ""), body)
                self._send_json({"ok": True, "name": body.get("to")})
            elif path == "/api/chat":
                conn = body.get("connection")
                if not self.store.find_connection(conn or ""):
                    raise ApiError(404, "Unknown connection: %r" % (conn,))
                model = str(body.get("model") or "").strip()
                if not model:
                    raise ApiError(400, "No model specified")
                temperature = body.get("temperature", 0.4)
                try:
                    temperature = min(2.0, max(0.0, float(temperature)))
                except Exception:
                    raise ApiError(400, "Invalid temperature")
                max_tokens = body.get("maxTokens", 4000)
                try:
                    max_tokens = min(32000, max(16, int(max_tokens)))
                except Exception:
                    raise ApiError(400, "Invalid maxTokens")
                timeout = body.get("timeout", 300)
                try:
                    timeout = min(600, max(10, int(timeout)))
                except Exception:
                    raise ApiError(400, "Invalid timeout")
                started = time.time()
                result = provider_chat(
                    self.store, conn, model, body.get("messages") or [],
                    temperature, max_tokens, timeout,
                )
                result["elapsedMs"] = int((time.time() - started) * 1000)
                self._send_json(result)
            elif path == "/api/test-connection":
                # Test a connection with the values currently in the form.
                # apiKey absent/"" -> reuse the stored key when the name matches.
                name = body.get("name")
                ctype = body.get("type")
                base = (body.get("baseUrl") or "").strip()
                key = body.get("apiKey")
                existing = self.store.find_connection(name if name else "") if name else None
                if ctype not in CONNECTION_TYPES:
                    ctype = (existing or {}).get("type", "openai")
                if not base:
                    base = (existing or {}).get("baseUrl") or DEFAULT_BASES.get(ctype, "")
                if key is None or key == "":
                    key = (existing or {}).get("apiKey")
                conn = {
                    "name": name or "(test)",
                    "type": ctype,
                    "baseUrl": base.rstrip("/"),
                    "apiKey": key,
                }
                self._send_json(provider_list_models_conn(conn, self.store))
            elif path == "/api/debug":
                dbg_line(self.store, "taskpane",
                         str(body.get("category") or "event"), body.get("data"))
                self._send_json({"ok": True})
            else:
                self._error(404, "Unknown API endpoint: %s" % path)
        except ApiError as exc:
            self._error(exc.status, exc.message)
        except BrokenPipeError:
            pass
        except Exception as exc:  # pragma: no cover
            import traceback

            traceback.print_exc()
            self._error(500, "Internal error: %s" % exc)

    def do_PUT(self):
        try:
            path = urllib.parse.urlsplit(self.path).path
            body = self._body_json()
            if path == "/api/settings":
                changed = False
                if "approvals" in body and body["approvals"] is not None:
                    self.store.set_approvals(body["approvals"])
                    changed = True
                if "defaultProfile" in body:
                    dp = body["defaultProfile"]
                    if dp is not None and not self.store.find_profile(dp):
                        raise ApiError(400, "Unknown profile: %r" % (dp,))
                    self.store.data["defaultProfile"] = dp
                    changed = True
                if "debug" in body and body["debug"] is not None:
                    self.store.set_debug(bool(body["debug"]))
                    changed = True
                if changed:
                    self.store.save()
                self._send_json({"ok": True, "approvals": self.store.data["approvals"]})
            else:
                self._error(404, "Unknown API endpoint: %s" % path)
        except ApiError as exc:
            self._error(exc.status, exc.message)
        except Exception as exc:  # pragma: no cover
            import traceback

            traceback.print_exc()
            self._error(500, "Internal error: %s" % exc)

    def do_DELETE(self):
        try:
            path = urllib.parse.urlsplit(self.path).path
            if path.startswith("/api/connections/"):
                name = urllib.parse.unquote(path[len("/api/connections/"):])
                if not self.store.delete_connection(name):
                    raise ApiError(404, "Unknown connection: %s" % name)
                self._send_json({"ok": True})
            elif path.startswith("/api/profiles/"):
                name = urllib.parse.unquote(path[len("/api/profiles/"):])
                if not self.store.delete_profile(name):
                    raise ApiError(404, "Unknown profile: %s" % name)
                self._send_json({"ok": True})
            else:
                self._error(404, "Unknown API endpoint: %s" % path)
        except ApiError as exc:
            self._error(exc.status, exc.message)
        except Exception as exc:  # pragma: no cover
            import traceback

            traceback.print_exc()
            self._error(500, "Internal error: %s" % exc)


def _make_handler(store, static_dir, verbose=False, port=0):
    return type("Handler", (SlideKickHandler,), {
        "store": store,
        "static_dir": str(static_dir),
        "verbose": verbose,
        "serving_port": port,
    })


def _tls_wrap(server, tls):
    """tls = (certfile, keyfile) or None. Uses TLS 1.2+ server defaults."""
    if not tls:
        return
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.minimum_version = ssl.TLSVersion.TLSv1_2
    ctx.load_cert_chain(tls[0], tls[1])
    server.socket = ctx.wrap_socket(server.socket, server_side=True)


def build_server(host, port, store, static_dir, verbose=False, tls=None):
    server = ThreadingHTTPServer((host, port), _make_handler(store, static_dir, verbose, port))
    _tls_wrap(server, tls)
    return server


def build_dual_servers(port, store, static_dir, verbose=False, tls=None):
    """Both loopback stacks on one port: IPv4 (127.0.0.1) + IPv6 (::1).

    macOS resolves 'localhost' to ::1 first, and WKWebView treats the immediate
    ECONNREFUSED as final without falling back to IPv4 ('Load failed') - a
    server listening on 127.0.0.1 only therefore breaks the pane. Listening on
    both loopback stacks keeps 'localhost', 127.0.0.1 and [::1] working.
    Falls back to IPv4-only on hosts without IPv6. Loopback addresses only."""
    import socket as socket_mod

    servers = []
    handler = _make_handler(store, static_dir, verbose, port)
    try:
        class V6Server(ThreadingHTTPServer):
            address_family = socket_mod.AF_INET6

        v6 = V6Server(("::1", port), handler)
        try:
            v6.socket.setsockopt(socket_mod.IPPROTO_IPV6, socket_mod.IPV6_V6ONLY, 1)
        except OSError:
            pass
        _tls_wrap(v6, tls)
        servers.append(v6)
        port = v6.server_address[1]  # learn the bound (possibly ephemeral) port
    except OSError:
        pass
    try:
        v4 = ThreadingHTTPServer(("127.0.0.1", port), handler)
        _tls_wrap(v4, tls)
        servers.append(v4)
    except OSError:
        pass
    return servers


def find_free_port(start, tries=20):
    for port in range(start, start + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return None


def main(argv=None):
    parser = argparse.ArgumentParser(description=APP + " local server")
    parser.add_argument("--host", default="",
                        help="bind a single explicit host instead of both loopbacks (127.0.0.1 + ::1)")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument(
        "--config", default=str(Path.home() / ".slidekick" / "config.json")
    )
    parser.add_argument(
        "--static-dir",
        default=str(Path(__file__).resolve().parent / "static"),
    )
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument(
        "--debug", action="store_true",
        help="force debug logging on (AI interactions, provider payloads)",
    )
    parser.add_argument("--cert", default="", help="TLS certificate (PEM) - enables https")
    parser.add_argument("--key", default="", help="TLS private key (PEM)")
    args = parser.parse_args(argv)

    store = Store(args.config)
    store.load()
    store.force_debug = args.debug
    tls = (args.cert, args.key) if (args.cert and args.key) else None

    if args.host:
        try:
            servers = [build_server(args.host, args.port, store, args.static_dir,
                                    args.verbose, tls)]
        except OSError as exc:
            print("Cannot bind %s:%s (%s)" % (args.host, args.port, exc), file=sys.stderr)
            alt = find_free_port(args.port)
            if alt:
                print("Port %s busy; suggest --port %s" % (args.port, alt), file=sys.stderr)
            sys.exit(2)
    else:
        servers = build_dual_servers(args.port, store, args.static_dir, args.verbose, tls)
        if not servers:
            print("Cannot bind port %s on any loopback stack" % args.port, file=sys.stderr)
            alt = find_free_port(args.port)
            if alt:
                print("Port %s busy; suggest --port %s" % (args.port, alt), file=sys.stderr)
            sys.exit(2)

    port = servers[0].server_address[1]
    scheme = "https" if tls else "http"
    for s in servers:
        host = s.server_address[0]
        shown = "[%s]" % host if ":" in str(host) else str(host)
        print(
            "%s server v%s  |  %s://%s:%d/taskpane.html  |  config: %s"
            % (APP, VERSION, scheme, shown, port, store.path),
            flush=True,
        )
    print(
        "Simulator: %s://127.0.0.1:%d/simulator.html  |  Ctrl+C to stop" % (scheme, port),
        flush=True,
    )
    if store.debug_enabled():
        print(
            "DEBUG logging ON -> %s/logs/debug-<YYYY-MM-DD>.log" % store.path.parent,
            flush=True,
        )
    # serve the first stack in the main thread, the others as daemons
    for s in servers[1:]:
        threading.Thread(target=s.serve_forever, daemon=True).start()
    try:
        servers[0].serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        for s in servers:
            s.server_close()


if __name__ == "__main__":
    main()