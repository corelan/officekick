#!/usr/bin/env python3
"""
OfficeKick - local companion server for the Corelan Office add-ins
(SlideKick for PowerPoint, DocKick for Word).

Responsibilities:
  * Serve the task-pane UI (static files, no-cache).
  * Store platform connections, profiles and the per-action approval policy
    in a local config file (API keys never leave this machine's config).
  * Proxy AI chat + model-list requests to the configured provider
    (openai / anthropic / openrouter / generic-openai / ollama / llamacpp).

Runs on 127.0.0.1 only. Python 3.8+ standard library, no third-party deps.
"""

import argparse
import base64
import errno
import json
import os
import socket
import ssl
import sys
import threading
import time
import re
import urllib.error
import urllib.parse
import urllib.request
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

APP = "OfficeKick"
VERSION = "0.9.32"

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
    "create_table": "auto",
    "move_shape": "auto",
    "insert_paragraphs": "auto",
    "edit_paragraph": "auto",
    "insert_comment": "auto",
    "delete_paragraph": "confirm",
    "rewrite_doc": "confirm",
    "insert_image": "confirm",
    "apply_layout": "confirm",
    "draft_reply": "auto",
    "forward": "auto",
    "new_message": "confirm",
    "compose_body": "confirm",
    "set_categories": "auto",
    "save_draft": "auto",
    "delete_slide": "confirm",
    "create_image": "confirm",
    "format_text": "confirm",
    "format_selection": "confirm",
    "set_background": "confirm",
    "add_hyperlink": "auto",
    "set_alt_text": "auto",
    "move_slide": "confirm",
    "create_template": "confirm",
    "create_shape": "auto",
    "accept_changes": "confirm",
    "set_track_changes": "confirm",
    "add_hyperlink": "confirm",
    "format_selection": "confirm",
    "style_shape": "confirm",
    "new_meeting": "auto",
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
IMAGE_TIMEOUT = 300   # image generation is slow (60-180s)


class ApiError(Exception):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status
        self.message = message


# --------------------------------------------------------------------------- #
# Config store
# --------------------------------------------------------------------------- #
# --------------------------------------------------------------------------- #
# Conversation sessions (local files, resumable from any pane)
# --------------------------------------------------------------------------- #

SESSION_ID_RE = re.compile(r"^[a-f0-9-]{8,64}$")


class SessionStore:
    def __init__(self, base_dir):
        self.dir = Path(base_dir) / "sessions"
        try:
            self.dir.mkdir(mode=0o700, parents=True, exist_ok=True)
            os.chmod(self.dir, 0o700)
        except OSError:
            pass
        for f in self.dir.glob("*.json"):
            try:
                os.chmod(f, 0o600)
            except OSError:
                pass

    def _valid(self, sid):
        return bool(SESSION_ID_RE.match(sid or ""))

    def list(self, app=None):
        items = []
        for f in self.dir.glob("*.json"):
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
                doc_app = data.get("app")
                # sessions are scoped per application; rows without an app are
                # legacy/shared and stay visible in every pane
                if app and doc_app and doc_app != app:
                    continue
                items.append({
                    "id": data.get("id", f.stem),
                    "name": data.get("name") or "session",
                    "updated": data.get("updated") or f.stat().st_mtime,
                    "messages": len(data.get("exchanges") or []),
                    "app": doc_app,
                })
            except Exception:
                continue
        items.sort(key=lambda x: x["updated"], reverse=True)
        return items

    def get(self, sid):
        if not self._valid(sid):
            raise ApiError(400, "Invalid session id")
        f = self.dir / (sid + ".json")
        if not f.is_file():
            raise ApiError(404, "Unknown session: %s" % sid)
        try:
            return json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            raise ApiError(500, "Session file unreadable")

    def create(self, name, exchanges, app=None):
        sid = uuid.uuid4().hex
        return self.upsert(sid, name, exchanges, app=app)

    def upsert(self, sid, name, exchanges, existing=None, app=None):
        if not self._valid(sid):
            raise ApiError(400, "Invalid session id")
        # the stored file is the source of truth: if the caller does not
        # pass the current exchanges (or passes a shorter history), keep
        # the longer stored list instead of shrinking the session
        f = self.dir / (sid + ".json")
        base = existing
        if base is None and f.is_file():
            try:
                base = json.loads(f.read_text(encoding="utf-8"))
            except Exception:
                base = None
        stored_ex = (base or {}).get("exchanges") or []
        new_ex = self._clean_exchanges(exchanges)
        ex = new_ex if len(new_ex) >= len(stored_ex) else stored_ex
        doc = {
            "id": sid,
            "name": (name or (base or {}).get("name") or "session")[:80],
            "updated": time.time(),
            "exchanges": ex,
            "app": app if app in ("slidekick", "wordkick", "mailkick")
                   else (base or {}).get("app"),
        }
        f = self.dir / (sid + ".json")
        f.write_text(json.dumps(doc, indent=1), encoding="utf-8")
        os.chmod(f, 0o600)
        return {"id": sid, "name": doc["name"], "updated": doc["updated"], "messages": len(ex)}

    def rename(self, sid, name):
        doc = self.get(sid)
        # title-repair: the model may answer with reasoning fragments, fences,
        # labels ("Title: x") or long essays - derive a clean short title
        import re as _re
        n = str(name or "")
        n = _re.sub(r"<think>[\s\S]*?</think>", "", n, flags=_re.I)
        n = _re.sub(r"^\s*```.*?```\s*$", "", n, flags=_re.S)
        n = _re.sub(r"^\s*(title|session title|titel)\s*[:\-]\s*", "", n, flags=_re.I)
        n = n.replace('"', "").replace("`", "").strip()
        nl = [ln for ln in n.splitlines() if ln.strip()]
        n = nl[0].strip() if nl else ""
        words = n.split()
        if len(words) > 7:
            n = " ".join(words[:7])
        if not n:
            n = ""   # empty = keep the stored/derived name
        return self.upsert(sid, n[:80], doc.get("exchanges"), doc)

    def auto_clean(self, days, now=None):
        """Delete sessions whose NEWEST exchange timestamp is older than `days`
        days. Timestamps in exchanges are ms-epoch strings from the panes;
        fall back to the session's `updated` float, then the file mtime.
        Returns the ids of the removed sessions."""
        now = now if now is not None else time.time()
        cutoff = now - max(1, int(days)) * 86400.0
        removed = []
        for f in self.dir.glob("*.json"):
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
            except Exception:
                continue
            ts_list = []
            for ex in data.get("exchanges") or []:
                try:
                    if ex and ex.get("ts"):
                        ts_list.append(float(ex["ts"]))
                except (TypeError, ValueError):
                    continue
            norm = [t / 1000.0 if t > 1e12 else t for t in ts_list]
            newest = max(norm) if norm else None
            if newest is None:
                newest = data.get("updated") or f.stat().st_mtime
            if newest < cutoff:
                try:
                    f.unlink()
                    removed.append(data.get("id") or f.stem)
                except OSError:
                    pass
        return {"count": len(removed), "removed": removed}

    def delete(self, sid):
        if not self._valid(sid):
            raise ApiError(400, "Invalid session id")
        f = self.dir / (sid + ".json")
        if not f.is_file():
            raise ApiError(404, "Unknown session: %s" % sid)
        f.unlink()
        return {"ok": True}

    @staticmethod
    def _clean_exchanges(exchanges):
        """Keep role/content only, cap volume (last 40 messages)."""
        out = []
        for m in exchanges or []:
            if not isinstance(m, dict):
                continue
            role = m.get("role")
            content = m.get("content")
            if role in ("user", "assistant") and isinstance(content, str):
                entry = {"role": role, "content": content[:8000]}
                ts = str(m.get("ts") or "").strip()
                if ts:
                    entry["ts"] = ts[:32]   # when the exchange happened
                out.append(entry)
        return out[-40:]


# --------------------------------------------------------------------------- #
# Named templates (MailKick): small reusable content blocks, stored locally
# --------------------------------------------------------------------------- #

class TemplateStore:
    def __init__(self, base_dir):
        self.dir = Path(base_dir) / "templates"
        try:
            self.dir.mkdir(mode=0o700, parents=True, exist_ok=True)
            os.chmod(self.dir, 0o700)
        except OSError:
            pass
        for f in self.dir.glob("*.json"):
            try:
                os.chmod(f, 0o600)
            except OSError:
                pass

    def list(self):
        items = []
        for f in self.dir.glob("*.json"):
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
                items.append({
                    "id": data.get("id", f.stem),
                    "name": data.get("name") or "template",
                    "updated": data.get("updated") or f.stat().st_mtime,
                    "len": len(data.get("body") or ""),
                })
            except Exception:
                continue
        items.sort(key=lambda x: x["name"].lower())
        return items

    def get(self, tid):
        if not tid or len(tid) > 64 or not re.match(r"^[a-f0-9-]{1,64}$", tid):
            raise ApiError(400, "Invalid template id")
        f = self.dir / (tid + ".json")
        if not f.is_file():
            raise ApiError(404, "Unknown template: %s" % tid)
        try:
            return json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            raise ApiError(500, "Template file unreadable")

    def _write(self, tid, name, body, updated=None):
        doc = {"id": tid, "name": name[:64], "body": body[:20000], "updated": updated or time.time()}
        f = self.dir / (tid + ".json")
        f.write_text(json.dumps(doc, indent=1), encoding="utf-8")
        os.chmod(f, 0o600)
        return {"id": tid, "name": doc["name"], "updated": doc["updated"], "len": len(doc["body"])}

    def create(self, name, body):
        name = str(name or "").strip()
        if not name or len(name) > 64:
            raise ApiError(400, "Template name must be 1-64 chars")
        if not isinstance(body, (str,)) or not body.strip():
            raise ApiError(400, "Template body is empty")
        return self._write(uuid.uuid4().hex, name, body.strip()[:20000])

    def update(self, tid, name=None, body=None):
        doc = self.get(tid)
        nn = str(name or doc.get("name") or "").strip()[:64]
        if not nn:
            raise ApiError(400, "Template name cannot be empty")
        nb = (body if isinstance(body, str) and body.strip() else doc.get("body") or "")
        return self._write(tid, nn, nb.strip()[:20000], doc.get("updated"))

    def delete(self, tid):
        self.get(tid)
        (self.dir / (tid + ".json")).unlink()
        return {"ok": True}


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
        self.sessions = SessionStore(self.path.parent)
        self.templates = TemplateStore(self.path.parent)

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
        self.data.setdefault("autoUpdateCheck", True)   # check GitHub at boot (no git)
        self.data.setdefault("autoApplyApproved", False)  # legacy global auto-apply (fallback for autoApplyByApp)
        self.data.setdefault("autoApplyByApp", {})
        self.data.setdefault("trackedChanges", False)  # DocKick revision setting  # per-add-in auto-apply: slidekick/wordkick/mailkick
        self.data.setdefault("autoClean", True)        # delete old sessions on pane open
        self.data.setdefault("autoCleanDays", 120)     # session older than N days -> removed
        self.data.setdefault("connections", [])
        self.data.setdefault("profiles", [])
        self.data.setdefault("defaultProfile", None)
        approvals = self.data.setdefault("approvals", dict(DEFAULT_APPROVALS))
        for op in OP_TYPES:
            approvals.setdefault(op, DEFAULT_APPROVALS[op])
        self.data.setdefault("debug", False)
        # speaker-notes generation is opt-in (off by default)
        self.data.setdefault("generateNotes", False)
        # AI edits as tracked changes (Word; opt-in)
        self.data.setdefault("trackedChanges", True)
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

    def upsert_profile(self, name, content, other, image=None):
        prof = {
            "name": name,
            "content": {"connection": content[0], "model": content[1]},
            "other": {"connection": other[0], "model": other[1]},
        }
        if image:
            prof["image"] = {"connection": image[0], "model": image[1]}
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
            for slot in ("content", "other", "image"):
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
        for slot in ("content", "other", "image"):
            if slot not in updates and slot != "image" and slot not in prof:
                continue
            ref = updates.get(slot)
            if ref is None or ref == {}:
                if slot == "image" and ref == {}:
                    prof.pop("image", None)   # explicit empty = clear the image model
                continue
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

    _marker_path_cached = None

    def auto_clean_tick(self, min_interval=3600.0):
        """Run the session auto-cleanup at most once per hour (pane opens
        hammer /api/config). Honors the autoClean setting."""
        if not bool(self.data.get("autoClean", True)):
            return
        marker = self.path.parent / ".lastautoclean"
        try:
            if marker.exists() and time.time() - marker.stat().st_mtime < min_interval:
                return
            marker.write_text(str(time.time()), encoding="utf-8")
            os.chmod(marker, 0o600)
            self.sessions.auto_clean(int(self.data.get("autoCleanDays", 120)))
        except OSError:
            pass

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
            "generateNotes": bool(d.get("generateNotes", False)),
            "trackedChanges": bool(d.get("trackedChanges", False)),
            "autoClean": bool(d.get("autoClean", True)),
            "autoCleanDays": int(d.get("autoCleanDays", 120)),
            "autoUpdateCheck": bool(d.get("autoUpdateCheck", True)),
            "autoApplyApproved": bool(d.get("autoApplyApproved", False)),
            "autoApplyByApp": {
                app: bool((d.get("autoApplyByApp") or {}).get(app, d.get("autoApplyApproved", False)))
                for app in ("slidekick", "wordkick", "mailkick")
            },
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
# Image generation (providers with an OpenAI-style images API)
# --------------------------------------------------------------------------- #

UNSUPPORTED_IMAGE_TYPES = ("anthropic", "ollama", "llamacpp")


def provider_generate_image(conn, model, prompt, size=None):
    """POST {base}/images/generations; returns base64 image bytes or raises ApiError."""
    if conn["type"] in UNSUPPORTED_IMAGE_TYPES:
        raise ApiError(400, "Provider type '%s' does not support image generation"
                       % (conn["type"],))
    base = (conn.get("baseUrl") or "").rstrip("/")
    if not base:
        raise ApiError(400, "Connection '%s' has no Base URL" % (conn["name"],))
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + (conn.get("apiKey") or ""),
    }
    if conn["type"] == "openrouter":
        headers["HTTP-Referer"] = "https://github.com/corelan/officekick"
        headers["X-Title"] = "OfficeKick"
    payload = {"model": model, "prompt": prompt, "n": 1}
    if size:
        payload["size"] = size
    body = json.dumps(payload).encode("utf-8")

    # /v1 in base -> {base}/images/generations; otherwise guess both spellings.
    # openrouter's dedicated image API lives at {base}/images (response shape is
    # the OpenAI-style {"data":[{"b64_json"...}]}) - try it first there.
    if conn["type"] == "openrouter":
        urls = [base + "/images"]
        if not base.endswith("/v1"):
            urls.append(base + "/v1/images")
        urls.append(base + "/images/generations")
        # preflight: reject text-only models with actionable suggestions
        try:
            murls = [base + "/models"] if base.endswith("/v1") else [base + "/v1/models", base + "/models"]
            raw = []
            for murl in murls:
                try:
                    raw = _parse_openai_models_raw(http_request(murl, headers=headers, timeout=30))
                    if raw:
                        break
                except ApiError:
                    raw = []
            entry = next((m for m in raw if str(m.get("id") or "") == model), None)
            if entry is not None and not _model_is_image_capable(entry):
                sugg = _image_model_suggestion(model, [m["id"] for m in raw if _model_is_image_capable(m)])
                raise ApiError(
                    400,
                    "Model '%s' cannot generate images (OpenRouter lists it as text-only). "
                    "Image-capable alternative(s): %s%s"
                    % (model, ", ".join("'%s'" % s for s in sugg) or "(none on this account)",
                       " - they are the list shown with 'image-capable only'." if sugg else ""))
        except ApiError:
            raise
        except Exception:
            pass  # preflight is best-effort; the provider's own error is the fallback
    elif "/v1" in base:
        urls = [base + "/images/generations"]
    else:
        urls = [base + "/v1/images/generations", base + "/images/generations"]

    print("[image] conn=%s model=%s prompt=%r start" % (conn["name"], model, prompt[:60]), flush=True)
    t0 = time.time()
    last_err = None
    for url in urls:
        for extra in ({"response_format": "b64_json"}, None):
            try:
                req = urllib.request.Request(
                    url, data=body, headers=dict(headers, **(extra or {})),
                    method="POST")
                with urllib.request.urlopen(req, timeout=IMAGE_TIMEOUT) as resp:
                    raw = resp.read()          # read INSIDE the with-block
            except urllib.error.HTTPError as exc:
                detail = ""
                try:
                    detail = exc.read(400).decode("utf-8", "replace")
                except Exception:
                    pass
                last_err = ApiError(exc.code or 502,
                                    "image request failed: HTTP %s %s"
                                    % (exc.code, detail[:200]))
                break                          # HTTP error: next URL, not next variant
            except (urllib.error.URLError, socket.timeout, TimeoutError) as exc:
                last_err = ApiError(504, "image request timed out or failed: %s" % exc)
                break
            try:
                data = json.loads(raw.decode("utf-8", "replace"))
            except Exception:
                last_err = ApiError(502, "image response was not JSON")
                continue
            items = data.get("data") or []
            if not items:
                last_err = ApiError(502, "image response contained no data")
                continue
            if items[0].get("b64_json"):
                print("[image] conn=%s done in %.1fs (%d b64 chars)" % (conn["name"], time.time() - t0, len(items[0]["b64_json"])), flush=True)
                return items[0]["b64_json"]
            if items[0].get("url"):
                try:
                    dl = urllib.request.Request(
                        items[0]["url"],
                        headers={"Authorization": "Bearer " + (conn.get("apiKey") or "")})
                    with urllib.request.urlopen(dl, timeout=IMAGE_TIMEOUT) as r2:
                        return base64.b64encode(r2.read()).decode("ascii")
                except Exception as exc:
                    raise ApiError(502, "could not download image url: %s" % exc)
            last_err = ApiError(502, "image response had neither b64_json nor url")
    print("[image] conn=%s FAILED in %.1fs: %s" % (conn["name"], time.time() - t0, last_err), flush=True)
    raise last_err or ApiError(502, "image generation failed")


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
        reason = getattr(exc, "reason", exc)
        hint = _reachability_hint(reason, url)
        raise ApiError(
            502,
            "Cannot reach %s (%s). Is the service running?%s" % (_host_of(url), reason, hint),
        )
    except (socket.timeout, TimeoutError):
        raise ApiError(502, "Timeout talking to %s (>%ss)" % (_host_of(url), timeout))
    try:
        return json.loads(text)
    except Exception:
        return text


def _reachability_hint(reason, url):
    """macOS local-network privacy: a server process without the macOS
    "Local Network" permission gets Errno 65 (EHOSTUNREACH / "No route to
    host") for LAN targets even though the same request works from an
    interactive shell. Attach a targeted hint for that case."""
    errno_ = getattr(reason, "errno", None)
    if errno_ is None and getattr(reason, "args", None):
        first = reason.args[0]
        if isinstance(first, int):
            errno_ = first
    if errno_ is None or errno_ != errno.EHOSTUNREACH:
        return ""
    try:
        import ipaddress as _ip
        import urllib.parse as _up
        host = _up.urlsplit(url).hostname or ""
        private = _ip.ip_address(host).is_private
    except (ValueError, AttributeError):
        private = False   # hostnames resolve via DNS, not the local-network wall
    if not private:
        return ""
    return (" - macOS local-network check: give the OfficeKick server "
            "'Local Network' access (System Settings > Privacy & Security > "
            "Local Network) or restart the server from a permitted session "
            "(e.g. 'officekick run' in a Terminal window)")


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


IMAGE_MODEL_NAME_HINTS = (
    "image", "imagen", "seedream", "dall", "flux", "banana", "recraft",
    "ideogram", "kolors", "stable-diffusion", "sdxl", "sd3", "playground",
)


def _image_model_suggestion(requested, capable_ids):
    """Closest image-capable ids for a non-image model id (openrouter preflight)."""
    ids = [str(x) for x in capable_ids]
    cand = requested + "-image"
    if cand in ids:
        return [cand]
    low = requested.lower()
    near = sorted(
        {i for i in ids if i.lower().startswith(low) or low.startswith(i.lower())},
        key=len)[:3]
    return near or ids[:5]


def _model_is_image_capable(entry):
    """True for text->image models; prefers provider metadata (openrouter's
    architecture.output_modalities), falls back to id heuristics."""
    if not isinstance(entry, dict):
        return False
    ident = str(entry.get("id") or entry.get("name") or "")
    arch = entry.get("architecture") or {}
    outs = arch.get("output_modalities") or []
    if outs:
        return "image" in [str(x).lower() for x in outs]
    return any(h in ident.lower() for h in IMAGE_MODEL_NAME_HINTS)


def _parse_openai_models_raw(data):
    """OpenAI-style /models entries kept whole (metadata needed for filtering)."""
    if isinstance(data, dict) and isinstance(data.get("data"), list):
        return [m for m in data["data"] if isinstance(m, dict) and m.get("id")]
    if isinstance(data, list):
        return [m for m in data if isinstance(m, dict) and m.get("id")]
    return []


def _parse_openai_models(data):
    return [str(m["id"]) for m in _parse_openai_models_raw(data)]


def _dedupe(seq):
    seen = set()
    out = []
    for item in seq:
        if item and item not in seen:
            seen.add(item)
            out.append(item)
    return out


def filter_image_models(store, conn_name):
    conn = store.find_connection(conn_name)
    if conn is None:
        raise ApiError(404, "Unknown connection: %r" % (conn_name,))
    t, base, headers = conn["type"], normalize_base(conn), conn_headers(conn)
    try:
        if t == "openrouter":
            raw = []
            for murl in (base + "/models", base + "/v1/models"):
                try:
                    raw = _parse_openai_models_raw(
                        http_request(murl, headers=headers, timeout=30))
                    if raw:
                        break
                except ApiError:
                    continue
            out = [m["id"] for m in raw if _model_is_image_capable(m)]
            if out:
                return out
            # metadata missing -> same list but by-name
            ids = _parse_openai_models({"data": raw})
            return [i for i in ids if any(h in i.lower() for h in IMAGE_MODEL_NAME_HINTS)]
        if t == "ollama":
            data = http_request(base + "/api/tags", headers=headers, timeout=30)
            out = [str(m.get("name") or m.get("model") or "")
                   for m in (data.get("models") or []) if isinstance(m, dict)]
        elif t == "anthropic":
            out = []  # anthropic cannot generate images
        else:
            raw = []
            for url in (base + "/v1/models", base + "/models"):
                try:
                    raw = _parse_openai_models_raw(http_request(url, headers=headers, timeout=30))
                    if raw:
                        break
                except ApiError:
                    continue
            out = [str(m["id"]) for m in raw if _model_is_image_capable(m)]
        return [i for i in out if i and any(h in i.lower() for h in IMAGE_MODEL_NAME_HINTS)]
    except ApiError:
        raise
    except Exception as exc:
        raise ApiError(502, "could not list models: %s" % exc)



UPDATE_CACHE = {"t": 0.0, "data": None}


def test_connection(store, name):
    """Reach the provider: cheap list-models call, auth/base errors surface."""
    if not name:
        return {"ok": False, "error": "which connection? (unknown connection name)"}
    t0 = time.time()
    try:
        resp = provider_list_models(store, name)
        models = resp.get("models") if isinstance(resp, dict) else resp
        n = len(models) if isinstance(models, (list, tuple)) else 0
        return {"ok": True, "connection": name, "models": n,
                "ms": int((time.time() - t0) * 1000)}
    except Exception as exc:
        return {"ok": False, "connection": name,
                "error": "%s" % exc, "ms": int((time.time() - t0) * 1000)}


def _version_tuple(v):
    parts = re.findall(r"\d+", str(v or ""))
    return tuple(int(x) for x in parts) or (0,)


def update_check(force_latest=None):
    """Compare the running VERSION with the latest in the GitHub repo.

    Reads the VERSION file from GitHub (contents API, raw fallback; no git
    involved). Cached for 10 minutes; ?latest= overrides the remote value
    (used by tests). A private repo needs a read-only token in the env var
    OFFICEKICK_GITHUB_TOKEN or the file ~/.officekick/github_token.
    """
    now = time.time()
    if not force_latest:
        cached = UPDATE_CACHE["data"]
        if cached and now - UPDATE_CACHE["t"] < 600:
            return cached
    latest = force_latest or None
    if latest is None:
        headers = {"User-Agent": "officekick/" + VERSION,
                   "Accept": "application/vnd.github.raw"}
        raw_url = "https://raw.githubusercontent.com/corelan/officekick/main/VERSION"
        try:
            req = urllib.request.Request(raw_url, headers=headers)
            with urllib.request.urlopen(req, timeout=8) as r:
                latest = r.read().decode("utf-8", "replace").strip()
        except Exception as exc:
            return {"ok": False, "error": "could not check for updates: %s" % exc,
                    "current": VERSION}
    latest = (latest or "").strip()
    if not latest:
        return {"ok": False, "error": "no version published in the repo", "current": VERSION}
    up_to_date = _version_tuple(latest) <= _version_tuple(VERSION)
    data = {"ok": True, "current": VERSION, "latest": latest,
            "upToDate": up_to_date,
            "repo": "https://github.com/corelan/officekick",
            "howToUpdate": "To update: run python3 install.py --update in your OfficeKick "
                           "clone (it pulls with git, redeploys the add-ins and restarts the "
                           "local server), and reload the task pane (reopen the add-in; restart Office if it "
                           "keeps showing the old UI). No automatic git pull is attempted."}
    if not force_latest:
        UPDATE_CACHE["t"] = now
        UPDATE_CACHE["data"] = data
    return data


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
        if t in ("generic-openai", "openrouter") and not base.endswith("/v1"):
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
        data, last_error, text = None, None, ""
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
    for slot in ("content", "other", "image"):
        ref = body.get(slot)
        if ref is None or ref == {}:
            if slot == "image":
                continue  # image model is optional
            raise ApiError(400, "Profile %s slot missing" % slot)
        conn, model = ref.get("connection"), ref.get("model")
        if not store.find_connection(conn):
            raise ApiError(
                400, "Profile %s references unknown connection: %r" % (slot, conn)
            )
        if not isinstance(model, str) or not model.strip():
            raise ApiError(400, "Profile %s has no model selected" % slot)
        refs.append((slot, conn, model.strip()))
    by_slot = {slot: (conn, model) for slot, conn, model in refs}
    return (
        name.strip(),
        by_slot["content"],
        by_slot["other"],
        by_slot.get("image"),
    )


# --------------------------------------------------------------------------- #
# HTTP server
# --------------------------------------------------------------------------- #
_ASSET_RE = re.compile(
    r'(src|href)="((?!(?:https?:)?//|#)[^"?]+?\.(?:js|css))"'
)


def _decorate_html(html_bytes):
    """Append ?v=<VERSION> to relative js/css asset refs in served HTML.

    Without this, hosts (and the simulator) can execute a heuristically
    cached OLD build of the pane scripts after an update - asset URLs that
    carry the version never collide with stale cache entries.
    """
    if not html_bytes:
        return html_bytes
    try:
        html = html_bytes.decode("utf-8")
    except UnicodeDecodeError:
        return html_bytes
    html = _ASSET_RE.sub(lambda m: '%s="%s?v=%s"' % (m.group(1), m.group(2), VERSION), html)
    tag = '<script>window.__SK_BUILD = "%s";</script>' % VERSION
    if "<head" in html:
        html = html.replace("<head>", "<head>" + tag, 1) if html.count("<head>") else html
    if "window.__SK_BUILD" not in html:
        # head variant with attributes (or no head at all): prepend right after <html>
        html = html.replace("<html", "<html", 1)
        if "<body" in html:
            html = html.replace("<body", tag + "<body", 1)
    return html.encode("utf-8")


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
        # always-on access log: every request line + User-Agent, so pane-load
        # problems at Office launch are visible after the fact. Paths/UA only,
        # never request bodies (those go to the opt-in debug log).
        line = fmt % args
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        ua = self.headers.get("User-Agent", "-") if hasattr(self, "headers") else "-"
        try:
            logdir = Path.home() / ".officekick" / "logs"
            logdir.mkdir(mode=0o700, parents=True, exist_ok=True)
            with open(logdir / ("access-%s.log" % time.strftime("%Y-%m-%d")),
                      "a", encoding="utf-8") as fh:
                fh.write("%s | %s | ua=%s\n" % (stamp, line, ua))
        except OSError:
            pass
        if self.verbose:
            sys.stderr.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), line))

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
        # legacy name: the MailKick pane lived under outlook/ before 0.8.7;
        # older sideloaded manifests keep working without a reinstall
        if rel.startswith("outlook/"):
            rel = "mailkick/" + rel[len("outlook/"):]
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
            src = (target.parent / "taskpane.html").read_text(encoding="utf-8")
            src = src.replace(
                '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>',
                '<script src="sim.js"></script>',
            )
            src = src.replace("<title>SlideKick</title>",
                              "<title>SlideKick - simulator</title>")
            data = src.encode("utf-8")
        else:
            data = target.read_bytes()
        if suffix == ".html":
            data = _decorate_html(data)
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
                self.store.auto_clean_tick()
                pub = self.store.public()
                pub["port"] = self.serving_port
                pub["connectionTypes"] = list(CONNECTION_TYPES)
                pub["defaultBases"] = DEFAULT_BASES
                pub["approvalModes"] = list(APPROVAL_MODES)
                pub["opTypes"] = list(OP_TYPES)
                self._send_json(pub)
            elif path == "/api/models":
                conn = (query.get("connection") or [""])[0]
                caps = (query.get("capability") or [""])[0]
                resp = provider_list_models(self.store, conn)
                if caps == "image":
                    resp = dict(resp)
                    resp["models"] = filter_image_models(self.store, conn)
                self._send_json(resp)
            elif path == "/api/sessions":
                self._send_json({"sessions": self.store.sessions.list(query.get("app", [None])[0])})
            elif path.startswith("/api/sessions/"):
                sid = path[len("/api/sessions/"):]
                self._send_json(self.store.sessions.get(sid))
            elif path == "/api/test-connection":
                name = (query.get("connection") or [""])[0]
                result = test_connection(self.store, name)
                self._send_json(result)
            elif path == "/api/update-check":
                self._send_json(update_check((query.get("latest") or [""])[0] or None))
            elif path == "/api/templates":
                self._send_json({"templates": self.store.templates.list()})
            elif path.startswith("/api/templates/"):
                tid = path[len("/api/templates/"):]
                self._send_json(self.store.templates.get(tid))
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
                name, content, other, image = validate_profile_body(self.store, body)
                self.store.upsert_profile(name, content, other, image)
                self._send_json({"ok": True, "name": name})
            elif path == "/api/connections/rename":
                self.store.rename_connection(
                    str(body.get("from") or ""), str(body.get("to") or ""), body)
                self._send_json({"ok": True, "name": body.get("to")})
            elif path == "/api/profiles/rename":
                self.store.rename_profile(
                    str(body.get("from") or ""), str(body.get("to") or ""), body)
                self._send_json({"ok": True, "name": body.get("to")})
            elif path == "/api/sessions":
                doc = self.store.sessions.create(body.get("name") or "session",
                                                 body.get("exchanges") or [],
                                                 body.get("app"))
                self._send_json({"ok": True, "session": doc})
            elif path == "/api/templates":
                doc = self.store.templates.create(body.get("name"), body.get("body"))
                self._send_json({"ok": True, "template": doc})
            elif path.startswith("/api/sessions/rename/"):
                sid = path[len("/api/sessions/rename/"):]
                doc = self.store.sessions.rename(sid, body.get("name"))
                self._send_json({"ok": True, "session": doc})
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
            elif path == "/api/image":
                conn_name = body.get("connection")
                conn = self.store.find_connection(conn_name or "")
                if not conn:
                    raise ApiError(404, "Unknown connection: %r" % (conn_name,))
                model = body.get("model") or ""
                prompt = body.get("prompt") or ""
                if not model or not prompt:
                    raise ApiError(400, "image request needs 'model' and 'prompt'")
                size = body.get("size") or None
                b64 = provider_generate_image(conn, model, prompt, size)
                self._send_json({"ok": True, "b64": b64, "model": model})
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
                if "generateNotes" in body and body["generateNotes"] is not None:
                    self.store.data["generateNotes"] = bool(body["generateNotes"])
                    changed = True
                if "autoApplyApproved" in body and body["autoApplyApproved"] is not None:
                    self.store.data["autoApplyApproved"] = bool(body["autoApplyApproved"])
                    changed = True
                if "autoApplyByApp" in body and body["autoApplyByApp"] is not None:
                    for _app, _on in dict(body["autoApplyByApp"]).items():
                        if _app in ("slidekick", "wordkick", "mailkick"):
                            byapp = self.store.data.setdefault("autoApplyByApp", {})
                            byapp[_app] = bool(_on)
                            changed = True
                if "autoUpdateCheck" in body and body["autoUpdateCheck"] is not None:
                    self.store.data["autoUpdateCheck"] = bool(body["autoUpdateCheck"])
                    changed = True
                if "trackedChanges" in body and body["trackedChanges"] is not None:
                    self.store.data["trackedChanges"] = bool(body["trackedChanges"])
                    changed = True
                if "autoClean" in body and body["autoClean"] is not None:
                    self.store.data["autoClean"] = bool(body["autoClean"])
                    changed = True
                if "autoCleanDays" in body and body["autoCleanDays"] is not None:
                    try:
                        _acd = int(body["autoCleanDays"])
                    except (TypeError, ValueError):
                        raise ApiError(400, "autoCleanDays must be a number (days)")
                    if _acd < 1 or _acd > 3650:
                        raise ApiError(400, "autoCleanDays out of range (1-3650)")
                    self.store.data["autoCleanDays"] = _acd
                    changed = True
                if changed:
                    self.store.save()
                self._send_json({"ok": True, "approvals": self.store.data["approvals"]})
            elif path.startswith("/api/sessions/"):
                sid = urllib.parse.unquote(path[len("/api/sessions/"):])
                try:
                    existing = self.store.sessions.get(sid)
                except ApiError:
                    existing = None
                new_name = body.get("name")
                if new_name is None and existing:
                    new_name = existing.get("name")   # keep the stored title
                doc = self.store.sessions.upsert(
                    sid, new_name, body.get("exchanges") or [],
                    existing=existing, app=body.get("app"))
                self._send_json({"ok": True, "session": doc})
            elif path.startswith("/api/templates/"):
                tid = urllib.parse.unquote(path[len("/api/templates/"):])
                doc = self.store.templates.update(tid, body.get("name"), body.get("body"))
                self._send_json({"ok": True, "template": doc})
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
            elif path.startswith("/api/sessions/"):
                sid = urllib.parse.unquote(path[len("/api/sessions/"):])
                self._send_json(self.store.sessions.delete(sid))
            elif path.startswith("/api/templates/"):
                tid = urllib.parse.unquote(path[len("/api/templates/"):])
                self._send_json(self.store.templates.delete(tid))
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
    # adopt a legacy SlideKick-era config when the default path is empty
    _legacy_cfg = Path.home() / ".slidekick" / "config.json"
    _default_cfg = Path.home() / ".officekick" / "config.json"
    try:
        if not _default_cfg.exists() and _legacy_cfg.exists():
            _default_cfg.parent.mkdir(parents=True, exist_ok=True)
            import shutil
            shutil.copy2(_legacy_cfg, _default_cfg)
            os.chmod(_default_cfg, 0o600)
            print("migrated legacy config:", _legacy_cfg, "->", _default_cfg)
    except OSError:
        pass

    parser.add_argument("--host", default="",
                        help="bind a single explicit host instead of both loopbacks (127.0.0.1 + ::1)")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument(
        "--config", default=str(Path.home() / ".officekick" / "config.json")
    )
    parser.add_argument(
        "--static-dir",
        default=str(Path(__file__).resolve().parent / "slidekick" / "static"),
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
            "%s server v%s  |  %s://%s:%d/slidekick/static/taskpane.html  |  config: %s"
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