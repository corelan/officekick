#!/usr/bin/env python3
"""
Mock AI provider for developing and testing SlideKick.

Serves one of the three wire formats SlideKick speaks:
  --flavor openai     GET /v1/models, POST /v1/chat/completions
  --flavor anthropic  GET /v1/models, POST /v1/messages
  --flavor ollama     GET /api/tags,   POST /api/chat

Options:
  --require-key KEY   reject requests missing the expected credential header
  --reply-file FILE   reply with the file's text (e.g. a canned slide_ops JSON)
  --echo-request      reply with a JSON dump of the parsed request payload
                      (lets tests assert exactly what the adapter sent)
  --fail-models       GET model-list endpoints return HTTP 500 (fallback tests)

Standard library only.
"""

import argparse
import base64
import json
import struct
import sys
import threading
import time
import zlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

DEFAULT_MODELS = ["mock-mini", "mock-large", "mock-vision"]
DEFAULT_REPLY = json.dumps({
    "kind": "slide_ops",
    "ops": [
        {"type": "create_slide", "layout": "title_content",
         "title": "Mock slide", "body": ["bullet one", "bullet two"], "notes": None},
        {"type": "delete_slide", "slide": 0},
    ],
})



def make_png_1x1(color=(255, 0, 0)):
    """A tiny valid PNG (mock image payloads)."""
    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xffffffff)
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    raw = zlib.compress(bytes([0, color[0], color[1], color[2]]))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", raw) + chunk(b"IEND", b""))

TINY_PNG = base64.b64encode(make_png_1x1()).decode("ascii")

class MockState:
    reply = DEFAULT_REPLY
    echo_request = False
    fail_models = False
    require_key = None
    last_request = None


STATE = MockState()


def make_handler(flavor):
    class MockHandler(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.1"

        def log_message(self, fmt, *args):
            pass

        def _json(self, obj, status=200):
            body = json.dumps(obj).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _authorized(self):
            if not STATE.require_key:
                return True
            if flavor == "anthropic":
                return self.headers.get("x-api-key") == STATE.require_key
            return self.headers.get("Authorization") == "Bearer " + STATE.require_key

        def do_GET(self):
            if STATE.require_key and not self._authorized():
                self._json({"error": {"message": "unauthorized (mock)"}}, 401)
                return
            if STATE.fail_models:
                self._json({"error": {"message": "mock models endpoint broken"}}, 500)
                return
            if flavor in ("openai", "anthropic", "llamacpp") and self.path in ("/v1/models",):
                self._json({"data": [{"id": m} for m in DEFAULT_MODELS]})
            elif flavor == "openai" and self.path == "/models":
                self._json({"data": [{"id": m} for m in DEFAULT_MODELS]})
            elif flavor == "ollama" and self.path == "/api/tags":
                self._json({"models": [{"name": m} for m in DEFAULT_MODELS]})
            else:
                self._json({"error": {"message": "not found: %s" % self.path}}, 404)

        def do_POST(self):
            length = int(self.headers.get("Content-Length") or 0)
            raw = self.rfile.read(length).decode("utf-8", "replace") if length else "{}"
            try:
                req = json.loads(raw)
            except Exception:
                req = {"raw": raw}
            if STATE.require_key and not self._authorized():
                self._json({"error": {"message": "unauthorized (mock)"}}, 401)
                return
            STATE.last_request = req
            if STATE.echo_request:
                reply = json.dumps(req)
            else:
                reply = STATE.reply
            if flavor in ("openai", "llamacpp") and self.path in ("/v1/chat/completions", "/chat/completions"):
                self._json({"choices": [{"message": {"role": "assistant", "content": reply}}]})
            elif flavor == "anthropic" and self.path == "/v1/messages":
                self._json({"content": [{"type": "text", "text": reply}]})
            elif flavor == "ollama" and self.path == "/api/chat":
                self._json({"message": {"role": "assistant", "content": reply}, "done": True})
            elif flavor in ("openai", "generic") and self.path.rstrip("/").endswith("images/generations") or \
                    self.path.rstrip("/").endswith("/v1/images"):
                # a 1x1 red PNG; the prompt is echoed back as metadata
                self._json({"created": 0, "data": [{
                    "b64_json": TINY_PNG,
                    "revised_prompt": str((STATE.last_request or {}).get("prompt", ""))[:120]}]})
            else:
                self._json({"error": {"message": "not found: %s" % self.path}}, 404)

    return MockHandler


def main():
    parser = argparse.ArgumentParser(description="Mock AI provider for SlideKick tests")
    parser.add_argument("--port", type=int, default=8901)
    parser.add_argument("--flavor", default="openai",
                        choices=["openai", "anthropic", "ollama"])
    parser.add_argument("--reply-file")
    parser.add_argument("--echo-request", action="store_true")
    parser.add_argument("--fail-models", action="store_true")
    parser.add_argument("--require-key")
    args = parser.parse_args()

    if args.reply_file:
        STATE.reply = open(args.reply_file, "r", encoding="utf-8").read()
    STATE.echo_request = args.echo_request
    STATE.fail_models = args.fail_models
    STATE.require_key = args.require_key

    server = ThreadingHTTPServer(("127.0.0.1", args.port), make_handler(args.flavor))
    print("mock %s provider on 127.0.0.1:%d" % (args.flavor, args.port), flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()