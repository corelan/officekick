#!/usr/bin/env python3
"""
End-to-end tests for the SlideKick local server.

Runs the real server in-process against an in-process mock AI provider
(tools/mock_provider.py). Standard library only:

    python3 tests/test_slidekick_server.py
"""

import json
import os
import sys
import tempfile
import threading
import unittest
import urllib.request
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tools"))

import slidekick_server as sk  # noqa: E402
import mock_provider as mp  # noqa: E402
from http.server import ThreadingHTTPServer  # noqa: E402

LT, GT, SLASH = chr(60), chr(62), chr(47)


def tag(tagname, body, close=True):
    out = LT + tagname + GT + body
    return out + (LT + SLASH + tagname + GT) if close else out


class ServerHarness:
    def __init__(self):
        self.tmp = tempfile.mkdtemp(prefix="slidekick-test-")
        cfg = os.path.join(self.tmp, "config.json")
        self.store = sk.Store(cfg)
        self.store.load()
        self.server = sk.build_server(
            "127.0.0.1", 0, self.store, str(REPO / "src" / "static")
        )
        self.port = self.server.server_address[1]
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    @property
    def base(self):
        return "http://127.0.0.1:%d" % self.port

    def req(self, path, method="GET", body=None, expect_error=False):
        data = json.dumps(body).encode() if body is not None else None
        r = urllib.request.Request(self.base + path, data=data, method=method)
        if data is not None:
            r.add_header("Content-Type", "application/json")
        try:
            with urllib.request.urlopen(r, timeout=30) as resp:
                return resp.status, json.loads(resp.read().decode() or "{}")
        except urllib.error.HTTPError as exc:
            payload = json.loads(exc.read().decode() or "{}")
            if not expect_error:
                raise AssertionError("unexpected HTTP %s: %r" % (exc.code, payload))
            return exc.code, payload

    def stop(self):
        self.server.shutdown()
        self.server.server_close()


class MockHarness:
    def __init__(self, flavor="openai"):
        self.flavor = flavor
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), mp.make_handler(flavor))
        self.port = self.server.server_address[1]
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    @property
    def url(self):
        return "http://127.0.0.1:%d" % self.port

    def stop(self):
        self.server.shutdown()
        self.server.server_close()


class TestSlideKickServer(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.S = ServerHarness()
        # saved & restored around each test
        cls._mock_state = dict(
            reply=mp.STATE.reply,
            echo_request=mp.STATE.echo_request,
            fail_models=mp.STATE.fail_models,
            require_key=mp.STATE.require_key,
        )
        cls.mocks = []

    @classmethod
    def tearDownClass(cls):
        cls.S.stop()
        for m in cls.mocks:
            m.stop()

    def setUp(self):
        S, self.S = self.S, self.S
        # fresh config for every test
        self.S.store.data = {
            "version": 1, "connections": [], "profiles": [],
            "defaultProfile": None, "approvals": dict(sk.DEFAULT_APPROVALS),
            "debug": False,
        }
        # reset mock state
        mp.STATE.reply = mp.DEFAULT_REPLY
        mp.STATE.echo_request = False
        mp.STATE.fail_models = False
        mp.STATE.require_key = None

    def mock(self, flavor="openai"):
        m = MockHarness(flavor)
        self.mocks.append(m)
        return m

    # ---------------------------------------------------------------- #
    def test_health_and_config(self):
        st, data = self.S.req("/api/health")
        self.assertEqual(st, 200)
        self.assertTrue(data["ok"])
        st, data = self.S.req("/api/config")
        self.assertEqual(st, 200)
        self.assertEqual(data["approvals"]["delete_slide"], "confirm")
        self.assertEqual(data["approvals"]["create_slide"], "auto")
        self.assertIn("openrouter", data["connectionTypes"])

    def test_connection_crud_and_key_masking(self):
        m = self.mock("openai")
        st, _ = self.S.req("/api/connections", "POST", {
            "name": "main", "type": "generic-openai", "baseUrl": m.url,
            "apiKey": "secret-key-123",
        })
        self.assertEqual(st, 200)
        st, cfg = self.S.req("/api/config")
        conn = [c for c in cfg["connections"] if c["name"] == "main"][0]
        self.assertTrue(conn["hasKey"])
        # the API key must never leave the store via the API
        self.assertNotIn("secret-key-123", json.dumps(cfg))
        # re-save without key -> key preserved
        self.S.req("/api/connections", "POST", {
            "name": "main", "type": "generic-openai", "baseUrl": m.url,
        })
        st, _ = self.S.req("/api/models?connection=main", expect_error=False)
        # mock requires no key; but header check that key flowed via test-connection
        st, _ = self.S.req("/api/connections/main", "DELETE")
        self.assertEqual(st, 200)
        st, cfg2 = self.S.req("/api/config")
        self.assertEqual([c for c in cfg2["connections"] if c["name"] == "main"], [])

    def test_bad_connection_inputs(self):
        st, data = self.S.req("/api/connections", "POST", {"name": "x", "type": "nope"}, True)
        self.assertEqual(st, 400)
        st, data = self.S.req("/api/connections", "POST", {"name": "", "type": "openai"}, True)
        self.assertEqual(st, 400)
        st, data = self.S.req("/api/connections", "POST", {
            "name": "g", "type": "generic-openai", "baseUrl": "ftp://x"}, True)
        self.assertEqual(st, 400)

    def test_approvals_settings(self):
        st, _ = self.S.req("/api/settings", "PUT", {
            "approvals": {"delete_slide": "auto", "apply_layout": "deny"}})
        self.assertEqual(st, 200)
        st, cfg = self.S.req("/api/config")
        self.assertEqual(cfg["approvals"]["delete_slide"], "auto")
        self.assertEqual(cfg["approvals"]["apply_layout"], "deny")
        self.assertEqual(cfg["approvals"]["create_slide"], "auto")  # untouched
        st, data = self.S.req("/api/settings", "PUT", {"approvals": {"delete_slide": "bogus"}}, True)
        self.assertEqual(st, 400)
        st, data = self.S.req("/api/settings", "PUT", {"approvals": {"not_an_op": "auto"}}, True)
        self.assertEqual(st, 400)
        # debug toggle
        st, _ = self.S.req("/api/settings", "PUT", {"debug": True})
        st, cfg = self.S.req("/api/config")
        self.assertTrue(cfg["debug"])

    def test_profiles(self):
        self.S.req("/api/connections", "POST", {
            "name": "ollama-local", "type": "ollama",
            "baseUrl": "http://localhost:11434"})
        st, _ = self.S.req("/api/profiles", "POST", {
            "name": "daily",
            "content": {"connection": "ollama-local", "model": "llama3:8b"},
            "other": {"connection": "ollama-local", "model": "llama3:8b"},
        })
        self.assertEqual(st, 200)
        st, cfg = self.S.req("/api/config")
        self.assertEqual(cfg["defaultProfile"], "daily")
        # unknown connection ref
        st, data = self.S.req("/api/profiles", "POST", {
            "name": "bad", "content": {"connection": "ghost", "model": "x"},
            "other": {"connection": "ollama-local", "model": "x"}}, True)
        self.assertEqual(st, 400)
        # default profile update
        st, _ = self.S.req("/api/settings", "PUT", {"defaultProfile": "daily"})
        st, _ = self.S.req("/api/settings", "PUT", {"defaultProfile": "nope"}, True)
        self.assertEqual(st, 400)

    def test_delete_protected_connection(self):
        # a connection referenced by a profile must not be deletable (409)
        self.S.req("/api/connections", "POST", {
            "name": "used-conn", "type": "openai",
            "baseUrl": "https://api.openai.com/v1"})
        self.S.req("/api/profiles", "POST", {
            "name": "guard", "content": {"connection": "used-conn", "model": "gpt"},
            "other": {"connection": "used-conn", "model": "gpt"}})
        st, data = self.S.req("/api/connections/used-conn", "DELETE", None, True)
        self.assertEqual(st, 409)
        self.assertIn("used by profile", data["error"])
        st, cfg = self.S.req("/api/config")
        self.assertTrue(any(c["name"] == "used-conn" for c in cfg["connections"]))
        # after removing the profile, deletion succeeds
        self.S.req("/api/profiles/guard", "DELETE")
        st, _ = self.S.req("/api/connections/used-conn", "DELETE")
        self.assertEqual(st, 200)

    def test_models_openai_compatible_with_key(self):
        m = self.mock("openai")
        mp.STATE.require_key = "test-key"
        self.S.req("/api/connections", "POST", {
            "name": "keyed", "type": "generic-openai", "baseUrl": m.url,
            "apiKey": "test-key"})
        st, data = self.S.req("/api/models?connection=keyed")
        self.assertEqual(st, 200)
        self.assertEqual(data["models"], ["mock-mini", "mock-large", "mock-vision"])
        # wrong key must surface a provider error
        self.S.req("/api/connections", "POST", {
            "name": "wrong", "type": "generic-openai", "baseUrl": m.url,
            "apiKey": "bad"})
        st, data = self.S.req("/api/models?connection=wrong", expect_error=True)
        self.assertEqual(st, 502)
        self.assertIn("401", data["error"])

    def test_models_v1_fallback(self):
        m = self.mock("openai")
        # base WITHOUT /v1: {base}/models is served by the mock too, but the
        # llamacpp flavor only exposes /v1/models -> use it
        m2 = self.mock("llamacpp")
        self.S.req("/api/connections", "POST", {
            "name": "llama", "type": "llamacpp", "baseUrl": m2.url})
        st, data = self.S.req("/api/models?connection=llama")
        self.assertEqual(data["models"], ["mock-mini", "mock-large", "mock-vision"])

    def test_models_anthropic_fallback_list(self):
        m = self.mock("anthropic")
        mp.STATE.fail_models = True
        self.S.req("/api/connections", "POST", {
            "name": "ant", "type": "anthropic", "baseUrl": m.url})
        st, data = self.S.req("/api/models?connection=ant")
        self.assertEqual(data["source"], "fallback")
        self.assertEqual(len(data["models"]), len(sk.ANTHROPIC_FALLBACK_MODELS))

    def test_models_ollama(self):
        m = self.mock("ollama")
        self.S.req("/api/connections", "POST", {
            "name": "ol", "type": "ollama", "baseUrl": m.url})
        st, data = self.S.req("/api/models?connection=ol")
        self.assertEqual(data["models"], ["mock-mini", "mock-large", "mock-vision"])

    def test_chat_generic_openai_returns_ops(self):
        m = self.mock("openai")
        self.S.req("/api/connections", "POST", {
            "name": "g", "type": "generic-openai", "baseUrl": m.url})
        st, data = self.S.req("/api/chat", "POST", {
            "connection": "g", "model": "mock-mini",
            "messages": [{"role": "system", "content": "sys"}, {"role": "user", "content": "hi"}],
        })
        self.assertEqual(st, 200)
        self.assertEqual(data["text"], mp.DEFAULT_REPLY)
        self.assertIn("slide_ops", data["text"])
        self.assertGreater(data["elapsedMs"], 0)
        # payload reached mock with both messages intact
        sent = mp.STATE.last_request
        self.assertEqual(len(sent["messages"]), 2)

    def test_chat_strips_reasoning(self):
        m = self.mock("openai")
        self.S.req("/api/connections", "POST", {
            "name": "g", "type": "generic-openai", "baseUrl": m.url})
        mp.STATE.reply = (
            tag("think", "let me think... ") + mp.DEFAULT_REPLY +
            tag("reason", " more reasoning", close=False)
        )
        st, data = self.S.req("/api/chat", "POST", {
            "connection": "g", "model": "mock-mini",
            "messages": [{"role": "user", "content": "hi"}]})
        self.assertEqual(data["text"], mp.DEFAULT_REPLY)

        # truncated reasoning only -> empty answer, not a crash
        mp.STATE.reply = tag("thinking", "pure ramble", close=False)
        st, data = self.S.req("/api/chat", "POST", {
            "connection": "g", "model": "mock-mini",
            "messages": [{"role": "user", "content": "hi"}]})
        self.assertEqual(data["text"], "")

    def test_chat_anthropic_payload_shape(self):
        m = self.mock("anthropic")
        self.S.req("/api/connections", "POST", {
            "name": "ant", "type": "anthropic", "baseUrl": m.url,
            "apiKey": "ant-key"})
        mp.STATE.echo_request = True
        st, data = self.S.req("/api/chat", "POST", {
            "connection": "ant", "model": "claude-x",
            "messages": [
                {"role": "system", "content": "BE SYSTEM"},
                {"role": "user", "content": "user one"},
                {"role": "user", "content": "user two"},
            ]})
        self.assertEqual(st, 200)
        # the mock echoed the request; parse and verify the anthropic mapping
        echo = json.loads(data["text"])
        self.assertEqual(echo["system"], "BE SYSTEM")
        self.assertEqual(echo["model"], "claude-x")
        roles = [m2["role"] for m2 in echo["messages"]]
        self.assertEqual(roles[0], "user")
        # consecutive user messages must be merged (anthropic requires alternation)
        self.assertEqual(len(echo["messages"]), 1)
        self.assertIn("user one", echo["messages"][0]["content"])
        self.assertIn("user two", echo["messages"][0]["content"])

    def test_chat_ollama(self):
        m = self.mock("ollama")
        self.S.req("/api/connections", "POST", {
            "name": "ol", "type": "ollama", "baseUrl": m.url})
        mp.STATE.reply = json.dumps({"kind": "reply", "text": "ollama says hi"})
        st, data = self.S.req("/api/chat", "POST", {
            "connection": "ol", "model": "llama3",
            "messages": [{"role": "user", "content": "yo"}]})
        self.assertEqual(data["text"], json.dumps({"kind": "reply", "text": "ollama says hi"}))
        self.assertFalse(mp.STATE.last_request["stream"])

    def test_chat_errors(self):
        st, data = self.S.req("/api/chat", "POST", {
            "connection": "ghost", "model": "x", "messages": []}, True)
        self.assertEqual(st, 404)
        # dead endpoint -> provider unreachable
        self.S.req("/api/connections", "POST", {
            "name": "dead", "type": "ollama", "baseUrl": "http://127.0.0.1:1"})
        st, data = self.S.req("/api/chat", "POST", {
            "connection": "dead", "model": "x",
            "messages": [{"role": "user", "content": "a"}]}, True)
        self.assertEqual(st, 502)
        self.assertIn("reach", data["error"].lower())

    def test_static_files(self):
        html = urllib.request.urlopen(self.S.base + "/taskpane.html").read().decode()
        self.assertIn("SlideKick", html)
        js = urllib.request.urlopen(self.S.base + "/taskpane.js").read().decode()
        self.assertIn("SlideKick", js)
        sim = urllib.request.urlopen(self.S.base + "/simulator.html").read().decode()
        self.assertIn("sim.js", sim)
        st, data = self.S.req("/nope.js", "GET", None, True)
        self.assertEqual(st, 404)

    def test_debug_log_and_permissions(self):
        # config file is 0600
        mode = os.stat(self.S.store.path).st_mode & 0o777
        self.assertEqual(mode, 0o600)
        # debug on -> /api/debug appends to the daily file
        self.S.req("/api/settings", "PUT", {"debug": True})
        self.S.req("/api/debug", "POST", {"category": "ui.test", "data": {"x": 1}})
        logdir = self.S.store.path.parent / "logs"
        logs = list(logdir.glob("debug-*.log"))
        self.assertEqual(len(logs), 1)
        entry = json.loads(logs[0].read_text().strip().splitlines()[-1])
        self.assertEqual(entry["category"], "ui.test")
        self.assertEqual(entry["data"], {"x": 1})
        self.assertEqual(entry["source"], "taskpane")
        # debug off -> nothing more written
        self.S.req("/api/settings", "PUT", {"debug": False})
        self.S.req("/api/debug", "POST", {"category": "ui.test", "data": {"x": 2}})
        self.assertEqual(len(logs[0].read_text().strip().splitlines()), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)