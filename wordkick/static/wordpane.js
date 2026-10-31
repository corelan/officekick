/*
 * DocKick task pane (runs inside Microsoft Word via Office.js, or in the
 * browser via sim.js for offline testing).
 *
 * Pipeline: free-text instruction -> profile/route -> /api/chat -> reasoning
 * stripping -> JSON ops extraction (hidden repair round on unclear output)
 * -> validation against deck context -> approval gate -> Office.js apply.
 */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };

  var SK = {
    app: "wordkick",
    sessionId: null,
    sessionName: "",
    sessionSaveTimer: null,
    profImageFilter: true,
    sim: false,
    deck: null,
    cfg: null,
    busy: false,
    modelsCache: {},     // connName -> {list:[...], at:ms}
    sessionApprovals: {},// opType -> 'auto' when approved for this session
    history: [],         // conversation turns shown to the model (session-only)
    editingConn: null,
    editingProf: null,
    connBaseDirty: false,
    lastCtx: null,
    activeTab: "chat",
  };

  var CONN_TYPES = [
    { id: "openai", label: "OpenAI (native)", base: "https://api.openai.com/v1", keyHint: "sk-..." },
    { id: "anthropic", label: "Anthropic (native)", base: "https://api.anthropic.com", keyHint: "sk-ant-..." },
    { id: "openrouter", label: "OpenRouter", base: "https://openrouter.ai/api/v1", keyHint: "sk-or-..." },
    { id: "generic-openai", label: "OpenAI-compatible (generic)", base: "", keyHint: "optional" },
    { id: "ollama", label: "Ollama", base: "http://localhost:11434", keyHint: "usually empty" },
    { id: "llamacpp", label: "llama.cpp (llama-server)", base: "http://localhost:8080", keyHint: "usually empty" },
  ];
  var CONN_TYPE_BY_ID = {};
  CONN_TYPES.forEach(function (t) { CONN_TYPE_BY_ID[t.id] = t; });

  // styles for insert_paragraphs live in WORD_STYLES above

  var WORD_STYLES = ["normal", "title", "heading1", "heading2", "heading3", "quote", "listParagraph"];

  var OPS_META = {
    insert_paragraphs: { label: "insert paragraphs", hint: "Adds paragraphs with a style (title/heading/quote/list/normal) at the end of the document." },
    edit_paragraph:    { label: "edit paragraph", hint: "Replaces the text of one paragraph, keeping its style." },
    replace_text:      { label: "replace text", hint: "Finds exact text and replaces it (all matches by default)." },
    create_table:      { label: "insert table", hint: "Inserts a table with the given cell values (WordApi 1.3)." },
    insert_comment:    { label: "insert comment", hint: "Adds a comment anchored to the matching text (WordApi 1.4)." },
    delete_paragraph:  { label: "delete paragraph", hint: "Removes a paragraph by index." },
    rewrite_doc:       { label: "rewrite document", hint: "Replaces the document body with a new set of paragraphs (destructive - review)." },
    insert_image:      { label: "insert image", hint: "Generates a picture from a description with the profile's image model and inserts it at the end of the document (or at the cursor with the checkbox)." },
    format_text:      { label: "format text", hint: "bold/italic/color/size/underline/allCaps/highlight + alignment on one paragraph." },
    create_template:  { label: "create template block", hint: "appends a titled template block with fillable content controls (text, dropdown, checkbox, date)." },
  };
  var OP_TYPES = Object.keys(OPS_META);

  function FriendlyError(message) {
    this.name = "FriendlyError";
    this.friendly = true;
    this.message = message;
  }
  FriendlyError.prototype = Object.create(Error.prototype);

  // ------------------------------------------------------------------ //
  // Misc helpers
  // ------------------------------------------------------------------ //
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }

  function setStatus(text, cls) {
    var el = $("chat-status");
    el.className = "status" + (cls ? " " + cls : "") + (text ? "" : " hidden");
    el.textContent = text || "";
  }

  // theme: follow the Office host (isDarkTheme covers Black/Dark Gray; fall
  // back to body-color luminance; simulator follows the browser or ?theme=)
  function luminance(hexColor) {
    var m = /^\s*#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i.exec(hexColor || "");
    if (!m) return 1;
    return (0.2126 * parseInt(m[1], 16) +
            0.7152 * parseInt(m[2], 16) +
            0.0722 * parseInt(m[3], 16)) / 255;
  }

  function applyTheme() {
    var dark = false, mode = "none";
    try {
      var t = Office && Office.context && Office.context.officeTheme;
      if (t) {
        if (typeof t.isDarkTheme === "boolean") { dark = t.isDarkTheme; mode = "host.isDarkTheme"; }
        else if (t.bodyBackgroundColor) { dark = luminance(t.bodyBackgroundColor) < 0.5; mode = "host.luminance"; }
      }
    } catch (e) { /* theme API not available */ }
    if (mode === "none") {
      var q = new URLSearchParams(String(window.location.search || ""));
      var forced = q.get("theme");
      if (forced) { dark = forced.toLowerCase() === "dark"; mode = "forced"; }
      else if (window.matchMedia) { dark = window.matchMedia("(prefers-color-scheme: dark)").matches; mode = "browser"; }
    }
    document.body.classList.toggle("dark", dark);
    SK.theme = { dark: dark, via: mode };
    dbg("theme", SK.theme);
  }

  function watchTheme() {
    try {
      Office.context.addHandlerAsync(Office.EventType.OfficeThemeChanged, function () {
        applyTheme();
        setStatus("Switched to " + (SK.theme.dark ? "dark" : "light") + " theme.", "");
      }, function (asyncResult) { /* subscribed (or unsupported) */ });
    } catch (e) { /* event not available on this host */ }
    document.addEventListener("visibilitychange", function () {
      if (!document.hidden) applyTheme();
    });
  }

  function setFormMsg(el, text, ok) {
    el.textContent = text || "";
    el.className = "formmsg" + (ok ? " ok" : text ? " bad" : "");
  }

  function dbg(category, data) {
    try {
      if (!SK.cfg || !SK.cfg.debug) return;
      fetch("/api/debug", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ category: "ui." + category, data: data }),
      }).catch(function () {});
    } catch (e) { /* never break the pane over logging */ }
  }

  function logApi(entry) {
    SK.apiLog = SK.apiLog || [];
    entry.at = entry.at || Date.now();
    SK.apiLog.unshift(entry);
    if (SK.apiLog.length > 80) SK.apiLog.pop();
    if (SK.activeTab === "requests") renderRequests();
  }

  function fmtClock(ms) {
    var d = new Date(ms);
    return ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2) +
      ":" + ("0" + d.getSeconds()).slice(-2);
  }

  function renderRequests() {
    var box = $("req-list");
    if (!box) return;
    var entries = SK.apiLog || [];
    if (!entries.length) {
      box.innerHTML = '<div class="opnote">No requests yet. Every AI call (chat, models, image) is logged here with status, duration and payloads.</div>';
      return;
    }
    // newest on top; show up to 6 boxes; only the newest starts open;
    // click a box header to open/close it
    var shown = entries.slice(0, 6);
    var frag = document.createDocumentFragment();
    if (entries.length > shown.length) {
      var more = document.createElement("div");
      more.className = "opnote";
      more.textContent = "(" + (entries.length - shown.length) + " older request(s) not shown)";
      frag.appendChild(more);
    }
    shown.forEach(function (e, i) {
      var grpb = document.createElement("div");
      grpb.className = "reqgrp " + (e.ok ? "reqok" : "reqbad");
      var head = document.createElement("div");
      head.className = "reqhead2";
      head.title = "Click to open/close this request/response";
      var ts = document.createElement("b");
      ts.textContent = fmtClock(e.at) + "  ";
      head.appendChild(ts);
      var badge = document.createElement("span");
      badge.className = "reqbadge " + (e.ok ? "ok" : "bad");
      badge.textContent = e.ok ? ("HTTP " + e.status) : (e.status ? ("HTTP " + e.status) : "NETWORK");
      head.appendChild(badge);
      var lab = document.createElement("span");
      lab.textContent = " " + e.method + " " + e.path +
        (e.conn ? "  [" + e.conn + (e.model ? ":" + e.model : "") + "]" : "") +
        "  " + e.ms + " ms";
      head.appendChild(lab);
      var body = document.createElement("div");
      body.className = "reqbody";
      if (e.err) {
        var err = document.createElement("div");
        err.className = "reqerr";
        err.textContent = e.err;
        body.appendChild(err);
      }
      if (e.req || e.resp) {
        var pre = document.createElement("pre");
        pre.textContent = (e.req ? "Request:\n" + e.req : "") + (e.resp ? "\nResponse:\n" + e.resp : "");
        body.appendChild(pre);
      } else {
        var none = document.createElement("div");
        none.className = "opnote";
        none.textContent = "(no payload captured)";
        body.appendChild(none);
      }
      if (i === 0) body.classList.add("open");
      head.addEventListener("click", function () { body.classList.toggle("open"); });
      grpb.appendChild(head);
      grpb.appendChild(body);
      frag.appendChild(grpb);
    });
    box.innerHTML = "";
    box.appendChild(frag);
  }

  function clearRequests() {
    SK.apiLog = [];
    renderRequests();
  }


  function testConnRow(name, msgEl) {
    msgEl.textContent = "Testing connection '" + name + "'…";
    apiJson("/api/test-connection?connection=" + encodeURIComponent(name), "GET")
      .then(function (res) {
        if (res && res.ok) msgEl.textContent = "'" + name + "' reachable: " + res.models + " model(s), " + res.ms + " ms";
        else msgEl.textContent = "'" + name + "' FAILED: " + ((res && res.error) || res || "unknown");
      }).catch(function (e) { msgEl.textContent = "'" + name + "' FAILED: " + (e && e.message || e); });
  }


  // ---- live web lookups (optional checkbox): ONE call through a model with
  // web access; the MODEL does the lookup (no local search API keys). Any
  // model on an openrouter connection gets the ":online" suffix so
  // openrouter injects fresh web results + source citations. Other
  // providers are asked too; models without web access answer
  // NO_WEB_ACCESS and the pane says so in the status. ----
  function webLookup(prof, instruction) {
    var ref = prof.other && prof.other.connection && prof.other.model ? prof.other : prof.content;
    var model = ref.model;
    var conn = (SK.cfg.connections || []).filter(function (c) { return c.name === ref.connection; })[0] || {};
    if (conn.type === "openrouter") model = model + ":online";
    var sys = "You have LIVE WEB ACCESS. Research the user's request with web search now. " +
      "Return 4-8 short facts that help answer it; after EACH fact put its source URL in ( ): " +
      "like \"fact text (https://source.example/page)\". Be current; mention dates when known. " +
      "If you actually cannot access the web, reply with exactly NO_WEB_ACCESS.";
    return askRetryable(function () {
      return apiJson("/api/chat", "POST", {
        connection: ref.connection, model: model,
        messages: [{ role: "system", content: sys },
                   { role: "user", content: instruction.slice(0, 600) }],
      });
    }).then(function (resp) {
      var t = String(resp.text || "").trim();
      if (!t || /^NO_WEB_ACCESS\b/i.test(t) || /cannot access (the )?web/i.test(t) ||
          /can'?t browse/i.test(t) || /no web access/i.test(t)) return "";
      return "MODEL-SIDE WEB LOOKUP via " + model + ":\n" + t.slice(0, 5000);
    });
  }

  function autoApplyWanted() {
    return !!(SK.cfg && SK.cfg.autoApplyApproved);
  }

  function syncAutoApplyCheckbox() {
    var box = $("chat-autoapply");
    if (box && autoApplyWanted()) box.checked = true;
    if (box && !SK.autoApplyUserTouched) box.checked = autoApplyWanted();
  }

  function wireAutoApplySetting() {
    var cb = $("set-autoapply");
    if (!cb) return;
    cb.checked = autoApplyWanted();
    cb.addEventListener("change", function () {
      var on = !!cb.checked;
      apiJson("/api/settings", "PUT", { autoApplyApproved: on }).then(function () {
        if (SK.cfg) SK.cfg.autoApplyApproved = on;
        var box = $("chat-autoapply");
        if (box) box.checked = on;           // follow the global immediately
        syncAutoApplyCheckbox();
        var msg = $("set-autoapply-msg");
        if (msg) msg.textContent = on ? "ON - every run auto-applies approved actions." : "OFF - the per-run checkbox stays manual.";
      }).catch(function (e) {
        var msg = $("set-autoapply-msg");
        if (msg) msg.textContent = "Save failed: " + ((e && e.message) || e);
      });
    });
  }

  function updCheck() {
    var btn = $("upd-check"), info = $("upd-info");
    if (!btn || !info) return;
    btn.disabled = true;
    info.textContent = "Checking for updates…";
    apiJson("/api/update-check", "GET").then(function (data) {
      btn.disabled = false;
      if (!data || data.ok === false) {
        info.textContent = data && data.error ? data.error : "could not check for updates";
        return;
      }
      if (data.upToDate) {
        info.textContent = "You are up to date (" + data.current + ", latest " + data.latest + ").";
      } else {
        info.textContent = "Update available: " + data.latest + " (you run " + data.current + "). " +
          (data.howToUpdate || "");
      }
    }).catch(function (e) {
      btn.disabled = false;
      info.textContent = "could not check for updates: " + (e && e.message ? e.message : e);
    });
  }

  function apiJson(path, method, body) {
    var t0 = Date.now();
    var verb = method || "GET";
    var opts = { method: verb, headers: {} };
    var bodyStr = "";
    if (body !== undefined) {
      opts.headers["Content-Type"] = "application/json";
      bodyStr = JSON.stringify(body);
      opts.body = bodyStr;
    }
    var fetchPromise;
    try {
      fetchPromise = fetch(path, opts);
    } catch (syncErr) {
      logApi({ path: path, method: verb, status: 0, ok: false,
        err: "network (sync): " + String(syncErr && syncErr.message || syncErr),
        req: bodyStr.slice(0, 1600) });
      throw new FriendlyError("Cannot reach the OfficeKick server (" + path + ")");
    }
    return fetchPromise.then(function (resp) {
      return resp.text().then(function (text) {
        var data = null;
        try { data = text ? JSON.parse(text) : {}; } catch (e) { /* non-json */ }
        var okResp = resp.ok;
        logApi({
          path: path, method: verb,
          conn: body && body.connection || null,
          model: body && body.model || null,
          status: resp.status, ms: Date.now() - t0, ok: !!okResp,
          err: okResp ? null : (data && data.error ? String(data.error) : "HTTP " + resp.status),
          req: bodyStr.slice(0, 1600), resp: (text || "").slice(0, 1600),
        });
        if (!okResp) {
          var msg = data && data.error ? data.error : "HTTP " + resp.status;
          throw new FriendlyError(msg);
        }
        return data || {};
      });
    }, function (fetchErr) {
      logApi({ path: path, method: verb, status: 0, ok: false,
        conn: body && body.connection || null, model: body && body.model || null,
        err: "network: " + String(fetchErr && fetchErr.message || fetchErr),
        req: bodyStr.slice(0, 1600) });
      throw new FriendlyError("Cannot reach the OfficeKick server (" + path + "): " +
        String(fetchErr && fetchErr.message || fetchErr));
    });
  }

  // ------------------------------------------------------------------ //
  // Reasoning stripping (chain-of-thought wrappers). Tag strings are built
  // via character codes so they never appear verbatim in this file.
  // ------------------------------------------------------------------ //
  var STRIP_RES = (function () {
    var lt = String.fromCharCode(60), gt = String.fromCharCode(62);
    var tags = ["think", "thinking", "reason", "reasoning"];
    var res = [];
    tags.forEach(function (tag) {
      var opener = lt + tag + gt;
      var closer = lt + "/" + tag + gt;
      res.push(new RegExp(escapeRe(opener) + "[\\s\\S]*?" + escapeRe(closer) + "\\s*", "gi"));
      res.push(new RegExp(escapeRe(opener) + "[\\s\\S]*$", "i"));
    });
    return res;
  })();
  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

  function stripReasoning(text) {
    var out = String(text == null ? "" : text);
    STRIP_RES.forEach(function (rx) { out = out.replace(rx, ""); });
    return out.trim();
  }

  // ------------------------------------------------------------------ //
  // JSON extraction (tolerant): direct parse, fence strip, balanced scan.
  // ------------------------------------------------------------------ //
  var FENCE = String.fromCharCode(96, 96, 96);

  function extractObject(text) {
    var t = String(text == null ? "" : text).trim();
    var candidates = [t];
    var noFence = t.split(FENCE).join("").trim(); // remove all fence markers
    if (noFence !== t) candidates.push(noFence);
    for (var i = 0; i < candidates.length; i++) {
      try { return JSON.parse(candidates[i]); } catch (e) { /* continue */ }
    }
    // scan for the first balanced { ... } (string-aware)
    var start = t.indexOf("{");
    if (start >= 0) {
      var depth = 0, inStr = false, escNext = false;
      for (var k = start; k < t.length; k++) {
        var c = t.charAt(k);
        if (inStr) {
          if (escNext) escNext = false;
          else if (c === "\\") escNext = true;
          else if (c === "\"") inStr = false;
          continue;
        }
        if (c === "\"") { inStr = true; continue; }
        if (c === "{") depth++;
        else if (c === "}") {
          depth--;
          if (depth === 0) {
            try { return JSON.parse(t.slice(start, k + 1)); } catch (e) { break; }
          }
        }
      }
    }
    return undefined;
  }

  var SYSTEM_PROMPT = [
    "You are DocKick, an AI assistant inside a Microsoft Word task pane.",
    "You get the document outline and a user request. Reply with ONLY one JSON",
    "object - never commentary, never code fences, never reasoning. Format A",
    "(conversational):",
    '{"kind":"reply","text":"your answer"}',
    "Format B (document changes):",
    '{"kind":"doc_ops","ops":[ ... ]}',
    'Accepted also: {"kind":"slide_ops","ops":[...]} (treated like doc_ops).',
    "",
    "Allowed op types:",
    '- {"type":"insert_paragraphs","items":[{"text":"...","style":"normal|title|heading1|heading2|heading3|quote|listParagraph"}]} (appended at the end of the document; start a rewrite with a heading)',
    '- {"type":"edit_paragraph","para":N,"text":"..."} (keeps the paragraph style)',
    '- {"type":"replace_text","find":"exact text","replace":"new text","all":true} (all matches by default)',
    '- {"type":"create_table","values":[["h1","h2"],["a","b"]]} (inserted at the end of the document)',
    '- {"type":"insert_comment","find":"exact anchor text","text":"comment"} (comments need WordApi 1.4)',
    '- {"type":"delete_paragraph","para":N} (para = the paragraph number shown in the outline; 1 = the first paragraph - never invent them)',
    '- {"type":"rewrite_doc","items":[{"text":"...","style":"..."}]} (DESTROYS the current body - only when explicitly asked for a full rewrite)',
    '- {"type":"create_template","title":"Template name","fields":[{"label":"Field label","type":"text|dropdown|checkbox|date","options":["only for dropdown"]}]} (appends a titled template block with fillable controls; use when the user asks for a template/form/template block; dropdown needs an options array)',
    '- {"type":"insert_image","prompt":"detailed image description"} (generate a picture from the description and insert it at the end of the document; only when the user asks for a picture/image/illustration/logo)',
    '- {"type":"format_text","para":N,"font":{"bold":true},"align":"center"} (format ONE paragraph from the outline: font = {"bold","italic","color","size","name","underline","allCaps","highlightColor"} subset; align = left|center|right|justify; only what is requested)',
    '- to format only PART of a paragraph, add "find":"exact words appearing in that paragraph" (all matches within the paragraph are formatted; e.g. {"type":"format_text","para":N,"find":"AI","font":{"bold":true}})',
    "",
    "Rules:",
    "- Paragraph numbers are 1-based (para 1 = the first paragraph) and MUST come from the outline below. If the user says \"the 3rd paragraph\", that is para 3. Never invent them.",
    "- Rewrite/improve requests: keep the meaning and structure; reword only.",
    "- If the request is a question/explanation/brainstorm rather than a document change, use format A.",
    "- Image requests: always return an insert_image op (never generate the picture inline); DocKick fetches it with the user's configured image model. Include a concrete, self-contained prompt in the op.",
    "- Image refinement ('refine/change the image'): write a FRESH self-contained image prompt describing the FULL desired result. Never embed previous prompts, the conversation, or op JSON into it - only the scene essentials plus the requested change. Prefer 200-600 words.",
    "- LIVE WEB LOOKUPS: when the request needs up-to-date info you cannot reliably know (current events, news, prices, recent releases, rankings, live data), reply with ONLY {\"kind\":\"web_lookup\"} - nothing else. SlideKick then runs a web-enabled lookup and re-asks you with fresh citable results. Never use it for deck/document/mail edits, styling, general concepts, or what the outline/context already answers.",

    "- Picture placement: 'here / at the cursor / current position' = the image lands at the caret (the pane's 'insert at cursor' checkbox); default = end of document. Keep insert_image minimal (type + prompt only).",
  ].join("\n");

  var NOTES_OFF_RULE = "";

  var REPAIR_ASK =
    "Your previous answer did not follow the required format. " +
    "Reply again with ONLY the JSON object (format A or B) and nothing else: " +
    'either {"kind":"reply","text":"..."} or {"kind":"slide_ops","ops":[...]} - ' +
    "no reasoning, no fences, no commentary.";

  var HISTORY_MAX_EXCHANGES = 6; // oldest turns are dropped, newest kept
  var HISTORY_MAX_CHARS = 2000;  // per stored message

  // ------------------------- local sessions ------------------------- //
  function derivedSessionName() {
    var first = "";
    for (var i = 0; i < SK.history.length; i++) {
      if (SK.history[i].role === "user") { first = SK.history[i].content; break; }
    }
    return String(first).replace(/\s+/g, " ").trim().slice(0, 40) || "new session";
  }

  function updateSessionLabel() {
    var el = $("chat-session-label");
    if (!el) return;
    if (SK.history.length && SK.sessionId) {
      el.textContent = "Session: " + (SK.sessionName || derivedSessionName());
      el.style.display = "";
    } else {
      el.textContent = "";
      el.style.display = "none";
    }
  }

  function scheduleSaveSession() {
    if (SK.sessionSaveTimer) clearTimeout(SK.sessionSaveTimer);
    SK.sessionSaveTimer = setTimeout(saveSessionNow, 600);
  }

  function saveSessionNow() {
    if (SK.testMode) return;   // test runs never end up in history
    if (!SK.history.length) return;
    var payload = { exchanges: SK.history.slice(), app: SK.app };
    var createPromise;
    if (!SK.sessionId) {
      payload.name = derivedSessionName();
      createPromise = apiJson("/api/sessions", "POST", payload).then(function (res) {
        if (res && res.session) {
          SK.sessionId = res.session.id;
          SK.sessionName = res.session.name;
          aiNameSession(SK.sessionId);
        }
        return res;
      });
    } else {
      createPromise = apiJson("/api/sessions/" + SK.sessionId, "PUT", payload).then(function (up) {
        var doc2 = up && up.session;
        if (doc2 && doc2.exchanges && doc2.exchanges.length >= 2 && doc2.exchanges.length < SK.history.length) {
          // the store did not keep the history (old-server edge): forget the
          // dead id, the next save starts a fresh session with the full history
          SK.sessionId = null;
          delete payload.name;
        } else if (doc2 && doc2.name) {
          SK.sessionName = doc2.name;
        }
        return up;
      });
    }
    return createPromise.then(function () {
      dbg("session.saved", { id: SK.sessionId, name: SK.sessionName, messages: SK.history.length });
    }).catch(function (err) {
      dbg("session.saveFailed", { msg: (err && err.message) || String(err) });
    });
  }

  function aiNameSession(sid) {
    if (SK.testMode) return;   // no AI calls for test sessions
    // one tiny call through the profile's other slot; the derived name stays
    // if anything fails, and the server repairs/cleans what the model returns
    var prof = currentProfile();
    if (!prof || !prof.other || !prof.other.connection || !prof.other.model) return;
    var first = derivedSessionName();
    var renameFallback = function () {
      apiJson("/api/sessions/rename/" + sid, "POST", { name: first.slice(0, 48) });
    };
    apiJson("/api/chat", "POST", {
      connection: prof.other.connection, model: prof.other.model,
      temperature: 0.2, maxTokens: 300,
      messages: [
        { role: "system", content: "You generate very short chat-session titles. Reply with ONLY the title: 2-4 words, no quotes, no punctuation at the end." },
        { role: "user", content: "Session start: " + first.slice(0, 120) },
      ],
    }).then(function (resp) {
      var name = stripReasoning(resp.text || "").replace(/"/g, "").replace(/^[:\- ]+/, "").trim().slice(0, 48);
      // reject non-title junk (models that ignore the instruction and echo ops/json)
      var junk = /^[{\[]/.test(name) || /"(kind|ops|type)"/.test(name) || name.length < 2 || name.length > 60;
      if (junk) { renameFallback(); return; }
      return apiJson("/api/sessions/rename/" + sid, "POST", { name: name }).then(function () {
        SK.sessionName = name;
        dbg("session.named", { id: sid, name: name });
        updateSessionLabel();
        if (SK.activeTab === "sessions") renderSessions();
      });
    }).catch(function () {
      dbg("session.namingFailed", { id: sid });   // derived name stays
    });
  }

  function resumeSession(sessId) {
    apiJson("/api/sessions/" + sessId, "GET").then(function (doc) {
      SK.history = (doc.exchanges || []).slice(-HISTORY_MAX_EXCHANGES * 2);
      SK.sessionId = doc.id;
      SK.sessionName = doc.name || "";
      var inp1 = $("chat-input");
      if (inp1) { inp1.disabled = false; inp1.value = ""; }
      switchTab("chat");
      setStatus("Session resumed: \u201C" + SK.sessionName + "\u201D (" + SK.history.length + " message(s) loaded).", "ok");
      updateSessionLabel();
      pushHistoryStubRender();
      // history must show real titles: a generic leftover "Session" gets an
      // AI title based on the interactions, stored in the session file
      if (/^session\b\s*$/i.test(String(SK.sessionName || ""))) aiNameSession(doc.id);
    }).catch(function (err) {
      setStatus("Could not load session: " + ((err && err.message) || err), "error");
    });
  }

  function pushHistoryStubRender() {
    var el = $("chat-output");
    if (!el) return;
    el.innerHTML = "";
    var note = document.createElement("div");
    note.className = "opnote";
    note.textContent = "Resumed session \u201C" + SK.sessionName + "\u201D - the AI sees the last " +
      Math.min(HISTORY_MAX_EXCHANGES, SK.history.length / 2) + " exchange(s). The last turns are in the Previous requests box below (collapsed; click its header to open). Continue typing below.";
    el.appendChild(note);
    // show the last up-to-6 messages (3 exchanges) as collapsed request/response
    // boxes, styled like the Inspector groups; future runs append below
    // the last 6 interactions (newest first) as flat rows in the box
    var list = SK.history.slice(0);
    while (list.length % 2) list.shift();
    var turns = [];
    for (var mi = 0; mi + 1 < list.length; mi += 2) turns.push([list[mi], list[mi + 1]]);
    var pastBody = $("past-body");
    if (pastBody) {
      pastBody.innerHTML = "";
      for (var ti = turns.length - 1; ti >= 0; ti--) {   // every run of the session, newest first
        pastBody.appendChild(pastTurnRow(turns[ti][0], turns[ti][1]));
      }
      wirePastHeader();   // wires the open/close click + sets the visibility (the pm glyph)
    }
    renderSessions();  // refresh the "current" label
  }


  function deleteSession(sessId) {
    apiJson("/api/sessions/" + sessId, "DELETE").then(function () {
      if (SK.sessionId === sessId) { SK.sessionId = null; SK.sessionName = ""; }
      renderSessions();
      setStatus("Session deleted.", "ok");
    }).catch(function (err) {
      setStatus("Could not delete session: " + ((err && err.message) || err), "error");
    });
  }

  function renameSessionInline(sessId, nameEl) {
    var input = document.createElement("input");
    input.type = "text";
    input.maxLength = 48;
    input.value = SK.sessionNames[sessId] || "";
    input.className = "sess-rename";
    nameEl.textContent = "";
    nameEl.appendChild(input);
    input.focus();
    input.select();
    var done = function (commit) {
      var val = input.value.trim();
      if (commit && val) {
        apiJson("/api/sessions/rename/" + sessId, "POST", { name: val }).catch(function () {}).then(function () {
          SK.sessionNames[sessId] = val;
          if (SK.sessionId === sessId) SK.sessionName = val;
          renderSessions();
        });
      } else {
        renderSessions();
      }
    };
    input.addEventListener("keydown", function (e) {
      if (e.key === "Enter") done(true);
      else if (e.key === "Escape") done(false);
      else if (e.key === "Tab") { e.preventDefault(); done(true); }
    });
    input.addEventListener("blur", function () { done(false); });
  }

  function showSessionChat(sessId, card) {
    var box = card.querySelector(".sesschat");
    if (!box.classList.contains("hidden")) { box.classList.add("hidden"); box.innerHTML = ""; return; }
    apiJson("/api/sessions/" + sessId, "GET").then(function (doc) {
      var msgs = (doc.exchanges || []);
      // cap the peek at 6 entries: the last 3 requests + 3 responses
      var shown = msgs.slice(-6);
      var frag = document.createDocumentFragment();
      if (msgs.length > shown.length) {
        var more = document.createElement("div");
        more.className = "opnote";
        more.textContent = "(showing the last " + shown.length + " of " + msgs.length + " messages)";
        frag.appendChild(more);
      }
      shown.forEach(function (m) {
        var b = document.createElement("div");
        b.className = "bubble " + (m.role === "user" ? "user" : "assistant");
        var who = document.createElement("b");
        who.textContent = (m.role === "user" ? "You" : "AI") + ": ";
        var c = String(m.content || "");
        if (m.role === "user" && m.ts) {
          var its = document.createElement("span");
          its.className = "ptts";
          its.textContent = m.ts + "  ";
          b.appendChild(its);
        }
        if (m.role === "user") {
          b.style.cursor = "pointer";
          var tclick = String(m.content || "");
          b.addEventListener("click", function () { copyRequestToInput(tclick); });
        }
        if (c.length > 240) c = c.slice(0, 240) + "…";
        var txt = document.createElement("span");
        txt.textContent = c;
        b.appendChild(who); b.appendChild(txt);
        frag.appendChild(b);
      });
      box.innerHTML = "";
      box.appendChild(frag);
      box.classList.remove("hidden");
    }).catch(function (err) {
      box.innerHTML = '<div class="opnote">Could not load: ' + esc((err && err.message) || String(err)) + "</div>";
      box.classList.remove("hidden");
    });
  }

  SK.sessionNames = {};
  function renderSessions() {
    apiJson("/api/sessions?app=" + encodeURIComponent(SK.app), "GET").then(function (res) {
      var rows = (res.sessions || []);
      var list = $("sess-list");
      if (!list) return;
      var frag = document.createDocumentFragment();
      if (!rows.length) {
        var empty = document.createElement("div");
        empty.className = "opnote";
        empty.textContent = "No sessions yet. Every conversation is saved locally after the first exchange; come back and resume it anytime.";
        frag.appendChild(empty);
      }
      rows.forEach(function (row) {
        SK.sessionNames[row.id] = row.name;
        var card = document.createElement("div");
        card.className = "sesscard" + (row.id === SK.sessionId ? " current" : "");
        var head = document.createElement("div");
        head.className = "sesshead";
        var name = document.createElement("span");
        name.className = "sessname";
        name.title = "Click to show the last interactions";
        name.textContent = row.name;
        name.addEventListener("click", function () { showSessionChat(row.id, card); });
        head.appendChild(name);
        var when = new Date(row.updated * 1000);
        var whenStr = ("0" + when.getHours()).slice(-2) + ":" + ("0" + when.getMinutes()).slice(-2) + " · " +
          when.toLocaleDateString();
        var meta = document.createElement("span");
        meta.className = "sessmeta";
        meta.textContent = row.messages + " msg" + (row.messages === 1 ? "" : "s") + " · " + whenStr +
          (row.id === SK.sessionId ? " · current session" : "");
        head.appendChild(meta);
        var actions = document.createElement("span");
        actions.className = "sessactions";
        var renameBtn = document.createElement("button");
        renameBtn.className = "btn subtle";
        renameBtn.textContent = "Rename";
        renameBtn.title = "Rename this session";
        renameBtn.addEventListener("click", function () {
          var input = document.createElement("input");
          input.type = "text"; input.maxLength = 48;
          input.value = SK.sessionNames[row.id] || row.name;
          input.className = "sess-rename";
          name.textContent = ""; name.appendChild(input);
          input.focus(); input.select();
          var done = function (commit) {
            var val = input.value.trim();
            if (commit && val) {
              apiJson("/api/sessions/rename/" + row.id, "POST", { name: val }).catch(function () {}).then(function () {
                SK.sessionNames[row.id] = val;
                SK.sessionName = val;
                updateSessionLabel();
                renderSessions();
              });
            } else { renderSessions(); }
          };
          input.addEventListener("keydown", function (e) {
            if (e.key === "Enter") done(true);
            else if (e.key === "Escape") done(false);
            else if (e.key === "Tab") { e.preventDefault(); done(true); }
          });
          input.addEventListener("blur", function () { done(false); });
        });
        actions.appendChild(renameBtn);
        [["Resume", false], ["Delete", true]].forEach(function (pair) {
          var b = document.createElement("button");
          b.className = "btn subtle" + (pair[1] ? " danger" : "");
          b.textContent = pair[0];
          b.addEventListener("click", function () {
            if (pair[1] && b.dataset.armed !== "1") {
              // two-stage confirm: the second click within 3.5 s deletes
              b.dataset.armed = "1";
              b.textContent = "Really delete?";
              b.classList.add("armed");
              setTimeout(function () {
                if (b.parentNode) { b.dataset.armed = "0"; b.textContent = "Delete"; b.classList.remove("armed"); }
              }, 3500);
              return;
            }
            if (pair[1]) deleteSession(row.id); else resumeSession(row.id);
          });
          actions.appendChild(b);
        });
        head.appendChild(actions);
        card.appendChild(head);
        var chat = document.createElement("div");
        chat.className = "sesschat hidden";
        card.appendChild(chat);
        frag.appendChild(card);
      });
      list.innerHTML = "";
      list.appendChild(frag);
      var cur = $("sess-current-label");
      if (cur) cur.textContent = SK.sessionId ? ("current: " + (SK.sessionName || SK.sessionId)) : "current: (new / unsaved)";
    }).catch(function (err) {
      var list = $("sess-list");
      if (list) list.innerHTML = '<div class="opnote">Cannot reach the sessions store: ' + esc((err && err.message) || String(err)) + "</div>";
    });
  }

  // wire the sessions tab buttons (called from wireEvents)
  function wireSessions() {
    var nw = $("sess-new");
    if (nw) nw.addEventListener("click", function () {
      SK.history = [];
      SK.sessionId = null;
      SK.sessionName = "";
      var inp0 = $("chat-input");
      if (inp0) { inp0.disabled = false; inp0.value = ""; }
      switchTab("chat");
      setStatus("New session - the current one stays saved.", "");
      var el = $("chat-output");
      if (el) el.innerHTML = "";
    });
    var rf = $("sess-refresh");
    if (rf) rf.addEventListener("click", renderSessions);
  }

  var CONTENT_ROUTE_RE = new RegExp(
    "\\b(document|doc|paragraph|paragraphs|text|wording|reword|rewrite|rephrase|outline|comment|comments|table|tables|add|insert|create|make|generate|tighten|restructure|expand|delete|tone|image|images|picture|pictures|photo|illustration|logo|bold|italic|underline|color|font|highlight|format|highlight)\\b|template|form|dropdown|checkbox|date", "i"
  );

  function assistantHistoryText(parsed) {
    // a compact recap for the next round; never the raw ops JSON (it bloats every
    // later request and pushes the models toward echoing blobs into new prompts)
    if (!parsed || !parsed.ok) {
      var rawT = (SK.lastModelText && SK.lastModelText.text) || "";
      return String(rawT).slice(0, 400);
    }
    if (parsed.kind === "reply") return String(parsed.text || "").slice(0, HISTORY_MAX_CHARS);
    // store only the action types and how many per type were executed -
    // no descriptions, no prompts (keeps the session/context compact)
    var ops = parsed.ops || [];
    var counts = {};
    ops.forEach(function (op) {
      var t = op && op.type ? String(op.type) : "unknown";
      counts[t] = (counts[t] || 0) + 1;
    });
    var byType = Object.keys(counts).map(function (t) { return t + " x" + counts[t]; }).join(", ");
    return ("Executed " + ops.length + " action(s)" + (byType ? ": " + byType : "")).slice(0, HISTORY_MAX_CHARS);
  }

  function historyCap(text) {
    var s = String(text || "");
    return s.length > HISTORY_MAX_CHARS ? s.slice(0, HISTORY_MAX_CHARS) + " [truncated]" : s;
  }

  function nowStamp() {
    var d = new Date();
    function z(n) { return (n < 10 ? "0" : "") + n; }
    return d.getFullYear() + "-" + z(d.getMonth() + 1) + "-" + z(d.getDate()) + " " +
      z(d.getHours()) + ":" + z(d.getMinutes());
  }

  // clicking a stored request reloads it into the input box, but only when
  // the input is empty (never clobber a typed draft)
  function copyRequestToInput(fullText) {
    var inp = $("chat-input");
    if (!inp || inp.value) return;
    inp.value = String(fullText || "");
    inp.focus();
  }

  function pushHistory(userText, assistantText) {
    var stamp = nowStamp();
    SK.history.push({ role: "user", content: historyCap(userText), ts: stamp });
    SK.history.push({ role: "assistant", content: historyCap(assistantText), ts: stamp });
    // persist every request/response immediately - a pane closing right after
    // an exchange must never lose it (no debounce here)
    saveSessionNow();
    while (SK.history.length > HISTORY_MAX_EXCHANGES * 2) SK.history.shift();
    updateChatMeta();
    dbg("history.push", { exchanges: Math.floor(SK.history.length / 2) });
  }

  function resetHistory() {
    SK.history = [];
    SK.sessionId = null;      // the next exchange saves as a NEW session
    SK.sessionName = "";
    updateChatMeta();
    updateSessionLabel();
    renderSessions();
    dbg("history.reset", {});
  }

  function updateChatMeta() {
    var el = $("chat-history-note");
    if (!el) return;
    var n = Math.floor(SK.history.length / 2);
    el.textContent = n ? "AI sees the last " + n + " exchange(s)" : "new session";
  }

  function routeToModel(prof, route, instruction) {
    if (route === "content") return prof.content;
    if (route === "other") return prof.other;
    return CONTENT_ROUTE_RE.test(instruction) ? prof.content : prof.other;
  }

  // ------------------------------------------------------------------ //
  // Word document adapter (thin wrapper around Word.run)
  // ------------------------------------------------------------------ //
    function WordDeck() {
    this.word = true;
    this._tablesApi = undefined; // undefined = probe needed
    this._cmtApi = undefined;
    this._caps = null;
  }

  WordDeck.prototype.capabilities = function () {
    if (!this._caps) {
      var c = { text: true, comments: false, tables: false, selection: true, host: "" };
      try { c.comments = Office.context.requirements.isSetSupported("WordApi", "1.4"); } catch (e) {}
      try { c.tables = Office.context.requirements.isSetSupported("WordApi", "1.3"); } catch (e) {}
      try { c.host = "Word " + ((Office.context.diagnostics || {}).version || ""); } catch (e) {}
      this._caps = c;
    }
    return this._caps;
  };

  WordDeck.prototype._wrap = function (err) {
    var raw = (err && (err.message || "")) || String(err);
    var msg = raw;
    if (/insertTable|PropertyOrMethodNotSupported|APINotAvailable/i.test(raw)) {
      msg = "A required Word.js API is missing in this build (" + raw + "). Update Microsoft 365.";
    }
    return new FriendlyError(msg);
  };

  WordDeck.prototype.run = function (name, fn) {
    var t0 = Date.now();
    return Word.run(fn).then(
      function (result) {
        dbg("office", { method: name, ms: Date.now() - t0, ok: true });
        return result;
      },
      function (err) {
        dbg("office", { method: name, ms: Date.now() - t0, ok: false, error: String(err) });
        throw this._wrap(err);
      }.bind(this)
    );
  };

  WordDeck.prototype.getContext = function (opts) {
    opts = opts || {};
    var self = this;
    return this.run("getContext", function (ctx) {
      var body = ctx.document.body;
      var paras = body.paragraphs;
      paras.load("text, style");
      var sel = ctx.document.getSelection();
      sel.load("text");
      var tables = body.tables;
      tables.load("itemCount");
      return ctx.sync().then(function () {
        var n = Math.min(opts.maxParas || paras.items.length, 300);
        var t0 = null;
        var items = [];
        for (var i = 0; i < n; i++) {
          items.push({ idx: i, text: paras.items[i].text, style: paras.items[i].style || "Normal" });
        }
        return ctx.sync().then(function () {
          return { documentName: "document", paragraphCount: paras.items.length,
                   paragraphs: items, tables: [], selected: sel.text || "", host: self.capabilities().host };
        });
      });
    });
  };

  WordDeck.prototype.insertParagraphs = function (op) {
    return this.run("insertParagraphs", function (ctx) {
      var body = ctx.document.body;
      var items = op.items || [];
      items.forEach(function (it) {
        var p = body.insertParagraph(String(it.text == null ? "" : it.text), "End");
        var st = styleNameOf(it.style);
        if (st) p.style = st;
      });
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  WordDeck.prototype.editParagraph = function (index, text) {
    return this.run("editParagraph", function (ctx) {
      // Word.js collections: use .items after load+sync (no getItemAt here)
      var paras = ctx.document.body.paragraphs;
      paras.load("text, style");
      return ctx.sync().then(function () {
        var p = paras.items[index];
        if (!p) throw new FriendlyError("Paragraph " + index + " does not exist");
        p.insertText(String(text), "Replace");
        return ctx.sync().then(function () { return { ok: true }; });
      });
    });
  };

  WordDeck.prototype.replaceText = function (find, replace, all) {
    return this.run("replaceText", function (ctx) {
      var results = ctx.document.body.search(String(find), { matchCase: false, matchWholeWord: false });
      results.load("items");
      return ctx.sync().then(function () {
        var items = all ? results.items : (results.items.length ? [results.items[0]] : []);
        items.forEach(function (r) { r.insertText(String(replace), "Replace"); });
        return ctx.sync().then(function () { return { ok: true, hits: items.length }; });
      });
    });
  };

  WordDeck.prototype.insertTable = function (op) {
    if (!this.capabilities().tables) {
      return Promise.reject(new FriendlyError("Tables need WordApi 1.3 - this Word build does not support it (update Microsoft 365)."));
    }
    return this.run("insertTable", function (ctx) {
      var vals = op.values;
      ctx.document.body.insertTable(vals.length, vals[0].length, "End", vals);
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  WordDeck.prototype.insertComment = function (find, text) {
    if (!this.capabilities().comments) {
      // graceful: hand back the text for a manual paste
      return Promise.resolve({ fallback: true, reason: "comments need WordApi 1.4" });
    }
    return this.run("insertComment", function (ctx) {
      var results = ctx.document.body.search(String(find), { matchCase: false });
      results.load("items");
      return ctx.sync().then(function () {
        if (!results.items.length) return { ok: true, hits: 0 };
        results.items[0].insertComment(String(text));
        return ctx.sync().then(function () { return { ok: true, hits: 1 }; });
      });
    });
  };

  WordDeck.prototype.deleteParagraph = function (index) {
    return this.run("deleteParagraph", function (ctx) {
      var paras = ctx.document.body.paragraphs;
      paras.load("text");
      return ctx.sync().then(function () {
        var p = paras.items[index];
        if (!p) throw new FriendlyError("Paragraph " + index + " does not exist");
        p.delete();
        return ctx.sync().then(function () { return { ok: true }; });
      });
    });
  };

  WordDeck.prototype.formatText = function (index, font) {
    return this.run("formatText", function (ctx) {
      var para = ctx.document.body.paragraphs.getItemAt(index);
      var range = para.getAsRange
        ? para.getAsRange()
        : para.getRange("Whole");
      var rf = range.font;
      if (font.bold != null) rf.bold = font.bold;
      if (font.italic != null) rf.italic = font.italic;
      if (font.allCaps != null) rf.allCaps = font.allCaps;
      if (font.underline != null) rf.underline = font.underline ? "Single" : "None";
      if (font.color) rf.color = font.color;
      if (font.name) rf.name = font.name;
      if (font.size) rf.size = font.size;
      if (font.highlightColor) rf.highlightColor = font.highlightColor;
      if (font.align) para.alignment = { left: "Left", center: "Center", right: "Right", justify: "Justified" }[font.align];
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  WordDeck.prototype.createTemplateBlock = function (title, fields) {
    return this.run("createTemplateBlock", function (ctx) {
      if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) {
        throw new FriendlyError("content controls need WordApi 1.5 - update Word (16.70+ on Mac).");
      }
      var body = ctx.document.body;
      body.insertParagraph(title, "End").styleBuiltIn = "Heading2";
      var notes = [];
      fields.forEach(function (f) {
        var para = body.insertParagraph(f.label + ": ", "End");
        var cc = para.getRange("Whole").insertContentControl("richText");
        cc.title = f.label;
        cc.tag = "officekick-" + f.type;
        var typed = "";
        if (f.type === "text") { try { cc.type = "plainText"; typed = "plainText"; } catch (e2) { typed = "richText"; } }
        else if (f.type === "checkbox") { try { cc.type = "checkBox"; typed = "checkBox"; } catch (e3) { typed = "richText"; notes.push(f.label + ": kept as text (checkbox not supported by this build)"); } }
        else if (f.type === "date") { try { cc.type = "datePicker"; typed = "datePicker"; } catch (e4) { typed = "richText"; notes.push(f.label + ": kept as text (date picker not supported by this build)"); } }
        else if (f.type === "dropdown") {
          try {
            cc.type = "dropDownList"; typed = "dropDownList";
            (f.options || []).forEach(function (o) {
              try {
                if (cc.dropDownListContentControl && cc.dropDownListContentControl.addListItem) {
                  cc.dropDownListContentControl.addListItem({ displayText: o, value: o });
                } else { cc.placeholderText = o; }
              } catch (e5) { notes.push(f.label + ": option '" + o + "' rejected by Word"); }
            });
          } catch (e6) { typed = "richText"; notes.push(f.label + ": kept as text (dropdown not supported by this build)"); }
        }
        void typed;
      });
      return ctx.sync().then(function () { return { ok: true, notes: notes }; });
    });
  };

  WordDeck.prototype.setTrackChanges = function (on) {
    return this.run("setTrackChanges", function (ctx) {
      var doc = ctx.document;
      if (typeof doc.changeTrackingMode === "undefined") {
        throw new FriendlyError("this Word build does not support change-tracking mode - edits will be applied directly.");
      }
      doc.changeTrackingMode = on ? "trackAll" : "off";
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  function q34() { return String.fromCharCode(34); }
  WordDeck.prototype.formatTextFind = function (index, findText, font) {
    return this.run("formatTextFind", function (ctx) {
      var para = ctx.document.body.paragraphs.getItemAt(index);
      var found = para.search(String(findText), {});
      found.load("items/font");
      return ctx.sync().then(function () {
        var items = found.items || [];
        if (!items.length) throw new FriendlyError(q34() + findText + q34() + " not found in that paragraph.");
        items.forEach(function (rng) {
          if (font.bold != null) rng.font.bold = font.bold;
          if (font.italic != null) rng.font.italic = font.italic;
          if (font.allCaps != null) rng.font.allCaps = font.allCaps;
          if (font.underline != null) rng.font.underline = font.underline;
          if (font.color) rng.font.color = font.color;
          if (font.name) rng.font.name = font.name;
          if (font.size) rng.font.size = font.size;
          if (font.highlightColor) rng.font.highlightColor = font.highlightColor;
        });
        return ctx.sync().then(function () { return { ok: true, count: items.length }; });
      });
    });
  };

  WordDeck.prototype.addImage = function (payload) {
    return this.run("addImage", function (ctx) {
      if (payload.atCursor) {
        // WordApi 1.2: insert at the caret / selection
        ctx.document.getSelection().insertInlinePictureFromBase64(String(payload.b64), "Replace");
      } else {
        ctx.document.body.insertInlinePictureFromBase64(String(payload.b64), "End");
      }
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  WordDeck.prototype.rewriteDoc = function (op) {
    return this.run("rewriteDoc", function (ctx) {
      var body = ctx.document.body;
      body.clear();
      var last = null;
      (op.items || []).forEach(function (it) {
        last = body.insertParagraph(String(it.text == null ? "" : it.text), last ? "After" : "End");
        var st = styleNameOf(it.style);
        if (st && last) last.style = st;
      });
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  WordDeck.prototype.renderDebug = function (el) {
    var self = this;
    this.getContext().then(function (ctx) {
      var rows = [];
      rows.push("<h3>Simulator document</h3>");
      rows.push('<div class="opnote">' + ctx.paragraphCount + " paragraphs - " +
        (ctx.tables || []).length + " table(s) - " + (ctx.comCount || 0) + " comment(s)</div>");
      (ctx.paragraphs || []).forEach(function (p) {
        rows.push('<div class="simrow"><b>[' + p.idx + "]</b> <i>" + esc(p.style) + "</i> " + esc(String(p.text).slice(0, 120)) + "</div>");
      });
      (ctx.tables || []).forEach(function (t, i) {
        rows.push('<div class="simrow"><b>table ' + (i + 1) + "</b>: " +
          (t.values || []).map(function (r) { return r.join(" | "); }).join(" // ") + "</div>");
      });
      el.innerHTML = rows.join("");
    }, function (err) {
      el.innerHTML = '<div class="status error">' + esc(err.message) + "</div>";
    });
  };

  function styleNameOf(kind) {
    if (!kind || String(kind).toLowerCase() === "normal") return null; // doc default
    var k = String(kind).toLowerCase();
    if (k === "title") return "Title";
    if (k === "heading1") return "Heading 1";
    if (k === "heading2") return "Heading 2";
    if (k === "heading3") return "Heading 3";
    if (k === "quote") return "Quote";
    if (k === "listparagraph") return "List Paragraph";
    return null;
  }

    // ------------------------------------------------------------------ //
  // Word ops validation + formatting
  // ------------------------------------------------------------------ //
  function approvalMode(opType) {
    return SK.sessionApprovals[opType] || (SK.cfg && SK.cfg.approvals && SK.cfg.approvals[opType]) || "confirm";
  }

  function slideOkF(idx, ctx) {
    return typeof idx === "number" && ctx && ctx.paragraphs && ctx.paragraphs[idx - 1];
  }
function validateOps(rawOps, ctx) {
    var ops = [], invalid = [];
    (rawOps || []).forEach(function (op, index) {
      if (!op || typeof op !== "object") { invalid.push({ op: op, reason: "not an object" }); return; }
      var t = op.type;
      if (OP_TYPES.indexOf(t) < 0) { invalid.push({ op: op, reason: "unknown op type '" + t + "'" }); return; }
      if (!ctx && t !== "insert_paragraphs" && t !== "create_table") {
        invalid.push({ op: op, reason: "needs document context (enable 'include document context')" });
        return;
      }
      var bad = function (reason) { invalid.push({ op: op, reason: reason }); };
      var paraOk = function (v) { return typeof v === "number" && Math.floor(v) === v && v >= 1 && (!ctx || v <= ctx.paragraphCount); };
      var out = { type: t };
      if (t === "insert_paragraphs") {
        var items = Array.isArray(op.items) ? op.items : null;
        if (!items) { bad("missing items"); return; }
        var itemsOk = [];
        for (var i2 = 0; i2 < items.length; i2++) {
          var it = items[i2];
          if (!it || (typeof it.text !== "string" && it.text != null)) { bad("items must have string text"); return; }
          itemsOk.push({ text: it.text == null ? "" : String(it.text), style: it.style || "normal" });
        }
        if (!itemsOk.length) { bad("empty items"); return; }
        out.items = itemsOk;
      } else if (t === "edit_paragraph") {
        if (!paraOk(op.para)) { bad("bad paragraph index"); return; }
        if (typeof op.text !== "string") { bad("missing text"); return; }
        out.para = op.para; out.text = op.text;
      } else if (t === "replace_text") {
        if (typeof op.find !== "string" || op.find.length < 2) { bad("find must be at least 2 chars"); return; }
        if (typeof op.replace !== "string") { bad("missing replace"); return; }
        out.find = op.find; out.replace = op.replace; out.all = op.all !== false;
      } else if (t === "create_table") {
        var vals = Array.isArray(op.values) ? op.values : null;
        if (!vals) { bad("missing values (2D array)"); return; }
        var rowsT = [], cols = 0;
        for (var vi = 0; vi < vals.length; vi++) {
          var rowT = vals[vi];
          if (!Array.isArray(rowT)) { bad("values must be an array of arrays"); return; }
          if (vi === 0) cols = rowT.length;
          if (rowT.length !== cols || cols < 1) { bad("values rows must be rectangular"); return; }
          rowsT.push(rowT.map(function (c) { return c == null ? "" : String(c); }));
        }
        out.values = rowsT;
      } else if (t === "insert_comment") {
        if (typeof op.find !== "string" || op.find.length < 2) { bad("find must be at least 2 chars"); return; }
        if (typeof op.text !== "string" || !op.text) { bad("missing comment text"); return; }
        out.find = op.find; out.text = op.text;
      } else if (t === "delete_paragraph") {
        if (!paraOk(op.para)) { bad("bad paragraph index"); return; }
        out.para = op.para;
      } else if (t === "rewrite_doc") {
        var items2 = Array.isArray(op.items) ? op.items : (Array.isArray(op.paragraphs) ? op.paragraphs : null);
        if (!items2) { bad("missing items"); return; }
        var items3 = [];
        for (var i3 = 0; i3 < items2.length; i3++) {
          var it3 = items2[i3];
          items3.push({ text: (it3 && it3.text != null) ? String(it3.text) : "", style: (it3 && it3.style) || "normal" });
        }
        out.items = items3;
      } else if (t === "create_template") {
        if (typeof op.title !== "string" || !op.title.trim()) { bad("template needs a title"); return; }
        var fields = Array.isArray(op.fields) ? op.fields : null;
        if (!fields || !fields.length || fields.length > 10) { bad("1-10 fields required"); return; }
        var clean = [];
        for (var fi = 0; fi < fields.length; fi++) {
          var f = fields[fi] || {};
          if (typeof f.label !== "string" || !f.label.trim()) { bad("field " + (fi + 1) + " needs a label"); return; }
          var ft = f.type;
          if (ft !== "text" && ft !== "dropdown" && ft !== "checkbox" && ft !== "date") {
            bad("field '" + f.label + "': type must be text, dropdown, checkbox or date"); return;
          }
          var entry = { label: f.label.trim().slice(0, 100), type: ft };
          if (ft === "dropdown") {
            if (!Array.isArray(f.options) || !f.options.length || f.options.length > 20) { bad("dropdown '" + f.label + "' needs 1-20 options"); return; }
            entry.options = f.options.map(function (o) { return String(o).slice(0, 100); });
          }
          clean.push(entry);
        }
        out.title = op.title.trim().slice(0, 100); out.fields = clean;
      } else if (t === "format_text") {
        var paraOk = slideOkF(op.para, ctx);
        if (!paraOk) { bad("bad paragraph index"); return; }
        var fontW = {}, hasAnyW = false;
        ["bold", "italic", "allCaps", "underline"].forEach(function (k3) {
          if (op.font && typeof op.font[k3] === "boolean") { fontW[k3] = op.font[k3]; hasAnyW = true; }
        });
        if (op.font && typeof op.font.color === "string" && /^[\w#\(\), .%]+$/.test(op.font.color)) { fontW.color = op.font.color; hasAnyW = true; }
        if (op.font && typeof op.font.highlightColor === "string" && /^[\w#]+$/.test(op.font.highlightColor)) { fontW.highlightColor = op.font.highlightColor; hasAnyW = true; }
        if (op.font && typeof op.font.name === "string" && op.font.name) { fontW.name = op.font.name; hasAnyW = true; }
        if (op.font && typeof op.font.size === "number" && op.font.size >= 6 && op.font.size <= 400) { fontW.size = op.font.size; hasAnyW = true; }
        if (typeof op.align === "string" && /^(left|center|right|justify)$/i.test(op.align)) { fontW.align = op.align.toLowerCase(); hasAnyW = true; }
        if (!hasAnyW) { bad("no format changes requested"); return; }
        out.para = op.para; out.font = fontW;
        if (typeof op.find === "string" && op.find.trim()) {
          var findW = op.find.trim().slice(0, 200);
          var paraTxtW = (ctx.paragraphs[op.para - 1] || {}).text || "";
          if (paraTxtW.toLowerCase().indexOf(findW.toLowerCase()) === -1) { bad("find text not found in that paragraph"); return; }
          out.find = findW;
        }
      } else if (t === "insert_image") {
        if (typeof op.prompt !== "string" || op.prompt.trim().length < 3) { bad("missing image prompt"); return; }
        if (op.prompt.length > 4000) { bad("image prompt too long (max 4000 chars)"); return; }
        out.prompt = op.prompt.trim();
      }
      ops.push(out);
    });
    return { ops: ops, invalid: invalid, notesDropped: 0 };
  }

  var OP_EFFECT = {
    create_slide: "add", edit_slide: "edit", replace_text: "edit", set_notes: "edit",
    create_table: "add", create_shape: "add", create_image: "add", insert_image: "add",
    apply_layout: "edit", format_text: "edit", format_selection: "edit",
    move_shape: "edit", move_slide: "edit", set_background: "edit", add_hyperlink: "edit",
    set_alt_text: "edit", set_categories: "edit", draft_reply: "add", new_message: "add",
    compose_body: "edit", save_draft: "add", create_template: "add", web_lookup: "edit",
    delete_slide: "del", delete_shape: "del"
  };

  function fmtOp(op, ctx) {
    var paraRef = function (i) {
      if (!ctx || !ctx.paragraphs || !ctx.paragraphs[i - 1]) return "";
      var t = ctx.paragraphs[i].text;
      return t ? ' "' + String(t).slice(0, 60) + '"' : "";
    };
    switch (op.type) {
      case "insert_paragraphs":
        return "Insert " + op.items.length + " paragraph(s) -> end of document (" +
          op.items.map(function (i) { return i.style; }).filter(function (s, k, a) { return a.indexOf(s) === k; }).join(", ") + ")";
      case "edit_paragraph":
        return "Paragraph " + op.para + paraRef(op.para) + ": -> " + op.text.slice(0, 80);
      case "replace_text":
        return "'" + op.find.slice(0, 60) + "' -> '" + op.replace.slice(0, 60) + "'" + (op.all ? " (all)" : " (first)");
      case "create_table":
        return "Table " + op.values.length + "x" + op.values[0].length + " -> end of document";
      case "insert_comment":
        return "Comment on '" + op.find.slice(0, 50) + "' -> " + op.text.slice(0, 60);
      case "delete_paragraph":
        return "Delete paragraph " + op.para + paraRef(op.para);
      case "rewrite_doc":
        return "REWRITE document: " + op.items.length + " paragraph(s)";
      case "format_text": {
        var fp = [];
        Object.keys(op.font).forEach(function (kf2) {
          var v2 = op.font[kf2];
          fp.push(kf2 + "=" + (typeof v2 === "boolean" ? (v2 ? "on" : "off") : v2));
        });
        return "Paragraph " + op.para + paraRef(op.para) + ": " + fp.join(", ") +
          (op.find ? " (in " + q34() + op.find + q34() + ")" : "");
      }
      case "create_template":
        return "template '" + op.title + "' with " + op.fields.length + " field(s): " +
          op.fields.map(function (f) { return f.label + " (" + f.type + ")"; }).join(", ");
      case "insert_image": {
        var prI = op.prompt && op.prompt.length > 60 ? op.prompt.slice(0, 60) + "…" : (op.prompt || "");
        return "Image at end of document: '" + prI + "' (profile image model)";
      }
    }
    return op.type;
  }

function applyOps(deck, ops, ctx) {
    var lines = [], manualNotes = [];
    var deleted = [];
    var revNote = null;
    var revOn = !!(SK.cfg && SK.cfg.trackedChanges);
    var pre = Promise.resolve();
    if (revOn && deck.setTrackChanges) {
      pre = deck.setTrackChanges(true).catch(function (err) {
        revNote = (err && err.message) || String(err);
        SK.cfg.trackedChanges = false; // do not retry per-op
      });
    }
    var adj = function (i) {
      var d = 0;
      deleted.forEach(function (x) { if (x < i) d++; });
      return i - d;
    };
    var runOne = function (op) {
      switch (op.type) {
        case "insert_paragraphs":
          return deck.insertParagraphs(op).then(function () {
            lines.push({ ok: true, text: "insert " + op.items.length + " paragraph(s) -> end of document" });
          });
        case "edit_paragraph":
          return deck.editParagraph(adj(op.para - 1), op.text).then(function () {
            lines.push({ ok: true, text: "paragraph " + op.para + " edited" });
          });
        case "replace_text":
          return deck.replaceText(op.find, op.replace, op.all !== false).then(function (r) {
            lines.push({ ok: true, text: "replace '" + op.find.slice(0, 50) + "' -> " +
              (r && r.hits != null ? r.hits : "?") + " hit(s)" });
          });
        case "create_table":
          return deck.insertTable(op).then(function () {
            lines.push({ ok: true, text: "table " + op.values.length + "x" + op.values[0].length +
              " -> end of document" });
          });
        case "insert_comment":
          return deck.insertComment(op.find, op.text).then(function (r) {
            if (r && r.fallback) {
              manualNotes.push({ slide: 0, text: "COMMENT on '" + op.find + "': " + op.text });
              lines.push({ ok: true, text: "comment needs manual paste (" + r.reason + ")" });
            } else {
              lines.push({ ok: true, text: "comment added on '" + op.find.slice(0, 40) + "'" });
            }
          });
        case "delete_paragraph":
          var i2 = adj(op.para - 1);
          return deck.deleteParagraph(i2).then(function () {
            deleted.push(op.para - 1);
            lines.push({ ok: true, text: "paragraph " + op.para + " deleted" });
          });
        case "rewrite_doc":
          return deck.rewriteDoc(op).then(function () {
            lines.push({ ok: true, text: "document rewritten: " + op.items.length + " paragraph(s)" });
          });
        case "format_text": {
          if (op.find) {
            return deck.formatTextFind(adj(op.para - 1), op.find, op.font).then(function () {
              lines.push({ ok: true, text: "format paragraph " + op.para + " (" + op.find + "): " +
                Object.keys(op.font).join(", ") });
            });
          }
          return deck.formatText(adj(op.para - 1), op.font).then(function () {
            lines.push({ ok: true, text: "format paragraph " + op.para + ": " + Object.keys(op.font).join(", ") });
          });
        }
        case "create_template":
          return deck.createTemplateBlock(op.title, op.fields).then(function (r) {
            var extra = (r && r.notes && r.notes.length) ? " (notes: " + r.notes.join("; ") + ")" : "";
            lines.push({ ok: true, text: "template '" + op.title + "' appended: " +
              op.fields.map(function (f) { return f.label + " (" + f.type + ")"; }).join(", ") + extra });
          });
        case "insert_image": {
          var prof = currentProfile();
          var imgRef = prof && prof.image;
          if (!imgRef || !imgRef.connection || !imgRef.model) {
            throw new FriendlyError(
              "No image model on profile '" + (prof ? prof.name : "?") +
              "' - set one in Profiles (connection + model with an images API).");
          }
          setStatus("Generating image… (often 30-90 s depending on the model)", "busy");
          return Promise.race([
            apiJson("/api/image", "POST", {
              connection: imgRef.connection, model: imgRef.model, prompt: op.prompt,
            }),
            new Promise(function (_, rej) {
              setTimeout(function () {
                rej(new FriendlyError("image generation timed out after 300 s (no answer from '" + imgRef.model + "')"));
              }, 300000);
            }),
          ]).then(function (img) {
            if (!img || !img.b64) throw new FriendlyError("server returned no image data");
            return deck.addImage({ b64: img.b64, atCursor: $("chat-cursor") && $("chat-cursor").checked }).then(function () {
              var atCursor = $("chat-cursor") && $("chat-cursor").checked;
              lines.push({ ok: true, text: "image inserted " + (atCursor ? "at cursor/selection" : "at end of document") +
                " (" + imgRef.connection + ":" + imgRef.model + ")" });
            });
          });
        }
      }
      return Promise.resolve();
    };
    var chain = Promise.resolve();
    ops.forEach(function (op) {
      chain = chain.then(function () { return runOne(op); }, function (err) {
        lines.push({ ok: false, text: fmtOp(op, ctx) + " - " + ((err && err.message) || String(err)) });
        return Promise.resolve();
      });
    });
    return pre.then(function () { return chain; }).then(function () {
      var res = { lines: lines, manualNotes: manualNotes };
      if (revNote) lines.push({ ok: false, text: "revisions: " + revNote });
      var off = (SK.cfg && SK.cfg.trackedChanges && deck.setTrackChanges)
        ? deck.setTrackChanges(false).catch(function () { return null; })
        : Promise.resolve(null);
      return off.then(function () { return res; });
    });
  }

  // ------------------------------------------------------------------ //
  // Chat / prompt flow
  // ------------------------------------------------------------------ //
  function formatContext(ctx) {
    var lines = [];
    lines.push("DOCUMENT OUTLINE: " + ctx.paragraphCount + " paragraphs, " +
      (ctx.tables || []).length + " table(s), host: " + ctx.host);
    if (ctx.selected && ctx.selected.length) {
      lines.push('CURRENT SELECTION: "' + String(ctx.selected).slice(0, 200) + '"');
    }
    (ctx.paragraphs || []).slice(0, 400).forEach(function (p) {
      var t = String(p.text || "").replace(/\r|\n/g, " ");
      if (t.length > 180) t = t.slice(0, 180) + "…";
      if (!t.trim()) return;
      lines.push("    para " + (p.idx + 1) + " (" + p.style + '): "' + t + '"');
    });
    (ctx.tables || []).forEach(function (t) {
      lines.push("    table " + (t.idx + 1) + ": " +
        (t.values || []).slice(0, 8).map(function (r) { return r.join(" | "); }).join(" // "));
    });
    var total = lines.join("\n");
    if (total.length > 20000) total = total.slice(0, 20000) + "\n…[truncated]";
    return total;
  }

  function currentProfile() {
    var name = $("chat-profile").value;
    if (!SK.cfg) return null;
    var prof = (SK.cfg.profiles || []).find(function (p) { return p.name === name; });
    if (prof) return prof;
    var def = SK.cfg.defaultProfile;
    return (SK.cfg.profiles || []).find(function (p) { return p.name === def; }) || (SK.cfg.profiles || [])[0] || null;
  }

  function setRunEnabled(on) {
    $("chat-run").disabled = !on || SK.busy;
  }

  function runPrompt() {
    if (SK.busy) return;
    var input = $("chat-input");
    var instruction = input.value.trim();
    if (!instruction) return;
    var prof = currentProfile();
    if (!prof) { setStatus("No profile yet - create one on the Profiles tab first.", "error"); switchTab("settings"); switchSettingsSub("profiles"); return; }
    var route = $("chat-route").value;
    var ref = routeToModel(prof, route, instruction);
    if (!ref || !ref.connection || !ref.model) {
      setStatus("Profile '" + prof.name + "' has no model assigned for this route. Open Profiles and pick connection + model.", "error");
      return;
    }
    filePreviousRun();  // the previous request/response moves into the history box NOW
    SK.busy = true;
    input.disabled = true;   // the text stays visible but locked while working
    setRunEnabled(false);
    chatTimerStart();
    var wantCtx = $("chat-context").checked;
    setStatus("Collecting document context…", "busy");
    var ctxPromise = wantCtx ? SK.deck.getContext({}).catch(function (err) {
      setStatus("Could not read the deck (" + err.message + ") - continuing without context.", "error");
      return null;
    }) : Promise.resolve(null);

    ctxPromise.then(function (ctx) {
      SK.lastCtx = ctx;
      SK.webResults = "";
      var webPromise = Promise.resolve("");
      return webPromise.then(function () {
      var sys = SYSTEM_PROMPT + "\nHost: " + (SK.sim ? "DocKick simulator" : SK.deck.capabilities().host);
            var user = (ctx ? formatContext(ctx) : "(no deck context requested)") + "\n\nUSER REQUEST:\n" + instruction;
      if (SK.webResults) user += "\n\nWEB RESULTS (live lookup, cite them when relevant):\n" + SK.webResults;
      var messages = [
        { role: "system", content: sys },
      ].concat(SK.history.slice(-(HISTORY_MAX_EXCHANGES * 2)))
       .concat([{ role: "user", content: user }]);
      dbg("prompt", { profile: prof.name, route: route, model: ref.model, connection: ref.connection, instruction: instruction, hadContext: !!ctx, messages: messages });
      setStatus("Asking " + ref.model + " (via " + ref.connection + ")…", "busy");
      var askedLookup = false;
      function mainAsk() {
        return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: messages }); })
          .then(function (resp) {
            var pre = tryParse(resp.text);
            if (pre.ok && pre.kind === "web_lookup" && !askedLookup) {
              askedLookup = true;
              setStatus("Doing a live web lookup through the model…", "busy");
              return webLookup(prof, instruction).then(function (ws) {
                if (!ws) {
                  return { ok: true, kind: "reply",
                           text: "No web-enabled model available for a live lookup (the model reported no web access). Re-phrase or check the profile/connection." };
                }
                SK.webResults = ws;
                messages = messages.concat([{ role: "assistant", content: pre.text || '{"kind":"web_lookup"}' },
                                            { role: "user", content: "WEB LOOKUP RESULTS:\n" + ws +
                                              "\n\nAnswer the ORIGINAL USER REQUEST now using these results (cite sources); the required JSON format rules apply again." }]);
                setStatus("Answering with the fresh web results…", "busy");
                return mainAsk2();
              }).catch(function (e) {
                return { ok: true, kind: "reply", text: "The live web lookup failed: " + ((e && e.message) || e) };
              });
            }
            return resp;
          });
      }
      function mainAsk2() {
        return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: messages }); });
      }
      return mainAsk().then(function (resp) {
        dbg("response", { text: resp.text, elapsedMs: resp.elapsedMs });
        SK.lastModelText = { text: resp.text };
        var parsed = tryParse(resp.text);
        if (!parsed.ok) {
          setStatus("Model output unclear - silently re-asking for clean JSON…", "busy");
          var repair = messages.concat([
            { role: "assistant", content: resp.text },
            { role: "user", content: REPAIR_ASK },
          ]);
          dbg("repair.request", {});
          return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: repair }); }).then(function (resp2) {
            dbg("repair.response", { text: resp2.text, elapsedMs: resp2.elapsedMs });
            SK.lastModelText = { text: resp2.text };
            return tryParse(resp2.text);
          });
        }
        return parsed;
      }).then(function (parsed) {
        renderParsed(parsed, ctx, instruction);
        pushHistory(instruction, assistantHistoryText(parsed));
        });
      });
    }).catch(function (err) {
      var msg = err && err.friendly ? err.message : (err && err.message) || String(err);
      setStatus(msg, "error");
      // the request left the pane but failed - keep it in the session too
      pushHistory(instruction, "(failed: " + msg + ")");
      // give the request back so the user can just hit Run again
      if (input && !input.value) input.value = instruction;
    }).then(function () {
      stopChatTimer();
      SK.busy = false;
      setRunEnabled(true);
      input.disabled = false;
      if ($( "chat-output" ) && $("chat-output").children.length) input.value = "";
    });
  }

  function tryParse(text) {
    var stripped = stripReasoning(text);
    if (!stripped) return { ok: false, text: text };
    var obj = extractObject(stripped);
    if (obj && typeof obj === "object") {
      if (obj.kind === "reply" && typeof obj.text === "string") return { ok: true, kind: "reply", text: obj.text };
      if (obj.kind === "slide_ops" && Array.isArray(obj.ops)) return { ok: true, kind: "slide_ops", ops: obj.ops };
      if (Array.isArray(obj.ops)) return { ok: true, kind: "slide_ops", ops: obj.ops }; // tolerate missing kind
    }
    return { ok: false, text: stripped };
  }



  // ---- auto-retry for model asks: 5 attempts, 10 s apart, on transport
  // errors/server hiccups only (4xx business errors fail immediately) ----
  function askRetryable(askFn, attempt) {
    attempt = attempt || 1;
    return askFn().catch(function (err) {
      var status = err && err.status;
      var retries_429 = status === 429;
      var serverish = !status || status >= 500 || retries_429;
      if (!serverish || attempt >= 5) throw err;
      var secs = 10;
      setStatus("Attempt " + attempt + "/5 failed (" + ((err && err.message) || err) +
        ") - retrying in " + secs + " s…", "busy");
      return new Promise(function (res) { setTimeout(res, secs * 1000); }, function () {}).then(function () {
        return askRetryable(askFn, attempt + 1);
      });
    });
  }

  function chatTimerStart() {
    stopChatTimer();
    var el = $("chat-timer");
    if (!el) return;
    SK.chatTimer = { iv: null, t0: Date.now(), el: el };
    function tick() {
      var t = Math.floor((Date.now() - SK.chatTimer.t0) / 1000);
      var h = Math.floor(t / 3600), m = Math.floor((t % 3600) / 60), s2 = t % 60;
      var mm = (m < 10 ? "0" : "") + m;
      var ss = (s2 < 10 ? "0" : "") + s2;
      el.textContent = (h ? h + "h " : "") + mm + "min " + ss + "sec(s)";
    }
    tick();
    SK.chatTimer.iv = setInterval(tick, 500);
  }

  function stopChatTimer() {
    if (SK.chatTimer && SK.chatTimer.iv) { clearInterval(SK.chatTimer.iv); SK.chatTimer.iv = null; }
    // the final elapsed time stays visible (frozen) until the next run
  }



  function pmGlyph(open) { return open ? "\u2212" : "+"; }   // boxed minus / plus

  function pastTurnRow(u, a) {
    // one collapsible turn box: header = +/- and timestamp + the request
    // snippet; inside = the full request (clickable to refill the input)
    // and the response/gate summary
    var grp = document.createElement("div");
    grp.className = "reqgrp pastturn";
    var head = document.createElement("div");
    head.className = "reqhead2";
    head.title = "Click to open/close this request/response";
    var pm = document.createElement("span");
    pm.className = "pmbox";
    pm.textContent = "+";
    head.appendChild(pm);
    var uc = String(u.content || "");
    var ts1 = document.createElement("span");
    ts1.className = "ptts";
    ts1.textContent = u.ts || "no timestamp";   // timestamp first
    head.appendChild(ts1);
    var snip = document.createElement("div");
    snip.className = "pt-text";                // the request text BELOW the ts
    snip.textContent = uc.length > 50 ? uc.slice(0, 50) + "…" : uc;
    var body = document.createElement("div");
    body.className = "reqbody";
    var q1 = document.createElement("div");
    q1.className = "pt-req";
    q1.style.cursor = "pointer";
    q1.title = "Click to load this request into the input box (only when the box is empty)";
    var b1 = document.createElement("b");
    b1.textContent = "Request: ";
    q1.appendChild(b1);
    var fullReq = uc;
    q1.appendChild(document.createTextNode(fullReq));
    q1.addEventListener("click", function () { copyRequestToInput(fullReq); });
    body.appendChild(q1);
    var q2 = document.createElement("div");
    q2.className = "pt-res";
    var b2 = document.createElement("b");
    b2.textContent = "Response: ";
    q2.appendChild(b2);
    var ac = String(a.content || "");
    q2.appendChild(document.createTextNode(ac.length > 400 ? ac.slice(0, 400) + "…" : ac));
    body.appendChild(q2);
    head.addEventListener("click", function () {
      var isOpen = body.classList.toggle("open");
      pm.textContent = pmGlyph(isOpen);
    });
    grp.appendChild(head);
    grp.appendChild(snip);
    grp.appendChild(body);
    return grp;
  }

  function updatePastBox() {
    // the box title is just a LABEL: the container shows every run box of
    // the session, open by default; only the individual boxes have +/-
    var box = $("pastbox"), body = $("past-body"), head = $("past-head");
    if (!box || !body || !head) return;
    var n = body.querySelectorAll(".pastturn").length;
    head.textContent = "Previous requests (" + n + ")";
    head.style.cursor = "default";
    body.classList.add("open");   // the container list itself stays open
    box.style.display = n ? "" : "none";
  }

  // On Run: the PREVIOUS run (request + its reply / proposed-actions text)
  // files itself into the box immediately and the top zone blanks, so only
  // the current run is between the prompt and the history box.
  function filePreviousRun() {
    var zone = $("chat-output");
    var body = $("past-body");
    if (!body || !$("pastbox")) return;
    if (zone && zone.childNodes.length && SK.lastInstruction) {
      var lastUser = null;
      for (var hi = SK.history.length - 1; hi >= 0; hi--) {
        if (SK.history[hi].role === "user") { lastUser = SK.history[hi]; break; }
      }
      var row = pastTurnRow(
        { content: SK.lastInstruction, ts: (lastUser && lastUser.ts) || "" },
        { content: zone.textContent || "", ts: (lastUser && lastUser.ts) || "" });
      body.insertBefore(row, body.firstChild);   // newest on top
      while (zone.firstChild) zone.removeChild(zone.firstChild);
    }
    updatePastBox();
  }

  function wirePastHeader() {
    // the title is not clickable; the per-run boxes carry the +/- toggles
    var head = $("past-head");
    if (!head) return;
    head.onclick = null;
    updatePastBox();
  }

  function renderParsed(parsed, ctx, instruction) {
    SK.lastInstruction = instruction;
    var out = $("chat-output");
    out.innerHTML = "";
    // the request prints above the response/summary (echoed by the clear)
    var reqEcho = document.createElement("div");
    reqEcho.className = "reqecho";
    reqEcho.textContent = instruction;
    out.appendChild(reqEcho);
    wirePastHeader();
    if (!parsed.ok) {
      var box = document.createElement("div");
      box.className = "reply-text";
      box.textContent = parsed.text || "(empty response)";
      out.appendChild(box);
      var note = document.createElement("div");
      note.className = "opnote";
      note.textContent = "Shown as text: the model did not return a valid action JSON. Enable debug logging (Settings) to inspect the raw output.";
      out.appendChild(note);
      setStatus("Done (reply as text).", "ok");
      return;
    }
    if (parsed.kind === "reply") {
      var box2 = document.createElement("div");
      box2.className = "reply-text";
      box2.textContent = parsed.text;
      out.appendChild(box2);
      setStatus("Done.", "ok");
      return;
    }
    var validation = validateOps(parsed.ops, ctx);
    if (!validation.ops.length && !validation.invalid.length) {
      var box3 = document.createElement("div");
      box3.className = "reply-text";
      box3.textContent = "(model returned an empty action list)";
      out.appendChild(box3);
      setStatus("Done.", "ok");
      return;
    }
    // auto-apply: skip the preview entirely when every action's policy is 'auto'
    // and nothing is invalid/blocked - approvals still halt the run otherwise
    var allAuto = validation.invalid.length === 0 &&
      validation.ops.length > 0 &&
      validation.ops.every(function (op) { return approvalMode(op.type) === "auto"; });
    if ($("chat-autoapply").checked && allAuto && SK.deck.capabilities().text) {
      dbg("ops.autoapply", { count: validation.ops.length });
      var autoBox = document.createElement("div");
      autoBox.className = "opsbox";
      var autoH = document.createElement("h3");
      autoH.textContent = "Auto-applied " + validation.ops.length + " action(s) (auto-apply)" +
        (validation.notesDropped ? " - " + validation.notesDropped + " notes action(s) skipped (setting off)" : "");
      autoBox.appendChild(autoH);
      var autoResult = document.createElement("div");
      autoBox.appendChild(autoResult);
      out.appendChild(autoBox);
      setStatus("Auto-apply: running " + validation.ops.length + " action(s)…", "busy");
      applyOps(SK.deck, validation.ops, ctx).then(function (res) {
        renderApplyResult(autoResult, res);
        var fails = res.lines.filter(function (l) { return !l.ok; }).length;
        setStatus(fails ? "Auto-applied with " + fails + " error(s)." : "Auto-applied " + validation.ops.length + " action(s).", fails ? "error" : "ok");
        renderSimDeck();
        updateDeckSummary();
        dbg("ops.applied", { lines: res.lines, autoapply: true });
      }).catch(function (err) {
        setStatus((err && err.message) || String(err), "error");
      });
      return;
    }
    renderOpsPreview(out, validation, ctx);
  }

  // ------------------------------------------------------------------ //
  // Ops preview + approval gate
  // ------------------------------------------------------------------ //
  function renderApplyResult(result, res) {
    result.innerHTML = "";
    res.lines.forEach(function (l) {
      var d = document.createElement("div");
      d.className = "resultline " + (l.ok ? "ok" : "fail");
      d.textContent = l.text;
      result.appendChild(d);
    });
    if (res.manualNotes.length) {
      var mb = document.createElement("div");
      mb.className = "manualbox";
      mb.appendChild(document.createElement("h4")).textContent = "DocKick could not write this automatically - copy manually:";
      res.manualNotes.forEach(function (mn) {
        var ta = document.createElement("textarea");
        ta.value = mn.text;
        mb.appendChild(ta);
      });
      result.appendChild(mb);
    }
  }

  function renderOpsPreview(out, validation, ctx) {
    var caps = SK.deck.capabilities();
    var box = document.createElement("div");
    box.className = "opsbox";
    var h = document.createElement("h3");
    syncAutoApplyCheckbox();
    h.textContent = "Proposed actions" + ($("chat-autoapply").checked ? " (auto-apply)" : "");
    box.appendChild(h);


    if (!caps.text) {
      var warn = document.createElement("div");
      warn.className = "status error";
      warn.textContent = "This Word build lacks the required Word.js APIs. DocKick can read but not modify. Update Microsoft 365 to enable edits.";
      box.appendChild(warn);
    }

    var confirmTypes = {};
    var rows = [];
    validation.ops.forEach(function (op) {
      var mode = approvalMode(op.type);
      rows.push({ op: op, mode: mode });
      if (mode === "confirm") confirmTypes[op.type] = true;
    });
    validation.invalid.forEach(function (entry) {
      rows.push({ op: entry.op || { type: "?" }, mode: "invalid", reason: entry.reason });
    });

    var tableRows = document.createElement("div");
    rows.forEach(function (row) {
      var r = document.createElement("div");
      r.className = "oprow " + (OP_EFFECT[row.op.type] === "del" ? "op-del" : "op-mod") +
        (row.mode === "deny" ? " blocked" : row.mode === "invalid" ? " invalid" : "");
      var badge = document.createElement("span");
      badge.className = "optype " + row.op.type;
      badge.textContent = (OPS_META[row.op.type] || { label: row.op.type }).label;
      var desc = document.createElement("span");
      desc.className = "opdesc";
      var text = row.mode === "deny" ? fmtOp(row.op, ctx) + "  - blocked by approval policy"
        : row.mode === "invalid" ? (row.op && row.op.type ? row.op.type + ": " : "") + (row.reason || "invalid")
        : fmtOp(row.op, ctx);
      desc.textContent = text;
      r.appendChild(badge);
      r.appendChild(desc);
      if (row.mode === "confirm") {
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.className = "opcheckbox";
        cb.checked = SK.sessionApprovals[row.op.type] === "auto"; // remembered => pre-approved
        cb.dataset.optype = row.op.type;
        var lab = document.createElement("span");
        lab.className = "opnote";
        lab.textContent = cb.checked ? "(approved for this session)" : "(needs approval)";
        cb.addEventListener("change", function () {
          lab.textContent = cb.checked ? (SK.sessionApprovals[row.op.type] === "auto" ? "(approved for this session)" : "(approved once)") : "(needs approval)";
          updateApplyLabel();
        });
        r.appendChild(cb);
        r.appendChild(lab);
        row._cb = cb;
      } else if (row.mode === "auto") {
        var note2 = document.createElement("span");
        note2.className = "opnote";
        note2.textContent = "auto";
        r.appendChild(note2);
      } else if (row.mode === "deny") {
        var note3 = document.createElement("span");
        note3.className = "opnote";
        note3.textContent = "denied";
        r.appendChild(note3);
      } else {
        var note4 = document.createElement("span");
        note4.className = "opnote";
        note4.textContent = row.reason || "";
        r.appendChild(note4);
      }
      tableRows.appendChild(r);
    });
    box.appendChild(tableRows);

    // remember-for-session checkboxes (one per confirm type present)
    var remember = {};
    Object.keys(confirmTypes).forEach(function (type) {
      if (SK.sessionApprovals[type] === "auto") return; // already remembered
      var lab2 = document.createElement("label");
      lab2.className = "field-inline check";
      var cb2 = document.createElement("input");
      cb2.type = "checkbox";
      cb2.dataset.optype = type;
      remember[type] = cb2;
      lab2.appendChild(cb2);
      lab2.appendChild(document.createTextNode("don't ask again for '" + OPS_META[type].label + "' this session"));
      box.appendChild(lab2);
    });

    var autoOps = rows.filter(function (r) { return r.mode === "auto"; }).length;
    var confirmOps = rows.filter(function (r) { return r.mode === "confirm"; }).length;
    var denyOps = rows.filter(function (r) { return r.mode === "deny"; }).length;
    var invalidOps = validation.invalid.length;

    var btnrow = document.createElement("div");
    btnrow.className = "row";
    var applyBtn = document.createElement("button");
    applyBtn.className = "btn primary";
    var updateApplyLabel = function () {
      var approved = rows.filter(function (r) {
        return r.mode === "confirm" && r._cb && r._cb.checked;
      }).length;
      applyBtn.textContent = "Apply " + autoOps + " auto" +
        (confirmOps ? " + " + approved + "/" + confirmOps + " approved" : "");
      applyBtn.disabled = (autoOps + approved === 0) || !caps.text;
    };
    updateApplyLabel();
    if (confirmOps > 0) {
      var selectAllBtn = document.createElement("button");
      selectAllBtn.className = "btn subtle";
      selectAllBtn.textContent = "Select all (" + confirmOps + ")";
      selectAllBtn.disabled = !caps.text;
      selectAllBtn.addEventListener("click", function (ev) {
        ev.preventDefault();
        rows.forEach(function (r) {
          if (r.mode === "confirm" && r._cb && !r._cb.checked) {
            r._cb.checked = true;
            r._cb.dispatchEvent(new Event("change"));
          }
        });
        dbg("ops.selectall", { count: confirmOps });
      });
      btnrow.appendChild(selectAllBtn);
    }
    var discardBtn = document.createElement("button");
    discardBtn.className = "btn subtle";
    discardBtn.textContent = "Discard";
    btnrow.appendChild(applyBtn);
    btnrow.appendChild(discardBtn);
    if (denyOps || invalidOps || validation.notesDropped) {
      var info = document.createElement("span");
      info.className = "opnote";
      info.textContent = (denyOps ? denyOps + " blocked, " : "") + (invalidOps ? invalidOps + " invalid, " : "") +
        (validation.notesDropped ? validation.notesDropped + " notes action(s) skipped (Generate speaker notes is off), " : "") +
        (autoOps + denyOps + invalidOps + validation.notesDropped ? "rest not affected" : "");
      btnrow.appendChild(info);
    }
    box.appendChild(btnrow);

    var result = document.createElement("div");
    box.appendChild(result);

    discardBtn.addEventListener("click", function () {
      box.parentNode.removeChild(box);
      setStatus("Discarded proposed actions.", "");
      dbg("ops.discarded", {});
    });

    applyBtn.addEventListener("click", function () {
      var chosen = [];
      rows.forEach(function (row) {
        if (row.mode === "auto") chosen.push(row.op);
        else if (row.mode === "confirm" && row._cb && row._cb.checked) chosen.push(row.op);
      });
      Object.keys(remember).forEach(function (type) {
        if (remember[type].checked && chosen.some(function (op) { return op.type === type; })) {
          SK.sessionApprovals[type] = "auto";
        }
      });
      if (!chosen.length) { setStatus("Nothing approved - nothing applied.", ""); return; }
      dbg("ops.apply", { approved: chosen.map(function (op) { return op.type; }), count: chosen.length });
      applyBtn.disabled = true;
      setStatus("Applying " + chosen.length + " op(s)…", "busy");
      applyOps(SK.deck, chosen, SK.lastCtx || ctx).then(function (res) {
        result.innerHTML = "";
        res.lines.forEach(function (l) {
          var d = document.createElement("div");
          d.className = "resultline " + (l.ok ? "ok" : "fail");
          d.textContent = l.text;
          result.appendChild(d);
        });
        if (res.manualNotes.length) {
          var mb = document.createElement("div");
          mb.className = "manualbox";
          mb.appendChild(document.createElement("h4")).textContent = "Notes could not be written automatically - copy manually:";
          res.manualNotes.forEach(function (mn) {
            var ta = document.createElement("textarea");
            ta.value = mn.text;
            mb.appendChild(ta);
          });
          result.appendChild(mb);
        }
        var fails = res.lines.filter(function (l) { return !l.ok; }).length;
        setStatus(fails ? "Applied with " + fails + " error(s)." : "Applied " + chosen.length + " op(s).", fails ? "error" : "ok");
        renderSimDeck();
        updateDeckSummary();
        dbg("ops.applied", { lines: res.lines, manualNotes: res.manualNotes });
      });
    });

    out.appendChild(box);
    setStatus(confirmOps ? "Approval needed before applying." : "Review proposed actions.", "pending");
    dbg("ops.preview", { total: rows.length, auto: autoOps, confirm: confirmOps, deny: rows.filter(function (r) { return r.mode === "deny"; }).length, invalid: invalidOps });
  }

  // ------------------------------------------------------------------ //
  // Deck summary + sim renderer
  // ------------------------------------------------------------------ //
  function updateDeckSummary() {
    var el = $("deck-summary");
    if (!SK.deck) { el.textContent = ""; return; }
    SK.deck.getContext({ maxParas: 1 }).then(function (ctx) {
      var sel = ctx.selected && ctx.selected.length ? " (text selected)" : "";
      el.textContent = "Document: " + ctx.paragraphCount + " paragraph(s) - " + ctx.host + sel;
    }).catch(function (err) {
      el.textContent = "Document: unreadable (" + err.message + ")";
    });
  }

  function renderSimDeck() {
    if (!SK.sim) return;
    SK.deck.renderDebug($("sim-deck-view"));
  }

  // ------------------------------------------------------------------ //
  // Tabs
  // ------------------------------------------------------------------ //
  function switchSettingsSub(id) {
    var bar = document.getElementById("settings-subtabs");
    if (!bar) return;
    Array.prototype.forEach.call(bar.querySelectorAll("button[data-subtab]"), function (b) {
      b.classList.toggle("active", b.dataset.subtab === id);
    });
    Array.prototype.forEach.call(document.querySelectorAll("#tab-settings .subpanel"), function (p2) {
      p2.classList.toggle("active", p2.id === "sub-" + id);
    });
    if (id === "inspector" && typeof renderRequests === "function") renderRequests();
  }

  function switchTab(id) {
    SK.activeTab = id;
    var inp = $("chat-input");
    if (inp && !SK.busy) inp.disabled = false;   // returning to a sheet always frees the input
    if (id === "requests") renderRequests();
    if (id === "sessions") renderSessions();
    Array.prototype.forEach.call(document.querySelectorAll("#tabs button"), function (b) {
      b.classList.toggle("active", b.dataset.tab === id);
    });
    Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (s) {
      s.classList.toggle("active", s.id === "tab-" + id);
    });
  }

  // ------------------------------------------------------------------ //
  // Connections tab
  // ------------------------------------------------------------------ //
  function renderConnections() {
    var ts = $("conn-type");
    if (!ts.options.length) {
      ts.innerHTML = CONN_TYPES.map(function (t) {
        return '<option value="' + t.id + '">' + esc(t.label) + "</option>";
      }).join("");
    }
    var list = $("conn-list");
    var rows = [ "<tr><th>Name</th><th>Type</th><th>Base URL</th><th>Key</th></tr>" ];
    (SK.cfg.connections || []).forEach(function (c) {
      var used = (SK.cfg.profiles || []).filter(function (p) {
        return (p.content && p.content.connection === c.name) ||
               (p.other && p.other.connection === c.name);
      }).map(function (p) { return p.name; });
      // click the name to edit; deletion lives in the edit form (with confirm)
      var nameCell =
        '<button class="namebtn" data-edit="' + esc(c.name) + '" title="edit">' +
        esc(c.name) + "</button>" +
        (used.length ? '<br /><span class="opnote">in use: ' + esc(used.join(", ")) + "</span>" : "");
      rows.push(
        "<tr><td>" + nameCell + "</td><td>" + esc(c.type) + "</td><td>" + esc(c.baseUrl || "-") + "</td>" +
        "<td>" + (c.hasKey ? "set" : "—") + '</td><td>' +
        '<button class="btn" data-test="' + esc(c.name) + '">Test</button>' +
        '<span class="opnote" data-testmsg="' + esc(c.name) + '"></span></td></tr>'
      );
    });
    list.innerHTML = rows.join("") || "<tr><td>No connections yet - add one below.</td></tr>";
    Array.prototype.forEach.call(list.querySelectorAll("[data-edit]"), function (b) {
      b.addEventListener("click", function () { editConnection(b.dataset.edit); });
    });
    Array.prototype.forEach.call(list.querySelectorAll("[data-test]"), function (b) {
      b.addEventListener("click", function () {
        testConnRow(b.dataset.test, list.querySelector('[data-testmsg="' + b.dataset.test + '"]'));
      });
    });
  }

  function editConnection(name) {
    var c = (SK.cfg.connections || []).find(function (x) { return x.name === name; });
    if (!c) return;
    SK.editingConn = name;
    SK.connBaseDirty = true;
    $("conn-form-title").textContent = "Edit connection: " + name;
    $("conn-name").value = c.name;
    $("conn-type").value = c.type;
    $("conn-base").value = c.baseUrl || "";
    $("conn-key").value = "";
    $("conn-key").placeholder = "saved - leave empty to keep";
    setFormMsg($("conn-formmsg"), "", true);
    $("conn-testout").innerHTML = "";
  }

  function resetConnForm() {
    SK.editingConn = null;
    SK.connBaseDirty = false;
    $("conn-form-title").textContent = "New connection";
    $("conn-name").value = "";
    $("conn-type").value = "openai";
    $("conn-base").value = CONN_TYPE_BY_ID["openai"].base;
    $("conn-key").value = "";
    $("conn-key").placeholder = "";
    setFormMsg($("conn-formmsg"), "", true);
    $("conn-testout").innerHTML = "";
  }

  function connFormValues() {
    return {
      name: $("conn-name").value.trim(),
      type: $("conn-type").value,
      baseUrl: $("conn-base").value.trim(),
      apiKey: $("conn-key").value,
    };
  }

  function saveConnection() {
    var v = connFormValues();
    if (!v.name) { setFormMsg($("conn-formmsg"), "Name is required."); return; }
    if (!v.baseUrl) { setFormMsg($("conn-formmsg"), "Base URL is required (type " + v.type + " has no default)."); return; }
    var req;
    if (SK.editingConn && SK.editingConn !== v.name) {
      // rename: moves the connection and re-points profile references
      var rb = { from: SK.editingConn, to: v.name, type: v.type, baseUrl: v.baseUrl };
      if (v.apiKey) rb.apiKey = v.apiKey;
      req = apiJson("/api/connections/rename", "POST", rb);
    } else {
      var body = { name: v.name, type: v.type, baseUrl: v.baseUrl };
      if (v.apiKey) body.apiKey = v.apiKey;
      req = apiJson("/api/connections", "POST", body);
    }
    req.then(function () {
      setFormMsg($("conn-formmsg"), "Saved connection '" + v.name + "'.", true);
      SK.editingConn = v.name;
      SK.connBaseDirty = false;
      return loadConfigAndRender();
    }).catch(function (err) { setFormMsg($("conn-formmsg"), err.message); });
  }

  function testConnection() {
    var v = connFormValues();
    if (SK.editingConn && v.name === SK.editingConn && !v.apiKey) v.apiKey = "";
    setFormMsg($("conn-formmsg"), "Testing…", true);
    apiJson("/api/test-connection", "POST", v).then(function (res) {
      var models = res.models || [];
      $("conn-testout").innerHTML =
        '<span class="m">' + esc(models.length) + " model(s)" + (res.source === "fallback" ? " (fallback list)" : "") + ":</span> " +
        esc(models.slice(0, 6).join(", ")) + (models.length > 6 ? ", …" : "");
      setFormMsg($("conn-formmsg"), "Connection OK.", true);
    }).catch(function (err) {
      $("conn-testout").innerHTML = "";
      setFormMsg($("conn-formmsg"), "Test failed: " + err.message);
    });
  }

  // ------------------------------------------------------------------ //
  // Profiles tab
  // ------------------------------------------------------------------ //
  function renderProfiles() {
    // keep the profile-form connection dropdowns in sync with the current
    // config: connections created since the form was opened must show up
    ["prof-content-conn", "prof-other-conn", "prof-image-conn"].forEach(function (id) {
      var sel = $(id);
      fillConnSelect(sel, sel.value || "");
    });
    var list = $("prof-list");
    var rows = [ "<tr><th>Name</th><th>Content model</th><th>Other model</th><th>Image model</th></tr>" ];
    (SK.cfg.profiles || []).forEach(function (p) {
      var star = p.name === SK.cfg.defaultProfile ? "*" : "";
      // click the name to edit; deletion lives in the edit form (with confirm)
      var nameCell =
        '<button class="namebtn" data-edit="' + esc(p.name) + '" title="edit">' +
        esc(p.name) + "</button>" + (star ? " <span class=\"opnote\">(default)</span>" : "");
      rows.push(
        "<tr><td>" + nameCell + "</td>" +
        "<td>" + esc(modelRefLabel(p.content)) + "</td><td>" + esc(modelRefLabel(p.other)) + "</td><td>" + esc(modelRefLabel(p.image)) + "</td></tr>"
      );
    });
    list.innerHTML = rows.join("") || "<tr><td>No profiles yet - add one below.</td></tr>";
    Array.prototype.forEach.call(list.querySelectorAll("[data-edit]"), function (b) {
      b.addEventListener("click", function () { editProfile(b.dataset.edit); });
    });
  }

  function modelRefLabel(ref) {
    if (!ref || !ref.connection || !ref.model) return "(not set)";
    return ref.connection + " : " + ref.model;
  }

  function fillConnSelect(sel, selected) {
    sel.innerHTML = "";
    var opts = [ '<option value="">(choose connection)</option>' ];
    (SK.cfg.connections || []).forEach(function (c) {
      opts.push('<option value="' + esc(c.name) + '"' + (c.name === selected ? " selected" : "") + ">" + esc(c.name + " [" + c.type + "]") + "</option>");
    });
    sel.innerHTML = opts.join("");
  }

  function modelsCacheGet(connName, force, capability) {
    var key = connName + (capability ? "::" + capability : "");
    var cached = SK.modelsCache[key];
    if (cached && !force && Date.now() - cached.at < 5 * 60 * 1000 && cached.list.length) {
      return Promise.resolve(cached.list);
    }
    var url = "/api/models?connection=" + encodeURIComponent(connName) +
      (capability ? "&capability=" + encodeURIComponent(capability) : "");
    return apiJson(url, "GET").then(function (res) {
      var list = (res.models || []).map(String);
      SK.modelsCache[key] = { list: list, at: Date.now(), source: res.source };
      return list;
    });
  }

  function populateModelSelect(slot, connName, prefer, capability) {
    var sel = $(slot + "-conn") && $(slot + "-model");
    var modelSel = $(slot + "-model");
    var custom = $(slot + "-custom");
    modelSel.innerHTML = '<option value="">loading…</option>';
    custom.classList.add("hidden");
    if (!connName) {
      modelSel.innerHTML = '<option value="">(choose a connection first)</option>';
      return Promise.resolve();
    }
    return modelsCacheGet(connName, false, capability).then(function (list) {
      modelSel.innerHTML = "";
      var sorted = list.slice().sort(function (a, b) {
        var la = String(a || "").toLowerCase(), lb = String(b || "").toLowerCase();
        return la === lb ? String(a).localeCompare(String(b)) : (la < lb ? -1 : 1);
      });
      var opts = ['<option value="">(select model)</option>'];
      sorted.forEach(function (m) {
        opts.push('<option value="' + esc(m) + '"' + (m === prefer ? " selected" : "") + ">" + esc(m) + "</option>");
      });
      opts.push('<option value="__custom__">custom model id…</option>');
      modelSel.innerHTML = opts.join("");
      if (prefer && list.indexOf(prefer) < 0) {
        modelSel.value = "__custom__";
        custom.classList.remove("hidden");
        custom.value = prefer;
      } else {
        custom.value = "";
      }
      modelSel.value = modelSel.value || (list.length ? list[0] : "");
    }).catch(function (err) {
      modelSel.innerHTML = '<option value="">error: ' + esc(err.message) + '</option>';
    });
  }

  function editProfile(name) {
    var p = (SK.cfg.profiles || []).find(function (x) { return x.name === name; });
    if (!p) return;
    SK.editingProf = name;
    $("prof-form-title").textContent = "Edit profile: " + name;
    $("prof-name").value = p.name;
    fillConnSelect($("prof-content-conn"), p.content && p.content.connection);
    fillConnSelect($("prof-other-conn"), p.other && p.other.connection);
    fillConnSelect($("prof-image-conn"), p.image && p.image.connection);
    populateModelSelect("prof-content", p.content && p.content.connection, p.content && p.content.model);
    populateModelSelect("prof-other", p.other && p.other.connection, p.other && p.other.model);
    populateModelSelect("prof-image", p.image && p.image.connection, p.image && p.image.model, SK.profImageFilter ? "image" : null);
    setFormMsg($("prof-formmsg"), "", true);
  }

  function resetProfForm() {
    SK.editingProf = null;
    $("prof-form-title").textContent = "New profile";
    $("prof-name").value = "";
    fillConnSelect($("prof-content-conn"), "");
    fillConnSelect($("prof-other-conn"), "");
    fillConnSelect($("prof-image-conn"), "");
    $("prof-content-model").innerHTML = '<option value="">(choose a connection first)</option>';
    $("prof-other-model").innerHTML = '<option value="">(choose a connection first)</option>';
    $("prof-image-model").innerHTML = '<option value="">(optional - choose a connection first)</option>';
    $("prof-content-custom").classList.add("hidden");
    $("prof-other-custom").classList.add("hidden");
    $("prof-image-custom").classList.add("hidden");
    setFormMsg($("prof-formmsg"), "", true);
  }

  function profileFromForm() {
    var slotValue = function (slot) {
      var model = $(slot + "-model").value;
      if (model === "__custom__") model = $(slot + "-custom").value.trim();
      return {
        connection: $(slot + "-conn").value,
        model: model,
      };
    };
    var imageConn = $("prof-image-conn").value;
    var imageModel = $("prof-image-model").value;
    if (imageModel === "__custom__") imageModel = $("prof-image-custom").value.trim();
    return {
      name: $("prof-name").value.trim(),
      content: slotValue("prof-content"),
      other: slotValue("prof-other"),
      image: imageConn && imageModel ? { connection: imageConn, model: imageModel } : null,
    };
  }

  function saveProfile() {
    var p = profileFromForm();
    if (!p.name) { setFormMsg($("prof-formmsg"), "Profile name is required."); return; }
    if (!p.content.connection || !p.content.model) { setFormMsg($("prof-formmsg"), "Pick connection + model for the content model."); return; }
    if (!p.other.connection || !p.other.model) { setFormMsg($("prof-formmsg"), "Pick connection + model for the other-actions model."); return; }
    var req;
    if (SK.editingProf && SK.editingProf !== p.name) {
      // rename: moves the profile; the default-profile pointer follows
      req = apiJson("/api/profiles/rename", "POST", {
        from: SK.editingProf, to: p.name, content: p.content, other: p.other,
        image: p.image || {},
      });
    } else {
      req = apiJson("/api/profiles", "POST", p);
    }
    req.then(function () {
      setFormMsg($("prof-formmsg"), "Saved profile '" + p.name + "'.", true);
      SK.editingProf = p.name;
      return loadConfigAndRender();
    }).catch(function (err) { setFormMsg($("prof-formmsg"), err.message); });
  }

  // ------------------------------------------------------------------ //
  // Settings tab (approvals + debug)
  // ------------------------------------------------------------------ //
  function renderSettings() {
    var table = $("approvals-table");
    var rows = ['<tr><th>Action</th><th>What it does</th><th>Policy</th></tr>'];
    OP_TYPES.forEach(function (t) {
      var mode = (SK.cfg.approvals || {})[t] || "confirm";
      var radios = ["auto", "confirm", "deny"].map(function (m) {
        return '<label style="font-weight:400;margin:0 6px 0 0"><input type="radio" name="appr-' + esc(t) + '" value="' + m + '"' +
          (m === mode ? " checked" : "") + " /> " + m + "</label>";
      }).join("");
      rows.push("<tr><td><b>" + esc(OPS_META[t].label) + "</b></td><td class='muted'>" + esc(OPS_META[t].hint) + "</td><td>" + radios + "</td></tr>");
    });
    table.innerHTML = rows.join("");
    $("settings-debug").checked = !!SK.cfg.debug;
    $("settings-trackedchanges").checked = !!SK.cfg.trackedChanges;
        $("debug-path").textContent = SK.cfg.configPath.replace(/config\.json$/, "logs") + "/debug-YYYY-MM-DD.log";
  }

  function saveApprovals() {
    var approvals = {};
    OP_TYPES.forEach(function (t) {
      var el = document.querySelector('input[name="appr-' + t + '"]:checked');
      approvals[t] = el ? el.value : "confirm";
    });
    apiJson("/api/settings", "PUT", { approvals: approvals, generateNotes: SK.cfg ? !!SK.cfg.generateNotes : false }).then(function () {
      setFormMsg($("settings-msg"), "Approval policy + speaker-notes setting saved.", true);
      return loadConfigAndRender();
    }).catch(function (err) { setFormMsg($("settings-msg"), err.message); });
  }

  function saveDebugOption() {
    apiJson("/api/settings", "PUT", { trackedChanges: $("settings-trackedchanges").checked }).then(function () {
      SK.cfg.trackedChanges = $("settings-trackedchanges").checked;
      setFormMsg($("settings-tc-msg"), $("settings-trackedchanges").checked
        ? "ON - AI edits will land as tracked changes."
        : "OFF - AI edits apply directly.", true);
    }).catch(function (err) { setFormMsg($("settings-tc-msg"), err.message); });
    apiJson("/api/settings", "PUT", { debug: $("settings-debug").checked }).then(function () {
      return loadConfigAndRender();
    }).then(function () {
      setFormMsg($("settings-msg"), SK.cfg.debug ? "Debug logging ON - writing to the daily log." : "Debug logging off.", true);
    }).catch(function (err) { setFormMsg($("settings-msg"), err.message); });
  }

  // ------------------------------------------------------------------ //
  // About tab
  // ------------------------------------------------------------------ //
  function renderAbout() {
    var kv = [
      ["DocKick add-in", SK.cfg ? "v" + (SK.cfg.version || "0.1.3") : "0.1.3"],
      ["Server", "v" + (SK.cfg ? SK.cfg.version || "0.1.0" : "?") + " on 127.0.0.1:" + (SK.cfg ? SK.cfg.port : "?")],
      ["Config file", SK.cfg ? SK.cfg.configPath : "?"],
      ["Connections", SK.cfg ? (SK.cfg.connections || []).length : "?"],
      ["Profiles", SK.cfg ? (SK.cfg.profiles || []).length : "?"],
      ["Debug logging", SK.cfg && SK.cfg.debug ? "on" : "off"],
    ];
    $("about-info").innerHTML = kv.map(function (r) {
      return "<tr><td>" + esc(r[0]) + "</td><td>" + esc(r[1]) + "</td></tr>";
    }).join("");
  }

  // ------------------------------------------------------------------ //
  // Boot + wiring
  // ------------------------------------------------------------------ //
  function renderChat() {
    var sel = $("chat-profile");
    var opts = [];
    var def = SK.cfg.defaultProfile;
    (SK.cfg.profiles || []).forEach(function (p) {
      opts.push('<option value="' + esc(p.name) + '"' + (p.name === def ? " selected" : "") + ">" + esc(p.name) + "</option>");
    });
    if (!opts.length) opts.push('<option value="">(no profiles yet)</option>');
    sel.innerHTML = opts.join("");
  }

  function loadConfigAndRender() {
    return apiJson("/api/config", "GET").then(function (cfg) {
      SK.cfg = cfg;
      dbg("config.loaded", { connections: (cfg.connections || []).length, profiles: (cfg.profiles || []).length, debug: cfg.debug });
      renderChat();
      renderConnections();
      renderProfiles();
      renderSettings();
      renderAbout();
      var onboarding = $("chat-output");
      if (!(cfg.connections || []).length && !document.getElementById("onboarding-banner")) {
        var b = document.createElement("div");
        b.id = "onboarding-banner";
        b.className = "status error";
        b.textContent = "No AI connections configured yet. Open the Connections tab and add one (e.g. Ollama on http://localhost:11434 - no key needed).";
        onboarding.insertBefore(b, onboarding.firstChild);
      }
      if ((cfg.connections || []).length && document.getElementById("onboarding-banner")) {
        var el = document.getElementById("onboarding-banner");
        el.parentNode.removeChild(el);
      }
      $("statusbar").textContent = "DocKick v" + (cfg.version || "0.1.0") +
        (window.location && window.location.protocol === "https:" ? " - \u{1F512} " : " - ") +
        "server 127.0.0.1:" + cfg.port + (cfg.debug ? " - debug ON" : "");
      return cfg;
    });
  }

  function wireEvents() {
    Array.prototype.forEach.call(document.querySelectorAll("#tabs button"), function (b) {
      b.addEventListener("click", function () { switchTab(b.dataset.tab); });
    });

    $("chat-run").addEventListener("click", runPrompt);
    $("chat-input").addEventListener("keydown", function (ev) {
      if ((ev.metaKey || ev.ctrlKey) && ev.key === "Enter") { ev.preventDefault(); runPrompt(); }
    });
    $("chat-new").addEventListener("click", function () {
      resetHistory();
      $("chat-output").innerHTML = "";
      $("chat-input").value = "";
      setStatus("Conversation history cleared.", "ok");
    });
    $("chat-refresh-deck").addEventListener("click", function () { updateDeckSummary(); setStatus("Document refreshed.", ""); });

    $("conn-new").addEventListener("click", resetConnForm);
    $("conn-save").addEventListener("click", saveConnection);
    $("conn-test").addEventListener("click", testConnection);
    $("conn-delete").addEventListener("click", function () {
      if (!SK.editingConn) {
        setFormMsg($("conn-formmsg"), "Nothing selected to delete - click a connection name first.");
        return;
      }
      var used = (SK.cfg.profiles || []).filter(function (p) {
        return (p.content && p.content.connection === SK.editingConn) ||
               (p.other && p.other.connection === SK.editingConn);
      }).map(function (p) { return p.name; });
      if (used.length) {
        window.alert("Connection '" + SK.editingConn + "' is used by profile(s): " +
          used.join(", ") + ". Remove it from those profiles first.");
        return;
      }
      if (!window.confirm("Delete connection '" + SK.editingConn + "'?")) return;
      apiJson("/api/connections/" + encodeURIComponent(SK.editingConn), "DELETE")
        .then(function () { resetConnForm(); return loadConfigAndRender(); })
        .catch(function (err) { window.alert(err.message); });
    });
    $("conn-type").addEventListener("change", function () {
      var t = CONN_TYPE_BY_ID[$("conn-type").value];
      if (t && !SK.connBaseDirty) $("conn-base").value = t.base;
      $("conn-key").placeholder = t && t.keyHint ? t.keyHint : "";
    });
    $("conn-base").addEventListener("input", function () { SK.connBaseDirty = true; });

    $("prof-new").addEventListener("click", resetProfForm);
    $("prof-save").addEventListener("click", saveProfile);
    $("prof-delete").addEventListener("click", function () {
      if (!SK.editingProf) {
        setFormMsg($("prof-formmsg"), "Nothing selected to delete - click a profile name first.");
        return;
      }
      if (!window.confirm("Delete profile '" + SK.editingProf + "'?")) return;
      apiJson("/api/profiles/" + encodeURIComponent(SK.editingProf), "DELETE")
        .then(function () { resetProfForm(); return loadConfigAndRender(); })
        .catch(function (err) { window.alert(err.message); });
    });
    $("prof-default").addEventListener("click", function () {
      var name = $("prof-name").value.trim();
      if (!name) { setFormMsg($("prof-formmsg"), "Fill in the profile name first."); return; }
      apiJson("/api/settings", "PUT", { defaultProfile: name }).then(loadConfigAndRender).then(function () {
        setFormMsg($("prof-formmsg"), "Default profile set to '" + name + "'.", true);
      }).catch(function (err) { setFormMsg($("prof-formmsg"), err.message); });
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-refresh]"), function (btn) {
      btn.addEventListener("click", function () {
        var slot = "prof-" + btn.dataset.refresh;
        var connName = $(slot + "-conn").value;
        if (!connName) return;
        var isImage = slot.indexOf("-image") > 0;
        var cap = isImage && SK.profImageFilter ? "image" : null;
        delete SK.modelsCache[connName + (cap ? "::image" : "")];
        populateModelSelect(slot, connName, $(slot + "-model").value, cap);
      });
    });
    $("prof-image-toggle").addEventListener("click", function () {
      SK.profImageFilter = !SK.profImageFilter;
      $("prof-image-toggle").textContent = SK.profImageFilter ? "image-capable only" : "show all models";
      var connName = $("prof-image-conn").value;
      if (connName) {
        delete SK.modelsCache[connName + (SK.profImageFilter ? "::image" : "")];
        populateModelSelect("prof-image", connName, "", SK.profImageFilter ? "image" : null);
      }
    });
    ["content", "other", "image"].forEach(function (slot) {
      $("prof-" + slot + "-conn").addEventListener("change", function () {
        populateModelSelect("prof-" + slot, $("prof-" + slot + "-conn").value, "",
          slot === "image" && SK.profImageFilter ? "image" : null);
      });
      $("prof-" + slot + "-model").addEventListener("change", function () {
        var isCustom = $("prof-" + slot + "-model").value === "__custom__";
        $("prof-" + slot + "-custom").classList.toggle("hidden", !isCustom);
        if (isCustom) $("prof-" + slot + "-custom").focus();
      });
    });
    // convenience: picking a content model with an untouched 'other' slot
    // copies the choice (connection + model, incl. custom ids) over
    $("prof-content-model").addEventListener("change", function () {
      var model = $("prof-content-model").value;
      var otherModel = $("prof-other-model");
      if (otherModel.value) return; // 'other' already has a choice - leave it
      var conn = $("prof-content-conn").value;
      if (!conn || !model) return;
      $("prof-other-conn").value = conn;
      populateModelSelect("prof-other", conn, model).then(function () {
        if (model === "__custom__") {
          $("prof-other-model").value = "__custom__";
          $("prof-other-custom").classList.remove("hidden");
          $("prof-other-custom").value = $("prof-content-custom").value;
        }
      });
    });

    $("settings-save").addEventListener("click", saveApprovals);
    Array.prototype.forEach.call(document.querySelectorAll("#settings-subtabs button"), function (b) {
      b.addEventListener("click", function () { switchSettingsSub(b.dataset.subtab); });
    });
    $("settings-debug-save").addEventListener("click", saveDebugOption);
    var updBtn = $("upd-check");
    if (updBtn) updBtn.addEventListener("click", updCheck);
    wireAutoApplySetting();
    var updAuto = $("upd-auto");
    if (updAuto) {
      // remember the preference (on by default); the boot check honors it
      apiJson("/api/config", "GET").then(function (cfg) {
        updAuto.checked = !(cfg && cfg.autoUpdateCheck === false);
        if (updAuto.checked) updCheck();   // silent availability check at boot
      }, function () { updAuto.checked = true; updCheck(); });
      updAuto.addEventListener("change", function () {
        apiJson("/api/settings", "PUT", { autoUpdateCheck: !!updAuto.checked });
      });
    }
    $("settings-tc-save").addEventListener("click", function () {
      apiJson("/api/settings", "PUT", { trackedChanges: $("settings-trackedchanges").checked }).then(function () {
        SK.cfg.trackedChanges = $("settings-trackedchanges").checked;
        setFormMsg($("settings-tc-msg"), "Revision setting saved.", true);
      }).catch(function (err) { setFormMsg($("settings-tc-msg"), err.message); });
    });
    updateSessionLabel();
    wireSessions();
    $("req-clear").addEventListener("click", clearRequests);
    $("req-refresh").addEventListener("click", renderRequests);
    $("sim-reset").addEventListener("click", function () {
      window.SimDeck.resetAll();
      renderSimDeck();
      updateDeckSummary();
      setStatus("Simulator document reset.", "");
    });
  }

  function bootOffice() {
    var badge = $("host-badge");
    if (typeof Office === "undefined") {
      badge.style.display = "";
      badge.textContent = "NO OFFICE.JS";
      setStatus("Office.js failed to load. Check the internet connection, then reopen the pane.", "error");
      return;
    }
    Office.onReady(function (info) {
      badge.style.display = "";
      badge.textContent = "DOC";
      SK.deck = new WordDeck();
      applyTheme();
      watchTheme();
      finishBoot();
    });
  }

  function finishBoot() {
    wireEvents();
    bootLoad(0);
  }

  function bootSuccess(cfg) {
    updateDeckSummary();
    renderSimDeck();
    if (!(cfg.connections || []).length) { switchTab("settings"); switchSettingsSub("connections"); }
    setStatus("Ready. Debug: " + (cfg.debug ? "on" : "off") + ".", "ok");
  }

  // retries cover the update window: the pane can be opened while the server
  // is restarting (installer / LaunchAgent)
  function bootLoad(attempt) {
    loadConfigAndRender().then(bootSuccess, function (err) { bootFail(err, attempt); });
  }

  function bootFail(err, attempt) {
    if (attempt < 3) {
      setStatus("Server not reachable yet - retrying (" + (attempt + 1) + "/3)…", "busy");
      setTimeout(function () { bootLoad(attempt + 1); }, 1000);
      return;
    }
    setStatus("Cannot reach the DocKick server: " + err.message + " Is it running? (install.sh / install.ps1 / LaunchAgent)", "error");
    $("statusbar").textContent = "server unreachable: " + err.message;
  }

  function boot() {
    var params = new URLSearchParams(String(window.location.search || ""));
    SK.sim = window.__SK_SIM__ === true || params.has("sim");

    // internal test-run flag: simulator runs and ?nosave=1 never write
    // to the session history (and skip AI-naming); force with &save=1
    SK.forceSave = /[?&]save=1/.test(window.location.search);
    SK.testMode = (SK.sim || /\?.*nosave=1/.test(window.location.search)) && !SK.forceSave;
    if (SK.sim) {
      $("host-badge").style.display = "";
      $("host-badge").textContent = "SIM";
      var simTab = $("sim-panel");
      simTab.style.display = "";
      var nav = document.createElement("button");
      nav.textContent = "Sim document";
      nav.dataset.tab = "sim";
      $("tabs").appendChild(nav);
      nav.addEventListener("click", function () { switchTab("sim"); });
      SK.deck = new window.SimDeck();
      applyTheme();
      finishBoot();
    } else {
      bootOffice();
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();