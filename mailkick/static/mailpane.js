/*
 * MailKick task pane (runs inside Microsoft Outlook via Office.js, or in the
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
    app: "mailkick",
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
    draft_reply:    { label: "draft reply", hint: "Opens a reply window with AI-written HTML content - you review and send." },
    compose_body:   { label: "replace compose text", hint: "Rewrites the body of the email you are composing (review first)." },
    set_categories: { label: "categories", hint: "Adds Outlook categories to the selected message; names are checked against your mailbox master list when Outlook's permission level allows reading it." },
    save_draft:     { label: "save draft", hint: "Saves the message you are composing as a draft." },
    new_message:    { label: "new message from template", hint: "Opens a prefilled NEW email window (to/subject/body) - you review and send." },
    forward:        { label: "forward", hint: "Opens a forward window for the selected mail (quoted original + your optional note) - you review and send." },
    new_meeting:    { label: "new meeting", hint: "Opens a MEETING/calendar form (subject, attendees, start/end, location, body) - you review and send the invite." },
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
    // per-add-in setting first; the legacy global is the fallback
    var byApp = SK.cfg && SK.cfg.autoApplyByApp;
    if (byApp && byApp[SK.app] !== undefined) return !!byApp[SK.app];
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
      SK.autoApplyUserTouched = true;   // boot config loads must not undo this
      var byAppPut = {};
      byAppPut[SK.app] = on;   // module-specific: this add-in only
      apiJson("/api/settings", "PUT", { autoApplyByApp: byAppPut }).then(function () {
        if (SK.cfg) {
          if (!SK.cfg.autoApplyByApp) SK.cfg.autoApplyByApp = {};
          SK.cfg.autoApplyByApp[SK.app] = on;
        }
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

  // (in SYSTEM_PROMPT: explicit newline rule)
  var SYSTEM_PROMPT = [
    "You are MailKick, an AI assistant inside a Microsoft Outlook task pane.",
    "When the context lists MULTI-SELECTED ITEMS, treat them as the batch the user wants processed (e.g. suggest or set categories per item).",
    "You get the current email context and a user request. In text fields use real line breaks, never literal backslash-n. Reply with ONLY one JSON",
    "object - never commentary, never code fences, never reasoning. Format A",
    "(conversational info ONLY: questions, summaries, translations, explanations about the content):",
    '{"kind":"reply","text":"your answer"}',
    "Format B (mailbox actions):",
    '{"kind":"mail_ops","ops":[ ... ]}',
    "",
    "Allowed op types:",
    '- {"type":"draft_reply","html":"<p>reply text</p>","all":true} (READ mode: opens a reply window with this content; add "all":true to reply to EVERYONE on the thread; simple inline tags only - p, br, b, i, a, ul, li; write the reply in the SAME LANGUAGE and tone as the conversation)',
    '- {"type":"new_meeting","subject":"...","required":["a@b.c"],"optional":[],"start":"2026-11-02T14:00","end":"2026-11-02T15:00","location":"...","body":"..."} (opens a NEW MEETING form; required/optional = email addresses; ISO date-times; the form opens for YOU to review - MailKick never sends anything itself)',
    '- {"type":"compose_body","html":"<p>...</p>"} (COMPOSE mode only: replaces the body of the email being written)',
    '- {"type":"set_categories","names":["Category"]} (names should exist in the KNOWN CATEGORIES list from the context when present; if the context says the master list is unavailable, still use sensible category names - the pane will report unchecked ones)',
    '- {"type":"save_draft"} (COMPOSE mode: saves the message being composed as a draft)',
    '- {"type":"new_message","html":"..."} (read or compose: opens a NEW email window prefilled; simple inline tags only)',
    '- {"type":"forward","comment":"<p>short note above the quoted original (optional)</p>","to":["a@b.c"]} (READ mode: opens a FORWARD window for the current mail - subject FW:..., quoted original, your optional note; recipients optional)',
    "",
    "Rules:",
    "- Named templates: when the request mentions a template, its content is listed in the EMAIL CONTEXT (NAMED TEMPLATE '<name>'). Use it as the structural/wording base for draft_reply/compose_body/new_message, then adapt per the user's instructions and the email context.",
    "- Choose ops that match the mode stated in the EMAIL CONTEXT; wrong-mode ops are rejected.",
    "- ACTION RULE (critical): any request to reply, draft, answer, respond, forward, fwd, send, write, notify, thank, confirm, decline, request, categorize or otherwise ACT on the mail is an ACTION request: answer with format B ops (draft_reply / forward / new_message / set_categories / compose_body / save_draft), NEVER with format A prose. Never write the reply/forward text into the chat, never say 'here is what I would say', never ask whether to do it - propose the op. Format A is ONLY for questions, summaries, translations and explanations about the content.",
    "- 'reply'/'answer'/'respond' on the open mail means draft_reply (add \"all\":true when the user says reply to all/everyone). 'forward'/'fwd' means the forward op. Meeting/calendar/invite requests mean new_meeting.",
    "- The context may list ATTACHMENTS (names + sizes) of the open mail: reference them by name in drafts/explanations, but you CANNOT read their contents.",
    "- Reply/summary/translate/brainstorm requests that need no change: use format A.",
    "- Keep reply drafts short and human; greeting + body + signoff unless told otherwise.",
    "- LIVE WEB LOOKUPS: when the request needs up-to-date info you cannot reliably know (current events, news, prices, recent releases, rankings, live data), reply with ONLY {\"kind\":\"web_lookup\"} - nothing else. MailKick then runs a web-enabled lookup and re-asks you with fresh citable results. Never use it for mail edits, styling, general concepts, or what the email context already answers.",
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
  var editingTpl = null;

  function renderTemplates() {
    apiJson("/api/templates", "GET").then(function (res) {
      var rows = res.templates || [];
      SK.templateIds = {};
      var list = $("tpl-list");
      if (!list) return;
      var frag = document.createDocumentFragment();
      var cnt = $("tpl-count");
      if (cnt) cnt.textContent = rows.length + " template(s)";
      if (!rows.length) {
        var e = document.createElement("div");
        e.className = "opnote";
        e.textContent = "No templates yet. Create one, then say \u201Creply using template <name> \u2026\u201D.";
        frag.appendChild(e);
      }
      rows.forEach(function (t) {
        SK.templateIds[String(t.id)] = t.name;
        var card = document.createElement("div");
        card.className = "tplcard";
        var head = document.createElement("div");
        head.className = "sesshead";
        var name = document.createElement("span");
        name.className = "tplname";
        name.title = "Click to show the template content";
        name.textContent = t.name;
        var when = new Date(t.updated * 1000);
        var meta = document.createElement("span");
        meta.className = "sessmeta";
        meta.textContent = t.len + " chars · " + ("0" + when.getHours()).slice(-2) + ":" + ("0" + when.getMinutes()).slice(-2) + " · " + when.toLocaleDateString();
        var actions = document.createElement("span");
        actions.className = "sessactions";
        var useBtn = document.createElement("button");
        useBtn.className = "btn subtle";
        useBtn.textContent = "Use";
        useBtn.title = "Prefill the prompt to use this template";
        useBtn.addEventListener("click", function () {
          var inp = $("chat-input");
          inp.value = "reply using template " + t.name + ": ";
          switchTab("chat");
          inp.focus();
        });
        var editBtn = document.createElement("button");
        editBtn.className = "btn subtle";
        editBtn.textContent = "Edit";
        editBtn.addEventListener("click", function () { editTemplate(t.id); });
        var delBtn = document.createElement("button");
        delBtn.className = "btn subtle danger";
        delBtn.textContent = "Delete";
        delBtn.addEventListener("click", function () {
          apiJson("/api/templates/" + t.id, "DELETE").then(function () { renderTemplates(); setFormMsg($("tpl-formmsg"), "Template deleted.", true); })
            .catch(function (err) { setFormMsg($("tpl-formmsg"), err.message); });
        });
        actions.appendChild(useBtn); actions.appendChild(editBtn); actions.appendChild(delBtn);
        head.appendChild(name); head.appendChild(meta); head.appendChild(actions);
        card.appendChild(head);
        var show = document.createElement("div");
        show.className = "tplbody hidden";
        name.addEventListener("click", function () {
          apiJson("/api/templates/" + t.id, "GET").then(function (doc) {
            if (show.classList.contains("hidden")) { show.textContent = doc.body || "(empty)"; show.classList.remove("hidden"); }
            else { show.classList.add("hidden"); }
          });
        });
        card.appendChild(show);
        frag.appendChild(card);
      });
      list.innerHTML = "";
      list.appendChild(frag);
    }).catch(function (err) {
      var list = $("tpl-list");
      if (list) list.innerHTML = '<div class="opnote">Cannot reach the templates store: ' + esc((err && err.message) || String(err)) + "</div>";
    });
  }

  function editTemplate(tid) {
    apiJson("/api/templates/" + tid, "GET").then(function (doc) {
      editingTpl = doc.id;
      $("tpl-form-title").textContent = "Edit template: " + doc.name;
      $("tpl-name").value = doc.name;
      $("tpl-body").value = doc.body || "";
      $("tpl-delete").classList.remove("hidden");
      $("tpl-cancel").classList.remove("hidden");
      setFormMsg($("tpl-formmsg"), "", true);
      $("tpl-name").focus();
    });
  }

  function resetTplForm() {
    editingTpl = null;
    $("tpl-form-title").textContent = "New template";
    $("tpl-name").value = "";
    $("tpl-body").value = "";
    $("tpl-delete").classList.add("hidden");
    $("tpl-cancel").classList.add("hidden");
    setFormMsg($("tpl-formmsg"), "", true);
  }

  function saveTemplate() {
    var name = $("tpl-name").value.trim();
    var body = $("tpl-body").value;
    if (!name) { setFormMsg($("tpl-formmsg"), "A name is required."); return; }
    if (!body.trim()) { setFormMsg($("tpl-formmsg"), "The template has no content."); return; }
    var req = editingTpl
      ? apiJson("/api/templates/" + editingTpl, "PUT", { name: name, body: body })
      : apiJson("/api/templates", "POST", { name: name, body: body });
    req.then(function () {
      setFormMsg($("tpl-formmsg"), "Template saved.", true);
      editingTpl = editingTpl || null;
      openTemplatesRender();
    }).catch(function (err) { setFormMsg($("tpl-formmsg"), err.message); });
  }

  function openTemplatesRender() { renderTemplates(); }

  function wireTemplates() {
    var nw = $("tpl-new");
    if (nw) nw.addEventListener("click", resetTplForm);
    var rf = $("tpl-refresh");
    if (rf) rf.addEventListener("click", renderTemplates);
    var sv = $("tpl-save");
    if (sv) sv.addEventListener("click", saveTemplate);
    var cn = $("tpl-cancel");
    if (cn) cn.addEventListener("click", resetTplForm);
    var dl = $("tpl-delete");
    if (dl) dl.addEventListener("click", function () {
      if (!editingTpl) return;
      apiJson("/api/templates/" + editingTpl, "DELETE").then(function () {
        resetTplForm(); renderTemplates();
      }).catch(function (err) { setFormMsg($("tpl-formmsg"), err.message); });
    });
  }

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
    "\\b(email|mail|message|reply|respond|respond|draft|summary|summarize|translate|categorize|category|categories|subject|body|sender|attach|attachment|compose|send)\\b", "i"
  );

  function cleanLiteralNL(t) {
    // a model that writes \\n inside its JSON strings would render a literal "\n";
    // turn those into real line breaks (display + history + draft html)
    return String(t).replace(/\\n/g, "\n");
  }

  function assistantHistoryText(parsed) {
    // a compact recap for the next round; never the raw ops JSON
    if (!parsed || !parsed.ok) {
      var rawT = (SK.lastModelText && SK.lastModelText.text) || "";
      return String(rawT).slice(0, 400);
    }
    if (parsed.kind === "reply") return cleanLiteralNL(parsed.text || "").slice(0, HISTORY_MAX_CHARS);
    // store only the action types and how many per type were executed -
    // no descriptions (keeps the session/context compact)
    var ops = parsed.ops || [];
    var counts = {};
    ops.forEach(function (op) {
      var t = op && op.type ? String(op.type) : "unknown";
      counts[t] = (counts[t] || 0) + 1;
    });
    var byType = Object.keys(counts).map(function (t) { return t + " x" + counts[t]; }).join(", ");
    return ("Executed " + ops.length + " action(s)" + (byType ? ": " + byType : "")).slice(0, HISTORY_MAX_CHARS);
  }

  function getTemplateBlock(instruction) {
    var text = String(instruction || "");
    var m = /\btemplates?\s+(?:named\s+)?["\u201c]?([a-zA-Z0-9][a-zA-Z0-9 _.-]{1,48})["\u201c]?/i.exec(text);
    return apiJson("/api/templates", "GET").then(function (res) {
      var tpls = res.templates || [];
      if (!m || !tpls.length) return "";
      var want = m[1].trim().toLowerCase();
      var hit = null;
      tpls.forEach(function (t) {
        var n = String(t.name).toLowerCase();
        if (!hit && (n === want || n.startsWith(want) || want.startsWith(n))) hit = t;
      });
      if (!hit) return "";
      return apiJson("/api/templates/" + hit.id, "GET").then(function (doc) {
        if (!doc || !doc.body) return "NAMED TEMPLATE '" + hit.name + "': (body missing)";
        return "NAMED TEMPLATE '" + doc.name + "':\n" + String(doc.body).slice(0, 6000);
      });
    }).catch(function () { return ""; });
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
  // Mailbox adapter (Outlook mailbox.item via Office.js)
  // ------------------------------------------------------------------ //
    function MailDeck() {}

  MailDeck.prototype.run = function (name, fn) {
    // Outlook's mailbox API = a plain object once Office.onReady resolved;
    // no Word.run-style context batching - each op = its own async sequence
    var self = this;
    return fn(self);
  };

  MailDeck.prototype.capabilities = function () {
    return { text: true, host: "Outlook" };
  };

  function promisify(fn) {
    return new Promise(function (resolve, reject) {
      try { fn(resolve, reject); } catch (e) { reject(e); }
    });
  }

  function toHtml(text) {
    // the model returns draft_reply/compose_body as TEXT; wrap to minimal HTML
    var escd = cleanLiteralNL(String(text))
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    return "<p>" + escd.split(/\n{2,}/).join("</p><p>").replace(/\n/g, "<br />") + "</p>";
  }

  MailDeck.prototype.getContext = function (opts) {
    var self = this;
    return MailDeck.prototype._getContextCore.call(self, opts).then(function (ctx) {
      return masterCategoriesList().then(function (list) {
        if (list) ctx.masterCategories = list;
        return ctx;
      });
    }).then(function (ctx) {
      // Mailbox 1.13: enrich with the items the user multi-selected (batch triage)
      return new Promise(function (resMulti) {
        try {
          if (Office.context.requirements.isSetSupported("Mailbox", "1.13") &&
              Office.context.mailbox && Office.context.mailbox.getSelectedItemsAsync) {
            Office.context.mailbox.getSelectedItemsAsync(function (mr) {
              try {
                if (!mr || mr.status !== "succeeded" || !Array.isArray(mr.value) || !mr.value.length) return resMulti(null);
                resMulti(mr.value.slice(0, 30).map(function (it) {
                  var d = it.itemData || it || {};
                  var frm = d.from || {};
                  return {
                    subject: String((d.subject != null ? d.subject : "") || ""),
                    sender: String((frm && (frm.displayName || frm.emailAddress)) || ""),
                  };
                }));
              } catch (e2) { resMulti(null); }
            });
          } else { resMulti(null); }
        } catch (e3) { resMulti(null); }
      }).then(function (multi) {
        if (multi) ctx.selectedItems = multi;
        return ctx;
      });
    });
  };

  MailDeck.prototype._getContextCore = function (opts) {
    var minimal = !!(opts && opts.minimal);
    return promisify(function (resolve, reject) {
      try {
        var item = Office.context.mailbox.item;
        if (!item) {
          reject(new FriendlyError("Open or select a message first (the reading pane must stay open)."));
          return;
        }
        var out = { host: "Outlook", mode: "", subject: "", sender: "", to: "", bodyText: "", categories: [], warnings: [], minimal: minimal };
        var isCompose = (item.getAsync !== undefined && item.subject && item.subject.getAsync);
        out.mode = isCompose ? "compose" : "read";
        Promise.all([
          new Promise(function (res, rej) {
            if (isCompose) item.subject.getAsync(function (r) { r.status === "succeeded" ? res(r.value) : rej(new FriendlyError(r.error.message)); });
            else res(String(item.subject || ""));
          })
        ]).then(function (vals) {
          out.subject = String(vals[0] || "");
          if (out.mode === "read" && item.from) {
            out.sender = (item.from.displayName || item.from.emailAddress || "");
          }
          if (minimal) { resolve(out); return; }
          // optional steps: degrade to defaults on failure instead of failing the whole context
          var opt = function (fn, fallback) {
            return new Promise(function (res) {
              try { fn(function (r) { (r && r.status === "succeeded") ? res(r.value) : res(fallback); }, res); }
              catch (e) { res(fallback); }
            });
          };
          return Promise.all([
            opt(function (ok, fail) { item.to.getAsync(ok, fail); }, ""),
            opt(function (ok, fail) { item.categories ? item.categories.getAsync(ok, fail) : ok([]); }, []),
          ]).then(function (vals2) {
            out.to = String(vals2[0] || "");
            out.categories = (vals2[1] || []).map(function (c) { return c.displayName || c; });
            try {
              // read-mode attachment listing is a sync item property (Mailbox 1.1)
              if (!isCompose && item.attachments && typeof item.attachments.length === "number") {
                out.attachments = item.attachments.slice(0, 12).map(function (a) {
                  return { name: String(a.name || "?"), size: a.size || 0 };
                });
              }
            } catch (eAtt) { /* listing not available */ }
            var done = function (txt) { out.bodyText = String(txt || ""); resolve(out); };
            try {
              item.body.getAsync("text", function (r) {
                if (r.status === "succeeded") done(r.value);
                else { out.warnings.push("body: " + (r.error && r.error.message || "read failed")); resolve(out); }
              });
            } catch (e) { out.warnings.push("body: " + e.message); resolve(out); }
          });
        }, reject);
      } catch (e) { reject(e); }
    });
  };

  MailDeck.prototype.draftReply = function (html, all) {
    return promisify(function (resolve, reject) {
      var item = Office.context.mailbox.item;
      if (!item || (!all && !item.displayReplyForm) || (all && !item.displayReplyAllForm)) {
        reject(new FriendlyError(all
          ? "Open a received message first (a reply-all needs an existing email)."
          : "Open a received message first (a reply needs an existing email)."));
        return;
      }
      if (all) item.displayReplyAllForm(String(html));   // Mailbox 1.1
      else item.displayReplyForm(String(html));
      resolve({ ok: true });
    });
  };

  MailDeck.prototype.newMeeting = function (op) {
    return promisify(function (resolve, reject) {
      var mb = Office.context.mailbox;
      if (!mb || !mb.displayNewAppointmentForm) {
        reject(new FriendlyError("this Outlook build cannot open meeting forms (needs Mailbox 1.1+)."));
        return;
      }
      var startM = new Date(String(op.start || "")), endM = new Date(String(op.end || ""));
      if (isNaN(startM.getTime()) || isNaN(endM.getTime())) {
        reject(new FriendlyError("start/end must be ISO date-times like 2026-11-02T14:00."));
        return;
      }
      if (endM < startM) { var tmp = endM; endM = startM; startM = tmp; }
      mb.displayNewAppointmentForm({
        requiredAttendees: (op.required || []).slice(0),
        optionalAttendees: (op.optional || []).slice(0),
        start: startM,
        end: endM,
        location: String(op.location || ""),
        subject: String(op.subject || ""),
        body: String(op.body || ""),
      });
      resolve({ ok: true });
    });
  };

  MailDeck.prototype.forwardMessage = function (op, ctx) {
    return promisify(function (resolve, reject) {
      var mb = Office.context.mailbox;
      if (!mb || !mb.displayNewMessageForm) {
        reject(new FriendlyError("Forward needs a newer Outlook build (displayNewMessageForm is missing)."));
        return;
      }
      var c = ctx || {};
      if (!c.bodyText) {
        reject(new FriendlyError("Forward needs the full email context - enable the include-email-context checkbox."));
        return;
      }
      var esc = function (sv) {
        return String(sv == null ? "" : sv).replace(/[&<>"]/g, function (ch) {
          var q = String.fromCharCode(34); return { "&": "&amp;", "<": "&lt;", ">": "&gt;", q: "&quot;" }[ch];
        });
      };
      var bodyHtml = esc(c.bodyText).split(/\r?\n/).join("<br />");
      var head = "-----Original Message-----<br />" +
        "From: " + esc(c.sender || "") + "<br />" +
        (Array.isArray(c.to) ? "To: " + esc(c.to.join(", ")) : "To: " + esc(c.to || "")) + "<br />" +
        "Subject: " + esc(c.subject || "") + "<br /><br />";
      var note = op.html ? op.html + "<br /><br />" : "";
      mb.displayNewMessageForm({
        toRecipients: op.to || [],
        subject: "FW: " + String(c.subject || ""),
        htmlBody: note + head + bodyHtml,
      });
      resolve({ ok: true });
    });
  };

  MailDeck.prototype.composeBody = function (html) {
    return promisify(function (resolve, reject) {
      var item = Office.context.mailbox.item;
      if (!item || !item.body || !item.body.setAsync) {
        reject(new FriendlyError("replace compose text works only while composing an email."));
        return;
      }
      item.body.setAsync(String(html), { coercionType: Office.CoercionType.Html }, function (r) {
        r.status === "succeeded" ? resolve({ ok: true }) : reject(new FriendlyError(r.error.message));
      });
    });
  };

  function masterCategoriesList() {
    // mailbox-wide category master list - needs the elevated ReadWriteMailbox
    // permission; any failure (incl. "Elevated permission is required")
    // resolves null instead of throwing
    return promisify(function (resolve) {
      try {
        var mb = Office.context.mailbox;
        if (!mb || !mb.masterCategories || !mb.masterCategories.getAsync) { resolve(null); return; }
        mb.masterCategories.getAsync(function (r) {
          if (!r || r.status !== "succeeded") { resolve(null); return; }
          resolve((r.value || []).map(function (c) { return c.displayName; }));
        });
      } catch (e) { resolve(null); }
    });
  }

  MailDeck.prototype.setCategories = function (names) {
    return promisify(function (resolve, reject) {
      var item = Office.context.mailbox.item;
      if (!item || !item.categories || !item.categories.addAsync) {
        reject(new FriendlyError("this Outlook build does not support categories from add-ins (Mailbox 1.8)."));
        return;
      }
      masterCategoriesList().then(function (list) {
        var apply = function (note) {
          item.categories.addAsync(names.map(String), function (r) {
            r.status === "succeeded" ? resolve({ ok: true, note: note }) : reject(new FriendlyError(r.error.message));
          });
        };
        if (!list) {
          apply("(applied WITHOUT checking the master list - reading it needs elevated mailbox permissions; new names may get created in your mailbox)");
          return;
        }
        var known = {};
        list.forEach(function (c) { known[String(c).toLowerCase()] = true; });
        var unknown = (names || []).filter(function (n) { return !known[String(n).toLowerCase()]; });
        if (unknown.length) {
          reject(new FriendlyError("category not in your master list: " + unknown.join(", ") + " - create it in Outlook first."));
          return;
        }
        apply("");
      });
    });
  };

  MailDeck.prototype.newMessage = function (payload) {
    return promisify(function (resolve, reject) {
      var mb = Office.context.mailbox;
      if (!mb || !mb.displayNewMessageFormAsync) {
        reject(new FriendlyError("this Outlook build cannot open new-message windows from add-ins (Mailbox 1.9)."));
        return;
      }
      mb.displayNewMessageFormAsync({
        htmlBody: String(payload.html),
        subject: payload.subject || "",
        toRecipients: (payload.to || []).map(function (a) { return { emailAddress: String(a) }; }),
      }, function (r) {
        r.status === "succeeded" ? resolve({ ok: true }) : reject(new FriendlyError(r.error && r.error.message || "failed"));
      });
    });
  };

  MailDeck.prototype.saveDraft = function () {
    return promisify(function (resolve, reject) {
      var item = Office.context.mailbox.item;
      if (!item || !item.saveAsync) {
        reject(new FriendlyError("save draft works only while composing an email."));
        return;
      }
      item.saveAsync(function (r) {
        r.status === "succeeded" ? resolve({ ok: true }) : reject(new FriendlyError(r.error.message));
      });
    });
  };

  MailDeck.prototype.renderDebug = function (el) { el.innerHTML = ""; };
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

  function validateOps(rawOps, ctx) {
    var ops = [], invalid = [];
    (rawOps || []).forEach(function (op, index) {
      if (!op || typeof op !== "object") { invalid.push({ op: op, reason: "not an object" }); return; }
      var t = op.type;
      if (OP_TYPES.indexOf(t) < 0) { invalid.push({ op: op, reason: "unknown op type '" + t + "'" }); return; }
      if (!ctx && t !== "draft_reply" && t !== "set_categories" && t !== "new_meeting") {
        invalid.push({ op: op, reason: "needs email context (enable \u0027include email context\u0027)" });
        return;
      }
      var bad = function (reason) { invalid.push({ op: op, reason: reason }); };
      var paraOk = function (v) { return typeof v === "number" && Math.floor(v) === v && v >= 0 && (!ctx || v < ctx.paragraphCount); };
      var out = { type: t };
      if (t === "draft_reply") {
        var html = op.html;
        if (typeof html !== "string" || html.length < 2) {
          if (typeof op.text === "string" && op.text.length >= 2) html = op.text;
          else { bad("missing reply text"); return; }
        }
        if (html.length > 20000) { bad("reply too long (max 20000 chars)"); return; }
        out.html = html;
        out.all = op.all === true;
      } else if (t === "compose_body") {
        if (!ctx || ctx.mode !== "compose") { bad("compose_body works only on an email being composed"); return; }
        var cb = op.html;
        if (typeof cb !== "string" || cb.length < 2) {
          if (typeof op.text === "string" && op.text.length >= 2) cb = op.text;
          else { bad("missing body text"); return; }
        }
        if (cb.length > 60000) { bad("body too long (max 60000 chars)"); return; }
        out.html = cb;
      } else if (t === "set_categories") {
        var names = op.names || op.categories;
        if (!Array.isArray(names) || !names.length || names.length > 6) { bad("missing names (1-6)"); return; }
        if (!names.every(function (nm) { return typeof nm === "string" && nm.length >= 1 && nm.length <= 64; })) {
          bad("bad category names"); return;
        }
        out.names = names.map(String);
      } else if (t === "forward") {
        if (!ctx || ctx.mode !== "read") { bad("forward works only on a received email (READ mode)"); return; }
        var note = typeof op.comment === "string" ? op.comment : (typeof op.html === "string" ? op.html : "");
        if (note.length > 8000) { bad("forward note too long (max 8000 chars)"); return; }
        var fwTo = Array.isArray(op.to) ? op.to.filter(function (x) { return typeof x === "string" && x; }) : [];
        if (fwTo.length > 50) { bad("too many recipients (max 50)"); return; }
        if (note) out.html = note;
        out.to = fwTo;
      } else if (t === "new_meeting") {
        if (typeof op.subject !== "string" || !op.subject.trim()) { bad("missing subject"); return; }
        out.subject = op.subject.trim().slice(0, 150);
        var emailOkM = function (arr) {
          return arr != null && Array.isArray(arr) && arr.every(function (e) {
            return typeof e === "string" && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.trim());
          });
        };
        var reqM = [], optM = [];
        if (op.required != null) {
          if (!emailOkM(op.required)) { bad("required must be an array of email addresses"); return; }
          reqM = op.required.map(function (e) { return e.trim(); }).slice(0, 50);
        }
        if (op.optional != null) {
          if (!emailOkM(op.optional)) { bad("optional must be an array of email addresses"); return; }
          optM = op.optional.map(function (e) { return e.trim(); }).slice(0, 50);
        }
        if (reqM.length || optM.length) { out.required = reqM; out.optional = optM; }
        var startM = new Date(String(op.start || "")), endM = new Date(String(op.end || ""));
        if (isNaN(startM.getTime()) || isNaN(endM.getTime())) {
          bad("start/end must be ISO date-times like 2026-11-02T14:00"); return;
        }
        out.start = String(op.start); out.end = String(op.end);
        if (typeof op.location === "string" && op.location.trim()) out.location = op.location.trim().slice(0, 200);
        if (typeof op.body === "string" && op.body.trim()) out.body = op.body.slice(0, 4000);
      } else if (t === "save_draft") {
        if (!ctx || ctx.mode !== "compose") { bad("save draft works only while composing"); return; }
      } else if (t === "new_message") {
        var toList = Array.isArray(op.to) ? op.to.filter(function (x) { return typeof x === "string" && x; }) : [];
        var nmHtml = op.html;
        if (typeof nmHtml !== "string" || nmHtml.length < 2) {
          if (typeof op.text === "string" && op.text.length >= 2) nmHtml = op.text;
          else { bad("missing message text"); return; }
        }
        if (toList.length > 50) { bad("too many recipients (max 50)"); return; }
        if (nmHtml.length > 60000) { bad("message too long (max 60000 chars)"); return; }
        out.to = toList;
        out.subject = typeof op.subject === "string" ? op.subject.slice(0, 200) : "";
        out.html = nmHtml;
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
    compose_body: "edit", save_draft: "add", create_template: "add", web_lookup: "edit", forward: "add", new_meeting: "add",
    delete_slide: "del", delete_shape: "del"
  };

  function fmtOp(op, ctx) {
    switch (op.type) {
      case "draft_reply": {
        var pr = op.html.length > 70 ? op.html.slice(0, 70) + "…" : op.html;
        return (op.all ? "Reply-ALL window (everyone on the thread) with: " : "Reply window with: ") + pr;
      }
      case "new_meeting": {
        var whoM = (((op.required || []).join(", ")) + ((op.required || []).length && (op.optional || []).length ? ", " : "") + ((op.optional || []).join(", "))) || "(you pick attendees)";
        return "Meeting form: " + op.subject + " - " + op.start + " to " + op.end + " - " + whoM;
      }
      case "forward": {
        var fwN = (op.html ? op.html : "").replace(/<[^>]*>/g, " ").trim();
        fwN = fwN.length > 60 ? fwN.slice(0, 60) + "\u2026" : fwN;
        var fwTo2 = (op.to && op.to.length) ? " -> " + op.to.join(", ") : " -> (you pick recipients)";
        return "Forward window for this mail" + fwTo2 + (fwN ? " - note: " + fwN : "");
      }
      case "compose_body":
        return "REPLACE the email you are composing (" + op.html.length + " chars) - review";
      case "set_categories":
        return "Add categories: " + op.names.join(", ");
      case "save_draft":
        return "Save the message being composed as draft";
      case "new_message": {
        var prN = op.html.length > 70 ? op.html.slice(0, 70) + "…" : op.html;
        var toN = (op.to && op.to.length) ? " -> " + op.to.join(", ") : "";
        return "New email window" + toN + (op.subject ? " '" + op.subject + "'" : "") + ": " + prN + " - review and send";
      }
    }
    return op.type;
  }

function applyOps(deck, ops, ctx) {
    var lines = [], manualNotes = [];
    var runOne = function (op) {
      switch (op.type) {
        case "draft_reply":
          return deck.draftReply(op.html, !!op.all).then(function () {
            lines.push({ ok: true, text: (op.all ? "reply-ALL window opened" : "reply window opened") + " with AI content - review and send" });
          });
        case "forward":
          return deck.forwardMessage(op, ctx).then(function () {
            lines.push({ ok: true, text: "forward window opened (quoted original" + (op.to && op.to.length ? ", " + op.to.length + " recipient(s)" : "") + ") - review and send" });
          });
        case "compose_body":
          return deck.composeBody(op.html).then(function () {
            lines.push({ ok: true, text: "compose body replaced (" + op.html.length + " chars)" });
          });
        case "set_categories":
          return deck.setCategories(op.names).then(function (r) {
            lines.push({ ok: true, text: "categories added: " + op.names.join(", ") + (r && r.note ? " " + r.note : "") });
          });
        case "save_draft":
          return deck.saveDraft().then(function () {
            lines.push({ ok: true, text: "draft saved" });
          });
        case "new_message":
          return deck.newMessage({ to: op.to, subject: op.subject, html: op.html }).then(function () {
            lines.push({ ok: true, text: "new email window opened - review and send" });
          });
        case "new_meeting":
          return deck.newMeeting(op).then(function () {
            lines.push({ ok: true, text: "meeting form opened (subject: " + op.subject + ") - review and send the invite yourself" });
          });
      }
      return Promise.resolve();
    };
    var chain = Promise.resolve();
    ops.forEach(function (op) {
      chain = chain.then(function () {
        return runOne(op).catch(function (err) {
          lines.push({ ok: false, text: fmtOp(op, ctx) + " - " + ((err && err.message) || String(err)) });
        });
      }, function (err) {
        lines.push({ ok: false, text: fmtOp(op, ctx) + " - " + ((err && err.message) || String(err)) });
        return Promise.resolve();
      });
    });
    return chain.then(function () {
      return { lines: lines, manualNotes: manualNotes };
    });
  }

  // ------------------------------------------------------------------ //
  // Chat / prompt flow
  // ------------------------------------------------------------------ //
  function formatContext(ctx) {
    var lines = [];
    lines.push("EMAIL CONTEXT (host: " + ctx.host + ", mode: " + ctx.mode + "):");
    if (ctx.sender) lines.push("FROM: " + ctx.sender);
    if (ctx.to) lines.push("TO: " + ctx.to);
    if (ctx.subject) lines.push("SUBJECT: " + ctx.subject);
    if (ctx.categories && ctx.categories.length) lines.push("CATEGORIES: " + ctx.categories.join(", "));
    if (ctx.attachments && ctx.attachments.length) {
      lines.push("ATTACHMENTS: " + ctx.attachments.map(function (a) {
        return a.name + (a.size ? " (" + (a.size > 1048576 ? (a.size / 1048576).toFixed(1) + " MB" : Math.round(a.size / 1024) + " kB") + ")" : "");
      }).join(", "));
    }
    var body = String(ctx.bodyText || "");
    if (body.length > 4000) body = body.slice(0, 4000) + "…[truncated]";
    lines.push("BODY:\n" + body);
    if (ctx.warnings && ctx.warnings.length) lines.push("(warnings: " + ctx.warnings.join("; ") + ")");
    if (ctx.masterCategories && ctx.masterCategories.length)
      lines.push("KNOWN CATEGORIES (mailbox master list): " + ctx.masterCategories.join(", "));
    else if (!ctx.minimal)
      lines.push("(category master list not readable at this permission level - propose sensible category names anyway)");
    if (ctx.minimal) lines.push("(minimal context: the full email body was not included - the 'include email context' box was unticked)");
    if (ctx.selectedItems && ctx.selectedItems.length) {
      lines.push("MULTI-SELECTED ITEMS (" + ctx.selectedItems.length + "):");
      ctx.selectedItems.forEach(function (si, k) {
        lines.push("  " + (k + 1) + ". " + (si.subject || "(no subject)") + (si.sender ? " - " + si.sender : ""));
      });
      lines.push("(for triage requests, work through these items one by one)");
    }
    return lines.join("\n");
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

  function runQuickReply(stance) {
    // reply + / reply - quick actions: a canned instruction runs through the
    // normal pipeline (the open email's conversation is attached as context);
    // the model proposes a draft_reply action with the AI-written text and
    // the usual approval gate applies (MailKick never sends by itself)
    if (SK.busy) return;
    var prof = currentProfile();
    if (!prof) { setStatus("No profile yet - create one on the Profiles tab first.", "error"); switchTab("settings"); switchSettingsSub("profiles"); return; }
    var inp = $("chat-input");
    if (!inp) return;
    var stanceText = stance === "pos"
      ? "positive and agreeable"
      : "politely negative (declining) but still courteous";
    inp.disabled = false;
    inp.value = "Draft a polite reply to this email conversation in the tone the conversation uses (formal or informal). Stance: " + stanceText +
      ". Reply entirely in the same language as the conversation, in the same tone (formal or informal). Then propose one draft_reply action with the AI-generated reply text (simple inline HTML only). Do not send anything.";
    runPrompt();
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
    // the deck/document/email readout refreshes automatically on every run
    updateDeckSummary();
    renderSimDeck();
    // the global auto-apply setting must tick the per-run box BEFORE the
    // all-auto evaluation happens (it used to tick inside the gate render)
    syncAutoApplyCheckbox();
    var wantCtx = $("chat-context").checked;
    setStatus("Collecting email context…", "busy");
    var ctxPromise = SK.deck.getContext({ minimal: !wantCtx }).catch(function (err) {
      setStatus("Could not read the email (" + err.message + ") - continuing without context.", "error");
      return null;
    });

    ctxPromise.then(function (ctx) {
      SK.lastCtx = ctx;
      var sys = SYSTEM_PROMPT + "\nHost: " + (SK.sim ? "MailKick simulator" : SK.deck.capabilities().host);
      SK.webResults = "";
      var webPromise = Promise.resolve("");
      return webPromise.then(function () {
      if (!(SK.cfg && SK.cfg.generateNotes)) sys += "\n" + NOTES_OFF_RULE;
      var ctxText = ctx ? formatContext(ctx) : "(no email could be read - make sure an email is open or selected and the reading pane stays open)";
      return getTemplateBlock(instruction).then(function (tplBlock) {
        var user = ctxText + "\n\n" + (tplBlock ? tplBlock + "\n" : "") + "\nUSER REQUEST:\n" + instruction;
        if (SK.webResults) user += "\n\nWEB RESULTS (live lookup, cite them when relevant):\n" + SK.webResults;
        var messages = [
          { role: "system", content: sys },
        ].concat(SK.history.slice(-(HISTORY_MAX_EXCHANGES * 2)))
         .concat([{ role: "user", content: user }]);
        dbg("prompt", { profile: prof.name, route: route, model: ref.model, connection: ref.connection, instruction: instruction, hadContext: !!ctx, templateUsed: !!tplBlock, history: SK.history.length, messages: messages });
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
      box2.textContent = cleanLiteralNL(parsed.text);
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
        noteSimWindows();
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
      mb.appendChild(document.createElement("h4")).textContent = "MailKick could not write this automatically - copy manually:";
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
      warn.textContent = "This Outlook build lacks the required mailbox APIs. MailKick can read but not modify. Update Microsoft 365 to enable edits.";
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
      // narrow pane: line 1 = action tag + checkbox + cancel (right aligned);
      // the description lands on its own line BELOW
      var line1 = document.createElement("span");
      line1.className = "oprowline1";
      r.appendChild(line1);
      var badge = document.createElement("span");
      badge.className = "optype " + row.op.type;
      badge.textContent = (OPS_META[row.op.type] || { label: row.op.type }).label;
      var desc = document.createElement("span");
      desc.className = "opdesc";
      var text = row.mode === "deny" ? fmtOp(row.op, ctx) + "  - blocked by approval policy"
        : row.mode === "invalid" ? (row.op && row.op.type ? row.op.type + ": " : "") + (row.reason || "invalid")
        : fmtOp(row.op, ctx);
      desc.textContent = text;
      line1.appendChild(badge);
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
        line1.appendChild(cb);
        line1.appendChild(lab);
        row._cb = cb;
      } else if (row.mode === "auto") {
        var note2 = document.createElement("span");
        note2.className = "opnote";
        note2.textContent = "auto";
        line1.appendChild(note2);
      } else if (row.mode === "deny") {
        var note3 = document.createElement("span");
        note3.className = "opnote";
        note3.textContent = "denied";
        line1.appendChild(note3);
      } else {
        var note4 = document.createElement("span");
        note4.className = "opnote";
        note4.textContent = row.reason || "";
        line1.appendChild(note4);
      }
      var x = document.createElement("button");
      x.type = "button";
      x.className = "op-x";
      x.textContent = "\u00D7";
      x.title = "Remove this action from the proposed list (second click confirms)";
      x.addEventListener("click", function () {
        if (x.dataset.armed !== "1") {
          x.dataset.armed = "1";   // confirm: second click within 3.5 s removes
          x.classList.add("armed");
          x.textContent = "sure?";
          setTimeout(function () {
            if (!x.isConnected) return;
            x.dataset.armed = "0"; x.classList.remove("armed"); x.textContent = "\u00D7";
          }, 3500);
          return;
        }
        row.removed = true;
        r.style.display = "none";
        updateApplyLabel();
      });
      line1.appendChild(x);
      tableRows.appendChild(r);
      r.appendChild(desc);   // the action explanation, below the action line
    });
    box.appendChild(tableRows);

    // horizontal divider between the action list and the "don't ask again"
    // session checkboxes (only when there are any)
    if (Object.keys(confirmTypes).length) {
      var gatehr = document.createElement("div");
      gatehr.className = "gate-divider";
      box.appendChild(gatehr);
    }

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
      // removed actions (X) no longer count towards the apply totals
      var liveAuto = rows.filter(function (r) { return !r.removed && r.mode === "auto"; }).length;
      var liveConfirm = rows.filter(function (r) { return !r.removed && r.mode === "confirm"; }).length;
      var approved = rows.filter(function (r) {
        return !r.removed && r.mode === "confirm" && r._cb && r._cb.checked;
      }).length;
      applyBtn.textContent = "Apply " + liveAuto + " auto" +
        (liveConfirm ? " + " + approved + "/" + liveConfirm + " approved" : "");
      applyBtn.disabled = (liveAuto + approved === 0) || !caps.text;
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
          if (r.removed) return;
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
        if (row.removed) return;   // removed with the X icon
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
        noteSimWindows();
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
      el.textContent = "Email: " + (ctx.subject || "(open an email)") + " - mode " + ctx.mode + " - " + ctx.host;
    }).catch(function (err) {
      el.textContent = "Document: unreadable (" + err.message + ")";
    });
  }

  function noteSimWindows() {
    if (!SK.sim || !SK.deck || !SK.deck.lastWindow) return;
    var w = SK.deck.lastWindow();
    if (!w) return;
    var out = $("chat-output");
    if (!out) return;
    var head, body;
    if (w.kind === "reply") {
      head = "Simulated reply window - real Outlook opens it for review (you still send it yourself)";
      body = w.html || "";
    } else if (w.kind === "meeting") {
      head = "Simulated meeting window - real Outlook opens the appointment form (you send it yourself)";
      body = "<p><b>" + esc(w.subject || "") + "</b><br />" + esc(w.when || "") +
        (w.location ? " - " + esc(w.location) : "") + "</p>" +
        (w.attendees ? "<p><b>Attendees:</b> " + esc(w.attendees) + "</p>" : "") + (w.body || "");
    } else {
      head = "Simulated forward window - FW: " + esc(w.subject || "") + " - real Outlook opens it for review";
      body = (w.to && w.to.length ? "<p><b>To:</b> " + esc(w.to.join(", ")) + "</p>"
              : "<p><b>To:</b> (you pick recipients in Outlook)</p>") + (w.note ? w.note : "");
    }
    var card = document.createElement("div");
    card.className = "simwin";
    card.innerHTML = '<div class="simwin-head">' + esc(head) + "</div>" +
      '<div class="simwin-body">' + body + "</div>";
    out.appendChild(card);
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
    if (id === "templates") renderTemplates();
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
    // narrow pane: one block per connection, info stacked below each other
    var list = $("conn-list");
    var blocks = [];
    (SK.cfg.connections || []).forEach(function (c) {
      var used = (SK.cfg.profiles || []).filter(function (p) {
        return (p.content && p.content.connection === c.name) ||
               (p.other && p.other.connection === c.name);
      }).map(function (p) { return p.name; });
      // click the name to edit; deletion lives in the edit form (with confirm)
      blocks.push(
        '<div class="connblock">' +
          '<div class="connline">' +
            '<button class="namebtn" data-edit="' + esc(c.name) + '" title="edit">' + esc(c.name) + "</button>" +
            '<button class="btn" data-test="' + esc(c.name) + '">Test</button>' +
            '<span class="opnote" data-testmsg="' + esc(c.name) + '"></span>' +
          "</div>" +
          '<div class="connfield"><span class="connlabel">type</span>' + esc(c.type) + "</div>" +
          '<div class="connfield"><span class="connlabel">url</span>' + esc(c.baseUrl || "-") + "</div>" +
          '<div class="connfield"><span class="connlabel">key</span>' + (c.hasKey ? "set" : "not set") + "</div>" +
          (used.length ? '<div class="connfield"><span class="connlabel">in use</span><span class="opnote">' + esc(used.join(", ")) + "</span></div>" : "") +
        "</div>");
    });
    list.innerHTML = blocks.join("") || '<div class="opnote">No connections yet - add one below.</div>';
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
    // narrow pane: one block per action - action + description on top,
    // the three policy options below them
    var table = $("approvals-table");
    var blocks = [];
    OP_TYPES.forEach(function (t) {
      var mode = (SK.cfg.approvals || {})[t] || "confirm";
      var radios = ["auto", "confirm", "deny"].map(function (m) {
        return '<label class="apvopt"><input type="radio" name="appr-' + esc(t) + '" value="' + m + '"' +
          (m === mode ? " checked" : "") + " /> " + m + "</label>";
      }).join("");
      blocks.push(
        '<div class="apvblock">' +
          '<div class="apvhead"><b>' + esc(OPS_META[t].label) + "</b>" +
            '<span class="muted">' + esc(OPS_META[t].hint) + "</span>" +
          "</div>" +
          '<div class="apvopts">' + radios + "</div>" +
        "</div>");
    });
    table.innerHTML = blocks.join("");
    $("settings-debug").checked = !!SK.cfg.debug;
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

  function saveAutoClean() {
    var days = parseInt($("settings-autoclean-days").value, 10);
    if (!days || days < 1 || days > 3650) {
      setFormMsg($("settings-autoclean-msg"), "Days must be between 1 and 3650.");
      return;
    }
    apiJson("/api/settings", "PUT", { autoClean: $("settings-autoclean").checked, autoCleanDays: days })
      .then(function () { return loadConfigAndRender(); })
      .then(function () {
        setFormMsg($("settings-autoclean-msg"),
          SK.cfg.autoClean
            ? "Auto-clean ON - sessions newer than " + SK.cfg.autoCleanDays + " day(s) are kept."
            : "Auto-clean off - nothing is deleted.", true);
      }).catch(function (err) { setFormMsg($("settings-autoclean-msg"), err.message); });
  }
  function saveDebugOption() {
    apiJson("/api/settings", "PUT", { debug: $("settings-debug").checked }).then(function () {
      return loadConfigAndRender();
    }).then(function () {
      setFormMsg($("settings-msg"), SK.cfg.debug ? "Debug logging ON - writing to the daily log." : "Debug logging off.", true);
    }).catch(function (err) { setFormMsg($("settings-msg"), err.message); });
  }

  // ------------------------------------------------------------------ //
  // About tab
  // ------------------------------------------------------------------ //
  function paneBuildLabel() {
    var build = window.__SK_BUILD;
    if (!build) return "(unknown - reload the pane)";
    var server = SK.cfg && SK.cfg.version ? String(SK.cfg.version) : null;
    if (server && server !== String(build)) {
      return "v" + build + " - STALE, close and reopen the pane (server runs v" + server + ")";
    }
    return "v" + build + (server ? " (current)" : "");
  }

  function renderAbout() {
    var kv = [
      ["Corelan MailKick add-in", SK.cfg ? "v" + (SK.cfg.version || "0.1.3") : "0.1.3"],
      ["Server", "v" + (SK.cfg ? SK.cfg.version || "0.1.0" : "?") + " on 127.0.0.1:" + (SK.cfg ? SK.cfg.port : "?")],
      ["Pane build", paneBuildLabel()],
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
      // the auto-apply checkbox was wired before the config arrived - keep it
      // in sync with the stored setting from now on (the "off after reopen" bug)
      var sap2 = $("set-autoapply");
      if (sap2 && !SK.autoApplyUserTouched) sap2.checked = autoApplyWanted();
      syncAutoApplyCheckbox();   // and the per-run box follows the global too
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
      $("statusbar").textContent = "MailKick v" + (cfg.version || "0.1.0") +
        (window.location && String(window.location.protocol) === "https:" ? " - 🔒 " : " - ") +
        "server 127.0.0.1:" + cfg.port + (cfg.debug ? " - debug ON" : "");
      return cfg;
    });
  }

  function wireEvents() {
    Array.prototype.forEach.call(document.querySelectorAll("#tabs button"), function (b) {
      b.addEventListener("click", function () { switchTab(b.dataset.tab); });
    });

    $("chat-run").addEventListener("click", runPrompt);
    var rpos = $("reply-pos"), rneg = $("reply-neg");
    if (rpos) rpos.addEventListener("click", function () { runQuickReply("pos"); });
    if (rneg) rneg.addEventListener("click", function () { runQuickReply("neg"); });
    $("chat-input").addEventListener("keydown", function (ev) {
      if ((ev.metaKey || ev.ctrlKey) && ev.key === "Enter") { ev.preventDefault(); runPrompt(); }
    });
    $("chat-new").addEventListener("click", function () {
      resetHistory();
      $("chat-output").innerHTML = "";
      $("chat-input").value = "";
      setStatus("Conversation history cleared.", "ok");
    });
    $("chat-refresh-deck").addEventListener("click", function () { updateDeckSummary(); setStatus("Email refreshed.", ""); });

    $("conn-new").addEventListener("click", function () {
      resetConnForm();
      var nf = $("conn-name");
      if (nf) nf.focus();
    });
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
    $("settings-autoclean-save").addEventListener("click", saveAutoClean);
    var updBtn = $("upd-check");
    if (updBtn) updBtn.addEventListener("click", updCheck);
    wireAutoApplySetting();
    var updAuto = $("upd-auto");
    if (updAuto) {
      // remember the preference (on by default); the boot check honors it
      apiJson("/api/config", "GET").then(function (cfg) {
        updAuto.checked = !(cfg && cfg.autoUpdateCheck === false);
        var acBox = $("settings-autoclean"), acDays = $("settings-autoclean-days");
        if (acBox && acDays) {
          acBox.checked = !(cfg && cfg.autoClean === false);
          acDays.value = String((cfg && cfg.autoCleanDays) || 120);
        }
        if (updAuto.checked) updCheck();   // silent availability check at boot
      }, function () { updAuto.checked = true; updCheck(); });
      updAuto.addEventListener("change", function () {
        apiJson("/api/settings", "PUT", { autoUpdateCheck: !!updAuto.checked });
      });
    }
    updateSessionLabel();
    wireSessions();
    wireTemplates();
    $("req-clear").addEventListener("click", clearRequests);
    $("req-refresh").addEventListener("click", renderRequests);
    $("sim-reset").addEventListener("click", function () {
      window.SimDeck.resetAll();
      renderSimDeck();
      updateDeckSummary();
      setStatus("Simulator email reset.", "");
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
      badge.textContent = "MAIL";
      SK.deck = new MailDeck();
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
    setStatus("Cannot reach the OfficeKick server: " + err.message + " Is it running? (install.sh / install.ps1 / LaunchAgent)", "error");
    $("statusbar").textContent = "server unreachable: " + err.message;
  }

  function boot() {
    var params = new URLSearchParams(String(window.location.search || ""));
    SK.sim = window.__MAILKICK_SIM__ === true || window.__SK_SIM__ === true || params.has("sim");
    SK.forceSave = /[?&]save=1/.test(window.location.search);
    SK.testMode = (SK.sim || /\?.*nosave=1/.test(window.location.search)) && !SK.forceSave;
    if (SK.sim) {
      $("host-badge").style.display = "";
      $("host-badge").textContent = "SIM";
      var simTab = $("sim-panel");
      simTab.style.display = "";
      var nav = document.createElement("button");
      nav.textContent = "Sim email";
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
