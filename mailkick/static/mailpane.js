/*
 * MailKick task pane (runs inside Microsoft Outlook via Office.js, or in the
 * browser via sim.js for offline testing).
 *
 * Pipeline: free-text instruction -> profile/route -> /api/chat -> reasoning
 * stripping -> JSON ops extraction (hidden repair round on unclear output)
 * -> validation against deck context -> approval gate -> Office.js apply.
 */
(function () {

  // PANIC TRAP (0.9.182): every uncaught pane error (window.onerror /
  // unhandledrejection) reports to the local server -> a timestamped,
  // app-tagged JSON line in ~/.officekick/logs/pane-errors-<day>.log.
  // Installed FIRST so even boot/wiring crashes are reported; never throws.
  var _errLastMsg = "", _errLastT = 0;
  function _paneErr(kind, msg, extra) {
    try {
      var t = +new Date();
      if (msg === _errLastMsg && t - _errLastT < 4000) return;   // 4 s dedupe
      _errLastMsg = msg; _errLastT = t;
      var payload = {
        kind: String(kind || "uncaught"),
        message: String(msg == null ? "" : msg).slice(0, 800),
        stack: String((extra && extra.stack) || "").slice(0, 2000),
        at: String((extra && extra.fileName) || "").slice(0, 200),
        line: (extra && extra.lineNumber) || 0,
        ua: String(navigator.userAgent).slice(0, 200),
        pane: String(window.location.pathname).slice(0, 120),
        href: String(window.location.href).slice(0, 200)
      };
      var xr = new XMLHttpRequest();
      xr.open("POST", "/api/pane-error", true);
      xr.setRequestHeader("Content-Type", "application/json");
      xr.send(JSON.stringify(payload));
    } catch (eP) { /* the trap must never break the pane */ }
  }
  window.onerror = function (m, s, l, c, e) {
    _paneErr("uncaught", m + " @" + (s || "?") + ":" + l, e);
  };
  window.addEventListener("unhandledrejection", function (ev) {
    var r = ev && ev.reason;
    _paneErr("rejection", String((r && (r.message || r.stack)) || r), (r instanceof Error) ? r : "");
  });

  "use strict";

  var $ = function (id) { return document.getElementById(id); };

  var SK = {
    app: "mailkick",
    sessionId: null,
    sessionName: "",
    sessionSaveTimer: null,
    profImageFilter: true,
    profWebFilter: true,    // the web-slot model select shows web-capable models by default
    sim: false,
    deck: null,
    cfg: null,
    busy: false,
    runCancelled: false,     // cancel button: interrupt the current run
    titleBusy: false,        // a session-title AI update is in flight
    activeFetch: null,       // AbortController of the in-flight pane request
    modelsCache: {},     // connName -> {list:[...], at:ms}
    sessionApprovals: {},// opType -> 'auto' when approved for this session
    sessionManual: false,  // the user renamed this session: pause AI titles
    history: [],         // conversation turns shown to the model (session-only)
    editingConn: null,
    editingProf: null,
    connBaseDirty: false,
    lastCtx: null,
    activeTab: "chat",
    mailEmailId: null,            // 0.9.233: identity (itemId) of the open email
    mailStaleResumeDropUntil: 0,  // one-shot window in which a boot resume is dropped
    mnsUserTouched: false,        // user changed the new-session-per-email toggle
    clarifyAnswer: null,          // 0.9.235: pending user pick from a clarify round
    clarifyChain: 0,              // clarify rounds served since the original ask
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
    if (SK.actRunOn && text) actLog(text);   // the activity box mirrors the status trail
  }
  // Run activity log (0.9.169): high-level trail of what the add-in does
  // while it works with the model - sent requests, lookups, page reads and
  // applies. Most recent first; cleared on every Run press.
  function actLog(text) {
    var box = $("chat-activity");
    if (!box) return;
    var t = new Date();
    var hh = ("0" + t.getHours()).slice(-2) + ":" + ("0" + t.getMinutes()).slice(-2) + ":" + ("0" + t.getSeconds()).slice(-2);
    var row = document.createElement("div");
    row.className = "actrow";
    var ts = document.createElement("span");
    ts.className = "actts";
    ts.textContent = hh;
    var tx = document.createElement("span");
    tx.className = "acttx";
    tx.textContent = text;
    row.appendChild(ts);
    row.appendChild(tx);
    box.insertBefore(row, box.firstChild);
    while (box.children.length > 80) box.removeChild(box.lastChild);
    box.style.display = "";
  }
  function actReset() {
    SK.actRunOn = true;   // only run-time statuses mirror; boot/idle ones never do
    var box = $("chat-activity");
    if (box) { box.innerHTML = ""; box.style.display = ""; }
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
    // boot-time lines (pane.boot, theme) fire BEFORE /api/config arrives ---
    // queue them and replay once the config (and its debug flag) lands, so
    // early diagnostics never get swallowed
    try {
      if (!SK.cfg || !SK.cfg.debug) {
        SK.dbgQueue = SK.dbgQueue || [];
        if (SK.dbgQueue.length < 16) SK.dbgQueue.push({ category: category, data: data });
        return;
      }
      if (SK.dbgQueue && SK.dbgQueue.length) {
        var q = SK.dbgQueue.splice(0);
        SK.dbgQueue = null;
        q.forEach(function (qe2) { dbgFlush(qe2.category, qe2.data); });
      }
      dbgFlush(category, data);
    } catch (e) { /* never break the pane over logging */ }
  }
  function dbgFlush(category, data) {
    fetch("/api/debug", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ category: "ui." + category, data: data }),
    }).catch(function () {});
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
  // fetch a source URL's page text via the server (the pane's research
  // channel: search results, then the DOCUMENTS they point at)
  function fetchPageText(url, from) {
    return apiJson("/api/fetch-url?url=" + encodeURIComponent(url) + (from > 0 ? "&from=" + parseInt(from, 10) : ""), "GET").then(function (r) {
      return (r && r.ok) ? r : null;
    }).catch(function () { return null; });
  }
  function gatherPageContents(factsText) {
    // up to 2 distinct http(s) URLs from the lookup facts, each = page text
    var urls = [];
    var re = /https?:\/\/[^\s)\]]+/g, m;
    while ((m = re.exec(String(factsText || ""))) !== null) {
      var u = m[0].replace(/[.,;'"]+$/, "");
      if (urls.indexOf(u) < 0) urls.push(u);
      if (urls.length >= 2) break;
    }
    return Promise.all(urls.map(function (u) { return fetchPageText(u); })).then(function (pages) {
      var blocks = "";
      pages.forEach(function (p) {
        if (p && p.text) {
          blocks += "\n\n--- PAGE: " + (p.title || p.url) + " (" + p.url + ") ---\n" + p.text;
        }
      });
      return blocks;
    });
  }
  // ------------------------- search profiles ------------------------- //
  function renderSearchProfiles() {
    if (!$("srp-steps")) return;
    // the key-state labels follow the fresh config too
    ["brave", "tavily", "firecrawl"].forEach(function (eng) { renderSearchKeyState(eng); });
    SK.srpDraft = null;   // rebuild the working copy of the ONE shared sequence
    currentSearchProfile();
    renderSearchSteps();
  }

  function currentSearchProfile() {
    if (!SK.srpDraft) {
      SK.srpDraft = { name: "", steps: (SK.cfg && SK.cfg.searchSteps || []).map(function (s) { return JSON.parse(JSON.stringify(s)); }) };
    }
    return SK.srpDraft;
  }

  function buildStepEditor(host, st) {
    host.innerHTML = "";
    if (st.kind === "server") {
      var engines = ["duckduckgo", "brave", "tavily", "firecrawl"];
      var es = document.createElement("select");
      es.innerHTML = engines.map(function (e) { return '<option value="' + e + '">' + e + "</option>"; }).join("");
      es.value = engines.indexOf(st.engine) >= 0 ? st.engine : "duckduckgo";
      es.addEventListener("change", function () { st.engine = es.value; });
      host.appendChild(es);
    } else {
      var conns = (SK.cfg && SK.cfg.connections) || [];
      var cs = document.createElement("select");
      cs.innerHTML = conns.map(function (c) { return '<option value="' + esc(c.name) + '">' + esc(c.name) + " (" + esc(c.type) + ")</option>"; }).join("");
      cs.value = st.connection || (conns.length ? conns[0].name : "");
      if (!st.connection && conns.length) st.connection = cs.value;   // default = written back at build time
      cs.addEventListener("change", function () { st.connection = cs.value; st.model = ""; renderSearchSteps(); });
      host.appendChild(cs);
      var ms = document.createElement("select");
      ms.id = "srp-step-model-" + (st.connection || "x");
      host.appendChild(ms);
      var hint = document.createElement("span");
      hint.className = "opnote";
      hint.textContent = connLookupCapability(cs.value);
      host.appendChild(hint);
      fillWebModelSelect(st, cs, ms);
    }
    return host;
  }

  function fillWebModelSelect(st, cs, ms) {
    if (!ms) {
      ms = document.createElement("select");   // not used
    }
    ms.innerHTML = '<option value="">(connection default)</option>';
    if (cs.value) {
      modelsCacheGet(cs.value, true, "web").then(function (list) {
        (list || []).forEach(function (m) {
          var o = document.createElement("option");
          o.value = m; o.textContent = m;
          ms.appendChild(o);
        });
        if (st.model && list.indexOf(st.model) >= 0) ms.value = st.model;
      }).catch(function () {});
    }
    ms.addEventListener("change", function () { st.model = ms.value; });
  }

  function renderSearchSteps() {
    var box = $("srp-steps");
    if (!box) return;
    var p = currentSearchProfile();
    var steps = (p && p.steps) || [];
    box.innerHTML = "";
    steps.forEach(function (st, i) {
      var row = document.createElement("div");
      row.className = "row row-tight";
      var grip = document.createElement("span");
      grip.className = "dragh";
      grip.title = "Drag to reorder";
      grip.draggable = true;
      grip.addEventListener("dragstart", function (e) {
        SK.dragStep = i;
        try { e.dataTransfer.setData("text/plain", String(i)); e.dataTransfer.effectAllowed = "move"; } catch (e2) {}
        row.classList.add("dragging");
      });
      grip.addEventListener("dragend", function () {
        row.classList.remove("dragging");
        SK.dragStep = null;
      });
      row.addEventListener("dragover", function (e) {
        if (SK.dragStep === null || SK.dragStep === undefined) return;
        var from = SK.dragStep;
        if (from === i) return;
        e.preventDefault();
        try { e.dataTransfer.dropEffect = "move"; } catch (e2) {}
        var r2 = row.getBoundingClientRect();
        var above = (e.clientY - r2.top) < r2.height / 2;
        row.classList.toggle("drop-above", above);
        row.classList.toggle("drop-below", !above);
      });
      row.addEventListener("dragleave", function () { row.classList.remove("drop-above", "drop-below"); });
      row.addEventListener("drop", function (e) {
        e.preventDefault();
        row.classList.remove("drop-above", "drop-below");
        var from = SK.dragStep === undefined ? null : SK.dragStep;
        if (from === null || from === undefined) return;
        SK.dragStep = null;
        var r2 = row.getBoundingClientRect();
        var above = (e.clientY - r2.top) < r2.height / 2;
        var to = above ? i : i + 1;
        if (to === from || to === from + 1) return;   // dropped in place
        var moved = p.steps.splice(from, 1)[0];
        if (to > from) to -= 1;
        p.steps.splice(to, 0, moved);
        renderSearchSteps();
      });
      row.appendChild(grip);
      var n = document.createElement("span");
      n.className = "opnote";
      n.textContent = (i + 1) + ".";
      row.appendChild(n);
      var kindSel = document.createElement("select");
      kindSel.innerHTML = '<option value="server">Search engine</option><option value="online">Model lookup</option>';
      kindSel.value = st.kind;
      kindSel.addEventListener("change", function () {
        st.kind = kindSel.value;
        // switching the type refreshes the row's option selects; normalize the
        // fields so a switched step stays a valid {server|online} step
        if (st.kind === "server") {
          delete st.connection; delete st.model;
          if (!st.engine) st.engine = "duckduckgo";
        } else {
          delete st.engine;
          if (!st.connection) st.connection = (((SK.cfg || {}).connections || [])[0] || {}).name || "";
          if (typeof st.model !== "string") st.model = "";
        }
        renderSearchSteps();
      });
      row.appendChild(kindSel);
      row.appendChild(buildStepEditor(document.createElement("span"), st));
      var del = document.createElement("button");
      del.className = "btn subtle danger";
      del.type = "button";
      del.textContent = "\u00d7";
      del.title = "Remove this step";
      del.addEventListener("click", function () { p.steps.splice(i, 1); renderSearchSteps(); });
      row.appendChild(del);
      box.appendChild(row);
    });
    if (!steps.length) {
      var empty = document.createElement("div");
      empty.className = "opnote";
      empty.textContent = "(no steps yet - add one below)";
      box.appendChild(empty);
    }
  }

  function connLookupCapability(connName) {
    var c = ((SK.cfg && SK.cfg.connections) || []).filter(function (x) { return x.name === connName; })[0];
    if (!c) return "(unknown connection)";
    if (c.type === "openrouter") return "pin a model - :online attaches to it (empty = step skipped)";
    if (c.type === "openai" || c.type === "generic-openai") return "needs a search-preview model";
    return "no native web search - this step skips itself";
  }

  function wireSearchProfiles() {
    if (!$("srp-steps")) return;
    $("srp-add").addEventListener("click", addSearchStepFn({ kind: "server", engine: "duckduckgo" }));
    $("srp-save").addEventListener("click", saveSearchProfile);
    wireSearchKeys();
  }

  // per-engine API keys: type to replace, leave blank to clear (never stored
  // in the pane or echoed back - only a set/unset marker is public)
  function wireSearchKeys() {
    var engines = ["brave", "tavily", "firecrawl"];
    engines.forEach(function (eng) {
      var inp = $("searchkey-" + eng), st = $("searchkey-" + eng + "-state");
      if (!inp || !st) return;
      renderSearchKeyState(eng);
      inp.addEventListener("input", function () { renderSearchKeyState(eng); });
    });
    var btn = $("searchkeys-save");
    if (!btn) return;
    btn.addEventListener("click", function () {
      var payload = {};
      engines.forEach(function (eng) {
        var inp2 = $("searchkey-" + eng);
        if (inp2) payload[eng] = inp2.value || null;   // blank = clear
      });
      apiJson("/api/settings", "PUT", { searchKeys: payload }).then(function () {
        engines.forEach(function (eng) {
          var inp3 = $("searchkey-" + eng);
          if (inp3) inp3.value = "";
        });
        flashSaved("saved-searchkeys", "\u2713 saved");
        return loadConfigAndRender();
      }).then(function () {
        engines.forEach(function (eng) { renderSearchKeyState(eng); });   // state text = after the fresh cfg
      }).catch(function (e) { flashSaved("saved-searchkeys", "save failed"); });
    });
  }

  function renderSearchKeyState(eng) {
    var inp = $("searchkey-" + eng), st = $("searchkey-" + eng + "-state");
    if (!inp || !st) return;
    var set = !!(SK.cfg && SK.cfg.searchKeysSet && SK.cfg.searchKeysSet[eng]);
    var typing = !!inp.value;
    st.textContent = typing ? "replacing key" : (set ? "key set" : "no key");
    st.className = typing ? "opnote" : (set ? "opnote keystate-ok" : "opnote keystate-no");
  }

  function addSearchStepFn(st) {
    return function () {
      var p = currentSearchProfile();   // creates the working copy lazily
      st = JSON.parse(JSON.stringify(st));
      if (st.kind === "online") st.connection = (((SK.cfg || {}).connections || [])[0] || {}).name || "";
      p.steps.push(st);
      renderSearchSteps();
    };
  }

  function saveSearchProfile() {
    var p = currentSearchProfile();
    var steps = (p && p.steps) || [];
    apiJson("/api/settings", "PUT", { searchSteps: steps }).then(function () {
      flashSaved("saved-srp", "\u2713 saved");
      return loadConfigAndRender();
    }).then(function () {
      SK.srpDraft = null;
      currentSearchProfile();
      renderSearchSteps();
      $("srp-status").textContent = "";
    }).catch(function (e) {
      $("srp-status").textContent = "Save failed: " + ((e && e.message) || e);
    });
  }

  // the RUN order = the selected search profile's steps; no profile = the
  // free DuckDuckGo engine. An online step checks the connection's actual
  // capability and skips itself (no model = no search).
  function buildSearchLadder(prof) {
    var attempts = [];
    var skipped = [];
    var savedSteps = (SK.cfg && SK.cfg.searchSteps) || [];
    var steps = savedSteps.length ? savedSteps : [{ kind: "server", engine: "duckduckgo" }];
    steps.forEach(function (st) {
      if (st.kind === "server") {
        var eng = st.engine || "duckduckgo";
        if (eng !== "duckduckgo" && !(SK.cfg && SK.cfg.searchKeysSet && SK.cfg.searchKeysSet[eng])) {
          // no key = the step cannot run - skip it like a modelless lookup step
          skipped.push(eng + " (no API key set)");
          return;
        }
        attempts.push({ server: true, engine: eng, label: "server:" + eng });
      } else {
        var c = ((SK.cfg && SK.cfg.connections) || []).filter(function (x) { return x.name === st.connection; })[0];
        if (!c) { skipped.push(String(st.connection) + " (connection not found)"); return; }
        if (c.type === "ollama" || c.type === "llamacpp") { skipped.push(String(st.connection) + " (no native web search)"); return; }
        var model = String(st.model || "");
        if ((c.type === "openai" || c.type === "generic-openai") && !/search-preview/i.test(model)) {
          // no native searching model selected: skip this step
          skipped.push(String(st.connection) + " (no search-preview model selected)");
          return;
        }
        // ":online" cannot run bare - OpenRouter needs a base model id; a
        // step with no model pinned = "no search model = no search" -> skip
        if (c.type === "openrouter" && !model) { skipped.push(String(st.connection) + " (no model pinned)"); return; }
        var online = c.type === "openrouter" ? model + ":online" : model;
        if (!online) return;
        attempts.push({ connection: st.connection, model: online, label: st.connection + " : " + online });
      }
    });
    if (!savedSteps.length && ((prof.other || {}).connection)) {
      // free-only is the DEFAULT when no steps are configured; still try the
      // provider's own channel after
      var own = ((SK.cfg && SK.cfg.connections) || []).filter(function (x) { return x.name === prof.other.connection; })[0];
      if (own && own.type === "openrouter" && prof.other.model) {
        attempts.push({ connection: prof.other.connection, model: prof.other.model + ":online" });
      }
    }
    if (!attempts.length) attempts.push({ server: true, engine: "duckduckgo", label: "server:duckduckgo" });
    return { attempts: attempts, skipped: skipped, profileUsed: savedSteps.length ? "steps" : "" };
  }

  function webLookup(prof, instruction) {
    var ref = prof.other && prof.other.connection && prof.other.model ? prof.other : prof.content;
    var sys = "You have LIVE WEB ACCESS. Research the user's request with web search now. " +
      "Return 4-8 short facts that help answer it; after EACH fact put its source URL in ( ): " +
      "like \"fact text (https://source.example/page)\". Be current; mention dates when known. " +
      "If you actually cannot access the web, reply with exactly NO_WEB_ACCESS.";
    var connOf = function (name) {
      return (SK.cfg.connections || []).filter(function (c) { return c.name === name; })[0] || {};
    };
    // capability ladder: try the profile's own model first (with OpenRouter's
    // :online web plugin when the connection is OpenRouter); if this profile's
    // connection has no web plugin, fall back to ANY OpenRouter connection
    // with a cheap always-available web-enabled model. The pane owns the
    // lookup channel - the model the user chats with often has none.
    // capability ladder: OpenRouter's :online plugin first (native search);
    // then the SERVER-SIDE keyless search (works for every connection kind);
    // then an OpenAI search-preview model (native search) when present.
    var ladder = buildSearchLadder(prof);
    var attempts = ladder.attempts;
    if (ladder.skipped && ladder.skipped.length) actLog("Lookup steps skipped - " + ladder.skipped.join(", "));    // (the explicit single-engine unshift is retired: the search profile
    // tab owns the sequence now - see buildSearchLadder)
    var webNoHit = function (t) {
      return !t || /^NO_WEB_ACCESS\b/i.test(t) || /cannot access (the )?web/i.test(t) ||
        /can'?t browse/i.test(t) || /no web access/i.test(t) ||
        /lookup channel isn'?t available/i.test(t);
    };
    var one = function (a) {
      actLog(a && a.server ? "Web search via " + (a.engine || "duckduckgo")
                            : "Model lookup via " + (a.connection || "?") + (a.model ? " (model " + a.model + ")" : ""));
      if (a && a.server) {
        return apiJson("/api/web-search?q=" + encodeURIComponent(instruction.slice(0, 240)) + "&engine=" + encodeURIComponent(a.engine || "duckduckgo"), "GET")
          .then(function (res) {
            if (!res || !res.ok || !(res.results || []).length) return "";
            var lines = (res.results || []).slice(0, 8).map(function (r) {
              return "- " + r.title + (r.snippet ? " - " + r.snippet : "") + "\n  " + r.url;
            }).join("\n");
            return "SERVER-SIDE WEB SEARCH (live, DuckDuckGo via the local server) for the lookup:\n" + lines;
          }).catch(function () { return ""; });
      }
      return askRetryable(function () {
        return apiJson("/api/chat", "POST", {
          connection: a.connection, model: a.model,
          messages: [{ role: "system", content: sys },
                     { role: "user", content: instruction.slice(0, 600) }],
        });
      }).then(function (resp) {
        var t = String(resp.text || "").trim();
        if (webNoHit(t)) return "";
        return "MODEL-SIDE WEB LOOKUP via " + a.model + ":\n" + t.slice(0, 5000);
      }).catch(function () { return ""; });  // this model can't serve a lookup - try the next
    };
    return attempts.reduce(function (acc, a) {
      return acc.then(function (got) { return got ? got : one(a); });
    }, Promise.resolve(""));
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

  function setUpdateFlag(latest) {
    // 0.9.321: glowing "Update available" marker at the bottom of the pane,
    // just above the version/server line - removed (footer back to its
    // original single-line size) as soon as a check reports up to date
    var foot = document.querySelector(".sk-footer");
    if (!foot) return;
    var el = document.getElementById("update-flag");
    if (!latest) {
      if (el) { el.parentNode.removeChild(el); }
      return;
    }
    if (!el) {
      el = document.createElement("div");
      el.id = "update-flag";
      foot.insertBefore(el, foot.firstChild);
    }
    el.textContent = "\u2B06\uFE0F Update available: " + latest +
      (SK.cfg && SK.cfg.version ? " (you run " + SK.cfg.version + ")" : "");
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
        setUpdateFlag(null);
        info.className = "opnote";
        info.textContent = "You are up to date (" + data.current + ", latest " + data.latest + ").";
      } else {
        setUpdateFlag(data.latest);
        info.className = "opnote updavail";   // bold; gold/purple per theme
        info.textContent = "UPDATE AVAILABLE: " + data.latest + " (you run " + data.current + "). " +
          (data.howToUpdate || "");
        if (data.pullAvailable) {
          var pullBtn = document.createElement("button");
          pullBtn.className = "btn";
          pullBtn.textContent = "Update now (pull + redeploy + restart)";
          pullBtn.addEventListener("click", function () {
            pullBtn.disabled = true;
            info.textContent = "Starting the full update (pull + redeploy + restart)\u2026";
            apiJson("/api/update-run", "POST", {}).then(function (ur) {
              if (!(ur && ur.ok)) {
                pullBtn.disabled = false;
                info.className = "opnote updavail";
                info.textContent = "Update could not start: " + ((ur && ur.error) || "unknown error");
                return;
              }
              info.className = "opnote";
              info.textContent = "Update running (pull + redeploy + restart) - the server drops mid-run and comes back on the new build\u2026";
              var oldGen = "";
              apiJson("/api/health", "GET").then(function (h0) {
                oldGen = (h0 && h0.gen) || "";
              }).catch(function () { }).then(function () {
                var tries = 0;
                var poll = function () {
                  apiJson("/api/health", "GET").then(function (h1) {
                    if (h1 && h1.gen && h1.gen !== oldGen) {
                      info.className = "opnote";
                      info.textContent = "Update done - the server is on " + (h1.version || "the new build") +
                      setUpdateFlag(null);
                        ". Reload this pane (About tab -> Reload pane, or quit and reopen the app) to load the new pane files.";
                    } else if (h1 && h1.gen && h1.gen === oldGen && tries > 12) {
                      info.className = "opnote updavail";
                      info.textContent = "The update did not restart the server (still on " + (h1.version || "the old build") + ")."
                      info.textContent += " Check install.log in the clone for the cause. Updates complete even with Office apps open (0.9.309+: files copy + server restart; only the WEF re-registration is skipped) - on an older build, quit all Office apps completely and click Update again.";
                    } else if (++tries > 150) {
                      info.className = "opnote updavail";
                      info.textContent = "Update is taking unusually long or the server did not come back - check install.log in " +
                        (ur.repoPath || "the clone") + " plus update_run.log in the config folder.";
                    } else { setTimeout(poll, 1000); }
                  }).catch(function () {
                    // the old server drops mid-run; keep waiting for the new one
                    if (++tries > 150) {
                      info.className = "opnote updavail";
                      info.textContent = "Update is taking unusually long or the server did not come back - check install.log in " +
                        (ur.repoPath || "the clone") + " plus update_run.log in the config folder.";
                    } else { setTimeout(poll, 1000); }
                  });
                };
                setTimeout(poll, 2000);
              });
            }).catch(function (e2) {
              pullBtn.disabled = false;
              info.className = "opnote updavail";
              info.textContent = "Update could not start: " + ((e2 && e2.message) || e2);
            });
          });
          info.appendChild(document.createElement("br"));
          info.appendChild(pullBtn);
        }
      }
    }).catch(function (e) {
      btn.disabled = false;
      info.textContent = "could not check for updates: " + (e && e.message ? e.message : e);
    });
  }
    var rsBtn2 = $("btn-restart-server");
    if (rsBtn2) rsBtn2.addEventListener("click", function () {
      rsBtn2.disabled = true;
      apiJson("/api/health", "GET").then(function (h0) {
        var oldGen = (h0 && h0.gen) || "";
        return apiJson("/api/server-restart", "POST", { cause: "manual restart from the pane" }).then(function () {
          setStatus("Restarting the server\u2026", "busy");
          var tries = 0;
          var poll = function () {
            apiJson("/api/health", "GET").then(function (h1) {
              tries++;
              rsBtn2.disabled = false;
              if ((h1 && h1.gen && h1.gen !== oldGen) || tries > 60) {
                setStatus(tries > 60 ? "Server did not answer after the restart - start it manually." : "Server restarted (new generation).", tries > 60 ? "error" : "ok");
              } else { setTimeout(poll, 700); }
            }).catch(function () { if (++tries < 61) setTimeout(poll, 700); else { rsBtn2.disabled = false; setStatus("Server did not answer after the restart - start it manually.", "error"); } });
          };
          setTimeout(poll, 1200);
        });
      }).catch(function (e3) { rsBtn2.disabled = false; setStatus("Could not restart: " + (e3 && e3.message ? e3.message : e3), "error"); });
    });


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
    SK.activeFetch = (typeof AbortController !== "undefined") ? new AbortController() : null;
    if (SK.activeFetch) {
      opts.signal = SK.activeFetch.signal;
      SK.activeFetch.path = path;
    }
    try {
      fetchPromise = fetch(path, opts);
    } catch (syncErr) {
      logApi({ path: path, method: verb, status: 0, ok: false,
        err: "network (sync): " + String(syncErr && syncErr.message || syncErr),
        req: bodyStr.slice(0, 1600) });
      throw new FriendlyError("Cannot reach the OfficeKick server (" + path + ")");
    }
    var settleClear = function () { SK.activeFetch = null; };
    fetchPromise.then(settleClear, settleClear);
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
      if (SK.runCancelled || (fetchErr && fetchErr.name === "AbortError") ||
          /abort/i.test(String(fetchErr && fetchErr.name || "")))
        throw new FriendlyError("Run cancelled.");
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
    "- 'reply'/'answer'/'respond' on the open mail means draft_reply (add \"all\":true when the user says reply to all/everyone). 'forward'/'fwd' means the forward op. Meeting/calendar/invite requests mean new_meeting; resolve 'today'/'tomorrow'/'next week' from the TODAY line in the context (NEVER from the email's own date) and never schedule a meeting or reminder in the past.",
    "- The context may list ATTACHMENTS (names + sizes) of the open mail: reference them by name in drafts/explanations, but you CANNOT read their contents.",
    "- Reply/summary/translate/brainstorm requests that need no change: use format A.",
    "- Keep reply drafts short and human; greeting + body + signoff unless told otherwise.",
    "- LIVE WEB LOOKUPS: when the request needs up-to-date info you cannot reliably know (current events, news, prices, recent releases, rankings, live data), reply with ONLY {\"kind\":\"web_lookup\"} - nothing else. MailKick then runs a web-enabled lookup and re-asks you with fresh citable results. Never use it for mail edits, styling, general concepts, or what the email context already answers. Lookups are INTERNAL actions - free (never gated/approved) - and you may reply with later rounds of {\"kind\":\"web_lookup\"} as well (up to 3 lookups per run). After a lookup you may also ask for the FULL text of one page from its results by replying with ONLY {\"kind\":\"read_page\",\"url\":\"https://...\"} (internal, free, up to 4 pages per run). When a page's text carries a \"…[page continues]\" ending and the info you need may sit deeper, read the NEXT part by replying with ONLY {\"kind\":\"read_page\",\"url\":\"https://...\",\"from\":<n>} using the part start the pane gave you (continuations are free and do not use the 4-page budget). You have no web access yourself and cannot run a search; the lookup channel belongs to MailKick - never claim you tried a lookup or that the lookup channel is unavailable, and never refuse for lack of live data: reply with {\"kind\":\"web_lookup\"} instead.",
  ].join("\n");

  var CLARIFY_RULE =
    "\nAMBIGUOUS DELIVERABLE or SCOPE: when the request could plausibly produce a DIFFERENT deliverable and the email " +
    "context does not decide it (summarize the thread vs. draft replies; one reply vs. one per open thread; who " +
    "should be answered; which of the MULTI-SELECTED items when several are open), you may ask the user ONCE instead " +
    "of guessing. Reply exactly:" +
    " {\"kind\":\"clarify\",\"question\":\"<one short question>\",\"options\":[\"<short option>\",\"<short option>\"]} " +
    "(2-4 short options). This replaces the ops/reply round - it is NOT an action: the pane shows the question; " +
    "the user's choice arrives in your next ask (look for CLARIFICATION) - then produce the deliverable for that " +
    "choice without asking again. Only when the wording and the context do not already decide it; never for " +
    "tone/length details you can decide yourself, and never twice in one chain.";
  var NOTES_OFF_RULE = "";

  var REPAIR_ASK =
    "Your previous answer did not follow the required format. " +
    "Reply again with ONLY the JSON object (format A or B) and nothing else: " +
    'either {"kind":"reply","text":"..."} or {"kind":"slide_ops","ops":[...]} - ' +
    "no reasoning, no fences, no commentary.";

  var REPAIR_EMPTY_ASK =
    "Your previous reply was EMPTY. An empty response is NEVER acceptable. " +
    "Always produce something useful: the ops JSON with actions for the " +
    "user's request, or format A (reply) with the answer/reason. Respond NOW " +
    "with the JSON object - no fences, no commentary, never an empty reply.";

  var CONTINUE_ASK =
    "Your previous reply was cut off mid-way through the op list. Reply with a " +
    "slide_ops JSON containing ONLY the ops still needed to FINISH the redesign " +
    "(slides not yet covered), same JSON format. If nothing is missing, reply " +
    '{"kind":"reply","text":"done"}';
  function continueAsk(ctx) {
    return CONTINUE_ASK +
      "\nREMINDERS: reference only paragraphs/targets that appeared in the context summary; do not re-cover content already handled in your previous reply.";
  }

  var HISTORY_MAX_EXCHANGES = 16; // oldest turns are dropped, newest kept (0.9.210: 6 -> 16 - much longer working memory)
  var HISTORY_MAX_CHARS = 10000; // per stored message (long answers must survive resume/history)

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
      el.innerHTML = "";
      var lbl = document.createElement("span");
      lbl.textContent = "Session: " + (SK.sessionName || derivedSessionName());
      el.appendChild(lbl);
      el.appendChild(document.createTextNode(" "));
      var ref = document.createElement("button");
      ref.className = "btn subtle slabelbtn";
      ref.type = "button";
      ref.textContent = "\u21bb";
      ref.title = "Ask the AI for a new session title";
      ref.addEventListener("click", function (ev) {
        ev.stopPropagation();
        if (SK.titleBusy || !SK.sessionId) return;
        SK.sessionManual = false;   // an AI refresh re-enables auto updates
        aiNameSession(SK.sessionId, true);
      });
      var ed = document.createElement("button");
      ed.className = "btn subtle slabelbtn";
      ed.type = "button";
      ed.textContent = "\u270e";
      ed.title = "Type your own title (stops automatic AI renames until you press refresh)";
      ed.addEventListener("click", function (ev) {
        ev.stopPropagation();
        renameSessionLabelInline();
      });
      el.appendChild(ref);
      el.appendChild(ed);
      el.style.display = "";
    } else {
      el.innerHTML = "";
      el.style.display = "none";
    }
  }

  function renameSessionLabelInline() {
    var el = $("chat-session-label");
    if (!el || !SK.sessionId) return;
    if (el._editing) return;
    el._editing = true;
    var lbl = el.querySelector("span");
    var input = document.createElement("input");
    input.type = "text";
    input.maxLength = 48;
    input.value = SK.sessionName || derivedSessionName();
    input.className = "sess-rename";
    var done = function (commit) {
      el._editing = false;
      var val = input.value.trim();
      el.innerHTML = "";
      updateSessionLabel();
      if (commit && val && val !== SK.sessionName) {
        SK.sessionManual = true;   // user-titled: no more automatic AI renames
        apiJson("/api/sessions/rename/" + SK.sessionId, "POST", { name: val, manual: true })
          .then(function () { SK.sessionName = val; updateSessionLabel(); renderSessions(); setStatus("Session renamed - AI title updates pause until you press refresh.", "ok"); })
          .catch(function (e) { setStatus("Rename failed: " + ((e && e.message) || e), "error"); });
      }
    };
    input.addEventListener("keydown", function (e) {
      if (e.key === "Enter") done(true);
      else if (e.key === "Escape") done(false);
    });
    input.addEventListener("blur", function () { done(false); });
    lbl.textContent = "";
    lbl.appendChild(input);
    input.focus(); input.select();
  }


  function scheduleSaveSession() {
    if (SK.sessionSaveTimer) clearTimeout(SK.sessionSaveTimer);
    SK.sessionSaveTimer = setTimeout(saveSessionNow, 600);
  }

  function saveSessionNow() {
    if (SK.testMode) return;   // test runs never end up in history
    if (!SK.history.length) return;
    var payload = { exchanges: SK.history.slice(), app: SK.app, manualName: !!SK.sessionManual };
    var createPromise;
    if (!SK.sessionId) {
      payload.name = derivedSessionName();
      createPromise = apiJson("/api/sessions", "POST", payload).then(function (res) {
        if (res && res.session) {
          SK.sessionId = res.session.id;
          try { localStorage.setItem("sk-resume-session-mailkick", SK.sessionId); } catch (eS) {}
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
    var wasFirstSave = !SK.sessionIdBeforeName;
    return createPromise.then(function () {
      dbg("session.saved", { id: SK.sessionId, name: SK.sessionName, messages: SK.history.length });
      maybeUpdateSessionTitle();
    }).catch(function (err) {
      dbg("session.saveFailed", { msg: (err && err.message) || String(err) });
    });
  }

  function aiNameSession(sid, updateMode) {
    if (SK.testMode) return;   // no AI calls for test sessions
    // one tiny call through the profile's other slot; the derived name stays
    // if anything fails, and the server repairs/cleans what the model returns
    var prof = currentProfile();
    if (!prof || !prof.other || !prof.other.connection || !prof.other.model) return;
    if (SK.titleBusy) return;   // a title update is already running
    SK.titleBusy = true;
    var first = derivedSessionName();
    var renameFallback = function () {
      apiJson("/api/sessions/rename/" + sid, "POST", { name: first.slice(0, 48) });
    };
    var sysPrompt = "You generate very short chat-session titles. Reply with ONLY the title: 2-4 words, no quotes, no punctuation at the end.";
    var userMsg = "Session start: " + first.slice(0, 120);
    if (updateMode) {
      // the title must follow the session's CURRENT overall focus: consider
      // up to the last 6 requests together with the existing title (models
      // over-index on the newest ask otherwise), KEEP = title still fits
      var asks = [];
      SK.history.slice(-HISTORY_MAX_EXCHANGES * 2).forEach(function (h) {
        if (h.role === "user" && asks.length < 6) asks.push(String(h.content || "").replace(/\s+/g, " ").slice(0, 90));
      });
      userMsg = "Existing session title: " + String(SK.sessionName || first).slice(0, 60) +
        "\nLast " + asks.length + " request(s) (oldest first):";
      asks.forEach(function (a) { userMsg += "\n- " + a; });
      userMsg += "\nConsider ALL of those requests together (not only the newest)." +
        "\nReply with ONLY the 2-4 word title that fits the session's overall current focus, or reply with exactly KEEP if the existing title still covers it.";
      sysPrompt += " Titles must follow the session's overall focus across the listed requests, never just the last one. If the current title still fits, reply with exactly KEEP.";
    }
    apiJson("/api/chat", "POST", {
      connection: prof.other.connection, model: prof.other.model,
      temperature: 0.2, maxTokens: 300,
      messages: [
        { role: "system", content: sysPrompt },
        { role: "user", content: userMsg },
      ],
    }).then(function (resp) {
      var name = stripReasoning(resp.text || "").replace(/"/g, "").replace(/^[:\- ]+/, "").trim().slice(0, 48);
      // the model may answer KEEP (= the current title still fits): must not
      // rename the session to the literal word
      if (String(name).toUpperCase() === "KEEP") { SK.titleBusy = false; return; }
      // reject non-title junk (models that ignore the instruction and echo ops/json)
      var junk = /^[{\[]/.test(name) || /"(kind|ops|type)"/.test(name) || name.length < 2 || name.length > 60;
      if (junk) { renameFallback(); SK.titleBusy = false; return; }
      return apiJson("/api/sessions/rename/" + sid, "POST", { name: name }).then(function () {
        SK.sessionName = name;
        dbg("session.named", { id: sid, name: name });
        SK.titleBusy = false;
        updateSessionLabel();
        if (SK.activeTab === "sessions") renderSessions();
      });
    }).catch(function () {
      dbg("session.namingFailed", { id: sid });   // derived name stays
      SK.titleBusy = false;
    });
  }

  // as the session continues, keep the AI title in line with what the
  // session is now about (bounded: one update in flight, cheap calls)
  function maybeUpdateSessionTitle() {
    if (SK.testMode || !SK.sessionId) return;
    if (SK.sessionManual) return;   // user-titled: only the refresh button renames
    if (SK.titleBusy) return;
    if (!SK.history || SK.history.length < 4) return;   // nothing meaningful yet
    try {
      aiNameSession(SK.sessionId, true);
    } catch (eT) { /* naming stays */ }
  }

  function resumeSession(sessId) {
    apiJson("/api/sessions/" + sessId, "GET").then(function (doc) {
      if (SK.mailStaleResumeDropUntil && Date.now() < SK.mailStaleResumeDropUntil) {
        // 0.9.233: the resumed session belongs to a different email and the
        // new-session-per-email setting is on - drop it once, don't apply
        SK.mailStaleResumeDropUntil = 0;
        dbg("mail.resumeDropped", { id: sessId });
        return;
      }
      SK.history = (doc.exchanges || []).slice(-HISTORY_MAX_EXCHANGES * 2);
      SK.sessionApprovals = {};   // a remembered approval never travels between sessions
      try { localStorage.removeItem("sk-approvals-mailkick"); } catch (eRC) {}
      SK.sessionId = doc.id;
      try { localStorage.setItem("sk-resume-session-mailkick", doc.id); } catch (eS) {}
      SK.sessionName = doc.name || "";
      restoreSessionApprovals();   // approvals ride the resumed session (0.9.267)
      SK.sessionManual = !!doc.manualName;   // absent = AI-titled (updates allowed)
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
      Math.min(HISTORY_MAX_EXCHANGES, SK.history.length / 2) + " exchange(s). The last turns are in the Previous requests box (collapsed; click its header to open). Just type in the request box to continue.";
    el.appendChild(note);
    // show the last recent messages as collapsed request/response
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
    var wasCurrent = (SK.sessionId === sessId);
    apiJson("/api/sessions/" + sessId, "DELETE").then(function () {
      try { delete SK.sessionNames[sessId]; } catch (eN) {}
      if (wasCurrent) {
        // removing the ACTIVE session = wipe its slate + ready a new one
        SK.sessionId = null; SK.sessionName = ""; SK.sessionManual = false;
        SK.history = [];
        try { localStorage.removeItem("sk-resume-session-' + app + '"); } catch (eS) {}
        var inp0 = $("chat-input");
        if (inp0) { inp0.disabled = false; inp0.value = ""; }
        var el0 = $("chat-output");
        if (el0) el0.innerHTML = "";
        var body0 = $("past-body");
        if (body0) body0.innerHTML = "";
        var cslot = $("clarify-slot");
        if (cslot) { cslot.innerHTML = ""; cslot.style.display = "none"; }
        updatePastBox();
        updateSessionLabel();
        renderSessions();
        switchTab("chat");
        setStatus("Current session removed - a new session is ready.", "ok");
        return;
      }
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

  // 0.9.295: the full-response viewer - History keeps only a 240-char
  // bubble preview (lost formatting + cut-off); clicking an assistant
  // bubble opens the COMPLETE response, lightly markdown-formatted.
  function escapeHtmlTk(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function mdToSafeHtml(md) {
    var src = String(md || "");
    var blocks = [];
    src = src.replace(/```[\s\S]*?```/g, function (code) {
      blocks.push('<pre class="mdcode">' + escapeHtmlTk(code.replace(/^```\w*\n?/, "").replace(/```$/, "")) + "</pre>");
      return "\u0000B" + (blocks.length - 1) + "\u0000";
    });
    var inline = function (s) {
      return escapeHtmlTk(s)
        .replace(/`([^`]+)`/g, "<code>$1</code>")
        .replace(/\*\*([^*\n]+)\*\*/g, "<b>$1</b>")
        .replace(/(^|[\s(])\*([^*\n]+)\*/g, "$1<i>$2</i>");
    };
    var out = src.split(/\n{2,}/).map(function (block) {
      var t0 = block.trim();
      if (!t0) return "";
      var marker = t0.match(/^\u0000B(\d+)\u0000$/);
      if (marker) return blocks[parseInt(marker[1], 10)];
      var lines = t0.split("\n");
      var isList = lines.every(function (l) { return /^[-*]\s+/.test(l.trim()); });
      var isNum = lines.every(function (l) { return /^\d+[.)]\s+/.test(l.trim()); });
      if (isList) return "<ul>" + lines.map(function (l) { return "<li>" + inline(l.replace(/^[-*]\s+/, "")) + "</li>"; }).join("") + "</ul>";
      if (isNum) return "<ol>" + lines.map(function (l) { return "<li>" + inline(l.replace(/^\d+[.)]\s+/, "")) + "</li>"; }).join("") + "</ol>";
      var allHeads = lines.length && lines.every(function (l) { return /^#{1,4}\s+/.test(l || ""); });
      if (allHeads) {
        return lines.map(function (l) {
          var h2 = /^#{1,4}\s+/.exec(l);
          var lvl = Math.min(h2[0].trim().length, 4);
          return '<div class="mdh mdh' + lvl + '">' + inline(l.replace(/^#+\s+/, "")) + "</div>";
        }).join("");
      }
      var hm = /^#{1,4}\s+/.exec(lines[0]);
      if (hm && lines.length === 1) {
        var lvl = Math.min(hm[0].trim().length, 4);
        return '<div class="mdh mdh' + lvl + '">' + inline(lines[0].replace(/^#+\s+/, "")) + "</div>";
      }
      return "<p>" + lines.map(inline).join("<br>") + "</p>";
    }).join("");
    return out || escapeHtmlTk(src);
  }
  function openTextViewer(title, content) {
    var scrim = document.getElementById("resp-viewer");
    if (!scrim) {
      scrim = document.createElement("div");
      scrim.id = "resp-viewer";
      scrim.innerHTML = '<div class="rv-card"><div class="rv-head"><b id="rv-title"></b>' +
        '<button class="btn subtle" id="rv-close" type="button">\u2715 close</button></div>' +
        '<div id="rv-body" class="rv-body"></div></div>';
      document.body.appendChild(scrim);
      document.getElementById("rv-close").addEventListener("click", function () { scrim.classList.add("hidden"); });
      scrim.addEventListener("click", function (e) { if (e.target === scrim) scrim.classList.add("hidden"); });
      document.addEventListener("keydown", function (e) {
        if (e.key === "Escape" && !document.getElementById("resp-viewer").classList.contains("hidden")) {
          document.getElementById("resp-viewer").classList.add("hidden");
        }
      });
    }
    document.getElementById("rv-title").textContent = title || "Response";
    document.getElementById("rv-body").innerHTML = mdToSafeHtml(content);
    scrim.classList.remove("hidden");
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
        var hintFull = false;
        if (m.role !== "user") {
          // 0.9.295: a text response opens FULL + formatted on click
          var fullResp = String(m.content || "");
          b.style.cursor = "pointer";
          b.title = "Click to open the full response";
          b.addEventListener("click", function () { openTextViewer("AI response", fullResp); });
        }
        if (c.length > 240) {
          c = c.slice(0, 240) + "…";
          if (m.role !== "user") hintFull = true;
        }
        var txt = document.createElement("span");
        txt.textContent = c;
        b.appendChild(who); b.appendChild(txt);
        if (hintFull) {
          var hint = document.createElement("span");
          hint.className = "ptts";
          hint.textContent = "  (click for the full response)";
          b.appendChild(hint);
        }
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
  function setupMissingNote() {
    // the sessions tab is meaningless until a connection + profile exist;
    // return a guidance string when still unconfigured, null when usable
    var cfg = SK.cfg || {};
    var hasConn = (cfg.connections || []).length > 0;
    var hasProf = (cfg.profiles || []).length > 0;
    if (hasConn && hasProf) return null;
    if (!hasConn) return "No setup yet. Create a connection in Settings first (Connections tab), then a profile in Settings (Profiles tab). Sessions become available once a profile exists.";
    return "No profile yet. Create one in Settings (Profiles tab: add or pick models). Sessions become available once a profile exists.";
  }
  function renderSessions() {
    var list0 = $("sess-list");
    if (!list0) return;
    var missing = setupMissingNote();
    if (missing) {
      // unconfigured: nothing interactive - guidance instead of the list
      list0.innerHTML = "";
      var note2 = document.createElement("div");
      note2.className = "opnote";
      note2.textContent = missing;
      list0.appendChild(note2);
      var head2 = $("sess-head");
      var tail = "";
      if (typeof updatePastBox === "function") { /* keep counters untouched */ }
      if (SK.cfg && SK.cfg._sessCounts) tail = "";
      list0.className = "sesslist disabled";
      return;
    }
    list0.className = "sesslist";
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
              // stage 1 confirm: the next click deletes - but the CURRENT
              // session demands an explicit second confirmation (0.9.183)
              b.dataset.armed = "1";
              b.textContent = (SK.sessionId === row.id)
                ? "Delete CURRENT session?" : "Really delete?";
              b.classList.add("armed");
              setTimeout(function () {
                if (b.parentNode) { b.dataset.armed = "0"; b.dataset.armed2 = "0"; b.textContent = "Delete"; b.classList.remove("armed"); }
              }, 3500);
              return;
            }
            if (pair[1] && SK.sessionId === row.id && b.dataset.armed2 !== "1") {
              // stage 2: the wording makes clear this is the ACTIVE session
              b.dataset.armed2 = "1";
              b.textContent = "Really remove the CURRENT session?";
              b.classList.add("armed");
              setStatus("You are about to remove the CURRENT session - click once more to confirm (a fresh session starts).", "busy");
              setTimeout(function () {
                if (b.parentNode) { b.dataset.armed = "0"; b.dataset.armed2 = "0"; b.textContent = "Delete"; b.classList.remove("armed"); }
              }, 6000);
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

  // shared "new session" flow (chat-tab New button + History-tab New session):
  // armed confirm, keep the session in History, then a clean slate everywhere
  function startNewSession(btn) {
    if (!btn) return;
    if (btn.dataset.armed !== "1") {
      btn.dataset.armed = "1";
      btn.dataset.was = btn.textContent;
      btn.textContent = "Really start a new session?";
      btn.classList.add("armed");
      setTimeout(function () {
        if (btn.parentNode) { btn.dataset.armed = "0"; btn.textContent = btn.dataset.was || btn.textContent; btn.classList.remove("armed"); }
      }, 3500);
      return;
    }
    btn.dataset.armed = "0";
    btn.textContent = btn.dataset.was || btn.textContent;
    btn.classList.remove("armed");
    filePreviousRun();   // the current chat's last exchange = in the session already
    SK.history = [];
    SK.sessionId = null;
    SK.sessionName = "";
    SK.sessionManual = false;
    SK.sessionApprovals = {};   // a new session asks again (session-scoped approvals)
    try { localStorage.removeItem("sk-approvals-mailkick"); } catch (eRC) {}
    try { localStorage.removeItem("sk-res...kick"); } catch (eR) {}
    var inp0 = $("chat-input");
    if (inp0) { inp0.disabled = false; inp0.value = ""; }
    var el = $("chat-output");
    if (el) el.innerHTML = "";
    var body0 = $("past-body");
    if (body0) body0.innerHTML = "";   // a new session = a clean Previous requests box
    var cslot = $("clarify-slot");
    if (cslot) { cslot.innerHTML = ""; cslot.style.display = "none"; }
    updatePastBox();
    switchTab("chat");
    setStatus("New session - the current one stays saved in History.", "ok");
    renderSessions();
  }

  function wireSessions() {
var nw = $("sess-new");
    if (nw) nw.addEventListener("click", function () {
      // two-stage confirm (same armed pattern as the session Delete):
      // a fresh session is only meaningful when the current one is saved
      if (nw.dataset.armed !== "1") {
        nw.dataset.armed = "1";
        nw.dataset.was = nw.textContent;
        nw.textContent = "Really start a new session?";
        nw.classList.add("armed");
        setTimeout(function () {
          if (nw.parentNode) { nw.dataset.armed = "0"; nw.textContent = nw.dataset.was || "New session"; nw.classList.remove("armed"); }
        }, 3500);
        return;
      }
      nw.dataset.armed = "0";
      nw.textContent = nw.dataset.was || "New session";
      nw.classList.remove("armed");
      // the current session's last exchange moves into the Previous requests
      // box first (it may not be saved yet if the user never ran a follow-up)
      filePreviousRun();
      // a new session begins on a clean slate
      var body0 = $("past-body");
      if (body0) body0.innerHTML = "";
      var cslot = $("clarify-slot");
      if (cslot) { cslot.innerHTML = ""; cslot.style.display = "none"; }
      updatePastBox();
      SK.history = [];
      SK.sessionId = null;
      SK.sessionName = "";
      SK.sessionManual = false;
      var inp0 = $("chat-input");
      if (inp0) { inp0.disabled = false; inp0.value = ""; }
      switchTab("chat");
      setStatus("New session - the current one stays saved in History.", "ok");
      var el = $("chat-output");
      if (el) el.innerHTML = "";
      try { localStorage.removeItem("sk-res...kick"); } catch (eR) {}
      renderSessions();
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
    if (parsed && parsed.kind === "clarify") return "(asked the user to pick: " + String(parsed.question || "").slice(0, 140) + ")";
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
    // always copy to the clipboard; refill the input only when it is empty
    var txt = String(fullText || "");
    var copied = false;
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(txt).catch(function () {});
        copied = true;
      }
    } catch (eC) { /* fall through to the execCommand fallback */ }
    if (!copied) {
      try {
        var ta = document.createElement("textarea");
        ta.value = txt;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
      } catch (eF) { /* clipboard unavailable - ignore */ }
    }
    var inp = $("chat-input");
    if (!inp || inp.value) return;
    inp.value = txt;
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
    SK.sessionApprovals = {};   // approvals are session-scoped - a new session asks again
    try { localStorage.removeItem("sk-approvals-mailkick"); } catch (eRC) {}
    SK.sessionId = null;
    try { localStorage.removeItem("sk-resume-session-mailkick"); } catch (eS) {}      // the next exchange saves as a NEW session
    SK.sessionName = "";
    SK.sessionManual = false;
    updateChatMeta();
    updateSessionLabel();
    renderSessions();
    dbg("history.reset", {});
  }

  // ---------------- new session per email (0.9.233, MailKick) -------- //
  // Setting "create a new session when opening a new email" (off by
  // default): the identity of the open email (Outlook itemId; the sim maps
  // it to its own ids) is remembered whenever the pane sees it. When it
  // changes, the chat drops into a clean new session. Sessions with no
  // requests are never persisted (saveSessionNow returns on empty history),
  // and the old conversation stays saved in History.
  function mailIdentityFromItem(item) {
    try {
      if (!item) return "";
      var isCompose = (item.getAsync !== undefined && item.subject && item.subject.getAsync);
      if (isCompose) return "";   // a compose draft has no stable itemId
      return String(item.itemId || "");
    } catch (eId) { return ""; }
  }
  function mailCurrentIdentity() {
    if (SK.sim) {
      try {
        return (window.SimAPI && window.SimAPI.currentEmailId) ? String(window.SimAPI.currentEmailId()) : "";
      } catch (eS) { return ""; }
    }
    var item = null;
    try { item = Office.context.mailbox.item; } catch (eO) { return ""; }
    return mailIdentityFromItem(item);
  }
  function mailIdentityRemember(id) {
    try { if (id) localStorage.setItem("mailkick-lastmail", String(id).slice(0, 400)); } catch (eL) {}
  }
  function mailIdentityLast() {
    try { return localStorage.getItem("mailkick-lastmail") || ""; } catch (eL) { return ""; }
  }
  function mailBlankChat() {
    var zone0 = $("chat-output");
    if (zone0) zone0.innerHTML = "";   // a stale resume notice / last-run render leaves the top zone
    var body0 = $("past-body");
    if (body0) body0.innerHTML = "";   // a new session = a clean Previous requests box
    var cslot = $("clarify-slot");
    if (cslot) { cslot.innerHTML = ""; cslot.style.display = "none"; }
    updatePastBox();
  }
  function mailSessionForIdentity(id, source) {
    id = String(id || "").slice(0, 400);
    if (!(SK.cfg && SK.cfg.newSessionPerEmail)) return false;
    if (!id) return false;
    if (SK.mailEmailId === null) {   // first sighting: baseline only, no reset
      SK.mailEmailId = id;
      mailIdentityRemember(id);
      return false;
    }
    if (id === SK.mailEmailId) return false;
    resetHistory();   // clears history + the boot-resume pointer
    mailBlankChat();
    SK.mailEmailId = id;
    mailIdentityRemember(id);
    setStatus("New email opened - started a new session.", "ok");
    actLog("New session for this email (" + source + ").");
    dbg("mail.newSessionPerEmail", { source: source });
    return true;
  }
  function applyMailSessionBootPolicy() {
    // on every config load: set the baseline and drop a resumed session
    // when the open email is not the one the last conversation belongs to
    if (!(SK.cfg && SK.cfg.newSessionPerEmail)) return;
    var cur = mailCurrentIdentity();
    if (!cur) return;
    var last = mailIdentityLast();
    if (SK.mailEmailId === null) {
      // first policy pass after boot: when the pane remembers a DIFFERENT
      // email, any in-flight boot resume is a different email's session
      // - block it for a bounded window and clean the slate
      if (last && last !== cur) {
        SK.mailStaleResumeDropUntil = Date.now() + 10000;
        resetHistory();
        mailBlankChat();
        setStatus("New session for this email (the previous conversation belongs to a different email).", "ok");
        dbg("mail.staleResumeDropped", {});
      }
      SK.mailEmailId = cur;
      mailIdentityRemember(cur);
      return;
    }
    if (last && last !== cur) {
      // a boot resume of another email's session may still be in flight -
      // block it for a bounded window, then clean the slate
      SK.mailStaleResumeDropUntil = Date.now() + 10000;
      resetHistory();
      mailBlankChat();
      setStatus("New session for this email (the previous conversation belongs to a different email).", "ok");
      dbg("mail.staleResumeDropped", {});
    }
    SK.mailEmailId = cur;
    mailIdentityRemember(cur);
  }
  var mailSelTimer = null;
  function onMailSelectionChanged() {
    if (SK.busy) return;   // never reset mid-run: the next run's check catches it
    if (mailSelTimer) clearTimeout(mailSelTimer);
    mailSelTimer = setTimeout(function () {
      mailSelTimer = null;
      if (SK.busy) return;
      mailSessionForIdentity(mailCurrentIdentity(), "selection");
    }, 350);
  }
  function watchMailSelection() {
    try {
      if (SK.sim || !SK.deck) return;
      if (!Office.context.requirements.isSetSupported("Mailbox", "1.13")) return;
      if (!Office.context.mailbox || !Office.context.mailbox.addHandlerAsync) return;
      Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, onMailSelectionChanged,
        function (ar) { dbg("mail.selection.watch", { ok: !!(ar && ar.status === "succeeded") }); });
    } catch (eW) { dbg("mail.selection.watchFailed", { msg: String((eW && eW.message) || eW) }); }
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
        try { out.itemId = mailIdentityFromItem(item); } catch (eI0) { out.itemId = ""; }
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
  function restoreSessionApprovals() {
    // approvals ride the session: on a pane reload the session resumes and
    // the remembered approvals come back with it (same session id only)
    try {
      var saved = JSON.parse(localStorage.getItem("sk-approvals-mailkick") || "null");
      if (saved && saved.map && (saved.sid || "") === (SK.sessionId || "")) {
        SK.sessionApprovals = saved.map;
        return true;
      }
    } catch (eR2) { /* corrupt entry - start clean */ }
    return false;
  }

  function localDateStr(d) {
    var z = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "-" + z(d.getMonth() + 1) + "-" + z(d.getDate());
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
        // DATE GUARD (0.9.246): anchor "today" to the REAL current date, not
        // to the email's date or whatever the model remembers. A past date is
        // always wrong for a reminder/meeting; an explicit today/tonight ask
        // must land on today. Self-diagnosing reasons feed the repair rounds.
        var todayStr = localDateStr(new Date());
        var startDay = String(op.start).slice(0, 10);
        if (startDay < todayStr) {
          dbg("date.mismatch", { op: t, found: startDay, today: todayStr, past: true });
          bad("WRONG DATE: the meeting is scheduled for " + startDay +
              ", which is in the PAST. Today is " + todayStr +
              " - re-emit the op with a correct future date (resolve 'today' from" +
              " the TODAY line in the context).");
          return;
        }
        if (/\b(today|tonight|this evening)\b/i.test(SK.lastInstruction || "") && startDay !== todayStr) {
          dbg("date.mismatch", { op: t, found: startDay, today: todayStr, askedToday: true });
          bad("WRONG DATE: the user asked for TODAY (" + todayStr +
              "), but the op is dated " + startDay +
              " - re-emit the op dated " + todayStr + ".");
          return;
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
    var failedOps = [];   // ops whose runOne link rejected - retried once after the pass
    var chain = Promise.resolve();
    var progress = function (op, l0, ok) {
      if (typeof SK.opProgress === "function") SK.opProgress(op, lines.slice(l0), ok);
    };
    ops.forEach(function (op) {
      chain = chain.then(function () {
        if (SK.runCancelled) {
          lines.push({ ok: false, text: "cancel - remaining op(s) not executed (" + (op && op.type ? op.type : "unknown") + " skipped)" });
          return Promise.resolve();
        }
        var l0 = lines.length;
        return runOne(op).then(function () {
          progress(op, l0, true);
        }, function (err) {
          lines.push({ ok: false, text: fmtOp(op, ctx) + " - " + ((err && err.message) || String(err)) });
          failedOps.push(op);
          progress(op, l0, false);
        });
      }, function (err) {
        lines.push({ ok: false, text: fmtOp(op, ctx) + " - " + ((err && err.message) || String(err)) });
        failedOps.push(op);
      });
    });
    return chain.then(function () {
      if (!failedOps.length) return { lines: lines, manualNotes: manualNotes };
      lines.push({ ok: true, text: "retrying " + failedOps.length +
        " failed action(s) once after a short pause\u2026" });
      var rseq = new Promise(function (res) { setTimeout(res, 1500); });
      failedOps.forEach(function (op) {
        rseq = rseq.then(function () {
          var l0 = lines.length;
          return runOne(op).then(function () { progress(op, l0, true); }, function (err2) {
            lines.push({ ok: false, text: "retry of " + op.type + " failed again: " +
              ((err2 && err2.message) || String(err2)) });
            progress(op, l0, false);
          });
        });
      });
      return rseq.then(function () { return { lines: lines, manualNotes: manualNotes }; });
    });
  }

  // ------------------------------------------------------------------ //
  // Chat / prompt flow
  // ------------------------------------------------------------------ //
  function formatContext(ctx) {
    var lines = [];
    lines.push("EMAIL CONTEXT (host: " + ctx.host + ", mode: " + ctx.mode + "):");
    {
      var dNow = new Date();
      var zN = function (n) { return (n < 10 ? "0" : "") + n; };
      var wdNow = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday",
                   "Friday", "Saturday"][dNow.getDay()];
      lines.push("TODAY: " + wdNow + " " + dNow.getFullYear() + "-" +
                 zN(dNow.getMonth() + 1) + "-" + zN(dNow.getDate()) +
                 " (relative words in the user's request - today/tomorrow/next week -" +
                 " refer to THIS date, never to the email's own date)");
    }
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

  // ---- cancel the current run: flips the flag (checked between steps) and
  // aborts the in-flight request for an immediate interrupt ----
  function cancelRun() {
    if (!SK.busy) return;
    SK.runCancelled = true;
    SK.cancelErase = true;   // a canceled request is wiped from the record (0.9.302)
    if (SK.activeFetch && SK.activeFetch.abort) {
      try { SK.activeFetch.abort(); } catch (eA) { /* already settled */ }
    }
    setStatus("Cancelling\u2026", "busy");
    dbg("run.cancel", {});
  }

  function setRunEnabled(on) {
    $("chat-run").disabled = !on || SK.busy;
    var cb = $("chat-cancel");
    if (cb) {
      cb.disabled = !SK.busy;                          // only meaningful mid-run
      cb.style.display = SK.busy ? "" : "none";        // appears only while a run is in progress
    }
  }

  function runQuickReply(stance, profile) {
    // reply + / reply - / propose quick actions: a canned instruction runs
    // through the normal pipeline (the open email's conversation is attached
    // as context); the model proposes draft_reply action(s) with the
    // AI-written text and the usual approval gate applies (MailKick never
    // sends by itself)
    if (SK.busy) return;
    var prof = currentProfile();
    if (!prof) { setStatus("No profile yet - create one on the Profiles tab first.", "error"); switchTab("settings"); switchSettingsSub("profiles"); return; }
    var inp = $("chat-input");
    if (!inp) return;
    var instruction;
    if (stance === "propose") {
      // 0.9.229: analyze the conversation and PROPOSE several plausible
      // replies - not one canned stance. Each proposal = a draft_reply op,
      // so they land in the approval list side by side.
      instruction =
        "Analyze this email conversation and decide yourself how many and which " +
        "reply proposals would genuinely help - this includes the judgment of " +
        "how many variants are useful (if only one natural reply exists, propose " +
        "only that one). Consider at least: a QUESTION, a positive/agreeable " +
        "answer, a politely negative answer, a concrete DECISION or commitment, " +
        "or a sensible alternative you judge better. Return one draft_reply action " +
        "PER proposal, each starting with a short marker like " +
        "[Proposal 1 - <kind you picked>] at the top of its body (simple inline " +
        "HTML only). Reply entirely in the conversation's language and tone. " +
        "Do not send anything.";
      if (profile) {
        // 0.9.305: the user picked a proposal profile - that kind wins
        instruction += " The user picked the PROPOSAL PROFILE \"" + String(profile) +
          "\": every proposal must fit exactly that kind (stance and content " +
          "tailored to it); drop the other kinds and the free choice above.";
      }
    } else {
      var stanceText = stance === "pos"
        ? "positive and agreeable"
        : "politely negative (declining) but still courteous";
      instruction = "Draft a polite reply to this email conversation in the tone the conversation uses (formal or informal). Stance: " + stanceText +
        ". Reply entirely in the same language as the conversation, in the same tone (formal or informal). Then propose one draft_reply action with the AI-generated reply text (simple inline HTML only). Do not send anything.";
    }
    inp.disabled = false;
    inp.value = instruction;
    runPrompt();
  }

  // 0.9.305: the propose quick-action asks WHICH proposal profile first
  // (same visual language as a model clarify, in the clarify slot).
  var PROPOSE_PROFILES = [
    "Question", "Reminder", "Follow-up", "Accept (positive)", "Decline (negative)",
    "Set up meeting", "Decision / commitment", "Other"
  ];
  function renderProposePicker() {
    if (SK.busy) return;
    var prof = currentProfile();
    if (!prof) { setStatus("No profile yet - create one on the Profiles tab first.", "error"); switchTab("settings"); switchSettingsSub("profiles"); return; }
    var slot = $("clarify-slot");
    if (!slot) { runQuickReply("propose"); return; }   // no slot host: free propose
    slot.innerHTML = "";
    slot.style.display = "";
    var box = document.createElement("div");
    box.className = "clarifybox";
    var q = document.createElement("div");
    q.className = "clarify-q";
    q.textContent = "What kind of reply proposal?";
    box.appendChild(q);
    var close = function () { slot.innerHTML = ""; slot.style.display = "none"; };
    var pick = function (val) { close(); runQuickReply("propose", val); };
    var row = document.createElement("div");
    row.className = "clarify-row";
    var ti = document.createElement("input");
    ti.type = "text";
    ti.className = "clarify-input";
    ti.placeholder = "other - type your own keyword\u2026";
    var go = document.createElement("button");
    go.className = "btn";
    go.type = "button";
    go.textContent = "Propose";
    go.disabled = true;
    ti.addEventListener("input", function () { go.disabled = !ti.value.trim(); });
    ti.addEventListener("keydown", function (eK) {
      if (eK.key === "Enter") { eK.preventDefault(); if (ti.value.trim()) pick(ti.value.trim()); }
    });
    go.addEventListener("click", function () { if (ti.value.trim()) pick(ti.value.trim()); });
    PROPOSE_PROFILES.forEach(function (p) {
      var b = document.createElement("button");
      b.className = "btn clarify-opt";
      b.type = "button";
      b.textContent = p;
      b.addEventListener("click", function () {
        if (p === "Other") { ti.focus(); return; }   // Other = the free-text row
        pick(p);
      });
      box.appendChild(b);
    });
    row.appendChild(ti); row.appendChild(go);
    box.appendChild(row);
    slot.appendChild(box);
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
    SK.runCancelled = false;
    SK.cancelErase = false;  // fresh run: cancellable until it renders (renderParsed re-clears)
    input.disabled = true;   // the text stays visible but locked while working
    setRunEnabled(false);
    chatTimerStart();
    actReset();
    actLog("Sending request to " + ref.connection + (ref.model ? " (" + ref.model + ")" : ""));
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
      // 0.9.233: when the selection event was missed (hosts without it, or
      // the sim), the run start detects the different email and resets the
      // session BEFORE the history is consumed for this ask
      mailSessionForIdentity((ctx && ctx.itemId) || mailCurrentIdentity(), "run");
      var sys = SYSTEM_PROMPT + "\nHost: " + (SK.sim ? "MailKick simulator" : SK.deck.capabilities().host);
      var _cp = SK.cfg ? (String(SK.cfg.customPrompt || "").trim() ? String(SK.cfg.customPrompt).trim() : String(SK.cfg.customPromptDefault || "").trim()) : "";
      if (_cp) {
        sys += "\n\nUSER CUSTOM INSTRUCTIONS (from Settings - obey these):\n" + _cp;
      }
      SK.webRounds = 0;
      SK.pageReads = 0;
      SK.webResults = "";
      if (!(SK.clarifyAnswer && SK.clarifyAnswer.expect === instruction)) SK.clarifyChain = 0;   // a free new ask resets the clarify chain
      var webPromise = Promise.resolve("");
      return webPromise.then(function () {
      if (!(SK.cfg && SK.cfg.generateNotes)) sys += "\n" + NOTES_OFF_RULE;
      sys += "\n" + CLARIFY_RULE;
      var ctxText = ctx ? formatContext(ctx) : "(no email could be read - make sure an email is open or selected and the reading pane stays open)";
      var clarifyBlock = "";
      if (SK.clarifyAnswer && SK.clarifyAnswer.expect === instruction) {
        clarifyBlock = "\n\nCLARIFICATION: you asked the user: \"" + SK.clarifyAnswer.q + "\". The user chose: " + SK.clarifyAnswer.a +
          ". Produce the deliverable for that choice now - do not ask again.";
        SK.clarifyAnswer = null;
        dbg("clarify.answer", { delivered: true });
      }
      return getTemplateBlock(instruction).then(function (tplBlock) {
        var user = ctxText + "\n\n" + (tplBlock ? tplBlock + "\n" : "") + "\nUSER REQUEST:\n" + instruction + clarifyBlock;
        if (SK.webResults) user += "\n\nWEB RESULTS (live lookup, cite them when relevant):\n" + SK.webResults;
        var messages = [
          { role: "system", content: sys },
        ].concat(SK.history.slice(-(HISTORY_MAX_EXCHANGES * 2)))
         .concat([{ role: "user", content: user }]);
        SK.lastAsk = { messages: messages, connection: ref.connection, model: ref.model, rounds: 0 };   // 0.9.217: repair rounds on validation rejects
        dbg("prompt", { profile: prof.name, route: route, model: ref.model, connection: ref.connection, instruction: instruction, hadContext: !!ctx, templateUsed: !!tplBlock, history: SK.history.length, messages: messages });
        setStatus("Asking " + ref.model + " (via " + ref.connection + ")…", "busy");
        var askedLookup = false;
      function mainAsk() {
        return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: messages }); })
          .then(function (resp) {
            var pre = tryParse(resp.text);
            var txt0 = String(resp.text || "");
            var noWebSaid = !pre.ok &&
                /lookup channel|don'?t have (web|live)|no (web|live|internet) (access|connection)|can'?t (access|browse|search|pull|verify|run|do) [a-z ]{0,24}(web|internet|search|lookup|online|current|fresh|live)|won'?t guess|cannot verify|can'?t verify|without (a )?(live |real )?(web|internet|search|lookup|live access)|stale\b/i.test(txt0);
            if (!askedLookup && !SK.noLookupRedo &&
                (noWebSaid || (pre.ok && pre.kind === "reply" &&
                  /can'?t verify|cannot verify|can'?t confirm|cannot confirm|no (live |web |internet )?(access|search|lookup|connection|data)|without (a )?(live )?(web|internet|search|live access)|would need (a )?(live|current)|requires? (a )?(live|web) (lookup|search|check|connection)|check a live|won'?t guess|pull live|didn'?t return a search result|did not return a search result/i.test(txt0))) &&
                /(rankings?|current|latest|today|news|price|prices|release|releases|live|up[- ]?to[- ]?date|fresh|who is|what happened|version|release|check|search|web|online)/i.test(instruction || "")) {
              // The model believes IT must search and refuses. The pane owns
              // the lookup channel - run the web lookup HERE and re-ask once
              // with the results.
              askedLookup = true;
              setStatus("Model said it can't search - doing the pane-side web lookup\u2026", "busy");
              return webLookup(prof, instruction).then(function (ws) {
                if (!ws) return resp;   // nothing to add - show the original answer
                SK.webResults = ws;
                setStatus("Reading the found documents\u2026", "busy");
                return gatherPageContents(ws).then(function (pages) {
                  SK.webResults = ws + pages;
                  messages = messages.concat([{ role: "assistant", content: txt0.slice(0, 4000) },
                                              { role: "user", content: "WEB LOOKUP RESULTS:\n" + ws +
                                                (pages ? "\n\nThe lookups also RETURNED THE FULL TEXT of these documents/pages - read them before answering:" + pages : "") +
                                               "\n\nAnswer the ORIGINAL USER REQUEST now using these results (cite sources); the required JSON format rules apply again." }]);
                  setStatus("Answering with the fresh web results\u2026", "busy");
                  return mainAsk2();
                });
              }).catch(function (e) { return resp; });
            }
            if (pre.ok && pre.kind === "web_lookup" && !askedLookup) {
              askedLookup = true;
              return webLookupThen(resp);
            }
            if (pre.ok && pre.kind === "read_page") return readPageThen(resp);
            return resp;
          });
      }
      // web_lookup: the model asked for a live web lookup. Internal action -
      // free (never gated, no approval) and repeatable (up to 3 rounds per
      // run): each round re-asks the model with the results, and the model
      // decides what to do next (ops, a reply, or another lookup).
      function webLookupThen(resp) {
        var pre = tryParse(resp.text);
        if ((SK.webRounds || 0) >= 3) {
          return { ok: true, kind: "reply",
                   text: "Web lookups were already run the maximum of 3 times this run. Decide with the results you have, or state what you could not verify." };
        }
        SK.webRounds = (SK.webRounds || 0) + 1;
        dbg("web_lookup", { round: SK.webRounds });
        setStatus("Doing a live web lookup (round " + SK.webRounds + ")\u2026", "busy");
        return webLookup(prof, instruction).then(function (ws) {
          if (!ws) {
            var ladderInfo = buildSearchLadder(prof);
            var ran = (ladderInfo.attempts || []).map(function (a) { return a.label || ((a.connection || "") + ":" + (a.model || "")); }).join(" -> ");
            var skippedTxt = (ladderInfo.skipped || []).join(", ");
            return { ok: true, kind: "reply",
                     text: "No live results came back from any lookup step in the search profile (" + (ran || "no steps configured") +
                           (skippedTxt ? "); skipped as unconfigured: " + skippedTxt : "") +
                           ". The model was not re-asked with web content." };
          }
          SK.webResults = ws;
          if (pre.ops && pre.ops.length) SK.pendingWebOps = pre.ops;   // look-up-then-act in one reply
          setStatus("Reading the found documents\u2026", "busy");
          return gatherPageContents(ws).then(function (pages) {
            SK.webResults = ws + pages;
            messages = messages.concat([{ role: "assistant", content: pre.text || '{"kind":"web_lookup"}' },
                                        { role: "user", content: "WEB LOOKUP RESULTS:\n" + ws +
                                          (pages ? "\n\nThe lookups also RETURNED THE FULL TEXT of these documents/pages - read them before answering:" + pages : "") +
                                          "\n\nAnswer the ORIGINAL USER REQUEST now using these results (cite sources); the required JSON format rules apply again." }]);
            setStatus("Answering with the fresh web results\u2026", "busy");
            return mainAsk2();
          });
        }).catch(function (e) {
          return { ok: true, kind: "reply", text: "The live web lookup failed: " + ((e && e.message) || e) };
        });
      }
      // read_page: the model wants the FULL text of one page (usually from a
      // lookup result). Internal + free; up to 4 reads per run.
      function readPageThen(resp) {
        var pre = tryParse(resp.text);
        var url = String(pre.url || "");
        var frm = (typeof pre.from === "number" && pre.from > 0) ? Math.floor(pre.from) : 0;
        if ((SK.pageReads || 0) >= 4) {
          return { ok: true, kind: "reply",
                   text: "Page reads were already run the maximum of 4 times this run. Decide with the text you have, or state what you could not verify." };
        }
        if (frm > 0 && (SK.pageConts || 0) >= 6) {
          return { ok: true, kind: "reply",
                   text: "This run already extended page text " + (SK.pageConts || 0) + " times. Decide with the text you have, or state what you could not verify." };
        }
        if (frm > 0) { SK.pageConts = (SK.pageConts || 0) + 1; }
        else { SK.pageReads = (SK.pageReads || 0) + 1; }
        dbg("read_page", { round: frm > 0 ? "cont-" + (SK.pageConts || 0) : SK.pageReads, from: frm, url: url.slice(0, 160) });
        setStatus(frm > 0 ? "Reading more of the requested page (from " + frm + ")\u2026"
                          : "Reading the requested page (" + SK.pageReads + ")\u2026", "busy");
        return fetchPageText(url, frm).then(function (pg) {
          if (!pg || !pg.text) {
            return { ok: true, kind: "reply",
                     text: "Could not read " + url + (frm ? " at offset " + frm : "") + " from the server's fetcher. Pick a different page from the results, or decide with what you have." };
          }
          if (pre.ops && pre.ops.length) SK.pendingWebOps = pre.ops;   // page-read-then-act in one reply
          var trailer = pg.more ?
            "\n\n(The text shown above is PART of the page - it ends at char " + ((pg.offset || 0) + pg.chars) + " of " + pg.total +
            ". If what you need may sit deeper, read the NEXT part now with ONLY " +
            '{"kind":"read_page","url":"' + url + '","from":' + pg.next + "} (continuations do not use the page-read budget).)" : "";
          messages = messages.concat([{ role: "assistant", content: pre.text || JSON.stringify({ kind: "read_page", url: url, from: frm }) },
                                      { role: "user", content: "PAGE CONTENT (" + (pg.title || pg.url || url) + ")" + (frm ? " [part from char " + frm + "]" : "") + " - " + url + ":\n" + pg.text + trailer +
                                        "\n\nUse it for the ORIGINAL USER REQUEST; cite the page. The required JSON format rules apply again." }]);
          setStatus("Answering with the requested page\u2026", "busy");
          return mainAsk2();
        }).catch(function (e) {
          return { ok: true, kind: "reply", text: "Page read failed: " + ((e && e.message) || e) };
        });
      }
      function mainAsk2() {
        return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: messages }); })
          .then(function (resp) {
            var pre = tryParse(resp.text);
            if (pre.ok && pre.kind === "web_lookup") return webLookupThen(resp);
            if (pre.ok && pre.kind === "read_page") return readPageThen(resp);
            return resp;
          });
      }
      return mainAsk().then(function (resp) {
          dbg("response", { text: resp.text, elapsedMs: resp.elapsedMs });
          SK.lastModelText = { text: resp.text };
          var parsed = tryParse(resp.text);
          if (parsed.ok && parsed.kind === "web_lookup" && parsed.ops) {
              SK.pendingWebOps = parsed.ops;
              return webLookupThen(resp);
          }
          if (parsed.ok && parsed.kind === "web_lookup") {
              if (!SK.webResults) return webLookupThen(resp);
              return parsed;
          }
          if (parsed.ok && parsed.kind === "read_page" && parsed.ops) {
              SK.pendingWebOps = parsed.ops;
              return readPageThen(resp);
          }
          if (parsed.ok && parsed.kind === "read_page") {
              if (!SK.webResults) return readPageThen(resp);
              return parsed;
          }
          if (parsed.ok && parsed.recovered) {
              // truncated reply: complete ops salvaged - ask once for the rest, merge
              SK.lastRecovered = parsed.ops.length;
              setStatus("Reply was cut off - recovered " + parsed.ops.length + " action(s), asking for the rest…", "busy");
              var cont = messages.concat([
                { role: "assistant", content: resp.text || "" },
                { role: "user", content: continueAsk(ctx) },
              ]);
              return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: cont }); })
                .then(function (r3) {
                  dbg("continue.response", { text: r3.text, elapsedMs: r3.elapsedMs });
                  SK.lastModelText = { text: r3.text };
                  var p3 = tryParse(r3.text);
                  var extra = (p3.ok && !p3.recovered && p3.ops ? p3.ops : salvageOps(r3.text)) || [];
                  return { ok: true, kind: "slide_ops", ops: parsed.ops.concat(extra), recovered: true };
                });
            }
            if (parsed.ok) return parsed;   // clean parse: never repair (the missing escape re-asked every clean mail run)
          var repairRound = function (nudge, prevText, level) {
            var repairMsgs = messages.concat([
              { role: "assistant", content: prevText || "" },
              { role: "user", content: nudge },
            ]);
            dbg("repair.request", { level: level, empty: !String(prevText || "").trim() });
            setStatus("Model output unclear - silently re-asking for clean JSON…", "busy");
            return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: repairMsgs }); })
              .then(function (r2) {
                dbg("repair.response", { level: level, text: r2.text, elapsedMs: r2.elapsedMs });
                SK.lastModelText = { text: r2.text };
                return { text: r2.text, parsed: tryParse(r2.text) };
              });
          };
          return repairRound(REPAIR_ASK, resp.text, 1).then(function (r1) {
            if (r1.parsed.ok) return r1.parsed;
            var wasEmpty = !String(r1.text || "").trim();
            var prevEmpty = !String(resp.text || "").trim();
            var nudge2 = (wasEmpty || prevEmpty) ? REPAIR_EMPTY_ASK : REPAIR_ASK;
            setStatus(wasEmpty || prevEmpty ? "Empty/unclear reply - re-asking firmly…" : "Still unclear - one more clean-JSON round…", "busy");
            return repairRound(nudge2, r1.text, 2).then(function (r2b) {
              if (r2b.parsed.ok) return r2b.parsed;
              var best = r2b.text && r2b.text.length > (resp.text || "").length ? r2b.text : resp.text;
              var salv = salvageOps(best);
              if (!salv || !salv.length) return r2b.parsed;
              SK.lastRecovered = salv.length;
              setStatus("Reply was cut off - recovered " + salv.length + " action(s), asking for the rest…", "busy");
              var cont = messages.concat([
                { role: "assistant", content: best },
                { role: "user", content: continueAsk(ctx) },
              ]);
              return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: cont }); })
                .then(function (r3) {
                  dbg("continue.response", { text: r3.text, elapsedMs: r3.elapsedMs });
                  SK.lastModelText = { text: r3.text };
                  var p3 = tryParse(r3.text);
                  var extra = (p3.ok && p3.ops ? p3.ops : salvageOps(r3.text)) || [];
                  return { ok: true, kind: "slide_ops", ops: salv.concat(extra), recovered: true };
                });
            });
          });
        }).then(function (parsed) {
          if (parsed && parsed.kind === "clarify") {
            // 0.9.235: the model asks the user to pick the deliverable
            SK.clarifyChain = (SK.clarifyChain || 0) + 1;
            dbg("clarify.asked", { chain: SK.clarifyChain, q: String(parsed.question || "").slice(0, 120) });
            if (SK.clarifyChain > 2) {
              parsed = { ok: true, kind: "reply", text: "The clarification limit for this chain is reached (your question, answered in text: " + String(parsed.question || "") + ")" };
            } else {
              SK.lastInstruction = instruction;   // the clarify pair files into the Previous box on the next run
              renderClarify(parsed, instruction);
              setStatus("Pick one of the options (or write your own answer) to continue.", "pending");
              pushHistory(instruction, assistantHistoryText(parsed));
              return Promise.resolve(null);
            }
          }
          var waitApply = renderParsed(parsed, ctx, instruction);
          pushHistory(instruction, assistantHistoryText(parsed));
          return Promise.resolve(waitApply);   // timer/busy wait for the auto-apply to finish
        });
        });
      });
    }).catch(function (err) {
      var msg = err && err.friendly ? err.message : (err && err.message) || String(err);
      // 0.9.302: a CANCELED request never existed - no history pair, no
      // Previous-requests filing (the zone is empty, so the next run's
      // filePreviousRun skips it too), and the model context never sees it.
      // Cancels AFTER the run rendered (auto-apply under way) are real:
      // renderParsed cleared the flag and ops may have landed.
      if (SK.cancelErase && /Run cancelled/i.test(msg)) {
        setStatus("Run cancelled.", "pending");
        dbg("run.cancelled.forget", {});
      } else {
        setStatus(msg, "error");
        // the request left the pane but failed - keep it in the session too
        pushHistory(instruction, "(failed: " + msg + ")");
        // give the request back so the user can just hit Run again
        if (input && !input.value) input.value = instruction;
      }
    }).then(function () {
      stopChatTimer();
      SK.busy = false;
      SK.actRunOn = false;
      setRunEnabled(true);
      input.disabled = false;
      if ($( "chat-output" ) && $("chat-output").children.length) input.value = "";
    });
  }

    // ---- salvage: recover complete ops from a truncated reply (JSON cut off) ----
  function salvageOps(text) {
    var t = String(text == null ? "" : "");
    try { t = String(text); } catch (eT) { return null; }
    var i = t.indexOf('"ops"');
    if (i < 0) return null;
    var arrStart = t.indexOf("[", i);
    if (arrStart < 0) return null;
    var depth = 0, inStr = false, esc = false, objStart = -1, ops = [];
    for (var p = arrStart + 1; p < t.length; p++) {
      var ch = t.charAt(p);
      if (inStr) {
        if (esc) esc = false;
        else if (ch === "\\") esc = true;
        else if (ch === '"') inStr = false;
        continue;
      }
      if (ch === '"') { inStr = true; continue; }
      if (ch === "{") { if (depth === 0) objStart = p; depth++; }
      else if (ch === "}") {
        depth--;
        if (depth === 0 && objStart >= 0) {
          var raw = t.slice(objStart, p + 1);
          try {
            var o = JSON.parse(raw);
            if (o && o.type) ops.push(o);
          } catch (eO) { /* skip incomplete op */ }
          objStart = -1;
        }
      } else if (ch === "]" && depth === 0) break;   // array closed properly
    }
    return ops.length ? ops : null;
  }

  function tryParse(text) {
    var stripped = stripReasoning(text);
    if (!stripped) return { ok: false, text: text };
    var obj = extractObject(stripped);
    if (obj && typeof obj === "object") {
      if (obj.kind === "read_page" && typeof obj.url === "string" && /^https?:\/\//i.test(obj.url)) {
        return { ok: true, kind: "read_page", url: obj.url.split("#")[0],
                 from: (typeof obj.from === "number" && obj.from > 0) ? Math.floor(obj.from) : 0,
                 ops: Array.isArray(obj.ops) ? obj.ops : null, text: stripped };
      }
      if (obj.kind === "web_lookup") {
        // an internal lookup intent: no ops, no approval - the pane runs the
        // search profile and re-asks (must parse before the repair chain)
        // Models sometimes emit look-up-plus-actions in one reply: keep the
        // ops so the runner can apply them after the lookup round.
        return { ok: true, kind: "web_lookup", ops: Array.isArray(obj.ops) ? obj.ops : null, text: stripped };
      }
      if (obj.kind === "clarify" && typeof obj.question === "string" && obj.question.trim()) {
        // 0.9.235: the model asks the user to pick the deliverable; internal
        // action, never gated - the pane renders the question and the re-ask
        // carries the user's decision
        return { ok: true, kind: "clarify", question: obj.question.trim().slice(0, 300),
                 options: (Array.isArray(obj.options) ? obj.options : []).slice(0, 4)
                   .map(function (o) { return String(o == null ? "" : o).trim(); }).filter(Boolean),
                 text: stripped };
      }
      if (obj.kind === "reply" && typeof obj.text === "string") return { ok: true, kind: "reply", text: obj.text };
      if (obj.kind === "slide_ops" && Array.isArray(obj.ops)) return { ok: true, kind: "slide_ops", ops: obj.ops };
      if (Array.isArray(obj.ops)) return { ok: true, kind: "slide_ops", ops: obj.ops }; // tolerate missing kind
    }
    // truncated reply: salvage every complete op from the cut-off stream
    var salv = salvageOps(stripped);
    if (salv) return { ok: true, kind: "slide_ops", ops: salv, recovered: true };
    return { ok: false, text: stripped };
  }



  // ---- auto-retry for model asks: 5 attempts, 10 s apart, on transport
  // errors/server hiccups only (4xx business errors fail immediately) ----
  function askRetryable(askFn, attempt) {
    attempt = attempt || 1;
    if (SK.runCancelled) return Promise.reject(new FriendlyError("Run cancelled."));
    return askFn().catch(function (err) {
      if (SK.runCancelled || (err && (err.name === "AbortError" || /abort/i.test(String(err && err.message)))))
        return Promise.reject(new FriendlyError("Run cancelled."));
      var status = err && err.status;
      var retries_429 = status === 429;
      // provider business errors surface as "Provider returned HTTP 4xx..."
      // inside the message when the pane carries no status object: 429 is
      // retryable, every other 4xx throws immediately (retrying a bad
      // model id 5 times just wastes 50 seconds)
      var emsg = String((err && err.message) || err);
      var biz4xx = !status && /HTTP 4\d\d\b/.test(emsg) && !/HTTP 429\b/.test(emsg);
      var serverish = !status || status >= 500 || retries_429;
      if (!serverish || biz4xx || attempt >= 5) throw err;
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
      // progressive: seconds only, then minutes, then hours
      if (t < 60) { el.textContent = s2 + "s"; }
      else if (t < 3600) { el.textContent = m + "min " + ss + "s"; }
      else { el.textContent = h + "h " + mm + "min " + ss + "s"; }
    }
    tick();
    SK.chatTimer.iv = setInterval(tick, 500);
  }

  function stopChatTimer() {
    if (SK.chatTimer && SK.chatTimer.iv) { clearInterval(SK.chatTimer.iv); SK.chatTimer.iv = null; }
    // the final elapsed time stays visible (frozen) until the next run
  }



  function pmGlyph(open) { return open ? "\u2212" : "+"; }   // boxed minus / plus


  // ---- delete one request/response pair from the session ----
  function removeHistoryPair(ts, grpEl) {
    var n0 = SK.history.length;
    var ts2 = String(ts || "");
    SK.history = SK.history.filter(function (e) { return String(e.ts || "") !== ts2; });
    var removed = n0 - SK.history.length;
    if (grpEl && grpEl.parentNode) grpEl.parentNode.removeChild(grpEl);
    updateChatMeta();
    updatePastBox();
    updateSessionLabel();
    // persist NOW: the server session is rewritten WITHOUT this pair - it
    // will not come back after a reload and future prompts do not include it
    if (SK.sessionId && removed) {
      apiJson("/api/sessions/remove-pair/" + encodeURIComponent(SK.sessionId), "PUT", { ts: ts2 })
        .then(function (r) { dbg("session.pairRemoved", { ts: ts2, removed: r.removed }); })
        .catch(function (e2) { dbg("session.pairRemoveFailed", { ts: ts2, err: String(e2 && e2.message || e2) }); setStatus("Pair removed locally; server sync failed: " + ((e2 && e2.message) || e2), "error"); });
    }
    setStatus(removed ? ("Removed " + (removed / 2) + " exchange(s). Future prompts will not include them.") :
      "No matching exchange found (nothing removed).", removed ? "ok" : "");
  }

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
    var cpIc = document.createElement("span");
    cpIc.className = "pt-copy";
    cpIc.innerHTML = copyIconSvg();               // copy icon next to the step title
    cpIc.title = "Copy this request to the clipboard (and into the input box if it is empty)";
    cpIc.addEventListener("click", function (ev) {
      ev.stopPropagation();
      copyRequestToInput(uc);
    });
    var snip = document.createElement("div");
    snip.className = "pt-text";                // the request text BELOW the ts
    snip.textContent = uc.length > 50 ? uc.slice(0, 50) + "…" : uc;
    snip.appendChild(cpIc);                    // copy affordance on the title line
    var body = document.createElement("div");
    body.className = "reqbody";
    // delete = the same two-stage button as the sessions list (real Office
    // WebViews swallow window.confirm, so the dialog form never appeared)
    var delBtn = document.createElement("button");
    delBtn.className = "btn subtle danger pt-delbtn";
    delBtn.textContent = "Delete";
    delBtn.title = "Remove this request/response pair from the session - a second click within 3.5 s confirms (future prompts will no longer include it)";
    delBtn.addEventListener("click", function (ev) {
      ev.stopPropagation();
      if (delBtn.dataset.armed !== "1") {
        delBtn.dataset.armed = "1";
        delBtn.textContent = "Really delete?";
        delBtn.classList.add("armed");
        setTimeout(function () {
          if (delBtn.parentNode) { delBtn.dataset.armed = "0"; delBtn.textContent = "Delete"; delBtn.classList.remove("armed"); }
        }, 3500);
        return;
      }
      removeHistoryPair(u.ts, grp);
    });
    body.appendChild(delBtn);
    var q1 = document.createElement("div");
    q1.className = "pt-req";
    q1.style.cursor = "pointer";
    q1.title = "Copy this request to the clipboard (and into the input box if it is empty)";
    var b1 = document.createElement("b");
    b1.textContent = "Request: ";
    q1.appendChild(b1);
    var fullReq = uc;
    q1.appendChild(document.createTextNode(fullReq));
    q1.addEventListener("click", function () { copyRequestToInput(fullReq); });
    body.appendChild(q1);
    var q2 = document.createElement("div");
    q2.className = "pt-res";
    q2.style.cursor = "pointer";
    q2.title = "Click to open the full response";
    var b2 = document.createElement("b");
    b2.textContent = "Response: ";
    q2.appendChild(b2);
    var ac = String(a.content || "");
    // 0.9.296: the response row opens COMPLETE + formatted (the same
    // viewer as History) - the box row stays a 400-char preview
    q2.addEventListener("click", function () { openTextViewer("AI response", ac); });
    q2.appendChild(document.createTextNode(ac.length > 400 ? ac.slice(0, 400) + "…" : ac));
    if (ac.length > 400) {
      var hint2 = document.createElement("span");
      hint2.className = "ptts";
      hint2.textContent = "  (click for the full response)";
      q2.appendChild(hint2);
    }
    body.appendChild(q2);
    head.addEventListener("click", function () {
      var isOpen = body.classList.toggle("open");
      pm.textContent = pmGlyph(isOpen);
    });
    snip.addEventListener("click", function () {
      // the excerpt line expands/collapses the step just like +/timestamp
      var isOpen = body.classList.toggle("open");
      pm.textContent = pmGlyph(isOpen);
    });
    snip.style.cursor = "pointer";
    snip.title = "Click to open/close this request/response - the small icon copies it";
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
    var clSlot = $("clarify-slot");
    var clTxt = clSlot ? String(clSlot.textContent || "").replace(/\s+/g, " ").trim() : "";
    if (zone && (zone.childNodes.length || (clTxt && SK.lastInstruction))) {
      var lastUser = null;
      for (var hi = SK.history.length - 1; hi >= 0; hi--) {
        if (SK.history[hi].role === "user") { lastUser = SK.history[hi]; break; }
      }
      var row = pastTurnRow(
        { content: SK.lastInstruction, ts: (lastUser && lastUser.ts) || "" },
        { content: (clTxt ? clTxt + (zone.textContent ? "\n" + zone.textContent : "") : (zone.textContent || "")), ts: (lastUser && lastUser.ts) || "" });
      body.insertBefore(row, body.firstChild);   // newest on top
      while (zone.firstChild) zone.removeChild(zone.firstChild);
      if (clSlot) { clSlot.innerHTML = ""; clSlot.style.display = "none"; }   // the clarify box lives between status and activity
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

  // the last-request echo: click (or its copy button) always copies the
  // request text; when the input box is empty it is also filled with it
  function copyIconSvg() {
    return '<svg width="12" height="12" viewBox="0 0 12 12" style="vertical-align:-2px" aria-hidden="true">' +
      '<rect x="1" y="1" width="8" height="8" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.4"/>' +
      '<rect x="3.5" y="3.5" width="7.5" height="7.5" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.4"/>' +
    '</svg>';
  }
  function copyTextToClipboard2(txt, done) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(txt).then(function () { done(true); }, function () { done(false); });
        return;
      }
    } catch (eN) { /* fall through */ }
    try {
      var ta = document.createElement("textarea");
      ta.value = txt; ta.style.position = "fixed"; ta.style.opacity = "0";
      document.body.appendChild(ta); ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
      done(true);
    } catch (eC) { done(false); }
  }
  function makeReqEchoCopyable(reqEcho, instruction) {
    reqEcho.classList.add("copyable");
    reqEcho.title = "Click to copy this request - it also fills the input box when that is empty";
    var cpBtn = document.createElement("button");
    cpBtn.className = "btn subtle reqecho-copy";
    cpBtn.type = "button";
    cpBtn.innerHTML = copyIconSvg();
    cpBtn.title = "Copy this request to the clipboard";
    var doCopy = function (after) {
      var inp = $("chat-input");
      var wasEmpty = !!(inp && !inp.value.trim());
      copyTextToClipboard2(instruction, function (okc) {
        if (okc && wasEmpty && inp) { inp.value = instruction; }
        if (okc) setStatus(wasEmpty ? "Copied + filled the empty input box." : "Copied to the clipboard.", "ok");
        else setStatus("Could not copy automatically - select the text manually.", "error");
        if (after) after(okc);
      });
    };
    reqEcho.addEventListener("click", function () { doCopy(function (okc) { if (okc) { cpBtn.textContent = "\u2713"; } }); });
    cpBtn.addEventListener("click", function (ev) {
      ev.stopPropagation();
      doCopy(function (okc) { cpBtn.textContent = okc ? "\u2713" : "\u00d7"; setTimeout(function () { cpBtn.innerHTML = copyIconSvg(); }, 1600); });
    });
    reqEcho.appendChild(cpBtn);
  }
  // ---- clarify round (0.9.235): the model may ask the user to pick the ----
  // deliverable instead of guessing; the pick re-asks internally.
  function renderClarify(parsed, instruction) {
    var out = $("clarify-slot");   // sits right below the status line, above the activity box
    if (!out) return;
    out.innerHTML = "";   // one clarify box at a time
    out.style.display = "";
    var box = document.createElement("div");
    box.className = "clarifybox";
    var q = document.createElement("div");
    q.className = "clarify-q";
    q.textContent = String(parsed.question || "Which result do you want?");
    box.appendChild(q);
    var pick = function (val) {
      Array.prototype.forEach.call(box.querySelectorAll("button, input"), function (el) { el.disabled = true; });
      var pickedRow = document.createElement("div");
      pickedRow.className = "clarify-picked";
      pickedRow.textContent = "Your answer: " + String(val);
      box.appendChild(pickedRow);
      // the choice rides INSIDE the instruction: every later re-ask of this
      // chain (pre-lookup / lookup / repair rounds echo the original request)
      // carries it, not only the first ask after the pick
      var augText = String(instruction) + "\n(your earlier choice on \"" + String(parsed.question || "") + "\": " + String(val) + ")";
      SK.clarifyAnswer = { q: String(parsed.question || ""), a: String(val), expect: augText };
      var inp = $("chat-input");
      if (inp) { inp.value = augText; }
      runPrompt();   // the re-ask carries the decision as a CLARIFICATION block
    };
    var opts = (parsed.options || []).slice(0, 4).filter(function (o) { return String(o || "").trim(); });
    opts.forEach(function (o) {
      var b = document.createElement("button");
      b.className = "btn clarify-opt";
      b.type = "button";
      b.textContent = String(o);
      b.addEventListener("click", function () { pick(String(o)); });
      box.appendChild(b);
    });
    var row = document.createElement("div");
    row.className = "clarify-row";
    var ti = document.createElement("input");
    ti.type = "text";
    ti.className = "clarify-input";
    ti.placeholder = "something else - type your own answer\u2026";
    var go = document.createElement("button");
    go.className = "btn";
    go.type = "button";
    go.textContent = "Answer";
    go.disabled = true;
    ti.addEventListener("input", function () { go.disabled = !ti.value.trim(); });
    var send = function () {
      var v = ti.value.trim();
      if (v) pick(v);
    };
    go.addEventListener("click", send);
    ti.addEventListener("keydown", function (eK) {
      if (eK.key === "Enter") { eK.preventDefault(); send(); }
    });
    row.appendChild(ti);
    row.appendChild(go);
    box.appendChild(row);
    out.appendChild(box);
  }

  function renderParsed(parsed, ctx, instruction) {
    SK.cancelErase = false;   // from here the run is real: even a cancel keeps its record
    SK.lastInstruction = instruction;
    var out = $("chat-output");
    out.innerHTML = "";
    // the request prints above the response/summary (echoed by the clear)
    var reqEcho = document.createElement("div");
    reqEcho.className = "reqecho";
    reqEcho.textContent = instruction;
    makeReqEchoCopyable(reqEcho, instruction);
    out.appendChild(reqEcho);
    wirePastHeader();
    if (!parsed.ok) {
      var box = document.createElement("div");
      box.className = "reply-text";
      var showTxt = String(parsed.text || "");
      if (showTxt.length > 420) showTxt = showTxt.slice(0, 420) + "\u2026 (truncated preview - enable debug logging for the full raw output)";
      box.textContent = showTxt || "(empty response)";
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
    if (!validation.ops.length && !validation.invalid.length && (SK.pendingWebOps || []).length) {
      parsed = { ok: true, kind: "slide_ops", ops: SK.pendingWebOps };
      SK.pendingWebOps = null;
      validation = validateOps(parsed.ops, ctx);
    }
    if (!validation.ops.length && !validation.invalid.length) {
      var box3 = document.createElement("div");
      box3.className = "reply-text";
      box3.textContent = (SK.webResults || SK.pendingWebOps) ?
        "(model returned an empty action list - web lookup ran; the model chose not to change the email)" :
        "(model returned an empty action list)";
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
      return applyOps(SK.deck, validation.ops, ctx).then(function (res) {
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

    if (!validation.ops.length && validation.invalid.length > 0) {
      // 0.9.217 - don't reject fast: re-ask the model WITH the reasons first
      var reasons = (validation.invalid || []).map(function (entry) {
        var ot = (entry.op && entry.op.type) || "?";
        return "  - " + ot + ": " + (entry.reason || "unknown");
      }).join("\n");
      var ask = SK.lastAsk;
      if (ask && ask.messages && ask.rounds < 2) {
        ask.rounds++;
        setStatus("All " + validation.invalid.length + " action(s) were rejected - re-asking the model with the reasons (round " + ask.rounds + "/2)…", "busy");
        var repairMsgs = ask.messages.concat([
          { role: "assistant", content: (SK.lastModelText && SK.lastModelText.text) || "" },
          { role: "user", content:
            "NONE of your proposed actions passed validation. The EXACT reasons:\n" + reasons +
            "\nFix every problem and reply again with ONLY the JSON object " +
            '({"kind":"reply","text":"..."} or {"kind":"slide_ops","ops":[...]}) - no prose, no fences. ' +
            "Use only the allowed action types with complete required fields." },
        ]);
        return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ask.connection, model: ask.model, messages: repairMsgs }); })
          .then(function (rR) {
            SK.lastModelText = { text: rR.text };
            var pR;
            try { pR = tryParse(rR.text); } catch (eP) { pR = { ok: false, text: rR.text }; }
            try {
              return renderParsed(pR.ok ? pR : { ok: false, text: rR.text }, ctx, SK.lastInstruction || "");
            } catch (eR2) {
              // 0.9.221: log the FULL trace into the activity box so a
              // broken re-ask self-reports instead of staying opaque
              try { actLog("RE-ASK RENDER FAILED: " + ((eR2 && eR2.stack) || eR2)); } catch (eL) { }
              setStatus("Re-ask render failed: " + ((eR2 && eR2.message) || eR2), "error");
            }
          }).catch(function (eR) {
            try { actLog("RE-ASK FAILED: " + ((eR && eR.stack) || eR)); } catch (eL) { }
            setStatus("Re-ask failed: " + ((eR && eR.message) || eR), "error");
          });
      }
      var allbad = document.createElement("div");
      allbad.className = "status error";
      // self-diagnosing: name the ACTUAL reject reasons (e.g. WRONG DATE),
      // never a one-size-fits-all guess
      var reasons = validation.invalid.slice(0, 3).map(function (iv) {
        return String(iv.reason || "invalid action");
      });
      var rs = reasons.length ? " - " + reasons.join("; ") : "";
      if (validation.invalid.length > 3) {
        rs += " (+" + (validation.invalid.length - 3) + " more; details in the activity log)";
      }
      allbad.textContent = "None of the " + validation.invalid.length + " action(s) survived validation" +
        (SK.lastRecovered ? " (they came from a cut-off reply)" : "") +
        ". No changes were applied." + rs;
      box.appendChild(allbad);
      var reRow = document.createElement("div");
      reRow.className = "row";
      var reBtn = document.createElement("button");
      reBtn.className = "btn";
      reBtn.type = "button";
      reBtn.textContent = "Ask again (same request)";
      reBtn.addEventListener("click", function () {
        var inp = $("chat-input");
        if (SK.busy) return;
        if (!inp.value.trim() && SK.lastInstruction) { inp.value = SK.lastInstruction; }
        runPrompt();
      });
      reRow.appendChild(reBtn);
      box.appendChild(reRow);
      SK.lastRecovered = 0;
      out.appendChild(box);
      return box;
    }

    var tableRows = document.createElement("div");
    rows.forEach(function (row) {
      var r = document.createElement("div");
      row.op._rowEl = r; row.op._gateRow = row;   // progressive apply removes done rows
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
        if (r.parentNode) r.parentNode.removeChild(r);
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
      var liveAuto = rows.filter(function (r) { return !r.removed && !r.done && r.mode === "auto"; }).length;
      var liveConfirm = rows.filter(function (r) { return !r.removed && !r.done && r.mode === "confirm"; }).length;
      var approved = rows.filter(function (r) {
        return !r.removed && !r.done && r.mode === "confirm" && r._cb && r._cb.checked;
      }).length;
      applyBtn.textContent = "Apply " + liveAuto + " auto" +
        (liveConfirm ? " + " + approved + "/" + liveConfirm + " approved" : "");
      applyBtn.disabled = (liveAuto + approved === 0) || !caps.text;
      // partial selection = some confirm rows unticked -> dim as a visual hint
      applyBtn.classList.toggle("halfapply", !!liveConfirm && approved < liveConfirm);
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
          if (r.removed || r.done) return;
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
    if (confirmOps) {
      // make the confirm mechanics impossible to misread: unticked rows are
      // skipped on Apply - they are not "failed", they were never approved
      var cn = document.createElement("div");
      cn.className = "opnote";
      cn.textContent = "(needs approval) actions only run when their checkbox is"
        + " ticked (or use Select all) - unticked ones are skipped on Apply.";
      box.appendChild(cn);
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
      var rememberMsg = "";
      rows.forEach(function (row) {
        if (row.removed || row.done) return;   // X-ed away / already completed
        if (row.mode === "auto") chosen.push(row.op);
        else if (row.mode === "confirm" && row._cb && row._cb.checked) chosen.push(row.op);
      });
      Object.keys(remember).forEach(function (type) {
        if (remember[type].checked) {
          SK.sessionApprovals[type] = "auto";
          var rNote = document.createElement("div");
          rNote.className = "opnote";
          rNote.textContent = "remembered for this session: later '" + OPS_META[type].label +
            "' actions run without asking";
          result.appendChild(rNote);
        }
      });
      // the Office webview can reload the pane page between runs (heavy deck
      // edits) - in-page memory dies with it. Persist the approvals per
      // session id and restore them on the next boot (0.9.267).
      try {
        var sid2 = SK.sessionId;
        if (!sid2) { try { sid2 = localStorage.getItem("mk-resume-mailkick"); } catch (eE1) {} }
        localStorage.setItem("sk-approvals-mailkick", JSON.stringify({ sid: sid2 || "", map: SK.sessionApprovals }));
      } catch (eAp) { /* storage may be unavailable */ }      // unapproved confirm rows would silently vanish - ask before skipping
      var skips = rows.filter(function (row) {
        return !row.removed && !row.done && row.mode === "confirm" && row._cb && !row._cb.checked;
      });
      var counts = {};
      skips.forEach(function (row) { counts[OPS_META[row.op.type].label] = (counts[OPS_META[row.op.type].label] || 0) + 1; });
      var skipList = Object.keys(counts).map(function (l) { return counts[l] + " " + l; }).join(", ");
      var skipNote = (skipList && chosen.length) ? skips.length + " action(s) skipped - not approved: " + skipList : "";
      if (!chosen.length) {
        setStatus(rememberMsg || "Nothing approved - nothing applied.", "");
        return;
      }
      dbg("ops.apply", { approved: chosen.map(function (op) { return op.type; }), count: chosen.length });
      applyBtn.disabled = true;
      // progressive tracking: executed actions leave the proposal list and
      // grow into a completed list one by one; failed actions stay in the
      // list above (numbers + Apply re-arm track the survivors)
      var doneBox = document.createElement("div");
      doneBox.className = "opsbox donebox";
      var doneH = document.createElement("h3");
      var doneCount = 0;
      doneH.textContent = "Completed actions (0)";
      doneBox.appendChild(doneH);
      var doneList = document.createElement("div");
      doneBox.appendChild(doneList);
      box.parentNode.insertBefore(doneBox, box.nextSibling);   // sibling: survives the proposal box closing
      SK.opProgress = function (op, addedLines, ok) {
        if (ok) {
          doneCount++;
          doneH.textContent = "Completed actions (" + doneCount + ")";
          var dr = document.createElement("div");
          dr.className = "oprowline1 doneline";
          dr.textContent = "\u2713 " + ((OPS_META[op.type] || { label: op.type }).label);
          doneList.insertBefore(dr, doneList.firstChild);
          if (op._rowEl && op._rowEl.parentNode) op._rowEl.parentNode.removeChild(op._rowEl);
          if (op._gateRow) op._gateRow.done = true;
        } else if (op._rowEl && op._rowEl.parentNode) {
          op._rowEl.classList.add("opfailrow");
        }
        updateApplyLabel();
      };
      setStatus("Applying " + chosen.length + " op(s)…", "busy");
      applyOps(SK.deck, chosen, SK.lastCtx || ctx).then(function (res) {
        result.innerHTML = "";
        if (skipNote) {
          var sn = document.createElement("div");
          sn.className = "opnote";
          sn.textContent = skipNote + " - tick them and Apply again to run them.";
          result.appendChild(sn);
        }
        res.lines.forEach(function (l) {
          var d = document.createElement("div");
          d.className = "resultline " + (l.ok ? "ok" : "fail");
          d.textContent = l.text;
          result.appendChild(d);
        });
        var liveRows = rows.filter(function (r) { return !r.removed && !r.done; }).length;
        if (!liveRows) {
          box.parentNode.removeChild(box);   // all done: the completed list remains
          SK.opProgress = null;
        } else {
          applyBtn.disabled = false;
          updateApplyLabel();
        }
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
    if (id === "general") {   // a fresh General view: no stale check-now result
      var info = document.getElementById("upd-info");
      if (info) info.textContent = "";
    }
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
    req.then(function (res) {
      var needsKey = !!(res && res.needsKey);
      setFormMsg($("conn-formmsg"),
        needsKey
          ? "Saved connection '" + v.name + "' WITHOUT an API key - every model or chat call will fail with HTTP 401 (Missing Authentication header) until you open the connection and set the key."
          : "Saved connection '" + v.name + "'. Ready for the next one.",
        !needsKey);
      // fields cleared right after saving - the form is fresh for the next connection
      SK.editingConn = null;
      SK.connBaseDirty = false;
      $("conn-form-title").textContent = "New connection";
      $("conn-name").value = "";
      $("conn-type").value = "openai";
      $("conn-base").value = CONN_TYPE_BY_ID["openai"].base;
      $("conn-key").value = "";
      $("conn-key").placeholder = "";
      $("conn-testout").innerHTML = "";
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
    ["prof-content-conn", "prof-other-conn", "prof-image-conn", "prof-web-conn"].forEach(function (id) {
      var sel = $(id);
      fillConnSelect(sel, sel.value || "");
    });
    var list = $("prof-list");
    var rows = [ "<tr><th>Name</th><th>Content model</th><th>Other model</th><th>Image model</th><th>Web model</th></tr>" ];
    (SK.cfg.profiles || []).forEach(function (p) {
      var star = p.name === SK.cfg.defaultProfile ? "*" : "";
      // click the name to edit; deletion lives in the edit form (with confirm)
      var nameCell =
        '<button class="namebtn" data-edit="' + esc(p.name) + '" title="edit">' +
        esc(p.name) + "</button>" + (star ? " <span class=\"opnote\">(default)</span>" : "");
      rows.push(
        "<tr><td>" + nameCell + "</td>" +
        "<td>" + esc(modelRefLabel(p.content)) + "</td><td>" + esc(modelRefLabel(p.other)) + "</td><td>" + esc(modelRefLabel(p.image)) + "</td><td>" + esc(modelRefLabel(p.web)) + "</td></tr>"
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
    fillConnSelect($("prof-web-conn"), p.web && p.web.connection);
    populateModelSelect("prof-web", p.web && p.web.connection, p.web && p.web.model, SK.profWebFilter ? "web" : null);
    setFormMsg($("prof-formmsg"), "", true);
  }

  function resetProfForm() {
    SK.editingProf = null;
    $("prof-form-title").textContent = "New profile";
    $("prof-name").value = "";
    fillConnSelect($("prof-content-conn"), "");
    fillConnSelect($("prof-other-conn"), "");
    fillConnSelect($("prof-image-conn"), "");
    fillConnSelect($("prof-web-conn"), "");
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
    var webConn = $("prof-web-conn").value;
    var webModel = $("prof-web-model").value;
    if (webModel === "__custom__") webModel = $("prof-web-custom").value.trim();
    return {
      name: $("prof-name").value.trim(),
      content: slotValue("prof-content"),
      other: slotValue("prof-other"),
      image: imageConn && imageModel ? { connection: imageConn, model: imageModel } : null,
      web: webConn && webModel ? { connection: webConn, model: webModel } : null,
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
        image: p.image || {}, web: p.web || {},
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
    var fsSel = $("settings-fontsize");
    if (fsSel) fsSel.value = String(parseInt(SK.cfg.paneFontSize, 10) || 100);
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
  function flashSaved(id, text) {
    var el = $(id);
    if (!el) return;
    el.textContent = (text || "\u2713 saved");
    el.classList.add("show");
    clearTimeout(el._t);
    el._t = setTimeout(function () { el.classList.remove("show"); }, 1800);
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
      ["Pane build (JS)", window.__SK_BUILD || "?"],
      ["Add-in identity (this manifest)", (SK.cfg && SK.cfg.manifests && SK.cfg.manifests["mailkick"]) ? "v" + SK.cfg.manifests["mailkick"] : "?"],
      ["All manifest identities", (SK.cfg && SK.cfg.manifests) ? Object.keys(SK.cfg.manifests).map(function (k5) { return k5 + " v" + SK.cfg.manifests[k5]; }).join(" / ") : "?"],
      ["Office host", (function () { try { return Office.context.diagnostics.host + " " + Office.context.diagnostics.version; } catch (e5) { return "n/a"; } })()],
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

  function applyPaneFontSize(pct) {
    // scale the whole pane (text, inputs, paddings) - zoom works in the
    // Office WebViews (Chromium/ WKWebView), px-based CSS would ignore a
    // root font-size change
    var v = Math.max(80, Math.min(160, parseInt(pct, 10) || 100));
    document.body.style.zoom = (v / 100);
  }
  // 0.9.289: the working tab ("Sessions") is disabled until a profile
  // exists. A fresh pane boots straight into Settings > Connections with a
  // purple two-step box (connection first, then profile); the Sessions tab
  // unlocks the moment a profile is saved (every config load re-evaluates).
  function refreshProfileGate() {
    var none = !((((SK.cfg || {}).profiles) || []).length);
    var chatBtn = document.querySelector('#tabs button[data-tab="chat"]');
    if (chatBtn) {
      chatBtn.disabled = none;
      chatBtn.title = none ? "Create a provider connection and a profile first (Settings > Connections)"
                           : "";
    }
    if (none && SK.activeTab === "chat") {
      switchTab("settings");
      switchSettingsSub("connections");
    }
    var panel = document.getElementById("sub-connections");
    if (panel) {
      var warnBox = document.getElementById("profile-gate-warn");
      if (none) {
        if (!warnBox) {
          warnBox = document.createElement("div");
          warnBox.id = "profile-gate-warn";
          panel.insertBefore(warnBox, panel.firstChild);
        }
        warnBox.className = "opnote updavail";
        warnBox.textContent = "Set up in two steps: (1) create at least one CONNECTION to a model provider below (Base URL + optional API key; Test lists its models), then (2) create a PROFILE that points at it (Connections and Profiles tabs are side by side). The Sessions tab unlocks as soon as a profile exists.";
      } else if (warnBox) {
        warnBox.parentNode.removeChild(warnBox);
      }
    }
  }

  function loadConfigAndRender() {
    return apiJson("/api/config", "GET").then(function (cfg) {
      SK.cfg = cfg;
      applyMailSessionBootPolicy();   // 0.9.233: baseline + stale-resume drop
      var nspBox2 = $("settings-newsessionpermail");
      if (nspBox2 && !SK.mnsUserTouched) nspBox2.checked = !!SK.cfg.newSessionPerEmail;
      applyPaneFontSize(cfg.paneFontSize || 100);
      // custom-prompt field: show the effective prompt (custom, else the
      // built-in default); only when the user is not typing in it right now
      var cpEl2 = $("settings-customprompt");
      if (cpEl2 && document.activeElement !== cpEl2) {
        cpEl2.value = String((SK.cfg && SK.cfg.customPrompt) || (SK.cfg && SK.cfg.customPromptDefault) || "");
      }
      var st2 = $("style-image");
      if (st2 && document.activeElement !== st2) st2.value = (SK.cfg && SK.cfg.imageStyle) || "photorealistic";
      // the auto-apply checkbox was wired before the config arrived - keep it
      // in sync with the stored setting from now on (the "off after reopen" bug)
      var sap2 = $("set-autoapply");
      if (sap2 && !SK.autoApplyUserTouched) sap2.checked = autoApplyWanted();
      syncAutoApplyCheckbox();   // and the per-run box follows the global too
      dbg("config.loaded", { connections: (cfg.connections || []).length, profiles: (cfg.profiles || []).length, debug: cfg.debug });
      renderChat();
      renderConnections();
      renderSearchProfiles();
      renderProfiles();
      renderSettings();
      renderAbout();
      refreshProfileGate();
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
      if ($("statusbar")) $("statusbar").className = (window.location && window.location.protocol === "https:") ? "" : "http";   // https = green + lock, http = blue/cyan
      return cfg;
    });
  }

  function wireEvents() {
    Array.prototype.forEach.call(document.querySelectorAll("#tabs button"), function (b) {
      b.addEventListener("click", function () { switchTab(b.dataset.tab); });
    });

    $("chat-run").addEventListener("click", runPrompt);
    var cancelBtn = $("chat-cancel");
    if (cancelBtn) cancelBtn.addEventListener("click", cancelRun);
    var rpos = $("reply-pos"), rneg = $("reply-neg"), rprop = $("reply-propose");
    if (rpos) rpos.addEventListener("click", function () { runQuickReply("pos"); });
    if (rneg) rneg.addEventListener("click", function () { runQuickReply("neg"); });
    if (rprop) rprop.addEventListener("click", renderProposePicker);   // 0.9.305: profile picker first
    $("chat-input").addEventListener("keydown", function (ev) {
      if ((ev.metaKey || ev.ctrlKey) && ev.key === "Enter") { ev.preventDefault(); runPrompt(); }
    });
    $("chat-new").addEventListener("click", function () { startNewSession($("chat-new")); });
    $("chat-refresh-deck").addEventListener("click", function () {
      var mailReset = mailSessionForIdentity(mailCurrentIdentity(), "refresh");
      updateDeckSummary();
      if (!mailReset) setStatus("Email refreshed.", "");   // a mail change reports itself
    });

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
      var imAllToggle = $("prof-image-all");
  var wbAllToggle = $("prof-web-all");
  if (imAllToggle) imAllToggle.checked = !SK.profImageFilter;   // checkbox = SHOW ALL
  if (wbAllToggle) wbAllToggle.checked = !SK.profWebFilter;
  if (imAllToggle) imAllToggle.addEventListener("change", function () {
    // unchecked (default) = image-capable models only; checked = full list
    SK.profImageFilter = !imAllToggle.checked;
    var connName = $("prof-image-conn").value;
    if (connName) populateModelSelect("prof-image", connName, "", SK.profImageFilter ? "image" : null);
  });
  if (wbAllToggle) wbAllToggle.addEventListener("change", function () {
    SK.profWebFilter = !wbAllToggle.checked;
    var connName = $("prof-web-conn").value;
    if (connName) populateModelSelect("prof-web", connName, "", SK.profWebFilter ? "web" : null);
  });

    
    ["content", "other", "image", "web"].forEach(function (slot) {
      $("prof-" + slot + "-conn").addEventListener("change", function () {
        populateModelSelect("prof-" + slot, $("prof-" + slot + "-conn").value, "",
          slot === "image" && SK.profImageFilter ? "image" : slot === "web" && SK.profWebFilter ? "web" : null);
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
        var stEl = $("style-image");
    if (stEl) {
      stEl.innerHTML = [["photorealistic", "photorealistic (default)"], ["cartoon", "cartoon"], ["anime", "anime"],
        ["watercolor", "watercolor"], ["oil painting", "oil painting"], ["flat vector", "flat vector"],
        ["3d render", "3d render"], ["pixel art", "pixel art"],
        ["children's book illustration", "children's book illustration"], ["line art", "line art"], ["infographic", "infographic"]]
        .map(function (o) { return '<option value="' + o[0] + '">' + o[1] + "</option>"; }).join("");
      $("style-image-save").addEventListener("click", function () {
        var v = stEl.value;
        apiJson("/api/settings", "PUT", { imageStyle: v }).then(function (res) {
          if (SK.cfg) SK.cfg.imageStyle = (res && res.imageStyle) || v;
          flashSaved("saved-imagestyle", "\u2713 style saved: " + v);
        }).catch(function () { flashSaved("saved-imagestyle", "save failed"); });
      });
    }
    var cpEl = $("settings-customprompt");
    var cpSave = $("settings-customprompt-save");
    if (cpEl) {
      cpEl.value = String((SK.cfg && SK.cfg.customPrompt)
        || (SK.cfg && SK.cfg.customPromptDefault) || "");   // empty = the built-in prompt
      if (cpSave) cpSave.addEventListener("click", function () {
        apiJson("/api/settings", "PUT", { customPrompt: cpEl.value }).then(function () {
          if (SK.cfg) SK.cfg.customPrompt = cpEl.value;
          flashSaved("saved-customprompt", "\u2713 saved");
        }).catch(function () { flashSaved("saved-customprompt", "save failed"); });
      });
    }
    $("settings-debug").addEventListener("change", function () { saveDebugOption(); flashSaved("saved-debug", SK.cfg && SK.cfg.debug ? "\u2713 saved - debug logging ON" : "\u2713 saved - debug logging OFF"); });
    var _acTimer = null;
    var _acSave = function () {
      clearTimeout(_acTimer);
      _acTimer = setTimeout(function () { saveAutoClean(); flashSaved("saved-autoclean"); }, 500);
    };
    $("settings-autoclean").addEventListener("change", _acSave);
    $("settings-autoclean-days").addEventListener("change", _acSave);
    var _nspBox = $("settings-newsessionpermail");
    if (_nspBox) _nspBox.addEventListener("change", function () {
      SK.mnsUserTouched = true;   // background config loads must not undo this
      var nspOn = !!_nspBox.checked;
      apiJson("/api/settings", "PUT", { newSessionPerEmail: nspOn }).then(function () {
        if (SK.cfg) SK.cfg.newSessionPerEmail = nspOn;
        flashSaved("saved-newsessionpermail", nspOn ? "\u2713 saved - every new email starts a new session" : "\u2713 saved - the session continues across emails");
        mailSessionForIdentity(mailCurrentIdentity(), "setting");   // enabling: sets the baseline, never resets
      }).catch(function () { flashSaved("saved-newsessionpermail", "save failed"); });
    });
    var fsSel2 = $("settings-fontsize");
    if (fsSel2) fsSel2.addEventListener("change", function () {
      var v = parseInt(fsSel2.value, 10) || 100;
      apiJson("/api/settings", "PUT", { paneFontSize: v }).then(function () {
        if (SK.cfg) SK.cfg.paneFontSize = v;
        applyPaneFontSize(v);
        var msg = $("settings-fontsize-msg");
        if (msg) setFormMsg(msg, "Text size saved (" + v + "%).", true);
      }).catch(function () {
        var msg2 = $("settings-fontsize-msg");
        if (msg2) setFormMsg(msg2, "Could not save the text size.", false);
      });
    });
    var updBtn = $("upd-check");
    if (updBtn) updBtn.addEventListener("click", updCheck);
    var cfgBtn = $("cfg-refresh");
    if (cfgBtn) cfgBtn.addEventListener("click", function () {
      cfgBtn.disabled = true;
      setStatus("Re-reading connections and profiles…", "busy");
      loadConfigAndRender().then(function () {
        cfgBtn.disabled = false;
        setStatus("Config re-read: connections and profiles refreshed.", "ok");
      }).catch(function (e2) {
        cfgBtn.disabled = false;
        setStatus("Config re-read failed: " + ((e2 && e2.message) || e2), "error");
      });
    });
    var relBtn = $("pane-reload");
    if (relBtn) relBtn.addEventListener("click", function () {
      dbg("pane.reload", {});
      // hard reload of the pane from the running server, busting any cache;
      // boots into Settings -> General (where updates/user settings live)
      var qs = String(location.search || "").replace(/cb=\d+(&|$)/g, "").replace(/^\?/, "");
      if (/reloadnav=general(&|$)/.test("?" + qs + "&")) qs = qs.replace(/&?reloadnav=general/, "");
      location.replace(location.pathname + "?" + (qs ? qs + "&" : "") + "reloadnav=general&cb=" + Date.now());
    });
    // a pane reload (or back-end update) resumes the session that was active
    var bootResume = null;
    try { bootResume = localStorage.getItem("sk-resume-session-mailkick"); } catch (eS) {}
    if (bootResume) resumeSession(bootResume);
    else restoreSessionApprovals();   // a plain reload restores the same-session approvals
    // a pane reload lands on Settings -> General (the caller tags the URL)
    try {
      var qnav = new URLSearchParams(location.search);
      if (qnav.get("reloadnav") === "general") {
        switchTab("settings");
        switchSettingsSub("general");
        qnav.delete("reloadnav");   // not sticky: a plain reload keeps the normal boot
        var cleanUrl = location.pathname + ("" + qnav ? "?" + qnav.toString() : "");
        try { history.replaceState(null, "", cleanUrl); } catch (eH) { location.replace(cleanUrl); }
      }
    } catch (eQR) { /* older parser: leave the boot view */ }
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
        if (updAutoServer2) updAutoServer2.checked = !(cfg && cfg.serverAutoUpdate === false);
        if (updAuto.checked) updCheck();   // silent availability check at boot
      }, function () { updAuto.checked = true; updCheck(); });
      var updAutoServer2 = $("upd-auto-server");
      updAuto.addEventListener("change", function () {
        apiJson("/api/settings", "PUT", { autoUpdateCheck: !!updAuto.checked });
      });
      if (updAutoServer2) {
        updAutoServer2.addEventListener("change", function () {
          apiJson("/api/settings", "PUT", { serverAutoUpdate: !!updAutoServer2.checked });
        });
      }
    }
    updateSessionLabel();
    wireSessions();
    wireSearchProfiles();
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
      watchMailSelection();   // 0.9.233: reset the session when a different email opens
      applyTheme();
      watchTheme();
      finishBoot();
    });
  }

  // server heartbeat: a small dot at the right end of the tab strip; green
  // while the LOCAL server answers /api/health (this pane's own origin), red
  // when it does not - polled every 5 s, tooltip carries the last check.
  function beatStamp() {
    var d = new Date();
    var p2 = function (n) { return (n < 10 ? "0" : "") + n; };
    return p2(d.getHours()) + ":" + p2(d.getMinutes()) + ":" + p2(d.getSeconds());
  }
  function paintStatusbar(good, ver, stamp) {
    var el = $("statusbar");
    if (!el) return;
    if (good) {
      el.className = (window.location && window.location.protocol === "https:") ? "" : "http";
      el.title = "Local server reachable (v" + (ver || "?") + ") - checked " + (stamp || "");
    if (ver && SK.cfg && SK.cfg.version && SK.cfg.version !== ver) {
      // the server updated (or restarted) under the pane: refresh the
      // bottom banner + the About tab so the version is never stale (0.9.184)
      SK.cfg.version = ver;
      el.textContent = "MailKick v" + ver +
        (window.location && String(window.location.protocol) === "https:" ? " - \u{1F512} " : " - ") +
        "server 127.0.0.1:" + SK.cfg.port + (SK.cfg.debug ? " - debug ON" : "");
      try { renderAbout(); } catch (eVer) {}
    }
      var hintEl2 = $("statusbar-hint");
      if (hintEl2) { hintEl2.style.display = "none"; hintEl2.textContent = ""; }
    } else {
      el.className = "bad";   // red + pulsing (css)
      el.title = "Local server NOT reachable - last check " + (stamp || "");
      var hintEl = $("statusbar-hint");
      if (!hintEl) {
        hintEl = document.createElement("div");
        hintEl.id = "statusbar-hint";
        var foot2 = document.querySelector(".sk-footer");
        if (foot2) foot2.appendChild(hintEl);
      }
      hintEl.style.display = "";
      hintEl.className = "sb-hint";
      hintEl.textContent = startHintText();
    }
  }
  function startHintText() {
    // OS/setup-aware one-liner for the red state
    var p = (SK.cfg && SK.cfg.platform) || (navigator.platform && navigator.platform.indexOf("Win") >= 0 ? "windows" : "macos");
    if (p === "windows") {
      return "Start it again: schtasks /run /tn OfficeKickServer   (or in your clone: py -3 install.py --start)";
    }
    return "Start it again: python3 install.py --start   (in your OfficeKick clone)";
  }
  function srvBeat() {
    var el = $("srvbeat");
    if (!el) return;
    var xr;
    try { xr = new XMLHttpRequest(); } catch (e0) { return; }
    xr.open("GET", "/api/health", true);
    xr.timeout = 4000;
    var stamp = beatStamp();
    xr.onload = function () {
      var good = xr.status === 200, ver = "";
      if (good) {
        try {
          var d = JSON.parse(xr.responseText);
          good = !!d.ok;
          ver = d.version || "";
        } catch (e1) { good = false; }
      }
      el.className = "beat " + (good ? "ok" : "bad");
      el.title = (good ? "Local server reachable (v" + ver + ")" : "Local server NOT reachable")
        + " - last check " + stamp + (good ? "" : ". Start it with: python3 install.py");
      paintStatusbar(good, ver, stamp);
    };
    xr.onerror = xr.ontimeout = function () {
      el.className = "beat bad";
      el.title = "Local server NOT reachable - last check " + stamp
        + ". Start it with: python3 install.py";
      paintStatusbar(false, "", stamp);
    };
    try { xr.send(); } catch (e2) { xr.onerror(); }
  }
  function finishBoot() {
    wireEvents();
    bootLoad(0);
    srvBeat();
    setInterval(srvBeat, 5000);
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
    setStatus("Cannot reach the OfficeKick server: " + err.message + " Is it running? (python3 install.py)", "error");
    $("statusbar").textContent = "server unreachable: " + err.message;
    $("statusbar").className = "bad";
  }

  function boot() {
    // Add-in Error diagnostics: proves the pane JS booted at all (the Office
    // toast fires when the pane never gets this far). Debug logging only.
    dbg("pane.boot", { href: String(window.location.href).slice(0, 160),
      ua: String(navigator.userAgent).slice(0, 140),
      lang: navigator.language, office: (typeof Office !== "undefined") });
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
