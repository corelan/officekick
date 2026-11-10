/*
 * DocKick task pane (runs inside Microsoft Word via Office.js, or in the
 * browser via sim.js for offline testing).
 *
 * Pipeline: free-text instruction -> profile/route -> /api/chat -> reasoning
 * stripping -> JSON ops extraction (hidden repair round on unclear output)
 * -> validation against deck context -> approval gate -> Office.js apply.
 */
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
    app: "wordkick",
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
    clarifyAnswer: null,   // 0.9.235: pending user pick from a clarify round
    clarifyChain: 0,       // clarify rounds served since the original ask
    sessionManual: false,  // the user renamed this session: pause AI titles
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
    move_text:         { label: "move text", hint: "CUT+paste: relocates the paragraph(s) containing find (through until, or count paragraphs) to sit right after place (an anchor text; or end). Styles preserved." },
    move_image:        { label: "move image", hint: "Relocates an inline image (or several in the same paragraph) to sit right after 'place' (an anchor text, or end). find = text inside/near the image's paragraph; needs WordApi 1.6." },
    copy_text:         { label: "copy text", hint: "Paste-COPY: the same block semantics as move_text but the original paragraph(s) stay put." },
    rewrite_doc:       { label: "rewrite document", hint: "Replaces the document body with a new set of paragraphs (destructive - review)." },
    insert_image:      { label: "insert image", hint: "Generates a picture from a description with the profile's image model and inserts it WHERE THE MODEL CHOOSES (place: cursor / an anchor text / end); without a place value, the pane's insert-at-cursor setting applies." },
    format_text:      { label: "format text", hint: "bold/italic/color/size/underline/allCaps/highlight + alignment on one paragraph." },
    create_template:  { label: "create template block", hint: "appends a titled template block with fillable content controls (text, dropdown, checkbox, date)." },
    accept_changes:   { label: "accept tracked changes", hint: "Accepts ALL tracked changes currently in the document (finalizes the AI edits made with revisions on)." },
    set_track_changes: { label: "set track changes", hint: "Turns the document's own Track Changes mode on or off (the Review-tab toggle). Does not touch existing revisions." },
    add_hyperlink:    { label: "add hyperlink", hint: "Links the FIRST occurrence of the given text to a URL (WordApi 1.3)." },
    format_selection: { label: "format selected text", hint: "Applies the same font changes to the text YOU selected right now." },
    reject_changes:   { label: "reject tracked changes", hint: "Rejects ALL tracked changes currently in the document (undo of the AI edits made with revisions on)." },
    set_header:       { label: "set header", hint: "Writes text into the page header of the document (section 1)." },
    set_footer:       { label: "set footer", hint: "Writes text into the page footer of the document (section 1)." },
    create_toc:       { label: "insert table of contents", hint: "Inserts a real TOC field (heading levels 1-3) with an optional title at the end of the document; Word fills the entries when fields update (F9 / print / right-click -> Update Field)." },
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
        info.className = "opnote";
        info.textContent = "You are up to date (" + data.current + ", latest " + data.latest + ").";
      } else {
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
                        ". Reload this pane (About tab -> Reload pane, or quit and reopen the app) to load the new pane files.";
                    } else if (h1 && h1.gen && h1.gen === oldGen && tries > 12) {
                      info.className = "opnote updavail";
                      info.textContent = "The update stopped early - the installer refuses re-registration while Office apps are open (quit Word/PowerPoint/Outlook completely, then click Update again). The server is still on " + (h1.version || "the old build") + ".";
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
    '- {"type":"insert_paragraphs","items":[{"text":"...","style":"normal|title|heading1|heading2|heading3|quote|listParagraph"}],"place":"cursor | end | an anchor text"} (same placement as insert_image; default end; anchor places are what you want for placeholders; an item text prefixed with [L2]..[L9] becomes a real list item at that indent, deepest [L9], [L1] = a top-level bullet)',
    '- {"type":"move_text","find":"existing text","until":"last paragraph text","place":"anchor text"} (CUT one or more consecutive paragraphs - from the one containing find through the one containing until, or when using "count":N the N paragraphs from find - and paste right after the anchor; preserves styles)',
    '- {"type":"copy_text","find":"existing text","until":"last paragraph text or omit","place":"anchor text"} (a paste-COPY: same as move_text but the original block stays put)',
    '- {"type":"move_image","image":N,"place":"anchor text or end"} (relocate the Nth inline picture; image paragraphs are marked [image xN] in the outline)',
    '- {"type":"edit_paragraph","para":N,"text":"..."} (keeps the paragraph style)',
    '- {"type":"replace_text","find":"exact text","replace":"new text","all":true} (all matches by default)',
    '- {"type":"create_table","values":[["h1","h2"],["a","b"]]} (inserted at the end of the document)',
    '- {"type":"insert_comment","find":"exact anchor text","text":"comment"} (comments need WordApi 1.4)',
    '- {"type":"delete_paragraph","para":N} (para = the paragraph number shown in the outline; 1 = the first paragraph - never invent them)',
    '- {"type":"rewrite_doc","items":[{"text":"...","style":"..."}]} (DESTROYS the current body - only when explicitly asked for a full rewrite; item texts prefixed with [L2]..[L9] become real list items at that indent)',
    '- {"type":"create_template","title":"Template name","fields":[{"label":"Field label","type":"text|dropdown|checkbox|date","options":["only for dropdown"]}]} (appends a titled template block with fillable controls; use when the user asks for a template/form/template block; dropdown needs an options array)',
    '- {"type":"insert_image","prompt":"detailed image description","place":"cursor"} (generate a picture from the description and insert it where place says; only when the user asks for a picture/image/illustration/logo)',
    '- {"type":"format_text","para":N,"font":{"bold":true},"align":"center"} (format ONE paragraph from the outline: font = {"bold","italic","color","size","name","underline","allCaps","highlightColor"} subset; align = left|center|right|justify; only what is requested)',
    '- to format only PART of a paragraph, add "find":"exact words appearing in that paragraph" (all matches within the paragraph are formatted; e.g. {"type":"format_text","para":N,"find":"AI","font":{"bold":true}})',
    '- {"type":"add_hyperlink","find":"exact text to link","url":"https://..."} (links the FIRST occurrence of that text; url must be http(s) or mailto:)',
    '- {"type":"format_selection","font":{...same font subset}} (formats the TEXT THE USER HAS SELECTED right now; no para field)',
    '- format_text may also carry "style":"Heading 2|Heading 3|Title|Subtitle|List Bullet|List Number|Quote|Normal" (re-styles the whole paragraph, e.g. make para 4 a bullet list; then format props still apply on top)',
    '- {"type":"accept_changes"} (accepts ALL tracked changes in the document; use ONLY when the user asks to accept/approve/finalize the changes or revisions; list it AFTER the edits that produced them)',
    '- {"type":"reject_changes"} (rejects ALL tracked changes in the document; only when the user explicitly asks to reject/discard the revisions)',
    '- {"type":"set_header","text":"..."} / {"type":"set_footer","text":"..."} (writes text into the page header / footer of section 1; use when the user asks for a header/footer, e.g. a draft marker)',
    '- {"type":"create_toc","title":"Table of Contents"(optional),"levels":"1-3"(optional)} (inserts a REAL table-of-contents FIELD for the given heading levels at the end of the document; Word fills the entries when fields are updated - never fake TOC entries as plain text)',

    '- {"type":"set_track_changes","on":true|false} (turns the DOCUMENT\'s Track Changes mode on/off - the Review-tab toggle; use when the user asks to start/stop tracking; turning it off does NOT remove existing revisions - offer accept_changes for those; the pane\'s separate "AI edits as tracked changes" lives in Settings)',
    "",
    "- LIST INDENTATION (writable [L] markers): the outline shows list items as \"[L2] text\" - no marker = a normal paragraph or top-level bullet. WRITABLE: prefix an item's text with [L2] (second level) up to [L9] (deepest, 9 indent levels) in insert_paragraphs / rewrite_doc - the pane turns it into a real list item with that indentation. Never write [L...] markers anywhere else (not in find values).\n",
    "- STYLES AND SIZES (document typography law): RUNNING TEXT (the body of an essay, article, report or letter) is ALWAYS style \"normal\". Use \"title\" for AT MOST ONE document title paragraph, \"heading1\" for section titles, \"heading2\" for subsections. Emphasis inside body text = bold/italic, never a larger size. If you set font.size, keep it near the document's existing scale (body typically 10-14pt, Title at most 28pt) - never upscale running text.",
    "",
    "Rules:",
    "- Paragraph numbers are 1-based (para 1 = the first paragraph) and MUST come from the outline below. If the user says \"the 3rd paragraph\", that is para 3. Never invent them.",
    "- Rewrite/improve requests: keep the meaning and structure; reword only.",
    "- If the request is a question/explanation/brainstorm rather than a document change, use format A.",
    "- Image requests: always return an insert_image op (never generate the picture inline); DocKick fetches it with the user's configured image model. Include a concrete, self-contained prompt in the op.",
    "- Image refinement ('refine/change the image'): write a FRESH self-contained image prompt describing the FULL desired result. Never embed previous prompts, the conversation, or op JSON into it - only the scene essentials plus the requested change. Prefer 200-600 words.",
    "- LIVE WEB LOOKUPS: when the request needs up-to-date info you cannot reliably know (current events, news, prices, recent releases, rankings, live data), reply with ONLY {\"kind\":\"web_lookup\"} - nothing else. SlideKick then runs a web-enabled lookup and re-asks you with fresh citable results. Never use it for deck/document/mail edits, styling, general concepts, or what the outline/context already answers. Lookups are INTERNAL actions - free (never gated/approved) - and you may reply with later rounds of {\"kind\":\"web_lookup\"} as well (up to 3 lookups per run). After a lookup you may also ask for the FULL text of one page from its results by replying with ONLY {\"kind\":\"read_page\",\"url\":\"https://...\"} (internal, free, up to 4 pages per run). When a page's text carries a \"…[page continues]\" ending and the info you need may sit deeper, read the NEXT part by replying with ONLY {\"kind\":\"read_page\",\"url\":\"https://...\",\"from\":<n>} using the part start the pane gave you (continuations are free and do not use the 4-page budget). You have no web access yourself and cannot run a search; the lookup channel belongs to SlideKick - never claim you tried a lookup or that the lookup channel is unavailable, and never refuse for lack of live data: reply with {\"kind\":\"web_lookup\"} instead. For CONTENT-CREATION requests (write me an essay/article/report/letter/plan/...): the pane has ALREADY run a web lookup before this ask and any fetched results sit in the context above - verify your claims against them, keep facts current and cite the sources you use; if you need MORE, reply {\"kind\":\"web_lookup\"} again (up to 3 rounds per run).",

    "- Picture placement: ALWAYS set 'place' on insert_image - \"cursor\" = the caret/current selection, \"end\" = end of document, or give an ANCHOR (a short exact-ish text from the document) to insert the image right AFTER that text, e.g. \"place\":\"Figure 1: Results\". Omit it only when you truly cannot decide (then the pane's setting applies).",
  ].join("\n");

  var CLARIFY_RULE =
    "\nAMBIGUOUS RESULT FORM or PLACEMENT: (FORM) when the user's ask could plausibly deliver a DIFFERENT FORM of the " +
    "result (e.g. 'make it visual' = an AI-GENERATED PICTURE vs. a TABLE vs. formatted text; 'infographic' vs. " +
    "structured text), or (PLACEMENT) when it is unclear WHERE in the document the result belongs, you may ask the " +
    "user ONCE instead of guessing. Reply exactly:" +
    " {\"kind\":\"clarify\",\"question\":\"<one short question>\",\"options\":[\"<short option>\",\"<short option>\"]} " +
    "(2-4 short options, no JSON inside). The pane shows the question; the user's choice arrives in your next ask " +
    "(look for CLARIFICATION) - then deliver the final result for that choice without asking again. Never use clarify " +
    "for style/tone details you can decide yourself, and never twice in one chain.";
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
          try { localStorage.setItem("sk-resume-session-wordkick", SK.sessionId); } catch (eS) {}
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
      SK.history = (doc.exchanges || []).slice(-HISTORY_MAX_EXCHANGES * 2);
      SK.sessionApprovals = {};   // a remembered approval never travels between sessions
      try { localStorage.removeItem("sk-approvals-wordkick"); } catch (eRC) {}
      SK.sessionId = doc.id;
      try { localStorage.setItem("sk-resume-session-wordkick", doc.id); } catch (eS) {}
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
    try { localStorage.removeItem("sk-approvals-wordkick"); } catch (eRC) {}
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
    "\\b(document|doc|paragraph|paragraphs|text|wording|reword|rewrite|rephrase|outline|comment|comments|table|tables|add|insert|create|make|generate|tighten|restructure|expand|delete|tone|image|images|picture|pictures|photo|illustration|logo|bold|italic|underline|color|font|highlight|format|highlight)\\b|template|form|dropdown|checkbox|date", "i"
  );

  function assistantHistoryText(parsed) {
    // a compact recap for the next round; never the raw ops JSON (it bloats every
    // later request and pushes the models toward echoing blobs into new prompts)
    if (parsed && parsed.kind === "clarify") return "(asked the user to pick: " + String(parsed.question || "").slice(0, 140) + ")";
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
    try { localStorage.removeItem("sk-approvals-wordkick"); } catch (eRC) {}
    SK.sessionId = null;
    try { localStorage.removeItem("sk-resume-session-wordkick"); } catch (eS) {}      // the next exchange saves as a NEW session
    SK.sessionName = "";
    SK.sessionManual = false;
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
    // Word.run does NOT bind the request-builder callback: `this` inside it
    // is not the deck, so this._wrap/this._wholeParaRangeOf and every
    // .bind(this) chain below threw (field: anchored insert_image
    // "this._wholeParaRangeOf is not a function" even after 0.9.283 - the
    // binds chained from the unbound root). Bind ONCE here; every op
    // callback and its nested then-bindings now see the deck.
    return Word.run(fn.bind(this)).then(
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
      paras.load("text, style, isListItem, inlinePictures/items");   // 0.9.226: image markers in the outline
      var sel = ctx.document.getSelection();
      sel.load("text");
      var tables = body.tables;
      tables.load("itemCount");
      var pics = body.inlinePictures;
      pics.load("items");
      return ctx.sync().then(function () {
        // 0.9.301: bake list indent levels as "[L2] text" markers, the same
        // notation SlideKick's deck read uses. Guarded per item - builds or
        // paragraphs without the ListItem API stay flat (no markers).
        var lvlIdx = {};
        for (var iL = 0; iL < paras.items.length; iL++) {
          try {
            if (paras.items[iL].isListItem) {
              paras.items[iL].listItem.load("level");   // access throws ItemNotFound if not a list item
              lvlIdx[iL] = 1;
            }
          } catch (eLI) { /* not a list item - flat */ }
        }
        return ctx.sync().then(function () {
        var n = Math.min(opts.maxParas || paras.items.length, 300);
        var t0 = null;
        var items = [];
        var imageCount = 0;
        for (var i = 0; i < n; i++) {
          var ptxt = paras.items[i].text || "";
          var np = 0;
          try { np = (paras.items[i].inlinePictures && paras.items[i].inlinePictures.items) ? paras.items[i].inlinePictures.items.length : 0; } catch (eI) { }
          if (np > 0) {
            imageCount += np;
            ptxt = (ptxt ? ptxt + " " : "") + "[image x" + np + "]";
          }
          var lvW = 0;
          if (lvlIdx[i]) {
            try { lvW = Number(paras.items[i].listItem.level) || 0; } catch (eLv) { lvW = 0; }
          }
          if (lvW > 0) ptxt = "[L" + Math.min(lvW + 1, 9) + "] " + ptxt;
          items.push({ idx: i, text: ptxt, style: paras.items[i].style || "Normal" });
        }
        return ctx.sync().then(function () {
          return { documentName: "document", paragraphCount: paras.items.length,
                   images: imageCount,
                   paragraphs: items, tables: [], selected: sel.text || "", host: self.capabilities().host };
        });
        });
      });
    });
  };

  WordDeck.prototype.insertParagraphs = function (op) {
    // 0.9.211: same placement vocabulary as insert_image -
    //   place "cursor" -> at the selection; "end" (default) -> append;
    //   any other string = ANCHOR: insert after the first body match.
    // 0.9.301: [L2]..[L9] item prefixes become real list indentation
    // (level writes run AFTER the text/style sync; failures degrade to a
    // visible note on the action row and can never lose the text).
    var self = this;
    return this.run("insertParagraphs", function (ctx) {
      var body = ctx.document.body;
      var items = op.items || [];
      var place = (typeof op.place === "string" && op.place.trim()) ? op.place.trim() : "end";
      var parsed = [];
      items.forEach(function (it) {
        var pl = parseListMarkers([String(it.text == null ? "" : it.text)]);
        parsed.push({ text: pl.lines[0], style: it.style, lv: pl.levels[0] || 0, markered: pl.any });
      });
      var mkStyle = function (it) {
        // a markered item is a list line: force the built-in list style for
        // its level (ListBullet..ListBullet5), overriding the model's style
        return it.markered ? wordBulletToken(it.lv) : styleNameOf(it.style);
      };
      var lvlQueue = [];   // {p: paragraph proxy, lv: 0-based level}
      if (place === "end") {
        parsed.forEach(function (it) {
          var p = body.insertParagraph(it.text, "End");
          setWordStyle(p, mkStyle(it));
          if (it.markered) lvlQueue.push({ p: p, lv: it.lv });
        });
        return ctx.sync().then(function () {
          return self._applyListLevels(ctx, lvlQueue, "end of document");
        });
      }
      if (place === "cursor") {
        var sel = ctx.document.getSelection();
        parsed.forEach(function (it) {
          var p = sel.insertParagraph(it.text, "After");
          setWordStyle(p, mkStyle(it));
          if (it.markered) lvlQueue.push({ p: p, lv: it.lv });
        });
        return ctx.sync().then(function () {
          return self._applyListLevels(ctx, lvlQueue, "at caret/selection");
        });
      }
      var rs = body.search(place);
      rs.load("items");
      return ctx.sync().then(function () {
        var match = rs && rs.items && rs.items[0];
        var how = match ? "after the paragraph containing '" + place.slice(0, 60) + "'"
                        : "anchor '" + place.slice(0, 60) + "' not found - end of document";
        if (!match) {
          parsed.forEach(function (it2) {
            var p2 = body.insertParagraph(it2.text, "End");
            setWordStyle(p2, mkStyle(it2));
            if (it2.markered) lvlQueue.push({ p: p2, lv: it2.lv });
          });
          return ctx.sync().then(function () {
            return self._applyListLevels(ctx, lvlQueue, how);
          });
        }
        return _wholeParaRange(ctx, match).then(function (pr) {
          var paraRange = pr || match;
          parsed.forEach(function (it3) {
            var p3 = paraRange.insertParagraph(it3.text, "After");
            setWordStyle(p3, mkStyle(it3));
            if (it3.markered) lvlQueue.push({ p: p3, lv: it3.lv });
          });
          return ctx.sync().then(function () {
            return self._applyListLevels(ctx, lvlQueue, how);
          });
        });
      });
    });
  };

  WordDeck.prototype._applyListLevels = function (ctx, queue, how) {
    // 0.9.301: ListItem.level for the markered paragraphs, after their
    // text/style landed. Word.js p.listItem THROWS ItemNotFound at ACCESS
    // when the paragraph is not a list item - those are counted flat. A
    // level write rejected at SYNC degrades to a note; the text stays.
    var wrote = 0, failed = 0;
    (queue || []).forEach(function (q) {
      try { q.p.listItem.level = q.lv; wrote++; } catch (eLI) { failed++; }
    });
    return ctx.sync().then(function () {
      var note = "";
      if ((queue || []).length) {
        note = failed
          ? " - list indentation set on " + wrote + "/" + queue.length + " item(s) (" + failed + " kept flat)"
          : " - list indentation applied (" + wrote + ")";
      }
      return { ok: true, how: how + note };
    }, function (eLev) {
      return { ok: true, how: how + " - this Word build rejected list indentation (" +
        String((eLev && eLev.message) || eLev).slice(0, 80) + ") - text kept flat" };
    });
  };

  WordDeck.prototype.rewriteDoc = function (op) {
    var self = this;
    return this.run("rewriteDoc", function (ctx) {
      var body = ctx.document.body;
      body.clear();
      var last = null;
      var lvlQueue = [];
      (op.items || []).forEach(function (it) {
        var pl = parseListMarkers([String(it.text == null ? "" : it.text)]);
        var txt = pl.lines[0];
        var st = pl.any ? wordBulletToken(pl.levels[0] || 0) : styleNameOf(it.style);
        var p = body.insertParagraph(txt, last ? "After" : "End");
        last = p;
        setWordStyle(p, st);
        if (pl.any) lvlQueue.push({ p: p, lv: pl.levels[0] || 0 });
      });
      return ctx.sync().then(function () {
        return self._applyListLevels(ctx, lvlQueue, "document rewritten");
      });
    });
  };
  WordDeck.prototype.editParagraph = function (index, text) {
    return this.run("editParagraph", function (ctx) {
      // Word.js collections: use .items after load+sync (no getItemAt here)
      var paras = ctx.document.body.paragraphs;
      paras.load("text, style");
      return ctx.sync().then(function () {
        var p = paras.items[index];
        if (!p) throw new FriendlyError("Paragraph " + (index + 1) + " does not exist");   // 0.9.307: op numbering is 1-based
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

  WordDeck.prototype.moveImage = function (op) {
    // 0.9.220 - {"type":"move_image","image":N,"place":"anchor text or end"} (relocate the Nth inline picture; image paragraphs appear as "[image x1]" in the outline - also accepts "find":"text around the image")',
    // Find the paragraph whose text contains `find`; if that paragraph
    // contains inline pictures, copy their base64 sources, insert them after
    // the NEW anchor (or at the end), then delete the source paragraph.
    return this.run("moveImage", function (ctx) {
      var find = String(op.find || "");
      if (!find) throw new FriendlyError("move_image: missing find text");
      var place = (typeof op.place === "string" && op.place.trim()) ? op.place.trim() : "end";
      var paras = ctx.document.body.paragraphs;
      paras.load("text, style");   // 0.9.222
      var found = ctx.document.body.search(find, { matchCase: false, matchWholeWord: false });
      found.load("items");
      return ctx.sync().then(function () {
        var srcPara = null;
        if (typeof op.image === "number" && op.image >= 1) {
          // 0.9.226: image:N - walk the paragraphs until the Nth inline picture is reached
          var target = Math.floor(op.image);
          var seen = 0;
          (paras.items || []).some(function (pr) {
            try { pr.load("inlinePictures/items"); } catch (eIp) { return false; }
            return false;
          });
          return ctx.sync().then(function () {
            for (var ip = 0; ip < (paras.items || []).length; ip++) {
              var pIp = paras.items[ip];
              var cnt = 0;
              try { cnt = (pIp.inlinePictures && pIp.inlinePictures.items) ? pIp.inlinePictures.items.length : 0; } catch (eIc) { }
              if (cnt > 0) {
                seen += cnt;
                if (seen >= target) { srcPara = pIp; break; }
              }
            }
            if (!srcPara) throw new FriendlyError("move_image: inline picture " + target + " not found (document has fewer images)");
            var srcText0 = srcPara.text || "";
            return (function () {
              var b64s = [];
              try {
                if (srcPara.pictures) { srcPara.pictures.load("items"); }
              } catch (eP2) { }
              var collect2 = ctx.sync().then(function () {
                if (srcPara.pictures && srcPara.pictures.items) {
                  srcPara.pictures.items.forEach(function (p) {
                    try { b64s.push(p.getBase64ImageSrc()); } catch (eG2) { }
                  });
                  return ctx.sync();
                }
                return null;
              }).then(function () { return b64s; });
              return collect2.then(function () {
                if (!b64s.length) throw new FriendlyError("move_image: picture " + target + " could not be read");
                var inserter2 = function (rng) {
                  var chainB = Promise.resolve();
                  b64s.forEach(function (getSrc) {
                    chainB = chainB.then(function () {
                      return Promise.resolve(getSrc).then(function (src) {
                        (rng || ctx.document.body).insertInlinePictureFromBase64(src, rng ? "After" : "End");
                        return ctx.sync();
                      });
                    });
                  });
                  return chainB.then(function () {
                    // image-only paragraphs are deleted whole; text paragraphs keep their text (a copy was inserted)
                    var deleted = false;
                    if (!String(srcPara.text || "").trim() && srcPara.pictures && srcPara.pictures.items && srcPara.pictures.items.length === b64s.length) {
                      try { srcPara.delete(); deleted = true; } catch (eD) { }
                      return ctx.sync().then(function () {
                        return { ok: true, how: (b64s.length) + " image(s) - " + (place === "end" ? "to end of document" : "after '" + place.slice(0, 60) + "'") + (deleted ? " (moved)" : " (copy; original kept because sharing the paragraph with text)") };
                      });
                    }
                    return Promise.resolve({ ok: true, how: (b64s.length) + " image(s) - " + (place === "end" ? "to end of document" : "after '" + place.slice(0, 60) + "'") + " (copy; image shares a paragraph with text - delete the original manually)" });
                  });
                };
                if (place === "end") return inserter2(ctx.document.body);
                var rs = ctx.document.body.search(place);
                rs.load("items");
                return ctx.sync().then(function () {
                  var match = rs && rs.items && rs.items[0];
                  if (!match) return inserter2(ctx.document.body).then(function (r) {
                    r.how = "anchor '" + place.slice(0, 60) + "' not found - " + r.how;
                    return r;
                  });
                  return _wholeParaRange(ctx, match).then(function (pr) {
                    return inserter2(pr || match);
                  }.bind(this));
                }.bind(this));
              });
            })();
          });
        }
        if (!(found.items && found.items.length)) throw new FriendlyError("move_image: text not found: " + find.slice(0, 60));
        var srcPara = null;
        (paras.items || []).some(function (pr) {
          if (pr.text && pr.text.indexOf(find) >= 0) { srcPara = pr; return true; }
          return false;
        });
        if (!srcPara) throw new FriendlyError("move_image: paragraph containing '" + find.slice(0, 50) + "' not resolved");
        // getBase64ImageSrc (WordApi 1.6): fetch the inline pictures of THAT paragraph
        return ctx.sync().then(function () {
          var b64calls = [];
          try {
            if (srcPara.pictures) { srcPara.pictures.load("items"); }
          } catch (eP) { /* WordApi too old - b64calls stays empty */ }
          var b64s = [];
          var collect = ctx.sync().then(function () {
            if (srcPara.pictures && srcPara.pictures.items) {
              srcPara.pictures.items.forEach(function (p) {
                try { b64s.push(p.getBase64ImageSrc()); } catch (eG) { }
              });
              return ctx.sync();
            }
            return null;
          }).then(function () { return b64s; });
          return collect.then(function () {
            if (!b64s.length) throw new FriendlyError("move_image: no inline picture in the paragraph containing '" + find.slice(0, 50) + "'");
            // insert at the target first, THEN delete the source paragraph
            var inserter = function (rng) {
              var chainB = Promise.resolve();
              b64s.forEach(function (getSrc) {
                chainB = chainB.then(function () {
                  return Promise.resolve(getSrc).then(function (src) {
                    (rng || ctx.document.body).insertInlinePictureFromBase64(src, rng ? "After" : "End");
                    return ctx.sync();
                  });
                });
              });
              return chainB.then(function () {
                srcPara.delete();
                return ctx.sync().then(function () {
                  return { ok: true, how: (b64s.length) + " image(s) " + (place === "end" ? "to end of document" : "after '" + place.slice(0, 60) + "'") };
                });
              });
            };
            if (place === "end") return inserter(ctx.document.body);
            var rs = ctx.document.body.search(place);
            rs.load("items");
            return ctx.sync().then(function () {
              var match = rs && rs.items && rs.items[0];
              if (!match) return inserter(ctx.document.body).then(function (r) {
                r.how = "anchor '" + place.slice(0, 60) + "' not found - " + r.how;
                return r;
              });
              return _wholeParaRange(ctx, match).then(function (pr) {
                return inserter(pr || match);
              }.bind(this));
            }.bind(this));
          });
        });
      });
    });
  };

  WordDeck.prototype._blockSlice = function (paras, op) {
    // 0.9.212: resolve the paragraph block op.describe: from the para
    // containing op.find, up to (including) para containing op.until, or
    // op.count consecutive paragraphs (default 1).
    var find = String(op.find || "");
    if (!find) throw new FriendlyError("missing find text");
    var items = (paras.items || []);
    var fromI = -1;
    for (var i = 0; i < items.length; i++) {
      var txt = items[i].text || "";
      if (txt.indexOf(find) >= 0) { fromI = i; break; }
    }
    if (fromI < 0) throw new FriendlyError("text not found: " + find.slice(0, 60));
    var toI = fromI;
    if (typeof op.until === "string" && op.until.trim()) {
      var until = op.until.trim();
      for (var j = fromI; j < items.length; j++) {
        if ((items[j].text || "").indexOf(until) >= 0) { toI = j; break; }
      }
      if (toI === fromI && (items[fromI].text || "").indexOf(until) < 0) {
        throw new FriendlyError("until text not found after '" + find.slice(0, 50) + "'");
      }
    } else if (typeof op.count === "number" && op.count >= 1) {
      toI = Math.min(fromI + Math.floor(op.count) - 1, items.length - 1);
    }
    var block = [];
    for (var k = fromI; k <= toI; k++) {
      try { items[k].load("text, style"); } catch (eL1) { }   // 0.9.223: belt+braces materialize on each object
      block.push(items[k]);
    }
    return block;
  };

  WordDeck.prototype._insertBlockAfter = function (ctx, block, place) {
    // inserts block paragraphs (text+style) after `place` ("end" or anchor);
    // returns a string describing the result.
    var body = ctx.document.body;
    var doInsert = function (rng) {
      var last = null;
      block.forEach(function (pr) {
        last = rng ? rng.insertParagraph(pr.text || "", "After") : body.insertParagraph(pr.text || "", "End");
        var st = null;
        try { st = pr.style; } catch (eS) { }   // 0.9.223: never crash on an un-materialized style read
        if (st) { try { last.style = st; } catch (eW) { } }
        rng = last;   // keep appending AFTER the previously inserted one
      });
      return ctx.sync();
    };
    if (!place || place === "end") {
      return doInsert(null).then(function () { return { how: "end of document" }; });
    }
    var rs = body.search(place);
    rs.load("items");
    return ctx.sync().then(function () {
      var match = rs && rs.items && rs.items[0];
      if (!match) {
        return doInsert(null).then(function () {
          return { how: "anchor '" + place.slice(0, 60) + "' not found - end of document" };
        });
      }
      return _wholeParaRange(ctx, match).then(function (pr) {
        return doInsert(pr || match).then(function () {
          return { how: "after the paragraph containing '" + place.slice(0, 60) + "'" };
        });
      }.bind(this));
    }.bind(this));
  };

  WordDeck.prototype._pasteBlock = function (op, cut) {
    var self = this;
    return this.run(cut ? "moveText" : "copyText", function (ctx) {
      var paras = ctx.document.body.paragraphs;
      paras.load("text");
      return ctx.sync().then(function () {
        var block = self._blockSlice(paras, op);
        var place = (typeof op.place === "string" && op.place.trim()) ? op.place.trim() : "end";
        return self._insertBlockAfter(ctx, block, place).then(function (res) {
          if (cut) block.forEach(function (pr) { pr.delete(); });
          return ctx.sync().then(function () {
            return { ok: true, how: res.how, moved: cut ? block.length : 0, count: block.length };
          });
        });
      });
    });
  };

  WordDeck.prototype.moveText = function (op) {
    // 0.9.212: cut/paste-grade - moves the paragraph containing `find`
    // (optionally through `until` or `count` paragraphs) to sit right
    // after `place` (anchor) or at the end. 0.9.211 single-para kept.
    return this._pasteBlock(op, true);
  };

  WordDeck.prototype.copyText = function (op) {
    // duplicate the block: same find/until/count/place, no delete.
    return this._pasteBlock(op, false);
  };

  WordDeck.prototype.deleteParagraph = function (index) {
    return this.run("deleteParagraph", function (ctx) {
      var paras = ctx.document.body.paragraphs;
      paras.load("text");
      return ctx.sync().then(function () {
        var p = paras.items[index];
        if (!p) throw new FriendlyError("Paragraph " + (index + 1) + " does not exist");   // 0.9.307: op numbering is 1-based
        p.delete();
        return ctx.sync().then(function () { return { ok: true }; });
      });
    });
  };

  WordDeck.prototype.formatText = function (index, font, style) {
    return this.run("formatText", function (ctx) {
      // Word.js collections have no getItemAt: load + sync, then index .items
      var paras = ctx.document.body.paragraphs;
      paras.load("text");
      return ctx.sync().then(function () {
        var para = paras.items[index];
        if (!para) {
          throw new FriendlyError("Paragraph " + (index + 1) + " not found (document has " + paras.items.length + " paragraphs).");
        }
        var range = para.getAsRange
          ? para.getAsRange()
          : para.getRange("Whole");
        if (style) range.styleBuiltIn = style;   // WordApi 1.3 built-in token
        var rf = range.font;
      if (font.bold != null) rf.bold = font.bold;
      if (font.italic != null) rf.italic = font.italic;
      if (font.allCaps != null) rf.allCaps = font.allCaps;
      if (font.underline != null) rf.underline = font.underline ? "Single" : "None";
      if (font.color) rf.color = font.color;
      if (font.name) rf.name = font.name;
      if (font.size) rf.size = font.size;
      if (font.highlightColor) rf.highlightColor = font.highlightColor;
        if (font.align) para.alignment = { left: "Left", center: "Centered", right: "Right", justify: "Justified" }[font.align];
        return ctx.sync().then(function () { return { ok: true }; });
      });
    });
  };

  WordDeck.prototype.createTemplateBlock = function (title, fields) {
    return this.run("createTemplateBlock", function (ctx) {
      if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) {
        throw new FriendlyError("content controls need WordApi 1.5 - update Word (16.70+ on Mac).");
      }
      var body = ctx.document.body;
      body.insertParagraph(title, "End").styleBuiltIn = "heading2";
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
      // reading/setting changeTrackingMode needs WordApi 1.6 AND an explicit
      // load + sync first - an unguarded proxy read throws
      // "The property 'changeTrackingMode' is not available..."
      if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) {
        throw new FriendlyError("this Word build does not support change-tracking mode (needs WordApi 1.6) - edits will be applied directly.");
      }
      var doc = ctx.document;
      doc.load("changeTrackingMode");
      return ctx.sync().then(function () {
        var previous = doc.changeTrackingMode;
        doc.changeTrackingMode = on ? "trackAll" : "off";
        return ctx.sync().then(function () {
          return { ok: true, previous: previous };
        });
      });
    });
  };

  WordDeck.prototype.acceptChanges = function () {
    return this.run("acceptChanges", function (ctx) {
      // TrackedChangeCollection.acceptAll() - WordApi 1.6
      if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) {
        throw new FriendlyError("accepting tracked changes needs WordApi 1.6 - update Word.");
      }
      var tcs = ctx.document.body.getTrackedChanges();
      tcs.acceptAll();
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  WordDeck.prototype.addHyperlink = function (findText, url) {
    return this.run("addHyperlink", function (ctx) {
      // Range.hyperlink exists since WordApi 1.3
      if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
        throw new FriendlyError("hyperlinks need WordApi 1.3 - update Word.");
      }
      var range = ctx.document.body.getRange("Whole");
      var found = range.search(String(findText), {});
      found.load("items");
      return ctx.sync().then(function () {
        var items = found.items || [];
        if (!items.length) {
          throw new FriendlyError("'" + findText + "' not found in the document.");
        }
        items[0].hyperlink = String(url);
        return ctx.sync().then(function () { return { ok: true, count: items.length }; });
      });
    });
  };

  WordDeck.prototype.formatSelection = function (font) {
    return this.run("formatSelection", function (ctx) {
      var sel = ctx.document.getSelection();
      if (font.bold != null) sel.font.bold = font.bold;
      if (font.italic != null) sel.font.italic = font.italic;
      if (font.allCaps != null) sel.font.allCaps = font.allCaps;
      if (font.underline != null) sel.font.underline = font.underline ? "Single" : "None";
      if (font.color) sel.font.color = font.color;
      if (font.highlightColor) sel.font.highlightColor = font.highlightColor;
      if (font.name) sel.font.name = font.name;
      if (font.size) sel.font.size = font.size;
      if (font.align) sel.paragraphFormat.alignment = { left: "Left", center: "Centered", right: "Right", justify: "Justified" }[font.align];
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  WordDeck.prototype.rejectChanges = function () {
    return this.run("rejectChanges", function (ctx) {
      // TrackedChangeCollection.rejectAll() - WordApi 1.6 (mirror of acceptAll)
      if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) {
        throw new FriendlyError("rejecting tracked changes needs WordApi 1.6 - update Word.");
      }
      var tcs = ctx.document.body.getTrackedChanges();
      tcs.rejectAll();
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  WordDeck.prototype.setHeaderFooter = function (text, isFooter) {
    return this.run("setHeaderFooter", function (ctx) {
      // sections -> getHeader/getFooter exist since WordApi 1.1
      if (!Office.context.requirements.isSetSupported("WordApi", "1.1")) {
        throw new FriendlyError("headers/footers need WordApi 1.1 - unsupported host.");
      }
      var sections = ctx.document.sections;
      sections.load("items");
      return ctx.sync().then(function () {
        var section = sections.items[0];
        if (!section) throw new FriendlyError("no section found in this document.");
        var hf = isFooter ? section.getFooter("primary") : section.getHeader("primary");
        if (typeof hf.insertText !== "function" || typeof hf.clear !== "function") {
          throw new FriendlyError("this Word build does not allow editing the " + (isFooter ? "footer" : "header") + " here.");
        }
        hf.clear();
        hf.insertText(String(text), "End");
        return ctx.sync().then(function () { return { ok: true }; });
      });
    });
  };

  WordDeck.prototype.createToc = function (op) {
    return this.run("createToc", function (ctx) {
      // insertOoxml exists since WordApi 1.1; a real TOC field (w:fldChar +
      // instrText) - Word fills the entries when the field updates (F9)
      var body = ctx.document.body;
      var lvl = String(op.levels || "1-3");
      var instr = " TOC \\o \"" + lvl + "\" \\h \\z \\u ";
      var escT = function (s) {
        return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      };
      var titleP = "";
      if (op.title && op.title.trim()) {
        titleP = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>' +
          '<w:r><w:t xml:space="preserve">' + escT(op.title.trim()) + "</w:t></w:r></w:p>";
      }
      var tocP = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve">' + escT(instr) + "</w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:t xml:space="preserve">Table of contents - select it and press F9 to fill the entries.</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
      return body.insertOoxml(titleP + tocP, Word.InsertLocation.end).then(function () {
        return ctx.sync();
      }).then(function () { return { ok: true }; });
    });
  };

  function q34() { return String.fromCharCode(34); }
  WordDeck.prototype.formatTextFind = function (index, findText, font) {
    return this.run("formatTextFind", function (ctx) {
      // Word.js collections have no getItemAt: load + sync, then index .items
      var paras = ctx.document.body.paragraphs;
      paras.load("text");
      return ctx.sync().then(function () {
        var para = paras.items[index];
        if (!para) {
          throw new FriendlyError("Paragraph " + (index + 1) + " not found (document has " + paras.items.length + " paragraphs).");
        }
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
    });
  };

  function _wholeParaRange(ctxX, matchRange) {
    // 0.9.285 - module-local twin of _wholeParaRangeOf (no `this` needed)
    // 0.9.227 - expands a text-match range to the WHOLE paragraph that
    // contains it (so inserts with "After" land after the paragraph, never
    // mid-sentence). Returns the paragraph range or null.
    var paras = ctxX.document.body.paragraphs;
    paras.load("text");
    return ctxX.sync().then(function () {
      var matchTxt = (matchRange.text || "");
      if (!matchTxt) return matchRange;
      for (var i = 0; i < paras.items.length; i++) {
        var ptxt = paras.items[i].text || "";
        if (ptxt.indexOf(matchTxt.trim()) >= 0) {
          var pr = paras.items[i];
          return pr.getAsRange ? pr.getAsRange() : pr.getRange("Whole");
        }
      }
      return matchRange;
    });
  }

  WordDeck.prototype._wholeParaRangeOf = _wholeParaRange;   // delegate

  WordDeck.prototype.addImage = function (payload) {
    // 0.9.209 - the MODEL (or the pane setting) decides where to insert:
    //   payload.place === "cursor"  -> the caret / selection ("Replace")
    //   payload.place === "end"     -> end of document
    //   any other non-empty string  -> ANCHOR text: insert right AFTER the
    //      first body match; anchor not found -> end of document (reported)
    //   missing/empty               -> caret when atCursor !== false, else end
    var place = typeof payload.place === "string" && payload.place.trim()
      ? payload.place.trim() : (payload.atCursor !== false ? "cursor" : "end");
    return this.run("addImage", function (ctx) {
      if (place === "cursor") {
        ctx.document.getSelection().insertInlinePictureFromBase64(String(payload.b64), "Replace");
        return ctx.sync().then(function () { return { ok: true, how: "at caret/selection" }; });
      }
      if (place === "end") {
        ctx.document.body.insertInlinePictureFromBase64(String(payload.b64), "End");
        return ctx.sync().then(function () { return { ok: true, how: "at end of document" }; });
      }
      var rs = ctx.document.body.search(place);
      rs.load("items");
      return ctx.sync().then(function () {
        if (rs && rs.items && rs.items.length) {
          return _wholeParaRange(ctx, rs.items[0]).then(function (pr) {
            (pr || rs.items[0]).insertInlinePictureFromBase64(String(payload.b64), "After");
            return ctx.sync().then(function () { return { ok: true, how: "after the paragraph containing '" + place.slice(0, 60) + "'" }; });
          }.bind(this));
        }
        ctx.document.body.insertInlinePictureFromBase64(String(payload.b64), "End");
        return ctx.sync().then(function () {
          return { ok: true, how: "anchor '" + place.slice(0, 60) + "' not found - inserted at end of document" };
        });
      }.bind(this));
    });
  };

  WordDeck.prototype.rewriteDoc = function (op) {
    return this.run("rewriteDoc", function (ctx) {
      var body = ctx.document.body;
      body.clear();
      var last = null;
      (op.items || []).forEach(function (it) {
        last = body.insertParagraph(String(it.text == null ? "" : it.text), last ? "After" : "End");
        setWordStyle(last, styleNameOf(it.style));
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

  // 0.9.281: insert_paragraphs/rewriteDoc assigned paragraph.style with
  // ENGLISH display names ("Title", "Heading 1") - .style is the LOCALIZED
  // style name, so French Word (Titre/Titre 1) threw InvalidArgument for
  // every styled paragraph (field: fresh Windows 11 fr-FR box - the whole
  // insert failed even though the paragraphs themselves inserted fine on
  // en-US builds). Assign styleBuiltIn with the Word.BuiltInStyleName
  // token instead: identical tokens on every UI language
  // (field report 0.9.281). The .style fallback keeps en-only builds whole;
  // if that also fails the paragraph keeps the default style and the op
  // still succeeds.
  function styleNameOf(kind) {
    if (!kind || String(kind).toLowerCase() === "normal") return null; // doc default
    var k = String(kind).toLowerCase();
    if (k === "title") return "title";
    if (k === "heading1") return "heading1";
    if (k === "heading2") return "heading2";
    if (k === "heading3") return "heading3";
    if (k === "quote") return "quote";
    if (k === "listparagraph") return "listParagraph";
    return null;
  }

  function setWordStyle(paragraph, token) {
    if (!token || !paragraph) return;
    try {
      paragraph.styleBuiltIn = token;   // WordApi 1.3 - localization-free
    } catch (e) {
      try { paragraph.style = token; } catch (e2) { /* keep default style */ }
    }
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
      var saved = JSON.parse(localStorage.getItem("sk-approvals-wordkick") || "null");
      if (saved && saved.map && (saved.sid || "") === (SK.sessionId || "")) {
        SK.sessionApprovals = saved.map;
        return true;
      }
    } catch (eR2) { /* corrupt entry - start clean */ }
    return false;
  }

  function slideOkF(idx, ctx) {
    return typeof idx === "number" && ctx && ctx.paragraphs && ctx.paragraphs[idx - 1];
  }
  function parseListMarkers(lines) {
    // 0.9.301: [L2]..[L9] prefixes on insert_paragraphs / rewrite_doc item
    // texts are WRITABLE list indentation - the same notation the doc read
    // renders (the read bakes markers on list items). Strips the markers
    // and returns per-line levels; Word's ListItem.level is 0-based (0..8
    // = 9 indent levels). One source for every Word write path.
    var ls = [], lv = [];
    (lines || []).forEach(function (ln) {
      var s = String(ln == null ? "" : ln);
      var m = /^\s*\[L(\d+)\]\s*/.exec(s);
      var n = 0;
      if (m) {
        var d = parseInt(m[1], 10);
        n = d <= 1 ? 0 : (d >= 9 ? 8 : d - 1);
        s = s.slice(m[0].length);
      }
      ls.push(s); lv.push(n);
    });
    var any = lv.some(function (v) { return v > 0; });
    return { lines: ls, levels: lv, any: any };
  }

  function wordBulletToken(lv) {
    // the built-in list styles carry real list templates; deepest token
    // is ListBullet5, deeper indent levels still work via listItem.level
    var toks = ["listBullet", "listBullet2", "listBullet3", "listBullet4", "listBullet5"];
    return toks[Math.min(lv, 4)];
  }

  function normalizeWordStyle(v) {
    // Word.BuiltInStyleName tokens (localization-safe via styleBuiltIn)
    var k = String(v || "").toLowerCase().replace(/\s+/g, "");
    var map = {
      "title": "title", "subtitle": "subtitle", "normal": "normal", "bodytext": "normal",
      "quote": "quote", "intensequote": "intenseQuote",
      "listbullet": "listBullet", "bullet": "listBullet", "bullets": "listBullet",
      "listnumber": "listNumber", "numbered": "listNumber", "numberedlist": "listNumber"
    };
    var h = k.match(/^heading([1-6])$/) || (k === "h1" ? ["x", "1"] : k === "h2" ? ["x", "2"] : k === "h3" ? ["x", "3"] : k === "h4" ? ["x", "4"] : null);
    if (h) map["heading" + h[1]] = "heading" + h[1];
    return map[k] || null;
  }

  function validateOps(rawOps, ctx) {
    var ops = [], invalid = [];
    // paragraph text mutations the batch itself makes (edit_paragraph /
    // replace_text earlier in the batch change what a later op's "find"
    // matches at apply time - the snapshot only knows the pre-batch text)
    // (accumulator rule: lives at validateOps scope, NOT per-op)
    var paraState = {};
    var paraTextOf = function (p) {
      if (paraState[p] !== undefined) return paraState[p];
      var snap = (ctx && ctx.paragraphs ? ctx.paragraphs[p - 1] : null) || null;
      return (snap && typeof snap.text === "string") ? snap.text : null;
    };
    (rawOps || []).forEach(function (op, index) {
      if (!op || typeof op !== "object") { invalid.push({ op: op, reason: "not an object" }); return; }
      var t = op.type;
      if (OP_TYPES.indexOf(t) < 0) { invalid.push({ op: op, reason: "unknown op type '" + t + "'" }); return; }
      if (!ctx && t !== "insert_paragraphs" && t !== "create_table" && t !== "accept_changes" && t !== "set_track_changes" && t !== "reject_changes" && t !== "set_header" && t !== "set_footer") {
        invalid.push({ op: op, reason: "needs document context (enable 'include document context')" });
        return;
      }
      var bad = function (reason) { invalid.push({ op: op, reason: reason }); };
      var paraOk = function (v) { return typeof v === "number" && Math.floor(v) === v && v >= 1 && (!ctx || v <= ctx.paragraphCount); };
      var out = { type: t };
      if ((t === "insert_image" || t === "insert_paragraphs") && typeof op.place === "string" && op.place.trim()) {
        var p0 = op.place.trim();
        if (p0 === "cursor" || p0 === "end") out.place = p0;
        else if (p0.length >= 3) out.place = p0.slice(0, 200);  // an anchor text to insert after
      }
      if (t === "move_text" || t === "copy_text") {
        if (typeof op.find === "string" && op.find.trim()) out.find = op.find.trim().slice(0, 200);
        else { invalid.push({ op: op, reason: t + " needs find (the text to " + (t === "move_text" ? "relocate" : "duplicate") + ")" }); return; }
        if (typeof op.until === "string" && op.until.trim()) out.until = op.until.trim().slice(0, 200);
        if (typeof op.count === "number" && op.count >= 1 && op.count <= 200) out.count = Math.floor(op.count);
        if (typeof op.place === "string" && op.place.trim() && op.place.trim() !== "end") {
          out.place = op.place.trim() === "cursor" ? "end" : op.place.trim().slice(0, 200);
        }
        if (typeof op.until !== "undefined" && typeof op.count !== "undefined") {
          invalid.push({ op: op, reason: t + ": give until OR count, not both" }); return;
        }
      }
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
        var headingsOnly = itemsOk.length >= 4;
        for (var ih = 0; ih < itemsOk.length; ih++) {
          var stH = itemsOk[ih].style || "normal";
          if (stH === "normal" || stH === "quote" || stH === "listParagraph" ||
              stH === "heading2" || stH === "heading3") { headingsOnly = false; break; }
        }
        if (headingsOnly) { bad("every item uses style title/heading1 - the document BODY must be 'normal' paragraphs; use title for ONE title paragraph and heading1 for section titles only, then continue the text as normal"); return; }
        out.items = itemsOk;
      } else if (t === "move_image") {
        if (typeof op.image === "number" && op.image >= 1) {
          out.image = Math.floor(op.image);
        } else if (typeof op.find === "string" && op.find.trim()) {
          out.find = op.find.trim().slice(0, 200);
        } else { invalid.push({ op: op, reason: "move_image needs either image:N (1-based inline picture from the outline) or find (text of the image paragraph)" }); return; }
        if (typeof op.place === "string" && op.place.trim()) {
          out.place = op.place.trim() === "cursor" ? "end" : op.place.trim().slice(0, 200);
        }
      } else if (t === "edit_paragraph") {
        if (!paraOk(op.para)) { bad("bad paragraph index"); return; }
        if (typeof op.text !== "string") { bad("missing text"); return; }
        out.para = op.para; out.text = op.text;
        paraState[op.para] = op.text;
      } else if (t === "replace_text") {
        if (typeof op.find !== "string" || op.find.length < 2) { bad("find must be at least 2 chars"); return; }
        if (typeof op.replace !== "string") { bad("missing replace"); return; }
        out.find = op.find; out.replace = op.replace; out.all = op.all !== false;
        for (var pi = 1; pi <= 4096; pi++) {
          var curW = paraTextOf(pi);
          if (typeof curW !== "string") continue;
          // document-wide replace: hit every paragraph whose text carries it
          var low = curW.toLowerCase(), f2 = op.find.toLowerCase(), ix;
          if (low.indexOf(f2) < 0) continue;
          var updated = curW;
          while ((ix = updated.toLowerCase().indexOf(f2)) >= 0) {
            updated = updated.slice(0, ix) + op.replace + updated.slice(ix + op.find.length);
          }
          paraState[pi] = updated;
        }
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
        var headOnly2 = items3.length >= 4;
        for (var ih3 = 0; ih3 < items3.length; ih3++) {
          var stH3 = items3[ih3].style || "normal";
          if (stH3 === "normal" || stH3 === "quote" || stH3 === "listParagraph" ||
              stH3 === "heading2" || stH3 === "heading3") { headOnly2 = false; break; }
        }
        if (headOnly2) { bad("every item uses style title/heading1 - the document BODY must be 'normal' paragraphs; title = ONE title paragraph, heading1 = section titles only"); return; }
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
      } else if (t === "set_track_changes") {
        if (typeof op.on !== "boolean") { bad("missing on:true|false"); return; }
        out.on = op.on;
      } else if (t === "set_header" || t === "set_footer") {
        if (typeof op.text !== "string" || !op.text.trim()) { bad("missing header/footer text"); return; }
        out.text = op.text.trim().slice(0, 600);
      } else if (t === "create_toc") {
        if (op.title != null && typeof op.title !== "string") { bad("title must be a string"); return; }
        out.title = typeof op.title === "string" ? op.title.trim().slice(0, 80) : "";
        var lvT = typeof op.levels === "string" && /^\s*1\s*-\s*[23]\s*$/.test(op.levels) ? op.levels.replace(/\s+/g, "") : "1-3";
        out.levels = lvT;
      } else if (t === "add_hyperlink") {
        if (typeof op.find !== "string" || !op.find.trim()) { bad("missing find text"); return; }
        var linkF = op.find.trim();
        if (linkF.length > 200) { bad("find text too long (200 chars max)"); return; }
        if (typeof op.url !== "string" || !/^(https?:\/\/|mailto:)/i.test(op.url.trim())) { bad("url must start with http(s):// or mailto:"); return; }
        out.find = linkF; out.url = op.url.trim();
      } else if (t === "format_selection") {
        var selFont = {}, selAny = false;
        ["bold", "italic", "allCaps", "underline"].forEach(function (kSel) {
          if (op.font && typeof op.font[kSel] === "boolean") { selFont[kSel] = op.font[kSel]; selAny = true; }
        });
        if (op.font && typeof op.font.color === "string" && /^[\w#\(\), .%]+$/.test(op.font.color)) { selFont.color = op.font.color; selAny = true; }
        if (op.font && typeof op.font.highlightColor === "string" && /^[\w#]+$/.test(op.font.highlightColor)) { selFont.highlightColor = op.font.highlightColor; selAny = true; }
        if (op.font && typeof op.font.name === "string" && op.font.name) { selFont.name = op.font.name; selAny = true; }
        if (op.font && typeof op.font.size === "number" && op.font.size >= 6 && op.font.size <= 400) { selFont.size = op.font.size; selAny = true; }
        if (op.font && typeof op.font.align === "string" && ["left", "center", "right", "justify"].indexOf(op.font.align) >= 0) { selFont.align = op.font.align; selAny = true; }
        if (!selAny) { bad("no format changes requested"); return; }
        out.font = selFont;
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
        if (op.font && typeof op.font.size === "number" && op.font.size !== 0) {
          if (op.font.size < 6) { bad("font size must be at least 6pt"); return; }
          if (op.font.size > 28) { bad("font.size " + op.font.size + "pt is far too large for document text - body text stays near the document's default (about 10-14pt; Word's Title style tops out at 28pt). For big section emphasis use Heading styles, not sizes."); return; }
          fontW.size = op.font.size; hasAnyW = true;
        }
        if (typeof op.align === "string" && /^(left|center|right|justify)$/i.test(op.align)) { fontW.align = op.align.toLowerCase(); hasAnyW = true; }
        if (typeof op.style === "string" && op.style.trim()) {
          var stN = normalizeWordStyle(op.style);
          if (!stN) { bad("unsupported style '" + op.style + "' (use Heading 1-3, Title, Subtitle, List Bullet, List Number, Quote or Normal)"); return; }
          out.style = stN;
        }
        if (!hasAnyW && !out.style) { bad("no format changes requested"); return; }
        out.para = op.para; out.font = fontW;
        if (typeof op.find === "string" && op.find.trim()) {
          var findW = op.find.trim().slice(0, 200);
          var paraTxtW = paraTextOf(op.para);   // earlier batch ops may already have rewritten this paragraph
          if (paraTxtW !== null && paraTxtW.toLowerCase().indexOf(findW.toLowerCase()) === -1) { bad("find text not found in that paragraph"); return; }
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
    delete_slide: "del", delete_shape: "del", accept_changes: "edit", set_track_changes: "edit",
    add_hyperlink: "edit", format_selection: "edit", reject_changes: "edit", set_header: "edit", set_footer: "edit"
  };

  function fmtOp(op, ctx) {
    var paraRef = function (i) {
      // op.para is 1-based; guard and read the SAME (0-based: i-1) paragraph
      if (!ctx || !ctx.paragraphs || !ctx.paragraphs[i - 1]) return "";
      var t = ctx.paragraphs[i - 1].text;
      return t ? ' "' + String(t).slice(0, 60) + '"' : "";
    };
    switch (op.type) {
      case "insert_paragraphs": {
        var pl2 = (typeof op.place === "string" && op.place.trim())
          ? { cursor: "caret/selection", end: "end of document" }[op.place.trim()] || ("after '" + op.place.trim().slice(0, 40) + "'")
          : "end of document";
        return "Insert " + op.items.length + " paragraph(s) -> " + pl2 + " (" +
          op.items.map(function (i) { return i.style; }).filter(function (s, k, a) { return a.indexOf(s) === k; }).join(", ") + ")"; }
      case "move_text":
        return "Move " + (op.until || op.count ? (op.until ? "'" + op.until.slice(0, 30) + "'" : String(op.count) + " paragraph(s)") : "1") + // from '" + op.find.slice(0, 40) + "'
          "' -> " +
          (op.place ? "after '" + op.place.slice(0, 40) + "'" : "end of document");
      case "move_image":
        return "Move " + (op.image ? "image " + op.image : "image (near '" + String(op.find || "").slice(0, 40) + "')") + " -> " +
          (op.place ? "after '" + op.place.slice(0, 40) + "'" : "end of document");
      case "copy_text":
        return "Copy " + (op.until || op.count ? (op.until ? "'" + op.until.slice(0, 30) + "'" : String(op.count) + " paragraph(s)") : "1") + // from '" + op.find.slice(0, 40) + "'
          "' -> " +
          (op.place ? "after '" + op.place.slice(0, 40) + "'" : "end of document") + " (original kept)";
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
      case "accept_changes":
        return "Accept ALL tracked changes in the document";
      case "reject_changes":
        return "Reject ALL tracked changes in the document";
      case "set_header":
        return "Page header: " + op.text.slice(0, 60);
      case "set_footer":
        return "Page footer: " + op.text.slice(0, 60);
      case "create_toc":
        return "Table of contents" + (op.title ? ': "' + op.title + '"' : "") + " (heading levels " + (op.levels || "1-3") + ")";
      case "add_hyperlink":
        return "Hyperlink '" + op.find.slice(0, 60) + "' -> " + op.url;
      case "format_selection":
        return "Format ONLY the selected text: " + Object.keys(op.font).join(", ");
      case "set_track_changes":
        return "Turn Track Changes " + (op.on ? "on" : "off") + " in the document";
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
        var pl = (typeof op.place === "string" && op.place.trim())
          ? { cursor: "caret/selection", end: "end of document" }[op.place.trim()] || ("after '" + op.place.trim().slice(0, 40) + "'")
          : "per the pane setting";
        return "Image (" + pl + "): '" + prI + "' (profile image model)";
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
      pre = deck.setTrackChanges(true).then(function (res) {
        SK.revPrevMode = res && res.previous ? String(res.previous) : null;
      }).catch(function (err) {
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
          return deck.insertParagraphs(op).then(function (res) {
            lines.push({ ok: true, text: "insert " + op.items.length + " paragraph(s) -> " +
              ((res && res.how) || "end of document") +
              (op.place && op.place !== "end" ? " (model chose)" : "") });
          });
        case "move_text":
          return deck.moveText(op).then(function (res2) {
            lines.push({ ok: true, text: "moved '" + op.find.slice(0, 50) + "' -> " +
              ((res2 && res2.how) || "after target") });
          });
        case "copy_text":
        case "move_image":
          return deck.moveImage(op).then(function (resM) {
            lines.push({ ok: true, text: "moved " + ((resM && resM.how) || "the image") });
          });
          return deck.copyText(op).then(function (res3) {
            lines.push({ ok: true, text: "copied " + (res3 && res3.count != null ? res3.count + " paragraph(s)" : String(op.find || "").slice(0, 50)) +
              " -> " + ((res3 && res3.how) || "after target") + " (original kept)" });
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
        case "accept_changes":
          return deck.acceptChanges().then(function () {
            lines.push({ ok: true, text: "all tracked changes accepted" });
          });
        case "add_hyperlink":
          return deck.addHyperlink(op.find, op.url).then(function (res) {
            lines.push({ ok: true, text: "hyperlink added on '" + op.find.slice(0, 50) + "' -> " + op.url });
          });
        case "format_selection":
          return deck.formatSelection(op.font).then(function () {
            lines.push({ ok: true, text: "formatted the selected text: " + Object.keys(op.font).join(", ") });
          });
        case "reject_changes":
          return deck.rejectChanges().then(function () {
            lines.push({ ok: true, text: "all tracked changes rejected" });
          });
        case "set_header":
          return deck.setHeaderFooter(op.text, false).then(function () {
            lines.push({ ok: true, text: "page header set: " + op.text.slice(0, 60) });
          });
        case "set_footer":
          return deck.setHeaderFooter(op.text, true).then(function () {
            lines.push({ ok: true, text: "page footer set: " + op.text.slice(0, 60) });
          });
        case "set_track_changes":
          return deck.setTrackChanges(op.on).then(function () {
            lines.push({ ok: true, text: "track changes turned " + (op.on ? "on" : "off") + " in the document" });
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
          return deck.formatText(adj(op.para - 1), op.font, op.style).then(function () {
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
            var opPlace = (typeof op.place === "string" && op.place.trim()) ? op.place.trim() : null;
            var uiAtCursor = $("chat-cursor") && $("chat-cursor").checked;
            var place = opPlace || (uiAtCursor ? "cursor" : "end");
            return deck.addImage({ b64: img.b64, place: place }).then(function (res) {
              lines.push({ ok: true, text: "image inserted " +
                ((res && res.how) || (place === "cursor" ? "at caret/selection" : "at end of document")) +
                (opPlace ? " (model chose)" : " (pane setting)") +
                " (" + imgRef.connection + ":" + imgRef.model + ")" });
            });
          });
        }
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
        // cancel = stop at the very next op boundary
        if (SK.runCancelled) {
          lines.push({ ok: false, text: "cancel - remaining op(s) not executed (" + (op && op.type ? op.type : "unknown") + " skipped)" });
          return Promise.resolve();
        }
        // per-op error capture INSIDE the link: a rejection from the last
        // op must not reject the whole chain (the tail - tracking restore +
        // final status - would never run)
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
        return Promise.resolve();
      });
    });
    return pre.then(function () { return chain; }).then(function () {
      var res = { lines: lines, manualNotes: manualNotes };
      var afterRetries = Promise.resolve();
      if (failedOps.length) {
        lines.push({ ok: true, text: "retrying " + failedOps.length +
          " failed action(s) once after a short pause\u2026" });
        var rseq = new Promise(function (res2) { setTimeout(res2, 1500); });
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
        afterRetries = rseq;
      }
      return afterRetries.then(function () {
        if (revNote) lines.push({ ok: false, text: "revisions: " + revNote });
      // restore whatever the change-tracking mode was before the AI-edits
      // batch (do not stomp a tracking mode the user had enabled themselves)
        var prevTrack = String(SK.revPrevMode || "");
        var wasOn = prevTrack.toLowerCase().indexOf("track") === 0;  // trackAll / TrackMineOnly
        var off = (SK.cfg && SK.cfg.trackedChanges && deck.setTrackChanges)
          ? deck.setTrackChanges(wasOn).catch(function () { return null; })
          : Promise.resolve(null);
        SK.revPrevMode = null;
        return off.then(function () { return res; });
      });
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
    setStatus("Collecting document context…", "busy");
    var ctxPromise = wantCtx ? SK.deck.getContext({}).catch(function (err) {
      setStatus("Could not read the deck (" + err.message + ") - continuing without context.", "error");
      return null;
    }) : Promise.resolve(null);

    ctxPromise.then(function (ctx) {
      SK.lastCtx = ctx;
      SK.webRounds = 0;
      SK.pageReads = 0;
      SK.webResults = "";
      if (!(SK.clarifyAnswer && SK.clarifyAnswer.expect === instruction)) SK.clarifyChain = 0;   // a free new ask resets the clarify chain
      var webPromise = Promise.resolve("");
      return webPromise.then(function () {
      var sys = SYSTEM_PROMPT + "\nHost: " + (SK.sim ? "DocKick simulator" : SK.deck.capabilities().host);
      var _cp = SK.cfg ? (String(SK.cfg.customPrompt || "").trim() ? String(SK.cfg.customPrompt).trim() : String(SK.cfg.customPromptDefault || "").trim()) : "";
      if (_cp) {
        sys += "\n\nUSER CUSTOM INSTRUCTIONS (from Settings - obey these):\n" + _cp;
      }
      sys += "\n" + CLARIFY_RULE;
            var user = (ctx ? formatContext(ctx) : "(no deck context requested)") + "\n\nUSER REQUEST:\n" + instruction;
      if (SK.clarifyAnswer && SK.clarifyAnswer.expect === instruction) {
        user += "\n\nCLARIFICATION: you asked the user: \"" + SK.clarifyAnswer.q + "\". The user chose: " + SK.clarifyAnswer.a +
          ". Deliver the final result for that choice now - do not ask again.";
        SK.clarifyAnswer = null;
        dbg("clarify.answer", { delivered: true });
      }
      if (SK.webResults) user += "\n\nWEB RESULTS (live lookup, cite them when relevant):\n" + SK.webResults;
      var messages = [
        { role: "system", content: sys },
      ].concat(SK.history.slice(-(HISTORY_MAX_EXCHANGES * 2)))
       .concat([{ role: "user", content: user }]);
      SK.lastAsk = { messages: messages, connection: ref.connection, model: ref.model, rounds: 0 };   // 0.9.216: repair rounds can re-ask on validation rejects
      dbg("prompt", { profile: prof.name, route: route, model: ref.model, connection: ref.connection, instruction: instruction, hadContext: !!ctx, messages: messages });
      setStatus("Asking " + ref.model + " (via " + ref.connection + ")…", "busy");
      var askedLookup = false;
      function MO_SYS(extra) {
        return "You build Word-document make-over actions. Reply ONLY with a JSON object " +
          '{\"kind\":\"slide_ops\",\"ops\":[...]} - no prose, no markdown fences. Word op schema: ' +
          '{"type":"format_text","para":N,"font":{"bold":true|false,"italic":true|false,"name":"..","size":N,"color":"#RRGGBB"},"align":"left|center|right|justify"}, ' +
          '{"type":"edit_paragraph","para":N,"text":".."}, {"type":"set_header","text":".."}, {"type":"set_footer","text":".."}, ' +
          '{"type":"add_hyperlink","para":N,"find":"exact text in that paragraph","url":"https://.."}, ' +
          '{"type":"insert_paragraphs","items":[{"text":"..","style":".."}]} (appends at the document end; use sparingly). ' +
          "para = the 1-based paragraph number from the outline. NEVER invent paragraph numbers outside the outline. " + (extra || "");
      }
      function wordMakeOverFlow() {
        var paras = (ctx && ctx.paragraphs) || [];
        var outline = [];
        paras.forEach(function (p) {
          var t = String(p.text || "").replace(/\s+/g, " ").trim();
          if (!t) return;
          outline.push("[" + (p.idx + 1) + "] " + ((p.style || "Normal") === "Normal" ? "" : (p.style + " | ")) + t.slice(0, 90));
        });
        var paraN = paras.length;
        // phantom guard: make-over ops only touch real, 1-based paragraphs
        function validOp(o) {
          if (!o || typeof o !== "object") return false;
          if (o.type === "format_text" || o.type === "edit_paragraph" || o.type === "add_hyperlink")
            return typeof o.para === "number" && o.para >= 1 && o.para <= paraN;
          if (o.type === "insert_paragraphs") return true;
          if (o.type === "set_header" || o.type === "set_footer")
            return typeof o.text === "string" && !!o.text.trim();
          return false;   // make-overs stay in the safe vocabulary
        }
        var moSeen = {};   // dedupe across passes (header/footer repeated by eager builders)
        function mergeFrom(r) {
          var pre = tryParse(r && r.text || "");
          var got = (pre.ok && pre.ops) || (salvageOps(r && r.text || "") || []);
          return got.filter(function (o) {
            if (!validOp(o)) return false;
            var sig = null;
            if (o.type === "set_header" || o.type === "set_footer") sig = o.type + "|" + String(o.text || "");
            if (o.type === "insert_paragraphs") sig = o.type + "|append";
            if (sig) {
              if (moSeen[sig]) return false;      // a later pass repeats an earlier extra - keep the first
              moSeen[sig] = true;
            }
            return true;
          });
        }
        var ops0FromDirector = null;
        var askBrief = "ORIGINAL USER REQUEST:\n" + String(instruction || "").slice(0, 400) +
          "\nDOCUMENT OUTLINE (1-based paragraph numbers; op.para uses these):\n" + outline.slice(0, 220).join("\n");
        var DIR_SYS = "You are the creative director for Word document make-overs. Invent a REAL design strategy - not font tweaks: " +
          "give the document an editorial identity. Reply ONLY with the strategy in short sections:\n" +
          "CONCEPT: the design idea with a memorable name (one paragraph)\n" +
          "THEME: heading font + body font (real font names), exact sizes for H1/H2/lead-in/body, 2-3 colors + 1 accent as #RRGGBB\n" +
          "STRUCTURE: how headings vs lead-ins vs body vs quotes are visually shaped (weights, allCaps, italics, alignment, spacing feel)\n" +
          "EXTRAS: exact page header text, table/link treatment if the document has any, and whether a pull-quote/summary block adds value";
        function passAsk(passName, userMsg) {
          if (SK.runCancelled) return Promise.reject(new FriendlyError("Run cancelled."));
          return askRetryable(function () {
            return apiJson("/api/chat", "POST", {
              connection: ref.connection, model: ref.model,
              temperature: 0.5, maxTokens: 2600,
              messages: [{ role: "system", content: MO_SYS("This round: " + passName + ".") },
                         { role: "user", content: userMsg }],
            });
          }).then(function (r) {
            dbg("makeover.pass", { pass: passName, len: String(r.text || "").length, elapsedMs: r.elapsedMs });
            return r;
          });
        }
        setStatus("AI director is inventing the document make-over…", "busy");
        return passAsk("DIRECTOR", "Design a make-over strategy for this document.\n\n" + askBrief).then(function (dr) {
          var concept = String(dr.text || "").slice(0, 2400);
          var preDr = tryParse(dr && dr.text || "");
          if (preDr.ok && preDr.ops && preDr.ops.length) {
            // disobedient director: answered with ops instead of a strategy -
            // keep its ops as the first results and give the builders a generic
            // editorial fallback strategy so their prompts still carry a plan
            ops0FromDirector = preDr.ops.filter(validOp);
            concept = "CONCEPT: Editorial refresh - one clean typography system. "
              + "THEME: system heading font on headings, regular body font, body 11-12pt, " +
              "headings one step larger; 1 accent color used sparingly. "
              + "STRUCTURE: headings bold and larger, lead-ins bold, quotes italic. "
              + "EXTRAS: a modest page header, no footer duplication.";
          }
          setStatus("Builder pass 1/3 - structure…", "busy");
          var b1 = "BUILDER PASS 1 - STRUCTURE. Follow the strategy below and emit ONLY ops that give the document a real hierarchy: " +
            "format_text (weights/italic/name + align) for headings, lead-ins and quotes; edit_paragraph only when a heading text must be tightened. " +
            "Use the exact 1-based paragraph numbers from the outline. Never repeat an op for the same paragraph.\n\n" +
            "DIRECTOR STRATEGY:\n" + concept + "\n\n" + askBrief;
          var ops = (typeof ops0FromDirector !== "undefined" ? (ops0FromDirector || []) : []);
          ops0FromDirector = null;
          if (ops.length) dbg("makeover.directorOps", { count: ops.length });
          return passAsk("BUILD1-STRUCTURE", b1).then(function (r1) {
            ops = ops.concat(mergeFrom(r1));
            setStatus("Builder pass 2/3 - typography & colors…", "busy");
            var b2 = "BUILDER PASS 2 - STYLE. ONLY format_text ops with the EXACT theme values from the strategy " +
              "(font name/size/color for every body paragraph, lead-ins and quotes). Cover EVERY content paragraph of the outline " +
              "that the ops list below does not already handle. Use the exact 1-based numbers.\n\n" +
              "OPS ALREADY COVERED:\n" + JSON.stringify(ops) + "\n\nDIRECTOR STRATEGY:\n" + concept + "\n\n" + askBrief;
            return passAsk("BUILD2-STYLE", b2).then(function (r2) {
              ops = ops.concat(mergeFrom(r2));
              setStatus("Builder pass 3/3 - extras…", "busy");
              var b3 = "BUILDER PASS 3 - EXTRAS. From the strategy: set_header with the exact header text (and set_footer if the strategy wants one), " +
                "add_hyperlink for real links, insert_paragraphs ONLY for a pull-quote/summary block the strategy explicitly asks for. " +
                'If none apply, reply {"kind":"slide_ops","ops":[]}.\n\nDIRECTOR STRATEGY:\n' + concept + "\n\n" + askBrief;
              return passAsk("BUILD3-EXTRAS", b3).then(function (r3) {
                ops = ops.concat(mergeFrom(r3));
                setStatus("Reviewing make-over coverage…", "busy");
                var rev = "REVIEW ROUND. Every content paragraph in the outline must be covered by a format_text op in the ops list below " +
                  "per the strategy. Emit ONLY the missing format_text ops (exact 1-based numbers). " +
                  'If the coverage is complete, reply {"kind":"slide_ops","ops":[]}.\n\n' +
                  "OPS SO FAR:\n" + JSON.stringify(ops) + "\n\nDIRECTOR STRATEGY:\n" + concept + "\n\nOUTLINE:\n" + outline.slice(0, 220).join("\n");
                return passAsk("REVIEW", rev).then(function (r4) {
                  ops = ops.concat(mergeFrom(r4));
                  dbg("makeover.merged", { ops: ops.length });
                  return { text: JSON.stringify({ kind: "slide_ops", ops: ops,
                           note: "AI document make-over: " + ops.length + " action(s) built in passes - structure -> style -> extras, then review rounds" }) };
                });
              });
            });
          });
        });
      }
      // 0.9.186: content-creation requests ALWAYS run the web lookup first
      // (standing rule: the pane owns the lookup channel - internal, free);
      // fetched results feed the writing ask; the 3-lookup budget still caps it.
      var CONTENT_WRITE_RE = /\b(write|draft|compose|create|generate|produce|make)\b[^\n]{0,90}?\b(essay|article|report|letter|proposal|story|summaries|summary|overview|introduction|conclusion|section|paragraph|paper|guide|description|blog|post|plan|bio|announcement|review|analysis|document|manual|notes|note|memo|whitepaper|white paper|case study|press release|abstract|bulletin|statement|brief)\b/i;
      var preLookupDone = null;
      function preLookupRun() {
        if (preLookupDone === null) {
          if (CONTENT_WRITE_RE.test(instruction || "")) {
            SK.webRounds = (SK.webRounds || 0) + 1;   // the shared 3-lookup budget
            setStatus("Checking the web for current info to write from\u2026", "busy");
            preLookupDone = webLookup(prof, instruction).then(function (ws) {
              if (!ws) { SK.webRounds = Math.max(0, (SK.webRounds || 1) - 1); return null; }
              SK.webResults = ws;
              setStatus("Reading the found documents\u2026", "busy");
              return gatherPageContents(ws).then(function (pages) {
                SK.webResults = ws + pages;
                messages = messages.concat([{ role: "user", content:
                  "WEB LOOKUP RESULTS (the pane fetched these proactively before this content-writing request - verify facts against them, keep them current, cite what you use; do NOT invent numbers/dates):\n" + ws +
                  (pages ? "\n\nThe lookups also RETURNED THE FULL TEXT of matched pages - read before writing:\n" + pages : "") }]);
                return true;
              });
            }).catch(function (ePl) {
              dbg("prelookup.failed", { msg: String(ePl && ePl.message || ePl) });
              SK.webRounds = Math.max(0, (SK.webRounds || 1) - 1);
              return null;
            });
          } else {
            preLookupDone = Promise.resolve(null);
          }
        }
        return preLookupDone;
      }
      function mainAsk() {
        return preLookupRun().then(function () { return mainAskInner(); });
      }
      function mainAskInner() {
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
      // ------------------------------------------------------------------
      // FULL DOCUMENT make-over: AI-driven strategy (like SlideKick's
      // orchestrated redesign). A creative-director round invents a design
      // strategy for THIS document, three builder passes translate it into
      // Word ops (structure -> style -> extras), a review round tops up
      // anything missed, then the usual validation + gate take over.
      // ------------------------------------------------------------------
      var WMAKEOVER_RE = /\b(make\s*over|makeover|redesign|re-?style|restyle|new (design|theme|look)|full design|design (the|this) (document|doc|word|whole)|modern|professional|fresh look|creative)\b/i;
      var wmoWanted = WMAKEOVER_RE.test(instruction || "") &&
        !!(ctx && ctx.paragraphs && ctx.paragraphs.length > 1);
      if (wmoWanted) {
        return wordMakeOverFlow().then(function (mr) { return mr; }, function (eMo) {
          dbg("makeover.failed", { msg: String(eMo && eMo.message || eMo) });
          return mainAsk().then(function (resp2) { return resp2; });
        });
      }
      return mainAsk().then(function (resp) {
        dbg("response", { text: resp.text, elapsedMs: resp.elapsedMs });
        SK.lastModelText = { text: resp.text };
        var parsed = tryParse(resp.text);
        if (parsed.ok && parsed.kind === "web_lookup" && parsed.ops) {
          // look-up-then-act in one reply: run the lookup (internal, free),
          // then the model is re-asked normally - but if THAT fails to parse,
          // still apply the ops the model attached to the intent.
          SK.pendingWebOps = parsed.ops;
          return webLookupThen(resp);
        }
        if (parsed.ok && parsed.kind === "web_lookup") {
          // bare lookup intent + nothing else: the pane has probably ALREADY
          // handled it mid-run (refusal trigger); if no results were
          // collected, run it now and do NOT end with an empty-action list.
          if (!SK.webResults) return webLookupThen(resp);
          return parsed;   // web results exist; fall through (empty ops => the
                           // "empty action list" box below now explains why)
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

        if (parsed.ok) return parsed;
        // escalate: up to 2 repair rounds; empty replies get the loudest nudge;
        // a cut-off op list is salvaged and the remaining ops asked for once
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
        if (parsed && parsed.kind !== "clarify" && !(SK.clarifyChain || 0) &&
            /infographics?\b/i.test(instruction) &&
            !/\b(image|picture|illustration|screenshot|photo|poster|graphic)\b/i.test(instruction)) {
          // 0.9.236: the word infographic ALWAYS surfaces the form question
          parsed = { ok: true, kind: "clarify",
                     question: "Infographic as an AI-generated image, a table/diagram, or formatted text?",
                     options: ["AI image", "Table or diagram", "Formatted text"] };
        }
        if (parsed && parsed.kind === "clarify") {
          // 0.9.235: the model asks the user to pick the form of the result
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
        // 0.9.235: the model asks the user to pick the form of the result;
        // internal action, never gated - the pane renders the question and
        // the re-ask carries the user's decision
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
  // form of the result instead of guessing; the pick re-asks internally.
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
      box2.textContent = parsed.text;
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
      // 0.9.224: a silent "Done" here hides what the model actually said -
      // show its reply text (why it chose to do nothing) + a one-click re-ask.
      var box3 = document.createElement("div");
      box3.className = "reply-text";
      var said = "";
      try { said = String((SK.lastModelText && SK.lastModelText.text) || ""); } catch (eM) { }
      box3.textContent = said ? ("The model proposed 0 actions. It answered:\n" + said.slice(0, 600))
                              : "(model returned an empty action list)";
      out.appendChild(box3);
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
      out.appendChild(reRow);
      if (SK.lastAsk) { SK.lastAsk.rounds = 0; }   // the re-request gets fresh repair rounds too
      setStatus("Done - the model proposed 0 actions.", "ok");
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

    if (!validation.ops.length && validation.invalid.length > 0) {
      // 0.9.216 - DON'T reject fast: when nothing survives validation, show
      // the reasons AND re-ask the model with them (up to 2 repair rounds),
      // only then fall back to the manual "Ask again" box.
      var reasons = (validation.invalid || []).map(function (entry) {
        var ot = (entry.op && entry.op.type) || "?";
        var plc = (entry.op && entry.op.place) ? " place='" + String(entry.op.place).slice(0, 40) + "'" : "";
        return "  - " + ot + plc + ": " + (entry.reason || "unknown");
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
            "For insert_paragraphs/move_text/copy_text use a 'place' that is an EXACT text snippet from the outline (or \"end\"/\"cursor\"); " +
            "for para-referencing ops use only outline numbers; for insert_paragraphs keep styles normal/heading1/heading2/quote/listParagraph." },
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
      allbad.textContent = "None of the " + validation.invalid.length + " action(s) survived validation" +
        (SK.lastRecovered ? " (they came from a cut-off reply)" : "") +
        " - no changes were applied. Reasons:";   // 0.9.215: the ACTUAL per-op reason(s) instead of a guess
      box.appendChild(allbad);
      validation.invalid.forEach(function (entry) {
        var r = document.createElement("div");
        r.className = "opnote";
        var ot = (entry.op && entry.op.type) || "?";
        var place = (entry.op && entry.op.place) ? " place='" + String(entry.op.place).slice(0, 40) + "'" : "";
        r.textContent = "  - " + ot + place + ": " + (entry.reason || "unknown");
        box.appendChild(r);
      });
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
      setStatus("Rejected - see the reasons above", "error");   // 0.9.215: the spinner DID keep running after this path
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
        if (!sid2) { try { sid2 = localStorage.getItem("wk-resume-wordkick"); } catch (eE1) {} }
        localStorage.setItem("sk-approvals-wordkick", JSON.stringify({ sid: sid2 || "", map: SK.sessionApprovals }));
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
      ["Pane build (JS)", window.__SK_BUILD || "?"],
      ["Add-in identity (this manifest)", (SK.cfg && SK.cfg.manifests && SK.cfg.manifests["dockick"]) ? "v" + SK.cfg.manifests["dockick"] : "?"],
      ["All manifest identities", (SK.cfg && SK.cfg.manifests) ? Object.keys(SK.cfg.manifests).map(function (k5) { return k5 + " v" + SK.cfg.manifests[k5]; }).join(" / ") : "?"],
      ["Office host", (function () { try { return Office.context.diagnostics.host + " " + Office.context.diagnostics.version; } catch (e5) { return "n/a"; } })()],
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
      $("statusbar").textContent = "DocKick v" + (cfg.version || "0.1.0") +
        (window.location && window.location.protocol === "https:" ? " - \u{1F512} " : " - ") +
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
    $("chat-input").addEventListener("keydown", function (ev) {
      if ((ev.metaKey || ev.ctrlKey) && ev.key === "Enter") { ev.preventDefault(); runPrompt(); }
    });
    $("chat-new").addEventListener("click", function () { startNewSession($("chat-new")); });
    $("chat-refresh-deck").addEventListener("click", function () { updateDeckSummary(); setStatus("Document refreshed.", ""); });

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
    try { bootResume = localStorage.getItem("sk-resume-session-wordkick"); } catch (eS) {}
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
    $("settings-tc-save").addEventListener("click", function () {
      apiJson("/api/settings", "PUT", { trackedChanges: $("settings-trackedchanges").checked }).then(function () {
        SK.cfg.trackedChanges = $("settings-trackedchanges").checked;
        setFormMsg($("settings-tc-msg"), "Revision setting saved.", true);
      }).catch(function (err) { setFormMsg($("settings-tc-msg"), err.message); });
    });
    updateSessionLabel();
    wireSessions();
    wireSearchProfiles();
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
      el.textContent = "DocKick v" + ver +
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
    setStatus("Cannot reach the DocKick server: " + err.message + " Is it running? (python3 install.py)", "error");
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