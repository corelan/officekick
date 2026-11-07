/*
 * SlideKick task pane (runs inside PowerPoint via Office.js, or in the
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
    app: "slidekick",
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
    lastRecovered: 0,      // op count salvaged from a cut-off reply (gate note)
    clarifyAnswer: null,   // 0.9.235: pending user pick from a clarify round
    clarifyChain: 0,       // clarify rounds served since the original ask
    sessionManual: false,  // the user renamed this session: pause AI titles
    recoveredDupes: 0,     // of those, duplicates removed during the merge
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

  var LAYOUT_KINDS = ["title_content", "title_only", "section", "blank"];

  var OPS_META = {
    create_slide: { label: "create slide", hint: "Appends a new slide (Office.js can only append; drag into place if needed)." },
    edit_slide: { label: "edit slide", hint: "Replaces title / body / speaker notes of an existing slide." },
    replace_text: { label: "replace text", hint: "Rewrites the text of one shape in place; with a find fragment, only that fragment is replaced (rest of the list untouched)." },
    set_notes: { label: "set notes", hint: "Writes speaker notes on a slide." },
    create_table: { label: "create table", hint: "Adds a table (new slide when slide is empty; position/size in points)." },
    move_shape: { label: "move/resize shape", hint: "Sets left/top/width/height (points) and optional rotation (degrees) of one shape - design passes use this." },
    copy_shape:        { label: "copy shape", hint: "Duplicates a shape (picture, text box, table...) onto another slide or the same one, keeping full fidelity; optional left/top re-placement. PowerPointApi 1.13+" },
    move_shape_to:     { label: "move shape to", hint: "CUT+paste: relocates a shape to another slide (position kept or re-specified); the original is removed. PowerPointApi 1.13+" },
    create_image: { label: "generate + insert image", hint: "Creates an image from a description with the profile's image model (OpenAI-style images API) and inserts it - lands on the slide you have open in Office." },
    apply_layout: { label: "apply layout", hint: "Applies a different slide layout (design change, text kept)." },
    format_text: { label: "format text", hint: "bold / italic / color / size / underline / allCaps / alignment / bullets on one shape's text (the numbers from the outline)." },
    format_selection: { label: "format selected text", hint: "the same font changes applied to the TEXT YOU SELECTED right now." },
    create_shape: { label: "insert shape", hint: "textbox (with text) or geometric shapes (rectangle/ellipse/roundedRectangle/hexagon/diamond/triangle) and straight lines at a position; optional rotation (degrees)." },
    move_slide: { label: "move slide", hint: "moves a slide so it becomes slide M (numbers as in the outline); runs alone in a batch." },
    duplicate_slide: { label: "duplicate slide", hint: "exact copy of one slide, inserted right after it (runs alone in a batch)." },
    duplicate_slide_to: { label: "copy slide over", hint: "copies one slide's full contents over another slide (target replaced); runs alone." },
    clear_slide: { label: "clear slide", hint: "empties one slide: clears all texts; the slide, layout and background stay." },
    set_background: { label: "slide background", hint: "AI-generated picture or a solid color as the slide background (optional transparency)." },
    set_shape_image: { label: "picture-fill shape", hint: "Fills one existing shape with an AI-generated picture (optional transparency) - images inside panels, circles, cut-outs." },
    add_hyperlink: { label: "add hyperlink", hint: "links one shape (from the outline) to a URL." },
    style_shape:  { label: "style shape", hint: "solid fill and/or outline color and optional transparency for one shape." },
    set_alt_text: { label: "alt text", hint: "sets the accessibility alt text of one shape." },
    delete_slide: { label: "delete slide", hint: "Removes a slide from the deck." },
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


  // ---- image generation as an async job: short 4s polls survive WebView
  // suspension, a killed fetch costs one cycle instead of the whole wait ----
  function imageJob(payload) {
    var deadline = Date.now() + 300000;   // same 5-minute budget as before
    function poll(jobId) {
      if (SK.runCancelled) throw new FriendlyError("Run cancelled.");
      if (Date.now() > deadline) {
        throw new FriendlyError("image generation timed out after 300 s (job " + jobId + ")");
      }
      return apiJson("/api/image/status/" + encodeURIComponent(jobId), "GET", undefined)
        .catch(function (err) {   // transport hiccup: keep polling - never give up early
          return { ok: true, status: "running" };
        })
        .then(function (st) {
          if (st && st.status === "done") return { ok: true, b64: st.b64 };
          if (st && st.status === "error") throw new FriendlyError("image generation failed: " + (st.error || "unknown provider error"));
          return new Promise(function (res) { setTimeout(res, 4000); }).then(function () { return poll(jobId); });
        });
    }
    return apiJson("/api/image/start", "POST", payload).then(function (j) {
      if (!j || !j.jobId) throw new FriendlyError("server did not start the image job");
      return poll(j.jobId);
    });
  }
  SK.imageJob = imageJob;
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
    // scan for balanced { ... } objects (string-aware); try each one -
    // the first braces may wrap chatter, the ops JSON can sit mid-text.
    var found = [];
    var n2 = t.length;
    for (var c2 = 0; c2 < n2; c2++) {
      if (t.charAt(c2) !== "{") continue;
      var depth = 0, inStr = false, escNext = false, end = -1;
      for (var k = c2; k < n2; k++) {
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
          if (depth === 0) { end = k; break; }
        }
      }
      if (end > c2) {
        var cut = t.slice(c2, end + 1);
        var o2 = undefined;
        try { o2 = JSON.parse(cut); } catch (e) { /* not JSON, keep scanning */ }
        if (o2 && typeof o2 === "object") {
          if (o2.kind === "slide_ops" && Array.isArray(o2.ops)) return o2;
          if (Array.isArray(o2.ops)) return o2;   // tolerate missing kind
          if (o2.kind === "reply" || o2.kind === "clarify" || o2.kind === "read_slide" || o2.kind === "web_lookup") return o2;
          found.push(o2);
        }
        c2 = end; // skip past this object
      }
    }
    return found.length ? found[0] : undefined;
  }

  // MAKE-OVER MODE: the creative director brief. Injected into the system
  // prompt when the instruction asks for a design make-over. Models default
  // to low-effort font tweaks; this brief forces a theme, original layouts
  // per role, image treatments (backgrounds, cut-out shapes, overlays) and
  // op-count minimums so the result reads as a redesign.
  var MAKEOVER_RE = /\b(make\s*over|makeover|reskin|restyle|redesign|re-?style|new (design|theme|look)|design (the|this) (deck|whole)|modern|professional|clean up the design|fresh look)\b/i;
  var MAKEOVER_BRIEF = [
    "",
    "MAKE-OVER MODE - invent a DESIGN, do not decorate.",
    "You are a first-rate presentation designer. The user asks for a redesign",
    "with no other instructions; they expect a confident, surprising result.",
    "Reuse of a previous direction is a failure - every redesign starts fresh.",
    "",
    "STEP 1 - INVENT THE CONCEPT (never a default look):",
    "Choose a named design direction (e.g. 'ink-and-paper editorial', 'bauhaus",
    "blocks', 'duotone gallery', 'mono brutalist', 'soft gradient tech'...)",
    "and write its three rules before touching anything:",
    "  r1 COLOR: deep/panel/accent hex trio + the matching text colors.",
    "  r2 SPACE: how content sits on the canvas (full-bleed vs large margins,",
    "     asymmetric splits, oversized numbers, vertical rhythm...).",
    "  r3 SHAPE + IMAGE LANGUAGE: the shapes that carry the identity (circle",
    "     badges, slim bars, huge half-panels, diagonal blocks...) and the one",
    "     photo style phrase for every generated image.",
    "If this session's history shows an earlier redesign, deliberately pick a",
    "CONTRASTING direction (different temperature, rhythm, image treatment).",
    "",
    "STEP 2 - STRUCTURE BEFORE STYLE: judge the deck's flow and reshape it:",
    "- no real cover? add one (create_slide layout:title_only) and design it.",
    "- long deck with distinct topics? insert topic-divider slides",
    "  (create_slide layout:section) where they are missing.",
    "- an old slide is only a placeholder? clear_slide it and rebuild, or",
    "  duplicate_slide_to a good slide and restyle the copy.",
    "- then move_slide the flow into shape (run it alone in its own reply).",
    "",
    "STEP 3 - COMPOSE each slide freely with the op vocabulary (backgrounds,",
    "panels, bars, overlays, textboxes, shapes, picture-filled shapes,",
    "formatting). No fixed recipes: the concept decides what a cover looks",
    "like, how a topic divider reads, and where photos go. Put photos INSIDE",
    "shapes when the concept wants cut-outs (circle badge, rounded panel,",
    "oversized half-crop); keep body text clear of busy photo areas (overlay",
    "or empty field). Every in-scope slide takes at least one structural op",
    "(format-only tweaks on a slide = that slide is not redesigned), and",
    "composition moves (move_shape) on most slides. Use images only when the",
    "concept calls for them - a photographic concept without generated",
    "images breaks the concept; a graphic concept may stand on shapes alone.",
    "",
    "USE THE FULL TOOLBOX: layering via op order (the last op paints on top), transparency stacking for tints and glows, rotation for diagonal compositions, gradient backgrounds for color washes, oversized numbers or keywords as graphic elements, thin accent bars, photo cut-outs inside circles, rounded or hexagon shapes - combinations of these are what make a design feel crafted.",
    "- NEVER destroy the user's content (pictures, charts, tables). Only elements you created in this redesign session may be removed.",
    "QUALITY BAR - READABILITY FIRST, then consistency: every text gets its",
    "color from the surface behind it (dark text on light fields, light text",
    "only on overlays/deep/panel colors with strong contrast; never",
    "mid-tone-on-mid-tone). Sizes: titles 28-44pt, body 16-20pt, captions",
    ">= 14pt, nothing under 12pt; keep the title/body hierarchy obvious.",
    "No body text directly on busy photos (overlay or empty field).",
    "Re-check every text you styled for contrast before finishing the slide.",
    "Consistent geometry for repeated slide kinds (covers look alike,",
    "dividers look alike); ops ordered: (1) structure ops,",
    "(2) backgrounds/panels/shapes, (3) moves + text styling, (4) image",
    "generation ops. Keep the user's wording unless they asked for copy",
  ].join("\n");

  var SYSTEM_PROMPT = [
    "You are SlideKick, an AI assistant inside a PowerPoint task pane.",
    "You get the current deck outline and a user request. Reply with ONLY one",
    "JSON object - never any commentary, never code fences, never reasoning or",
    "chain-of-thought. Format A (conversational):",
    '{"kind":"reply","text":"your answer"}',
    "Format B (deck changes):",
    '{"kind":"slide_ops","ops":[ ... ]}',
    "",
    "Allowed op types:",
    '- {"type":"create_slide","layout":"title_content|title_only|section|blank","title":"...","body":["bullet",...],"notes":"...","after":N} (body/notes optional; "after" optional: the 1-based slide number to place the NEW slide after - after:0 = at the very start. The pane moves the new slide into that position automatically)',
    '- {"type":"edit_slide","slide":N,"title":"..."|null,"body":[...]|null,"notes":"..."|null} (null means: leave unchanged)',
    '- {"type":"replace_text","slide":N,"shape":N,"text":"..."} (reword one shape; N refers to the shape numbers listed for that slide). PREFERRED for a bullet/list item edit: add "find":"<the exact current text of that one bullet/fragment from the outline>" and make "text" the REPLACEMENT for only that fragment - the rest of the shape\'s text is preserved. Without "find" the op replaces the shape\'s WHOLE text, so only do that when the outline shows the entire text (no …[text continues]) or you really want to rewrite it all',
    '- {"type":"set_notes","slide":N,"notes":"..."}',
    '- {"type":"create_table","slide":N|null,"values":[["h1","h2"],["x","y"]],"left":40,"top":150,"width":640} (all cells strings, rows rectangular; slide null = new title slide first; left/top/width optional points)',
    '- {"type":"move_shape","slide":N,"shape":N,"left":40,"top":150,"width":300,"height":60} (points; only the fields to change; keep shapes inside the slide bounds shown in the outline)',
    '- {"type":"copy_shape","slide":N,"shape":N,"to":M,"left":40,"top":150} (duplicate the shape onto slide M; "to" omitted = the same slide; left/top optional re-placement; full fidelity copy)',
    '- {"type":"move_shape_to","slide":N,"shape":N,"to":M,"left":40,"top":150} (CUT: relocate that shape to slide M - original removed)',
    '- {"type":"create_image","prompt":"detailed image description","slide":N|null,"width":400} (generate an image from the description and insert it; slide null = the slide the user currently has open; only when the user asks for a picture/image/illustration/logo)',
    '- {"type":"apply_layout","slide":N,"layout":"title_content|title_only|section|blank"} (only when the user asks for layout/design changes)',
    '- {"type":"format_text","slide":N,"shape":N,"font":{"bold":true},"align":"center","bullet":true} (format ONE shape from the outline: font = {"bold","italic","color","size","name","underline","allCaps"} subset only, colors as HTML hex/red/green/...; align = left|center|right; bullet = true adds bullets; only requested properties)',
    '- format_text needs a TEXT shape (textbox or placeholder): tables, charts and pictures have no formattable text range - for formatted prose next to a table, add a textbox (create_shape) and format that',
    '- When a bullet/slide contains code, commands, protocol fragments or syntax (not ordinary prose), make a separate format_text op setting font {"name":"Consolas"} (Courier New acceptable) on that shape so the technical text is monospaced; leave prose as-is.',
    '- Office.js cannot CREATE SmartArt, stock icons or 3D models. If the user asks for those, say so and offer shape/textbox-based alternatives (create_shape) or an AI-generated picture (create_image) instead. The outline marks existing SmartArt/model3D shapes so you can reference them.',
    '- {"type":"create_shape","slide":N,"kind":"textbox|rectangle|ellipse|roundedRectangle|hexagon|diamond|triangle|line","text":"only for textbox","left":120,"top":200,"width":240,"height":80} (points, slide-local; line = start point left/top + width/height = endpoint offsets; other shapes: no text; optional "fill":"#RRGGBB", "line":"#RRGGBB", "transparency":0.25 to style the created shape immediately)',
    '- to format only PART of a shape\'s text, add "find":"exact words appearing in that shape\'s text" (case-insensitive substring; e.g. {"type":"format_text","slide":N,"shape":N,"find":"security","font":{"bold":true}})',,
    '- {"type":"format_selection","font":{"bold":true}} (format ONLY the text the user has selected right now; same font subset; no slide/shape fields)',
    '- {"type":"set_background","slide":N,"kind":"image","prompt":"description","transparency":0.2} (an AI-generated picture as the slide background; transparency 0-1 optional) or {"type":"set_background","slide":N,"kind":"solid","color":"#RRGGBB","transparency":0.2} (solid color; transparency optional)',
    '- {"type":"add_hyperlink","slide":N,"shape":N,"url":"https://...","tip":"tooltip text"} (link the shape to the URL; url must be http(s) or mailto:)',
    '- {"type":"style_shape","slide":N,"shape":N,"fill":"#RRGGBB","line":"#RRGGBB","transparency":0.3} (solid fill and/or outline color and optional transparency 0-1 for ONE shape; colors as HTML hex or named; at least one of fill/line/transparency)',
    '- {"type":"set_alt_text","slide":N,"shape":N,"description":"alt text","title":"optional title"} (accessibility alt text for the shape)',
    '- {"type":"set_shape_image","slide":N,"shape":N,"prompt":"detailed image description","transparency":0.3} (picture-fill ONE existing shape (panel/circle/rectangle) with an AI-generated picture; transparency optional; use for cut-out photos, side panels, badges)',
    '- {"type":"move_slide","slide":N,"to":M} (move slide N so it becomes slide M; numbers as in the outline 1..N; run it ALONE in its own reply - do not combine with other ops. PowerPoint cannot place a brand-new slide at a position: create it first (it lands last) and the pane moves it into place automatically)',
    '- {"type":"duplicate_slide","slide":N} (exact copy of slide N, inserted right after it - deck grows by one; run it ALONE in its reply; the copy keeps the original\'s design)',
    '- {"type":"duplicate_slide_to","slide":N,"to":M} (copy slide N fully OVER slide M: M keeps its position but now shows N\'s complete contents; run it ALONE; this is how a slide gets inserted at a spot you like when no empty placeholder slide exists - duplicate the layout example then move it)',
    '- {"type":"clear_slide","slide":N} (empty slide N: clears every text frame on it; the slide, background and layout stay; combine with create_shape/move_shape/create_image in the SAME reply to rebuild it as a new design)',
    '- {"type":"delete_slide","slide":N}',
    "",
    "Rules:",
    "- Slide numbers are 1-based (slide 1 = the first slide in the filmstrip) and MUST come from the outline below. When the user says \"slide 2\", use \"slide\":2. Never invent them. Shape indexes keep the numbers shown in the outline.",
    "- One body element = one bullet/paragraph; keep bullets short unless told otherwise.",
    "- When a shape's text ends with \"…[text continues]\" the outline shows only a preview: NEVER rebuild the whole shape from a preview (you would erase the hidden part of the list). Edit list items with replace_text + \"find\" (fragment replace) instead.",
    "- Reword/tighten requests: only rewrite wording, keep the meaning and slide count.",
    "- LARGE DECKS: the outline lists every slide. If you cannot address all of them in one reply, work in batches of at most 20 ops and end each reply with a short note of exactly what remains.",
    "- If the request is a question/explanation/brainstorm rather than a deck change, use format A.",
    "- Design passes ('design this deck/slide', improve alignment/balance): emit ops (apply_layout, move_shape, create_table) computed from the geometry in the outline; when the user asks only for advice, use format A.",
    "- MAKE-OVERS ('give this deck a design makeover', 'redesign slides 3-8', 'modern, original, professional'): the outline marks each slide with a ROLE (title = cover/section slide, bullets, statement, table, image). The make-over is a THEME, not font tweaks: build ONE design system, then give EVERY SLIDE OF THE SAME ROLE the SAME layout blueprint; different roles use different blueprints. This makes the deck consistent where it is the same kind of slide and varied where it is not - all title slides look like siblings; title vs content slides differ.",
    "- Make-over design system, fixed for the whole scope: 2-3 colors + 1 accent (hex), a title size and a body size, one spacing grid (points; slides ~960x540), consistent title/body anchor positions per role. Decide the theme BEFORE emitting ops and apply it across every slide you touch in one reply.",
    "- Role blueprints (pick per theme, keep within the slide bounds): 'title/section' slides -> ONE shared treatment, e.g. full-bleed background photo (set_background kind:image) with a dark overlay shape (create_shape rectangle, fill #000000, transparency 0.45-0.6) and the title large in light color, same anchored position on every title slide. 'bullets' -> e.g. an accent bar (thin rectangle) beside the title, body as text boxes on a clean field, OR a side panel; same side of the slide for the whole family. 'statement' -> fewer/larger elements, generous whitespace. 'table' -> keep the table, restyle surroundings (header bar shape behind the title row, soft panel behind the table). 'image' slides -> the picture as a cut panel: picture-fill an existing shape (set_shape_image) or crop-position the photo on one side; never float the photo randomly.",
    "- NEVER give two slides with the same role different geometry patterns (that is the consistency failure). Never copy one slide's complete layout onto a different-role slide either.",
    "- Make-over building blocks: set_background (AI picture or solid color, optional transparency); create_shape with kind + geometry + fill/line/transparency for panels, accent bars, color blocks and overlays; style_shape with fill/line/transparency to restyle anything already on the slide; move_shape to reposition/resize (+rotation, degrees) or rotate anything you created or moved (rotation is how diagonal/tilted compositions are made); replace_text/format_text for the words; set_shape_image to fill an EXISTING shape (panel/circle) with an AI-generated picture. Shape kinds: textbox, rectangle, circle/ellipse, chevron, rightArrow, downArrow, upArrow, leftArrow, donut, plus, hexagon, star5, star8, teardrop, parallelogram, pentagon, octagon, cube, can, moon, lightningBolt, sun, heart, wave, doubleWave, arc, blockArc, pie, foldedCorner, halfFrame, smileyFace, roundedRectangle, diamond, triangle, line (use kind for arrows, chevrons and connector glyphs; a thin rectangle is a connector line).",
    "- Never remove or overwrite the user's own content: pictures, charts and tables stay. Redesign AROUND them (move, restyle, wrap in panels, put overlays above). clear_slide only on nearly-empty placeholder slides; duplicate_slide_to replaces its target completely - never aim it at a slide that still carries user content.",
    "- MODERN VISUAL DESIGN (applies whenever the ask creates or redesigns a visual: infographic, diagram, flow, timeline, process, icons, dashboard, comparison): invent ONE palette from the DECK'S own theme colors (one deep + one accent + 2 neutrals; reuse accents already present on nearby slides - do not import random rainbow palettes), then compose in LAYERS: (1) a background statement first - a soft tinted field, a transparent wash (create_shape rectangle, fill, transparency 0.10-0.25) or a color block on one side, (2) structural blocks + accents (thin bars, dots, chevron arrows, donut/star/teardrop glyphs, oversized numbers or keywords - NOT equal boxes in a row), (3) text LAST (ops apply in order - the last op paints on top). Icons come from picture-filled shapes: create_shape circle (or a small square) sized 90-120pt + set_shape_image on it with a flat vector icon prompt (one accent color on a soft neutral background, generous padding, no text inside); remember image ops are paid and slow - ONE icon set per slide, at most 4, else large numbers/letters carry the row. Layout laws: asymmetric compositions + generous whitespace, 2-point alignment grid (left-heavy or split), oversized scale contrast (one big, the rest small); FORBIDDEN 80s tells: 1px black outlines on every shape, saturated primary fills edge-to-edge, equal centered boxes, default thin line arrows when a chevron or rightArrow glyph would carry the flow. The Office JS API has NO shadows or blurs - build depth with transparency layers, never grey fake-shadow shapes. roundedRectangle renders as a sharp rectangle on current Mac builds - design with rectangles + spacing.",
    "- READABILITY LAW for every design or make-over pass: choose the TEXT color from the field it sits on - dark text (near-black) on light/plain/panel fields, light text (near-white) ONLY on deep backgrounds, dark overlays or panel colors with strong contrast; body text >= 16pt on colored fields (>= 14pt anywhere, never below 12pt), titles >= 28pt with an obvious title/body size gap; never place body text directly on busy photos - add a dark overlay or leave an empty field; avoid mid-tone text on mid-tone fills - pick the opposite extreme or set the text on a small solid panel.",
    "- A redesign pass must cover EVERY in-scope slide with its structural change in the same flow. If the op list would be very long, reply phase 1 and immediately continue with phase 2 as the next action JSON - never leave half the slides untouched.",
    "- Text over pictures: order the ops so the picture comes FIRST (set_background or a picture-filled shape), then the overlay (create_shape rectangle, fill #000000, transparency between 0.45 and 0.6), then the text (white or light via replace_text/format_text). Ops apply in the order you return them.",
    "- Make-over phase order per reply: (1) backgrounds + palette fields + overlay shapes, (2) layout moves + text restyle, (3) image generation requests. This keeps each approval batch reviewable and each design phase consistent.",
    "- Image requests: always return a create_image op (never generate the picture inline); SlideKick fetches it with the user's configured image model. Include a concrete, self-contained prompt in the op.",
    "- Image refinement ('refine/change the image'): write a FRESH self-contained image prompt describing the FULL desired result. Never embed previous prompts, the conversation, or op JSON into it - only the scene essentials plus the requested change. Prefer 200-600 words.",
    "- READ SLIDES FIRST: when the user asks you to READ, QUOTE, SHOW or PASTE a slide's content, ALWAYS reply read_slide for that slide BEFORE quoting anything - the outline previews are deliberately cut and quoting one produces wrong text. Full text then arrives from the pane. READ SLIDES ON DEMAND: the outline shows full shape text only for small decks and for the selected slide(s) - in LARGE decks every other slide shows just its title and a short preview. When you need a slide's ACTUAL content (to answer a question about it, or before rewriting it), reply with only {\"kind\":\"read_slide\",\"slides\":[N]} (1-based outline numbers, max 12). SlideKick re-asks you with the COMPLETE text of those slides. Never guess content from a title.",
    "- LIVE WEB LOOKUPS: when the request needs up-to-date info you cannot reliably know (current events, news, prices, recent releases, rankings, live data), reply with ONLY {\"kind\":\"web_lookup\"} - nothing else. SlideKick then runs a web-enabled lookup and re-asks you with fresh citable results. Never use it for deck/document/mail edits, styling, general concepts, or what the outline/context already answers. Lookups are INTERNAL actions - free (never gated/approved) - and you may reply with later rounds of {\"kind\":\"web_lookup\"} as well (up to 3 lookups per run). After a lookup you may also ask for the FULL text of one page from its results by replying with ONLY {\"kind\":\"read_page\",\"url\":\"https://...\"} (internal, free, up to 4 pages per run). When a page's text carries a \"…[page continues]\" ending and the info you need may sit deeper, read the NEXT part by replying with ONLY {\"kind\":\"read_page\",\"url\":\"https://...\",\"from\":<n>} using the part start the pane gave you (continuations are free and do not use the 4-page budget). You have no web access yourself and cannot run a search; the lookup channel belongs to SlideKick - never claim you tried a lookup or that the lookup channel is unavailable, and never refuse for lack of live data: reply with {\"kind\":\"web_lookup\"} instead. For CONTENT-CREATION requests (write an essay/article about... - or build/make the deck\'s slides): the pane has ALREADY run a web lookup before this ask and any fetched results sit in the context above - verify your claims against them, keep facts current and cite the sources you use; if you need MORE, reply {\"kind\":\"web_lookup\"} again (up to 3 rounds per run).",

  ].join("\n");

  var CLARIFY_RULE =
    "\nAMBIGUOUS RESULT FORM or PLACEMENT: (FORM) when the user's ask could plausibly deliver a DIFFERENT FORM " +
    "of the result (e.g. 'infographic', 'visual', 'diagram' = an AI-GENERATED PICTURE vs. EDITABLE SHAPES + text vs. " +
    "both; a paragraph vs. a bulleted list), or (PLACEMENT) when many slides are open and it is unclear WHERE the " +
    "result belongs (the open slide? a named one? a new slide after the named one?), you may ask the user ONCE " +
    "instead of guessing. Reply exactly:" +
    " {\"kind\":\"clarify\",\"question\":\"<one short question>\",\"options\":[\"<short option>\",\"<short option>\"]} " +
    "(2-4 short options, no JSON inside). The pane shows the question; the user's choice arrives in your next ask " +
    "(look for CLARIFICATION) - then deliver the final result for that choice without asking again. When the ask " +
    "DOES name a slide (e.g. 'turn slide 124 into...'), deliver on that slide and never silently move the result " +
    "to the open/selected slide. Never use clarify for style/tone/size details you can decide yourself, and never " +
    "twice in one chain.";
  var DELETE_RULE =
    "- THIS BUILD CANNOT DELETE SLIDES (the add-in API has no slide removal here): NEVER emit delete_slide. If your plan included one, drop it and rebuild around the existing slides - create_slide appends at the end, duplicate_slide_to replaces a target's contents, move_slide repositions. Deliver the complete result WITHOUT the delete; failing a whole story over a missing delete is unacceptable (field: a story plan that deleted the slide it had misplaced).";

  var NOTES_OFF_RULE =
    "- Speaker-notes AUTO-generation is OFF in the user's Settings. For ordinary slide work:" +
    " omit notes fields and never invent set_notes ops. When the user explicitly asks for" +
    " presenter/speaker notes (or clearly wants them for a slide): DO emit the set_notes op" +
    " (and keep notes fields out of create_slide/edit_slide) - the pane will ask the user to" +
    " confirm that one write before applying it; nothing is dropped. NEVER refuse an explicit" +
    " notes request, never tell the user to change Settings first, and never offer to put the" +
    " note text on the slide itself instead.";

  var REPAIR_ASK =
    "Your previous answer did not follow the required format. " +
    "Reply again with ONLY the JSON object (format A or B) and nothing else: " +
    'either {"kind":"reply","text":"..."} or {"kind":"slide_ops","ops":[...]} - ' +
    "no reasoning, no fences, no commentary. Never reply with nothing at all - " +
    "an empty response is a failure; if you cannot act, use format A with the " +
    "reason as text.";

    var CONTINUE_ASK =
    "Your previous reply was cut off mid-way through the op list. Reply with a " +
    "slide_ops JSON containing ONLY the ops still needed to FINISH the redesign " +
    "(slides not yet covered), same JSON format. If nothing is missing, reply " +
    '{"kind":"reply","text":"done"}';

  // cut-off recovery merge: the continuation very often repeats ops it had
  // already emitted (it sees its own partial list in the context). Exact
  // duplicates later in the merged list get dropped so actions do not run
  // twice; the count lands in the proposal note (recoveredDupes).
  function dedupeMergedOps(base, extra) {
    var seen = {}, out = [], dropped = 0;
    base.concat(extra || []).forEach(function (o) {
      var k = null;
      try { k = JSON.stringify(o); } catch (eK) { k = null; }
      if (k && seen[k]) { dropped++; return; }
      if (k) seen[k] = 1;
      out.push(o);
    });
    SK.recoveredDupes = dropped;
    return out;
  }
  function continueAsk(ctx) {
    var n = ctx && typeof ctx.slideCount === "number" ? ctx.slideCount : null;
    return CONTINUE_ASK +
      "\nREMINDERS: op slide numbers are 1-based and MUST be within 1.." + (n || "the current slide count") +
      (n ? " (this deck currently has " + n + " slides)" : "") +
      "; reference only shapes that appeared in the outline." +
      "\nCRITICAL: the operations from your previous (cut-off) reply ARE already kept and will run - do NOT repeat any of them. Emit ONLY genuinely new operations; if your previous list was actually complete, reply {\"kind\":\"reply\",\"text\":\"done\"} instead.";
  }

  var REPAIR_EMPTY_ASK =
    "Your previous reply was EMPTY. An empty response is NEVER acceptable. " +
    "Always produce something useful: the slide_ops JSON with actions for the " +
    "user's request, or format A (reply) with the answer/reason. Respond NOW " +
    "with the JSON object - no fences, no commentary, never an empty reply.";

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
      ref.textContent = "\u21bb";   // circular arrow
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
          try { localStorage.setItem("sk-resume-session-slidekick", SK.sessionId); } catch (eS) {}
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
      try { localStorage.removeItem("sk-approvals-slidekick"); } catch (eRC) {}
      SK.sessionId = doc.id;
      try { localStorage.setItem("sk-resume-session-slidekick", doc.id); } catch (eS) {}
      SK.sessionName = doc.name || "";
      restoreSessionApprovals();   // approvals ride the resumed session (0.9.267)
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
        try { localStorage.removeItem("sk-resume-session-slidekick"); } catch (eS) {}
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
    try { localStorage.removeItem("sk-approvals-slidekick"); } catch (eRC) {}
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

  // wire the sessions tab buttons (called from wireEvents)
  function wireSessions() {
    var nw = $("sess-new");
    if (nw) nw.addEventListener("click", function () { startNewSession(nw); });
    var rf = $("sess-refresh");
    if (rf) rf.addEventListener("click", renderSessions);
  }

  var CONTENT_ROUTE_RE = new RegExp(
    "\\b(slide|slides|deck|bullet|bullets|reword|rewrite|rephrase|outline|notes?|title|titles|add|create|make|generate|tighten|restructure|expand|insert|append|design|align|balance|spacing|table|tables|chart|bold|italic|underline|color|font|highlight|format|formatting|center|caps|image|images|picture|pictures|photo|logo|illustration)\\b", "i"
  );

  function assistantHistoryText(parsed) {
    // a compact recap for the next round; never the raw ops JSON (it bloats every
    // later request and pushes the models toward echoing blobs into new prompts)
    if (parsed && parsed.kind === "clarify") return "(asked the user to pick: " + String(parsed.question || "").slice(0, 140) + ")";
    if (!parsed || !parsed.ok) {
      var rawT = (SK.lastModelText && SK.lastModelText.text) || "";
      rawT = String(rawT);
      // never let a broken/unparsed reply (still containing ops JSON) into
      // history: the model would meet its own truncated JSON next round
      var brace = rawT.indexOf("{");
      if (brace >= 0) return "(the last answer could not be parsed as an action - it was re-asked)";
      return rawT.slice(0, 400);
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
    try { localStorage.removeItem("sk-approvals-slidekick"); } catch (eRC) {}
    SK.sessionId = null;      // the next exchange saves as a NEW session
    try { localStorage.removeItem("sk-resume-session-slidekick"); } catch (eS) {}
    SK.sessionName = "";
    SK.sessionManual = false;   // a new session starts AI-titled again
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
  // Office deck adapter (thin wrapper around PowerPoint.run)
  // ------------------------------------------------------------------ //
  function OfficeDeck() {
    this.office = true;
    this._textApi = undefined; // undefined = unknown, true/false once probed
    this._pfApi = undefined;   // placeholderFormat detection (PowerPointApi 1.8)
    this._caps = null;
  }

  OfficeDeck.prototype.capabilities = function () {
    var sel = false, version = "";
    try {
      sel = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
    } catch (e) { /* ignore */ }
    try { version = (Office.context.diagnostics || {}).version || ""; } catch (e) { /* ignore */ }
    return {
      text: this._textApi !== false,
      textProbed: this._textApi === true || this._textApi === false,
      placeholders: this._pfApi === true,
      selection: sel,
      host: "PowerPoint " + version,
    };
  };

  OfficeDeck.prototype._wrap = function (err) {
    var raw = (err && (err.message || "")) || String(err);
    var msg = raw;
    if (/PropertyNotFound|InvalidArgument.*text/i.test(raw)) {
      msg = "This Office build does not support the required text APIs (PowerPointApi ~1.8). Update Microsoft 365 to use deck edits.";
    } else if (/CannotInferLoadParameter|PropertyOrMethodNotSupported|APINotAvailable/i.test(raw)) {
      msg = "A required Office.js API is missing in this build (" + raw + "). Update Microsoft 365.";
    } else if (/not.*available|unsupported/i.test(raw) && /requirement/i.test(raw)) {
      msg = "API not available in this Office build: " + raw;
    }
    return new FriendlyError(msg);
  };

  OfficeDeck.prototype.run = function (name, fn) {
    var t0 = Date.now();
    return PowerPoint.run(fn).then(
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

  OfficeDeck.prototype._probeTextApi = function () {
    var self = this;
    return this.run("probe", function (ctx) {
      var slides = ctx.presentation.slides;
      slides.load("items/id");
      return ctx.sync().then(function () {
        if (!slides.items.length) return null;
        var shapes = slides.items[0].shapes;
        shapes.load("items/type,items/name,items/id");
        return ctx.sync().then(function () {
          if (!shapes.items.length) return null;
          var tf = shapes.items[0].getTextFrameOrNullObject();
          tf.load("isNullObject,hasText");
          return ctx.sync().then(function () {
            if (tf.isNullObject || !tf.hasText) return null;
            tf.textRange.load("text");
            return ctx.sync();
          });
        });
      });
    }).then(
      function () { self._textApi = true; },
      function () { self._textApi = false; }
    );
  };

  // PlaceholderFormat (PowerPointApi 1.8) is the reliable, localization-safe way
  // to find title/body placeholders. It throws when loaded on non-placeholder
  // shapes, so gate every load on shape.type === "placeholder" and probe once.
  // tables: GraphicFrame shapes have NO textFrame - read cell text through
  // the table API when the build has it (PowerPointApi 1.6+), else say so
  OfficeDeck.prototype._readTableText = function (ctx, sh) {
    var tryRead = function () {
      var tbl = sh.table;
      if (!tbl || !tbl.rows) {
        var bld = (Office.context && Office.context.diagnostics && Office.context.diagnostics.version) || "?";
        var err2 = new Error("this Office build (" + bld + ") does not expose the table APIs to add-ins - update Office to a build with the PowerPointApi 1.8 tables, then ask again");
        err2.noTableProp = true;
        throw err2;
      }
      // PowerPoint.Table/TableCell = 1.8, TableRow = 1.9; read row-by-row
      // (deep loads can be rejected; TableCell.text = a plain string)
      tbl.load("rowCount,columnCount");
      tbl.rows.load("items");
      return ctx.sync().then(function () {
        var cellQs = [];
        (tbl.rows.items || []).forEach(function (row) {
          (row.cells.items || []).forEach(function (cell) {
            cell.load("text");
            cellQs.push(cell);
          });
        });
        return ctx.sync().then(function () {
          var cells = cellQs.map(function (cell) { return String(cell.text || ""); });
          return { r: tbl.rowCount, c: tbl.columnCount, cells: cells };
        });
      });
    };
    return Promise.resolve().then(tryRead).catch(function (eT) {
      // some builds reject the batched cell loads: try the deep-load form
      return Promise.resolve().then(function () {
        var tbl = sh.table;
        if (!tbl || !tbl.rows) throw eT;
        tbl.load("rowCount,columnCount");
        tbl.rows.load("items/cells/text");
        return ctx.sync().then(function () {
          var cells = [];
          (tbl.rows.items || []).forEach(function (row) {
            (row.cells.items || []).forEach(function (cell) {
              cells.push(String(cell.text || ""));
            });
          });
          return { r: tbl.rowCount, c: tbl.columnCount, cells: cells };
        });
      }).catch(function () { throw eT; });
    });
  };
  OfficeDeck.prototype._probeGeometry = function () {
    var self = this;
    if (this._geoApi !== undefined) return Promise.resolve();
    var supported = false;
    try { supported = Office.context.requirements.isSetSupported("PowerPointApi", "1.4"); } catch (e) { /* optimistic */ }
    if (!supported) { this._geoApi = false; return Promise.resolve(); }
    return this.run("probeGeometry", function (ctx) {
      var slides = ctx.presentation.slides;
      slides.load("items/id");
      return ctx.sync().then(function () {
        if (!slides.items.length) { self._geoApi = true; return null; }
        var shapes = slides.items[0].shapes;
        shapes.load("items/left,items/width");
        return ctx.sync().then(function () {
          self._geoApi = true;
          return null;
        });
      });
    }).catch(function () { self._geoApi = false; });
  };

  OfficeDeck.prototype._probePlaceholders = function () {
    var self = this;
    if (this._pfApi !== undefined) return Promise.resolve();
    return this.run("probePlaceholders", function (ctx) {
      var slides = ctx.presentation.slides;
      slides.load("items/id");
      return ctx.sync().then(function () {
        if (!slides.items.length) return "skip";
        var shapes = slides.items[0].shapes;
        shapes.load("items/type,items/name,items/id");
        return ctx.sync().then(function () {
          var ph = null;
          shapes.items.forEach(function (sh) {
            if (sh.type === "placeholder" && !ph) ph = sh;
          });
          if (!ph) return "skip"; // nothing to probe; keep name heuristics
          ph.load("placeholderFormat");
          return ctx.sync().then(function () {
            ph.placeholderFormat.load("type");
            return ctx.sync();
          });
        });
      });
    }).then(
      function (res) { if (res !== "skip") self._pfApi = true; },
      function () { self._pfApi = false; } // older host: fall back to names
    );
  };

  OfficeDeck.prototype.getContext = function (opts) {
    opts = opts || {};
    var self = this;
    var pre = Promise.resolve();
    if (this._textApi === undefined) pre = pre.then(function () { return self._probeTextApi(); });
    if (this._pfApi === undefined) pre = pre.then(function () { return self._probePlaceholders(); });
    if (this._geoApi === undefined) pre = pre.then(function () { return self._probeGeometry(); });
    return pre.then(function () {
      var wantText = self._textApi !== false;
      var wantPf = self._pfApi === true;
      var limit = Math.min(opts.maxSlides || 200, 400);
      return self.run("getContext", function (ctx) {
        var pres = ctx.presentation;
        var slides = pres.slides;
        slides.load("items/id");
        return ctx.sync().then(function () {
          // 0.9.287: slides.removeAt exists only on builds with slide
          // deletion - surface it so the ask and the outline can forbid
          // delete_slide ops instead of letting them fail at apply time
          try { ctx.canDelete = (typeof slides.removeAt === "function"); }
          catch (eDel) { ctx.canDelete = undefined; }
          var count = slides.items.length;
          var n = Math.min(count, limit);
          for (var i = 0; i < n; i++) {
            slides.items[i].shapes.load("items/type,items/name,items/id" +
              (self._geoApi ? ",items/left,items/top,items/width,items/height" : ""));
          }
          return ctx.sync().then(function () {
            var data = [];
            var flat = [];     // same order as data shapes, flattened
            var tfs = [];
            var phList = [];
            for (var i2 = 0; i2 < n; i2++) {
              var shpProxies = slides.items[i2].shapes.items;
              var slideData = { idx: i2, shapes: [] };
              for (var j = 0; j < shpProxies.length; j++) {
                var sh = shpProxies[j];
                var entry = {
                  idx: j, name: sh.name || "",
                  type: sh.type || "", text: undefined, ptype: undefined,
                };
                if (self._geoApi) {
                  entry.left = sh.left; entry.top = sh.top;
                  entry.width = sh.width; entry.height = sh.height;
                }
                slideData.shapes.push(entry);
                flat.push(slideData.shapes[slideData.shapes.length - 1]);
                if (wantText) {
                  var tf = sh.getTextFrameOrNullObject();
                  tf.load("isNullObject,hasText");
                  tfs.push({ tf: tf, shape: flat[flat.length - 1] });
                }
                if (wantPf && sh.type === "placeholder") {
                  sh.load("placeholderFormat");
                  phList.push({ sh: sh, flatIdx: flat.length - 1 });
                }
              }
              data.push(slideData);
            }
            var afterLoads = (wantText || wantPf) ? ctx.sync() : Promise.resolve();
            return afterLoads.then(function () {
              phList.forEach(function (e) { e.sh.placeholderFormat.load("type"); });
              var afterPf = wantPf ? ctx.sync() : Promise.resolve();
              return afterPf.then(function () {
                phList.forEach(function (e) {
                  flat[e.flatIdx].ptype = e.sh.placeholderFormat.type;
                });
                var trs = [];
                tfs.forEach(function (entry) {
                  if (!entry.tf.isNullObject && entry.tf.hasText) {
                    entry.tf.textRange.load("text");
                    trs.push(entry);
                  }
                });
                var afterTr = wantText ? ctx.sync() : Promise.resolve();
                return afterTr.then(function () {
                  trs.forEach(function (entry) { entry.shape.text = entry.tf.textRange.text || ""; });
                  if (wantText) {
                    data.forEach(function (s) {
                      s.shapes.forEach(function (sh) { if (sh.text === undefined) sh.text = ""; });
                    });
                  }
                }).then(function () {
                  // tables: GraphicFrame shapes have no textFrame; pull the
                  // cell text through the table API (guarded per shape)
                  var tblShapes = [];
                  data.forEach(function (sd) {
                    sd.shapes.forEach(function (sh) { if (sh.type === "table") tblShapes.push(sh); });
                  });
                  if (!tblShapes.length) return null;
                  var tblApi = true;
                  try { tblApi = Office.context.requirements.isSetSupported("PowerPointApi", "1.8"); } catch (eTbl) { tblApi = false; }
                  var tries = tblShapes.map(function (sh) {
                    if (!tblApi) return { txt: "[table - this Office build does not expose table cell APIs to add-ins (needs the PowerPoint table API, set 1.8) - update Office, then ask again]" };
                    return self._readTableText(ctx, sh)
                      .then(function (info) {
                        if (!info || !info.cells || !info.cells.length) return "[table " + info.r + "x" + info.c + " - unreadable]";
                        var joined = info.cells.join(" | ");
                        var head = "[table " + info.r + "x" + info.c + "] ";
                        return head + joined;
                      })
                      .catch(function (eT) {
                        var why = eT && (eT.message || String(eT));
                        if (eT && eT.noTableProp) why = eT.message;
                        dbg("table.read", { failed: true, why: why });
                        return "[table - cell text unavailable in this Office build (" + why + ")]";
                      });
                  });
                  return Promise.all(tries).then(function (vals) {
                    tblShapes.forEach(function (sh, k) { sh.text = vals[k]; });
                    return null;
                  });
                }).then(function () {
                  var selected = [];
                  var selSupported = false;
                  try { selSupported = !!pres.getSelectedSlides; } catch (e) { selSupported = false; }
                  if (!selSupported) return { count: count, data: data, selected: [] };
                  var sel = pres.getSelectedSlides();
                  sel.load("items/id");
                  return ctx.sync().then(function () {
                    var idToIdx = {};
                    for (var i3 = 0; i3 < n; i3++) idToIdx[slides.items[i3].id] = i3;
                    (sel.items || []).forEach(function (s) {
                      if (idToIdx[s.id] !== undefined) selected.push(idToIdx[s.id]);
                    });
                    return { count: count, data: data, selected: selected };
                  });
                });
              });
            });
          });
        });
      }).then(function (r) {
        return {
          slideCount: r.count,
          slides: r.data.slice(0, limit),
          selected: r.selected || [],
          truncated: r.count > limit,
          host: "PowerPoint",
        };
      });
    });
  };

  OfficeDeck.prototype._pickLayout = function (masters, kind) {
    if (!masters || !masters.items || !masters.items.length) return null;
    var master = masters.items[0];
    var layouts = master.layouts && master.layouts.items ? master.layouts.items : [];
    var best = null, bestScore = 0;
    layouts.forEach(function (lay) {
      var s = scoreLayout(lay.name, kind);
      if (s > bestScore) { bestScore = s; best = lay; }
    });
    return best;
  };

  function scoreLayout(name, kind) {
    var n = (name || "").toLowerCase();
    switch (kind) {
      case "title_only":
        return /title only/.test(n) ? 3 : 0;
      case "section":
        return /section/.test(n) ? 3 : 0;
      case "blank":
        return n === "blank" ? 3 : /blank/.test(n) ? 2 : 0;
      case "title_content":
        if (/title/.test(n) && /(content|text|body)/.test(n) && !/vertical/.test(n)) return 3;
        if (/(content|body)/.test(n)) return 2;
        return 0;
    }
    return 0;
  }

  OfficeDeck.prototype.addSlide = function (op) {
    var kind = normalizeLayout(op.layout);
    var self = this;
    return this.run("addSlide", function (ctx) {
      var pres = ctx.presentation;
      var masters = pres.slideMasters;
      masters.load("id, name, layouts/items/name, layouts/items/id");
      return ctx.sync().then(function () {
        var lay = self._pickLayout(masters, kind);
        var opts = lay ? { slideMasterId: masters.items[0].id, layoutId: lay.id } : undefined;
        pres.slides.add(opts);
        return ctx.sync().then(function () {
          var slides = pres.slides;
          slides.load("items/id");
          return ctx.sync().then(function () {
            var newIndex = slides.items.length - 1;
            var slide = slides.getItemAt(newIndex);
            var shapes = slide.shapes;
            shapes.load("items/type,items/name,items/id");
            return ctx.sync().then(function () {
              var metas = [];
              var phList = [];
              var wantPf = self._pfApi === true;
              shapes.items.forEach(function (sh, j) {
                metas.push({ sh: sh, name: sh.name || "", type: sh.type || "", ptype: undefined, tf: null });
                var tf = sh.getTextFrameOrNullObject();
                tf.load("isNullObject,hasText");
                metas[j].tf = tf;
                if (wantPf && sh.type === "placeholder") {
                  sh.load("placeholderFormat");
                  phList.push({ sh: sh, idx: j });
                }
              });
              return ctx.sync().then(function () {
                phList.forEach(function (e) { e.sh.placeholderFormat.load("type"); });
                var afterPf = wantPf ? ctx.sync() : Promise.resolve();
                return afterPf.then(function () {
                  phList.forEach(function (e) { metas[e.idx].ptype = e.sh.placeholderFormat.type; });
                  var titleEntry = null, bodyEntry = null;
                  metas.forEach(function (m) {
                    if (m.tf.isNullObject) return;
                    if (!titleEntry && isTitleShape(m.name, m.type, m.ptype)) titleEntry = m;
                    else if (!bodyEntry && isBodyShape(m.name, m.type, m.ptype)) bodyEntry = m;
                  });
                  var toSet = [];
                  if (op.title && titleEntry) {
                    titleEntry.tf.textRange.load("text");
                    toSet.push({ e: titleEntry, v: String(op.title) });
                  }
                  if (op.body && op.body.length && bodyEntry) {
                    bodyEntry.tf.textRange.load("text");
                    toSet.push({ e: bodyEntry, v: op.body.join("\r") });
                  }
                  return ctx.sync().then(function () {
                    toSet.forEach(function (t) { t.e.tf.textRange.text = t.v; });
                    return ctx.sync().then(function () {
                      return { index: newIndex };
                    });
                  });
                });
              });
            });
          });
        });
      });
    });
  };

    OfficeDeck.prototype.addTable = function (op) {
    var self = this;
    // capability gate: PowerPointApi 1.8
    var tblOk = false;
    try { tblOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.8"); } catch (e) { /* ignore */ }
    if (!tblOk) {
      return Promise.reject(new FriendlyError(
        "Table creation needs PowerPointApi 1.8 - this PowerPoint build does not support it (update Microsoft 365)."));
    }
    return this.run("addTable", function (ctx) {
      var slide = ctx.presentation.slides.getItemAt(op.slide);
      var opt = { values: op.values };
      if (op.geometry && op.geometry.left != null) opt.left = op.geometry.left;
      if (op.geometry && op.geometry.top != null) opt.top = op.geometry.top;
      if (op.geometry && op.geometry.width != null) opt.width = op.geometry.width;
      slide.shapes.addTable(op.values.length, op.values[0].length, opt);
      return ctx.sync().then(function () {
        return { ok: true, newIndex: op.slide };
      });
    });
  };

  OfficeDeck.prototype.copyPasteShape = function (srcSlide, srcShape, dstSlide, geo, cut) {
    // 0.9.213 - the Word cut/paste idiom for PowerPoint objects:
    // copy the shape (full fidelity, incl. pictures/tables) via the
    // Office-JS shape copy/paste pair, drop it on the target slide,
    // optionally re-place it, and - for a MOVE - delete the source.
    return this.run(cut ? "moveShapeTo" : "copyShape", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.13")) {
        throw new FriendlyError("shape copy/paste needs PowerPointApi 1.13 (PowerPoint 16.108+): " +
          "update PowerPoint, or move the shape with move_shape (position/size) instead.");
      }
      var s = ctx.presentation.slides.getItemAt(srcSlide);
      var sh = s.shapes.getItemAt(srcShape);
      sh.copy();
      return ctx.sync().then(function () {
        var d = ctx.presentation.slides.getItemAt(dstSlide);
        d.shapes.paste();
        return ctx.sync().then(function () {
          var how = (cut ? "moved" : "copied") + " to slide " + (dstSlide + 1);
          var pasted = d.shapes.getItemAt(d.shapes.items.length - 1);
          var placed = false;
          ["left", "top", "width", "height"].forEach(function (p) {
            if (geo && geo[p] != null) { pasted[p] = geo[p]; placed = true; }
          });
          if (cut) sh.delete();
          return ctx.sync().then(function () {
            return { ok: true, how: how + (placed && geo && geo.left != null ? " (replaced: left " + Math.round(geo.left) + ", top " + Math.round(geo.top) + ")" : " (same position)") };
          });
        });
      });
    });
  };

  OfficeDeck.prototype.copyShape = function (srcSlide, srcShape, dstSlide, geo) {
    return this.copyPasteShape(srcSlide, srcShape, dstSlide, geo, false);
  };

  OfficeDeck.prototype.moveShapeTo = function (srcSlide, srcShape, dstSlide, geo) {
    return this.copyPasteShape(srcSlide, srcShape, dstSlide, geo, true);
  };

  OfficeDeck.prototype.moveShape = function (slideIndex, shapeIndex, geometry) {
    return this.run("moveShape", function (ctx) {
      var shape = ctx.presentation.slides.getItemAt(slideIndex).shapes.getItemAt(shapeIndex);
      ["left", "top", "width", "height"].forEach(function (p) {
        if (geometry[p] != null) shape[p] = geometry[p];
      });
      if (geometry.rotation != null) {
        if (!Office.context.requirements.isSetSupported("PowerPointApi", 1.10))
          throw new FriendlyError("rotating shapes needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
        shape.rotation = geometry.rotation;
      }
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

OfficeDeck.prototype.addImage = function (payload) {
  // Laddered insert (0.9.263): PowerPoint-Mac rejects some payloads with a
  // bare host error ("An internal error has occurred") - oversized PNGs are
  // the usual trigger (setSelectedDataAsync data limit). Attempt 1: the raw
  // bytes (prefix stripped). Attempt 2: re-encode via canvas to JPEG 0.85
  // (white-backed - transparency is filled, not kept) and retry with the
  // smaller payload. The failure carries the payload size so the next
  // diagnosis starts from evidence.
  function kb(s) { return (s.length * 0.75 / 1024).toFixed(0); }
  function stripPrefix(b64) {
    var s = String(b64 || "");
    var m = s.match(/^data:[^;]+;base64,/);
    return m ? s.slice(m[0].length) : s;
  }
  function reencode(b64, q) {
    return new Promise(function (res) {
      try {
        var img = new Image();
        var url = "data:image/png;base64," + b64;
        img.onload = function () {
          try {
            var c = document.createElement("canvas");
            c.width = img.naturalWidth || img.width || 1024;
            c.height = img.naturalHeight || img.height || 1024;
            var g2 = c.getContext("2d");
            g2.fillStyle = "#FFFFFF";
            g2.fillRect(0, 0, c.width, c.height);
            g2.drawImage(img, 0, 0, c.width, c.height);
            var out = String(c.toDataURL("image/jpeg", q)).split(",")[1];
            res(out);
          } catch (eC) { res(null); }
        };
        img.onerror = function () { res(null); };
        img.src = url;
      } catch (eT) { res(null); }
    });
  }
  function attempt(b64) {
    return new Promise(function (resolve, reject) {
      try {
        var opts = { coercionType: Office.CoercionType.Image };
        var geo = payload.geometry || {};
        if (geo.left != null) opts.imageLeft = geo.left;
        if (geo.top != null) opts.imageTop = geo.top;
        if (geo.width != null) opts.imageWidth = geo.width;
        if (geo.height != null) opts.imageHeight = geo.height;
        Office.context.document.setSelectedDataAsync(b64, opts, function (ar) {
          if (ar && ar.status === Office.AsyncResultStatus.Succeeded) resolve({ ok: true });
          else reject(new Error((ar && ar.error && ar.error.message) || (ar && ar.status) || "unknown"));
        });
      } catch (e) {
        reject(new Error(e && e.message || e));
      }
    });
  }
  var raw0 = stripPrefix(payload.b64);
  var self = this;
  function shapeFallback() {
    // attempt 3 (0.9.264): the direct image insert is failing on this host
    // regardless of payload - use the PROVEN primitive instead: a rectangle
    // at the requested geometry picture-filled via fill.setImage (the same
    // code path as set_shape_image, shipped at PowerPointApi 1.10).
    return self.addImageShape({ b64: raw0, slide: payload.slide, geometry: payload.geometry,
                                why: e0msg });
  }
  var e0msg = "";
  var raw0b = raw0;
  return attempt(raw0b).catch(function (e1) {
    e0msg = String(e1.message || e1);
    var msg1 = e0msg;
    return reencode(raw0b, 0.85).then(function (small) {
      if (!small) return shapeFallback();
      return attempt(small).then(function (r) {
        r.note = "inserted as JPEG " + kb(small) + " KB (the host rejected the original " +
          kb(raw0b) + " KB payload)";
        return r;
      }, function (e2) {
        dbg("office.addImage.jpegFail", { error: String(e2.message || e2).slice(0, 120) });
        return shapeFallback().catch(function (e3) {
          throw new FriendlyError("image insert failed: " + msg1 +
            " (payload " + kb(raw0b) + " KB; also as JPEG " + kb(small) + " KB: " +
            String((e2 && e2.message) || e2) + "; shape fallback: " +
            String((e3 && e3.message) || e3).slice(0, 80) + ")");
        });
      });
    }, function () { return shapeFallback(); });
  });
};

OfficeDeck.prototype.addImageShape = function (payload) {
  // (0.9.265): the PROVEN primitive for placing an image on an EXACT slide.
  // The direct setSelectedDataAsync insert always lands on the slide the
  // user has open (no slide-targeting exists in the PowerPoint JS API), so
  // batches that just created their target slide need this path. A
  // rectangle at the requested geometry is picture-filled via fill.setImage
  // (the same code path as set_shape_image, PowerPointApi 1.10).
  var kb = function (s) { return (s.length * 0.75 / 1024).toFixed(0); };
  var self = this;
  var raw0 = String(payload.b64 || "");
  var m0 = raw0.match(/^data:[^;]+;base64,/);
  if (m0) raw0 = raw0.slice(m0[0].length);
  var why = payload.why || "the direct insert always lands on the slide you have open";
  return self.run("addImageShape", function (ctx) {
    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
      throw new FriendlyError("placing an image on a specific slide needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
    }
    var geo = payload.geometry || {};
    var g0 = { left: geo.left != null ? geo.left : 80,
               top: geo.top != null ? geo.top : 110,
               width: geo.width != null ? geo.width : 800,
               height: geo.height != null ? geo.height : 450 };
    var sl = payload.slide != null ? payload.slide : 0;
    var shapes = ctx.presentation.slides.getItemAt(sl).shapes;
    var sh = shapes.addGeometricShape("rectangle", g0);
    try { sh.name = "«redacted:sk-…»"; } catch (eN) { /* naming optional */ }
    sh.fill.setImage(raw0);
    return ctx.sync().then(function () {
      return { ok: true,
        note: "inserted on slide " + (sl + 1) + " as a picture-filled rectangle (" + kb(raw0) + " KB)" +
          " - direct inserts land on the slide you have open (" + String(why).slice(0, 60) + ")" };
    });
  });
};

OfficeDeck.prototype.moveSlide = function (from, to) {
    // PowerPoint has no reorder primitive: export the slide (1.10), delete it,
    // then insertSlidesFromBase64 at the final position (after target / at the
    // beginning when no target).
    return this.run("moveSlide", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        throw new FriendlyError("moving slides needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
      }
      if (from === to) return Promise.resolve({ ok: true });
      var slides = ctx.presentation.slides;
      slides.load("items/id");
      return ctx.sync().then(function () {
        var count0 = slides.items.length;
        if (from < 0 || from >= count0) throw new FriendlyError("Slide " + (from + 1) + " does not exist");
        var expRes = slides.exportAsBase64Presentation([slides.getItemAt(from)]);   // ClientResult - value after sync
          return ctx.sync().then(function () {
            var b64 = expRes.value;
            slides.getItemAt(from).delete();
            return ctx.sync().then(function () {
            if (to <= 0) {
              ctx.presentation.insertSlidesFromBase64(String(b64), { formatting: "KeepSourceFormatting" });
              return ctx.sync().then(function () { return { ok: true }; });
            }
            slides.load("items/id");
            return ctx.sync().then(function () {
              var ids2 = slides.items.map(function (x) { return x.id; });
              var targetIdx = Math.min(to - 1, ids2.length - 1);
              if (targetIdx < 0) {
                ctx.presentation.insertSlidesFromBase64(String(b64), { formatting: "KeepSourceFormatting" });
              } else {
                ctx.presentation.insertSlidesFromBase64(String(b64), { formatting: "KeepSourceFormatting", targetSlideId: ids2[targetIdx] });
              }
              return ctx.sync().then(function () { return { ok: true }; });
            });
          });
        });
      });
    });
  };

  // Insert an exported b64 deck right after the slide with id targetIdAfter
  // (no targetSlideId = insert at the beginning).
  OfficeDeck.prototype._insertExported = function (ctx, b64, targetIdAfter) {
    var opts = { formatting: "KeepSourceFormatting" };
    if (targetIdAfter != null) opts.targetSlideId = targetIdAfter;
    ctx.presentation.insertSlidesFromBase64(String(b64), opts);
    return ctx.sync();
  };

  OfficeDeck.prototype.duplicateSlide = function (i) {
    // exact copy of slide i (export + insert right after itself; deck grows)
    return this.run("duplicateSlide", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        throw new FriendlyError("duplicating slides needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
      }
      var slides = ctx.presentation.slides;
      slides.load("items/id");
      return ctx.sync().then(function () {
        if (i < 0 || i >= slides.items.length) throw new FriendlyError("Slide " + (i + 1) + " does not exist");
        var expRes = slides.exportAsBase64Presentation([slides.getItemAt(i)]);   // ClientResult - value after sync
        return ctx.sync().then(function () {
          var b64 = expRes.value;
          ctx.presentation.insertSlidesFromBase64(String(b64),
            { formatting: "KeepSourceFormatting", targetSlideId: slides.items[i].id });
          return ctx.sync().then(function () { return { ok: true, at: i + 1 }; });
        });
      });
    });
  };

  OfficeDeck.prototype.duplicateSlideTo = function (from, to) {
    // copy slide from fully over slide to (both 0-based): export the source,
    // delete the target, insert the exact copy at the target's old position.
    return this.run("duplicateSlideTo", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        throw new FriendlyError("copying slides needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
      }
      if (from === to) return Promise.resolve({ ok: true });
      var slides = ctx.presentation.slides;
      slides.load("items/id");
      return ctx.sync().then(function () {
        var count0 = slides.items.length;
        if (from < 0 || from >= count0) throw new FriendlyError("Slide " + (from + 1) + " does not exist");
        if (to < 0 || to >= count0) throw new FriendlyError("Slide " + (to + 1) + " does not exist");
        var expRes = slides.exportAsBase64Presentation([slides.getItemAt(from)]);   // ClientResult - value after sync
        return ctx.sync().then(function () {
          var b64 = expRes.value;
          slides.getItemAt(to).delete();
          return ctx.sync().then(function () {
            slides.load("items/id");
            return ctx.sync().then(function () {
              var ids = slides.items.map(function (x) { return x.id; });
              // after deleting target 'to' the copy lands back at 0-based 'to':
              // the API inserts AFTER targetSlideId, so target = ids[to-1]
              // (0-based); with to == 0 there is no target = insert at top.
              var opts = { formatting: "KeepSourceFormatting" };
              if (ids.length && to > 0) opts.targetSlideId = ids[Math.min(to - 1, ids.length - 1)];
              ctx.presentation.insertSlidesFromBase64(String(b64), opts);
              return ctx.sync().then(function () { return { ok: true }; });
            });
          });
        });
      });
    });
  };

  OfficeDeck.prototype.clearSlide = function (i) {
    // empty slide i: wipe every text frame and drop text-free created lines.
    // Placeholders stay in place (their layout keeps its positions).
    // Windows notes (16.0.19x): deep "items/textFrame" collection loads
    // throw InvalidArgument (scalar-only loads there); and batch-issued
    // empty-string writes can be dropped silently, so the clear is issued
    // per shape with a read-back verification + retry pass.
    return this.run("clearSlide", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
        throw new FriendlyError("clearing a slide needs PowerPointApi 1.5 - update PowerPoint (16.63+ on Mac).");
      }
      var shapes = ctx.presentation.slides.getItemAt(i).shapes;
      shapes.load("items/type,items/hasTextFrame,items/name");
      return ctx.sync().then(function () {
        var tfs = [], removed = 0;
        shapes.items.forEach(function (shape) {
          if (shape.hasTextFrame && shape.textFrame) {
            try { shape.textFrame.load("textRange"); tfs.push(shape); } catch (eTf) { /* skip frame */ }
          } else if (shape.type === "Line" && String(shape.name || "").indexOf("sk-shape") === 0) {
            // only drop lines this add-in created - user lines stay
            try { shape.delete(); removed++; } catch (eLine) { /* keep line */ }
          }
        });
        return ctx.sync().then(function () {
          var per = Promise.resolve(), cleared = 0, undid = [];
          tfs.forEach(function (shape) {
            per = per.then(function () {
              try {
                shape.textFrame.textRange.text = "";
                cleared++;
                return ctx.sync();   // issue + flush per frame: Windows drops same-batch empty writes
              } catch (eW) {
                undid.push(shape);
                return Promise.resolve();
              }
            });
          });
          return per.then(function () {
            return { ok: true, cleared: cleared, removed: removed, failedFrames: undid.length };
          });
        });
      });
    });
  };

OfficeDeck.prototype.createShape = function (slideIndex, op) {
  // 0.9.233/234: PowerPoint 16.114 (mac) rejected shapes.addTextBox(text,
  // opts) with RichApi InvalidArgument at sync for EVERY textbox op of a
  // 12-op run while the four addGeometricShape calls with IDENTICAL
  // ShapeAddOptions succeeded, and MS documents no InvalidArgument
  // condition that those arguments could trigger (width/height are
  // positive ints). Ladder the textbox op across three creation paths,
  // each attempt in its OWN PowerPoint.run (a failed sync poisons the
  // context): v1 native addTextBox(text, opts); v2 the docs-sample form
  // addTextBox(text) with geometry written as shape properties; v3
  // emulated - a no-fill/no-outline rectangle + textFrame.textRange.text
  // (1.5). Non-textbox kinds keep the single native attempt.
  var self = this;
  var runVariant = function (vi) {
    var txt = String(op.text || "");
    return self.run("createShape(v" + vi + ")", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.4")) {
        throw new FriendlyError("shape creation needs PowerPointApi 1.4 - update PowerPoint (16.62+ on Mac).");
      }
      var shapes = ctx.presentation.slides.getItemAt(slideIndex).shapes;
      var opts = {};
      ["left", "top", "width", "height"].forEach(function (gk) {
        if (op.geometry && op.geometry[gk] != null) opts[gk] = op.geometry[gk];
      });
      var created;
      var txt = String(op.text || "");
      if (op.kind === "textbox") {
        if (vi === 1) {
          created = shapes.addTextBox(txt, opts);
        } else if (vi === 2) {
          created = shapes.addTextBox(txt);
          ["left", "top", "width", "height"].forEach(function (gk) {
            if (opts[gk] != null) { try { created[gk] = opts[gk]; } catch (eG) { /* geometry optional */ } }
          });
        } else {
          // v3 moved out of this run: the geometric path + an isolated text
          // run (runVariant3) - a textFrame rejection here must not take
          // the created shape with it. This branch is unreachable.
        }
      } else if (op.kind === "line") {
        created = shapes.addLine("Straight", opts);
      } else if (vi === 1) {
        created = shapes.addGeometricShape(op.kind, opts);
      } else if (vi === 2) {
        // v2 (0.9.261): the docs-sample form - the same repair that fixed
        // textboxes on 16.114: create BARE, then write the geometry as
        // shape properties. Some Mac builds reject ShapeAddOptions at sync
        // for specific geometric kinds (roundedRectangle threw InvalidArgument
        // on 16.115 while addTextBox/addGeometricShape("rectangle") with the
        // identical options object succeeded).
        created = shapes.addGeometricShape(op.kind);
        ["left", "top", "width", "height"].forEach(function (gk) {
          if (opts[gk] != null) { try { created[gk] = opts[gk]; } catch (eG) { /* geometry optional */ } }
        });
      } else {
        // v3 (0.9.264): the build rejects the KIND in both forms
        // (roundedRectangle: InvalidArgument in native + property form on
        // 16.115 while rectangle/line/triangle with the SAME options all
        // apply). Create a plain rectangle at the same geometry - the box
        // keeps its intent (corner radius is lost) and the row reports it.
        created = shapes.addGeometricShape("rectangle", opts);
      }
      try { created.name = "sk-shape-" + (op.kind || "shape"); } catch (eName) { /* naming optional */ }
      if (created && op.fill) {
        created.fill.setSolidColor(String(op.fill));
      }
      if (created && op.transparency != null) {
        if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
          throw new FriendlyError("fill transparency needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
        }
        created.fill.transparency = Number(op.transparency);
      }
      if (created && op.line) {
        created.lineFormat.visible = true;
        created.lineFormat.color = String(op.line);
      }
      if (created && op.geometry && op.geometry.rotation != null) {
        if (!Office.context.requirements.isSetSupported("PowerPointApi", 1.10))
          throw new FriendlyError("rotating shapes needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
        created.rotation = op.geometry.rotation;
      }
      return ctx.sync().then(function () {
        if (vi > 1) dbg("office.createShape.fallback", { variant: vi, kind: op.kind, slide: op.slide });
        if (vi === 3 && op.kind !== "textbox" && op.kind !== "line") {
          return { ok: true, variant: vi,
            note: "this PowerPoint build rejected '" + op.kind + "' - used a plain rectangle" };
        }
        return { ok: true, variant: vi };
      });
    }).catch(function (err) {
      var msg = String((err && err.message) || err);
      var tbLadder = op.kind === "textbox" && vi < 2 && !/needs PowerPointApi/.test(msg);
      var geoKind = op.kind !== "textbox" && op.kind !== "line";
      var geoLadder = geoKind && !/needs PowerPointApi/.test(msg) &&
        /RichApi|InvalidArgument|GeneralException|argument|invalid|error/i.test(msg);
      if (tbLadder || geoLadder) {
        // this build rejected the creation path - try the next variant
        // (textbox ladder 0.9.233; geometric ladder 0.9.261/0.9.264)
        dbg("office.createShape.attempt", { variant: vi, kind: op.kind, error: msg.slice(0, 120) });
        return runVariant(vi + 1);
      }
      throw err;
    });
    };
    var runVariant3 = function () {
    var txt = String(op.text || "");
    // the geometric path is the one PROVEN working on this build (rects
    // with styling applied fine while textbox calls all threw); the text
    // write runs ISOLATED afterwards so a textFrame rejection cannot take
    // the created shape with it
    return self.run("createShape(v3)", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.4")) {
        throw new FriendlyError("shape creation needs PowerPointApi 1.4 - update PowerPoint (16.62+ on Mac).");
      }
      var shapes = ctx.presentation.slides.getItemAt(slideIndex).shapes;
      var opts = {};
      ["left", "top", "width", "height"].forEach(function (gk) {
        if (op.geometry && op.geometry[gk] != null) opts[gk] = op.geometry[gk];
      });
      var created = shapes.addGeometricShape("rectangle", opts);
      try { created.name = "sk-shape-textbox"; } catch (eName) { /* naming optional */ }
      if (op.fill) {
        created.fill.setSolidColor(String(op.fill));
      }
      if (op.transparency != null) {
        if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
          throw new FriendlyError("fill transparency needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
        }
        created.fill.transparency = Number(op.transparency);
      }
      if (op.line) {
        created.lineFormat.visible = true;
        created.lineFormat.color = String(op.line);
      }
      if (op.geometry && op.geometry.rotation != null) {
        if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
          throw new FriendlyError("rotating shapes needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
        }
        created.rotation = op.geometry.rotation;
      }
      shapes.load("items/id,items/name");
      return ctx.sync().then(function () {
        var shapeId = "";
        for (var i = shapes.items.length - 1; i >= 0; i--) {
          if (String(shapes.items[i].name) === "sk-shape-textbox") { shapeId = String(shapes.items[i].id); break; }
        }
        dbg("office.createShape.fallback", { variant: 3, kind: "textbox", slide: op.slide, shapeId: shapeId });
        return shapeId;
      });
    }).then(function (shapeId) {
      if (!txt.trim()) return { ok: true, variant: 3 };
      return self.run("createShapeText(v3)", function (ctx) {
        if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
          throw new FriendlyError("writing textbox text needs PowerPointApi 1.5 - update PowerPoint (16.64+ on Mac).");
        }
        var shapes2 = ctx.presentation.slides.getItemAt(slideIndex).shapes;
        var sh2 = shapes2.getItem(shapeId);
        var tf = sh2.getTextFrameOrNullObject();
        tf.load("isNullObject");
        return ctx.sync().then(function () {
          if (tf.isNullObject) {
            dbg("office.createShape.textSkip", { slide: op.slide });
            return { ok: true, variant: 3, note: "created the shape but it carries no text frame" };
          }
          tf.wordWrap = true;
          tf.textRange.text = txt;
          return ctx.sync().then(function () { return { ok: true, variant: 3, textWritten: true }; });
        });
      }).catch(function (eTxt) {
        var msgTxt = String((eTxt && eTxt.message) || eTxt);
        dbg("office.createShape.textFailed", { slide: op.slide, error: msgTxt.slice(0, 120) });
        return { ok: true, variant: 3, note: "created the shape but the text write failed (" + msgTxt.slice(0, 80) + ")" };
      });
    });
    };
    return runVariant(1);
    };

OfficeDeck.prototype.formatSelection = function (font) {
    return this.run("formatSelection", function (ctx) {
      if (!ctx.presentation.getSelectedTextRange) {
        throw new FriendlyError("this PowerPoint build cannot read the text you selected (needs PowerPointApi 1.6).");
      }
      var range = ctx.presentation.getSelectedTextRange();
      range.load("isNullObject");
      return ctx.sync().then(function () {
        if (range.isNullObject) {
          throw new FriendlyError("Select some text in the presentation first, then run this again.");
        }
        if (font.bold != null) range.font.bold = font.bold;
        if (font.italic != null) range.font.italic = font.italic;
        if (font.allCaps != null) range.font.allCaps = font.allCaps;
        if (font.underline != null) range.font.underline = font.underline;
        if (font.color) range.font.color = font.color;
        if (font.name) range.font.name = font.name;
        if (font.size) range.font.size = font.size;
        return ctx.sync().then(function () { return { ok: true }; });
      });
    });
  };

OfficeDeck.prototype.formatTextSubstring = function (slideIndex, shapeIndex, findText, font) {
    return this.run("formatTextSubstring", function (ctx) {
      var shapesS = ctx.presentation.slides.getItemAt(slideIndex).shapes;
      shapesS.load("items/type");
      return ctx.sync().then(function () {
        if (shapeIndex >= shapesS.items.length) {
          throw new FriendlyError("shape " + (shapeIndex + 1) + " does not exist on slide " + (slideIndex + 1) +
            " (the slide has " + shapesS.items.length + " shape(s)) - the action probably targets a shape" +
            " that was not created or was removed; create the shape first, then format it.");
        }
        var shape = shapesS.getItemAt(shapeIndex);
        shape.load("type");
        var range = shape.textFrame.textRange;
        range.load("text");
        return ctx.sync().then(function () {
          var ty = String(shape.type || "");
        if (/^(table|chart|picture|media|ole|graphic|unsupported)/i.test(ty)) {
          throw new FriendlyError("shape " + (shapeIndex + 1) + " on slide " + (slideIndex + 1) +
            " has no text to format (it is a " + ty.toLowerCase() + " shape) - format_text needs a text shape (textbox or placeholder); a table\'s text lives in its cells.");
        }
        var idx = String(range.text || "").toLowerCase().indexOf(String(findText).toLowerCase());
        if (idx < 0) throw new FriendlyError("\"" + findText + "\" not found in that shape's current text.");
        var sub = range.getSubstring(idx, String(findText).length);
        if (font.bold != null) sub.font.bold = font.bold;
        if (font.italic != null) sub.font.italic = font.italic;
        if (font.allCaps != null) sub.font.allCaps = font.allCaps;
        if (font.underline != null) sub.font.underline = font.underline;
        if (font.color) sub.font.color = font.color;
        if (font.name) sub.font.name = font.name;
        if (font.size) sub.font.size = font.size;
        return ctx.sync().then(function () { return { ok: true }; });
        });
      });
    });
  };

OfficeDeck.prototype.setBackground = function (slideIndex, payload) {
    return this.run("setBackground", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        throw new FriendlyError("slide backgrounds need PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac) or use a solid design instead.");
      }
      var bg = ctx.presentation.slides.getItemAt(slideIndex).background;
      var fill = bg.fill;
      if (payload.kind === "image") {
        var opts = { imageBase64: String(payload.b64) };
        if (payload.transparency != null) opts.transparency = payload.transparency;
        fill.setPictureOrTextureFill(opts);
      } else if (payload.kind === "solid") {
        var sopts = { color: payload.color };
        if (payload.transparency != null) sopts.transparency = payload.transparency;
        fill.setSolidFill(sopts);
      } else {
        fill.setGradientFill({});
      }
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

OfficeDeck.prototype.addHyperlink = function (slideIndex, shapeIndex, url, tip) {
    return this.run("addHyperlink", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        throw new FriendlyError("hyperlinks need PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
      }
      var shape = ctx.presentation.slides.getItemAt(slideIndex).shapes.getItemAt(shapeIndex);
      var opts = { address: url };
      if (tip) opts.screenTip = tip;
      shape.setHyperlink(opts);
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

OfficeDeck.prototype.setAltText = function (slideIndex, shapeIndex, description, title) {
    return this.run("setAltText", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        throw new FriendlyError("alt text needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
      }
      var shape = ctx.presentation.slides.getItemAt(slideIndex).shapes.getItemAt(shapeIndex);
      shape.altTextDescription = description;
      if (title) shape.altTextTitle = title;
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

OfficeDeck.prototype.getSelectionInfo = function () {
    // selection -> { slides: [idx], shapes: [{slide,shape}], textPreview? } or null on any failure
    return this.run("getSelectionInfo", function (ctx) {
      var p = ctx.presentation;
      if (!p.getSelectedSlides || !p.getSelectedShapes) return Promise.resolve(null);
      var slidesColl, shapesColl, textRange = null;
      try {
        slidesColl = p.getSelectedSlides();
        shapesColl = p.getSelectedShapes();
        try { textRange = p.getSelectedTextRange(); } catch (e) { textRange = null; }
      } catch (e) { return Promise.resolve(null); }
      var allSlides = ctx.presentation.slides;
      slidesColl.load("id"); shapesColl.load("id"); allSlides.load("items/id");
      return ctx.sync().then(function () {
        var slideIdxOf = {};
        (allSlides.items || []).forEach(function (sl, idx) { slideIdxOf[sl.id] = idx; });
        var out = { slides: [], shapes: [], textPreview: "" };
        var selShapes = shapesColl.items;
        if (!selShapes.length) {
          out.slides = slidesColl.items
            .map(function (x) { return slideIdxOf[x.id]; })
            .filter(function (v) { return v != null; });
          if (!textRange) return out;
          return textRange.load("isNullObject,text");
          return ctx.sync().then(function () {
            if (!textRange.isNullObject) out.textPreview = String(textRange.text || "").slice(0, 200);
            return out;
          });
        }
        // map each selected shape to its slide index + shape index (by ids)
        selShapes.forEach(function (sh) { sh.slide.load("id"); });
        return ctx.sync().then(function () {
          var bySlide = {};   // slideObjId -> array of selected shapeIds
          selShapes.forEach(function (sh) {
            var sid = sh.slide.id;
            (bySlide[sid] = bySlide[sid] || []).push(sh.id);
          });
          var slideObjs = {};
          selShapes.forEach(function (sh) {
            if (!slideObjs[sh.slide.id]) {
              slideObjs[sh.slide.id] = sh.slide.shapes;
              sh.slide.shapes.load("items/id");
            }
          });
          return ctx.sync().then(function () {
            Object.keys(slideObjs).forEach(function (sid) {
              var sIdx = slideIdxOf[sid];
              var items = slideObjs[sid].items || [];
              var idToIdx = {};
              items.forEach(function (x, k) { idToIdx[x.id] = k; });
              (bySlide[sid] || []).forEach(function (shapeId) {
                if (sIdx != null && idToIdx[shapeId] != null) out.shapes.push({ slide: sIdx, shape: idToIdx[shapeId] });
              });
            });
            uniqueSorted(out.slides, slideIdxOf, slideObjs, out);
            if (!textRange) return out;
            return textRange.load("isNullObject,text");
          return ctx.sync().then(function () {
              if (!textRange.isNullObject) out.textPreview = String(textRange.text || "").slice(0, 200);
              return out;
            });
          });
        });
      }).catch(function () { return null; });
    });
  };

  function uniqueSorted(slidesOut, slideIdxOf, slideObjs, out) {
    var set = {};
    Object.keys(slideObjs).forEach(function (sid) {
      var v = slideIdxOf[sid];
      if (v != null) set[v] = true;
    });
    (out.slides = []).length = 0;
    out.slides = Object.keys(set).map(Number).sort(function (a, b) { return a - b; });
  }

OfficeDeck.prototype.formatText = function (slideIndex, shapeIndex, font) {
    return this.run("formatText", function (ctx) {
      var shapesX = ctx.presentation.slides.getItemAt(slideIndex).shapes;
      shapesX.load("items/type");
      return ctx.sync().then(function () {
        if (shapeIndex >= shapesX.items.length) {
          throw new FriendlyError("shape " + (shapeIndex + 1) + " does not exist on slide " + (slideIndex + 1) +
            " (the slide has " + shapesX.items.length + " shape(s)) - the action probably targets a shape" +
            " that was not created or was removed; create the shape first, then format it.");
        }
        var shape = shapesX.getItemAt(shapeIndex);
        shape.load("type");
        return ctx.sync().then(function () {
          var ty = String(shape.type || "");
          return _formatTextContinue(ctx, shape, shapeIndex, slideIndex, font);
        });
      });
    });
  };
  function _formatTextContinue(ctx, shape, shapeIndex, slideIndex, font) {
        if (/^(table|chart|picture|media|ole|graphic|unsupported)/i.test(tyTemp(shape))) {
          throw new FriendlyError("shape " + (shapeIndex + 1) + " on slide " + (slideIndex + 1) +
            " has no text to format (it is a " + tyTemp(shape).toLowerCase() + " shape) - format_text needs a text shape (textbox or placeholder); a table\'s text lives in its cells.");
        }
        var range = shape.textFrame.textRange;
        if (font.bold != null) range.font.bold = font.bold;
        if (font.italic != null) range.font.italic = font.italic;
        if (font.allCaps != null) range.font.allCaps = font.allCaps;
        if (font.name) range.font.name = font.name;
        if (font.color) range.font.color = font.color;
        if (font.size) range.font.size = font.size;
        if (font.underline != null) range.font.underline = font.underline;
        if (font.align) range.paragraphFormat.horizontalAlignment = font.align.charAt(0).toUpperCase() + font.align.slice(1);
        if (font.bullet != null) range.paragraphFormat.bulletFormat.visible = font.bullet;
        return ctx.sync().then(function () { return { ok: true }; });
  }
  function tyTemp(shape) { return String(shape.type || ""); }

OfficeDeck.prototype.setSlideTitle = function (i, text) { return this._setPlaceholderText(i, "title", text); };
  OfficeDeck.prototype.setSlideBody = function (i, lines) { return this._setPlaceholderText(i, "body", lines); };

  OfficeDeck.prototype._setPlaceholderText = function (i, role, lines) {
    var self = this;
    var joined = Array.isArray(lines) ? lines.join("\r") : String(lines == null ? "" : lines);
    return this.run("set" + role, function (ctx) {
      var slide = ctx.presentation.slides.getItemAt(i);
      var shapes = slide.shapes;
      shapes.load("items/type,items/name,items/id");
      return ctx.sync().then(function () {
        var metas = [];
        var phList = [];
        var wantPf = self._pfApi === true;
        shapes.items.forEach(function (sh, j) {
          metas.push({ sh: sh, name: sh.name || "", type: sh.type || "", ptype: undefined, tf: null });
          var tf = sh.getTextFrameOrNullObject();
          tf.load("isNullObject,hasText");
          metas[j].tf = tf;
          if (wantPf && sh.type === "placeholder") {
            sh.load("placeholderFormat");
            phList.push({ sh: sh, idx: j });
          }
        });
        return ctx.sync().then(function () {
          phList.forEach(function (e) { e.sh.placeholderFormat.load("type"); });
          var afterPf = wantPf ? ctx.sync() : Promise.resolve();
          return afterPf.then(function () {
            phList.forEach(function (e) { metas[e.idx].ptype = e.sh.placeholderFormat.type; });
            var pick = null;
            metas.forEach(function (m) {
              if (m.tf.isNullObject) return;
              if (role === "title") {
                if (!pick && isTitleShape(m.name, m.type, m.ptype)) pick = m;
              } else {
                if (!pick && isBodyShape(m.name, m.type, m.ptype)) pick = m;
              }
            });
            if (!pick) {
              if (role === "title") {
                throw new FriendlyError("No title placeholder found on slide " + i);
              }
              // 0.9.287: title-only/blank layouts have no body placeholder -
              // write the content into a NEW TEXT BOX instead of failing the
              // op (field: story slides on a fresh build were all title
              // slides; five body writes failed wholesale and the retry
              // repeated them). Geometry: the standard body band per the
              // make-over grid (960x540). Two creation forms - some builds
              // reject ShapeAddOptions at sync, then accept the bare form
              // with geometry as properties (the 0.9.233/0.9.261 repair).
              var std = { left: 48, top: 190, width: 864, height: 300 };
              var fallbackNote = "no body placeholder (slide " + (i + 1) + ") - content set in a new text box";
              try {
                var tb = slide.shapes.addTextBox(joined, std);
                try { tb.name = "sk-body-fallback-" + (i + 1); } catch (eN2) { /* optional */ }
                return ctx.sync().then(function () { return { ok: true, how: fallbackNote }; });
              } catch (eTb1) {
                try {
                  var tb2 = slide.shapes.addTextBox(joined);
                  ["left", "top", "width", "height"].forEach(function (gk) {
                    try { tb2[gk] = std[gk]; } catch (eG4) { /* optional */ }
                  });
                  try { tb2.name = "sk-body-fallback-" + (i + 1); } catch (eN3) { /* optional */ }
                  return ctx.sync().then(function () { return { ok: true, how: fallbackNote }; });
                } catch (eTb2) {
                  throw new FriendlyError("No body/content placeholder found on slide " +
                    i + ", and the text-box fallback failed (" +
                    String((eTb1 && eTb1.message) || eTb1).slice(0, 80) + ")");
                }
              }
            }
            if (!pick.tf.hasText && joined === "") return null;
            pick.tf.textRange.load("text");
            return ctx.sync().then(function () {
              pick.tf.textRange.text = joined;
              return ctx.sync();
            });
          });
        });
      });
    });
  };

  OfficeDeck.prototype.setShapeText = function (i, shapeIdx, text) {
    return this.run("setShapeText", function (ctx) {
      var shape = ctx.presentation.slides.getItemAt(i).shapes.getItemAt(shapeIdx);
      var tf = shape.getTextFrameOrNullObject();
      tf.load("isNullObject,hasText");
      return ctx.sync().then(function () {
        if (tf.isNullObject) {
          throw new FriendlyError("Shape " + shapeIdx + " on slide " + i + " has no text frame");
        }
        tf.textRange.load("text");
        return ctx.sync().then(function () {
          tf.textRange.text = String(text == null ? "" : text);
          return ctx.sync();
        });
      });
    });
  };

  OfficeDeck.prototype.setShapeTextFind = function (slideIndex, shapeIndex, findText, replText) {
    return this.run("setShapeTextFind", function (ctx) {
      var shape = ctx.presentation.slides.getItemAt(slideIndex).shapes.getItemAt(shapeIndex);
      var tf = shape.getTextFrameOrNullObject();
      tf.load("isNullObject,hasText");
      return ctx.sync().then(function () {
        if (tf.isNullObject || !tf.hasText) {
          throw new FriendlyError("Shape " + shapeIndex + " on slide " + slideIndex + " has no text");
        }
        var range = tf.textRange;
        range.load("text");
        return ctx.sync().then(function () {
          var full = String(range.text || "");
          var needle = String(findText);
          var idx = full.toLowerCase().indexOf(needle.toLowerCase());
          if (idx < 0) {
            throw new FriendlyError("\"" + findText + "\" not found in that shape's current text.");
          }
          var repl = String(replText == null ? "" : replText);
          try {
            var sub = range.getSubstring(idx, needle.length);
            sub.text = repl;
          } catch (eSub) { /* older host without getSubstring set - fall back below */ }
          return ctx.sync().then(function () {
            range.load("text");
            return ctx.sync().then(function () {
              var now = String(range.text || "");
              if (now.toLowerCase().indexOf(repl.toLowerCase()) >= 0) {
                return { ok: true, mode: "fragment" };
              }
              // host ignored the subrange write -> rewrite the full text ourselves
              var rebuilt = full.slice(0, idx) + repl + full.slice(idx + needle.length);
              range.text = rebuilt;
              return ctx.sync().then(function () { return { ok: true, mode: "full" }; });
            });
          });
        });
      });
    });
  };

  OfficeDeck.prototype.mergeDeck = function (b64) {
    // PowerPointApi 1.2: insert ALL slides of a .pptx after the current last slide
    return this.run("mergeDeck", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.2")) {
        throw new FriendlyError("merging decks needs PowerPointApi 1.2 - update PowerPoint.");
      }
      var slides = ctx.presentation.slides;
      slides.load("items/id");
      return ctx.sync().then(function () {
        var opts = { formatting: "KeepSourceFormatting" };
        if (slides.items.length) {
          opts.targetSlideId = slides.items[slides.items.length - 1].id;  // append at the end
        }
        ctx.presentation.insertSlidesFromBase64(String(b64), opts);
        return ctx.sync().then(function () { return { ok: true }; });
      });
    });
  };

  OfficeDeck.prototype.styleShape = function (slideIndex, shapeIndex, fill, line, transparency) {
    return this.run("styleShape", function (ctx) {
      // PowerPointApi 1.4: ShapeFill.setSolidColor + ShapeLineFormat; transparency: 1.10
      var shape = ctx.presentation.slides.getItemAt(slideIndex).shapes.getItemAt(shapeIndex);
      if (fill) {
        shape.fill.setSolidColor(String(fill));
      }
      if (line) {
        shape.lineFormat.visible = true;
        shape.lineFormat.color = String(line);
      }
      if (transparency != null) {
        if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
          throw new FriendlyError("fill transparency needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
        }
        shape.fill.transparency = Number(transparency);
      }
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

  OfficeDeck.prototype.setShapeImage = function (slideIndex, shapeIndex, b64, transparency) {
    return this.run("setShapeImage", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        throw new FriendlyError("picture-filled shapes need PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
      }
      var shape = ctx.presentation.slides.getItemAt(slideIndex).shapes.getItemAt(shapeIndex);
      shape.fill.setImage(String(b64));
      if (transparency != null) shape.fill.transparency = Number(transparency);
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

OfficeDeck.prototype.setNotes = function (i, text) {
    var self = this;
    return this.run("setNotes", function (ctx) {
      var slide = ctx.presentation.slides.getItemAt(i);
      if (!slide.notesPage) {
        throw new FriendlyError("Speaker notes are not editable via the API in this Office build");
      }
      var shapes = slide.notesPage.shapes;
      shapes.load("items/type,items/name,items/id");
      return ctx.sync().then(function () {
        var metas = [];
        var phList = [];
        var wantPf = self._pfApi === true;
        shapes.items.forEach(function (sh, j) {
          metas.push({ sh: sh, name: sh.name || "", type: sh.type || "", ptype: undefined, tf: null });
          var tf = sh.getTextFrameOrNullObject();
          tf.load("isNullObject,hasText");
          metas[j].tf = tf;
          if (wantPf && sh.type === "placeholder") {
            sh.load("placeholderFormat");
            phList.push({ sh: sh, idx: j });
          }
        });
        return ctx.sync().then(function () {
          phList.forEach(function (e) { e.sh.placeholderFormat.load("type"); });
          var afterPf = wantPf ? ctx.sync() : Promise.resolve();
          return afterPf.then(function () {
            phList.forEach(function (e) { metas[e.idx].ptype = e.sh.placeholderFormat.type; });
            var pick = null;
            metas.forEach(function (m) {
              if (m.tf.isNullObject) return;
              if (isBodyShape(m.name, m.type, m.ptype) && !pick) pick = m; // notes body is a Body placeholder
              else if (!pick && /body|notes|content|text/i.test(m.name)) pick = m;
              else if (!pick && m.tf.hasText) pick = m;
              else if (!pick && !m.ptype && !/slide|date|footer|header|number/i.test(m.name)) pick = m;
            });
            if (!pick) {
              return { manual: true, reason: "no notes body found on slide " + i };
            }
            pick.tf.textRange.load("text");
            return ctx.sync().then(function () {
              pick.tf.textRange.text = String(text == null ? "" : text);
              return ctx.sync().then(function () { return { manual: false }; });
            });
          });
        });
      });
    }).catch(function (err) {
      if (err && err.friendly) return { manual: true, reason: err.message };
      throw err;
    });
  };

  OfficeDeck.prototype.deleteSlide = function (i) {
    return this.run("deleteSlide", function (ctx) {
      var slides = ctx.presentation.slides;
      if (typeof slides.removeAt !== "function") {
        throw new FriendlyError("This PowerPoint build cannot delete slides via the add-in API");
      }
      slides.removeAt(i);
      return ctx.sync();
    });
  };

  // Full text read of specific 1-based slides (read_slide): every shape
  // with its complete text, the same shape numbering (1-based) the outline
  // uses. Unlike getContext it does not summarize - it returns the raw
  // per-shape text so the model can quote/rewrite the slide accurately.
  OfficeDeck.prototype.readSlides = function (ones) {
    var self = this;
    if (self._textApi === false)
      return Promise.reject(new FriendlyError("Office APIs are too old in this build to read shape text"));
    var list = [];
    (ones || []).forEach(function (n) {
      var i = Math.floor(Number(n));
      if (i >= 1 && list.indexOf(i) < 0) list.push(i);
    });
    if (!list.length) return Promise.reject(new FriendlyError("readSlides: no slide numbers given"));
    return this.run("readSlides", function (ctx) {
      var slides = ctx.presentation.slides;
      slides.load("items/id");
      return ctx.sync().then(function () {
        var targets = [];
        list.forEach(function (n) {
          if (n > slides.items.length)
            throw new FriendlyError("Slide " + n + " does not exist (deck has " + slides.items.length + " slide(s))");
          targets.push({ n: n, sl: slides.items[n - 1], shapes: [] });
        });
        targets.forEach(function (t) {
          t.sl.shapes.load("items/type,items/name,items/id");
        });
        return ctx.sync().then(function () {
          var tfs = [];
          targets.forEach(function (t) {
            var shpProxies = t.sl.shapes.items;
            for (var j = 0; j < shpProxies.length; j++) {
              var sh = shpProxies[j];
              var entry = { idx: j, name: sh.name || "", type: sh.type || "", text: null };
              t.shapes.push(entry);
              var tf = sh.getTextFrameOrNullObject();
              tf.load("isNullObject,hasText");
              tfs.push({ tf: tf, entry: entry });
            }
          });
          return ctx.sync().then(function () {
            var trs = [];
            tfs.forEach(function (e) {
              if (!e.tf.isNullObject && e.tf.hasText) {
                e.tf.textRange.load("text");
                trs.push(e);
              }
            });
            return ctx.sync().then(function () {
              trs.forEach(function (e) { e.entry.text = e.tf.textRange.text || ""; });
              // tables have no textFrame: read their cells via the table API
              var tblEntries = [];
              targets.forEach(function (t2) {
                t2.shapes.forEach(function (sh) { if (sh.type === "table") tblEntries.push(sh); });
              });
              var tblApi2 = true;
              try { tblApi2 = Office.context.requirements.isSetSupported("PowerPointApi", "1.8"); } catch (eTbl2) { tblApi2 = false; }
              var tblTries = tblEntries.map(function (sh) {
                if (!tblApi2) return Promise.resolve("[table - this Office build does not expose table cell APIs to add-ins (needs the PowerPoint table API, set 1.8) - update Office, then ask again]");
                return self._readTableText(ctx, sh)
                  .then(function (info) {
                    return info && info.cells && info.cells.length
                      ? "[table " + info.r + "x" + info.c + "] " + info.cells.join(" | ")
                      : "[table " + info.r + "x" + info.c + " - unreadable]";
                  })
                  .catch(function (eT) {
                    var why = eT && (eT.message || String(eT));
                    if (eT && eT.noTableProp) why = eT.message;
                    dbg("table.read", { failed: true, why: why });
                    return "[table - cell text unavailable in this Office build (" + why + ")]";
                  });
              });
              return Promise.all(tblTries).then(function (vals) {
                tblEntries.forEach(function (sh, k) { sh.text = vals[k]; });
                var out = [];
                targets.forEach(function (t) {
                  out.push("SLIDE " + t.n + " (slide " + t.n + " in the outline):");
                  if (!t.shapes.length) { out.push("  (no shapes)"); return; }
                  t.shapes.forEach(function (sh) {
                    var txt = sh.text == null ? "(text unavailable in this Office build)" : String(sh.text);
                    txt = txt.replace(/\r|\n/g, " | ");
                    out.push("  shape " + (sh.idx + 1) + " (" + sh.name + ", " + sh.type + '): "' + txt + '"');
                  });
                });
                var total = out.join("\n");
                if (total.length > 60000) total = total.slice(0, 60000) + "\n\u2026[truncated]";
                return total;
              });            // Promise.all(tblTries)
            });              // ctx.sync (textRange)
          });                // ctx.sync (tfs)
        });                  // ctx.sync (shapes)
      });                    // ctx.sync (ids)
    });                      // run callback
  };


  OfficeDeck.prototype.applyLayout = function (i, kind) {
    if (this._textApi === false) {
      return Promise.reject(new FriendlyError("Office APIs are too old in this build to modify slides"));
    }
    return this.run("applyLayout", function (ctx) {
      var slide = ctx.presentation.slides.getItemAt(i);
      var layouts = slide.slideMaster.layouts;
      layouts.load("items/id,items/name");
      return ctx.sync().then(function () {
        var items = (layouts.items || []).map(function (l) { return { id: l.id, name: l.name }; });
        var best = null, bestScore = 0;
        items.forEach(function (l) {
          var s = scoreLayout(l.name, normalizeLayout(kind));
          if (s > bestScore) { bestScore = s; best = l; }
        });
        if (!best) throw new FriendlyError("No layout matching '" + kind + "' found in this deck");
        var lay = layouts.getItemOrNullObject(best.id);
        slide.applyLayout(lay);
        return ctx.sync();
      });
    });
  };

  function normalizeLayout(kind) {
    var k = String(kind || "").toLowerCase();
    if (k.indexOf("content") >= 0) return "title_content";
    if (k.indexOf("only") >= 0) return "title_only";
    if (k.indexOf("section") >= 0) return "section";
    if (k === "title_content" || k === "title_only" || k === "section" || k === "blank") return k;
    return "blank";
  }

  function isTitleShape(name, type, ptype) {
    // placeholderFormat.type (PowerPointApi 1.8) wins: localization-safe on all hosts
    var t = String(ptype || "").toLowerCase();
    if (t) return t === "title" || t === "centertitle";
    return /placeholder/i.test(String(type || "")) && /title/i.test(String(name || ""));
  }
  function isBodyShape(name, type, ptype) {
    var t = String(ptype || "").toLowerCase();
    if (t) return t === "body" || t === "subtitle" || t === "content" || t === "verticalbody";
    var n = String(name || ""), ty = String(type || "");
    if (!/placeholder|textbox/i.test(ty)) return false;
    if (/title/i.test(n)) return false;
    return /content|text|body|subtitle/i.test(n);
  }

  // ------------------------------------------------------------------ //
  // Ops validation / approvals
  // ------------------------------------------------------------------ //
  function validateOps(rawOps, ctx, preInvalid) {
    var ops = [], invalid = [];
    var notesDropped = 0;
    var notesOff = !(SK.cfg && SK.cfg.generateNotes);
    // slides the batch itself creates (create_slide ops processed before the
    // op being validated) count as valid targets: the validator sees the deck
    // as it will be AFTER those creates, not only as it is right now
    var virtualCount = ctx ? ctx.slideCount : 0;
    // shapes the batch itself adds to an EXISTING slide (create_table /
    // create_shape / create_image) shift the post-batch shape numbering:
    // format/move ops may target those fresh shapes (declared OUTSIDE the
    // op loop - a per-op accumulator would reset on every op)
    var plannedShapes = {};
    // text mutations the batch itself makes (replace_text earlier in the
    // batch changes what a later op's "find" matches at apply time - the
    // snapshot only knows the pre-batch text). key = "slide:shape".
    // (accumulator rule: lives at validateOps scope, NOT per-op)
    var textState = {};
    (rawOps || []).forEach(function (op, index) {
      if (!op || typeof op !== "object") { invalid.push({ op: op, reason: "not an object" }); return; }
      var t = op.type;
      if (OP_TYPES.indexOf(t) < 0) { invalid.push({ op: op, reason: "unknown op type '" + t + "'" }); return; }
      if (!ctx && t !== "create_slide") {
        invalid.push({ op: op, reason: "needs deck context (enable 'include deck context')" });
        return;
      }
      var bad = function (reason) { invalid.push({ op: op, reason: reason }); };
      var slideOk = function (v) { return typeof v === "number" && Math.floor(v) === v && v >= 1 && (!ctx || v <= virtualCount); };
      // a slide the CURRENT deck does not have yet = created later in this
      // batch: shape/find checks cannot use the stored context - defer them
      // to apply time (a wrong shape number fails honestly there)
      var futureSlide = function (v) { return ctx && typeof v === "number" && v > ctx.slideCount; };
      var shapeOk = function (sl, sh) {
        if (!ctx || !ctx.slides[sl - 1]) return false;
        return (ctx.slides[sl - 1].shapes || []).some(function (s) { return s.idx === sh - 1 && s.text != null; });
      };
      // post-batch shape targets: counts the fresh shapes the batch added
      // (plannedShapes lives at validateOps scope, see above)
      var shapeIdxOk = function (sl, sh) {
        if (typeof sh !== "number" || sh < 1 || !ctx || !ctx.slides[sl - 1]) return false;
        var cs = ctx.slides[sl - 1].shapes;
        if (!cs || !cs.length) return true;   // ctx does not detail this slide (not selected) - defer to apply time
        return sh <= cs.length + (plannedShapes[sl] || 0);
      };
      // best-known text of a shape: what earlier batch ops left there,
      // else the pre-batch snapshot; null = unknown (defer find to apply)
      var textOf = function (sl, sh) {
        var k2 = sl + ":" + sh;
        if (textState[k2] !== undefined) return textState[k2];
        var snap = ((ctx && ctx.slides[sl - 1] && ctx.slides[sl - 1].shapes) || []).filter(function (s2) { return s2.idx === sh - 1; })[0];
        return (snap && typeof snap.text === "string") ? snap.text : null;
      };
      var cutReplace = function (base, findS, repS) {
        var i2 = base.toLowerCase().indexOf(findS.toLowerCase());
        if (i2 < 0) return null;
        return base.slice(0, i2) + repS + base.slice(i2 + findS.length);
      };
      var out = { type: t };
      if (t === "create_slide") {
        out.layout = normalizeLayout(op.layout);
        out.title = typeof op.title === "string" ? op.title : "";
        var body = [];
        if (Array.isArray(op.body)) {
          op.body.forEach(function (b) { if (b != null) body.push(String(b)); });
        } else if (typeof op.body === "string" && op.body.length) {
          body = op.body.split(/\r|\n/);
        }
        out.body = body;
        out.notes = notesOff ? null : (typeof op.notes === "string" ? op.notes : null);
        out.after = typeof op.after === "number" ? op.after : null;
        virtualCount += 1;   // later ops in this batch may target this slide
      } else if (t === "edit_slide") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = op.slide;
        out.title = typeof op.title === "string" ? op.title : null;
        var body2 = null;
        if (Array.isArray(op.body)) { body2 = []; op.body.forEach(function (b) { if (b != null) body2.push(String(b)); }); }
        else if (typeof op.body === "string" && op.body.length) body2 = op.body.split(/\r|\n/);
        out.body = body2;
        out.notes = notesOff ? null : (typeof op.notes === "string" ? op.notes : null);
        if (out.title === null && out.body === null && out.notes === null) { bad("nothing to change (all fields null)"); return; }
      } else if (t === "replace_text") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        if (!shapeOk(op.slide, op.shape)) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        if (typeof op.text !== "string" || !op.text.length) { bad("missing text"); return; }
        out.slide = op.slide; out.shape = op.shape; out.text = op.text;
        if (typeof op.find === "string" && op.find.trim()) {
          var findR = op.find.trim();
          if (findR.length > 200) { bad("find text too long (200 chars max)"); return; }
          var shapeTxtR = textOf(op.slide, op.shape);   // earlier batch ops may already have rewritten this shape
        if (!futureSlide(op.slide) && shapeTxtR !== null && shapeTxtR.toLowerCase().indexOf(findR.toLowerCase()) === -1) {
          bad("find text not found in that shape's current text"); return;
        }
          out.find = findR;
        }
        // remember what THIS op leaves in the shape for later find checks
        var curTr = textOf(op.slide, op.shape);
        textState[op.slide + ":" + op.shape] = out.find ?
          (typeof curTr === "string" ? (cutReplace(curTr, findR, op.text)) : null) : op.text;
      } else if (t === "set_notes") {
        // notes-off no longer drops the op: it goes through as "confirm"
        // (renderOpsPreview forces confirm when generateNotes is off)
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        if (typeof op.notes !== "string" || !op.notes) { bad("missing notes"); return; }
        out.slide = op.slide; out.notes = op.notes;
      } else if (t === "create_table") {
        var vals = Array.isArray(op.values) ? op.values : null;
        if (!vals) { bad("missing values (2D array)"); return; }
        var rowsT = [];
        var cols = 0;
        for (var vi = 0; vi < vals.length; vi++) {
          var rowT = vals[vi];
          if (!Array.isArray(rowT)) { bad("values must be an array of arrays"); return; }
          if (vi === 0) cols = rowT.length;
          if (rowT.length !== cols || cols < 1) { bad("values rows must be rectangular"); return; }
          rowsT.push(rowT.map(function (c) { return c == null ? "" : String(c); }));
        }
        if (op.slide != null && !slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = typeof op.slide === "number" ? op.slide : null;
        out.title = typeof op.title === "string" ? op.title : "";
        out.values = rowsT;
        var geoT = {}, srcG = op.geometry && typeof op.geometry === "object" ? op.geometry : null;
        ["left", "top", "width", "height"].forEach(function (p2) {
          var v = srcG ? srcG[p2] : op[p2];
          if (typeof v === "number" && v >= 0) geoT[p2] = v;
        });
        var rotV = srcG && srcG.rotation != null ? srcG.rotation : op.rotation;
        if (typeof rotV === "number" && rotV >= -360 && rotV <= 360) geoT.rotation = rotV;
        out.geometry = geoT;
      } else if (t === "move_shape") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        if (!futureSlide(op.slide) && !shapeIdxOk(op.slide, op.shape)) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        var anyM = false, geoM = {};
        ["left", "top", "width", "height"].forEach(function (p3) {
          if (typeof op[p3] === "number") { geoM[p3] = op[p3]; anyM = true; }
        });
        if (typeof op.rotation === "number" && op.rotation >= -360 && op.rotation <= 360) { geoM.rotation = op.rotation; anyM = true; }
        else if (op.rotation != null) { bad("rotation must be a number between -360 and 360"); return; }
        if (!anyM) { bad("nothing to change (no left/top/width/height/rotation)"); return; }
        out.slide = op.slide; out.shape = op.shape; out.geometry = geoM;
      } else if (t === "create_image") {
        if (typeof op.prompt !== "string" || op.prompt.trim().length < 3) { bad("missing image prompt"); return; }
        if (op.prompt.length > 4000) { bad("image prompt too long (max 4000 chars)"); return; }
        if (op.slide != null && !slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = typeof op.slide === "number" ? op.slide : null;
        out.prompt = op.prompt.trim();
        var geoI = {};
        ["left", "top", "width", "height"].forEach(function (p2) {
          if (typeof op[p2] === "number" && op[p2] >= 0) geoI[p2] = op[p2];
        });
        out.geometry = geoI;
      } else if (t === "format_text") {
        if (!slideOk(op.slide)) {
          bad(typeof op.slide === "number"
            ? "bad slide index: slide " + op.slide + " does not exist (deck has " + (ctx ? ctx.slideCount : 0) + " slides)"
            : "bad slide index: slide must be a number (1-" + (ctx ? ctx.slideCount : 0) + ")");
          return;
        }
        if (!futureSlide(op.slide) && !shapeIdxOk(op.slide, op.shape)) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        var fontF = {}, hasAny = false;
        ["bold", "italic", "allCaps", "underline"].forEach(function (k2) {
          if (typeof op.font !== "undefined" && typeof op.font[k2] === "boolean") { fontF[k2] = op.font[k2]; hasAny = true; }
        });
        if (op.font && typeof op.font.color === "string" && /^[\w#\(\), .%]+$/.test(op.font.color)) { fontF.color = op.font.color; hasAny = true; }
        if (op.font && typeof op.font.name === "string" && op.font.name) { fontF.name = op.font.name; hasAny = true; }
        if (op.font && typeof op.font.size === "number" && op.font.size >= 6 && op.font.size <= 400) { fontF.size = op.font.size; hasAny = true; }
        if (typeof op.align === "string" && /^(left|center|right)$/i.test(op.align)) { fontF.align = op.align.toLowerCase(); hasAny = true; }
        if (typeof op.bullet === "boolean") { fontF.bullet = op.bullet; hasAny = true; }
        if (!hasAny) { bad("no format changes requested"); return; }
        if (op.font && typeof op.font !== "object") { bad("font must be an object"); return; }
        out.slide = op.slide; out.shape = op.shape; out.font = fontF;
        if (typeof op.find === "string" && op.find.trim()) {
          var findS = op.find.trim();
          if (findS.length > 200) { bad("find text too long (200 chars max)"); return; }
          var shapeTxtF = textOf(op.slide, op.shape);   // earlier batch ops may already have rewritten this shape
        if (!futureSlide(op.slide) && shapeTxtF !== null && shapeTxtF.toLowerCase().indexOf(findS.toLowerCase()) === -1) {
          bad("find text not found in that shape's current text"); return;
        }
          out.find = findS;
        }
      } else if (t === "format_selection") {
        var fontS = {}, hasAnyS = false;
        ["bold", "italic", "allCaps", "underline"].forEach(function (k4) {
          if (typeof op.font !== "undefined" && typeof op.font[k4] === "boolean") { fontS[k4] = op.font[k4]; hasAnyS = true; }
        });
        if (op.font && typeof op.font.color === "string" && /^[\w#\(\), .%]+$/.test(op.font.color)) { fontS.color = op.font.color; hasAnyS = true; }
        if (op.font && typeof op.font.name === "string" && op.font.name) { fontS.name = op.font.name; hasAnyS = true; }
        if (op.font && typeof op.font.size === "number" && op.font.size >= 6 && op.font.size <= 400) { fontS.size = op.font.size; hasAnyS = true; }
        if (!hasAnyS) { bad("no format changes requested"); return; }
        out.font = fontS;
      } else if (t === "style_shape") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        var shapeInCtxS = ctx && ctx.slides[op.slide - 1] &&
          (ctx.slides[op.slide - 1].shapes || []).some(function (s) { return s.idx === op.shape - 1; });
        if (!futureSlide(op.slide) && !shapeInCtxS) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        var fillS = null, lineS = null;
        if (typeof op.fill === "string" && op.fill.trim()) {
          if (!/^[\w#\(\), .%]{1,40}$/.test(op.fill.trim())) { bad("bad fill color"); return; }
          fillS = op.fill.trim();
        }
        if (typeof op.line === "string" && op.line.trim()) {
          if (!/^[\w#\(\), .%]{1,40}$/.test(op.line.trim())) { bad("bad line color"); return; }
          lineS = op.line.trim();
        }
        var transS = (typeof op.transparency === "number" && op.transparency >= 0 && op.transparency <= 1) ? op.transparency : null;
        if (!fillS && !lineS && transS === null) { bad("fill or line color (or transparency) required"); return; }
        out.slide = op.slide; out.shape = op.shape;
        if (fillS) out.fill = fillS;
        if (lineS) out.line = lineS;
        if (transS !== null) out.transparency = transS;
      } else if (t === "set_background") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        if (op.kind !== "image" && op.kind !== "solid" && op.kind !== "gradient") { bad("kind must be image, solid or gradient"); return; }
        if (op.kind === "image") {
          if (typeof op.prompt !== "string" || !op.prompt.trim()) { bad("image background needs a prompt"); return; }
          out.prompt = op.prompt.trim().slice(0, 600);
        } else if (op.kind === "solid") {
          if (typeof op.color !== "string" || !/^[\w#]+$/.test(op.color)) { bad("solid background needs a color"); return; }
          out.color = op.color;
        }
        if (typeof op.transparency === "number" && op.transparency >= 0 && op.transparency <= 1) out.transparency = op.transparency;
        out.slide = op.slide; out.kind = op.kind;
      } else if (t === "set_shape_image") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        var shapeInCtxI = ctx && ctx.slides[op.slide - 1] &&
          (ctx.slides[op.slide - 1].shapes || []).some(function (sh) { return sh.idx === op.shape - 1; });
        if (!futureSlide(op.slide) && !shapeInCtxI) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        if (typeof op.prompt !== "string" || op.prompt.trim().length < 3) { bad("missing image prompt"); return; }
        if (op.prompt.length > 4000) { bad("image prompt too long (max 4000 chars)"); return; }
        out.slide = op.slide; out.shape = op.shape; out.prompt = op.prompt.trim();
        if (typeof op.transparency === "number" && op.transparency >= 0 && op.transparency <= 1) out.transparency = op.transparency;
      } else if (t === "add_hyperlink") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        var shapeInCtxH = ctx && ctx.slides[op.slide - 1] &&
          (ctx.slides[op.slide - 1].shapes || []).some(function (sh) { return sh.idx === op.shape - 1; });
        if (!futureSlide(op.slide) && !shapeInCtxH) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        if (typeof op.url !== "string" || !/^(https?:\/\/\S+|mailto:\S+)$/.test(op.url.trim())) { bad("url must be http(s)://... or mailto:..."); return; }
        out.slide = op.slide; out.shape = op.shape; out.url = op.url.trim();
        if (typeof op.tip === "string" && op.tip.trim()) out.tip = op.tip.trim().slice(0, 200);
      } else if (t === "set_alt_text") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        var shapeInCtxA = ctx && ctx.slides[op.slide - 1] &&
          (ctx.slides[op.slide - 1].shapes || []).some(function (sh) { return sh.idx === op.shape - 1; });
        if (!futureSlide(op.slide) && !shapeInCtxA) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        if (typeof op.description !== "string" || !op.description.trim()) { bad("alt text needs a description"); return; }
        out.slide = op.slide; out.shape = op.shape; out.description = op.description.trim().slice(0, 500);
        if (typeof op.title === "string" && op.title.trim()) out.title = op.title.trim().slice(0, 200);
      } else if (t === "create_shape") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        // curated PowerPoint.GeometricShapeType palette (0.9.266): anything
        // outside this list ladders down to a plain rectangle at apply time,
        // so keep the set rich but verified; common model aliases resolve too
        var KINDS = { textbox: 1, rectangle: 1, rect: 1, ellipse: 1, circle: 1, oval: 1,
          roundedRectangle: 1, hexagon: 1, diamond: 1, triangle: 1, line: 1,
          chevron: 1, rightArrow: 1, downArrow: 1, upArrow: 1, leftArrow: 1,
          donut: 1, plus: 1, octagon: 1, parallelogram: 1, pentagon: 1,
          teardrop: 1, cube: 1, can: 1, moon: 1, lightningBolt: 1, sun: 1,
          heart: 1, wave: 1, doubleWave: 1, star4: 1, star5: 1, star6: 1, star8: 1,
          cloud: 1, arc: 1, blockArc: 1, pie: 1, foldedCorner: 1, halfFrame: 1, smileyFace: 1 };
        var kindS = String(op.kind || "");
        kindS = Object.keys(KINDS).find(function (kk) { return kk.toLowerCase() === kindS.toLowerCase(); });
        if (!kindS) { bad("kind must be one of the supported shapes (textbox, rectangle, circle/ellipse, chevron, rightArrow, downArrow, upArrow, leftArrow, donut, plus, hexagon, star5, star8, teardrop, parallelogram, pentagon, octagon, cube, can, moon, lightningBolt, sun, heart, wave, doubleWave, arc, blockArc, pie, foldedCorner, halfFrame, smileyFace, roundedRectangle, diamond, triangle, line)"); return; }
        kindS = { rect: "rectangle", circle: "ellipse", oval: "ellipse" }[kindS] || kindS;
        var geom = {}, srcG = op.geometry && typeof op.geometry === "object" ? op.geometry : null;
        ["left", "top", "width", "height"].forEach(function (gk) {
          var v = srcG ? srcG[gk] : op[gk];
          if (typeof v === "number" && v >= -2000 && v <= 4000) geom[gk] = Math.round(v);
        });
        var rotT = srcG && srcG.rotation != null ? srcG.rotation : op.rotation;
        if (typeof rotT === "number" && rotT >= -360 && rotT <= 360) geom.rotation = rotT;
        else if (rotT != null) { bad("rotation must be a number between -360 and 360"); return; }
        var txtS = "";
        if (typeof op.text === "string" && op.text.trim()) {
          if (kindS !== "textbox") { bad("text applies only to the textbox kind"); return; }
          txtS = op.text.trim().slice(0, 400);
        }
        out.slide = op.slide; out.kind = kindS; out.geometry = geom;
        if (txtS) out.text = txtS;
        if (typeof op.fill === "string" && op.fill.trim()) {
          if (!/^[\w#\(\), .%]{1,40}$/.test(op.fill.trim())) { bad("bad fill color"); return; }
          out.fill = op.fill.trim();
        }
        if (typeof op.line === "string" && op.line.trim()) {
          if (!/^[\w#\(\), .%]{1,40}$/.test(op.line.trim())) { bad("bad line color"); return; }
          out.line = op.line.trim();
        }
        if (typeof op.transparency === "number" && op.transparency >= 0 && op.transparency <= 1) out.transparency = op.transparency;
      } else if (t === "move_slide") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        if (typeof op.to !== "number" || Math.floor(op.to) !== op.to || op.to < 1 || op.to > ctx.slideCount) { bad("to must be a slide number 1.." + ctx.slideCount); return; }
        out.slide = op.slide; out.to = Math.round(op.to);
      } else if (t === "duplicate_slide") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = op.slide;
      } else if (t === "duplicate_slide_to") {
        if (!slideOk(op.slide)) { bad("bad source slide index"); return; }
        if (typeof op.to !== "number" || Math.floor(op.to) !== op.to || op.to < 1 || op.to > ctx.slideCount) { bad("to must be a slide number 1.." + ctx.slideCount); return; }
        out.slide = op.slide; out.to = Math.round(op.to);
      } else if (t === "clear_slide") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = op.slide;
      } else if (t === "apply_layout") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = op.slide; out.layout = normalizeLayout(op.layout);
      } else if (t === "copy_shape" || t === "move_shape_to") {
        if (!slideOk(op.slide)) { bad("bad source slide index"); return; }
        if (!futureSlide(op.slide) && !shapeIdxOk(op.slide, op.shape)) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        out.slide = op.slide; out.shape = Math.round(op.shape);
        if (op.to == null) {
          out.to = op.slide;
        } else if (typeof op.to === "number" && Math.floor(op.to) === op.to && op.to >= 1 && op.to <= (ctx ? ctx.slideCount : op.slide)) {
          out.to = Math.round(op.to);
        } else { bad("to must be a slide number 1.." + (ctx ? ctx.slideCount : op.slide)); return; }
        var geoX = {};
        ["left", "top", "width", "height"].forEach(function (gk5) {
          if (typeof op[gk5] === "number") geoX[gk5] = op[gk5];
        });
        if (Object.keys(geoX).length) out.geometry = geoX;
      } else if (t === "delete_slide") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = op.slide;
      }
      if ((t === "create_table" || t === "create_shape" || t === "create_image") && typeof out.slide === "number") {
        // register the batch-created shape's TEXT for later ops' find checks
        // (0.9.265: a create_shape textbox earlier in the batch was invisible
        // to textOf - a format op's find read the PRE-batch snapshot and
        // false-rejected "2 · Wire" not found). Only when the slide's shape
        // count is detailed in the snapshot (else defer to apply time).
        var csX = ctx && ctx.slides[out.slide - 1] && ctx.slides[out.slide - 1].shapes;
        if (out.text && ctx && out.slide <= ctx.slideCount && csX && csX.length) {
          textState[out.slide + ":" + (csX.length + (plannedShapes[out.slide] || 0) + 1)] = out.text;
        }
        plannedShapes[out.slide] = (plannedShapes[out.slide] || 0) + 1;
      }
      ops.push(out);
    });
    // slide-lifecycle ops reindex the deck mid-run - they must run alone
    var ALONE_TYPES = ["move_slide", "duplicate_slide", "duplicate_slide_to"];
    var hasAlone = ops.some(function (op) { return ALONE_TYPES.indexOf(op.type) >= 0; });
    if (hasAlone && ops.length > 1) {
      var moved = ops.filter(function (op) { return ALONE_TYPES.indexOf(op.type) >= 0; });
      var others = ops.filter(function (op) { return ALONE_TYPES.indexOf(op.type) < 0; });
      moved.forEach(function (op) {
        invalid.push({ op: op, reason: op.type + " reindexes the deck - run it alone" });
      });
      ops = others.length ? others : moved.slice(); // restore if nothing else remains
      if (others.length) moved.length = 0; else moved.forEach(function (op) { invalid.pop(); });
    }
    (preInvalid || []).forEach(function (pi) {
      invalid.push({ op: pi.op || { type: "?" }, reason: pi.reason || "dropped by the target guard" });
    });
    // 0.9.270 HARD target guard: an insert-after ask may only write the named
    // slide and the slides THIS batch creates - any other write to an
    // EXISTING slide is dropped with a visible reason. Lives on SK for the
    // whole run so repair/recovery re-validations stay protected too; reset
    // at the next renderParsed.
    var ig = SK.insertGuard;
    if (ig && ig.named != null && ig.base != null && ctx) {
      var cCt = ops.filter(function (o5) { return o5.type === "create_slide"; }).length;
      for (var ix5 = ops.length - 1; ix5 >= 0; ix5--) {
        var o6 = ops[ix5];
        var s6 = o6 ? o6.slide : undefined;
        if (o6.type === "create_slide" || s6 == null) continue;
        if (s6 <= ctx.slideCount && s6 !== ig.named &&
            !(s6 > ig.named && s6 <= ig.named + cCt)) {
          invalid.push({ op: o6,
            reason: "target guard: this ask inserts slides after slide " + ig.named +
              " - this action wrote to existing slide " + s6 +
              ", which this batch does not create; dropped to protect your content" });
          ops.splice(ix5, 1);
        }
      }
    }
    return { ops: ops, invalid: invalid, notesDropped: notesDropped };
  }

  function approvalMode(opType) {
    return SK.sessionApprovals[opType] || (SK.cfg && SK.cfg.approvals && SK.cfg.approvals[opType]) || "confirm";
  }
  function restoreSessionApprovals() {
    // approvals ride the session: on a pane reload the session resumes and
    // the remembered approvals come back with it (same session id only)
    try {
      var saved = JSON.parse(localStorage.getItem("sk-approvals-slidekick") || "null");
      if (saved && saved.map && (saved.sid || "") === (SK.sessionId || "")) {
        SK.sessionApprovals = saved.map;
        return true;
      }
    } catch (eR2) { /* corrupt entry - start clean */ }
    return false;
  }

  var OP_EFFECT = {
    create_slide: "add", edit_slide: "edit", replace_text: "edit", set_notes: "edit",
    create_table: "add", create_shape: "add", create_image: "add", insert_image: "add", set_shape_image: "add",
    apply_layout: "edit", format_text: "edit", format_selection: "edit",
    move_shape: "edit", move_slide: "edit", set_background: "edit", add_hyperlink: "edit",
    set_alt_text: "edit", set_categories: "edit", draft_reply: "add", style_shape: "edit", new_message: "add",
    compose_body: "edit", save_draft: "add", create_template: "add", web_lookup: "edit",
    delete_slide: "del", delete_shape: "del",
    duplicate_slide: "add", duplicate_slide_to: "add", clear_slide: "edit"
  };

  function fmtOp(op, ctx) {
    var slideTitle = function (sl) {
      if (!ctx || !ctx.slides[sl - 1]) return "";
      var t = ctx.slides[sl - 1].title;
      return t ? ' "' + t + '"' : "";
    };
    switch (op.type) {
      case "create_slide":
        return "New slide" + (op.after != null && !SK.sim ? " (requested after " + op.after + ")" : "") + ': "' + op.title + '"' +
          (op.body && op.body.length ? " - " + op.body.length + " bullet(s)" : "") +
          (op.notes ? " + notes" : "");
      case "edit_slide": {
        var parts = [];
        if (op.title != null) parts.push("title");
        if (op.body != null) parts.push("body");
        if (op.notes != null) parts.push("notes");
        return "Slide " + op.slide + slideTitle(op.slide) + ": set " + parts.join(" + ") +
          (op.title != null ? ' -> "' + op.title + '"' : "");
      }
      case "replace_text":
        return "Slide " + op.slide + ", shape " + op.shape + (op.find ? ": \"" + op.find + "\" -> " : ": -> ") + op.text;
      case "set_notes":
        return "Slide " + op.slide + slideTitle(op.slide) + ": speaker notes (" + op.notes.length + " chars)";
      case "create_table":
        return (op.slide != null ? "Slide " + op.slide + slideTitle(op.slide) : "New slide") +
          ': table ' + op.values.length + "x" + op.values[0].length +
          (op.geometry && op.geometry.left != null ? " @ " + Math.round(op.geometry.left) + "," + Math.round(op.geometry.top) : "");
      case "move_shape":
        return "Slide " + op.slide + ", shape " + op.shape + ": " +
          Object.keys(op.geometry || {}).map(function (p) { return p + "->" + Math.round(op.geometry[p]); }).join(" ") || "(no geometry)";
      case "copy_shape":
      case "move_shape_to":
        return (op.type === "move_shape_to" ? "Cut " : "Copy ") + "slide " + op.slide + " shape " + op.shape +
          " -> slide " + op.to +
          (op.geometry && op.geometry.left != null ? " @ left " + Math.round(op.geometry.left) + ", top " + Math.round(op.geometry.top) : "") +
          (op.type === "move_shape_to" ? " (original removed)" : " (duplicate)");
      case "create_image": {
        var pr = op.prompt.length > 60 ? op.prompt.slice(0, 60) + "…" : op.prompt;
        return (op.slide != null ? "Slide " + op.slide + slideTitle(op.slide) : "Slide you have open") +
          ': image "' + pr + '" (profile image model)';
      }
      case "format_selection":
        return "the text you selected: " + Object.keys(op.font).join(", ");
      case "create_shape": {
        var gc = [];
        ["left", "top", "width", "height"].forEach(function (gk2) {
          if (op.geometry && op.geometry[gk2] != null) gc.push(gk2 + "=" + op.geometry[gk2]);
        });
        return "Slide " + op.slide + ": " + op.kind + (op.text ? " \"" + op.text + "\"" : "") +
          (gc.length ? " (" + gc.join(", ") + ")" : "");
      }
      case "duplicate_slide":
        return "Duplicate " + slideTitle(op.slide) + " as slide " + (op.slide + 1);
      case "duplicate_slide_to":
        return "Copy slide " + op.slide + slideTitle(op.slide) + " -> over slide " + op.to;
      case "clear_slide":
        return "Clear " + slideTitle(op.slide) + " (texts emptied, created shapes removed)";
      case "move_slide":
        return "Move slide " + op.slide + " -> slide " + op.to;
      case "set_background":
        return "Slide " + op.slide + " background: " + op.kind +
          (op.kind === "image" ? " (picture: " + op.prompt.slice(0, 50) + "…)" : op.kind === "solid" ? " (" + op.color + ")" : "");
      case "add_hyperlink":
        return "Slide " + op.slide + ", shape " + op.shape + " -> " + op.url;
      case "style_shape":
        return "Slide " + op.slide + ", shape " + op.shape + ": " +
          (op.fill ? "fill " + op.fill : "") + (op.fill && (op.line || op.transparency != null) ? ", " : "") +
          (op.line ? "line " + op.line : "") + (op.fill && op.line && op.transparency != null ? ", " : "") +
          (op.transparency != null ? "transparency " + op.transparency : "");
      case "set_shape_image": {
        var prI = op.prompt.length > 50 ? op.prompt.slice(0, 50) + "\u2026" : op.prompt;
        return "Slide " + op.slide + ", shape " + op.shape + ': picture-fill "' + prI + '"' +
          (op.transparency != null ? " transparency " + op.transparency : "") + " (profile image model)";
      }
      case "set_alt_text":
        return "Slide " + op.slide + ", shape " + op.shape + " alt text";
      case "format_text": {
        var fparts = [];
        Object.keys(op.font).forEach(function (kf) {
          var v = op.font[kf];
          fparts.push(kf + "=" + (typeof v === "boolean" ? (v ? "on" : "off") : v));
        });
        return "Slide " + op.slide + ", shape " + op.shape + ": " + fparts.join(", ") +
          (op.find ? " (in \"" + op.find + "\")" : "");
      }
      case "apply_layout":
        return "Slide " + op.slide + slideTitle(op.slide) + ": layout -> " + op.layout;
      case "delete_slide":
        return "Slide " + op.slide + slideTitle(op.slide);
    }
    return op.type;
  }

  // ------------------------------------------------------------------ //
  // Apply ops (shared by Office + simulator)
  // ------------------------------------------------------------------ //
  function applyOps(deck, ops, ctx) {
    var lines = [], manualNotes = [], shift = 0;
    var failedOps = [];      // ops that threw (retry pass re-runs these, in order)
    var abortedIdx = -1;     // set when an unexpected error aborts the chain
    var runOne = function (op) {
      var cur = op.slide !== undefined ? op.slide - 1 - shift : undefined;
      switch (op.type) {
        case "create_slide":
          return deck.addSlide({ layout: op.layout, title: op.title, body: op.body }).then(function (r) {
            // honor the requested position: 0-based index the new slide should END UP at
            var wantIdx = Math.min(Math.max(op.after, 0), r.index);   // after:N (1-based slide) => 0-based N
            var doNotes = function () {
              if (op.notes && op.notes.trim()) {
                return deck.setNotes(r.index, op.notes).then(function (nr) {
                  if (nr && nr.manual) manualNotes.push({ slide: r.index, text: op.notes, reason: nr.reason });
                  else lines.push({ ok: true, text: "set notes on new slide " + (r.index + 1) });
                });
              }
              return null;
            };
            if ((op.after != null) && r.index !== wantIdx) {
              return deck.moveSlide(r.index, wantIdx).then(function () {
                lines.push({ ok: true, text: "create slide -> inserted at position " + (wantIdx + 1) });
                return doNotes();
              }, function (eMove) {
                var note = " (requested position failed: " + ((eMove && eMove.message) || eMove) + " - left at end)";
                lines.push({ ok: true, text: "create slide -> appended at index " + (r.index + 1) + note });
                return doNotes();
              });
            }
            lines.push({ ok: true, text: "create slide -> appended at index " + (r.index + 1) });
            return doNotes();
          });
        case "edit_slide": {
          var chain = Promise.resolve();
          if (op.title != null) {
            chain = chain.then(function () {
              return deck.setSlideTitle(cur, op.title).then(function () {
                lines.push({ ok: true, text: "edit slide " + op.slide + ": title set" });
              });
            });
          }
          if (op.body != null) {
            chain = chain.then(function () {
              return deck.setSlideBody(cur, op.body).then(function () {
                lines.push({ ok: true, text: "edit slide " + op.slide + ": body set (" + op.body.length + " paragraphs)" });
              });
            });
          }
          if (op.notes != null) {
            chain = chain.then(function () {
              return deck.setNotes(cur, op.notes).then(function (nr) {
                if (nr && nr.manual) manualNotes.push({ slide: cur, text: op.notes, reason: nr.reason });
                else lines.push({ ok: true, text: "edit slide " + op.slide + ": notes set" });
              });
            });
          }
          return chain;
        }
        case "replace_text": {
          if (op.find) {
            return deck.setShapeTextFind(cur, op.shape - 1, op.find, op.text).then(function () {
              lines.push({ ok: true, text: "replace \"" + op.find + "\" -> \"" + op.text + "\" on slide " + op.slide + " shape " + op.shape });
            });
          }
          return deck.setShapeText(cur, op.shape - 1, op.text).then(function () {
            lines.push({ ok: true, text: "replace text on slide " + op.slide + " shape " + op.shape });
          });
        }
        case "create_table": {
          var runTable = function (slideIndex) {
            return deck.addTable({ values: op.values, slide: slideIndex, geometry: op.geometry }).then(function (r) {
              lines.push({ ok: true, text: "create table " + op.values.length + "x" + op.values[0].length + " -> " +
                (op.slide !== undefined && op.slide !== null ? "slide " + op.slide : "new slide " + (r && r.newIndex != null ? r.newIndex : slideIndex)) });
            });
          };
          if (op.slide !== undefined && op.slide !== null) return runTable(cur);
          return deck.addSlide({ layout: "title_only", title: op.title || "" }).then(function (created) {
            lines.push({ ok: true, text: "create slide -> appended at index " + created.index + " (for table)" });
            return runTable(created.index);
          });
        }
        case "move_shape":
          return deck.moveShape(cur, op.shape - 1, op.geometry || {}).then(function () {
            lines.push({ ok: true, text: "move shape " + op.shape + " on slide " + op.slide + ": " +
              Object.keys(op.geometry || {}).map(function (p) { return p + "->" + Math.round(op.geometry[p]); }).join(" ") });
          });
        case "copy_shape":
        case "move_shape_to":
          return deck.copyPasteShape(cur, op.shape - 1, op.to - 1, op.geometry || null,
            op.type === "move_shape_to").then(function (res) {
            lines.push({ ok: true, text: (op.type === "move_shape_to" ? "moved" : "copied") + " shape " +
              op.shape + " slide " + op.slide + " -> slide " + op.to + ": " + (res && res.how ? res.how.split("-> ")[1] || res.how : "") });
          });
        case "set_notes":
          return deck.setNotes(cur, op.notes).then(function (nr) {
            if (nr && nr.manual) manualNotes.push({ slide: cur, text: op.notes, reason: nr.reason });
            else lines.push({ ok: true, text: "set notes on slide " + op.slide });
          });
        case "create_shape":
          return deck.createShape(cur, op).then(function (r) {
            lines.push({ ok: true, text: fmtOp(op, ctx) + (r && r.note ? " [" + r.note + "]" : "") });
          });
        case "duplicate_slide":
          return deck.duplicateSlide(op.slide - 1).then(function (r) {
            lines.push({ ok: true, text: "duplicated slide " + op.slide + " -> inserted as slide " + (r.at + 1) });
          });
        case "duplicate_slide_to":
          return deck.duplicateSlideTo(op.slide - 1, op.to - 1).then(function () {
            lines.push({ ok: true, text: "copied slide " + op.slide + " onto slide " + op.to + " (target replaced)" });
          });
        case "clear_slide":
          return deck.clearSlide(op.slide - 1).then(function (r) {
            lines.push({ ok: true, text: "cleared slide " + op.slide + ": " + r.cleared + " text(s) emptied" +
              (r.removed ? ", " + r.removed + " line(s) removed" : "") });
          });
        case "move_slide": {
          if (op.slide === op.to) {
            lines.push({ ok: true, text: "slide " + op.slide + " is already in place" });
            return Promise.resolve();
          }
          return deck.moveSlide(op.slide - 1, op.to - 1).then(function () {
            lines.push({ ok: true, text: "moved slide " + op.slide + " -> now slide " + op.to });
          });
        }
        case "format_selection":
          return deck.formatSelection(op.font).then(function () {
            lines.push({ ok: true, text: "formatted your selection: " + Object.keys(op.font).join(", ") });
          });
        case "set_background": {
          if (op.kind === "image") {
            var profB = currentProfile();
            var imgRefB = profB && profB.image;
            if (!imgRefB || !imgRefB.connection || !imgRefB.model) {
              throw new FriendlyError(
                "No image model on profile '" + (profB ? profB.name : "?") +
                "' - set one in Profiles (connection + model with an images API).");
            }
            setStatus("Generating background… (often 30-90 s depending on the model)", "busy");
            return imageJob({
                connection: imgRefB.connection, model: imgRefB.model, prompt: op.prompt,
              }).then(function (img) {
              if (!img || !img.b64) throw new FriendlyError("server returned no image data");
              return deck.setBackground(cur, { kind: "image", b64: img.b64, transparency: op.transparency }).then(function () {
                lines.push({ ok: true, text: "slide " + op.slide + " background set (AI picture, " + imgRefB.connection + ":" + imgRefB.model + ")" });
              });
            });
          }
          return deck.setBackground(cur, op).then(function () {
            lines.push({ ok: true, text: "slide " + op.slide + " background set (" + op.kind + (op.kind === "solid" ? " " + op.color : "") + ")" });
          });
        }
        case "set_shape_image": {
          var profI = currentProfile();
          var imgRefI = profI && profI.image;
          if (!imgRefI || !imgRefI.connection || !imgRefI.model) {
            throw new FriendlyError(
              "No image model on profile '" + (profI ? profI.name : "?") +
              "' - set one in Profiles (connection + model with an images API).");
          }
          setStatus("Generating picture fill\u2026 (often 30-90 s depending on the model)", "busy");
          return imageJob({
              connection: imgRefI.connection, model: imgRefI.model, prompt: op.prompt,
            }).then(function (img) {
            if (!img || !img.b64) throw new FriendlyError("server returned no image data");
            return deck.setShapeImage(cur, op.shape - 1, img.b64, op.transparency).then(function () {
              lines.push({ ok: true, text: "picture-filled shape " + op.shape + " on slide " + op.slide +
                " (" + imgRefI.connection + ":" + imgRefI.model + ")" });
            });
          });
        }
        case "style_shape":
          return deck.styleShape(cur, op.shape - 1, op.fill, op.line, op.transparency).then(function () {
            lines.push({ ok: true, text: "styled shape " + op.shape + " on slide " + op.slide +
              (op.fill ? " fill=" + op.fill : "") + (op.line ? " line=" + op.line : "") +
              (op.transparency != null ? " transparency=" + op.transparency : "") });
          });
        case "add_hyperlink":
          return deck.addHyperlink(cur, op.shape - 1, op.url, op.tip).then(function () {
            lines.push({ ok: true, text: "hyperlink added on slide " + op.slide + ", shape " + op.shape });
          });
        case "set_alt_text":
          return deck.setAltText(cur, op.shape - 1, op.description, op.title).then(function () {
            lines.push({ ok: true, text: "alt text set on slide " + op.slide + ", shape " + op.shape });
          });
        case "format_text": {
          if (op.find) {
            return deck.formatTextSubstring(cur, op.shape - 1, op.find, op.font).then(function () {
              lines.push({ ok: true, text: "format slide " + op.slide + " shape " + op.shape +
                " (" + op.find + "): " + Object.keys(op.font).join(", ") });
            });
          }
          return deck.formatText(cur, op.shape - 1, op.font).then(function () {
            lines.push({ ok: true, text: "format slide " + op.slide + " shape " + op.shape + ": " +
              Object.keys(op.font).join(", ") });
          });
        }
        case "apply_layout":
          return deck.applyLayout(cur, op.layout).then(function () {
            lines.push({ ok: true, text: "apply layout '" + op.layout + "' to slide " + op.slide });
          });
        case "create_image": {
          var prof = currentProfile();
          var imgRef = prof && prof.image;
          if (!imgRef || !imgRef.connection || !imgRef.model) {
            throw new FriendlyError(
              "No image model on profile '" + (prof ? prof.name : "?") +
              "' - set one in Profiles (connection + model with an images API).");
          }
          setStatus("Generating image… (often 30-90 s depending on the model)", "busy");
          return imageJob({
              connection: imgRef.connection, model: imgRef.model, prompt: op.prompt,
            }).then(function (img) {
            if (!img || !img.b64) throw new FriendlyError("server returned no image data");
            // the direct insert reaches only the slide the user has OPEN -
            // when the op names a DIFFERENT slide (typical: the batch just
            // created it), go straight to the picture-filled-shape path
            // (0.9.265) so the image lands on the requested slide.
            var offSlide = !SK.sim && op.slide != null && ctx && op.slide - 1 !== ctx.activeSlide;
            var useShape = offSlide ? deck.addImageShape({ b64: img.b64, slide: cur, geometry: op.geometry })
                                    : deck.addImage({ b64: img.b64, slide: cur, geometry: op.geometry });
            return useShape.then(function (ri) {
              var note = (offSlide && !(ri && ri.note))
                ? " (targeting the slide you have open in PowerPoint; requested slide " + op.slide + ")" : "";
              if (ri && ri.note) note += (note ? " " : "") + "[" + ri.note + "]";
              lines.push({ ok: true, text: "image inserted (" + imgRef.connection + ":" + imgRef.model + ")" + note });
            });
          });
        }
        case "delete_slide":
          return deck.deleteSlide(cur).then(function () {
            lines.push({ ok: true, text: "delete slide " + op.slide + " (" + fmtOp(op, ctx) + ")" });
            shift++;
          });
      }
      return Promise.resolve();
    };

    var seq = Promise.resolve();
    ops.forEach(function (op, ix) {
      seq = seq.then(function () {
        if (SK.runCancelled) {
          lines.push({ ok: false, text: "cancel - remaining op(s) not executed (" + op.type + " skipped)" });
          return null;
        }
        var l0 = lines.length;
        return runOne(op).then(function () {
          if (typeof SK.opProgress === "function") SK.opProgress(op, lines.slice(l0), true);
        }, function (err) {
          var msg = err && err.friendly ? err.message : (err && err.message) || String(err);
          lines.push({ ok: false, text: op.type + ": " + msg });
          failedOps.push(op);
          if (typeof SK.opProgress === "function") SK.opProgress(op, lines.slice(l0), false);
          if (!(err && err.friendly)) {
            lines.push({ ok: false, text: "unexpected error - stopping remaining ops" });
            abortedIdx = ix;
            throw err; // aborts the chain below; remaining ops join the retry pass
          }
        });
      });
    });
    return seq.then(
      function () { return null; },
      function (err) {
        var msg = err && err.friendly ? err.message : (err && err.message) || String(err);
        lines.push({ ok: false, text: "aborted: " + msg });
        return msg;
      }
    ).then(function (abortedMsg) {
      var retryQueue = failedOps.slice();
      if (abortedMsg && abortedIdx >= 0) retryQueue = failedOps.concat(ops.slice(abortedIdx + 1));
      if (!retryQueue.length) return { lines: lines, manualNotes: manualNotes };
      // do its best to complete every step: one extra pass for everything
      // that failed (and for ops skipped by an unexpected error), same order
      lines.push({ ok: true, text: "retrying " + retryQueue.length +
        " uncompleted action(s) once after a short pause…" });
      var rseq = new Promise(function (res) { setTimeout(res, 1500); });
      retryQueue.forEach(function (op) {
        rseq = rseq.then(function () {
          var l0 = lines.length;
          return runOne(op).then(function () {
            // a rescued op moves to the completed list like any other success
            if (typeof SK.opProgress === "function") SK.opProgress(op, lines.slice(l0), true);
          }, function (err2) {
            var msg2 = err2 && err2.friendly ? err2.message : (err2 && err2.message) || String(err2);
            lines.push({ ok: false, text: "retry of " + op.type + " failed again: " + msg2 });
            if (typeof SK.opProgress === "function") SK.opProgress(op, lines.slice(l0), false);
          });
        });
      });
      return rseq.then(function () { return { lines: lines, manualNotes: manualNotes }; });
    });
  }

  // Slide role for design work: what the slide IS (title/section cover,
  // bullets, statement, table, image-led). The make-over prompt assigns one
  // layout blueprint per role so same-role slides read as one family.
  function slideRoleOf(slide) {
    var shapes = (slide && slide.shapes) || [];
    var texts = shapes.filter(function (sh) { return typeof sh.text === "string" && sh.text.trim(); });
    var words = texts.reduce(function (acc, sh) {
      return acc + String(sh.text).trim().split(/\s+/).filter(Boolean).length;
    }, 0);
    var sig = shapes.map(function (sh) { return (sh.type || "") + " " + (sh.name || ""); });
    if (sig.some(function (x) { return (/table/i).test(x); })) return "table";
    if (sig.some(function (x) { return (/picture|image|chart|media|model3d|graphicframe/i).test(x); }) &&
        words < 60) return "image";
    var bodyTexts = shapes.filter(function (sh) {
      return sh.isBody && typeof sh.text === "string" && sh.text.trim();
    });
    var bodyCount = 0;
    bodyTexts.forEach(function (sh) {
      var lns = String(sh.text).split(/\r|\n/).filter(function (ln) { return ln.trim(); });
      bodyCount += Math.max(lns.length, 1);
    });
    // cover/section = slide 1, or any slide without real body content
    if (slide.idx === 0 || (!bodyTexts.length && words <= 40)) return "title";
    if (bodyCount >= 2) return "bullets";
    if (words <= 30) return "statement";
    return "bullets";
  }

  // ------------------------------------------------------------------ //
  // Chat / prompt flow
  // ------------------------------------------------------------------ //
  function formatContext(ctx) {
    var lines = [];
    var sel = ctx.selected && ctx.selected.length ? " (slides are numbered 1..N; * = selected: " +
      ctx.selected.map(function (x) { return x + 1; }).join(",") + ")" : "";
    lines.push("DECK OUTLINE: " + ctx.slideCount + " slides" + sel + (ctx.truncated ? " (only the first slides are listed)" : "") + ", host: " + ctx.host +
      "\n(Per-shape text may be shortened in this outline. To check/refer to full slide text, first emit a read_slide op for that slide - its reply shows every shape in full; never claim you cannot see content without using read_slide.)");
    if (ctx.canDelete === false) {
      lines.push("SLIDE DELETION: NOT available on this build (the add-in API cannot remove slides) - never emit delete_slide; restructure with create_slide/duplicate_slide_to/move_slide instead.");
    }
    if (ctx.selection && (ctx.selection.shapes.length || ctx.selection.textPreview || ctx.selection.slides.length)) {
      var selBits = [];
      if (ctx.selection.slides.length) selBits.push("slides " + ctx.selection.slides.map(function (x) { return x + 1; }).join(","));
      if (ctx.selection.shapes.length)
        selBits.push("shapes " + ctx.selection.shapes.map(function (x) { return (x.slide + 1) + "." + (x.shape + 1); }).join(","));
      if (ctx.selection.textPreview) selBits.push("selected text: \"" + ctx.selection.textPreview + "\"");
      if (selBits.length) lines.push("SELECTION: " + selBits.join(" | "));
    }
    if (ctx.slides.length > 40) {
      var selMap = {};
      (ctx.selected || []).forEach(function (x) { selMap[x] = 1; });
      var bodyPrevF = function (s2) {
        var t2 = (s2.shapes || []).filter(function (sh) { return sh.isBody && sh.text; })
          .map(function (sh) { return sh.text; }).join(" ") || "";
        t2 = String(t2).replace(/\r|\n/g, " ");
        if (t2.length > 240) t2 = t2.slice(0, 240) + "…";
        return t2 ? " — " + t2 : "";
      };
      ctx.slides.forEach(function (s2) {
        var role2 = slideRoleOf(s2);
        var star2 = ctx.selected && ctx.selected.indexOf(s2.idx) >= 0 ? "*" : "";
        var tt2 = (s2.shapes || []).filter(function (sh) { return sh.isTitle || (sh.text && /title/i.test(sh.name)); })[0];
        var line2 = "[" + (s2.idx + 1) + star2 + "] (role: " + role2 + ") " +
          (tt2 && tt2.text ? tt2.text : "(no title)") + bodyPrevF(s2);
        if (line2.length > 340) line2 = line2.slice(0, 340);
        lines.push(line2);
        if (selMap[s2.idx]) {
          var selTxt = 0;
          (s2.shapes || []).forEach(function (sh) {
            var txt = sh.text == null ? "" : String(sh.text).replace(/\r|\n/g, " | ");
            if (!String(txt).trim()) txt = "(no readable text - table/gfx shape)";
            else if (txt.length > 900) txt = txt.slice(0, 900) + "…[text continues]";
            var role = sh.isTitle ? "title" : sh.isBody ? "body" : "text";
            lines.push("    shape " + (sh.idx + 1) + " (" + role + ", " + sh.name + '): "' + txt + '"');
            selTxt++;
          });
          if (!selTxt) lines.push("    (the selected slide has no readable text shapes)");
        }
      });
      lines.push("(compact outline for a large deck - " + ctx.slides.length + " slides listed; body text is a preview; the SELECTED slide(s) marked * show every shape with its complete text; full text of any other slide: reply {\"kind\":\"read_slide\",\"slides\":[N]})");
    } else
    ctx.slides.forEach(function (s) {
      var role = slideRoleOf(s);
      var star = ctx.selected && ctx.selected.indexOf(s.idx) >= 0 ? "*" : "";
      var title = (s.shapes || []).filter(function (sh) { return sh.isTitle || (sh.text && /title/i.test(sh.name)); })[0];
      lines.push("[" + (s.idx + 1) + star + "] (role: " + role + ") " + (title && title.text ? 'TITLE: "' + title.text + '"' : "(no title)"));
      (s.shapes || []).forEach(function (sh) {
        if (sh.text == null) return;
        // shapes are numbered from 1 in the outline listing below
        var role = sh.isTitle || (title && sh.idx === title.idx) ? "title" : sh.isBody ? "body" : "text";
        var txt = sh.text == null ? "" : String(sh.text).replace(/\r|\n/g, " | ");
        if (!txt.trim()) txt = "(no readable text - table/gfx shape)";
        else if (txt.length > 900) txt = txt.slice(0, 900) + "…[text continues]";
        lines.push("    shape " + (sh.idx + 1) + " (" + role + ", " + sh.name + '): "' + txt + '"');
      });
    });
    var total = lines.join("\n");
    if (total.length > 60000) total = total.slice(0, 60000) + "\n…[truncated]";
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
    SK.insertGuard = null;   // a new run rebuilds the target guard from its own parse (0.9.270)
    SK.targetConfirmDone = false;   // one AI target-confirmation round per run
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
    setStatus("Collecting deck context…", "busy");
    var ctxPromise = (wantCtx ? SK.deck.getContext({}) : Promise.resolve(null)).then(function (ctx) {
      if (ctx && SK.deck.getSelectionInfo) {
        return SK.deck.getSelectionInfo().then(function (sel) {
          if (sel) ctx.selection = sel;
          return ctx;
        }).catch(function () { return ctx; });
      }
      return ctx;
    }).catch(function (err) {
      setStatus("Could not read the deck (" + err.message + ") - continuing without context.", "error");
      return null;
    });

    ctxPromise.then(function (ctx) {
      SK.lastCtx = ctx;
      SK.readRounds = 0;
      SK.webRounds = 0;
      SK.pageReads = 0;
      SK.webResults = "";
      if (!(SK.clarifyAnswer && SK.clarifyAnswer.expect === instruction)) SK.clarifyChain = 0;   // a free new ask resets the clarify chain
      var webPromise = Promise.resolve("");
      return webPromise.then(function () {
      var sys = SYSTEM_PROMPT + "\nHost: " + (SK.sim ? "SlideKick simulator" : SK.deck.capabilities().host);
      var _cp = SK.cfg ? (String(SK.cfg.customPrompt || "").trim() ? String(SK.cfg.customPrompt).trim() : String(SK.cfg.customPromptDefault || "").trim()) : "";
      if (_cp) {
        sys += "\n\nUSER CUSTOM INSTRUCTIONS (from Settings - obey these):\n" + _cp;
      }
      if (!(SK.cfg && SK.cfg.generateNotes)) sys += "\n" + NOTES_OFF_RULE;
      if (ctx && ctx.canDelete === false) sys += "\n" + DELETE_RULE;
      sys += "\n" + CLARIFY_RULE;
      if (MAKEOVER_RE.test(instruction)) sys += "\n" + MAKEOVER_BRIEF;
      var user = (ctx ? formatContext(ctx) : "(no deck context requested)") + "\n\nUSER REQUEST:\n" + instruction;
      if (SK.clarifyAnswer && SK.clarifyAnswer.expect === instruction) {
        user += "\n\nCLARIFICATION: you asked the user: \"" + SK.clarifyAnswer.q + "\". The user chose: " + SK.clarifyAnswer.a +
          ". Deliver the final result for that choice now - do not ask again.";
        var namedCl = /(?:^|[\s(,.])slides?\s+(\d{1,4})\b/i.exec(instruction);
        if (namedCl) {
          user += "\nRemember: the request names slide " + namedCl[1] + " - deliver on slide " + namedCl[1] + " (1-based outline numbering), never on the open/selected slide.";
        }
        SK.clarifyAnswer = null;
        dbg("clarify.answer", { delivered: true });
      }
      if (SK.webResults) user += "\n\nWEB RESULTS (live lookup, cite them when relevant):\n" + SK.webResults;
      var messages = [
        { role: "system", content: sys },
      ].concat(SK.history.slice(-(HISTORY_MAX_EXCHANGES * 2)))
       .concat([{ role: "user", content: user }]);
      SK.lastAsk = { messages: messages, connection: ref.connection, model: ref.model, rounds: 0 };   // 0.9.216: repair rounds on validation rejects
      dbg("prompt", { profile: prof.name, route: route, model: ref.model, connection: ref.connection, instruction: instruction, hadContext: !!ctx, messages: messages });
      setStatus("Asking " + ref.model + " (via " + ref.connection + ")…", "busy");
      var askedLookup = false;
      // 0.9.186: content-creation requests ALWAYS run the web lookup first
      // (standing rule: the pane owns the lookup channel - internal, free);
      // fetched results feed the generation ask; the 3-lookup budget still caps it.
      // 0.9.189: when the user asks to READ/QUOTE a slide's content, the pane
      // fetches the FULL slide text itself before the first ask - the compact
      // outline's 900-char preview must never be the source of a quote.
      var READ_ASK_RE = /\b(read|quote|paste|show|display|list|repeat|give|what('s| is)\b|what does)\b[^\n]{0,60}\bslides?\b/i;   // large decks only: the outline preview would be quoted
      var SLIDE_NUM_RE = /\bslides?\s+((?:\d+\s*(?:,\s*|\+|to|through\s+)?\s*)+\d)\b/gi;
      var preReadDone = null;
      function preReadRun() {
        if (preReadDone === null) {
          var slideCount = SK.lastCtx ? SK.lastCtx.slideCount : 0;
          var compact = SK.lastCtx ? ((SK.lastCtx.slides || []).length > 40) : false;
          var wantRead = READ_ASK_RE.test(instruction || "");
          var nums = [];
          if (wantRead && instruction) {
            var mm, re2 = /\bslides?\s+(\d+\b[^)]{0,40})/gi;
            while ((mm = re2.exec(instruction)) !== null && nums.length < 6) {
              var digits = mm[1].match(/\d+/g) || [];
              for (var qi = 0; qi < digits.length; qi++) {
                var n2 = parseInt(digits[qi], 10);
                if (n2 >= 1 && (slideCount ? n2 <= slideCount : true) && nums.indexOf(n2) < 0) nums.push(n2);
                if (nums.length >= 6) break;
              }
            }
            nums.sort(function (a, b) { return a - b; });
          }
          if (wantRead && nums.length) {
            SK.readRounds = (SK.readRounds || 0) + 1;   // the shared 2-read budget
            setStatus("Reading slide(s) " + nums.join(", ") + " in full\u2026", "busy");
            preReadDone = SK.deck.readSlides(nums).then(function (full) {
              dbg("pre.read_slide", { slides: nums, chars: full.length });
              messages = messages.concat([{ role: "user", content:
                "FULL SLIDE TEXT of the slide(s) the user asked you to read (1-based, same numbering as the outline):\n" + full +
                "\n\nUse this complete text for the ORIGINAL USER REQUEST; quote it accurately; do NOT reply with read_slide for these slide(s) again." }]);
              return true;
            }).catch(function (ePr) {
              dbg("pre.read_slide.failed", { err: String(ePr && ePr.message || ePr).slice(0, 160) });
              SK.readRounds = Math.max(0, (SK.readRounds || 1) - 1);
              return null;
            });
          } else {
            preReadDone = Promise.resolve(null);
          }
        }
        return preReadDone;
      }
      var CONTENT_WRITE_RE = /\b(write|create|generate|produce|draft|build|compose|make)\b[^\n]{0,90}?\b(slides?|deck|presentation|slideshow)\b/i;
      var preLookupDone = null;
      function preLookupRun() {
        if (preLookupDone === null) {
          if (CONTENT_WRITE_RE.test(instruction || "")) {
            SK.webRounds = (SK.webRounds || 0) + 1;   // the shared 3-lookup budget
            setStatus("Checking the web for current info to build from\u2026", "busy");
            preLookupDone = webLookup(prof, instruction).then(function (ws) {
              if (!ws) { SK.webRounds = Math.max(0, (SK.webRounds || 1) - 1); return null; }
              SK.webResults = ws;
              setStatus("Reading the found documents\u2026", "busy");
              return gatherPageContents(ws).then(function (pages) {
                SK.webResults = ws + pages;
                messages = messages.concat([{ role: "user", content:
                  "WEB LOOKUP RESULTS (the pane fetched these proactively before this content-building request - verify facts against them, keep them current, cite what you use; do NOT invent numbers/dates):\n" + ws +
                  (pages ? "\n\nThe lookups also RETURNED THE FULL TEXT of matched pages - read before building:\n" + pages : "") }]);
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
        return preReadRun().then(function () { return preLookupRun(); }).then(function () { return mainAskInner(); });
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
            if (pre.ok && pre.kind === "read_slide") return readSlideThen(resp);
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
        dbg("web_lookup", { round: SK.webRounds, profile: prof.name });
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
      // read_page: the model wants the FULL text of one page (usually a link
      // from a lookup result). Internal + free; up to 4 reads per run.
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
      // read_slide: the model asked for the FULL text of specific slide(s);
      // fetch it from the deck and re-ask with it (same flow as web_lookup),
      // bounded to 2 read rounds per run so a model that keeps asking cannot
      // loop forever.
      function readSlideThen(resp) {
        var pre = tryParse(resp.text);
        var lim = SK.lastCtx ? SK.lastCtx.slideCount : 0;
        var idxs = [];
        (pre.slides || []).forEach(function (n) {
          if (idxs.length >= 12 || idxs.indexOf(n) >= 0) return;
          if (typeof n === "number" && Math.floor(n) === n && n >= 1 && (lim ? n <= lim : true)) idxs.push(n);
        });
        if (!idxs.length)
          return { ok: true, kind: "reply",
                   text: JSON.stringify({ kind: "reply", text: "read_slide needs 1-based slide numbers from the outline (this deck has " + lim + " slide(s)). Ask again with valid numbers, or answer from the outline." }) };
        if ((SK.readRounds || 0) >= 2)
          return { ok: true, kind: "reply",
                   text: JSON.stringify({ kind: "reply", text: "Full slide text was already provided twice in this run. Answer from it now, or ask the user for specifics." }) };
        SK.readRounds = (SK.readRounds || 0) + 1;
        setStatus("Reading slide(s) " + idxs.join(", ") + " in full…", "busy");
        return SK.deck.readSlides(idxs).then(function (full) {
          dbg("read_slide", { slides: idxs, chars: full.length });
          messages = messages.concat([
            { role: "assistant", content: resp.text },
            { role: "user", content: "FULL SLIDE TEXT (slide numbers are 1-based, same numbering as the outline):\n" + full +
              "\n\nAnswer the ORIGINAL USER REQUEST now using this complete text. The required JSON format rules apply again. Do NOT reply with read_slide for these slide(s) again." }
          ]);
          setStatus("Answering with the full slide text…", "busy");
          return mainAsk2();
        }, function (err) {
          return { ok: true, kind: "reply",
                   text: JSON.stringify({ kind: "reply", text: "Reading slide(s) " + idxs.join(",") + " failed: " + ((err && err.message) || err) }) };
        });
      }
      function mainAsk2() {
        return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: messages }); })
          .then(function (resp) {
            var pre = tryParse(resp.text);
            if (pre.ok && pre.kind === "web_lookup") return webLookupThen(resp);
            if (pre.ok && pre.kind === "read_page") return readPageThen(resp);
            if (pre.ok && pre.kind === "read_slide") return readSlideThen(resp);
            return resp;
          });
      }
      // AI-talks-to-AI make-over: a design-director round invents the concept
      // (no ops), then builder rounds translate it into ops chunk by chunk so
      // a full-deck redesign can never silently stop after slide N.
      var fullMo = MAKEOVER_RE.test(instruction) &&
        !/\bslides?\s*\d+\b/i.test(instruction) &&
        !!(ctx && ctx.slides && ctx.slides.length);
      var moConcept = "";
      function makeOverFlow() {
        var slideN = ctx.slides.length;
        var MO_CHUNK = 4, MO_MAX = 24;
        var STRUCT_TYPES = { set_background: 1, create_shape: 1, move_shape: 1, create_slide: 1, set_shape_image: 1, style_shape: 1, duplicate_slide: 1, duplicate_slide_to: 1, create_table: 1, insert_image: 1 };
        var director = messages.concat([{
          role: "user",
          content: "You are the DESIGN DIRECTOR for this deck. Do NOT emit ops. " +
            "Reply with format A (reply) holding ONLY the design concept: an " +
            "invented, bold style name; a deep/panel/accent hex trio plus which " +
            "text color goes on each surface; the type scale (title/body sizes); " +
            "one image-style phrase used by every picture; and then ONE EXACT " +
            "GEOMETRY TEMPLATE PER SLIDE ROLE on the 960x540-point canvas - " +
            "for each role list every element with its concrete position and " +
            "size in points and its fill/transparency (example: 'cover: full-bleed " +
            "bg 0,0,960,540; overlay 0,0,960,540 fill #000 transparency 0.5; " +
            "title box 96,210,680,110; accent bar 96,196,8,84 accent'). The " +
            "builders will apply YOUR numbers as the design - so be precise and " +
            "make every slide with that role come out identical in structure." +
            "Take compositional risks - the opposite of a default PowerPoint look. " +
            "If a earlier redesign appears in the history, pick a contrasting " +
            "direction. 200-350 words.",
        }]);
        setStatus("Consulting the AI design director\u2026", "busy");
        return askRetryable(function () {
          return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: director });
        }).then(function (dResp) {
          SK.lastModelText = { text: dResp.text };
          var dParsed = tryParse(dResp.text);
          moConcept = (dParsed.ok && dParsed.kind === "reply" && dParsed.text) ? String(dParsed.text) : String(dResp.text || "");
          SK.moConcept = moConcept;   // the approval gate shows the director's concept
          if (dParsed.ok && dParsed.kind === "slide_ops" && dParsed.ops) return dParsed.ops; // director ignored the format: treat its ops as the design
          return null;
        }).then(function (straightOps) {
          if (straightOps && straightOps.length) return { ok: true, kind: "slide_ops", ops: straightOps, orchestrated: 1 };
          var total = Math.min(slideN, MO_MAX);
          var all = [];   // every op emitted so far (passes append in order)
          function coverageNote() {
            var have = {};
            all.forEach(function (op) { if (op && op.slide != null && STRUCT_TYPES[op.type]) have[op.slide] = 1; });
            var missing = [];
            for (var i = 1; i <= total; i++) if (!have[i]) missing.push(i);
            return missing.length ? "Slides still without structural design: " + missing.join(", ") + "." : "All slides have structural design.";
          }
          function passAsk(passName, spec, x, y) {
            if (SK.runCancelled) throw new FriendlyError("Run cancelled.");
            var bm = messages.concat([
              { role: "assistant", content: moConcept || "" },
              { role: "user",
                content: "CONCEPT (from the design director):\n" + (moConcept || "(decide the concept yourself)") +
                  "\n\nRealize the director's geometry templates EXACTLY: same-role slides get the " +
                  "same template and the same numbers - apply them as-is (fill in only content and " +
                  "small content-driven sizing tweaks). Do not invent new compositions per round - " +
                  "consistency comes from the templates, creativity came from the director." +
                  "\n\n" + spec +
                  (x != null ? ("\n\nThis round covers slides " + x + ".." + y + " ONLY (1-based outline numbers). Slide count in the deck: " + slideN + ".") : "\n\nThis round covers the WHOLE deck.") +
                  (all.length ? ("\n\nOPS SO FAR: " + all.length + " action(s) - do not repeat them. " + coverageNote()) : "") +
                  "\nReply with the slide_ops JSON only." },
            ]);
            setStatus(passName + (x != null ? " (slides " + x + ".." + y + ")" : "") + "\u2026", "busy");
            var fire = function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: bm }); };
            return askRetryable(fire).then(function (b1) {
              SK.lastModelText = { text: b1.text };
              var p1 = tryParse(b1.text);
              if (p1.ok) return p1;
              var rep1 = bm.concat([
                { role: "assistant", content: b1.text || "" },
                { role: "user", content: REPAIR_ASK },
              ]);
              return askRetryable(function () {
                return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: rep1 });
              }).then(function (b2) {
                SK.lastModelText = { text: b2.text };
                return tryParse(b2.text);
              });
            });
          }
          function collect(pk, x, y) {
            return (pk && pk.ok && pk.ops) ? pk.ops.filter(function (op) {
              return op && op.type !== "web_lookup" && (x == null || op.slide == null || (op.slide >= x && op.slide <= y));
            }) : [];
          }
          function chunkRangePass(passName, spec) {
            var seq = Promise.resolve([]);
            for (var cs = 1; cs <= total; cs += MO_CHUNK) {
              (function (x) {
                var y = Math.min(x + MO_CHUNK - 1, total);
                seq = seq.then(function (acc) {
                  return passAsk(passName, spec, x, y).then(function (pk) {
                    var ops = collect(pk, x, y);
                    var have = {};
                    ops.forEach(function (op) { if (op.slide != null && STRUCT_TYPES[op.type]) have[op.slide] = 1; });
                    var miss2 = [];
                    for (var sv = x; sv <= y; sv++) if (!have[sv]) miss2.push(sv);
                    if (!miss2.length) return acc.concat(ops);
                    // loop until done: slides with no structural design get their own round
                    setStatus("Completing slide(s) " + miss2.join(", ") + "\u2026", "busy");
                    return askRetryable(function () {
                      return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: messages.concat([
                        { role: "assistant", content: moConcept || "" },
                        { role: "user",
                          content: "Slide(s) " + miss2.join(", ") + " received no design ops. CONCEPT:\n" + (moConcept || "(your concept)") +
                            "\n\nEmit slide_ops covering ONLY slide(s) " + miss2.join(", ") + " - follow the concept, each with at least one structural op plus moves. Slide_ops JSON only." },
                      ]) });
                    }).then(function (t2) {
                      SK.lastModelText = { text: t2.text };
                      var tp = tryParse(t2.text);
                      var tops = (tp.ok && tp.ops) ? tp.ops : (salvageOps(t2.text) || []);
                      return acc.concat(ops, tops.filter(function (op) {
                        return op && op.slide != null && miss2.indexOf(op.slide) >= 0;
                      }));
                    });
                  });
                });
              })(cs);
            }
            return seq;
          }
          function wholePass(passName, spec, liveShapes) {
            return passAsk(passName, spec + (liveShapes ? "\n\n" + liveShapes : ""), null, null).then(function (pk) { return collect(pk, null, null); });
          }
          function liveShapeInventory(lo, hi) {
            // the REAL current shapes per slide (1-based numbers like the ops use)
            return SK.deck.getContext({ maxSlides: 80 }).then(function (c2) {
              var out = ["CURRENT SHAPES (live deck; the numbers ARE the 1-based shape indices to use in ops):"];
              c2.slides.forEach(function (sl) {
                if (sl.idx + 1 < (lo || 1) || (hi && sl.idx + 1 > hi)) return;
                var parts = (sl.shapes || []).map(function (sh) {
                  return "[" + (sh.idx + 1) + "] " + sh.type +
                    (sh.width && sh.height ? " (" + Math.round(sh.width) + "x" + Math.round(sh.height) + ")" : "");
                }).join(", ");
                out.push("Slide " + (sl.idx + 1) + ": " + (parts || "(no shapes)"));
              });
              return out.join("\n");
            }, function () { return "CURRENT SHAPES: (deck read failed - only use set_background images this round)"; });
          }
          // PASS 1 - STRUCTURE: the skeleton, chunked, completeness enforced
          var p1 = chunkRangePass("STRUCTURE PASS",
            "STRUCTURE PASS - build the skeleton following the concept: backgrounds (solid/gradient), panels, accent bars, overlays, split compositions, repositioning and rotation. NO image generation, NO text formatting yet (exception: set a text color when the text would otherwise sit unreadably on a new dark field).");
          // PASS 2 - POLISH: typography + readability on the existing structure
          var p2 = p1.then(function (acc) {
            all = acc;
            return wholePass("POLISH PASS",
            "POLISH PASS - typography and readability on the structure that now exists: format_text titles/body per the concept's scale, keep every text above the readability law, strengthen accents and hierarchy. NO new structural shapes, NO images, no repeating existing ops.").then(function (more) {
              all = all.concat(more);
              return all;
            });
          });
          // PASS 3 - ART: imagery only where the concept calls for it
          var p3 = p2.then(function () {
            return liveShapeInventory().then(function (inv) {
            return wholePass("ART PASS",
            "ART PASS - imagery only where the concept calls for it: AI picture backgrounds for cover and section slides (set_background kind:image), and cut-out picture fills inside EXISTING shapes for content slides. Use set_shape_image ONLY on shape indices from the CURRENT SHAPES list below - the roundedRectangle/ellipse placeholders created in the structure pass are exactly your cut-out spots. NEVER guess indices. If a slide has no suitable shape yet, skip it and note it in your ops, or add create_shape with a solid fill instead - but never set_shape_image on a missing shape. Quality over quantity; set_background images per the concept's cover/section plan." + inv).then(function (more) {
              all = all.concat(more);
              return all;
            });
          });
          });
          // PASS 4 - REVIEW LOOP: the model audits its own redesign and keeps
          // fixing until it declares the deck done (max 2 fix rounds)
          var p4 = p3.then(function (acc3) {
            all = acc3;
            function reviewRound(n) {
              if (SK.runCancelled) throw new FriendlyError("Run cancelled.");
              var rm = messages.concat([
                { role: "assistant", content: moConcept || "" },
                { role: "user",
                  content: "REVIEW ROUND for the make-over. CONCEPT:\n" + (moConcept || "(your concept)") +
                    "\n\nOPS SO FAR: " + all.length + " action(s) across the deck. " + coverageNote() +
                    "\n\nAudit the redesign deck-wide: same-role slides must share one composition (the " +
                    "director's template) - flag slides that deviate; check imagery the concept promised " +
                    "and every text against the readability law. Reply with slide_ops fixing ONLY the " +
                    "slides that deviate or lack something. If the deck matches the concept, reply {\"kind\":\"reply\",\"text\":\"done\"}." },
              ]);
              setStatus("Review pass " + n + " - auditing the redesign\u2026", "busy");
              return askRetryable(function () {
                return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: rm });
              }).then(function (rp) {
                SK.lastModelText = { text: rp.text };
                var pp = tryParse(rp.text);
                if (pp.ok && pp.kind === "reply") return all;   // the auditor says done
                var fixes = (pp.ok && pp.ops) ? pp.ops : (salvageOps(rp.text) || []);
                if (!fixes.length) return all;
                all = all.concat(fixes);
                if (n >= 2) return all;
                return reviewRound(n + 1);
              });
            }
            return reviewRound(1);
          });
          return p4.then(function (finalOps) {
            if (!finalOps || !finalOps.length) return null;
            SK.moPassInfo = "built in AI-driven passes: structure -> polish -> art, then review rounds to completion (" + finalOps.length + " actions)";
            return { ok: true, kind: "slide_ops", ops: finalOps, orchestrated: 1 };
          });
        });
      }

      var baseChain = function () {
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
              return { ok: true, kind: "slide_ops", ops: dedupeMergedOps(parsed.ops, extra), recovered: true };
            });
        }

        if (parsed.ok) return parsed;
        // escalate: up to 2 repair rounds with increasingly strong nudges.
        // Empty replies get the loudest one - an empty response must never end a run.
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
            // last resort: salvage complete ops from the longest truncated attempt,
            // then ask once for the remaining ops and merge
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
                return { ok: true, kind: "slide_ops", ops: dedupeMergedOps(salv, extra), recovered: true };
              });
          });
        });
      });
      };
      var parsedChain = fullMo ? makeOverFlow() : baseChain();
      return parsedChain.then(function (parsed) {
        if (parsed === null) {
          // orchestration produced nothing usable - fall back to the plain chain
          return baseChain();
        }
        return parsed;
      }).then(function (parsed) {
        if (parsed && parsed.kind !== "clarify" && !(SK.clarifyChain || 0) &&
            /infographics?\b/i.test(instruction) &&
            !/\b(image|picture|illustration|screenshot|photo|poster|graphic)\b/i.test(instruction)) {
          // 0.9.236: the word infographic ALWAYS surfaces the form question -
          // models answer with shapes first, the user asked for a real ask
          parsed = { ok: true, kind: "clarify",
                     question: "Infographic as an AI-generated image, as editable shapes + text, or both?",
                     options: ["AI image", "Editable shapes", "Both"] };
        }
        return parsed;   // chain continues into the target-guard / clarify steps
      }).then(function (parsed) {
        if (parsed && parsed.ok && parsed.kind === "slide_ops" && (parsed.ops || []).length) {
          // 0.9.238: the ask names a specific slide, but the reply aims its
          // ops at a DIFFERENT one (typically the open slide; hit live:
          // "turn slide 124 into an infographic" with slide 125 open landed
          // on 125). Re-ask ONCE with the reason; keep the original when the
          // model stays put.
          var namedTarget = (function () {
            var m2 = /(?:^|[\s(,.])[\w-]{1,20}\s*slides?\s+(\d{1,4})\b/i.exec(instruction) ||
                     /(?:^|[\s(,.])slides?\s+(\d{1,4})\b/i.exec(instruction);
            return m2 ? parseInt(m2[1], 10) : 0;
          })();
          // 0.9.271: an interactive target answer (the pick embeds this
          // marker) is the user's explicit instruction - no re-ask, no drop
          var userTargetOverride = /your earlier choice on "Where should the work land?[\s\S]*the slides the plan aims at/i.test(instruction || "");
          if (userTargetOverride) SK.insertGuard = null;
          if (namedTarget && ctx && ctx.slides && namedTarget <= ctx.slides.length && !userTargetOverride) {
            var tmap = {};
            var createCount0 = 0;
            (parsed.ops || []).forEach(function (op2) {
              if (op2 && op2.slide) tmap[op2.slide] = 1;
              if (op2 && op2.type === "create_slide") createCount0++;
            });
            var tl = Object.keys(tmap).map(Number);
            // 0.9.267: "insert slides after N" - targets N+1.. are CORRECT when
            // the batch itself creates them (the old guard forced the model to
            // re-aim AT the named slide and an after-insert plan then painted
            // the user's real slide). Exempt: ask uses after/insert wording,
            // the plan carries create_slide, and all targets fit
            // named+1 .. named+creates within the virtual deck.
            var afterW = /\bafter\b|\binsert\b|\bnew slides?\b|\badd slides?\b/i.test(instruction);
            var hostsOk = afterW && createCount0 > 0 && tl.length > 0 &&
              tl.every(function (t3) { return t3 <= ctx.slides.length + createCount0; }) &&
              tl.every(function (t3) { return t3 >= namedTarget && t3 <= namedTarget + createCount0; });
            if (tl.length && tl.indexOf(namedTarget) < 0 && !hostsOk) {
              dbg("target.mismatch", { named: namedTarget, ops: tl.join(","), afterW: afterW, creates: createCount0 });
              // 0.9.269: an insert-after ask re-asked with "aim everything at
              // slide N" taught the model to DROP its create_slide ops and
              // paint the user's real slide (field-verified x2). The nudge
              // must carry the insert intent instead.
              var reBody;
              if (afterW) {
                reBody = "WRONG TARGET: the user asked to INSERT new slides after slide " + namedTarget +
                  " - your plan writes to EXISTING slides and carries " + (createCount0 > 0 ? "create_slide ops that do not cover targets " : "NO create_slide ops. ") +
                  "Re-emit the plan with create_slide ops added after slide " + namedTarget +
                  " and aim every NEW-slide op at the slides YOU create (post-insert 1-based numbering); " +
                  "only ops meant for slide " + namedTarget + "'s own content may write to it. Slide_ops JSON only, same rules.";
              } else {
                reBody = "WRONG TARGET: the user's request names slide " + namedTarget + " (1-based, same numbering as the outline), but your ops target slide " +
                  tl.join(", ") + " (that is the currently open slide). Re-emit the SAME plan with every slide number aimed at slide " + namedTarget +
                  ". Slide_ops JSON only, same rules.";
              }
              var reMsg = messages.concat([
                { role: "assistant", content: String((SK.lastModelText || {}).text || "") },
                { role: "user", content: reBody },
              ]);
              // 0.9.270: stash the insert constraint for the VALIDATOR - the
              // hard drop must hold through repair/recovery re-validations
              SK.insertGuard = { named: namedTarget, base: ctx.slides.length };
              setStatus("The ask names slide " + namedTarget + " but the plan aims at slide " + tl.join(", ") + " - re-asking\u2026", "busy");
              return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: reMsg }); })
                .then(function (r4) {
                  dbg("target.reask.response", { text: r4.text, elapsedMs: r4.elapsedMs });
                  SK.lastModelText = { text: r4.text };
                  var p4 = tryParse(r4.text);
                  return p4.ok ? p4 : parsed;   // validation enforces SK.insertGuard
                }).catch(function () { return parsed; });
            }
          }
        }
        return parsed;
      }).then(function (parsed) {
        // 0.9.271 stage 1 (case B): the ask carries a targeting VERB but no
        // slide number, and the plan writes exactly ONE existing slide -
        // the model may have guessed. One AI confirmation round: the model
        // deliberately re-confirms (same plan) or asks the user (clarify).
        if (parsed && parsed.ok && parsed.kind === "slide_ops" && (parsed.ops || []).length &&
            !(SK.targetConfirmDone) && ctx && ctx.slides && instruction &&
            /\b(insert|add|remove|delete|move|replace|change|write|apply|place|put)\b/i.test(instruction) &&
            !/slides?\s+\d|\bslide\s+\d|\bafter\b/i.test(instruction)) {
          var tSet = [];
          var hasCreatesB = parsed.ops.some(function (o5) { return o5 && o5.type === "create_slide"; });
          parsed.ops.forEach(function (o4) {
            if (o4 && o4.slide != null && o4.slide >= 1 && o4.slide <= ctx.slides.length &&
                tSet.indexOf(o4.slide) < 0) tSet.push(o4.slide);
          });
          if (tSet.length === 1 && !hasCreatesB) {
            SK.targetConfirmDone = true;   // one confirmation round per run
            var confMsg = messages.concat([
              { role: "assistant", content: String((SK.lastModelText || {}).text || "") },
              { role: "user", content: "TARGET CONFIRMATION: the user's request does not state WHERE - " +
                "your plan writes to slide " + tSet[0] + " of " + ctx.slides.length + " (1-based). " +
                "If that is certainly the intended target, reply with the SAME plan unchanged. " +
                "If the request leaves the target open, ask the user: reply {\"kind\":\"clarify\",\"question\":\"...\",\"options\":[2-4 short choices]} " +
                "(or reply with a CORRECTED plan if the wording clearly implies another target). " +
                "Slide_ops JSON or clarify JSON only." },
            ]);
            setStatus("Confirming the target slide with the model…", "busy");
            return askRetryable(function () { return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: confMsg }); })
              .then(function (r5) {
                dbg("target.confirm", { text: String(r5.text || "").slice(0, 160) });
                SK.lastModelText = { text: r5.text };
                var p5 = tryParse(r5.text);
                if (p5.ok && p5.kind === "clarify") {
                  // the model agrees the target is open - interactive clarify
                  renderClarify(p5, instruction);
                  SK.lastInstruction = instruction;
                  setStatus("Pick one of the options (or write your own answer) to continue.", "pending");
                  pushHistory(instruction, assistantHistoryText(p5));
                  return Promise.resolve(null);
                }
                return renderParsed(p5.ok ? p5 : parsed, ctx, instruction);
              }).catch(function () { return renderParsed(parsed, ctx, instruction); });
          }
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
      setStatus(msg, "error");
      // the request left the pane but failed - keep it in the session too
      pushHistory(instruction, "(failed: " + msg + ")");
      // give the request back so the user can just hit Run again
      if (input && !input.value) input.value = instruction;
    }).then(function () {
      stopChatTimer();
      SK.busy = false;
      SK.actRunOn = false;
      setRunEnabled(true);
      input.disabled = false;
      // the box clears the moment the result (and its echoed request)
      // is in the pane; a failed run keeps the text + stays enabled
      if ($("chat-output") && $("chat-output").children.length) input.value = "";
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
    if (Array.isArray(obj)) obj = { ops: obj };
    if (obj && typeof obj === "object") {
      if (obj.kind === "clarify" && typeof obj.question === "string" && obj.question.trim()) {
        // 0.9.235: the model asks the user to pick the form of the result;
        // internal action, never gated - the pane renders the question and
        // the re-ask carries the user's decision
        return { ok: true, kind: "clarify", question: obj.question.trim().slice(0, 300),
                 options: (Array.isArray(obj.options) ? obj.options : []).slice(0, 4)
                   .map(function (o) { return String(o == null ? "" : o).trim(); }).filter(Boolean),
                 text: stripped };
      }
      if (obj.kind === "reply") {
        if (typeof obj.text === "string") return { ok: true, kind: "reply", text: obj.text };
        return { ok: true, kind: "reply", text: String(obj.text == null ? "" : JSON.stringify(obj.text)) };
      }
      if (obj.kind === "read_page" && typeof obj.url === "string" && /^https?:\/\//i.test(obj.url)) {
        // an internal page-read intent: the pane fetches the page text and
        // re-asks (free, never gated)
        return { ok: true, kind: "read_page", url: obj.url.split("#")[0],
                 from: (typeof obj.from === "number" && obj.from > 0) ? Math.floor(obj.from) : 0,
                 ops: Array.isArray(obj.ops) ? obj.ops : null, text: stripped };
      }
      if (obj.kind === "web_lookup") {
        // an internal lookup intent: no ops, no approval - the pane runs the
        // search profile and re-asks (this must parse BEFORE the repair
        // chain, or the model gets nagged for "format" instead of results)
        // Models sometimes emit look-up-plus-actions in one reply: keep the
        // ops so the runner can apply them after the lookup round.
        return { ok: true, kind: "web_lookup", ops: Array.isArray(obj.ops) ? obj.ops : null, text: stripped };
      }
      if (obj.kind === "read_slide") {
        var rs = [];
        var src = Array.isArray(obj.slides) ? obj.slides : (typeof obj.slide === "number" ? [obj.slide] : []);
        src.forEach(function (n2) {
          if (typeof n2 === "number" && Math.floor(n2) === n2 && n2 >= 1 && rs.length < 12 && rs.indexOf(n2) < 0) rs.push(n2);
        });
        if (rs.length) return { ok: true, kind: "read_slide", slides: rs, text: stripped };
      }
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
    SK.lastInstruction = instruction;
    SK.opProgress = null;     // a fresh run gets fresh progressive tracking
    var out = $("chat-output");
    out.innerHTML = "";
    wirePastHeader();
    // the request prints above the response/summary (echoed by the clear)
    var reqEcho = document.createElement("div");
    reqEcho.className = "reqecho";
    reqEcho.textContent = instruction;
    makeReqEchoCopyable(reqEcho, instruction);
    out.appendChild(reqEcho);
    if (!parsed.ok) {
      var box = document.createElement("div");
      box.className = "reply-text";
      var showTxt = String(parsed.text || "");
      if (showTxt.length > 420) showTxt = showTxt.slice(0, 420) + "\u2026 (truncated preview - enable debug logging for the full raw output)";
      box.textContent = showTxt || "(empty response)";
      out.appendChild(box);
      var note = document.createElement("div");
      note.className = "opnote";
      note.textContent = "Shown as text: no action JSON could be recovered (SlideKick re-asked once). Enable debug logging (Settings) to inspect the raw output.";
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
    var validation = validateOps(parsed.ops, ctx, parsed.preInvalid);
    SK.contrastFixes = adjustForContrast(validation.ops);   // mechanical readability guard
    if (!validation.ops.length && !validation.invalid.length && (SK.pendingWebOps || []).length) {
      // the look-up-then-act ops attached to the intent = the real payload
      parsed = { ok: true, kind: "slide_ops", ops: SK.pendingWebOps };
      SK.pendingWebOps = null;
      validation = validateOps(parsed.ops, ctx);
      SK.contrastFixes = adjustForContrast(validation.ops);
    }
    if (!validation.ops.length && !validation.invalid.length) {
      var box3 = document.createElement("div");
      box3.className = "reply-text";
      box3.textContent = (SK.webResults || SK.pendingWebOps) ?
        "(model returned an empty action list - web lookup ran; the model chose not to change the deck)" :
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
      var autoDone = document.createElement("div");
      autoDone.className = "doneinline";
      var autoDoneN = 0;
      autoBox.appendChild(autoDone);
      out.appendChild(autoBox);
      SK.opProgress = function (op, addedLines, ok) {
        if (!ok) return;
        autoDoneN++;
        var dr = document.createElement("div");
        dr.className = "oprowline1 doneline";
        dr.textContent = "\u2713 " + ((OPS_META[op.type] || { label: op.type }).label) +
          (op.slide != null ? " - slide " + op.slide : "");
        autoDone.insertBefore(dr, autoDone.firstChild);
      };
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
  function copyIconSvg() {
    return '<svg width="12" height="12" viewBox="0 0 12 12" style="vertical-align:-2px" aria-hidden="true">' +
      '<rect x="1" y="1" width="8" height="8" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.4"/>' +
      '<rect x="3.5" y="3.5" width="7.5" height="7.5" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.4"/>' +
    '</svg>';
  }
  function copyTextToClipboard2(txt, done) {
    try {
      var ta2 = document.createElement("textarea");
      ta2.value = txt; ta2.style.position = "fixed"; ta2.style.opacity = "0";
      document.body.appendChild(ta2); ta2.select();
      document.execCommand("copy");
      document.body.removeChild(ta2);
      if (done) done(true); return;
    } catch (eC) { if (done) done(false); }
  }
  function manualNotesBox(res) {
    var mb = document.createElement("div");
    mb.className = "manualbox";
    mb.appendChild(document.createElement("h4")).textContent =
      "Notes could not be written automatically (this Office build blocks notes-page edits) - copy each into the notes pane:";
    res.manualNotes.forEach(function (mn) {
      var wrap = document.createElement("div");
      wrap.style.display = "flex"; wrap.style.alignItems = "flex-start"; wrap.style.gap = "6px";
      var ta = document.createElement("textarea");
      ta.value = "Slide " + (mn.slide + 1) + ": " + mn.text;
      ta.style.flex = "1 1 auto";
      var btn = document.createElement("button");
      btn.className = "btn subtle";
      btn.type = "button";
      btn.innerHTML = copyIconSvg();
      btn.title = "Copy these notes to the clipboard";
      btn.addEventListener("click", function () {
        copyTextToClipboard2(ta.value, function (okc) {
          btn.textContent = okc ? "\u2713 copied" : "copy failed";
          btn.disabled = true;
          setTimeout(function () { btn.innerHTML = copyIconSvg(); btn.disabled = false; }, 1800);
        });
      });
      wrap.appendChild(btn); wrap.appendChild(ta);
      mb.appendChild(wrap);
    });
    return mb;
  }

  function renderApplyResult(result, res) {
    result.innerHTML = "";
    res.lines.forEach(function (l) {
      var d = document.createElement("div");
      d.className = "resultline " + (l.ok ? "ok" : "fail");
      d.textContent = l.text;
      result.appendChild(d);
    });
    if (res.manualNotes.length) result.appendChild(manualNotesBox(res));
  }

  function hexLum(hex) {
    var m = /^#([0-9a-fA-F]{6})$/.exec(String(hex || "").trim());
    if (!m) return null;
    var v = parseInt(m[1], 16);
    var r = (v >> 16 & 255) / 255, g = (v >> 8 & 255) / 255, b = (v & 255) / 255;
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  }
  function adjustForContrast(ops) {
    var fixes = 0;
    var fieldLum = {};   // 1-based slide -> luminance of the last planned solid background
    ops.forEach(function (op) {
      if (op.type === "set_background" && op.kind === "solid" && op.color != null) {
        var L = hexLum(op.color);
        if (L != null) fieldLum[op.slide] = L;
        return;
      }
      if (op.type === "set_background" && op.kind != null && op.kind !== "solid") {
        delete fieldLum[op.slide];   // photo: contrast handled by overlays, not guessing
      }
      if (op.type === "format_text" && op.font && op.font.color != null && op.slide != null) {
        var field = fieldLum[op.slide];
        var c = hexLum(op.font.color);
        if (field == null || c == null) return;
        if (field < 0.4 && c < 0.4) { op.font.color = "#F7F5F0"; op._contrastFix = 1; fixes++; }
        else if (field > 0.6 && c > 0.6) { op.font.color = "#17171B"; op._contrastFix = 1; fixes++; }
      }
    });
    return fixes;
  }

  function renderOpsPreview(out, validation, ctx) {
    var caps = SK.deck.capabilities();
    var box = document.createElement("div");
    box.className = "opsbox";
    var h = document.createElement("h3");
    syncAutoApplyCheckbox();
    h.textContent = "Proposed actions" + ($("chat-autoapply").checked ? " (auto-apply)" : "");
    box.appendChild(h);

    if (SK.moConcept) {
      var moN = document.createElement("div");
      moN.className = "opnote";
      moN.textContent = "AI design director concept applied: " +
        String(SK.moConcept).slice(0, 110) + (String(SK.moConcept).length > 110 ? "\u2026" : "");
      box.appendChild(moN);
      SK.moConcept = "";
    }
    if (SK.contrastFixes > 0) {
      var cfN = document.createElement("div");
      cfN.className = "opnote";
      cfN.textContent = "Readability guard: " + SK.contrastFixes +
        " text color(s) automatically adjusted for contrast with their background.";
      box.appendChild(cfN);
    }
    if (SK.moPassInfo) {
      var passN = document.createElement("div");
      passN.className = "opnote";
      passN.textContent = SK.moPassInfo;
      box.appendChild(passN);
      SK.moPassInfo = "";
    }
    if (SK.lastRecovered) {
      var recN = document.createElement("div");
      recN.className = "opnote";
      recN.textContent = "Recovered " + SK.lastRecovered + " action(s) from a cut-off reply; the model was asked for the rest." +
        (SK.recoveredDupes > 0 ? " " + SK.recoveredDupes + " duplicated action(s) from the continuation were removed." : "") +
        " If some slides are still untouched, ask to continue.";
      box.appendChild(recN);
      SK.lastRecovered = 0;
      SK.recoveredDupes = 0;
    }

    if (!caps.text) {
      var warn = document.createElement("div");
      warn.className = "status error";
      warn.textContent = "This PowerPoint build lacks the text-editing APIs (PowerPointApi ~1.8). SlideKick can read but not modify this deck. Update Microsoft 365 to enable edits.";
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
      // 0.9.216 - don't reject fast: re-ask the model WITH the reasons first
      var reasons = (validation.invalid || []).map(function (entry) {
        var ot = (entry.op && entry.op.type) || "?";
        var sl = (entry.op && entry.op.slide) ? " slide='" + entry.op.slide + "'" : "";
        return "  - " + ot + sl + ": " + (entry.reason || "unknown");
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
            "Slide/shape numbers must come from the outline you were given (1-based); use only the allowed op types; keep geometry fields as numbers." },
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
      // nothing survived validation: no per-row wall of failures - one clear
      // box + a one-click re-ask of the same instruction
      // 0.9.271: when the target guard dropped everything on an insert-ask
      // (the model kept the wrong aim after its confirmation round), the
      // remaining ambiguity is the USER's to decide - interactive clarify.
      var guardDrops = validation.invalid.filter(function (iv) {
        return String(iv.reason || "").indexOf("target guard:") === 0;
      });
      if (!validation.ops.length && guardDrops.length && guardDrops.length === validation.invalid.length &&
          SK.insertGuard && SK.insertGuard.named != null) {
        var aimSet = [];
        guardDrops.forEach(function (gd) {
          var s3 = gd.op && gd.op.slide;
          if (s3 != null && aimSet.indexOf(s3) < 0) aimSet.push(s3);
        });
        var qText = "Where should the work land? Your ask names slide " + SK.insertGuard.named +
          ", and the plan aimed at slide(s) " + aimSet.join(", ") +
          " - the confirmation round did not resolve it.";
        var opts = ["slide " + SK.insertGuard.named + " (as my ask said)",
                    "the slides the plan aims at - " + aimSet.join(", "),
                    "cancel - do nothing"];
        var clP = { kind: "clarify", question: qText, options: opts };
        SK.insertGuard = null;   // the answered choice rebuilds the guard on the re-run
        SK.lastInstruction = SK.lastInstruction || instruction;
        renderClarify(clP, SK.lastInstruction);
        setStatus("Confirm where to apply the plan.", "pending");
        pushHistory(SK.lastInstruction, assistantHistoryText(clP));
        return;
      }
      var allbad = document.createElement("div");
      allbad.className = "status error";
      allbad.textContent = "None of the " + validation.invalid.length + " action(s) survived validation" +
        (SK.lastRecovered ? " (they came from a cut-off reply)" : "") +
        " - the slides/shapes they target do not match this deck. No changes were applied.";
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
      if (row.op) { row.op._rowEl = r; row.op._gateRow = row; }   // progressive apply removes this row
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
      if (row.op.notesOffConfirm) {
        text += "  - note generation is OFF in Settings; ticking this writes them anyway for this slide";
      }
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
        row._lab = lab;
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
        if (!sid2) { try { sid2 = localStorage.getItem("sk-resume-slidekick"); } catch (eE1) {} }
        localStorage.setItem("sk-approvals-slidekick", JSON.stringify({ sid: sid2 || "", map: SK.sessionApprovals }));
      } catch (eAp) { /* storage may be unavailable */ }
      // unapproved confirm rows would silently vanish - ask before skipping
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
      // progressive tracking: executed actions leave the proposal list and grow
      // into a completed list one by one
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
          dr.textContent = "\u2713 " + ((OPS_META[op.type] || { label: op.type }).label) +
            (op.slide != null ? " - slide " + op.slide : "");
          doneList.insertBefore(dr, doneList.firstChild);   // newest on top
          if (op._rowEl && op._rowEl.parentNode) op._rowEl.parentNode.removeChild(op._rowEl);
          if (op._gateRow) op._gateRow.done = true;
        } else if (op._rowEl && op._rowEl.parentNode) {
          op._rowEl.classList.add("opfailrow");
        }
        updateApplyLabel();   // numbers track the live (not-yet-done) rows
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
        // everything executed cleanly? close the proposal section; only the
        // completed list (+ any manual-notes copy boxes) stays
        var liveRows = rows.filter(function (r) { return !r.removed && !r.done; }).length;
        if (!liveRows) {
          if (res.manualNotes.length) doneBox.appendChild(manualNotesBox(res));
          box.parentNode.removeChild(box);   // the done box is a sibling
          SK.opProgress = null;
          setStatus("Applied " + chosen.length + " op(s).", "ok");
          var f2 = res && res.lines ? res.lines.filter(function (l) { return !l.ok; }).length : 0;
          if (f2) setStatus("Applied with " + f2 + " error(s).", "error");
          renderSimDeck();
          updateDeckSummary();
          dbg("ops.applied", { lines: res.lines, closedGate: true });
          return;
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
    SK.deck.getContext({ maxSlides: 1 }).then(function (ctx) {
      var sel = ctx.selected && ctx.selected.length ? " - selected " + ctx.selected.join(",") : "";
      el.textContent = "Deck: " + ctx.slideCount + " slide(s)" + sel + " - " + ctx.host;
    }).catch(function (err) {
      el.textContent = "Deck: unreadable (" + err.message + ")";
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
    populateModelSelect("prof-content", p.content && p.content.connection, p.content && p.content.model);
    populateModelSelect("prof-other", p.other && p.other.connection, p.other && p.other.model);
    SK.editingProfImage = p.image || null;
    fillConnSelect($("prof-image-conn"), p.image && p.image.connection);
    populateModelSelect("prof-image", p.image && p.image.connection, p.image && p.image.model, SK.profImageFilter ? "image" : null);
    SK.editingProfWeb = p.web || null;
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
    $("prof-content-model").innerHTML = '<option value="">(choose a connection first)</option>';
    $("prof-other-model").innerHTML = '<option value="">(choose a connection first)</option>';
    $("prof-image-model").innerHTML = '<option value="">(optional - choose a connection first)</option>';
    $("prof-content-custom").classList.add("hidden");
    $("prof-other-custom").classList.add("hidden");
    $("prof-image-custom").classList.add("hidden");
    SK.editingProfImage = null;
    SK.editingProfWeb = null;
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
    $("settings-notes").checked = !!SK.cfg.generateNotes;
    $("debug-path").textContent = SK.cfg.configPath.replace(/config\.json$/, "logs") + "/debug-YYYY-MM-DD.log";
  }

  function saveApprovals() {
    var approvals = {};
    OP_TYPES.forEach(function (t) {
      var el = document.querySelector('input[name="appr-' + t + '"]:checked');
      approvals[t] = el ? el.value : "confirm";
    });
    apiJson("/api/settings", "PUT", { approvals: approvals, generateNotes: $("settings-notes").checked }).then(function () {
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
  function saveNotesOnly() {
    var on = !!$("settings-notes").checked;
    apiJson("/api/settings", "PUT", { generateNotes: on }).then(function () {
      if (SK.cfg) SK.cfg.generateNotes = on;
      setStatus("Notes setting: " + (on ? "speaker notes ON" : "speaker notes OFF"), "ok");
    }).catch(function (e) { setStatus("Could not save the notes setting: " + ((e && e.message) || e), "error"); });
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
  function renderAbout() {
    var kv = [
      ["Pane build (JS)", window.__SK_BUILD || "?"],
      ["Add-in identity (this manifest)", (SK.cfg && SK.cfg.manifests && SK.cfg.manifests["slidekick"]) ? "v" + SK.cfg.manifests["slidekick"] : "?"],
      ["All manifest identities", (SK.cfg && SK.cfg.manifests) ? Object.keys(SK.cfg.manifests).map(function (k5) { return k5 + " v" + SK.cfg.manifests[k5]; }).join(" / ") : "?"],
      ["Office host", (function () { try { return Office.context.diagnostics.host + " " + Office.context.diagnostics.version; } catch (e5) { return "n/a"; } })()],
      ["SlideKick add-in", SK.cfg ? "v" + (SK.cfg.version || "0.1.3") : "0.1.3"],
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
      $("statusbar").textContent = "SlideKick v" + (cfg.version || "0.1.0") +
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
    $("chat-refresh-deck").addEventListener("click", function () { updateDeckSummary(); setStatus("Deck refreshed.", ""); });
    // merge another deck: file picker -> base64 -> insert all its slides at the end
    $("merge-deck").addEventListener("click", function () {
      var fi = $("merge-file");
      if (fi) fi.click();
    });
    $("merge-file").addEventListener("change", function () {
      var fi = $("merge-file");
      var f = fi.files && fi.files[0];
      if (!f) return;
      var name = f.name;
      if (!/\.pptx$/i.test(name)) { setStatus("Only .pptx files can be merged.", "error"); fi.value = ""; return; }
      if (!window.confirm("Merge ALL slides of '" + name + "' into this deck (appended at the end)?")) { fi.value = ""; return; }
      var reader = new FileReader();
      reader.onerror = function () { setStatus("Could not read '" + name + "'.", "error"); fi.value = ""; };
      reader.onload = function () {
        var b64 = String(reader.result || "");
        b64 = b64.slice(b64.indexOf(",") + 1);   // strip data: URI prefix
        setStatus("Merging slides from '" + name + "'…", "busy");
        SK.deck.mergeDeck(b64).then(function (res) {
          var cnt = res && res.count ? res.count : null;
          setStatus(cnt ? ("Merged " + cnt + " slide(s) from '" + name + "'.")
                        : "Merged slides from '" + name + "' appended to the deck.", "ok");
          return updateDeckSummary();
        }).then(function () { renderSimDeck(); fi.value = ""; })
          .catch(function (err) {
            setStatus("Merge failed: " + ((err && err.message) || err), "error");
            fi.value = "";
          });
      };
      reader.readAsDataURL(f);
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
    $("settings-notes").addEventListener("change", function () { saveNotesOnly(); flashSaved("saved-notes", "\u2713 saved"); });
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
    // a pane reload lands on Settings -> General (the caller tags the URL)
    try {
      var q = new URLSearchParams(location.search);
      if (q.get("reloadnav") === "general") {
        switchTab("settings");
        switchSettingsSub("general");
        q.delete("reloadnav");   // not sticky: a plain F5 keeps the normal boot
        var clean = location.pathname + ("" + q ? "?" + q.toString() : "");
        try { history.replaceState(null, "", clean); } catch (eH) {
          location.replace(clean);   // WebViews without replaceState
        }
      }
    } catch (eQR) { /* older parser: leave the boot view */ }

    // a pane reload (or back-end restart) resumes the session that was active
    var bootResume = null;
    try { bootResume = localStorage.getItem("sk-resume-session-slidekick"); } catch (eS) {}
    if (bootResume) resumeSession(bootResume);
    else restoreSessionApprovals();   // a plain reload restores the same-session approvals
    updateSessionLabel();
    wireSessions();
    wireSearchProfiles();
    $("req-clear").addEventListener("click", clearRequests);
    $("req-refresh").addEventListener("click", renderRequests);
    $("sim-reset").addEventListener("click", function () {
      window.SimDeck.resetAll();
      renderSimDeck();
      updateDeckSummary();
      setStatus("Simulator deck reset.", "");
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
      badge.textContent = "PPT";
      SK.deck = new OfficeDeck();
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
      el.textContent = "SlideKick v" + ver +
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
    setStatus("Cannot reach the SlideKick server: " + err.message + " Is it running? (python3 install.py)", "error");
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
      nav.textContent = "Sim deck";
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