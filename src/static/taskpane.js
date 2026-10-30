/*
 * SlideKick task pane (runs inside PowerPoint via Office.js, or in the
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

  var LAYOUT_KINDS = ["title_content", "title_only", "section", "blank"];

  var OPS_META = {
    create_slide: { label: "create slide", hint: "Appends a new slide (Office.js can only append; drag into place if needed)." },
    edit_slide: { label: "edit slide", hint: "Replaces title / body / speaker notes of an existing slide." },
    replace_text: { label: "replace text", hint: "Rewrites the text of one shape in place." },
    set_notes: { label: "set notes", hint: "Writes speaker notes on a slide." },
    create_table: { label: "create table", hint: "Adds a table (new slide when slide is empty; position/size in points)." },
    move_shape: { label: "move/resize shape", hint: "Sets left/top/width/height (points) of one shape - design passes use this." },
    apply_layout: { label: "apply layout", hint: "Applies a different slide layout (design change, text kept)." },
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

  function apiJson(path, method, body) {
    var opts = { method: method || "GET", headers: {} };
    if (body !== undefined) {
      opts.headers["Content-Type"] = "application/json";
      opts.body = JSON.stringify(body);
    }
    return fetch(path, opts).then(function (resp) {
      return resp.text().then(function (text) {
        var data = null;
        try { data = text ? JSON.parse(text) : {}; } catch (e) { /* non-json */ }
        if (!resp.ok) {
          var msg = data && data.error ? data.error : "HTTP " + resp.status;
          throw new FriendlyError(msg);
        }
        return data || {};
      });
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
    "You are SlideKick, an AI assistant inside a PowerPoint task pane.",
    "You get the current deck outline and a user request. Reply with ONLY one",
    "JSON object - never any commentary, never code fences, never reasoning or",
    "chain-of-thought. Format A (conversational):",
    '{"kind":"reply","text":"your answer"}',
    "Format B (deck changes):",
    '{"kind":"slide_ops","ops":[ ... ]}',
    "",
    "Allowed op types:",
    '- {"type":"create_slide","layout":"title_content|title_only|section|blank","title":"...","body":["bullet",...],"notes":"..."} (body/notes optional; new slides are appended at the end of the deck)',
    '- {"type":"edit_slide","slide":N,"title":"..."|null,"body":[...]|null,"notes":"..."|null} (null means: leave unchanged)',
    '- {"type":"replace_text","slide":N,"shape":N,"text":"..."} (reword one shape; N refers to the shape numbers listed for that slide)',
    '- {"type":"set_notes","slide":N,"notes":"..."}',
    '- {"type":"create_table","slide":N|null,"values":[["h1","h2"],["x","y"]],"left":40,"top":150,"width":640} (all cells strings, rows rectangular; slide null = new title slide first; left/top/width optional points)',
    '- {"type":"move_shape","slide":N,"shape":N,"left":40,"top":150,"width":300,"height":60} (points; only the fields to change; keep shapes inside the slide bounds shown in the outline)',
    '- {"type":"apply_layout","slide":N,"layout":"title_content|title_only|section|blank"} (only when the user asks for layout/design changes)',
    '- {"type":"delete_slide","slide":N}',
    "",
    "Rules:",
    "- Slide and shape indices are zero-based and MUST come from the outline below. Never invent them.",
    "- One body element = one bullet/paragraph; keep bullets short unless told otherwise.",
    "- Reword/tighten requests: only rewrite wording, keep the meaning and slide count.",
    "- If the request is a question/explanation/brainstorm rather than a deck change, use format A.",
    "- Design passes ('design this deck/slide', improve alignment/balance): emit ops (apply_layout, move_shape, create_table) computed from the geometry in the outline; when the user asks only for advice, use format A.",
  ].join("\n");

  var NOTES_OFF_RULE =
    "- Speaker-notes generation is DISABLED for this user (Settings). Never generate notes:" +
    " omit the notes field in create_slide/edit_slide ops and never emit set_notes ops." +
    " If the user explicitly asks for speaker notes, use format A (reply) and tell them to" +
    " enable 'Generate speaker notes' in Settings.";

  var REPAIR_ASK =
    "Your previous answer did not follow the required format. " +
    "Reply again with ONLY the JSON object (format A or B) and nothing else: " +
    'either {"kind":"reply","text":"..."} or {"kind":"slide_ops","ops":[...]} - ' +
    "no reasoning, no fences, no commentary.";

  var HISTORY_MAX_EXCHANGES = 6; // oldest turns are dropped, newest kept
  var HISTORY_MAX_CHARS = 2000;  // per stored message

  var CONTENT_ROUTE_RE = new RegExp(
    "\\b(slide|slides|deck|bullet|bullets|reword|rewrite|rephrase|outline|notes?|title|titles|add|create|make|generate|tighten|restructure|expand|insert|append|design|align|balance|spacing|table|tables|chart)\\b", "i"
  );

  function historyCap(text) {
    var s = String(text || "");
    return s.length > HISTORY_MAX_CHARS ? s.slice(0, HISTORY_MAX_CHARS) + " [truncated]" : s;
  }

  function pushHistory(userText, assistantText) {
    SK.history.push({ role: "user", content: historyCap(userText) });
    SK.history.push({ role: "assistant", content: historyCap(assistantText) });
    while (SK.history.length > HISTORY_MAX_EXCHANGES * 2) SK.history.shift();
    updateChatMeta();
    dbg("history.push", { exchanges: Math.floor(SK.history.length / 2) });
  }

  function resetHistory() {
    SK.history = [];
    updateChatMeta();
    dbg("history.reset", {});
  }

  function updateChatMeta() {
    var el = $("chat-history-note");
    if (!el) return;
    var n = Math.floor(SK.history.length / 2);
    el.textContent = n ? "AI sees the last " + n + " exchange(s)" : "new conversation";
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
      var limit = Math.min(opts.maxSlides || 60, 200);
      return self.run("getContext", function (ctx) {
        var pres = ctx.presentation;
        var slides = pres.slides;
        slides.load("items/id");
        return ctx.sync().then(function () {
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

  OfficeDeck.prototype.moveShape = function (slideIndex, shapeIndex, geometry) {
    return this.run("moveShape", function (ctx) {
      var shape = ctx.presentation.slides.getItemAt(slideIndex).shapes.getItemAt(shapeIndex);
      ["left", "top", "width", "height"].forEach(function (p) {
        if (geometry[p] != null) shape[p] = geometry[p];
      });
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

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
              throw new FriendlyError(
                role === "title"
                  ? "No title placeholder found on slide " + i
                  : "No body/content placeholder found on slide " + i
              );
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
  function validateOps(rawOps, ctx) {
    var ops = [], invalid = [];
    var notesDropped = 0;
    var notesOff = !(SK.cfg && SK.cfg.generateNotes);
    (rawOps || []).forEach(function (op, index) {
      if (!op || typeof op !== "object") { invalid.push({ op: op, reason: "not an object" }); return; }
      var t = op.type;
      if (OP_TYPES.indexOf(t) < 0) { invalid.push({ op: op, reason: "unknown op type '" + t + "'" }); return; }
      if (!ctx && t !== "create_slide") {
        invalid.push({ op: op, reason: "needs deck context (enable 'include deck context')" });
        return;
      }
      var bad = function (reason) { invalid.push({ op: op, reason: reason }); };
      var slideOk = function (v) { return typeof v === "number" && Math.floor(v) === v && v >= 0 && (!ctx || v < ctx.slideCount); };
      var shapeOk = function (sl, sh) {
        if (!ctx || !ctx.slides[sl]) return false;
        return (ctx.slides[sl].shapes || []).some(function (s) { return s.idx === sh && s.text != null; });
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
        if (typeof op.text !== "string") { bad("missing text"); return; }
        out.slide = op.slide; out.shape = op.shape; out.text = op.text;
      } else if (t === "set_notes") {
        if (notesOff) { notesDropped++; return; } // speaker-notes generation disabled
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
        var geoT = {};
        ["left", "top", "width", "height"].forEach(function (p2) {
          if (typeof op[p2] === "number" && op[p2] >= 0) geoT[p2] = op[p2];
        });
        out.geometry = geoT;
      } else if (t === "move_shape") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        var shapeInCtx = ctx && ctx.slides[op.slide] &&
          (ctx.slides[op.slide].shapes || []).some(function (s) { return s.idx === op.shape; });
        if (!shapeInCtx) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        var anyM = false, geoM = {};
        ["left", "top", "width", "height"].forEach(function (p3) {
          if (typeof op[p3] === "number") { geoM[p3] = op[p3]; anyM = true; }
        });
        if (!anyM) { bad("nothing to change (no left/top/width/height)"); return; }
        out.slide = op.slide; out.shape = op.shape; out.geometry = geoM;
      } else if (t === "apply_layout") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = op.slide; out.layout = normalizeLayout(op.layout);
      } else if (t === "delete_slide") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = op.slide;
      }
      ops.push(out);
    });
    return { ops: ops, invalid: invalid, notesDropped: notesDropped };
  }

  function approvalMode(opType) {
    return SK.sessionApprovals[opType] || (SK.cfg && SK.cfg.approvals && SK.cfg.approvals[opType]) || "confirm";
  }

  function fmtOp(op, ctx) {
    var slideTitle = function (sl) {
      if (!ctx || !ctx.slides[sl]) return "";
      var t = ctx.slides[sl].title;
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
        return "Slide " + op.slide + ", shape " + op.shape + ": -> " + op.text;
      case "set_notes":
        return "Slide " + op.slide + slideTitle(op.slide) + ": speaker notes (" + op.notes.length + " chars)";
      case "create_table":
        return (op.slide != null ? "Slide " + op.slide + slideTitle(op.slide) : "New slide") +
          ': table ' + op.values.length + "x" + op.values[0].length +
          (op.geometry && op.geometry.left != null ? " @ " + Math.round(op.geometry.left) + "," + Math.round(op.geometry.top) : "");
      case "move_shape":
        return "Slide " + op.slide + ", shape " + op.shape + ": " +
          Object.keys(op.geometry).map(function (p) { return p + "->" + Math.round(op.geometry[p]); }).join(" ");
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
    var runOne = function (op) {
      var cur = op.slide !== undefined ? op.slide - shift : undefined;
      switch (op.type) {
        case "create_slide":
          return deck.addSlide({ layout: op.layout, title: op.title, body: op.body }).then(function (r) {
            var note = "";
            if (!SK.sim && ctx && op.after != null && op.after < ctx.slideCount - 1) {
              note = " (requested position not supported - appended at end)";
            }
            lines.push({ ok: true, text: "create slide -> appended at index " + r.index + note });
            if (op.notes && op.notes.trim()) {
              return deck.setNotes(r.index, op.notes).then(function (nr) {
                if (nr && nr.manual) manualNotes.push({ slide: r.index, text: op.notes, reason: nr.reason });
                else lines.push({ ok: true, text: "set notes on new slide " + r.index });
              });
            }
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
        case "replace_text":
          return deck.setShapeText(cur, op.shape, op.text).then(function () {
            lines.push({ ok: true, text: "replace text on slide " + op.slide + " shape " + op.shape });
          });
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
          return deck.moveShape(cur, op.shape, op.geometry).then(function () {
            lines.push({ ok: true, text: "move shape " + op.shape + " on slide " + op.slide + ": " +
              Object.keys(op.geometry).map(function (p) { return p + "->" + Math.round(op.geometry[p]); }).join(" ") });
          });
        case "set_notes":
          return deck.setNotes(cur, op.notes).then(function (nr) {
            if (nr && nr.manual) manualNotes.push({ slide: cur, text: op.notes, reason: nr.reason });
            else lines.push({ ok: true, text: "set notes on slide " + op.slide });
          });
        case "apply_layout":
          return deck.applyLayout(cur, op.layout).then(function () {
            lines.push({ ok: true, text: "apply layout '" + op.layout + "' to slide " + op.slide });
          });
        case "delete_slide":
          return deck.deleteSlide(cur).then(function () {
            lines.push({ ok: true, text: "delete slide " + op.slide + " (" + fmtOp(op, ctx) + ")" });
            shift++;
          });
      }
      return Promise.resolve();
    };

    var seq = Promise.resolve();
    ops.forEach(function (op) {
      seq = seq.then(function () { return runOne(op); }).catch(function (err) {
        var msg = err && err.friendly ? err.message : (err && err.message) || String(err);
        lines.push({ ok: false, text: op.type + ": " + msg });
        if (!(err && err.friendly)) {
          lines.push({ ok: false, text: "unexpected error - stopping remaining ops" });
          throw err; // aborts the chain below
        }
      });
    });
    return seq.then(
      function () { return { lines: lines, manualNotes: manualNotes }; },
      function (err) {
        var msg = err && err.friendly ? err.message : (err && err.message) || String(err);
        lines.push({ ok: false, text: "aborted: " + msg });
        return { lines: lines, manualNotes: manualNotes };
      }
    );
  }

  // ------------------------------------------------------------------ //
  // Chat / prompt flow
  // ------------------------------------------------------------------ //
  function formatContext(ctx) {
    var lines = [];
    var sel = ctx.selected && ctx.selected.length ? " (zero-based, * = selected: " + ctx.selected.join(",") + ")" : "";
    lines.push("DECK OUTLINE: " + ctx.slideCount + " slides" + sel + (ctx.truncated ? " (only the first slides are listed)" : "") + ", host: " + ctx.host);
    ctx.slides.forEach(function (s) {
      var star = ctx.selected && ctx.selected.indexOf(s.idx) >= 0 ? "*" : "";
      var title = (s.shapes || []).filter(function (sh) { return sh.isTitle || (sh.text && /title/i.test(sh.name)); })[0];
      lines.push("[" + s.idx + star + "] " + (title && title.text ? 'TITLE: "' + title.text + '"' : "(no title)"));
      (s.shapes || []).forEach(function (sh) {
        if (sh.text == null) return;
        var role = sh.isTitle || (title && sh.idx === title.idx) ? "title" : sh.isBody ? "body" : "text";
        var txt = String(sh.text).replace(/\r|\n/g, " | ");
        if (txt.length > 220) txt = txt.slice(0, 220) + "…";
        if (!txt.trim()) return;
        lines.push("    shape " + sh.idx + " (" + role + ", " + sh.name + '): "' + txt + '"');
      });
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
    if (!prof) { setStatus("No profile yet - create one on the Profiles tab first.", "error"); switchTab("profiles"); return; }
    var route = $("chat-route").value;
    var ref = routeToModel(prof, route, instruction);
    if (!ref || !ref.connection || !ref.model) {
      setStatus("Profile '" + prof.name + "' has no model assigned for this route. Open Profiles and pick connection + model.", "error");
      return;
    }
    SK.busy = true;
    setRunEnabled(false);
    var wantCtx = $("chat-context").checked;
    setStatus("Collecting deck context…", "busy");
    var ctxPromise = wantCtx ? SK.deck.getContext({}).catch(function (err) {
      setStatus("Could not read the deck (" + err.message + ") - continuing without context.", "error");
      return null;
    }) : Promise.resolve(null);

    ctxPromise.then(function (ctx) {
      SK.lastCtx = ctx;
      var sys = SYSTEM_PROMPT + "\nHost: " + (SK.sim ? "SlideKick simulator" : SK.deck.capabilities().host);
      if (!(SK.cfg && SK.cfg.generateNotes)) sys += "\n" + NOTES_OFF_RULE;
      var user = (ctx ? formatContext(ctx) : "(no deck context requested)") + "\n\nUSER REQUEST:\n" + instruction;
      var messages = [
        { role: "system", content: sys },
      ].concat(SK.history.slice(-(HISTORY_MAX_EXCHANGES * 2)))
       .concat([{ role: "user", content: user }]);
      dbg("prompt", { profile: prof.name, route: route, model: ref.model, connection: ref.connection, instruction: instruction, hadContext: !!ctx, messages: messages });
      setStatus("Asking " + ref.model + " (via " + ref.connection + ")…", "busy");
      return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: messages }).then(function (resp) {
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
          return apiJson("/api/chat", "POST", { connection: ref.connection, model: ref.model, messages: repair }).then(function (resp2) {
            dbg("repair.response", { text: resp2.text, elapsedMs: resp2.elapsedMs });
            SK.lastModelText = { text: resp2.text };
            return tryParse(resp2.text);
          });
        }
        return parsed;
      }).then(function (parsed) {
        renderParsed(parsed, ctx, instruction);
        pushHistory(instruction, (SK.lastModelText && SK.lastModelText.text) || parsed.text || "");
      });
    }).catch(function (err) {
      var msg = err && err.friendly ? err.message : (err && err.message) || String(err);
      setStatus(msg, "error");
    }).then(function () {
      SK.busy = false;
      setRunEnabled(true);
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

  function renderParsed(parsed, ctx, instruction) {
    var out = $("chat-output");
    out.innerHTML = "";
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
      mb.appendChild(document.createElement("h4")).textContent = "Notes could not be written automatically - copy manually:";
      res.manualNotes.forEach(function (mn) {
        var ta = document.createElement("textarea");
        ta.value = "Slide " + (mn.slide + 1) + ": " + mn.text;
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
    h.textContent = "Proposed actions" + ($("chat-autoapply").checked ? " (auto-apply)" : "") + " - approval policy: " +
      OP_TYPES.filter(function (t) { return approvalMode(t) !== "auto"; }).map(function (t) { return OPS_META[t].label + "=" + approvalMode(t); }).join(", ");
    box.appendChild(h);

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

    var tableRows = document.createElement("div");
    rows.forEach(function (row) {
      var r = document.createElement("div");
      r.className = "oprow" + (row.mode === "deny" ? " blocked" : row.mode === "invalid" ? " invalid" : "");
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
            ta.value = "Slide " + (mn.slide + 1) + ": " + mn.text;
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
    setStatus(confirmOps ? "Approval needed before applying." : "Review proposed actions.", "");
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
  function switchTab(id) {
    SK.activeTab = id;
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
        "<td>" + (c.hasKey ? "set" : "—") + "</td></tr>"
      );
    });
    list.innerHTML = rows.join("") || "<tr><td>No connections yet - add one below.</td></tr>";
    Array.prototype.forEach.call(list.querySelectorAll("[data-edit]"), function (b) {
      b.addEventListener("click", function () { editConnection(b.dataset.edit); });
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
    ["prof-content-conn", "prof-other-conn"].forEach(function (id) {
      var sel = $(id);
      fillConnSelect(sel, sel.value || "");
    });
    var list = $("prof-list");
    var rows = [ "<tr><th>Name</th><th>Content model</th><th>Other model</th></tr>" ];
    (SK.cfg.profiles || []).forEach(function (p) {
      var star = p.name === SK.cfg.defaultProfile ? "*" : "";
      // click the name to edit; deletion lives in the edit form (with confirm)
      var nameCell =
        '<button class="namebtn" data-edit="' + esc(p.name) + '" title="edit">' +
        esc(p.name) + "</button>" + (star ? " <span class=\"opnote\">(default)</span>" : "");
      rows.push(
        "<tr><td>" + nameCell + "</td>" +
        "<td>" + esc(modelRefLabel(p.content)) + "</td><td>" + esc(modelRefLabel(p.other)) + "</td></tr>"
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

  function modelsCacheGet(connName, force) {
    var cached = SK.modelsCache[connName];
    if (cached && !force && Date.now() - cached.at < 5 * 60 * 1000 && cached.list.length) {
      return Promise.resolve(cached.list);
    }
    return apiJson("/api/models?connection=" + encodeURIComponent(connName), "GET").then(function (res) {
      var list = (res.models || []).map(String);
      SK.modelsCache[connName] = { list: list, at: Date.now(), source: res.source };
      return list;
    });
  }

  function populateModelSelect(slot, connName, prefer) {
    var sel = $(slot + "-conn") && $(slot + "-model");
    var modelSel = $(slot + "-model");
    var custom = $(slot + "-custom");
    modelSel.innerHTML = '<option value="">loading…</option>';
    custom.classList.add("hidden");
    if (!connName) {
      modelSel.innerHTML = '<option value="">(choose a connection first)</option>';
      return Promise.resolve();
    }
    return modelsCacheGet(connName).then(function (list) {
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
    setFormMsg($("prof-formmsg"), "", true);
  }

  function resetProfForm() {
    SK.editingProf = null;
    $("prof-form-title").textContent = "New profile";
    $("prof-name").value = "";
    fillConnSelect($("prof-content-conn"), "");
    fillConnSelect($("prof-other-conn"), "");
    $("prof-content-model").innerHTML = '<option value="">(choose a connection first)</option>';
    $("prof-other-model").innerHTML = '<option value="">(choose a connection first)</option>';
    $("prof-content-custom").classList.add("hidden");
    $("prof-other-custom").classList.add("hidden");
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
    return {
      name: $("prof-name").value.trim(),
      content: slotValue("prof-content"),
      other: slotValue("prof-other"),
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
      $("statusbar").textContent = "SlideKick v" + (cfg.version || "0.1.0") +
        " - server 127.0.0.1:" + cfg.port + " - default profile: " +
        (cfg.defaultProfile || "(none)") + " - debug " + (cfg.debug ? "ON" : "off") +
        " - connections: " + (cfg.connections || []).length;
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
      setStatus("Conversation history cleared.", "ok");
    });
    $("chat-refresh-deck").addEventListener("click", function () { updateDeckSummary(); setStatus("Deck refreshed.", ""); });
    Array.prototype.forEach.call(document.querySelectorAll("#chat-chips .chip"), function (chip) {
      chip.addEventListener("click", function () {
        $("chat-input").value = chip.dataset.prompt;
        $("chat-input").focus();
      });
    });

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
        delete SK.modelsCache[connName];
        populateModelSelect(slot, connName, $(slot + "-model").value);
      });
    });
    ["content", "other"].forEach(function (slot) {
      $("prof-" + slot + "-conn").addEventListener("change", function () {
        populateModelSelect("prof-" + slot, $("prof-" + slot + "-conn").value, "");
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
    $("settings-debug-save").addEventListener("click", saveDebugOption);
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

  function finishBoot() {
    wireEvents();
    bootLoad(0);
  }

  function bootSuccess(cfg) {
    updateDeckSummary();
    renderSimDeck();
    if (!(cfg.connections || []).length) switchTab("connections");
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
    setStatus("Cannot reach the SlideKick server: " + err.message + " Is it running? (install.sh / install.ps1 / LaunchAgent)", "error");
    $("statusbar").textContent = "server unreachable: " + err.message;
  }

  function boot() {
    var params = new URLSearchParams(String(window.location.search || ""));
    SK.sim = window.__SK_SIM__ === true || params.has("sim");
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