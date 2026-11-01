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
    app: "slidekick",
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

  var LAYOUT_KINDS = ["title_content", "title_only", "section", "blank"];

  var OPS_META = {
    create_slide: { label: "create slide", hint: "Appends a new slide (Office.js can only append; drag into place if needed)." },
    edit_slide: { label: "edit slide", hint: "Replaces title / body / speaker notes of an existing slide." },
    replace_text: { label: "replace text", hint: "Rewrites the text of one shape in place; with a find fragment, only that fragment is replaced (rest of the list untouched)." },
    set_notes: { label: "set notes", hint: "Writes speaker notes on a slide." },
    create_table: { label: "create table", hint: "Adds a table (new slide when slide is empty; position/size in points)." },
    move_shape: { label: "move/resize shape", hint: "Sets left/top/width/height (points) and optional rotation (degrees) of one shape - design passes use this." },
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
          if (o2.kind === "reply" || o2.kind === "read_slide" || o2.kind === "web_lookup") return o2;
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
    '- {"type":"create_slide","layout":"title_content|title_only|section|blank","title":"...","body":["bullet",...],"notes":"..."} (body/notes optional; new slides are appended at the end of the deck)',
    '- {"type":"edit_slide","slide":N,"title":"..."|null,"body":[...]|null,"notes":"..."|null} (null means: leave unchanged)',
    '- {"type":"replace_text","slide":N,"shape":N,"text":"..."} (reword one shape; N refers to the shape numbers listed for that slide). PREFERRED for a bullet/list item edit: add "find":"<the exact current text of that one bullet/fragment from the outline>" and make "text" the REPLACEMENT for only that fragment - the rest of the shape\'s text is preserved. Without "find" the op replaces the shape\'s WHOLE text, so only do that when the outline shows the entire text (no …[text continues]) or you really want to rewrite it all',
    '- {"type":"set_notes","slide":N,"notes":"..."}',
    '- {"type":"create_table","slide":N|null,"values":[["h1","h2"],["x","y"]],"left":40,"top":150,"width":640} (all cells strings, rows rectangular; slide null = new title slide first; left/top/width optional points)',
    '- {"type":"move_shape","slide":N,"shape":N,"left":40,"top":150,"width":300,"height":60} (points; only the fields to change; keep shapes inside the slide bounds shown in the outline)',
    '- {"type":"create_image","prompt":"detailed image description","slide":N|null,"width":400} (generate an image from the description and insert it; slide null = the slide the user currently has open; only when the user asks for a picture/image/illustration/logo)',
    '- {"type":"apply_layout","slide":N,"layout":"title_content|title_only|section|blank"} (only when the user asks for layout/design changes)',
    '- {"type":"format_text","slide":N,"shape":N,"font":{"bold":true},"align":"center","bullet":true} (format ONE shape from the outline: font = {"bold","italic","color","size","name","underline","allCaps"} subset only, colors as HTML hex/red/green/...; align = left|center|right; bullet = true adds bullets; only requested properties)',
    '- Office.js cannot CREATE SmartArt, stock icons or 3D models. If the user asks for those, say so and offer shape/textbox-based alternatives (create_shape) or an AI-generated picture (create_image) instead. The outline marks existing SmartArt/model3D shapes so you can reference them.',
    '- {"type":"create_shape","slide":N,"kind":"textbox|rectangle|ellipse|roundedRectangle|hexagon|diamond|triangle|line","text":"only for textbox","left":120,"top":200,"width":240,"height":80} (points, slide-local; line = start point left/top + width/height = endpoint offsets; other shapes: no text; optional "fill":"#RRGGBB", "line":"#RRGGBB", "transparency":0.25 to style the created shape immediately)',
    '- to format only PART of a shape\'s text, add "find":"exact words appearing in that shape\'s text" (case-insensitive substring; e.g. {"type":"format_text","slide":N,"shape":N,"find":"security","font":{"bold":true}})',,
    '- {"type":"format_selection","font":{"bold":true}} (format ONLY the text the user has selected right now; same font subset; no slide/shape fields)',
    '- {"type":"set_background","slide":N,"kind":"image","prompt":"description","transparency":0.2} (an AI-generated picture as the slide background; transparency 0-1 optional) or {"type":"set_background","slide":N,"kind":"solid","color":"#RRGGBB","transparency":0.2} (solid color; transparency optional)',
    '- {"type":"add_hyperlink","slide":N,"shape":N,"url":"https://...","tip":"tooltip text"} (link the shape to the URL; url must be http(s) or mailto:)',
    '- {"type":"style_shape","slide":N,"shape":N,"fill":"#RRGGBB","line":"#RRGGBB","transparency":0.3} (solid fill and/or outline color and optional transparency 0-1 for ONE shape; colors as HTML hex or named; at least one of fill/line/transparency)',
    '- {"type":"set_alt_text","slide":N,"shape":N,"description":"alt text","title":"optional title"} (accessibility alt text for the shape)',
    '- {"type":"set_shape_image","slide":N,"shape":N,"prompt":"detailed image description","transparency":0.3} (picture-fill ONE existing shape (panel/circle/rectangle) with an AI-generated picture; transparency optional; use for cut-out photos, side panels, badges)',
    '- {"type":"move_slide","slide":N,"to":M} (move slide N so it becomes slide M; numbers as in the outline 1..N; run it ALONE in its own reply - do not combine with other ops. PowerPoint cannot place a brand-new slide at a position: create it first (it lands last) and move_slide it in a following step)',
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
    "- Make-over building blocks: set_background (AI picture or solid color, optional transparency); create_shape with kind + geometry + fill/line/transparency for panels, accent bars, color blocks and overlays; style_shape with fill/line/transparency to restyle anything already on the slide; move_shape to reposition/resize (+rotation, degrees) or rotate anything you created or moved (rotation is how diagonal/tilted compositions are made); replace_text/format_text for the words; set_shape_image to fill an EXISTING shape (panel/circle) with an AI-generated picture. Shape kinds: textbox, rectangle, ellipse, roundedRectangle, hexagon, diamond, triangle, line.",
    "- Never remove or overwrite the user's own content: pictures, charts and tables stay. Redesign AROUND them (move, restyle, wrap in panels, put overlays above). clear_slide only on nearly-empty placeholder slides; duplicate_slide_to replaces its target completely - never aim it at a slide that still carries user content.",
    "- READABILITY LAW for every design or make-over pass: choose the TEXT color from the field it sits on - dark text (near-black) on light/plain/panel fields, light text (near-white) ONLY on deep backgrounds, dark overlays or panel colors with strong contrast; body text >= 16pt on colored fields (>= 14pt anywhere, never below 12pt), titles >= 28pt with an obvious title/body size gap; never place body text directly on busy photos - add a dark overlay or leave an empty field; avoid mid-tone text on mid-tone fills - pick the opposite extreme or set the text on a small solid panel.",
    "- A redesign pass must cover EVERY in-scope slide with its structural change in the same flow. If the op list would be very long, reply phase 1 and immediately continue with phase 2 as the next action JSON - never leave half the slides untouched.",
    "- Text over pictures: order the ops so the picture comes FIRST (set_background or a picture-filled shape), then the overlay (create_shape rectangle, fill #000000, transparency between 0.45 and 0.6), then the text (white or light via replace_text/format_text). Ops apply in the order you return them.",
    "- Make-over phase order per reply: (1) backgrounds + palette fields + overlay shapes, (2) layout moves + text restyle, (3) image generation requests. This keeps each approval batch reviewable and each design phase consistent.",
    "- Image requests: always return a create_image op (never generate the picture inline); SlideKick fetches it with the user's configured image model. Include a concrete, self-contained prompt in the op.",
    "- Image refinement ('refine/change the image'): write a FRESH self-contained image prompt describing the FULL desired result. Never embed previous prompts, the conversation, or op JSON into it - only the scene essentials plus the requested change. Prefer 200-600 words.",
    "- READ SLIDES ON DEMAND: the outline shows full shape text only for small decks and for the selected slide(s) - in LARGE decks every other slide shows just its title and a short preview. When you need a slide's ACTUAL content (to answer a question about it, or before rewriting it), reply with only {\"kind\":\"read_slide\",\"slides\":[N]} (1-based outline numbers, max 12). SlideKick re-asks you with the COMPLETE text of those slides. Never guess content from a title.",
    "- LIVE WEB LOOKUPS: when the request needs up-to-date info you cannot reliably know (current events, news, prices, recent releases, rankings, live data), reply with ONLY {\"kind\":\"web_lookup\"} - nothing else. SlideKick then runs a web-enabled lookup and re-asks you with fresh citable results. Never use it for deck/document/mail edits, styling, general concepts, or what the outline/context already answers.",

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
    "no reasoning, no fences, no commentary. Never reply with nothing at all - " +
    "an empty response is a failure; if you cannot act, use format A with the " +
    "reason as text.";

  var CONTINUE_ASK =
    "Your previous reply was cut off mid-way through the op list. Reply with a " +
    "slide_ops JSON containing ONLY the ops still needed to FINISH the redesign " +
    "(slides not yet covered), same JSON format. If nothing is missing, reply " +
    '{"kind":"reply","text":"done"}';

  var REPAIR_EMPTY_ASK =
    "Your previous reply was EMPTY. An empty response is NEVER acceptable. " +
    "Always produce something useful: the slide_ops JSON with actions for the " +
    "user's request, or format A (reply) with the answer/reason. Respond NOW " +
    "with the JSON object - no fences, no commentary, never an empty reply.";

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
          try { sessionStorage.setItem("sk-resume-session", SK.sessionId); } catch (eS) {}
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
      try { sessionStorage.setItem("sk-resume-session", doc.id); } catch (eS) {}
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
      if (SK.sessionId === sessId) { SK.sessionId = null; SK.sessionName = "";
        try { sessionStorage.removeItem("sk-resume-session"); } catch (eS) {} }
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
    "\\b(slide|slides|deck|bullet|bullets|reword|rewrite|rephrase|outline|notes?|title|titles|add|create|make|generate|tighten|restructure|expand|insert|append|design|align|balance|spacing|table|tables|chart|bold|italic|underline|color|font|highlight|format|formatting|center|caps|image|images|picture|pictures|photo|logo|illustration)\\b", "i"
  );

  function assistantHistoryText(parsed) {
    // a compact recap for the next round; never the raw ops JSON (it bloats every
    // later request and pushes the models toward echoing blobs into new prompts)
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
    SK.sessionId = null;      // the next exchange saves as a NEW session
    try { sessionStorage.removeItem("sk-resume-session"); } catch (eS) {}
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
      var limit = Math.min(opts.maxSlides || 200, 400);
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
      if (geometry.rotation != null) {
        if (!Office.context.requirements.isSetSupported("PowerPointApi", 1.10))
          throw new FriendlyError("rotating shapes needs PowerPointApi 1.10 - update PowerPoint (16.105+ on Mac).");
        shape.rotation = geometry.rotation;
      }
      return ctx.sync().then(function () { return { ok: true }; });
    });
  };

OfficeDeck.prototype.addImage = function (payload) {
  return new Promise(function (resolve, reject) {
    try {
      var opts = { coercionType: Office.CoercionType.Image };
      var geo = payload.geometry || {};
      if (geo.left != null) opts.imageLeft = geo.left;
      if (geo.top != null) opts.imageTop = geo.top;
      if (geo.width != null) opts.imageWidth = geo.width;
      if (geo.height != null) opts.imageHeight = geo.height;
      Office.context.document.setSelectedDataAsync(String(payload.b64), opts, function (ar) {
        if (ar && ar.status === Office.AsyncResultStatus.Succeeded) resolve({ ok: true });
        else reject(new FriendlyError("image insert failed: " + ((ar && ar.error && ar.error.message) || (ar && ar.status) || "unknown")));
      });
    } catch (e) {
      reject(new FriendlyError("image insert failed: " + (e && e.message || e)));
    }
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
        return slides.exportAsBase64Presentation([slides.getItemAt(from)]).then(function (b64) {
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
        return slides.exportAsBase64Presentation([slides.getItemAt(i)]).then(function (b64) {
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
        return slides.exportAsBase64Presentation([slides.getItemAt(from)]).then(function (b64) {
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
    return this.run("clearSlide", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
        throw new FriendlyError("clearing a slide needs PowerPointApi 1.5 - update PowerPoint (16.63+ on Mac).");
      }
      var shapes = ctx.presentation.slides.getItemAt(i).shapes;
      shapes.load("items/type,items/hasTextFrame,items/textFrame,items/name");
      return ctx.sync().then(function () {
        var cleared = 0, removed = 0;
        shapes.items.forEach(function (shape) {
          if (shape.hasTextFrame && shape.textFrame) {
            shape.textFrame.textRange.text = "";
            cleared++;
          } else if (shape.type === "Line" && String(shape.name || "").indexOf("sk-shape") === 0) {
            // only drop lines this add-in created - user lines stay
            shape.delete();
            removed++;
          }
        });
        return ctx.sync().then(function () { return { ok: true, cleared: cleared, removed: removed }; });
      });
    });
  };

OfficeDeck.prototype.createShape = function (slideIndex, op) {
    return this.run("createShape", function (ctx) {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.4")) {
        throw new FriendlyError("shape creation needs PowerPointApi 1.4 - update PowerPoint (16.62+ on Mac).");
      }
      var shapes = ctx.presentation.slides.getItemAt(slideIndex).shapes;
      var opts = {};
      ["left", "top", "width", "height"].forEach(function (gk) {
        if (op.geometry && op.geometry[gk] != null) opts[gk] = op.geometry[gk];
      });
      var created;
      if (op.kind === "textbox") {
        created = shapes.addTextBox(String(op.text || ""), opts);
      } else if (op.kind === "line") {
        created = shapes.addLine("Straight", opts);
      } else {
        created = shapes.addGeometricShape(op.kind, opts);
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
      return ctx.sync().then(function () { return { ok: true }; });
    });
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
      var shape = ctx.presentation.slides.getItemAt(slideIndex).shapes.getItemAt(shapeIndex);
      var range = shape.textFrame.textRange;
      range.load("text");
      return ctx.sync().then(function () {
        var idx = String(range.text || "").toLowerCase().indexOf(String(findText).toLowerCase());
        if (idx < 0) throw new FriendlyError("\"" + findText + "\" not found in that shape\'s current text.");
        var sub = range.getSubstring(idx, String(findText).length);
        if ((findText, font.bold) != null) sub.font.bold = font.bold;
        if (font.italic != null) sub.font.italic = font.italic;
        if (font.allCaps != null) sub.font.allCaps = font.allCaps;
        if (font.underline != null) sub.font.underline = font.underline;
        if (font.color) sub.font.color = font.color;
        if (font.name) sub.font.name = font.name;
        if (font.size) sub.font.size = font.size;
        return ctx.sync().then(function () { return { ok: true }; });
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

OfficeDeck.prototype.formatText =OfficeDeck.prototype.formatText = function (slideIndex, shapeIndex, font) {
    return this.run("formatText", function (ctx) {
      var shape = ctx.presentation.slides.getItemAt(slideIndex).shapes.getItemAt(shapeIndex);
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
              if (total.length > 60000) total = total.slice(0, 60000) + "\n…[truncated]";
              return total;
            });
          });
        });
      });
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
      var slideOk = function (v) { return typeof v === "number" && Math.floor(v) === v && v >= 1 && (!ctx || v <= ctx.slideCount); };
      var shapeOk = function (sl, sh) {
        if (!ctx || !ctx.slides[sl - 1]) return false;
        return (ctx.slides[sl - 1].shapes || []).some(function (s) { return s.idx === sh - 1 && s.text != null; });
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
        if (typeof op.text !== "string" || !op.text.length) { bad("missing text"); return; }
        out.slide = op.slide; out.shape = op.shape; out.text = op.text;
        if (typeof op.find === "string" && op.find.trim()) {
          var findR = op.find.trim();
          if (findR.length > 200) { bad("find text too long (200 chars max)"); return; }
          var shapeTxtR = ((ctx.slides[op.slide - 1].shapes || []).find(function (sh) { return sh.idx === op.shape - 1; }) || {}).text || "";
          if (shapeTxtR.toLowerCase().indexOf(findR.toLowerCase()) === -1) {
            bad("find text not found in that shape's current text"); return;
          }
          out.find = findR;
        }
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
        var shapeInCtx = ctx && ctx.slides[op.slide - 1] &&
          (ctx.slides[op.slide - 1].shapes || []).some(function (s) { return s.idx === op.shape - 1; });
        if (!shapeInCtx) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
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
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        var shapeInCtxF = ctx && ctx.slides[op.slide - 1] &&
          (ctx.slides[op.slide - 1].shapes || []).some(function (sh) { return sh.idx === op.shape - 1; });
        if (!shapeInCtxF) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
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
          var shapeTxtF = (ctx.slides[op.slide - 1].shapes.find(function (sh) { return sh.idx === op.shape - 1; }) || {}).text || "";
          if (shapeTxtF.toLowerCase().indexOf(findS.toLowerCase()) === -1) {
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
        if (!shapeInCtxS) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
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
        if (!shapeInCtxI) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        if (typeof op.prompt !== "string" || op.prompt.trim().length < 3) { bad("missing image prompt"); return; }
        if (op.prompt.length > 4000) { bad("image prompt too long (max 4000 chars)"); return; }
        out.slide = op.slide; out.shape = op.shape; out.prompt = op.prompt.trim();
        if (typeof op.transparency === "number" && op.transparency >= 0 && op.transparency <= 1) out.transparency = op.transparency;
      } else if (t === "add_hyperlink") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        var shapeInCtxH = ctx && ctx.slides[op.slide - 1] &&
          (ctx.slides[op.slide - 1].shapes || []).some(function (sh) { return sh.idx === op.shape - 1; });
        if (!shapeInCtxH) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        if (typeof op.url !== "string" || !/^(https?:\/\/\S+|mailto:\S+)$/.test(op.url.trim())) { bad("url must be http(s)://... or mailto:..."); return; }
        out.slide = op.slide; out.shape = op.shape; out.url = op.url.trim();
        if (typeof op.tip === "string" && op.tip.trim()) out.tip = op.tip.trim().slice(0, 200);
      } else if (t === "set_alt_text") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        var shapeInCtxA = ctx && ctx.slides[op.slide - 1] &&
          (ctx.slides[op.slide - 1].shapes || []).some(function (sh) { return sh.idx === op.shape - 1; });
        if (!shapeInCtxA) { bad("shape " + op.shape + " not in the context list for slide " + op.slide); return; }
        if (typeof op.description !== "string" || !op.description.trim()) { bad("alt text needs a description"); return; }
        out.slide = op.slide; out.shape = op.shape; out.description = op.description.trim().slice(0, 500);
        if (typeof op.title === "string" && op.title.trim()) out.title = op.title.trim().slice(0, 200);
      } else if (t === "create_shape") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        var KINDS = { textbox: 1, rectangle: 1, ellipse: 1, roundedRectangle: 1, hexagon: 1, diamond: 1, triangle: 1, line: 1 };
        var kindS = String(op.kind || "");
        kindS = Object.keys(KINDS).find(function (kk) { return kk.toLowerCase() === kindS.toLowerCase(); });
        if (!kindS) { bad("kind must be textbox, rectangle, ellipse, roundedRectangle, hexagon, diamond, triangle or line"); return; }
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
      } else if (t === "delete_slide") {
        if (!slideOk(op.slide)) { bad("bad slide index"); return; }
        out.slide = op.slide;
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
    return { ops: ops, invalid: invalid, notesDropped: notesDropped };
  }

  function approvalMode(opType) {
    return SK.sessionApprovals[opType] || (SK.cfg && SK.cfg.approvals && SK.cfg.approvals[opType]) || "confirm";
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
            var note = "";
            if (!SK.sim && ctx && op.after != null && op.after <= ctx.slideCount - 1) {
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
        case "set_notes":
          return deck.setNotes(cur, op.notes).then(function (nr) {
            if (nr && nr.manual) manualNotes.push({ slide: cur, text: op.notes, reason: nr.reason });
            else lines.push({ ok: true, text: "set notes on slide " + op.slide });
          });
        case "create_shape":
          return deck.createShape(cur, op).then(function () {
            lines.push({ ok: true, text: fmtOp(op, ctx) });
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
            return Promise.race([
              apiJson("/api/image", "POST", {
                connection: imgRefB.connection, model: imgRefB.model, prompt: op.prompt,
              }),
              new Promise(function (_, rej) {
                setTimeout(function () {
                  rej(new FriendlyError("image generation timed out after 300 s (no answer from '" + imgRefB.model + "')"));
                }, 300000);
              }),
            ]).then(function (img) {
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
          return Promise.race([
            apiJson("/api/image", "POST", {
              connection: imgRefI.connection, model: imgRefI.model, prompt: op.prompt,
            }),
            new Promise(function (_, rej) {
              setTimeout(function () {
                rej(new FriendlyError("image generation timed out after 300 s (no answer from '" + imgRefI.model + "')"));
              }, 300000);
            }),
          ]).then(function (img) {
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
            return deck.addImage({ b64: img.b64, slide: cur, geometry: op.geometry }).then(function (ri) {
              var note = !SK.sim && op.slide != null && ctx && op.slide - 1 !== ctx.activeSlide
                ? " (targeting the slide you have open in PowerPoint; requested slide " + op.slide + ")" : "";
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
      seq = seq.then(function () { return runOne(op); }).catch(function (err) {
        var msg = err && err.friendly ? err.message : (err && err.message) || String(err);
        lines.push({ ok: false, text: op.type + ": " + msg });
        failedOps.push(op);
        if (!(err && err.friendly)) {
          lines.push({ ok: false, text: "unexpected error - stopping remaining ops" });
          abortedIdx = ix;
          throw err; // aborts the chain below; remaining ops join the retry pass
        }
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
        rseq = rseq.then(function () { return runOne(op); }).catch(function (err2) {
          var msg2 = err2 && err2.friendly ? err2.message : (err2 && err2.message) || String(err2);
          lines.push({ ok: false, text: "retry of " + op.type + " failed again: " + msg2 });
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
    lines.push("DECK OUTLINE: " + ctx.slideCount + " slides" + sel + (ctx.truncated ? " (only the first slides are listed)" : "") + ", host: " + ctx.host);
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
        if (t2.length > 90) t2 = t2.slice(0, 90) + "…";
        return t2 ? " — " + t2 : "";
      };
      ctx.slides.forEach(function (s2) {
        var role2 = slideRoleOf(s2);
        var star2 = ctx.selected && ctx.selected.indexOf(s2.idx) >= 0 ? "*" : "";
        var tt2 = (s2.shapes || []).filter(function (sh) { return sh.isTitle || (sh.text && /title/i.test(sh.name)); })[0];
        var line2 = "[" + (s2.idx + 1) + star2 + "] (role: " + role2 + ") " +
          (tt2 && tt2.text ? tt2.text : "(no title)") + bodyPrevF(s2);
        if (line2.length > 170) line2 = line2.slice(0, 170);
        lines.push(line2);
        if (selMap[s2.idx]) {
          var selTxt = 0;
          (s2.shapes || []).forEach(function (sh) {
            if (sh.text == null || !String(sh.text).trim()) return;
            var role = sh.isTitle ? "title" : sh.isBody ? "body" : "text";
            var txt = String(sh.text).replace(/\r|\n/g, " | ");
            if (txt.length > 500) txt = txt.slice(0, 500) + "…[text continues]";
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
        var txt = String(sh.text).replace(/\r|\n/g, " | ");
        if (txt.length > 400) txt = txt.slice(0, 400) + "…[text continues]";
        if (!txt.trim()) return;
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
      SK.webResults = "";
      var webPromise = Promise.resolve("");
      return webPromise.then(function () {
      var sys = SYSTEM_PROMPT + "\nHost: " + (SK.sim ? "SlideKick simulator" : SK.deck.capabilities().host);
      if (!(SK.cfg && SK.cfg.generateNotes)) sys += "\n" + NOTES_OFF_RULE;
      if (MAKEOVER_RE.test(instruction)) sys += "\n" + MAKEOVER_BRIEF;
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
            if (pre.ok && pre.kind === "read_slide") return readSlideThen(resp);
            return resp;
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
            "one image-style phrase used by every picture; and one DISTINCT " +
            "composition idea per slide role. Take risks - the opposite of a " +
            "default PowerPoint look. If a earlier redesign appears in the " +
            "history, pick a contrasting direction. 150-300 words.",
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
            var bm = messages.concat([
              { role: "assistant", content: moConcept || "" },
              { role: "user",
                content: "CONCEPT (from the design director):\n" + (moConcept || "(decide the concept yourself)") +
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
          function wholePass(passName, spec) {
            return passAsk(passName, spec, null, null).then(function (pk) { return collect(pk, null, null); });
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
            return wholePass("ART PASS",
            "ART PASS - imagery only where the concept calls for it: AI picture backgrounds for cover and section slides, cut-out picture fills inside shapes for content slides, always in the concept's one image style. Use set_background kind:image and set_shape_image. Quality over quantity.").then(function (more) {
              all = all.concat(more);
              return all;
            });
          });
          // PASS 4 - REVIEW LOOP: the model audits its own redesign and keeps
          // fixing until it declares the deck done (max 2 fix rounds)
          var p4 = p3.then(function (acc3) {
            all = acc3;
            function reviewRound(n) {
              var rm = messages.concat([
                { role: "assistant", content: moConcept || "" },
                { role: "user",
                  content: "REVIEW ROUND for the make-over. CONCEPT:\n" + (moConcept || "(your concept)") +
                    "\n\nOPS SO FAR: " + all.length + " action(s) across the deck. " + coverageNote() +
                    "\n\nAudit the redesign against the concept AND the readability law. If specific slides still lack structural design, imagery the concept promised, or readable text, reply with slide_ops fixing ONLY those slides. If the deck is complete per the concept, reply {\"kind\":\"reply\",\"text\":\"done\"}." },
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
        if (parsed.ok && parsed.recovered) {
          // truncated reply: complete ops salvaged - ask once for the rest, merge
          SK.lastRecovered = parsed.ops.length;
          setStatus("Reply was cut off - recovered " + parsed.ops.length + " action(s), asking for the rest…", "busy");
          var cont = messages.concat([
            { role: "assistant", content: resp.text || "" },
            { role: "user", content: CONTINUE_ASK },
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
              { role: "user", content: CONTINUE_ASK },
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
      if (obj.kind === "reply") {
        if (typeof obj.text === "string") return { ok: true, kind: "reply", text: obj.text };
        return { ok: true, kind: "reply", text: String(obj.text == null ? "" : JSON.stringify(obj.text)) };
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
    cpIc.textContent = "\u29C9";               // copy icon next to the step title
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
    wirePastHeader();
    // the request prints above the response/summary (echoed by the clear)
    var reqEcho = document.createElement("div");
    reqEcho.className = "reqecho";
    reqEcho.textContent = instruction;
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
      recN.textContent = "Recovered " + SK.lastRecovered + " action(s) from a cut-off reply; the model was asked for the rest. If some slides are still untouched, ask to continue.";
      box.appendChild(recN);
      SK.lastRecovered = 0;
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
      // unapproved confirm rows would silently vanish - ask before skipping
      var skips = rows.filter(function (row) {
        return !row.removed && row.mode === "confirm" && row._cb && !row._cb.checked;
      });
      var counts = {};
      skips.forEach(function (row) { counts[OPS_META[row.op.type].label] = (counts[OPS_META[row.op.type].label] || 0) + 1; });
      var skipList = Object.keys(counts).map(function (l) { return counts[l] + " " + l; }).join(", ");
      if (skipList && chosen.length) {
        var okApply = window.confirm(
          skips.length + " (needs approval) action(s) are NOT approved and will be skipped: " +
          skipList + ". Apply the other " + chosen.length + " action(s) anyway?");
        if (!okApply) return;
      }
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
    populateModelSelect("prof-content", p.content && p.content.connection, p.content && p.content.model);
    populateModelSelect("prof-other", p.other && p.other.connection, p.other && p.other.model);
    SK.editingProfImage = p.image || null;
    fillConnSelect($("prof-image-conn"), p.image && p.image.connection);
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
    SK.editingProfImage = null;
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
      $("statusbar").textContent = "SlideKick v" + (cfg.version || "0.1.0") +
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
    $("settings-notes-save").addEventListener("click", saveNotesOnly);
    $("settings-autoclean-save").addEventListener("click", saveAutoClean);
    var updBtn = $("upd-check");
    if (updBtn) updBtn.addEventListener("click", updCheck);
    var relBtn = $("pane-reload");
    if (relBtn) relBtn.addEventListener("click", function () {
      dbg("pane.reload", {});
      // hard reload of the pane from the running server, busting any cache
      var qs = String(location.search || "").replace(/cb=\d+(&|$)/g, "").replace(/^\?/, "");
      location.replace(location.pathname + "?" + (qs ? qs + "&" : "") + "cb=" + Date.now());
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
        if (updAuto.checked) updCheck();   // silent availability check at boot
      }, function () { updAuto.checked = true; updCheck(); });
      updAuto.addEventListener("change", function () {
        apiJson("/api/settings", "PUT", { autoUpdateCheck: !!updAuto.checked });
      });
    }
    // a pane reload (or back-end restart) resumes the session that was active
    var bootResume = null;
    try { bootResume = sessionStorage.getItem("sk-resume-session"); } catch (eS) {}
    if (bootResume) resumeSession(bootResume);
    updateSessionLabel();
    wireSessions();
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
    };
    xr.onerror = xr.ontimeout = function () {
      el.className = "beat bad";
      el.title = "Local server NOT reachable - last check " + stamp
        + ". Start it with: python3 install.py";
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