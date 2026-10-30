// DocKick simulator: in-memory Word document (paragraphs + tables + comments)
// used to exercise the Word task pane in a regular browser. Same interface as
// window.SimDeck (SlideKick) so the pane boot code does not change.
(function () {
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]; }); }
  "use strict";

  var images = [];
  var initial = function () {
    return [
      { text: "Threat hunting with AI", style: "Heading 1" },
      { text: "Hunters collect signals across the environment and reason about them before an intrusion detonates.", style: "Normal" },
      { text: "This document is a working draft - it demonstrates DocKick's simulator.", style: "Normal" },
    ];
  };

  var paras = [];
  var tables = [];
  var comments = [];
  var templateBlocks = [];
  var trackChanges = false;

  function reset() {
    paras = initial();
    tables = [];
    comments = [];
    images = [];
    templateBlocks = [];
    trackChanges = false;
  }
  reset();

  class FriendlyError extends Error {}

  function styleName(kind) {
    if (!kind) return "Normal";
    var k = String(kind).toLowerCase();
    if (k.indexOf("title") === 0) return "Title";
    if (k === "heading1" || k === "heading 1") return "Heading 1";
    if (k === "heading2" || k === "heading 2") return "Heading 2";
    if (k === "heading3" || k === "heading 3") return "Heading 3";
    if (k.indexOf("quote") === 0) return "Quote";
    if (k === "list" || k.indexOf("bullet") === 0) return "List Paragraph";
    return "Normal";
  }

  var SimAPI = {
    sim: true,
    word: true,
    capabilities: function () {
      return { text: true, comments: true, tables: true, selection: true, host: "DocKick simulator" };
    },
    getContext: async function () {
      return {
        documentName: "DocKick simulator document",
        paragraphCount: paras.length,
        paragraphs: paras.map(function (p, i) { return { idx: i, text: p.text, style: p.style }; }),
        tables: tables.map(function (t, i) { return { idx: i, values: t.values }; }),
        comments: comments.slice(),
        selected: "",
        host: "DocKick simulator",
      };
    },
    insertParagraphs: async function (op) {
      (op.items || []).forEach(function (it) {
        paras.push({ text: String(it.text == null ? "" : it.text), style: styleName(it.style) });
      });
      return { ok: true, where: "end" };
    },
    editParagraph: async function (index, text) {
      if (!paras[index]) throw new FriendlyError("Paragraph " + index + " does not exist");
      paras[index].text = String(text);
      return { ok: true };
    },
    replaceText: async function (find, replace, all) {
      var hits = 0;
      for (var i = 0; i < paras.length && (all || hits === 0); i++) {
        if (paras[i].text.toLowerCase().indexOf(String(find).toLowerCase()) >= 0) {
          paras[i].text = paras[i].text.split(find).join(replace);
          hits++;
        }
      }
      return { ok: true, hits: hits };
    },
    insertTable: async function (op) {
      tables.push({ values: op.values });
      return { ok: true, index: tables.length - 1 };
    },
    insertComment: async function (find, text) {
      var target = null;
      for (var i = 0; i < paras.length; i++) {
        if (paras[i].text.toLowerCase().indexOf(String(find).toLowerCase()) >= 0) { target = i; break; }
      }
      comments.push({ anchor: target == null ? "(start)" : "para " + target + " (\"" + find + "\")", text: text });
      return { ok: true, anchor: target };
    },
    deleteParagraph: async function (index) {
      if (!paras[index]) throw new FriendlyError("Paragraph " + index + " does not exist");
      paras.splice(index, 1);
      return { ok: true };
    },
    createTemplateBlock: async function (title, fields) {
      templateBlocks.push({ title: title, fields: fields.map(function (f) { return f.label + ":" + f.type; }) });
      paras.push({ text: "[" + title + " template] " + fields.map(function (f) {
          return f.label + ": " + (f.options ? "{" + f.options.join("/") + "}" : "");
        }).join(" | "), style: "Normal" });
      return { ok: true, notes: [] };
    },
    setTrackChanges: async function (on) {
      trackChanges = !!on;
      return { ok: true };
    },
    formatText: async function (index, font) {
      if (!paras[index]) throw new FriendlyError("Paragraph " + index + " does not exist");
      paras[index].font = Object.assign(paras[index].font || {}, font);
      return { ok: true };
    },
    addImage: async function (payload) {
      // the sim has no caret: 'at cursor' just marks the line
      var img = { text: payload && payload.atCursor ? "[IMAGE at cursor]" : "[IMAGE]", style: images.length % 2 ? "Normal" : "Heading 3", image: true };
      paras.push(img);
      images.push({ b64: payload.b64 });
      return { ok: true };
    },
    rewriteDoc: async function (items) {
      paras = (items || []).map(function (it) {
        return { text: String(it.text == null ? "" : it.text), style: styleName(it.style) };
      });
      tables = [];
      return { ok: true };
    },
    resetAll: function () { images = []; reset(); },
    renderDebug: function (el) {
      var rows = [];
      rows.push("<h3>Simulator document</h3>");
      rows.push('<div class="opnote">' + paras.length + " paragraphs - " + tables.length +
        " table(s) - " + comments.length + " comment(s)</div>");
      paras.forEach(function (p, i) {
        rows.push('<div class="simrow"><b>[' + i + "]</b> <i>" + esc(p.style) + "</i> " +
          esc(String(p.text).slice(0, 120)) + "</div>");
      });
      tables.forEach(function (t, i) {
        rows.push('<div class="simrow"><b>table ' + (i + 1) + "</b>: " +
          (t.values || []).map(function (r) { return r.join(" | "); }).join(" // ") + "</div>");
      });
      comments.forEach(function (c) {
        rows.push('<div class="simrow"><b>comment</b> on ' + esc(c.anchor) + ": " + esc(String(c.text).slice(0, 100)) + "</div>");
      });
      el.innerHTML = rows.join("");
    },
  };
  // constructor that returns the singleton (boot does `new window.SimDeck()`)
  window.SimDeck = function () { return SimAPI; };
  window.SimAPI = SimAPI;
  window.__DOCKICK_SIM__ = true;
  window.__SK_SIM__ = true;  // the shared pane boot checks this flag
  window.__SimData = { paras: function () { return paras; }, tables: function () { return tables; }, comments: function () { return comments; } };
})();