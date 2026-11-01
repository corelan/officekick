/*
 * SlideKick simulator shim.
 *
 * Defines an in-memory deck + a Deck implementation with the same interface as
 * OfficeDeck, so taskpane.js runs unmodified in a plain browser (no Office.js).
 * Loaded by simulator.html BEFORE taskpane.js; sets window.__SK_SIM__ so the
 * task pane knows it is not inside PowerPoint.
 */
window.__SK_SIM__ = true;

(function () {
  "use strict";

  const fakeSelectedText = "AI agents are reshaping security work";
  const fakeSelectedFont = {};
  const INITIAL = [
    {
      layout: "title_content",
      title: "SlideKick simulator deck",
      body: ["run taskpane.js logic without PowerPoint", "ops are applied to this in-memory deck"],
      notes: "",
    },
    {
      layout: "title_only",
      title: "Second slide",
      body: [],
      notes: "",
    },
    {
      layout: "section",
      title: "Wrap-up",
      body: ["thanks for testing"],
      notes: "",
    },
  ];

  let slides = [];
  let selected = [0];

  function reset() {
    slides = JSON.parse(JSON.stringify(INITIAL)).map((s) => ({
      layout: s.layout,
      title: s.title,
      notes: s.notes || "",
      shapes: [
        { name: "Title 1", type: "placeholder", text: s.title || "", isTitle: true, isBody: false },
        { name: "Content Placeholder 2", type: "placeholder", text: Array.isArray(s.body) ? s.body.join("\n") : "", isTitle: false, isBody: true },
      ],
    }));
    selected = slides.length ? [0] : [];
  }
  reset();

  class FriendlyError extends Error {}

  function normalizeLayout(kind) {
    if (!kind) return "blank";
    const k = String(kind).toLowerCase();
    if (k.includes("content")) return "title_content";
    if (k.includes("only")) return "title_only";
    if (k.includes("section")) return "section";
    return "blank";
  }

  class SimDeck {
    constructor() {
      this.sim = true;
    }
    capabilities() {
      return { text: true, selection: true, host: "simulator" };
    }
    async getContext() {
      const out = slides.map((s, i) => ({
        idx: i,
        title: s.shapes[0] ? s.shapes[0].text : "",
        slideWidth: 960, slideHeight: 540,
        shapes: s.shapes.map((sh, j) => ({
          idx: j,
          name: sh.name,
          type: sh.type,
          text: sh.text,
          isTitle: sh.isTitle,
          isBody: sh.isBody,
          left: sh.left != null ? sh.left : 20,
          top: sh.top != null ? sh.top : 20,
          width: sh.width != null ? sh.width : 300,
          height: sh.height != null ? sh.height : 50,
          table: sh.table || null,
        })),
      }));
      return { slideCount: slides.length, slides: out, selected: selected.slice(), host: "simulator" };
    }
    async readSlides(idxs) {
      // read_slide: full per-shape text of the given 1-based slides
      const uniq = Array.from(new Set(idxs || [])).sort((a, b) => a - b);
      if (!uniq.length) throw new FriendlyError("readSlides: no slide numbers given");
      const parts = [];
      for (const n of uniq) {
        const sl = slides[n - 1];
        if (!sl) throw new FriendlyError("Slide " + n + " does not exist (deck has " + slides.length + " slide(s))");
        const lines = ["SLIDE " + n + " (slide " + n + " in the outline):"];
        const shs = sl.shapes || [];
        if (!shs.length) lines.push("  (no shapes)");
        shs.forEach((sh, j) => {
          const txt = String(sh.text == null ? "" : sh.text).replace(/\r|\n/g, " | ");
          const label = (sh.name || sh.type || "shape") + ", " + (sh.type || "");
          lines.push('  shape ' + (j + 1) + ' (' + label + '): "' + txt + '"');
        });
        parts.push(lines.join("\n"));
      }
      return parts.join("\n");
    }
    async addTable({ values, slide, geometry }) {
      const i = slide;
      if (!slides[i]) throw new FriendlyError("Slide " + i + " does not exist");
      geometry = geometry || {};
      slides[i].shapes.push({
        name: "Table " + values.length + "x" + values[0].length,
        type: "table",
        text: "",
        isTitle: false,
        isBody: false,
        left: geometry.left != null ? geometry.left : 40,
        top: geometry.top != null ? geometry.top : 150,
        width: geometry.width != null ? geometry.width : 600,
        height: 30 * values.length,
        table: { rows: values.length, cols: values[0].length, values: values },
      });
      return { ok: true };
    }
    async formatText(slideIndex, shapeIndex, font) {
      const sh = slides[slideIndex] && slides[slideIndex].shapes[shapeIndex];
      if (!sh) throw new FriendlyError("Shape " + shapeIndex + " on slide " + slideIndex + " does not exist");
      Object.assign(sh, { font: Object.assign(sh.font || {}, font) });
      return { ok: true };
    }
    async formatTextSubstring(slideIndex, shapeIndex, findText, font) {
      const sh = slides[slideIndex] && slides[slideIndex].shapes[shapeIndex];
      if (!sh) throw new FriendlyError("Shape " + shapeIndex + " on slide " + slideIndex + " does not exist");
      const t = String(sh.text || "");
      const i = t.toLowerCase().indexOf(String(findText).toLowerCase());
      if (i < 0) throw new FriendlyError('"' + findText + '" not found in the current text of that shape.');
      sh.fontRuns = sh.fontRuns || [];
      sh.fontRuns.push({ from: i, len: findText.length, font: font });
      return { ok: true };
    }
    async formatSelection(font) {
      if (!fakeSelectedText) throw new FriendlyError("Select some text in the presentation first, then run this again.");
      Object.assign(fakeSelectedFont, font);
      return { ok: true, text: fakeSelectedText };
    }
    async moveSlide(from, to) {
      if (from === to) return { ok: true };
      if (from < 0 || from >= slides.length) throw new FriendlyError("Slide " + (from + 1) + " does not exist");
      slides.splice(to, 0, slides.splice(from, 1)[0]);
      return { ok: true };
    }
    async createShape(slideIndex, op) {
      const sl = slides[slideIndex];
      if (!sl) throw new FriendlyError("Slide " + slideIndex + " does not exist");
      const g = op.geometry || {};
      sl.shapes.push({
        idx: sl.shapes.length,
        name: op.kind,
        type: op.kind,
        text: op.kind === "textbox" ? (op.text || "") : undefined,
        left: g.left, top: g.top, width: g.width, height: g.height,
        rotation: g.rotation != null ? g.rotation : undefined,
        fill: op.fill || undefined,
        line: op.line || undefined,
        transparency: op.transparency != null ? op.transparency : undefined,
      });
      return { ok: true };
    }
    async setBackground(slideIndex, payload) {
      const sl = slides[slideIndex];
      if (!sl) throw new FriendlyError("Slide " + slideIndex + " does not exist");
      sl.background = payload.kind === "image"
        ? { kind: "image", b64len: (payload.b64 || "").length, transparency: payload.transparency }
        : { kind: payload.kind, color: payload.color, transparency: payload.transparency };
      return { ok: true };
    }
    async duplicateSlide(from) {
      const src = slides[from];
      if (!src) throw new FriendlyError("Slide " + (from + 1) + " does not exist");
      const copy = JSON.parse(JSON.stringify(src));
      slides.splice(from + 1, 0, copy);
      return { ok: true, at: from + 1 };
    }
    async duplicateSlideTo(from, to) {
      if (from === to) return { ok: true };
      const src = slides[from], tgt = slides[to];
      if (!src || !tgt) throw new FriendlyError("Slide does not exist");
      slides[to] = JSON.parse(JSON.stringify(src));
      return { ok: true };
    }
    async clearSlide(from) {
      const sl = slides[from];
      if (!sl) throw new FriendlyError("Slide " + (from + 1) + " does not exist");
      let cleared = 0;
      sl.shapes.forEach((sh) => {
        if (typeof sh.text === "string") { sh.text = ""; cleared++; }
      });
      const before = sl.shapes.length;
      sl.shapes = sl.shapes.filter((sh) => sh.type !== "line" || typeof sh.text === "string");
      delete sl.background;
      return { ok: true, cleared, removed: before - sl.shapes.length };
    }
    async addHyperlink(slideIndex, shapeIndex, url, tip) {
      const sh = slides[slideIndex] && slides[slideIndex].shapes[shapeIndex];
      if (!sh) throw new FriendlyError("Shape " + shapeIndex + " on slide " + slideIndex + " does not exist");
      sh.hyperlink = { address: url, screenTip: tip || "" };
      return { ok: true };
    }
    async mergeDeck(b64) {
      // the sim cannot parse a .pptx: appends two placeholder slides to show
      // the merge flow
      const n = slides.length;
      for (let k = 1; k <= 2; k++) {
        slides.push({
          layout: "title_content",
          title: "Imported slide " + k + " (from source deck)",
          notes: "",
          shapes: [
            { name: "Title 1", type: "placeholder", text: "Imported slide " + k, isTitle: true, isBody: false },
            { name: "Content Placeholder 2", type: "placeholder", text: "imported from ~" + Math.round((b64 || "").length * 3 / 4 / 1024) + " kB source deck", isTitle: false, isBody: true },
          ],
        });
      }
      return { ok: true, count: 2, at: n + 1 };
    }
    async styleShape(slideIndex, shapeIndex, fill, line, transparency) {
      const sh = slides[slideIndex] && slides[slideIndex].shapes[shapeIndex];
      if (!sh) throw new FriendlyError("Shape " + shapeIndex + " on slide " + slideIndex + " does not exist");
      if (fill) sh.fill = String(fill);
      if (line) sh.line = String(line);
      if (transparency != null) sh.transparency = Number(transparency);
      return { ok: true };
    }
    async setShapeImage(slideIndex, shapeIndex, b64, transparency) {
      const sh = slides[slideIndex] && slides[slideIndex].shapes[shapeIndex];
      if (!sh) throw new FriendlyError("Shape " + shapeIndex + " on slide " + slideIndex + " does not exist");
      sh.fillImage = { b64len: (b64 || "").length };
      if (transparency != null) sh.transparency = Number(transparency);
      return { ok: true };
    }
    async setAltText(slideIndex, shapeIndex, description, title) {
      const sh = slides[slideIndex] && slides[slideIndex].shapes[shapeIndex];
      if (!sh) throw new FriendlyError("Shape " + shapeIndex + " on slide " + slideIndex + " does not exist");
      sh.altTextDescription = description;
      if (title) sh.altTextTitle = title;
      return { ok: true };
    }
    getSelectionInfo() {
      return Promise.resolve({ slides: [0], shapes: [], textPreview: fakeSelectedText });
    }
    async addImage({ b64, slide, geometry }) {
      const i = slide;
      if (!slides[i]) throw new FriendlyError("Slide " + i + " does not exist");
      geometry = geometry || {};
      slides[i].shapes.push({
        name: "Image (generated)",
        type: "image",
        text: "",
        isTitle: false,
        isBody: false,
        left: geometry.left != null ? geometry.left : 40,
        top: geometry.top != null ? geometry.top : 100,
        width: geometry.width != null ? geometry.width : 400,
        height: geometry.height != null ? geometry.height : 240,
        image: { b64: b64 },
      });
      return { ok: true };
    }
    async moveShape(slideIndex, shapeIndex, geometry) {
      const s = slides[slideIndex];
      if (!s || !s.shapes[shapeIndex]) throw new FriendlyError("Shape " + shapeIndex + " on slide " + slideIndex + " not available");
      ["left", "top", "width", "height"].forEach((p) => {
        if (geometry[p] != null) s.shapes[shapeIndex][p] = geometry[p];
      });
      if (geometry.rotation != null) s.shapes[shapeIndex].rotation = geometry.rotation;
      return { ok: true };
    }
    async addSlide({ layout, title, body, notes }) {
      const lay = normalizeLayout(layout);
      const sh = [
        { name: "Title 1", type: "placeholder", text: title || "", isTitle: true, isBody: false },
      ];
      if (lay !== "blank") {
        sh.push({
          name: "Content Placeholder 2",
          type: "placeholder",
          text: Array.isArray(body) ? body.join("\n") : "",
          isTitle: false,
          isBody: true,
        });
      }
      slides.push({ layout: lay, title: title || "", notes: typeof notes === "string" ? notes : "", shapes: sh });
      return { index: slides.length - 1, notesManual: null };
    }
    _shapeByRole(i, wantTitle) {
      if (!slides[i]) throw new FriendlyError("Slide " + i + " does not exist");
      return slides[i].shapes.find((sh) => (wantTitle ? sh.isTitle : sh.isBody));
    }
    async setSlideTitle(i, text) {
      const sh = this._shapeByRole(i, true);
      sh.text = String(text);
      slides[i].title = String(text);   // keep the convenience field in sync for the view
    }
    async setSlideBody(i, lines) {
      const sh = this._shapeByRole(i, false);
      sh.text = Array.isArray(lines) ? lines.join("\n") : String(lines);
      slides[i].body = Array.isArray(lines) ? lines : String(lines).split("\n");
    }
    async setShapeText(i, shapeIdx, text) {
      if (!slides[i]) throw new FriendlyError("Slide " + i + " does not exist");
      const sh = slides[i].shapes[shapeIdx];
      if (!sh) throw new FriendlyError("Shape " + shapeIdx + " does not exist on slide " + i);
      sh.text = String(text);
    }
    async setShapeTextFind(i, shapeIdx, findText, replText) {
      if (!slides[i]) throw new FriendlyError("Slide " + i + " does not exist");
      const sh = slides[i].shapes[shapeIdx];
      if (!sh) throw new FriendlyError("Shape " + shapeIdx + " does not exist on slide " + i);
      const full = String(sh.text || "");
      const idx = full.toLowerCase().indexOf(String(findText).toLowerCase());
      if (idx < 0) throw new FriendlyError('"' + findText + '" not found in that shape\'s current text.');
      sh.text = full.slice(0, idx) + String(replText == null ? "" : replText) + full.slice(idx + String(findText).length);
      return { ok: true, mode: "fragment" };
    }
    async setNotes(i, text) {
      if (!slides[i]) throw new FriendlyError("Slide " + i + " does not exist");
      slides[i].notes = String(text);
      return { manual: false };
    }
    async deleteSlide(i) {
      if (!slides[i]) throw new FriendlyError("Slide " + i + " does not exist");
      slides.splice(i, 1);
      selected = selected.includes(i) ? [] : selected;
    }
    async applyLayout(i, kind) {
      if (!slides[i]) throw new FriendlyError("Slide " + i + " does not exist");
      const lay = normalizeLayout(kind);
      slides[i].layout = lay;
      if (lay === "blank" && !slides[i].shapes.find((s) => s.isBody)) {
        // real PowerPoint: blank layout drops placeholders; keep sim simple
      }
    }
    async getSelection() {
      return selected.slice();
    }
    slideCount() {
      return slides.length;
    }
    renderDebug(el) {
      const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
      const lines = slides
        .map((s, i) => {
          const bullets = (s.shapes.find((sh) => sh.isBody) ? s.shapes.find((sh) => sh.isBody).text : "")
            .split(/\r|\n/)
            .filter((x) => x.trim())
            .map((x) => "<li>" + esc(x) + "</li>")
            .join("");
          return (
            '<div class="sim-slide"><span class="lay">' + esc(s.layout) + "</span>" +
            '<span class="idx">[' + (i + 1) + ']</span><span class="t">' + esc(s.title) + "</span>" +
            (bullets ? "<ul>" + bullets + "</ul>" : "") +
            (s.notes ? '<div class="notes">notes: ' + esc(s.notes) + "</div>" : "") +
            s.shapes.filter((sh) => sh.fill || sh.line)
              .map((sh, k) => '<div class="notes">shape ' + (s.shapes.indexOf(sh) + 1) + ": " +
                (sh.fill ? "fill " + esc(sh.fill) : "") + (sh.fill && sh.line ? ", " : "") +
                (sh.line ? "line " + esc(sh.line) : "") + "</div>")
              .join("") +
            "</div>"
          );
        })
        .join("");
      el.innerHTML = lines || "(empty deck)";
    }
    static resetAll() {
      reset();
    }
    static get slides() {
      return slides;
    }
  }

  window.SimDeck = SimDeck;
})();