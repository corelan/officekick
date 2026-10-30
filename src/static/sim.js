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
        shapes: s.shapes.map((sh, j) => ({
          idx: j,
          name: sh.name,
          type: sh.type,
          text: sh.text,
          isTitle: sh.isTitle,
          isBody: sh.isBody,
        })),
      }));
      return { slideCount: slides.length, slides: out, selected: selected.slice(), host: "simulator" };
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
    }
    async setSlideBody(i, lines) {
      const sh = this._shapeByRole(i, false);
      sh.text = Array.isArray(lines) ? lines.join("\n") : String(lines);
    }
    async setShapeText(i, shapeIdx, text) {
      if (!slides[i]) throw new FriendlyError("Slide " + i + " does not exist");
      const sh = slides[i].shapes[shapeIdx];
      if (!sh) throw new FriendlyError("Shape " + shapeIdx + " does not exist on slide " + i);
      sh.text = String(text);
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
            '<span class="idx">[' + i + ']</span><span class="t">' + esc(s.title) + "</span>" +
            (bullets ? "<ul>" + bullets + "</ul>" : "") +
            (s.notes ? '<div class="notes">notes: ' + esc(s.notes) + "</div>" : "") +
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