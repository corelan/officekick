// MailKick simulator: in-memory mailbox (one fake message + a compose draft)
// used to exercise the Outlook task pane in a regular browser. Exposes the same
// SimDeck interface the pane boot expects (window.__MAILKICK_SIM__).
(function () {
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]; }); }
  "use strict";

  var email = {
    from: "soc-team@corelan.example",
    to: "you@corelan.example",
    subject: "Threat hunting kickoff - draft notes",
    body: "Hunters collect signals across the environment and reason about them\n" +
          "before an intrusion detonates. This working draft covers: data sources,\n" +
          "hypothesis-driven hunting, and AI-assisted triage of phishing mail.\n" +
          "Next: we want a slide summary and some catchy titles for the course deck.",
    categories: ["Blue Team"],
  };
  var compose = null; // {subject, bodyText} when mode = compose
  var log = [];

  class FriendlyError extends Error {}

  var SimAPI = {
    sim: true,
    word: "Outlook",
    capabilities: function () { return { text: true, host: "Outlook" }; },
    getContext: async function (opts) {
      var minimal = !!(opts && opts.minimal);
      if (compose) {
        return { host: "Outlook", mode: "compose", sender: email.from, to: "you@corelan.example",
                 subject: compose.subject, bodyText: minimal ? "" : compose.bodyText, categories: [], warnings: [], minimal: minimal };
      }
      return { host: "Outlook", mode: "read", sender: email.from, to: email.to,
               subject: email.subject, bodyText: minimal ? "" : email.body, categories: email.categories.slice(), warnings: [], minimal: minimal };
    },
    draftReply: async function (html) {
      log.push("reply window opened with " + String(html).length + " chars");
      return { ok: true };
    },
    composeBody: async function (html) {
      if (!compose) throw new FriendlyError("compose_body works only on an email being composed");
      compose.bodyText = String(html);
      log.push("compose body replaced");
      return { ok: true };
    },
    setCategories: async function (names) {
      var known = { "blue team": 1, "red team": 1, "ai": 1 };
      (names || []).forEach(function (n) {
        if (!known[String(n).toLowerCase()]) throw new FriendlyError("category not in your master list: " + n);
      });
      log.push("categories added: " + names.join(", "));
      return { ok: true };
    },
    newMessage: async function (payload) {
      // opens a simulated compose draft prefilled - the user "sends" it
      compose = { subject: String(payload.subject || "(no subject)"),
                  bodyText: String(payload.html) };
      log.push("new message window opened (prefilled)");
      return { ok: true };
    },
    saveDraft: async function () {
      if (!compose) throw new FriendlyError("save draft works only while composing");
      log.push("draft saved");
      return { ok: true };
    },
    resetAll: function () { compose = null; log = []; },
    renderDebug: function (el) {
      var rows = [];
      rows.push("<h3>Simulator email</h3>");
      rows.push('<div class="opnote">mode: ' + (compose ? "COMPOSE (draft open)" : "READ (message selected)") +
        " - " + log.length + " applied op(s)</div>");
      if (compose) rows.push('<div class="simrow"><b>COMPOSE:</b> ' + esc(compose.subject) + "</div>");
      rows.push('<div class="simrow"><b>FROM</b> ' + esc(email.from) + " - <b>SUBJECT</b> " + esc(email.subject) + "</div>");
      rows.push('<div class="simrow"><b>BODY</b><br />' + esc(email.body).split("\n").join("<br />") + "</div>");
      (log || []).forEach(function (l) { rows.push('<div class="simrow"><i>' + esc(l) + "</i></div>"); });
      el.innerHTML = rows.join("");
    },
  };
  window.SimDeck = function () { return SimAPI; };
  window.SimAPI = SimAPI;
  window.__MAILKICK_SIM__ = true;
  window.__SK_SIM__ = true;
  window.__SimData = { email: function () { return email; }, log: function () { return log; } };
})();
