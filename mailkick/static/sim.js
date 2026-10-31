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
  var replyWin = null;      // simulated reply window (draft_reply)
  var forwardWin = null;    // simulated forward window (forward op)
  var meetingWin = null;    // simulated meeting window (new_meeting op)
  var lastWin = null;       // most recently opened simulated window
  var log = [];

  class FriendlyError extends Error {}

  var SimAPI = {
    sim: true,
    word: "Outlook",
    capabilities: function () { return { text: true, host: "Outlook" }; },
    getContext: async function (opts) {
      var minimal = !!(opts && opts.minimal);
      var multi = [
        { subject: "Phishing report - credential harvesting page", sender: "alerts@corelan.example" },
        { subject: "Quarterly QA review schedule", sender: "qa@corelan.example" },
        { subject: "Lunch orders for Friday", sender: "office@corelan.example" },
      ];
      if (compose) {
        return { host: "Outlook", mode: "compose", sender: email.from, to: "you@corelan.example",
                 subject: compose.subject, bodyText: minimal ? "" : compose.bodyText, categories: [], warnings: [], minimal: minimal, selectedItems: multi, masterCategories: ["Blue Team", "Red Team", "AI"] };
      }
      return { host: "Outlook", mode: "read", sender: email.from, to: email.to,
               subject: email.subject, bodyText: minimal ? "" : email.body, categories: email.categories.slice(), warnings: [], minimal: minimal, selectedItems: multi,
               attachments: [{ name: "hunt-plan.pptx", size: 1482334 }, { name: "ioc-list.csv", size: 41210 }] };
    },
    draftReply: async function (html, all) {
      // the sim can not open a real Outlook reply window - show the drafted
      // reply as a simulated reply window in the debug area instead
      replyWin = { html: String(html), all: !!all };
      lastWin = { kind: "reply", html: replyWin.html, all: !!all };
      log.push("reply window opened (simulated) with " + String(html).length + " chars");
      return { ok: true };
    },
    composeBody: async function (html) {
      if (!compose) throw new FriendlyError("compose_body works only on an email being composed");
      compose.bodyText = String(html);
      log.push("compose body replaced");
      return { ok: true };
    },
    setCategories: async function (names) {
      // sim mirrors the real degraded mode: the fake mailbox master list
      var known = ["Blue Team", "Red Team", "AI"];
      var unknown = (names || []).filter(function (n) { return known.indexOf(n) < 0; });
      log.push("categories added: " + names.join(", "));
      return { ok: true, note: unknown.length ? "(applied WITHOUT checking the master list; new: " + unknown.join(", ") + ")" : "" };
    },
    newMessage: async function (payload) {
      // opens a simulated compose draft prefilled - the user "sends" it
      compose = { subject: String(payload.subject || "(no subject)"),
                  bodyText: String(payload.html) };
      log.push("new message window opened (prefilled)");
      return { ok: true };
    },
    forwardMessage: async function (op, ctx) {
      var subj = (ctx && ctx.subject) || email.subject;
      forwardWin = { to: (op.to && op.to.slice(0)) || [], subject: subj,
                     note: String(op.html || ""), quoted: String((ctx && ctx.bodyText) || email.body) };
      lastWin = { kind: "forward", to: forwardWin.to, subject: forwardWin.subject, note: forwardWin.note, quoted: forwardWin.quoted };
      log.push("forward window opened (simulated) - FW: " + subj +
        ", note " + String(op.html || "").length + " chars" +
        (op.to && op.to.length ? ", to " + op.to.join(",") : ", recipients to pick"));
      return { ok: true };
    },
    saveDraft: async function () {
      if (!compose) throw new FriendlyError("save draft works only while composing");
      log.push("draft saved");
      return { ok: true };
    },
    newMeeting: async function (op) {
      meetingWin = {
        subject: String(op.subject || "(no subject)"),
        when: String(op.start || "?") + " to " + String(op.end || "?"),
        location: String(op.location || ""),
        attendees: [].concat(op.required || [], op.optional || []).join(", "),
        body: String(op.body || ""),
      };
      lastWin = { kind: "meeting", subject: meetingWin.subject, when: meetingWin.when,
                  location: meetingWin.location, attendees: meetingWin.attendees, body: meetingWin.body };
      log.push("meeting form opened (simulated) - " + meetingWin.subject);
      return { ok: true };
    },
    lastWindow: function () { return lastWin; },
    resetAll: function () { compose = null; replyWin = null; forwardWin = null; meetingWin = null; lastWin = null; log = []; },
    renderDebug: function (el) {
      var rows = [];
      rows.push("<h3>Simulator email</h3>");
      rows.push('<div class="opnote">mode: ' + (compose ? "COMPOSE (draft open)" : "READ (message selected)") +
        " - " + log.length + " applied op(s)</div>");
      if (compose) rows.push('<div class="simrow"><b>COMPOSE:</b> ' + esc(compose.subject) + "</div>");
      if (replyWin) rows.push('<div class="simrow sim-reply"><b>' + (replyWin.all ? "REPLY-ALL WINDOW (simulated):" : "REPLY WINDOW (simulated):") + '</b><br />' + replyWin.html + "</div>");
      if (meetingWin) rows.push('<div class="simrow sim-meeting"><b>MEETING FORM (simulated):</b> ' + esc(meetingWin.subject) +
        " - " + esc(meetingWin.when) + (meetingWin.location ? " - " + esc(meetingWin.location) : "") +
        (meetingWin.attendees ? "<br />required/optional: " + esc(meetingWin.attendees) : "") + "</div>");
      if (forwardWin) rows.push('<div class="simrow sim-forward"><b>FORWARD WINDOW (simulated):</b> FW: ' + esc(forwardWin.subject) +
        (forwardWin.to.length ? " - to: " + esc(forwardWin.to.join(", ")) : " - recipients: you pick") +
        (forwardWin.note ? "<br /><i>note:</i><br />" + forwardWin.note : "") + "</div>");
      rows.push('<div class="simrow"><b>FROM</b> ' + esc(email.from) + " - <b>SUBJECT</b> " + esc(email.subject) + "</div>");
      rows.push('<div class="simrow"><b>ATTACHMENTS</b> hunt-plan.pptx (1.4 MB), ioc-list.csv (40 kB)</div>');
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
