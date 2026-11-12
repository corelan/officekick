// MailKick simulator: in-memory mailbox (3 fake messages + a compose draft)
// used to exercise the Outlook task pane in a regular browser. Exposes the same
// SimDeck interface the pane boot expects (window.__MAILKICK_SIM__). Each
// message carries a stable id (itemId) so the pane can detect a different
// email being opened (0.9.233 session handling) - "open email" buttons in the
// debug panel simulate clicking another message in the Outlook list.
(function () {
  function esc(s) { return String(s).replace(/[&<>\"]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]; }); }
  "use strict";

  var messages = [
    {
      id: "sim-mail-1",
      from: "soc-team@corelan.example",
      to: "you@corelan.example",
      subject: "Threat hunting kickoff - draft notes",
      body: "Hunters collect signals across the environment and reason about them\n" +
            "before an intrusion detonates. This working draft covers: data sources,\n" +
            "hypothesis-driven hunting, and AI-assisted triage of phishing mail.\n" +
            "Next: we want a slide summary and some catchy titles for the course deck.",
      categories: ["Blue Team"],
    },
    {
      id: "sim-mail-2",
      from: "alerts@corelan.example",
      to: "you@corelan.example",
      subject: "Phishing report - credential harvesting page",
      body: "Received by 12 users so far. The link leads to a credential harvesting\n" +
            "page that clones our vpn-portal. Domain: vpn-portal.corelan.example.\n" +
            "Recommend: block the domain, sweep the mailboxes for the sender, and\n" +
            "draft a user-awareness note.",
      categories: ["Red Team"],
    },
    {
      id: "sim-mail-3",
      from: "qa@corelan.example",
      to: "you@corelan.example",
      subject: "Quarterly QA review schedule",
      body: "Quick reminder: the quarterly QA review is scheduled for next Tuesday.\n" +
            "Please add your test-plan updates to the shared tracker before Friday\n" +
            "so the report can be generated in time.",
      categories: [],
    },
  ];
  var activeIdx = 0;        // which message is "open" in the reading pane
  var compose = null; // {subject, bodyText} when mode = compose
  var replyWin = null;      // simulated reply window (draft_reply)
  var forwardWin = null;    // simulated forward window (forward op)
  var meetingWin = null;    // simulated meeting window (new_meeting op)
  var lastWin = null;       // most recently opened simulated window
  var log = [];

  class FriendlyError extends Error { constructor(m) { super(m); this.friendly = true; } }

  function currentMessage() { return messages[activeIdx]; }

  var SimAPI = {
    sim: true,
    word: "Outlook",
    capabilities: function () { return { text: true, host: "Outlook" }; },
    currentEmailId: function () {
      // stable identity of the "open" item - the pane keys new-session-per-
      // email resets on this (0.9.233)
      return compose ? "sim-compose-draft" : currentMessage().id;
    },
    getContext: async function (opts) {
      var minimal = !!(opts && opts.minimal);
      var msg = currentMessage();
      var multi = [
        { subject: "Phishing report - credential harvesting page", sender: "alerts@corelan.example" },
        { subject: "Quarterly QA review schedule", sender: "qa@corelan.example" },
        { subject: "Lunch orders for Friday", sender: "office@corelan.example" },
      ];
      if (compose) {
        return { host: "Outlook", mode: "compose", sender: msg.from, to: msg.to,
                 subject: compose.subject, bodyText: minimal ? "" : compose.bodyText, categories: [], warnings: [], minimal: minimal,
                 itemId: "sim-compose-draft", selectedItems: multi, masterCategories: ["Blue Team", "Red Team", "AI"] };
      }
      return { host: "Outlook", mode: "read", sender: msg.from, to: msg.to,
               subject: msg.subject, bodyText: minimal ? "" : msg.body, categories: (msg.categories || []).slice(), warnings: [], minimal: minimal,
               itemId: msg.id, selectedItems: multi,
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
    editBodyInPlace: async function (html) {
      // 0.9.348 sim: a read-mode draft in the simulator = always editable
      // (the real Outlook may refuse; the pane falls back to the text)
      if (compose) {
        compose.bodyText = String(html);
      } else if (item && item.bodyText !== undefined) {
        item.bodyText = String(html);
      } else {
        throw new FriendlyError("editing in place failed (simulator)");
      }
      log.push("draft edited in place (simulated) with " + String(html).length + " chars");
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
      var msg = currentMessage();
      var subj = (ctx && ctx.subject) || msg.subject;
      forwardWin = { to: (op.to && op.to.slice(0)) || [], subject: subj,
                     note: String(op.html || ""), quoted: String((ctx && ctx.bodyText) || msg.body) };
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
    resetAll: function () { activeIdx = 0; compose = null; replyWin = null; forwardWin = null; meetingWin = null; lastWin = null; log = []; },
    renderDebug: function (el) {
      function reRender() { SimAPI.renderDebug(el); }
      var rows = [];
      rows.push("<h3>Simulator email</h3>");
      var btns = messages.map(function (m, i) {
        return '<button class="btn subtle" data-sim-mail="' + i + '"' +
          (i === activeIdx ? ' disabled style="opacity:.5"' : "") +
          ">" + esc(m.subject.slice(0, 30)) + "</button>";
      }).join(" ");
      rows.push('<div class="opnote">open email (simulates clicking a message in the Outlook list): ' + btns + "</div>");
      rows.push('<div class="opnote">mode: ' + (compose ? "COMPOSE (draft open)" : "READ (message selected)") +
        " - <b>itemId</b>: " + esc(SimAPI.currentEmailId()) + " - " + log.length + " applied op(s)</div>");
      var msg = currentMessage();
      if (compose) rows.push('<div class="simrow"><b>COMPOSE:</b> ' + esc(compose.subject) + "</div>");
      if (replyWin) rows.push('<div class="simrow sim-reply"><b>' + (replyWin.all ? "REPLY-ALL WINDOW (simulated):" : "REPLY WINDOW (simulated):") + '</b><br />' + replyWin.html + "</div>");
      if (meetingWin) rows.push('<div class="simrow sim-meeting"><b>MEETING FORM (simulated):</b> ' + esc(meetingWin.subject) +
        " - " + esc(meetingWin.when) + (meetingWin.location ? " - " + esc(meetingWin.location) : "") +
        (meetingWin.attendees ? "<br />required/optional: " + esc(meetingWin.attendees) : "") + "</div>");
      if (forwardWin) rows.push('<div class="simrow sim-forward"><b>FORWARD WINDOW (simulated):</b> FW: ' + esc(forwardWin.subject) +
        (forwardWin.to.length ? " - to: " + esc(forwardWin.to.join(", ")) : " - recipients: you pick") +
        (forwardWin.note ? "<br /><i>note:</i><br />" + forwardWin.note : "") + "</div>");
      rows.push('<div class="simrow"><b>FROM</b> ' + esc(msg.from) + " - <b>SUBJECT</b> " + esc(msg.subject) + "</div>");
      rows.push('<div class="simrow"><b>ATTACHMENTS</b> hunt-plan.pptx (1.4 MB), ioc-list.csv (40 kB)</div>');
      rows.push('<div class="simrow"><b>BODY</b><br />' + esc(msg.body).split("\n").join("<br />") + "</div>");
      (log || []).forEach(function (l) { rows.push('<div class="simrow"><i>' + esc(l) + "</i></div>"); });
      el.innerHTML = rows.join("");
      var openBtns = el.querySelectorAll("[data-sim-mail]");
      Array.prototype.forEach.call(openBtns, function (b) {
        b.addEventListener("click", function () {
          activeIdx = parseInt(b.getAttribute("data-sim-mail"), 10) || 0;
          if (compose) { compose = null; log.push("compose draft closed (new email opened)"); }
          log.push("opened email: " + currentMessage().subject.slice(0, 50));
          reRender();
        });
      });
    },
  };
  window.SimDeck = function () { return SimAPI; };
  window.SimAPI = SimAPI;
  window.__MAILKICK_SIM__ = true;
  window.__SK_SIM__ = true;
  window.__SimData = { email: function () { return currentMessage(); }, log: function () { return log; } };
})();
