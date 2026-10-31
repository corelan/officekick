# Changelog
## [0.4.3] - 2026-10-30
### Fixed
- move_slide shipped without its OPS_META whitelist entry ("unknown op type
  'move_slide'" in the gate); the entry is in and the op is E2E-verified in
  the simulator (move last slide -> position 2 reorders correctly).

## [0.5.3] - 2026-10-30
### Fixed
- Settings sub-tab headers: readable in Office dark mode (lime) and a darker green in light mode.

## [0.6.4] - 2026-10-30
### Added / fixed
- Resuming a session now loads the last up to 6 messages as collapsed
  request/response boxes (Inspector-style, all collapsed, click to open)
  right above the prompt, so it visually resumes where you left off.
- The Settings sub-tab bar and the top tab strip wrap when the pane is
  narrow (no more Inspector falling out of view).

## [0.6.3] - 2026-10-30
### Added
- Every busy status (Asking ..., collecting context, generating image, applying, model-output repair) now shows a rotating indicator character so it reads as still-working.

## [0.9.5] - 2026-10-31
### Fixed - DocKick tracked-changes error after applying edits
- With "AI edits land as tracked changes" ON, every apply ended with
  "revisions: The property 'changeTrackingMode' is not available..." and the
  revision setting silently switched itself off. Cause: the pane read
  `document.changeTrackingMode` on the Word.js proxy without loading it first
  (modern Word.js throws exactly that error for unloaded properties). The
  read/write now follows the documented pattern - `isSetSupported("WordApi",
  "1.6")` guard, `document.load("changeTrackingMode")`, `context.sync()`,
  then set - and after the batch the pane RESTORES the tracking mode the
  document had before the AI edits instead of always switching it off
  (a mode you had enabled yourself stays enabled).
### Fixed - tracked-changes setting dropped on reopen
- The server's `/api/config` projection omitted `trackedChanges`: the pane
  only "remembered" the setting while it stayed open in the same session (it
  was assigned client-side right after saving) and silently lost it after
  every reopen - edits then landed as direct edits even though the setting
  said ON. The value is now part of `/api/config` (and stored by default),
  so the checkbox survives reloads like every other setting.

## [0.9.4] - 2026-10-31
### Changed - narrow-pane settings layout
- Connections view: instead of a wide table (Name | Type | Base URL | Key),
  each connection is now a stacked block - name + Test button on the first
  line, then type / url / key / "in use" below one another. Applied to all
  three add-ins.
- Approval policy table: the three policy options (auto / confirm / deny) now
  render BELOW the action name and its "what it does" description in a block
  per action, instead of a cramped third column. Radios keep their
  save/restore behaviour (verified round-trip).
- Containers changed from tables to stacked div blocks; word-break discipline
  kept; no horizontal overflow at pane width.

### Fixed - MailKick Connections/Profiles panels were dead
- The MailKick pane carried duplicate legacy sections (tab-connections +
  tab-profiles) whose element ids collided with the live Settings sub-tab
  panels: every id lookup bound to the hidden copies, so the Connections and
  Profiles views under Settings rendered empty and none of their buttons
  responded. The dead sections are removed; the live panels now populate and
  their edit forms fill on click (verified: connections list, edit form fill,
  profiles list, profile edit fill, approval policy radio round-trip).

## [0.9.3] - 2026-10-31
### Fixed - DocKick format_text on real Word
- Applying format actions in real Word failed with "ctx.document.body.paragraphs
  .getItemAt is not a function": Word.js collections have no `getItemAt`. The
  two format paths (whole paragraph and `find` fragment) now load the paragraph
  collection, sync, and index `.items` - the same pattern the existing
  edit/delete paragraph actions already use - with a clear error when the
  paragraph number is out of range.
- Gate descriptions for paragraph actions read the paragraph one position
  AFTER the requested one (off-by-one: guarded `paragraphs[i-1]`, read
  `paragraphs[i]`), which showed the wrong text and crashed the run when the
  target was the last paragraph of the document.
- E2E in the DocKick simulator: index-based + find-based format actions
  approved and applied together.

## [0.9.2] - 2026-10-31
### Fixed - reply +/- / forward "window" visibility
- In the simulator, draft replies and forwards now open a REAL visible
  simulated window: a bordered card appears directly in the chat output right
  after the run ("Simulated reply/forward window - real Outlook opens it for
  review"), and the forward window is also rendered in the Sim email panel
  (previously forwards were only a log line, replies only in the panel).
  Real Outlook behaviour is unchanged and was already correct: applying a
  draft_reply opens `item.displayReplyForm(...)`, applying a forward opens
  `Office.context.mailbox.displayNewMessageForm(...)` - the window appears
  after you press Apply (or immediately when auto-apply is on).
- About tab: "Pane build" row shows the build of the JS actually executing in
  the pane (injected by the server with every HTML response) and flags it
  STALE when it does not match the server version - the definitive answer to
  "am I running the new build?" in cached webviews.

## [0.9.1] - 2026-10-31
### Added - editing one item of a list (SlideKick)
- `replace_text` now supports a `find` fragment: with `{"find":"<the exact
  current words of that one bullet>"}` the pane replaces only that fragment and
  leaves the rest of the shape's text (and the bullet list) untouched. The
  fragment write is verified after applying; if a host ignores the sub-range
  write, the pane falls back to a full-shape rewrite of the same result.
  Without `find`, behaviour is unchanged (whole shape text).
- Deck outline sent to the model now carries up to 400 characters of each
  shape's text (was 220), and adds the rule: never rebuild a whole shape from a
  preview (outline previews end with `…[text continues]`) - edit list items
  with `replace_text` + `find` instead. This is the fix for "SlideKick claimed
  not to have access to the list" on larger slides/decks: the list text is
  included, and when only a preview fits, single items can still be edited
  safely in place.

## [0.9.0] - 2026-10-31
### Added - per-add-in auto-apply
- "Auto apply approved actions" is now a per-add-in setting: one value for
  SlideKick, one for DocKick, one for MailKick (stored as `autoApplyByApp`).
  The toggle in each pane's Settings > Approval policy writes only that
  add-in's value; existing setups keep working (the previous global
  `autoApplyApproved` is the fallback when an add-in has no value yet).

### Fixed - MailKick settings wiring
- MailKick's pane carried a duplicated boot/events tail: an older copy of
  wireEvents/boot won via JS last-definition-wins and silently disabled newer
  wiring, including the auto-apply settings toggle (this is why the setting
  looked like it never persisted in MailKick). The duplicated tail is removed;
  the settings checkbox now loads, writes and re-opens with the stored
  per-add-in value (E2E: OFF cycle, reopen, ON cycle, cross-add-in values).

### Fixed - stale pane scripts after updates
- The server now tags the pane's script/style assets with the current version
  (`?v=<version>` on served taskpane/simulator HTML), so browsers, Office hosts
  and the simulator can never execute a heuristically cached older build after
  an update; static responses were already `no-store`.

## [0.8.9] - 2026-10-31
### Fixed - settings round-trip
- Settings values saved in one visit were displayed OFF after closing and
  reopening the pane (the checkbox was wired before the config loaded and was
  never synced afterwards). The Settings "Auto apply approved actions" checkbox
  and the per-run auto-apply box now follow the stored setting on every config
  load; a manual toggle is never overridden by a later config refresh.
  (Persistence itself was never lost - the stored value was always correct.)

## [0.8.8] - 2026-10-31
### Changed - MailKick behaviour
- **draft_reply and forward no longer require approval** (default policy
  `auto`): MailKick opens the reply/forward window immediately and you send it
  yourself - nothing is ever sent automatically. The global
  "Auto apply approved actions" setting now also ticks the per-run box before
  the all-auto evaluation, so the first run of a session skips the preview
  correctly.
- **Action requests now propose ops, not prose**: a hard rule for the model -
  any reply/forward/draft/send/categorize request must come back as an op
  (`draft_reply` / `forward` / ...), never as chat text of "what I would say".
- **New `forward` op** (READ mode, default auto): opens a forward window for
  the selected mail with the quoted original, the model's optional note
  above it and optional recipients (Outlook build providing
  `displayNewMessageForm`; the simulator simulates the window).

## [0.8.7] - 2026-10-31
### Changed - naming
- The Outlook add-in folder/URL is now `mailkick/` everywhere, matching the
  add-in's name: docs moved to docs/mailkick.md, the add-in static tree is
  served at /mailkick/static/, and the manifest points there. Older
  sideloaded manifests (the legacy /outlook/static URL) keep working: the
  server aliases it to the mailkick tree, and installers clean up the old
  deployed folder. The Windows installer (install.ps1) now also deploys
  MailKick and renders its manifest into the trusted catalog.

## [0.8.6] - 2026-10-31
### Added
- The approval gate: a red **X** per proposed action to remove that action from
  the list (first click arms a "sure?", the second removes; the Apply totals
  and Select all respect the removal).
- **Run refreshes the context readout**: every Run updates the deck/document/
  email summary line and the simulator view, in all three add-ins.
### Changed / docs
- docs/architecture.md rewritten for the current reality: three add-ins
  (incl. MailKick + the outlook/ tree), the real API surface (sessions,
  templates, image, test-connection, update-check), current op lists per
  app, the sessions/Previous-requests model, model-side web lookups, gate
  features, and the token-free update check.
- README gained a "How to update OfficeKick" section.
- docs/dockick.md: the op table now lists insert image / format text /
  create template block with their real default policies.

## [0.8.5] - 2026-10-31
### Added
- MailKick quick actions `reply +` / `reply −` (green/red): one click drafts a
  polite reply from the open conversation in its own tone (formal/informal),
  positive or politely negative, and proposes a `draft_reply` action behind
  the approval gate (MailKick still never sends).
### Changed
- **Previous requests** is now the title of an always-open box: every run of
  the (resumed) session is its own +/- box with the timestamp first and the
  request text below it; the title itself is not a button anymore.
- Resuming a session titled just "Session" asks the AI for a better title
  based on the interactions and stores it (History shows real titles).
### Fixed
- The approval-policy table and the Behaviour group wrap their text instead of
  overflowing in narrow panes.

## [0.8.4] - 2026-10-31
### Changed - run/input interaction
- Run no longer blanks the input instantly: the request text stays visible
  with the box disabled while the run works (visibly dimmed); the Run button
  stays the only disabled-while-busy gate as before.
- When the run reports a result, the box clears and the pane prints
  "Request: <your text>" right above the response / proposed actions.
- On failure (unreachable model, retries exhausted): the input is re-enabled
  and keeps the text for a retry; the failed attempt is still persisted.
- The box is always re-enabled when switching tabsheets and back, clicking
  New session, or resuming a session.

## [0.8.3] - 2026-10-31
### Added
- A global "Auto apply approved actions" setting in Settings -> Approval
  policy (Behaviour group, persisted, default off). When on, every run's
  auto-apply checkbox follows it (tick immediately and it applies when
  gates render); when off the per-run checkbox stays manual. The user's
  manual per-run override is remembered for that pane.
  Verified: on -> persisted + the run checkbox ticks; off -> manual again.

## [0.8.2] - 2026-10-31
### Fixed / Changed
- Session titles: a reasoning model burned the previous 24-token title
  budget on hidden reasoning and returned nothing, so every AI title failed
  silently and sessions stayed "session". The title call now gets 300
  tokens; on failure the pane guarantees a rename to the derived request
  name; the server rename route additionally repairs model answers
  (reasoning fragments, code fences, "Title:"-labels, long essays).
- Titles confirmed working with the user's real profile ("Sliding Context
  Windows Explained").
- Collapsed state of the per-run boxes in Previous requests: the head now
  shows the first 50 characters of the request + the timestamp; + opens the
  full request (click-to-refill an empty input) and the response/actions
  summary; - collapses.
### Confirmed
- The session store keeps EVERY exchange (never truncated to 6); only the
  model context and the History-tab preview are capped at 6.

## [0.8.1] - 2026-10-31
### Changed
- Web lookups are no longer a user option: the checkbox is gone. The model
  decides - when a request needs up-to-date info it returns
  {"kind":"web_lookup"} (prompt rule everywhere), the pane runs the
  web-enabled call automatically, then re-asks with the fresh results
  (once; no loops; honesty fallbacks kept). Inspector-visible.
- Proposed actions are color-coded for dark-mode readability: green type
  labels for applications/edits/modifications (layout, formatting, notes,
  tables, shapes, images, reply/draft/message ops), red for deletions
  (delete_slide/shape) in light and dark mode (E2E-verified colors).

## [0.8.0] - 2026-10-31
### Changed - Previous-requests as group box with per-turn boxes
- "Previous requests" is the titled group-box head (boxed +/-, turn count);
  inside it every run is its own collapsible box: the head shows the boxed
  +/-, the request date/time and a snippet; opening it shows the full
  request (clicking it refills the input box, only when the input is empty)
  and the response / proposed-actions summary.
- Filing now happens the moment Run is clicked: the previous run moves out
  of the top zone into the box immediately, so during a run only the new
  activity sits between the prompt and the history box.
### Added
- Inspector: failed requests render in unmistakable red (border, header,
  badge) in light and dark mode.

## [0.7.9] - 2026-10-31
### Fixed
- Live lookup actually switched to the model side in all panes (0.7.8
  described it; the pane rework had not landed): one call through the
  profile's model, ":online" appended on openrouter connections, honesty
  fallbacks (NO_WEB_ACCESS / can't browse / no web access) -> the run
  continues and the status says the model has no web access.
### Verified
- Live E2E: model cites 5 fresh sources incl. an Oct-2026 benchmark score.

## [0.7.8] - 2026-10-31
### Changed
- Settings -> General regrouped into boxed sections per add-in: a
  "General (all add-ins)" box (update check + debug logging) and a
  tool-specific box ("PowerPoint-specific" = speaker notes; "Word-specific"
  = tracked changes; MailKick has none). The handlers/ids are unchanged.
- Web lookups now run through the MODEL exclusively: the previous
  local-search machinery (Brave/Tavily keys, DuckDuckGo scraping and the
  /api/websearch route) is removed. Models on an openrouter connection get
  the ":online" suffix for the lookup call; other providers are asked and
  honestly report "no web access" (status shows it, the run continues).
### Fixed
- The update check no longer mentions or uses GitHub tokens; it is a plain
  unauthenticated VERSION fetch (the repo goes public).

## [0.7.7] - 2026-10-31
### Added
- **Live web lookups** (optional): a "live web lookups" checkbox in the chat
  tab. When ticked, the run happens in two phases: a small model call
  proposes 0-3 web search queries, the server executes them (stdlib only;
  Brave or Tavily with a BRAVE_SEARCH_API_KEY / TAVILY_API_KEY env var,
  otherwise keyless DuckDuckGo HTML), and the results ride into the main
  prompt as WEB RESULTS the model can cite. Search failures degrade
  gracefully ("no results"/"failed"), visible in the status line.
  Verified live end to end through OpenRouter (searched + answered).
- **Auto-retry on model calls**: transport failures (connection refused,
  5xx, 429) retry automatically up to 5 attempts, 10 s apart, with the
  attempt counter in the status line; business errors (4xx) fail fast.
- **Timer format**: the request timer now reads "Mmin sssec(s)" (e.g.
  "0min 08sec(s)"), hours as "1h 02min 03sec(s)".
- **Connection test**: Settings -> Connections gets a Test button per
  connection (server /api/test-connection = cheap list-models roundtrip);
  the row shows reachability + model count + ms, or the failure reason.

## [0.7.6] - 2026-10-31
### Fixed - session-history loss (user-reported: a 4-request session with a
title was "gone" after reopening the add-in)
- The session PUT route overwrote the AI/derived title with the generic
  "session" (the name was not carried over), replacing what the History tab
  showed, and the store wrote the caller's exchange list over the stored one
  without reading the file first, so a shorter payload (e.g. a capped
  history after a resume) shrank a stored session.
- Now the stored file is the source of truth for titles and volume: the PUT
  keeps the existing name and keeps the longer of the stored/received
  exchange lists; the panes also verify an update actually stuck (a store
  that returns a suspiciously small session releases the dead id and saves
  as a fresh one next time).
- Reproduced before the fix (4 runs -> 2 msgs, title lost) and after the fix
  (4 runs -> 8 msgs, title kept, resume loads all 4 turns).

## [0.7.5] - 2026-10-31
### Changed
- Request history: each stored exchange carries a local timestamp; the
  Previous-requests rows and the session peek show "Request: <date time> ..."
  (older sessions without a timestamp simply show without it).
- Ops replies are stored as a compact recap only: "Executed N action(s):
  type x count, ..." (no action text, no image prompts; text replies are
  stored verbatim). Applies to the session store AND the model context.
- Clicking a stored request (previous-requests row or session-peek bubble)
  reloads the full request into the input box, but only when the box is
  empty - a typed draft is never clobbered. The chat input no longer shows
  an example; empty = empty.
### Fixed
- A failed request (model unreachable etc.) now puts the request text back
  into the input box so it can be retried with Run; the failed exchange is
  still saved to the session as "(failed: ...)".
- Session sanitizer dropped per-exchange timestamps; it now keeps them.

## [0.7.4] - 2026-10-30
### Changed
- Update check preferences: "Check for updates automatically at startup" is
  now a persisted setting (on by default) in Settings -> General, together
  with a "Check now" button. The boot check runs only when the setting is on.
- The instructions shown on an available update spell out the manual steps
  (git pull in the local clone, rerun install.sh / install.ps1, reload the
  task pane) and explicitly say no automatic git pull is attempted.

## [0.7.3] - 2026-10-30
### Changed
- The proposed-actions box shows only the actions now; the approval-policy
  summary line was removed (policy editing stays in Settings -> Approval
  policy, and the per-action approval controls remain on each row).

## [0.7.2] - 2026-10-30
### Changed
- Tab order: History now sits between Sessions and Settings in all three
  add-ins (Sessions | History | Settings | About [| Templates]).

## [0.7.1] - 2026-10-30
### Fixed
- After a resume, the "Previous requests (N)" box existed but was never made
  visible and its header click was not wired (updates lost on the resume
  path); resumes now show the collapsed box with the past turns, and the
  header toggles with the boxed +/- indicator. Verified end to end.

## [0.7.0] - 2026-10-30
### Added
- **Update check without git**: the server fetches the VERSION file from the
  GitHub repo (contents API with a raw fallback, cached 10 minutes) and
  compares it with the running version. Settings → General now has a
  "Check for updates" button; one silent check also runs at boot.
  Shows "up to date (x, latest y)" or "Update available: y (you run x)"
  with the update instructions (git pull + install.sh / install.ps1).
- Private-repo support: put a read-only token in the env var
  OFFICEKICK_GITHUB_TOKEN or ~/.officekick/github_token (never logged,
  never rendered in the panes).
- Test hook: /api/update-check?latest=X forces the remote value.

## [0.6.9] - 2026-10-30
### Fixed / Changed
- The "Previous requests (N)" box now actually shows its content: up to the
  6 most recent interactions as flat Request/Response rows, newest to oldest
  (resume fills it from the saved history; live runs file each finished turn).
  Each interaction shows the request text and the response/gate summary.
- The box header got a boxed +/− indicator (plus = collapsed, minus = open).
- Lost wirePastHeader during the refactor restored (that was the empty box).

## [0.6.8] - 2026-10-30
### Fixed
- Resume failed with "can't find variable jend" (a leaked variable from the
  past-requests refactor); resumes verified against a real mail session.
- The "Approval needed / Review proposed actions" status line now carries the
  amber pending treatment in light and dark mode; the proposed-actions box
  itself: light-green tint in dark mode, soft mint in light mode, with a
  plain "Proposed actions" title and a small muted policy note instead of
  the uppercase banner.

## [0.6.7] - 2026-10-30
### Added
- A request timer next to Run: clicking Run starts an H:MM:SS counter; it freezes at the elapsed time when the run finishes (or fails) and restarts on the next Run. Verified live.

## [0.6.6] - 2026-10-30
### Added
- **Test runs stay out of history**: a built-in nosave flag - simulator runs
  and any pane loaded with ?nosave=1 never write to the session store (and
  skip the AI-naming call); add &save=1 to a simulator URL to save anyway.
  Verified: a sim run added no session rows.
- History titles render immediately at tab open (already instant from the
  derived name; the AI name replaces it when it arrives).
### Verified
- Session storage integrity: every session file holds even, role-paired
  exchanges with correct app tags (2 live mailkick sessions intact).

## [0.6.5] - 2026-10-30
### Changed
- chat: the prompt sits at the top; below it the CURRENT response + proposed
  actions; below that all previous turns are filed into ONE collapsible
  "Previous requests (N)" box (newest first inside, click to expand; each
  turn opens/closes its response too). Resuming fills the box from the
  stored history; New session empties it.
- fixed: the busy-status spinner printed a literal unicode escape in CSS;
  it now renders the actual rotating character.

## [0.6.2] - 2026-10-30
### Changed
- **Inspector moved into Settings** as a sub-tab (General / Approval policy /
  Connections / Profiles / Inspector); the top strip is
  Sessions | Settings | About | History (- Templates in MailKick).
- Inspector entries render as **grouped boxes, newest on top**, capped at 6
  visible with a "(N older not shown)" note; each box shows a **timestamp
  + HTTP status + method/path + duration** in its header; only the newest
  box opens by default and all boxes open/close by clicking their header.
- Status bar no longer shows the connection count or "debug off" (it only
  prints "debug ON" when debugging is enabled).

## [0.6.1] - 2026-10-30
### Changed (chat flow)
- Run now turns the prompt box into a chat thread: each run prints its
  **Request** above the **Response**; older turns stack above (the thread
  keeps the last ~15 turns). Clicking a request line copies it back into the
  input box for a quick re-ask.
- The prompt box is smaller (4 rows) and clears the moment a request is sent.

## [0.6.0] - 2026-10-30
### Changed (Chat tab layout)
- The chat tab is now grouped: **Profile / route** in its own group box;
  **Current session** starts with the context info card (purple in light
  mode, cyan in dark), then the include-context + auto-apply checkboxes and
  the Refresh button, a subtle divider, and then the session label, prompt
  box, Run and the output. The New session button sits below the group box.
- Pane footer: a lock icon shows when the pane uses https; the default
  profile name is no longer printed. Also fixed 5 leaked DocKick strings
  inside MailKick (About row, API warnings, unreachable-server messages).

## [0.5.2] - 2026-10-30
### Changed
- Deck outline character cap raised from 40,000 to 60,000 (per request).

## [0.5.1] - 2026-10-30
### Fixed
- Large decks: the deck outline was hard-capped at 20,000 characters and 60
  slides - a 121-slide deck with rich text got truncated around slide 38 and
  the model correctly reported it could only see those. The outline now
  covers up to 200 slides; decks with more than 40 slides get a COMPACT
  outline (every slide with title + a 90-char body preview instead of
  per-shape dumps); the character cap is 40,000; the prompt instructs the
  model to work large tasks in batches of max 20 ops and state what remains.

## [0.5.0] - 2026-10-30
### Changed (UI restructure)
- Settings now hosts sub-tabs: General (behavior toggles + debug logging),
  Approval policy (its own sheet), Connections, Profiles. Top strip =
  Sessions | Settings | About | History | Inspector - Connections/Profiles
  moved under Settings; IDs kept; error-path navigation lands right.
- Word: AI edits as tracked changes is now the DEFAULT (on unless switched
  off in Settings > General; existing installs flipped on as well).

## [0.4.2] - 2026-10-30
### Added
- **move_slide** op (SlideKick): reorders a slide to any position. PowerPoint
  exposes no reorder primitive, so the deck composes it: export the slide
  (exportAsBase64Presentation, 1.10), delete it, then insertSlidesFromBase64
  at the final position (1.2 - after a target slide, or at the beginning when
  no target). Confirm-approve by default and it must run alone in a batch
  (it reindexes the deck). New slides still land at the end (no positional
  addSlide API) - the prompt tells the model to create-then-move.

## [0.4.1] - 2026-10-30
### Changed
- Sessions persist on every request/response pair: the autosave runs
  immediately when an exchange completes (no 600 ms debounce), and failed
  requests are kept too as "(failed: <reason>)", so resuming shows what was
  asked and what went wrong. Verified with a dead provider endpoint.

## [0.4.0] - 2026-10-30
### Changed (protocol: human numbering)
- **Slide and shape numbers in SlideKick ops are now 1-based** ("slide 2" and
  "shape 2" = the second slide/shape, exactly what a human means; previously
  they were zero-based - the model passed your number through and "slide 2"
  edited the third slide). The deck outline, the selection report and the
  simulator view number everything 1..N; the deck-adapter boundary converts
  back to zero-based internally.
- **Paragraph numbers in DocKick ops are now 1-based** for the same reason
  (previously zero-based; tables were already 1-based - inconsistent).
- Word tables and the MailKick batch list were already human-numbered; no
  indexed ops exist in MailKick.

## [0.3.6] - 2026-10-30
### Fixed / added
- Sessions tab: clicking a session name expands/collapses a preview of the
  last up to 6 entries (3 requests + 3 responses), with shorter snippets.
- live-report fix: the create_shape/move_shape formatters crashed on
  op.geometry of raw or failed ops ("undefined is not an object") and masked
  the underlying error; every geometry access is guarded now.

## [0.3.5] - 2026-10-30
### Added / clarified
- Honest-capability rule in the SlideKick prompt: SmartArt, stock icons and
  3D models cannot be created through Office.js (no authoring API exists).
  Existing SmartArt/model3D shapes are still detected in the deck context
  (ShapeType enum, PowerPointApi 1.4); the model now explains this and offers
  shape-based alternatives instead of inventing actions.

## [0.3.4] - 2026-10-30
### Added
- **create_shape** op (SlideKick): the model can insert text boxes (with text)
  and geometric shapes (rectangle/ellipse/roundedRectangle/hexagon/diamond/
  triangle) plus straight lines at a slide-local position. PowerPointApi 1.4
  (addTextBox/addGeometricShape/addLine) - far more available than the earlier
  roadmap note assumed. Auto-approve default; position/format ops can target
  the new shape afterwards from the refreshed context.

## [0.3.3] - 2026-10-30
### Changed
- **Sessions are scoped per application** (reversing the earlier shared
  design, by request). Each pane tags its autosaves with its app and lists
  only its own sessions: SlideKick / DocKick / MailKick. Sessions created
  before this change have no tag and remain visible in all three panes so
  nothing disappears; resuming one in a pane re-tags it to that app.
- GET /api/sessions?app=... filter + app field on session records; new test.

## [0.3.2] - 2026-10-30
### Fixed (live-report from PowerPoint)
- Office.js load() is not a promise - chaining .then() on it breaks on real
  hosts (the simulator could not catch this). Rewritten to the canonical
  load(...) + ctx.sync() pattern in formatTextSubstring (find-based
  formatting), formatSelection and the selection context probe.

## [0.3.1] - 2026-10-30
### Added
- **Sub-range formatting** for both `format_text` ops: the model can now
  format only PART of a shape/paragraph by including `"find":"exact words"`.
  Powered by PowerPoint `TextRange.getSubstring` (1.4, the deck code locates
  the phrase at runtime) and Word `paragraph.search` (1.1, all matches in the
  paragraph). Selection stays the fastest path for "what I have selected";
  find covers everything else.
- Verified separately: PowerPoint `TextRange.setSelected()` (1.8+) exists, so
  the add-in CAN programmatically move the selection on your build.

## [0.3.0] - 2026-10-30
### Added (Office.js API exploration wave)
- **SlideKick**: selection-driven ops - `format_selection` formats the text you
  have selected right now (PowerPointApi 1.6 getSelectedTextRange); the context
  now reports your selection (selected slides, shapes and text preview) so the
  model can target what you have open; `set_background` (AI-generated picture
  or solid color via 1.10 slide background fills); `add_hyperlink` (1.10
  shape.setHyperlink); `set_alt_text` (1.10 accessibility alt text).
- **DocKick**: `create_template` - appends a titled template block with real
  Word content controls (plain text, dropdown with AI-provided options,
  checkbox, date picker; WordApi 1.5 + 1.9 with graceful fallbacks); a
  Settings toggle "AI edits as tracked changes" applies your ops through
  Word's Track Changes so you can accept/reject them like human edits
  (changeTrackingMode, feature-detected).
- **MailKick**: Mailbox 1.13 multi-select awareness - the context now lists the
  emails you multi-selected (batch triage prompts for the model).
- New default approvals: format_selection/format_text/set_background/
  create_template = confirm; add_hyperlink/set_alt_text = auto.

## [0.2.4] - 2026-10-30
### Fixed
- **MailKick "there was no email"**: the model now always receives at least a
  minimal context (subject + sender + read/compose mode) even when the
  include-email-context box is unticked; optional reads (recipients /
  categories / body) degrade individually instead of failing the whole
  collection. Outlook wording fixed ("email", not "deck"/"document").
### Changed
- Outlook manifest: standard VersionOverrides for the read/compose ribbon
  buttons; `RequestedHeight` 480 -> 350 (schema max) - the official validator
  now passes; `install.sh` GUID refs reconciled to the shipped manifest Id.
### Repo
- `tests/` is untracked and gitignored (local-only); docs/mailkick.md rewritten
  (templates, sessions, Inspector, install/uninstall, limits).

## [0.2.3] - 2026-10-30
### Added
- **Formatting**: a `format_text` op. SlideKick: bold / italic / color / size /
  name / underline / allCaps + alignment and bullets on a shape's text
  (PowerPointApi 1.4, `textFrame.textRange.font`). DocKick: the same set +
  highlight color and justify alignment on a paragraph. The model gets the
  allowed attributes in the schema and emits only what you asked for.
## [0.2.2] - 2026-10-30
### UI
- Chat/Actions: the session name (AI-generated) shows above the Profile/Route
  controls; the prompt box has a title; "New conversation" moved to the bottom
  of the chat area (green-bordered).
## [0.2.1] - 2026-10-30
### Added
- **MailKick named templates**: create/edit/delete small reusable content blocks
  (Templates tab); say "reply using template <name> ..." or "new mail from
  template <name> ..." and MailKick injects only that template into the AI's
  context; for replies the template + the original mail + instructions drive the
  draft; for new mail a `new_message` op opens a prefilled compose window. You
  always review and send - MailKick never sends anything automatically.
- Templates list: name/date/length cards, name-click = content peek, Use button
  prefills the prompt.
### UI
- "New conversation" now also clears the prompt box and has a green border.
## [0.2.0] - 2026-10-30
### Added
- **Sessions** (all three add-ins): conversations persist locally in
  `~/.officekick/sessions/` (0600 files) and are auto-saved after every exchange.
  A new **Sessions tab** lists them with the AI-generated name, message count and
  the last-updated time, with Resume / Delete (+ click-to-rename). Resume loads
  the stored history into the Chat/Actions tab (last 6 exchanges stay in context).
  Names come from a tiny AI call through the profile's other model; a failed
  naming call falls back to the first user message.
### Fixed
- Outlook/OWA sideloading: the mail add-in manifest is now https-only (Outlook's
  dialog rejects http SourceLocations with "Installation failed"), and the loopback
  leaf certificate gained modern extensions (Subject/Autority Key Identifier,
  serverAuth) - strict TLS consumers like OWA's checker accept it now. Cert file
  and keychain CA verified end-to-end (localhost/127.0.0.1/[::1] all strict-verified).

## [0.1.8] - 2026-10-30
### Added
- **MailKick** - the Outlook add-in (docs/mailkick.md): email context, draft_reply /
  compose_body / set_categories / save_draft ops, read+compose modes, simulator,
  installer manifest render + Mac sideload instructions + the Windows catalog step.
### Fixed
- Image refinement flows: the conversation history no longer stores raw ops JSON -
  compact per-op summaries (+ the last image prompt, capped) are kept instead, so a
  "refine the image" round writes a fresh prompt instead of echoing blobs. Image
  prompt cap raised 1500 -> 4000 chars; system prompt rule: refinements = fresh
  self-contained prompts (delta only). Verified via the wire payload.
- SlideKick: OfficeDeck.addImage was missing (real-PowerPoint image runs aborted
  with 'deck.addImage is not a function'); implemented via the Common API
  setSelectedDataAsync(Image) with optional left/top/width/height (points) -
  targets the slide you have open.

## [0.1.7] - 2026-10-30
### Added
- Word: **insert at cursor** - a pane checkbox feeds the image to the caret/selected
  range (`getSelection().insertInlinePictureFromBase64(..., "Replace")`, WordApi 1.2);
  unchecked = end of document (unchanged).
- OpenRouter image pre-flight: text-only model ids are rejected up front with the
  closest image-capable suggestion(s), instead of a raw provider 404
  ("No model found"). Example: picking `google/gemini-3.1-flash-lite` now suggests
  `google/gemini-3.1-flash-lite-image`.
### Fixed
- `provider_chat` unbound `text` crash for unreachable openai-family providers.

## [0.1.6] - 2026-10-30
### Added
- **Inspector tab** in both panes (request/response log): every call to the local server (chat, models,
  image) is logged with HTTP status, duration, connection/model, and request and
  response snippets; errors show in red with the provider's message (billing,
  auth, rate limits, unknown model IDs, network failures). Client-side, last 80
  entries, Clear/Refresh buttons.
### Fixed
- Image generation: busy status ("Generating image…") while waiting, and a hard
  300 s timeout so a hung provider can no longer leave "Applying…" forever.
- Server logs each image call ([image] start/done/FAILED) to server.log.

## [0.1.5] - 2026-10-30
### Added
- **Pictures**: `create_image` (SlideKick) / `insert_image` (DocKick) ops - the model
  returns a detailed image prompt; the pane fetches pixels via a new server route
  `POST /api/image` (OpenAI-style images API; provider types without image
  generation answer with a clear error) and inserts them into the open PowerPoint
  slide / Word document end. Approval default: confirm (costs tokens).
- Optional third profile slot: **image model** (connection + model), shown as an
  extra column in the Profiles list; deleting a connection used by an image
  reference is blocked; renames re-point it.
- Mock provider serves `/v1/images/generations` with a tiny PNG; +1 test (22 total).
### Fixed
- About tab: copyright (Corelan Consulting bv) + repo & LICENSE links.
- Default manifests ship without ribbon commands (Office-Mac 16.114 registers but
  does not render them; see docs/slidekick.md).
