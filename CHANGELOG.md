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

## [0.9.11a] - docs, no release
### Docs
- New "Settings reference" section in all three add-in docs: the five
  Settings sub-tabs (General / Approval policy / Connections / Profiles /
  Inspector), what each setting does and its default (update check, debug
  logging, per-add-in auto-apply, policy modes + session approvals, gate
  behaviour, speaker notes on DocKick-only variants: tracked changes, and
  MailKick's never-send rules + context extras).

## 0.9.155
### Added
- The live-lookup channel now READS THE DOCUMENTS IT FINDS: after web
  lookup results come back, the pane fetches up to 2 source URLs through
  the new server endpoint GET /api/fetch-url (stdlib-only HTML-to-text
  extractor, page title + up to 8000 chars per page) and adds the page
  contents to the re-ask ("The lookups also RETURNED THE FULL TEXT of
  these documents/pages - read them before answering"). The model cites
  from the actual page text instead of just snippets. Guards: http(s)
  only, no localhost/private addresses (SSRF), 15 s timeout, 3 MB fetch
  cap. All three add-ins; status shows "Reading the found documents…".
## 0.9.154
### Fixed
- The chat tab's "New session" button (next to the request box) now runs
  the SAME flow as the History tab one - 0.9.153 only covered the History
  button, which is why a click could still say "Conversation history
  cleared." without a confirmation and leave the Previous requests box
  filled. Both buttons share one handler now: armed confirm ("Really
  start a new session?"), previous exchange kept in the session, request
  box cleared, Previous requests emptied, clean slate. All add-ins.
## 0.9.153
### Changed
- "New session" = two-stage confirm now (click once -> "Really start a
  new session?", click again within 3.5 s to proceed) - same armed-button
  pattern as session Delete.
- Before the reset, the current chat's last exchange is filed into the
  "Previous requests" box (so nothing disappears), and when the new
  session actually starts the Previous requests box is cleared - a new
  session begins on a clean slate. The resume hint marker is removed too,
  so the next run saves a genuinely fresh session. All three add-ins.
## 0.9.152
### Changed
- The table-unavailable note now states the platform truth: per
  Microsoft's requirement-set table, the PowerPoint JS table APIs
  (PowerPointApi 1.8) are Windows-only today - macOS Office reports them
  "Not available" on every Mac build so far, including the latest. So on
  macOS the model says plainly: no add-in can read (or edit) table cells
  there yet, independent of the Office version; on Windows it needs
  2504/build 18730+.
## 0.9.151
### Changed
- When a table cannot be read, the note the model returns is now
  user-actionable: it names the Office build and says exactly what fixes
  it ("update Office to a build with the PowerPointApi 1.8 tables, then
  ask again"), or shows the real load error. Confirmed on a Mac 16.114
  host: no shape.table on the add-in surface = no in-pane table reads or
  edits until Office ships the 1.8 table set to that build/channel.
## 0.9.150
### Fixed
- The outline never silently hides a shape any more: shapes whose text is
  empty/unreadable (tables in particular) are now listed as
  "(no readable text - table/gfx shape)" instead of vanishing. This is
  why slide 93's outline suddenly "lost" its table: the previous filter
  dropped shapes without text, so whether the model saw the table
  depended on the table-cell read succeeding - the shape was always
  there. Read_slide still lists every shape (it never filtered).
## 0.9.149
### Fixed
- (0.9.148 shipped the changelog text without the code - this commit
  carries the actual change.) Table reads now guard PowerPointApi **1.8**
  (Table/TableCell = 1.8, TableRow = 1.9, TableCell.text = a plain string
  per the Office JS docs) and read cells row-by-row
  (rows.load("items") + per-cell load("text")), with the deep-
  load ("items/cells/text") as the fallback. If the user's host still
  reports "cell text unavailable", the note now includes the exact cause
  (no shape.table on that build vs the load error) plus a table.read
  debug line.
## 0.9.148
### Fixed
- Table reads target the right API set: PowerPoint.Table/TableCell = 1.8
  (rows access 1.9), so the guard = 1.8 and cell text is read row-by-row
  (rows.load("items") + per-cell load("text")) instead of a deep
  items/cells/text load that can be rejected. The per-row fallback chain
  stays; failures keep surfacing their real reason.
### Verified (researched against the Office JS docs)
- PowerPoint.Table, TableCell = PowerPointApi 1.8; TableRow = 1.9;
  TableCell.text = a plain string. Older hosts than 1.8 = the honest
  "cell text unavailable" note.
## 0.9.147
### Added
- DocKick: "insert table of contents" op (create_toc, levels "1-3"|"1-2",
  optional title). Inserts a REAL Word TOC field (OOXML fldChar + instr
  TOC \o ".." \h \z \u) at the end of the document plus an optional
  TOC-Heading title paragraph; the model is told never to fake TOC entries
  as literal text. Word fills the entries when fields update (F9 / print).
  Simulator shim supported.
### Changed
- SlideKick table reads are now diagnosable: when shape.table is missing
  or the load fails, the note the model sees carries the real reason
  ("no shape.table here - a newer Office build is needed" / the load
  error) and a table.read debug line is written. Confirmed on a real
  Mac host: titles/text boxes/subtitles read fine; tables still depend
  on the host exposing shape.table.
## 0.9.146
### Fixed
- Tables were invisible to the model (correct claim - table shapes carry no
  text frame, so the outline and read_slide showed them as empty). Both
  Office paths (getContext + readSlides) now read table cells through the
  table API (PowerPointApi 1.6+, guarded): a table shape reports
  "[table RxC] cell1 | cell2 | ..." with the same per-shape text caps as
  other shapes. Old/unsupported builds degrade to an honest
  "[table - cell text unavailable in this Office build]" note instead of
  silence. The simulator shows the same cell text.
## 0.9.145
### Changed
- Deck-outline text limits raised: per-shape text cap 400 -> 900 chars,
  selected-slide deep listing 500 -> 900, big-deck per-slide body preview
  90 -> 240 and line cap 170 -> 340 (60 K total guard unchanged).
- The outline header now tells the model explicitly that per-shape text
  may be shortened and to emit a read_slide op to see a slide in full -
  "never claim you cannot see content without using read_slide" (the
  read-slide path shows every shape uncapped, up to the 60 K total).
  This answers cases like "I cannot see the last bullet of slide N": the
  outline cut the shape text, and the model now knows the fix.
## 0.9.144
### Changed
- The session label's title text is accented (indigo on light theme,
  bright indigo on dark) as are the History view's "current session"
  markers (the meta line + card highlight) - all three add-ins.
## 0.9.143
### Added
- The session label has two buttons: refresh (asks the AI for a new title,
  re-using the whole-recent-context ask) and edit (rename it inline; Enter
  confirms, Escape cancels).
### Changed
- A manual rename pauses automatic AI title updates for that session - the
  flag ("manualName") is stored in the session file, so it works for
  existing sessions too; sessions without the flag = AI-titled and stay
  eligible for automatic updates. Pressing refresh clears the pause. The
  flag round-trips through create/update/rename/list/get.
## 0.9.142
### Changed
- Session-title updates consider the session's whole recent picture:
  together with the existing title, the last up to 6 requests are listed
  (oldest first, 90 chars each) and the model is told to consider ALL of
  them - not only the newest - and to reply KEEP only when the existing
  title still covers that mix. All three add-ins.
## 0.9.141
### Changed
- Built-in prompt, Formatting block: the monospace trigger list now also
  covers filenames and folder paths ("shell, filenames, folder paths,
  commands, CLI flags, code, syntax, URLs or protocol elements").
## 0.9.140
### Changed
- "Reload pane" is now "Reload add-in", and after a reload the pane boots
  straight into Settings -> General (the reload tags the URL with a
  one-shot reloadnav=general parameter that is cleaned from the address
  right after boot, so a manual reload later behaves normally). All
  three add-ins.
## 0.9.139
### Changed
- The request box stands out: indigo border, faint indigo background tint
  and a soft blur shadow; a slightly brighter variant in dark mode. The
  focus outline now matches the border color. All three add-ins.
## 0.9.138
### Changed
- Wording fixes: the custom-prompt hint no longer says "below" (there is
  nothing below to point at); the resume note says "Just type in the
  request box to continue." instead of "Continue typing below" (the input
  box sits above the note) and drops the "below" from the Previous
  requests reference.
## 0.9.137
### Changed
- The custom system prompt editor is bigger: 16 rows, taller minimum
  height, capped at ~half the pane height, vertically resizable.
## 0.9.136
### Changed
- Built-in prompt, Formatting block: the monospace rule is now explicit and
  actionable - every line/cell with shell commands, CLI flags, code,
  syntax, URLs or protocol elements gets its separate Consolas format_text
  request ("never treat such lines as prose; when in doubt, ask").
## 0.9.135
### Fixed
- Cut-off recovery no longer runs actions twice: when the continuation's
  ops are merged with the salvaged ones, exact duplicate ops (the
  continuation very often repeats what it already emitted - it sees its
  own partial list in the context) are dropped deterministically, and the
  proposal note reports how many were removed.
- The continuation ask now says explicitly that the previously recovered
  ops are already kept and must NOT be repeated (emit only new ops, or
  reply done when nothing is missing).
## 0.9.134
### Changed
- Batch validation now understands slides the batch creates itself: a
  create_slide op raises a batch-local "virtual slide count" (the validator
  sees the deck as it will be after those creates), so later format_text /
  edit/replace ops in the same batch may legally target those new slides.
  Shape-in-context and find-text checks are deferred to apply time for
  slides that only exist after the batch's own creates (a wrong shape
  number still fails honestly there). Fixes "bad slide index: slide X
  does not exist" storms for multi-slide build-out batches.
- Built-in default prompt updated to the three blocks: Content (accuracy),
  Writing style, and Formatting (reuse the document's existing style/theme
  elements; code/syntax/protocol/command lines get a separate
  monospaced-font request, e.g. Consolas).
### Fixed
- Session titles: the model replying exactly "KEEP" (title still fits) no
  longer renames the session to the literal word KEEP - it keeps the
  current title (all three add-ins).
- Restart-flag watcher poll relaxed from 1 s to 10 s (a restart is never
  urgent; near-zero overhead).
## 0.9.133
### Added
- Server self-restart. The server watches <config dir>/restart.flag (1 s
  poll): on detection it removes the flag, generates a detached helper
  (restart_helper.py + restart_payload.json in the config dir), spawns it
  and stops serving. The helper waits for the port to free up, starts the
  next instance with the OS-appropriate launcher - macOS: the
  OfficeKickServer login-item app; Windows: the OfficeKickServer scheduled
  task (falling back to a detached pythonw relaunch of the same command
  line); other setups (manual/crontab): the exact original restart - waits
  until the new instance answers, then exits. Every server boot removes a
  stale flag first, so a restart cannot loop. Problems are appended to
  logs/restart.log.
- POST /api/server-restart arms a restart (optionally with a cause);
  /api/health now reports a "gen" boot stamp so a caller can see when a
  restart happened.
- Panes: a "Restart server" button in the update-check row (all three
  add-ins) + "Restart server now" directly after a successful git pull -
  both POST /api/server-restart and poll health until the generation
  changes ("Server restarted (new generation).").
## 0.9.132
### Changed
- Recovered-from-cut-off batches that then fail validation completely no
  longer paint a wall of dead rows: the proposal box collapses into one
  error line ("None of the N action(s) survived validation ... No changes
  were applied") plus an "Ask again (same request)" button that re-runs
  the same instruction (restores it from the last instruction when the
  input was cleared). Mixed batches (some valid) keep the row list.
- "bad slide index" reasons improved: "bad slide index: slide X does not
  exist (deck has N slides)" so the mismatch is obvious at a glance.
- The cut-off continuation ask is now dynamic (continueAsk(ctx)): it
  anchors the model to the actual slide count, 1-based numbering, the
  outline-only shape references, and tells it not to re-cover handled
  slides - this counters hallucinated slide indices in recovered batches.
  Word/Mail get the same all-invalid collapse + anchored continue ask.
## 0.9.131
### Changed
- Copy buttons now use the "two overlapping squares" copy icon instead of
  the clipboard glyph, drawn as an inline SVG (two stroked rounded rects)
  so it renders identically on every host font: the last-request echo
  button, the past-request copy affordance on step titles, and the
  manual-notes per-slide copy buttons. Word/Mail also gained the local
  clipboard helper they were missing behind the echo button (navigator
  .clipboard with an execCommand fallback), and the slidekick past-turn
  copy row uses it too.
## 0.9.130
### Added
- The last-request echo (the line above the proposed-actions box) is now
  interactive in all add-ins: clicking it always copies the request text
  to the clipboard, and when the input box is empty the click also fills
  the input with the same text ("Copied + filled the empty input box."
  vs "Copied to the clipboard."); a small clipboard button sits at the
  right of the echo with the same behavior (hover underline + cursor
  affordances).
## 0.9.129
### Changed
- Built-in prompt follows your rewording (style block): "Identify
  code/syntax/protocol/command instructions and make a separate request to
  put them in a monospaced font such as Consolas".
- SlideKick system prompt gained the matching operational rule so the ask
  is executable: when a bullet/shape contains code, commands, protocol
  fragments or syntax (not ordinary prose), emit a separate format_text op
  with font {"name":"Consolas"} (Courier New acceptable) on that shape.
  (format_text already supports font name; DocKick/MailKick have no
  font-name op, so there the style line reads as guidance only.)
## 0.9.128
### Changed
- Built-in quality/style prompt, style block: the monospaced-font line was
  reworded - "Identify code/syntax/protocol/command instructions and make
  a separate request to put them in a monospaced font such as Consolas"
  (identify + a separate format_request instead of a vague inline ask).
## 0.9.127
### Changed
- Connections / Profiles: the extra sheet headings were misplaced (they
  sat above the edit forms); removed. The form titles themselves
  ("New connection" / "New profile", and "Edit connection: X" /
  "Edit profile: X") now carry the stand-out color - deep teal on light,
  bright cyan on dark. The "Approval policy per action" title keeps it.
## 0.9.126
### Changed
- Built-in quality/style prompt updated: the accuracy block now also
  forbids inventing routines/solutions ("facts" emphasis), and the style
  block asks for code/syntax/protocol text in a monospaced font
  (e.g. Consolas) via a format_request. Dropped the "one idea per bullet"
  line in favor of the user's wording.
## 0.9.125
### Changed
- Section-title colors extended across the other Settings panes: the
  Approval policy sheet title and proper "Connections" / "Profiles" sheet
  headings (new, above the edit forms) all use the same stand-out style -
  deep teal on light, bright cyan on dark. The dark override is
  specificity-matched against the generic h2 gray so it wins in both
  themes (verified in the simulator).
## 0.9.124
### Changed
- Settings section titles (GENERAL (ALL ADD-INS), CUSTOM SYSTEM PROMPT,
  etc.) now stand out: bright cyan on dark theme, deep teal on light,
  uppercase with extra letter spacing - applied to every settings group
  in all three add-ins.
- Custom-prompt field: the prefill now lives where the config arrives, so
  the built-in default reliably shows on first open (previously it could
  stay empty until the next config render). While the user is typing,
  config reloads no longer clobber the field.
## 0.9.123
### Changed
- Custom system prompt auto-save reverted to a Save button (auto-save =
  only right for checkbox-type settings). The button writes the textarea
  verbatim with the green flash.
### Added
- Built-in quality/style prompt: when no custom prompt is set, the field
  is presented pre-filled with a default covering two labeled blocks -
  "### Content (accuracy)" (no invented facts, flag uncertainty, separate
  existing content from suggestions) and "### Writing style" (plain
  vocabulary without dumbing down, active voice, one idea per bullet).
  Clearing the field and saving falls back to this built-in prompt; the
  fallback also applies to the actual prompt assembly for fresh installs
  (customPromptDefault served by the server).
## 0.9.122
### Added
- Custom system prompt (Settings -> General, all add-ins): a textarea whose
  content is appended to every AI request's system prompt ("USER CUSTOM
  INSTRUCTIONS" block) - writing style, tone, formatting house rules.
  Auto-saves 1 s after the last keystroke with the green flash. Stored in
  the shared config (customPrompt, empty by default, so new installs get
  the key too; 4000-char cap).
## 0.9.121
### Changed
- Manual-notes fallback (when a host refuses notes-page writes): each note
  now sits next to a clipboard copy button - one click copies the exact
  text for that slide (button flashes "\u2713 copied"). All three manual-
  notes render sites use one shared builder; the box header names the
  cause (this Office build blocks notes-page edits). Notes-writing on
  Windows will get a deeper API fix once the affected host build can be
  probed directly; the sim shim writes notes successfully (the manual
  path only triggers where the host genuinely refuses).
## 0.9.120
### Changed
- After a successful "Pull update from git now", the success message tells
  the user to refresh the pane (About -> Reload pane / close+reopen) to
  load the new pane files first, then run the installer to redeploy.
## 0.9.119
### Changed
- Explicit speaker-notes requests no longer refuse while "Generate speaker
  notes" is OFF. set_notes ops now survive validation: they render as
  "(needs approval)" rows with an explanatory line ("note generation is
  OFF in Settings; ticking this writes them anyway for this slide"), so
  asking once confirms once. The model rule was rewritten accordingly:
  never refuse an explicit notes request, never demand a Settings change
  first, never offer to put the note on the slide instead; only spontaneous
  note-writing stays suppressed.
## 0.9.118
### Changed
- Settings -> General (all add-ins): settings save automatically on
  change - no more per-group Save buttons. Changed: debug logging
  (checkbox), speaker-notes generation, session auto-clean (checkbox +
  days, debounced 500 ms), update-check-at-startup (was already live).
  A small green "\u2713 saved" flash confirms each save next to the
  control; errors surface in the same spot.
## 0.9.117
### Docs
- Documented the blank icon in the Add-ins dialog as an expected cosmetic
  artifact of the http-only localhost transport (dialog webview skips
  http images; the icons themselves are served and used elsewhere).
  Decision: keep http-only, no external icon hosting.
## 0.9.116
### Docs
- Removed every how-to ribbon reference from the docs. The only remaining
  mentions are the warning: ribbon-commands manifests break the add-in on
  the tested builds (Office-Mac 16.114.3, Windows 16.0.19725), so the
  installer ships list-mode by default and --ribbon installs are not
  supported. The stale "Insert -> My Add-ins -> SHARED FOLDER" Windows
  route is gone too - every doc now points at Home -> Add-ins ->
  Advanced -> Shared Folder.
## 0.9.115
### Docs
- README: new "The local server" section - manual start (--start per OS,
  --status/--stop, menu 2b) and automatic start (macOS login item,
  Windows hardened logon task + schtasks quick restart), plus the
  health URL and the pane's red-footer hint.
## 0.9.114
### Fixed
- git calls made by the Windows server (update probe, git pull) no longer
  flash a console window per call (CREATE_NO_WINDOW).
## 0.9.113
### Added
- When the server is down, the pane footer shows an OS/setup-specific
  restart hint under the red status text: Windows = "schtasks /run /tn
  OfficeKickServer (or in your clone: py -3 install.py --start)", macOS =
  "python3 install.py --start (in your OfficeKick clone)". It hides again
  the moment the server answers.
## 0.9.112
### Added
- install.py --start / --stop / --status: bring the local server up (or
  down) without touching the install; the interactive menu gained
  "(2b) Start (or restart) the local server only".
- Windows: the logon task no longer self-kills the server (the
  stop-on-idle/battery/72h-limit settings used to end it silently - the
  likely cause of "server died"); the task is re-hardened at every install.
## 0.9.111
### Changed
- The "enable debug logging" explanation shortened to two lines: the log
  folder + the fact that the logs contain slide/document/message contents.
## 0.9.110
### Added
- Local-clone update checks: the installer records the checkout folder
  (repoPath + origin URL) into the config; the server then answers
  "check for updates" from the clone itself (git fetch + compare HEAD vs
  origin + read the repo's VERSION), no GitHub call needed. Without a
  recorded clone or without git, it falls back to the GitHub raw VERSION
  file. When the local clone shows a newer origin, the pane shows a
  "Pull update from git now" button (POST /api/repo-update runs
  git pull --ff-only; nothing pulls automatically).
- Windows update checks on this VM used to fail (no outbound result);
  with the local clone they work again.
## 0.9.109
### Fixed
- Behaviour group box: the remaining dead tail under the checkbox removed
  (empty feedback rows vanish, last row drops its bottom margin, group
  padding tightened) - 109px -> 69px, verified by measurement + screenshot.
## 0.9.108
### Changed
- REVERTS 0.9.107 on Windows: ribbon commands are opt-in again
  (--ribbon, both platforms default list-mode). Proven on PowerPoint
  16.0.19725: with a commands manifest the add-in disappears from the
  new Apps dialog AND the classic Shared Folder dialog; list-mode keeps
  the working Add-ins -> Advanced -> Shared Folder flow.
### Fixed
- Settings -> Approval policy: the Behaviour group box no longer renders
  a tall empty area under "Auto apply approved actions" (empty feedback
  rows collapse; group snugs to its content; checkbox row no longer
  stretches full width).
- Pane text size options now run 75% through 130%.
## 0.9.107
### Changed
- Windows installs include the ribbon commands by default now
  (--no-ribbon opts out; --ribbon force-includes). macOS stays list-mode
  by default (Office-Mac renders nothing from commands manifests) with
  --ribbon still available as a forced opt-in. Docs updated to point at
  the ribbon button first and the Shared Folder dialog as fallback.
## 0.9.106
### Changed
- "Update available" notice is printed bold and colored: gold in the dark
  theme, purple in the light theme (all three add-ins, manual check and
  the startup check).
## 0.9.105
### Fixed
- Pane reload now resumes the active session reliably (all three
  add-ins): the resume id moved from sessionStorage to localStorage -
  sessionStorage does not survive a pane reload in the Office WebViews,
  which is why a refresh used to drop back to a brand-new session. Keys
  are app-qualified (sk-resume-session-<app>) so panes never inherit
  another add-in's session, and deleting the current session also clears
  its resume id (Word/MailKick). The boot resume (id or newest session)
  happens before the first render.
## 0.9.104
### Added
- The pane footer text (next to the heartbeat dot) now shows server
  reachability too: green while the local server answers, red and gently
  pulsing while it does not (both themes), verified live through a
  server stop/start cycle on the same pane.
## 0.9.103
### Fixed
- clear_slide on Windows: the host accepted the text-wipe writes but
  silently dropped same-batch empty-string writes - slides reported
  "cleared" while every frame kept its text. The clear is now issued per
  shape with a sync after each write (Windows drops same-batch empty
  writes), deleted-lines handled in the first pass, and the success line
  reports any frames that had to be skipped.
## 0.9.102
### Fixed
- Windows session persistence: saves could die with "Permission denied"
  on the sessions dir when its ACL was restrictive/non-inherited
  (ssh-created dirs running under the logon task token). The server now
  skips chmod on Windows and the installer grants the current user full
  control over the config dir tree after deploy (icacls /grant, per-user
  data only). Verified: 4/4 live-server saves succeed after the ACL fix.
## 0.9.101
### Changed
- Group boxes: tighter top spacing so the group header (e.g. "PROFILE /
  ROUTE") sits close above the first control row (gap 16px -> 5px, all
  three add-ins).
## 0.9.100
### Changed
- Sessions tab (all three add-ins): while no connection or no profile
  exists, the list is replaced by guidance ("Create a connection in
  Settings first (Connections tab), then a profile in Settings (Profiles
  tab)") and every interactive element is gone until a profile exists.
## 0.9.99
### Added
- Pane text size in Settings -> General (Display group): 85% / 100% /
  115% / 130% - scales the whole pane immediately (zoom), saved in the
  shared config so all three add-ins follow it (server: paneFontSize,
  clamped 80-160).
### Docs
- README + slidekick/dockick docs now document the verified Windows path:
  Home -> Add-ins -> Advanced -> the classic Office Add-ins dialog ->
  SHARED FOLDER; the new "Apps" dialog (Manage) lists store/organization
  add-ins only.
## 0.9.98
### Changed
- Proposal gate (all three add-ins): executed actions leave the "Proposed
  actions" list into "Completed actions"; when every action finished
  successfully the proposal box closes and its buttons (Apply / Select
  all / Discard) and the "don't ask again" checkboxes disappear with it.
  Failed actions keep their row (marked failed), the counts and the
  re-armed Apply only track the not-yet-done actions - Apply retries
  exactly the survivors. Actions removed with the X now leave the list
  properly (they counted as phantom leftovers before and kept the box
  open on success). Removed the dead window.confirm before applying with
  unapproved rows (it never fired in the Office WebView and silently
  blocked Apply): the approved actions run, the skipped ones are listed
  in the results with a hint. Word and MailKick gain the same progressive
  gate (completed box, live counts, survivor retry) they lacked.
## 0.9.97
### Fixed
- clear_slide: the "items/textFrame" deep-load was rejected with
  InvalidArgument on Windows hosts (16.0.19x) - the only non-scalar load
  string in the pane. Rewritten as scalar loads + per-shape lazy
  textFrame/textRange loads + guarded line deletion; identical behavior
  on macOS.
### Changed
- install.py on Windows now also registers each add-in in the
  HKCU ...\Wef\Developer dev-sideload key (default value = manifest
  path, RefreshAddins=1): current Windows "Apps" dialog builds ignore
  Trust Center folder catalogs, and every Office app reloads Developer
  entries at launch. Outlook still needs Add from File once.
## 0.9.96
### Changed
- Renamed tools/repair_deck_paneref.py to tools/repair_paneref.py now that
  it covers Word documents as well as PowerPoint decks; docs and the
  installer tip updated to the new name (no behavior change).
## 0.9.95
### Changed
- tools/repair_deck_paneref.py is universal: accepts Word documents
  (.docx) alongside PowerPoint decks (.pptx) - same stamp, same orig/
  timestamped-backup + in-place repair. Folder scans pick up both kinds;
  other extensions still reported and ignored.
### Docs
- docs/dockick.md: Word troubleshooting entry for the "Add-in Error" toast
  with the tool usage; slidekick.md scope line updated to .pptx/.docx.
## 0.9.94
### Docs
- docs/slidekick.md: full troubleshooting section for the "Add-in Error"
  toast - cause (host-stamped deck reference), the repair tool with all
  its modes (--dry-run, folders, -r), the orig/ backup scheme, idempotency
  and the verified reopen behavior (Finder double-click auto-loads the
  pane; in-app open dialog skips restore silently).
- README.md Update section: quit-Office-before-install guidance and a
  pointer to the toast healing docs.
## 0.9.93
### Changed
- tools/repair_deck_paneref.py: --folder/--file ergonomics. Accepts any mix
  of .pptx files and folders; folders patch every .pptx inside (non-default
  recursion with -r / --recursive). Originals are copied to an orig/
  subfolder next to the deck(s) as <name>.pptx.<YYYYMMDD_HHMMSS>.bak, then
  the deck is patched in place keeping the same filename (atomic
  write-temp + os.replace). --dry-run touches nothing (no orig/ created).
  PowerPoint lock files (~$*.pptx) are reported as locked and skipped;
  non-deck files in a folder scan are listed and ignored; backups are
  written only on the first change (idempotent re-runs add no new .bak).
- Installer mac tip now shows the folder form.
## 0.9.92
### Changed
- Update protocol documented as verified end-to-end: repair a deck once
  (tools/repair_deck_paneref.py) -> save anytime with the pane open ->
  reopen by double-clicking the deck in Finder -> the pane auto-loads and
  no toast appears. The repaired pane reference survives pane-open saves;
  PowerPoint re-stamps the unresolvable Registry form only on decks that
  still carry it. Opening through the in-app open dialog skips pane
  restore silently (no error) - Finder double-click is the verified path.
  Installer's mac tip now states the verified behavior.
## 0.9.91
### Added
- Deck pane-reference repair tool (tools/repair_deck_paneref.py): removes
  the store="developer" storeType="Registry" stamp PowerPoint writes into
  decks saved with the pane open. On the Mac host (16.114) that stamp
  cannot be resolved at launch-hydration, so reopening any such deck
  throws the "Add-in Error: New Office Add-in" toast; without the stamp
  the host resolves the pane purely by add-in id (the path the manual
  ribbon open always used). Decks open cleanly afterwards; the pane does
  not auto-open (open it from the ribbon - one click, always works).
  Writes a .bak backup next to each deck, idempotent, --dry-run support.
  The installer prints the reminder + tool path after the sideload step.
  (Evidence: both a stale deck and a fresh deck stamped identically -
  same id, same 1.0.4.0, same broken store form - and the host build
  confirmed unchanged across the regression window.)
## 0.9.90
### Added
- Version transparency on both surfaces:
  1. The installer now prints, per add-in, the manifest identity version
     it renders + sideloads ("slidekick (id v1.0.4.0)") and the pane
     build in the manifest-written line, so update cycles log the exact
     identities in flight.
  2. The About tab in all three panes now lists every version number:
     Pane build (JS, window.__SK_BUILD), the pane's own manifest
     identity version, ALL deployed manifest identities, the Office
     host + version (Office.context.diagnostics), plus the server build
     it already showed. /api/config carries the manifest versions
     (deployed manifest dir) for this.
## 0.9.89
### Fixed
- Restore save-anytime + auto-reopen (user requirement: the pane must
  re-hydrate on deck open, as before 0.9.80):
  1. Manifest identity REVERTED to 1.0.4.0 - the 1.0.5.0 bump (0.9.80,
     a diagnostic-only change) is the prime suspect for the saved-pane
     hydration regression: decks saved under 1.0.4.0 fail to re-hydrate
     after a bump, while the Ribbon path tolerates it. Frozen permanently.
  2. The 0.9.88 boot-time AutoShowTaskpaneWithDocument removal is gone -
     it fought auto-reopen instead of supporting it.
  Install note: quit Office completely, run install.py --install (the
  WEF guard requires it), reopen.
## 0.9.88
### Fixed
- The deck-open "This add-in is no longer available" toast proven
  document-local (fresh deck: no toast; the long-lived test deck: toast,
  across reinstalls): decks saved with the pane open remember the
  task-pane reference, and after manifest churn the host fails to resolve
  it. All three panes now clear the auto-open marker
  (Office.AutoShowTaskpaneWithDocument) on boot: decks stop remembering
  the pane (it is opened via the ribbon, which always resolves), and
  opening the affected deck once with the add-in active and saving again
  repairs it permanently. Mail no-ops (no document settings there).
## 0.9.87
### Fixed
- The recurring "Add-in Error: This add-in is no longer available" on deck
  open is a host-side stale registration: install.py purged the WEF folder
  and re-sideloaded while Office was running (or the 1.0.4.0 -> 1.0.5.0
  version bump orphaned decks' pinned pane registrations - PowerPoint Mac
  re-opens task panes a document had open, and a missing/changed identity
  throws exactly that toast). Server panes were never implicated: every
  logged boot is healthy.
  1. install.py now refuses to purge/sideload WEF manifests while
     PowerPoint/Word/Outlook is running and tells you to quit Office
     first, so no update can ever create a dangling registration again.
  2. The manifest version is policy-locked at 1.0.5.0 (comment in the
     templates): the (Id, Version) pair is the host's registration
     identity - it stays stable so remembered panes always resolve.
## 0.9.86
### Fixed
- Cancel-button hardening requested before the test rounds - a code audit
  found three gaps and closed them in all three panes:
  1. The Cancel button existed but was never shown: setRunEnabled did not
     toggle its visibility, so the html-hidden button stayed invisible
     during runs. It now appears + enables only while a run is in
     progress and hides when the run ends.
  2. Word/Mail apply loops lacked the per-op boundary check: cancel mid-
     apply now stops at the very next op ("cancel - remaining op(s) not
     executed", failed/retry pass skipped) instead of reeling through
     every remaining action.
  3. The Word make-over passes (director/build/review) now refuse to
     start the next pass once a cancel is pending.
  Unchanged and verified present everywhere: instant transport abort
  (AbortController on the in-flight request), abort mapped to "Run
  cancelled.", retry-before-each-attempt checks, arm/disarm of the flag.
## 0.9.85
### Fixed
- Word make-over hardening (from the log forensics of a real run):
  1. make-over pass merges dedupe set_header / set_footer /
     insert_paragraphs across passes - an eager builder that repeats an
     earlier extra no longer doubles it in the merged list.
  2. A disobedient director that answers with ops JSON instead of a
     strategy no longer propagates as "strategy": its ops are kept as
     seed results, and the builders receive a concise editorial fallback
     strategy to plan with.
## 0.9.84
### Fixed
- Boot diagnostics were silent: dbg() bails before /api/config delivers
  the debug flag, which swallowed exactly the lines we needed (pane.boot,
  theme). dbg() now queues pre-config entries (max 16) and replays them
  once logging is on - pane.boot and the theme decision are always
  captured in all three panes.
## 0.9.83
### Fixed
- CRITICAL: the 0.9.80 pane.boot beacon edit dropped "var params" from
  boot() in all three panes - every boot since 0.9.80 threw
  "params is not defined" the moment the pane JavaScript started: no API
  calls, no debug lines, no theme detection (pane stuck light in a dark
  Office), and the Add-in Error toast on pane open. This is the
  reproducible cause of the reported update->relaunch Add-in Error.
  params is restored; the beacon now fires on every healthy boot.
## 0.9.82
### Fixed
- The 0.9.81 media-query conversion produced malformed CSS (missing rule
  braces, an explanatory comment parsed as a selector). The three dark
  blocks are now hand-verified valid: body.dark-scoped subtabs, group
  legends, context chips and request lines.
## 0.9.81
### Fixed
- Theme detection desync: the subtab titles / group legends / context
  chips used @media (prefers-color-scheme: dark) blocks, which the Office
  WebView evaluates independently of the pane's own body.dark logic -
  OS-dark + pane-light produced lime-on-light (unreadable), the reverse
  dark green-on-dark. All such blocks now follow body.dark like every
  other theme rule, so dark/light can never disagree again.
- Light-mode titles are azure blue (#0078d4, active #005a9e) instead of
  the saturated green; lime stays in dark mode only, where it reads.
## 0.9.80
### Changed
- Add-in Error investigation build: every pane now emits a pane.boot
  debug line the moment its JavaScript boots (URL, user agent, Office.js
  presence) - when the toast fires, the log will show whether the pane
  ever got this far or died before loading. Manifests bumped to 1.0.5.0
  so the update cycle matches the reported repro (update -> quit ->
  relaunch).
## 0.9.79
### Fixed
- The run timer froze and Run re-enabled while auto-apply was still
  executing: renderParsed kicked off the auto-apply without returning it,
  so the run chain finished (timer stopped, "busy" cleared) while long
  applies kept running - very visible on orchestrated make-over applies.
  renderParsed now returns the apply chain (auto-apply only; a parked
  gate still ends the run immediately) and all three run chains wait for
  it before freezing the timer / clearing busy.
## 0.9.78
### Added
- AI-driven document make-over in DocKick (the Word counterpart of the
  SlideKick orchestrated redesign). A creative-director round invents a
  real design strategy for THIS document (CONCEPT/THEME/STRUCTURE/EXTRAS
  - fonts, exact sizes, 2-3 colors + accent as hex, hierarchy rules,
  header text), then three builder passes turn it into ops with enforced
  op types per pass (1 structure: heading/lead-in/quote hierarchy, 2
  style: typography and colors for EVERY content paragraph, 3 extras:
  header/footer, links, pull-quote), and a review round emits any format
  ops still missing before the normal validation + gate take over.
  Merged ops pass the phantom guard (only real 1-based paragraphs) and
  the whole make-over falls back to the normal single ask if any pass
  fails.
## 0.9.77
### Fixed
- DocKick real-Word crash wave on a redesign ("Applied with 10 error(s)"):
  1. format_text InvalidArgument - Word.js uses the "Centered" alignment
     token, not PowerPoint's "Center"; both alignment maps (paragraph and
     selection formatting) now use the Word enum, so align=center ops
     apply instead of failing (and burning the retry pass).
  2. set_header "sections.getItemAt is not a function" - Word.js
     collections have no getItemAt; header/footer now loads the sections
     collection (load + sync + .items[0]), matching the documented
     collection contract used elsewhere in the pane.
## 0.9.76
### Fixed
- Pair delete in "Previous requests" never asked and never deleted in the
  real Office pane: the wastebasket row relied on window.confirm, which
  the Office WebView silently swallows (it only worked in the simulator).
- The delete is now the SAME two-stage button as the sessions list: an
  expanded step shows a small "Delete" button; the first click arms it to
  "Really delete?" for 3.5 s (stays deletable only then), the second click
  removes the pair everywhere (pane list, future model context, persisted
  session). The bin icon row is gone. All three add-ins, no dialog needed.
## 0.9.75
### Added
- Living session titles: as a session continues, every save now runs one
  cheap background AI call that reviews the current title against the
  session's recent requests and either replies with a new 2-4 word title
  (what the session is NOW about, which can drift far from the opening
  ask) or exactly KEEP. New titles are junk-filtered like first names,
  persisted via the rename endpoint and reflected live in the pane label
  and the sessions list; a user-set title stays in the loop the same way
  (the model always sees the current one). One update in flight at a
  time (titleBusy guard) in all three add-ins.
 ## 0.9.74
### Fixed
- Web lookups no longer stall on "I tried to run the live web lookup, but
  the lookup channel isn't available..." - three changes, in all add-ins:
  1. The pane now OWNS the lookup: when a reply is prose that says the
     model can't verify/search and the request wants live data (rankings,
     prices, news, dates...), the pane runs the web lookup itself (instead
     of waiting for the model to volunteer {"kind":"web_lookup"}) and
     re-asks once with the fresh results.
  2. Lookup capability ladder: the profile's model first (OpenRouter
     gets the :online web plugin), then ANY OpenRouter connection with
     "openai/gpt-4o-mini:online" as a guaranteed web-enabled fallback -
     so a no-web profile model (e.g. through a plain generic-openai or
     cheaper-inference route) can still serve a live lookup if an
     OpenRouter connection exists. A failed attempt is skipped, not fatal.
  3. The system prompt stops the hallucination at the source: the model
     is told it has no web access of its own, the channel belongs to the
     pane, and it must never claim it tried a lookup or that the channel
     is unavailable - either {"kind":"web_lookup"} or answer.
## 0.9.73
### Added
- A Cancel button next to Run (all three add-ins), visible and enabled only
  while a run is in progress. Clicking it interrupts the run as fast as
  possible: it aborts the in-flight request immediately (AbortController -
  the fetch dies the same moment) and the run stops at the very next step
  boundary - ask retries, make-over passes/chunks/top-up/review rounds,
  image-job poll cycles, and between apply ops (the remaining ops report
  "cancel - remaining op(s) not executed" and stay untouched). Aborted
  shows as "Run cancelled." with no error styling on the input, so the user
  can fire a new request right away. Cancelled runs are not stored as
  completed exchanges.
## 0.9.72
### Added
- Delete a request/response pair from "Previous requests": expanding a
  step reveals a wastebasket row; clicking it asks for confirmation and
  then removes the whole exchange pair everywhere at once - the pane's
  session history (so future prompts no longer carry it), the shown list,
  and the persisted server session (new PUT /api/sessions/remove-pair/<id>
  endpoint, which deliberately bypasses the keep-longer upsert rule so the
  deletion survives reloads). Cancel keeps everything. In all three
  add-ins.
## 0.9.71
### Fixed
- move_slide / duplicate_slide / duplicate_slide_to in real PowerPoint:
  exportAsBase64Presentation returns OfficeExtension.ClientResult (the
  value is read as .value after a ctx.sync()), not a promise - our code
  chained .then() on it and died with "...then is not a function" before
  touching the deck. All three paths now follow the documented contract
  (verified against Microsoft's API reference). The 0.9.54 retry pass even
  retried the broken op and reported it twice - the op itself was the bug.
- create_slide now honors "after": the pane documents "after":N in the
  prompt (N = 1-based slide number to place the new slide after; 0 = very
  start) and MOVES the fresh slide into that position automatically right
  after creating it (with a fallback note if the move fails). Previously
  the prompt even told the model a two-step create-then-move workaround
  that models skipped - inserting "somewhere" silently meant "at the end".
## 0.9.70
### Changed
- The Profile / route refresh glyph is bigger again (17px -> 20px). The
  Profile and Route selectors sit side by side as before (the refresh
  affordance lives only in the legend title).
## 0.9.69
### Fixed
- The status bar no longer stays on "Applying N op(s)…" when the gate closes
  itself after a fully successful apply: that path skipped the final status
  update, so the bar kept the apply message. It now reads "Applied N op(s)."
  (green) - or "Applied with N error(s)." when some failed.
## 0.9.68
### Changed
- The Profile / route refresh glyph is bigger (13px -> 17px).
## 0.9.67
### Changed
- The config refresh affordance moved from a small button next to Route
  into the "Profile / route" legend title itself (all three add-ins): a
  clear 13px bold refresh glyph, black in light mode / white in dark mode,
  hover scale-up; clicking it re-reads connections and profiles from the
  config files.
## 0.9.66
### Changed
- Saving a connection clears the form immediately (back to "New connection"
  defaults); the success message stays so you still see what was saved. The
  form is ready for the next connection without manual cleanup. Applied in
  all three add-ins.
## 0.9.65
### Added
- A small refresh button (circular arrows) next to the Profile/Route
  selector in all three add-ins: re-reads connections and profiles from the
  config files (loadConfigAndRender) without reloading the pane - useful
  after editing ~/.officekick/config.json by hand.
### Changed
- Asset cache-busting is now <version>-<mtime-ms> instead of version-only:
  fixes within one release no longer serve the stale cached pane scripts
  (this exact trap hit testing the refresh button itself).
## 0.9.64
### Fixed
- DocKick: "undefined is not an object (evaluating 'res.lines')" after a
  successful auto-apply. The apply result chain was mis-wired: the
  change-tracking restore ran as a detached promise and the return handed
  back the (empty) retry promise's own resolution instead of the result
  object whenever no action needed a retry. The chain is single again and
  always returns the result. MailKick's equivalent path was already
  correct.
## 0.9.63
### Fixed
- ART pass: set_shape_image ops no longer reference hallucinated shape
  indices ("shape 9 not in the context list" x N). Before the art round the
  pane re-reads the LIVE deck and hands the model the real per-slide shape
  inventory (1-based indices, sizes included); the prompt forbids guessed
  indices and points the model at the structure pass's roundedRectangle /
  ellipse placeholders as the cut-out spots.
- Repaired a broken build introduced during an interrupted edit (unterminated
  string in the make-over pass note + a mangled ART pass block); syntax and
  suite verified again.
## 0.9.62
### Added
- DocKick and MailKick get the same "Reload pane" button as SlideKick
  (Settings > General, next to Check now): cache-busting hard reload from
  the running server after a back-end update, plus the active-session
  auto-resume markers - a reload resumes the session that was open, and a
  new chat or session deletion clears it. All three add-ins now share the
  same reload/resume behavior and the same common code shape.
## 0.9.61
### Changed
- Gate lifecycle: when every action has been executed successfully the whole
  "Proposed actions" section (rows + Select all/Apply/Discard) closes
  itself; only the "Completed actions" list remains below where the buttons
  were. Failed actions stay in the proposal list and the Apply button
  re-arms after the pass so the failed ones can be retried with one click.
  Actions rescued by the automatic retry pass now also move to the
  completed list (they used to stay dimmed in the proposal list).
- The completed list is rendered as a sibling of the proposal section, so
  closing the section no longer removes it.
## 0.9.60
### Changed
- Image generation no longer holds one WebView fetch open for the whole
  provider wait - that is what produced the batch of "Cannot reach the
  OfficeKick server (/api/image): Load failed" lines: the server actually
  completed several of those generations (HTTP 200 in the access log) while
  the pane's long-lived fetch had already been killed.
  Server: new POST /api/image/start runs the provider call in a background
  thread and returns a jobId; GET /api/image/status/<id> hands back the
  result (job map is pruned by TTL, done results are consumed once).
  Pane: the three image ops (set_background picture / set_shape_image /
  create_image) now start the job and poll every 4 s with short requests -
  a dropped poll costs one cycle, not the whole generation. Transport
  hiccups during polling are tolerated; the 5-minute budget stays. The
  fail-once retry pass from 0.9.54 remains the last line of defense.
## 0.9.59
### Added
- Progressive action tracking in all apply paths: as each action executes it
  disappears from the proposal list; a "Completed actions (n)" list grows
  one check-mark at a time (newest on top). Actions that fail stay in the
  proposal list and are dimmed until the retry pass resolves them. The gate
  ends empty when every action succeeded.
- Readability is now enforced mechanically, not just prompted: the gate
  scans every planned op against the slide's planned background solid -
  dark text planned on a dark surface is flipped to near-white, light text
  on a light surface to near-black, with a "Readability guard: N text
  color(s) automatically adjusted" note in the gate. Photo backgrounds are
  excluded (their overlays govern contrast).
## 0.9.58
### Changed
- The make-over orchestration is now a multi-pass, self-driving pipeline.
  After the design-director round: PASS 1 (STRUCTURE, chunked over the deck)
  lays out backgrounds, panels, bars and repositioning with per-slide
  completeness forcing top-up rounds; PASS 2 (POLISH, whole deck) applies the
  concept's typography and readability; PASS 3 (ART, whole deck) adds the
  concept's imagery (still approval-gated since image generation bills);
  PASS 4 (REVIEW) makes the model audit its own redesign against the concept
  and the readability law and fix specific slides, looping until it replies
  "done" (max 2 review rounds). Every pass builds on the accumulated op list
  ("OPS SO FAR - do not repeat"), and the final single approval gate carries
  the director's concept plus a "built in AI-driven passes" note. All passes
  merge in order: 64 verified ops across a 15-slide E2E deck, every slide
  covered structurally, polished, and where the concept asked, imaged.
## 0.9.57
### Added
- AI-talks-to-AI make-over orchestration: a whole-deck redesign now runs in
  two stages. First a DESIGN DIRECTOR round invents the concept (style name,
  palette + text-color pairing, type scale, image language, a distinct
  composition idea per role) with ops forbidden. Then builder rounds
  translate the concept into concrete ops - chunk by chunk (4 slides at a
  time), so a 13-slide deck can never silently stop at slide 9. Slides that
  came back without design ops get one automatic top-up round. Everything
  is merged into a single approval gate, headed by the director's concept.
  Requests that name specific slides keep the single-pass path.
### Fixed
- Orchestration merge bugs found by E2E: chunk ops were doubled for fully
  covered ranges and the concept note never surfaced (local vs stored
  variable). The plain parse chain no longer races the orchestrated flow.
## 0.9.56
### Changed
- Readability rules hardened in the design prompt and the make-over brief.
  Present on EVERY request: pick text color from the surface behind it (dark
  text on light/plain/panel fields, near-white only on deep/overlaid/strongly
  contrasting fields), body >= 16pt on colored fields (>= 14pt anywhere,
  never below 12pt), titles >= 28pt with an obvious title/body gap, no body
  text directly on busy photos, no mid-tone text on mid-tone fills. In the
  make-over brief, the quality bar adds titles 28-44 / body 16-20 /
  captions >= 14 and a final per-slide contrast re-check.
## 0.9.55
### Changed
- Past-request excerpts are clickable: clicking the request snippet line
  expands/collapses the step exactly like the +/- and timestamp do.
- The copy icon on those steps is larger (15px) and clearly visible:
  black in light mode, white in dark mode (theme-aware), with a hover
  scale-up.
## 0.9.54
### Changed
- The apply flow now does its best to complete every step: after the main
  pass, any action that failed (and any actions skipped by an unexpected
  error) is retried once, in the original order, after a short pause. Each
  retry is reported ("retrying N uncompleted action(s) once..."), successes
  appear normally, and ops that fail a second time say so explicitly
  instead of vanishing. Applies to slide/word/mail add-ins.
## 0.9.53
### Added
- Past-request steps get a small copy icon (two joined squares) on their
  title line, and clicking the request line now always copies the request
  text to the clipboard as well; the input box is only refilled when empty
  (so a prompt you are building is never overwritten). Works in all three
  add-ins.
## 0.9.52
### Added
- Truncated-reply recovery in all three add-ins: when a model reply is cut
  off mid op-list, the pane salvages every complete op, automatically asks
  the model once for the remaining ops ("only what is still missing"), and
  merges both into one approval gate, marked "Recovered N action(s) from a
  cut-off reply". No more raw JSON shown as chat text or silently dropped
  redesigns. Raw unparseable replies are shown capped at 420 chars with a
  pointer to debug logging.
- Stability pipeline parity: word/mail panes now have the same escalated
  repair ladder as SlideKick (2 rounds, explicit "an empty response is
  NEVER acceptable" nudge for blank replies), the same salvage + continue
  merge, and the same no-raw-dump display policy.
### Changed
- Default max_tokens for chat raised 4000 -> 8192 (server-wide); full-deck
  redesigns no longer exhaust the cap and get cut mid-JSON as easily.
### Fixed
- Salvage order: a salvaged op list previously returned early and skipped
  the continuation (only the first chunk reached the gate).
## 0.9.51
### Added
- "Reload pane" button next to "Check now" (Settings > General): hard-reloads
  the task pane from the running server with a cache-buster, so a back-end
  update lands without reopening PowerPoint. The active session is remembered
  (in-session storage) and automatically resumed after the reload; starting a
  new chat or deleting the session clears it.
## 0.9.50
### Changed
- Empty model replies are no longer accepted: the pane escalates with up to
  two firm re-asks, and a blank response triggers an explicit "an empty
  response is NEVER acceptable" nudge. A run can only end after a usable
  JSON or a real text answer.
- Redesign guardrails in the prompt (always-on rules + make-over brief):
  never remove or overwrite user content (pictures/charts/tables stay -
  redesign around them), clear_slide only on nearly-empty placeholders,
  duplicate_slide_to never aimed at slides still carrying user content, and
  a pass must cover every in-scope slide (phase 1 + immediate phase 2 for
  very long op lists - never stop halfway).
- The make-over brief now names the full toolbox explicitly: op-order
  layering, transparency stacking, rotation diagonals, gradient washes,
  oversized numbers/keywords, accent bars, cut-outs in circles/rounded/
  hexagons - combinations make the design feel crafted.
### Fixed
- clear_slide deleted every line shape on the slide, including the user's
  own lines. Created lines are now tagged (sk-shape-*) and only tagged
  shapes can be removed; user lines survive a clear. Shapes created by
  create_shape are tagged at creation.
## 0.9.49
### Fixed
- create_shape accepted geometry only in flat form (left/top/...); when a
  model nests them in a geometry object the values were silently dropped and
  shapes landed at defaults. The validator now accepts both forms, plus the
  rotation key for create_shape/move_shape in the live validate branch.
- Simulator context now reports shape rotation (and move/create store it),
  so make-over verification in the simulator sees rotated elements.
## 0.9.48
### Changed
- Shape rotation joins the make-over vocabulary: create_shape and move_shape
  accept an optional rotation (degrees, -360..360; positive = clockwise).
  Applied via Shape.rotation (PowerPointApi 1.10, gated with the same
  friendly error as the other 1.10-only ops); simulation parity included.
  The prompt documents rotation as the primitive for diagonal/tilted
  compositions.
### Note
- Decision recorded: single Office.js codebase (no VBA/COM/AppleScript
  split) - creative make-overs are built on the existing op vocabulary
  (shapes, picture-filled shapes, cut-outs, positions, transparency,
  rotation).
## 0.9.47
### Changed
- Make-over brief v3 - an invention engine instead of fixed recipes. The
  model now invents a named design concept with three rules (color trio,
  spatial system, shape + image language), must pick a contrasting direction
  when an earlier redesign is visible in the session, and is told to
  re-STRUCTURE the deck first: add a real cover, insert topic dividers,
  clear/rebuild placeholder slides, duplicate a good slide over a bad one,
  and reorder with move_slide. Composition is free per slide (no canned
  full-bleed-photo covers); images are used where the concept calls for
  them (a photographic concept without generated images breaks the concept,
  a graphic concept may stand on shapes). Hardcoded cover/panel recipes from
  the previous brief removed.
## 0.9.46
### Changed
- The make-over brief rewritten for zero-guidance redesign requests: it now
  states "FULL REDESIGN, not decoration" and that a slide receiving only a
  font/color tweak makes the job incomplete; it fixes the theme (style +
  3-color palette + type + one shared image style) before any op; it prescribes
  concrete per-role blueprints (full-bleed photo + overlay + anchored large
  title + accent bar for covers, split panel with an in-shape cut-out photo +
  oversized number for content, restrained statement/table/image patterns);
  and it finishes with hard completeness rules - every in-scope slide takes a
  structural op, moves on most slides, at least 3 generated images for a
  full-deck pass (zero images with a photographic theme = failure), and a
  readability law (light text only on overlays/deep/panel colors, no short
  text on bare photos, nothing under 14pt on colored fields).
## 0.9.45
### Changed
- Partial-apply dimming refined: the Apply button dims while any confirm
  row is unticked (0/Z, 1/Z, ...), un-dims at full approval and when the
  selection turns fully automatic. The dimming lives in the apply label
  update, so it tracks every live selection change.
## 0.9.44
### Changed
- **Confirm before a partial Apply**: clicking "Apply X auto + Y/Z
  approved" when some confirm rows are still unticked now asks first
  ("N (needs approval) action(s) are NOT approved and will be skipped: ...
  Apply the other X anyway?"). Cancel = nothing applies. Previously the
  unticked actions silently vanished. All three add-ins; the Apply button
  also dims slightly while the selection is a partial one.
## 0.9.43
### Added
- **Make-over mode as a skill**: when the request asks for a design
  make-over (makeover / make over / redesign / new design, theme or look /
  modern / professional / restyle), SlideKick injects a full
  creative-director brief into the system prompt: pick a concrete theme and
  palette first, an image language shared by every generated picture, one
  layout blueprint per slide role, and hard creativity minimums (a
  structural element on every slide, repositioning on most, 2-5 generated
  images per full-deck pass with at least one cut-out in a shape). A
  reply that only changes colors/sizes is explicitly named a failure to
  the model. Normal edit requests are unaffected (brief absent,
  verified).
## 0.9.42
### Added
- Automatic **log cleanup**: `debug-*.log` / `access-*.log` files older than
  30 days are deleted. Runs once at server boot and together with the
  session auto-clean tick (max once/hour when panes open). The window is a
  setting (`logCleanDays`, 0 = never delete).
## 0.9.41
### Fixed
- The "Generate speaker notes" checkbox previously only saved through the
  "Save policy" button on the Approvals sub-tab - there was no visible Save
  for it where it sits (General). It now has its own "Save notes setting"
  button next to the checkbox.
## 0.9.40
### Changed
- **Loopback fallback for local AI servers**: when a request to a numeric
  private-LAN host fails with Errno 65 (the request never left the box -
  macOS Local Network wall or a silent firewall drop), the server retries
  once against 127.0.0.1 on the same port/path. Local Ollama/llama.cpp
  setups that listen on loopback keep working even when the LAN route is
  blocked. Public hosts and DNS names are never rerouted; when the
  loopback retry also fails the error mentions both attempts.
### Fixed
- The reachability hint is only attached for Errno 65, not any URL error.
## 0.9.39
### Fixed
- Model replies containing commentary plus a valid ops JSON are now
  executed: the extractor scans every balanced {...} in the answer instead
  of only the first, prefers any object that declares kind/ops, and also
  accepts a bare ops array or a reply whose text field is not a string.
  Plain-text answers (no JSON at all) render as text as before.
- A reply that fails to parse no longer feeds its raw (possibly truncated)
  JSON into the conversation history - the model would keep repeating its
  own broken format next round; history now stores a parse-failure note.
## 0.9.38
### Fixed
- duplicate_slide / duplicate_slide_to / clear_slide had no apply cases:
  the gate showed them, Apply accepted them, nothing executed. They now
  run their deck methods and report per-op result lines like the others.
## 0.9.37
### Fixed
- duplicate_slide / duplicate_slide_to / clear_slide were missing from the
  op metadata table (the model could emit them and validation rejected them
  as "unknown op type"). Now registered properly: gate labels, validation,
  and the alone-in-batch rule cover all three.
## 0.9.36
### Added
- Slide lifecycle ops (SlideKick, PowerPointApi 1.10):
  - **duplicate_slide** - exact copy of one slide right after it (export +
    `insertSlidesFromBase64`, runs alone like move_slide).
  - **duplicate_slide_to** - copies a slide's full contents over another
    slide (target replaced in place). This is the "insert at a position"
    primitive: duplicate a good slide, then move it.
  - **clear_slide** - empties a slide: all text frames cleared, text-free
    lines removed; the slide, layout and background stay. Rebuild in the
    same reply with create_shape/move_shape.
- Simulator parity for all three; system-prompt explains the
  duplicate/move/clear workflow for inserting slides at a chosen position.
## 0.9.35
### Fixed
- Role detection: a slide with a title plus short body text no longer counts
  as a cover. Only slide 1 and slides without actual body content (cover /
  section dividers) get the "title" role; everything with bullet text is
  "bullets"/"statement" as before.
## 0.9.34
### Added
- **Role-aware make-overs (theme engine)**: the deck outline now tags every
  slide with a detected ROLE - title (cover/section), bullets, statement,
  table, image (picture/chart-led). The make-over prompt turns these into
  layout blueprints: every slide with the same role gets the same treatment
  (e.g. every title slide = background photo + dark overlay + large light
  title), different roles get different blueprints. Palette, type scale and
  spacing grid are fixed once per theme and applied deck-wide.
### Changed
- Slide outline lines now show `(role: ...)` per slide, so role-based design
  stays consistent across requests and deck sizes (compact and full outline).
## 0.9.33
### Changed
- The **access log** is no longer always-on: it is written only when debug
  logging is enabled (Settings > General > "Debug logging" or `--debug`), and
  then alongside the debug log in `~/.officekick/logs/access-YYYY-MM-DD.log`.
  One flag produces both; with debug off nothing extra is written.
## 0.9.32
### Added
- **Design make-overs** (SlideKick): a new design-system prompt layer teaches the
  model to build one coherent system (palette, type scale, spacing grid,
  consistent anchors) and vary each slide's composition instead of copying one
  layout deck-wide. Includes phased op ordering (backgrounds/palette/overlays ->
  layout moves -> images) and text-over-picture layering (picture first, then a
  45-60% dark overlay shape, then light text).
- New op **set_shape_image** (SlideKick): fills an existing shape with an
  AI-generated picture via `ShapeFill.setImage` (PowerPointApi 1.10), optional
  transparency. Confirm-gated like other image ops; needs an image model on the
  profile.
- `style_shape` and `create_shape` now accept **transparency** (0-1) beside
  fill/line colors - overlay masks/panels for text-over-image slides
  (`ShapeFill.transparency`, PowerPointApi 1.10).
- `style_shape` now also accepts text-less shapes (design rectangles/panels),
  matching `move_shape`.
- Simulator: picture-fill, per-shape fill/line/transparency parity, so make-over
  flows are testable outside Office (fillImage/b64len in sim state).
## 0.9.31
### Changed
- Approval-gate clarity: when a gate contains "(needs approval)" actions, it
  now states the mechanics in plain words - unticked actions are skipped on
  Apply (they are not failures, they were never approved); tick their
  checkboxes or use "Select all" to include them. All three add-ins.
## 0.9.30
### Added
- Always-on server **access log**: `~/.officekick/logs/access-YYYY-MM-DD.log`.
  Every request line is written with the requesting client's User-Agent
  (paths only, never bodies), so it shows after the fact who loaded which
  pane/page and when - including the Office webview during a launch where an
  "Add-in Error" appeared. Previously request lines were suppressed unless
  debug was on.
## 0.9.29
### Documented / Changed
- **Ribbon verdict (retest, clean state)**: with the Office cache stores wiped,
  manifest Version 1.0.3.0 and the server live, the tested Office-Mac build
  (16.114.3) still registers the commands manifest without rendering the AI
  group and without keeping the Add-ins list entry - the add-in shows an
  "Add-in Error" and disappears from the list. Confirmed on a CLEAN state this
  time: this is a build limitation for developer file-sideloaded commands, not
  a stale-cache artifact. `--ribbon` stays opt-in; **list-mode is the only
  working manifest mode on this build**. After any `--ribbon` attempt, restore
  with a plain install re-run (it also bumps the manifest Version to
  1.0.4.0 here, so Office discards the cached commands registration outright).
### Fixed
- After a failed ribbon attempt, re-running a plain install used the SAME
  manifest Version (cached instances are keyed by id+version) - Office could
  keep the broken commands registration. Every install pass that switches
  modes now ships a fresh manifest Version; 1.0.4.0 = current.

## 0.9.28
### Fixed
- install.py's "is Office running?" check matched Office's widget/extension
  processes - e.g. Outlook's CalendarWidgetExtension.appex, which macOS keeps
  running even with every app closed - so the add-in cache clear was always
  skipped with "Office apps are running". The check now matches only the main
  app binaries (.../<app>.app/Contents/MacOS/Microsoft <App>).

## 0.9.27
### Changed
- **install.sh and install.ps1 are removed.** install.py is the only installer
  (macOS + Windows, same commands and flags). Every reference was cleaned up:
  README (restructured: lists, short sentences), docs/architecture.md,
  docs/slidekick.md, docs/dockick.md, docs/mailkick.md, the pane's
  "cannot reach the server" hint (now: python3 install.py) and the server's
  update-notice text.
- README rewritten for readability: shorter sentences and more step lists
  (Install, Update, Use it, permissions, Uninstall).
### Fixed
- docs/slidekick.md dev-mode command documented for install.py-less dev runs:
  serve straight from a checkout with python3 core/server.py

## 0.9.26
### Fixed
- "Add-in Error" at Office startup after the https -> http transport switch:
  Office for Mac reopens the last-open task pane from its WebKit session/page
  cache, which still held the https-era entries, so the auto-opened pane died
  ("Add-in Error") while re-selecting the add-in from the Add-ins menu
  (a fresh http pane) worked. install.py now ports install.sh's
  clear_office_caches: while no Office app is running it clears the per-app
  WebKit/ctrlstr caches and the WEF registration store (both rebuild at
  launch), and the sideloaded manifest <Version> is bumped to 1.0.3.0 so any
  cached manifest copy from earlier installs is discarded outright. The clear
  is skipped (with a note) when Office is still running.

## 0.9.25
### Changed
- The server-heartbeat dot moved from the tab strip into the footer, right
  next to the status line where the pane reports its state (same behavior:
  green while the local server answers, red when it does not, 5 s poll,
  tooltip with the last check / version / fix hint). Same fix-hint tooltip.

## 0.9.24
### Added
- Server heartbeat in the task panes: a small dot at the right end of the tab
  strip (all three add-ins) that goes green while the local OfficeKick server
  answers its health endpoint and red when it does not (checked every 5 s).
  The tooltip shows the last check time (+ the running version) and, in the
  red state, the fix hint ("start it with: python3 install.py").
### Fixed
- install.py: the server start commands (macOS shim, Windows task, the
  no-autostart hint) were missing --static-dir - the server answered the API
  but could not serve the pane/simulator pages. They now pass it exactly like
  install.sh always did.

## 0.9.23
### Added
- **install.py**: a unified, interactive, cross-platform installer
  (macOS + Windows, auto-detected) that replaces install.sh as the
  recommended install path. Menu: Install / Update / Uninstall; also
  --install (non-interactive), --update (git pull + update), --uninstall
  (--purge-config optional), --port, --ribbon (still opt-in: the ribbon
  commands remain experimental on this Office build), --dry-run. The
  autostart question (default Yes) explains what it registers: macOS login
  item via the in-bundle SMAppService helper (System Events fallback) /
  Windows per-user scheduled task (HKCU Run-key fallback). The transport is
  plain http://localhost:<port> only - Office allows localhost over http, so
  the installer no longer touches certificates at all; expected OS popups are
  listed up front. Windows sideload = HKCU trusted folder catalog (no admin);
  Outlook (Windows) needs the once-per-workstation "Add from File" dialog.
  Uninstall removes manifests, autostart and the installed files but keeps
  ~/.officekick (config + sessions) by default. install.sh remains for
  macOS-specific use.

## 0.9.22
### Fixed
- Background-items prompt: the server's login item was registered through the
  System Events AppleScript API, so macOS named its "can run in the
  background" notification "System Events.app" (scary, and misleading - the
  recorded item is just the autostart app), and every install re-registered
  the item, repeating the prompt. Registration now runs through a small
  compiled helper inside OfficeKickServer.app (tools/
  officekick-register-login-item.m, SMAppService): macOS attributes item and
  notification to "OfficeKickServer", repeat installs re-register nothing (no
  notification), leftovers from earlier installs are removed once, and the
  bundle is ad-hoc signed after every build. Falls back to the old
  System Events registration when clang is unavailable.

## 0.9.21
### Added
- SlideKick: **read_slide** - the model can ask for the FULL text of up to 12
  slides in one go by replying `{"kind":"read_slide","slides":[N]}` (1-based
  outline numbers). SlideKick reads those slides from the deck (real and
  simulated), re-asks the model with the complete text (same flow as live web
  lookups; max 2 read rounds per run), and the answer then uses the exact
  content. In large decks (more than 40 slides) the compact outline now also
  lists every shape of the SELECTED slide(s) with their complete text, so a
  slide the user has open is fully visible without the extra step.
### Changed
- In "Previous requests", the request excerpt of each collapsed turn now
  shows in blue (light blue in the dark theme) and semibold so it stands out
  from the request/response body (all three add-ins).

## 0.9.20
### Changed
- macOS server autostart moved from the LaunchAgent to a **login item**
  (System Events) running OfficeKickServer.app in the user session; the app
  is LSUIElement now and the wrapper keeps a KeepAlive-style respawn loop
  (restarts after crashes, stops cleanly on a graceful SIGTERM). The
  LaunchAgent stays installed but unloaded as a fallback.
### Fixed
- The app wrapper's default launch (no arguments) forgot --static-dir, which
  made all statics 404 ("Not found: slidekick/static/taskpane.html") after
  the login-item switch.

## 0.9.19
### Added
- Docs: README "macOS permissions (locally hosted models)" section - the
  Local Network grant flow for LAN providers (OfficeKickServer.app dialog,
  System Settings > Local Network fallback, the Errno 65 symptom, the
  terminal-session bypass) + pointers in all three add-in docs. Documents
  that `tccutil` cannot reset the LocalNetwork service on current macOS
  builds; no --prompt flag (would mislead).

## 0.9.18
### Added
- **macOS Local Network permission, properly attributable**: the installer now
  creates `~/Applications/OfficeKickServer.app` and runs the LaunchAgent's
  server THROUGH it, so macOS can raise the "wants to access devices on your
  local network" prompt for the server process (a bare launchd-spawned python
  is silently denied and LAN providers get Errno 65). Grant it once in the
  prompt or System Settings > Privacy & Security > Local Network.
- Manifest `<Version>` bumped to 1.0.2.0 in all three add-ins: Office keys the
  parsed-manifest cache by id+version, so an unchanged version keeps a stale
  registration (one more reason a ribbon attempt could look like a no-op).

## 0.9.17
### Changed
- The waiting-for-the-model timer now formats progressively: seconds only
  under a minute (7s), adds minutes past it (1min 05s) and hours past an
  hour (1h 02min 03s) - all three add-ins. Styled italic, dark green on the
  light pane and light green on the dark pane. Verified live in both
  branches plus the styles.

## 0.9.16
### Changed
- "+ New connection" now puts the cursor in the name field when the
  connection form opens (all three add-ins). Saving still leaves the focus
  where it was.

### Fixed
- Connection/provider test failures to LAN targets now explain the likely
  macOS cause: a server process without macOS "Local Network" permission
  gets Errno 65 (No route to host) for private IPs even though the same
  URL works from a shell. The hint lists the System Settings grant and the
  "restart the server from a permitted session" fallback.

## 0.9.15
### Changed
- Reverted the ribbon default to the list-mode manifest: on the tested
  Office-Mac build a commands manifest REPLACED the Add-ins-list entry and
  then rendered neither the AI group nor a list item (the add-in was
  unreachable). `--ribbon` remains as the opt-in; installer messages +
  README + docs updated with the exact symptom and the revert path.

## 0.9.14
### Fixed
- Ribbon manifests kept the raw `__COMMANDS_BEGIN__/__COMMANDS_END__`
  placeholder lines inside <OfficeApp> - valid XML but not valid manifest
  content; Office rejected the whole add-in ("Add-in Error: this add-in is no
  longer available"). Installers now strip the marker lines in ribbon mode
  (the commands block itself is kept).

## 0.9.13
### Changed
- The ribbon "AI" group is now the default install mode (`--ribbon` inverted:
  `--no-ribbon` restores the old list-mode manifest for 16.114-era Office-Mac
  builds). README + slidekick.md updated; installer messages refreshed.

## 0.9.12
### Added
- **Session auto-clean (global setting)**: "auto-clean up history after X days"
  on the Settings > General tab (Sessions group) - enabled by default with
  120 days. Opening any add-in deletes sessions whose NEWEST message is older
  than that (sweep throttled to once per hour); days are validated 1..3650 and
  the setting survives reopens (projected by /api/config + writable via
  /api/settings). Verified: 200-day-old session deleted at the default, kept
  at a custom 1-day setting matching a 1-day-old cutoff; docs updated.
- **Ribbon "AI" group (opt-in `--ribbon`)**: manifests now carry an AI group
  with the add-in icon and a "Open ...Kick" button on the built-in Home tab
  (SlideKick/DocKick; MailKick on the mail read/compose tabs via TabDefault).
  Group label simplified to "AI". Default install stays `--no-ribbon` (the
  16.114-era Office-Mac commands non-rendering observation); README + docs
  updated with the flag and the revert path.

## [0.9.11] - 2026-10-31
### Changed - approval gate
- A horizontal divider now separates the proposed-actions list from the
  "don't ask again for '<action>' this session" checkboxes (all three
  add-ins; divider only renders when at least one confirm-mode action is
  present).

## [0.9.10] - 2026-10-31
### Changed - proposed-actions layout
- The approval gate rows are reorganized for the narrow pane: line 1 carries
  the action tag, the approval checkbox (right-aligned) and the red cancel
  button; the action explanation moves to its own line BELOW instead of being
  squeezed next to them. Two-click confirm/revert on the cancel button and the
  checkbox semantics are unchanged. Applied to all three add-ins; verified at
  pane width and forced ~300 px (no horizontal overflow).

## [0.9.9] - 2026-10-31
### Added - API sweep round two
- SlideKick: "Merge deck…" button next to Refresh deck - pick a .pptx and
  all of its slides are appended after the last slide
  (`presentation.insertSlidesFromBase64`, PowerPointApi 1.2, source
  formatting kept).
- DocKick: `reject_changes` op - rejects ALL tracked changes
  (`TrackedChangeCollection.rejectAll()`, WordApi 1.6) - counterpart of
  accept_changes.
- DocKick: `set_header` / `set_footer` ops - write text into the page header /
  footer of section 1 (`section.getHeader/getFooter`, WordApi 1.1).
### Investigation result - slide comments NOT implementable
- PowerPoint slide comments are NOT available to Office.js add-ins in any
  released PowerPointApi requirement set (the only "comments" API on the
  PowerPoint side is the DocumentProperties metadata field, which Microsoft
  documents as having "no connection to comments made in slides"). There is
  no op for slide comments until Microsoft ships a released API for them.

## [0.9.8] - 2026-10-31
### Added - proactive API sweep (Outlook / Word / PowerPoint)
- MailKick: `new_meeting` op - opens the Outlook MEETING form (subject,
  required/optional attendees, ISO start/end, location, body) for review; the
  invite is never sent by MailKick (approval default `auto`).
- MailKick: `draft_reply` accepts `"all":true` - reply-ALL to everyone on the
  thread (`displayReplyAllForm`, Mailbox 1.1).
- MailKick: the email context lists ATTACHMENTS (name + size) of the open
  message, so drafts can reference them by name (contents are not read).
- DocKick: `add_hyperlink` op - links the first occurrence of a text to a URL
  (`Range.hyperlink`, WordApi 1.3).
- DocKick: `format_selection` op - the format changes apply to the text the
  user currently has selected.
- DocKick: `format_text` accepts a `style` (Heading 1-3, Title, Subtitle, List
  Bullet/Number, Quote, Normal) via `styleBuiltIn` - re-styles whole paragraphs
  (e.g. turn something into a bullet list).
- SlideKick: `style_shape` op - solid fill and/or outline color for one shape
  (`ShapeFill.setSolidColor` / `ShapeLineFormat`, PowerPointApi 1.4).
### Fixed
- A failing op in the MAIL apply chain could leave the pane stuck on
  "Applying..." (same per-op catch fix DocKick got in 0.9.6).

## [0.9.7] - 2026-10-31
### Added - set track changes (DocKick)
- New op `set_track_changes` with `on:true|false`: flips the DOCUMENT's own
  Track Changes mode (the Review-tab toggle, `document.changeTrackingMode`, 
  WordApi 1.6) - the model can now actually do "turn off track changes"
  instead of answering that it only works in Word. It does not touch existing
  revisions (offer accept_changes for those) and leaves the pane's separate
  "AI edits as tracked changes" setting in Settings alone.
- Approval default `confirm`. E2E in the simulator: gate shows the toggle,
  apply flips it, `trackedChanges` config untouched.

## [0.9.6] - 2026-10-31
### Added - accept tracked changes (DocKick)
- New op `accept_changes`: accepts ALL tracked changes in the document body
  (`body.getTrackedChanges().acceptAll()`, WordApi 1.6) - finalize the edits
  the AI made with revisions on. Ask for it ("accept the changes") and the
  model emits it after its edit ops; approval default `confirm` (accepting is
  final). Works with or without the tracked-changes setting enabled.
### Fixed - a failing LAST op stalled the apply
- An error thrown by the last applied op rejected the whole apply chain: the
  per-op error row, the revision-mode restore and the final status never ran
  and the pane stayed on "Applying…". Per-op errors are now captured inside
  their chain link. (Surfaced while E2E-testing the new op in the simulator.)

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
