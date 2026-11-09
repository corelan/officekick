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

## 0.9.277

### Fixed
- field report (french Office, fresh Windows box): the classic add-ins
  dialog said "veuillez ajouter ou activer des catalogues depuis le Centre
  de gestion de la confidentialite" - the installer's TrustedCatalogs entry
  wrote the enable value as `flags`; office reads `flag` (SINGULAR). With
  the wrong name the catalog existed but was never ENABLED. The value now
  ships as flag=1 ( Url + singular flag), stub-winreg battery asserts the
  exact layout Office honors, and docs gain the manual Trust Center +
  reg-add unblock for machines that installed older builds.

## 0.9.290

### Added
- a Style tab in the settings of all three panes. It carries the renamed
  "Custom style prompt" (moved there from General - the storage key and the
  built-in default are unchanged, existing text carries over), plus the new
  image-style selector: photorealistic (the default, also when nothing is
  selected), cartoon, anime, watercolor, oil painting, flat vector, 3d
  render, pixel art, children's book illustration, line art. The server
  stores the choice (imageStyle), echoes it in /api/config, accepts it via
  /api/settings (an empty value reads as the default), and appends it as
  the FINAL directive of every image generation prompt (both the async
  image jobs and the sync generate path, after the whitespace rule) - so
  an unstyled ask never falls back to the model's own taste. Verified by a
  unit battery (canonicalization incl. space/case-insensitive ids, the
  compose output, empty -> photorealistic) and a live roundtrip (PUT
  cartoon -> echoed; PUT "" -> photorealistic).

## 0.9.289

### Added
- the profile gate: the working tab ("Sessions") stays DISABLED until at
  least one profile exists. A pane that boots with no profiles opens
  straight into Settings > Connections and shows a purple box with the two
  setup steps ("(1) create at least one CONNECTION ... then (2) create a
  PROFILE ...; the Sessions tab unlocks as soon as a profile exists"). The
  gate re-evaluates on every config load: saving a profile re-enables the
  tab and removes the box, with no reload of the add-in needed beyond the
  normal pane refresh. All three panes (SlideKick, DocKick, MailKick) +
  the disabled-tab styling; verified live in the simulator: fresh state =
  disabled + directed + box; after one profile = enabled on the Sessions
  tab.

## 0.9.288

### Fixed
- uninstalling with --purge-config destroyed config.json unrecoverably
  (plain rmtree of ~/.officekick: API keys, profiles, sessions) while the
  flagged path and the interactive ask look identical in the output - the
  user reported his config vanishing from an uninstall he ran the day after
  reading the "everything" command. The purge now MOVES the folder to a
  timestamped sibling next to it (<config dir>-purged-<timestamp>, same
  name on every OS) - everything stays recoverable - and only wipes when
  the move itself fails. The interactive question now says the folder is
  moved aside, not deleted. Battery: purge keeps config.json + logs intact
  inside the moved folder, a second purge with nothing present is a no-op,
  and the keep path leaves the file untouched.

## 0.9.287

### Fixed
- "write text onto the slides" landed as five failures on a fresh build:
  every story slide came from a TITLE-ONLY layout (no body placeholder) and
  the body write threw "No body/content placeholder found on slide N" for
  each one, retry included - while a delete_slide op in the same plan hit
  "this build cannot delete slides". Two changes: edit_slide's body write
  now FALLS BACK to a new text box at the standard body band (48,190,864x300
  per the make-over grid) when the slide has no body placeholder - two
  creation forms mirroring the 0.9.233/0.9.261 textbox ladder, the friendly
  error only when both fail, and the row reports "content set in a new text
  box". And the deck context now carries canDeleteSlides (probed from
  slides.removeAt): the outline gains a SLIDE DELETION row and the ask gains
  a rule that forbids delete_slide ops on builds without it, telling the
  model to restructure by adding slides instead.

## 0.9.286

### Added
- one-click updates: the panes' check-for-updates button is now
  "Update now (pull + redeploy + restart)". It calls the new
  POST /api/update-run, which spawns install.py --update from the recorded
  clone as a detached process (git pull + full deploy + server restart -
  the exact flow that previously needed a command line), then polls
  /api/health: the server drops mid-run (the installer frees the port) and
  the button reports when the new build answers, including the version and
  the pane-reload nudge. The update.log trail: install.log in the clone +
  update_run.log in the config folder. Refuses a second run while one is
  fresh (10 min window); the pull-only flow remains available via the CLI.

## 0.9.285

### Fixed
- anchored inserts STILL threw "this._wholeParaRangeOf is not a function"
  after 0.9.283 - the actual root is one level deeper: Word.run(fn) does NOT
  bind the request-builder callback, so `this` inside every op callback was
  never the deck; the 0.9.283 .bind(this) chains were bound to that unbound
  base and the method lookup still missed. WordDeck.run now invokes
  fn.bind(this) once at the root (this._wrap and every nested bound chain
  resolve correctly), and all five anchor helpers go through a module-local
  _wholeParaRange() that needs no `this` at all. Both layers verified by
  node --check plus the anchor-site audit (no this.-based wholeParaRange
  callers remain).

## 0.9.284

### Changed
- flag-driven installer runs no longer pause for a keypress at the end
  (field: a scripted `install.py --update` sat waiting on "press any key").
  --update / --install / --uninstall / --stop / --start / --restart /
  --status / --transport / --cert-check / --dry-run all exit on completion;
  the interactive menu keeps the hold-the-screen behavior. Verified with a
  pseudo-tty run of --cert-check exiting 0 unattended.

## 0.9.283

### Fixed
- anchored images in DocKick died with "this._wholeParaRangeOf is not a
  function" (field: a story ask with three images placed after specific
  paragraphs - all three failed, retry included). The anchor branch of
  addImage called this._wholeParaRangeOf inside a ctx.sync().then(...) whose
  outer callback was not .bind(this)-ed, so `this` was undefined there.
  insertParagraphs' anchor branch had the same latent hole (end-of-document
  inserts never reach it, which is why the paragraphs in the same ask had
  worked). Both outer callbacks are now bound; the end/shortage branches
  are untouched.

## 0.9.282

### Fixed
- the leaf certificate regenerated (and re-consented + re-trusted) on EVERY
  install.py --update on a Windows box without an openssl binary.
  validate_existing_tls already had a cryptography fallback for the
  expiry/SAN checks, but the pair check (_pair_matches) was openssl-only:
  with openssl absent it returned False unconditionally, so ensure_leaf_certs
  saw "the private key does not match the certificate" every single run and
  walked the whole consent flow again (field: the fresh Windows 11 box -
  install at 17:47, update minutes later, certificates phase again). The
  check now falls back to DER-comparing the certificate's public key with
  the private key's public half via cryptography. Battery: openssl path,
  forced no-openssl fallback (match/mismatch), junk inputs.

## 0.9.281

### Fixed
- DocKick on non-English Word: insert_paragraphs (and rewriteDoc) assigned
  paragraph.style with ENGLISH display style names ("Title", "Heading 1").
  .style is the LOCALIZED style name property, so French Word ("Titre",
  "Titre 1") threw InvalidArgument for every styled paragraph and the op
  failed wholesale - even though the same op had always worked on en-US
  builds (field: fresh Windows 11 fr-FR box, "Insert 15 paragraph(s) ->
  end of document (title, normal) - InvalidArgument"). All five style
  assignments now go through setWordStyle(): styleBuiltIn with the
  Word.BuiltInStyleName token (identical tokens in every UI language,
  WordApi 1.3), falling back to the legacy .style assignment for very
  old builds, else the paragraph keeps its default style and the insert
  still succeeds. createTemplateBlock's "Heading2" token fixed to the
  enum spelling ("heading2"); the simulator maps the listParagraph token.

## 0.9.280

### Fixed
- openrouter sessions failing with HTTP 401 "Missing Authentication header":
  the connection had been saved WITHOUT an API key (the pane list even shows
  "key: not set" under the row, easy to miss), so the request builder omitted
  the Authorization header entirely - OpenRouter answers exactly that 401
  string for a header-less POST. Two changes: the save endpoint now returns
  needsKey=true for key-needing types (openai, anthropic, openrouter,
  generic-openai, llamacpp) when the stored key is/would be empty, and all
  three panes render a loud one-liner instead of the quiet "Ready for the
  next one". Setting the key later clears the warning (verified: keyless,
  with-key, ollama keyless, resave, key-after-the-fact).

## 0.9.279

### Changed
- the TrustedCatalogs Url now carries a trailing backslash. The office
  add-ins dialog on the field box still said "add/enable catalogs from the
  Trust Center" after the flag fix + a verified install - folder-catalog
  Urls are stored WITH the trailing backslash when added via the Trust
  Center UI, so the bootstrap entry now matches that form.

## 0.9.278

### Fixed
- the BIG one from the fresh-Word box: the Windows add-in registrations
  were DEAD CODE. sideload_windows_developer (the HKCU Wef\Developer
  launch-time sideload) and sideload_windows_catalog (the Trust Center
  folder catalog) were defined but never called anywhere - every Windows
  install since the restructuring deployed the manifests and then
  registered NOTHING, so Word's add-in list stayed empty no matter what
  (the field box: manifests present, both registry layers absent). The
  registering phase now runs both on every Windows install (dev entries
  for each rendered manifest + RefreshAddins; the TrustedCatalogs entry
  with Url + the singular flag=1), dry-run still writes nothing. Batteries:
  wiring (both layers fire; catalog layout; dry-run clean) and the dev
  layer with real rendered manifests (RefreshAddins + one default-value
  entry per add-in).

## 0.9.276

### Fixed
- the startup dependency install could fail OPAQUELY on a fresh box: the
  managed Python 3.15 runtime may lack a cryptography wheel (a source build
  needs Rust) or even pip itself. The startup ensure now (1) bootstraps pip
  with ensurepip when `pip --version` fails, (2) tries
  `--only-binary=:all:` FIRST so a missing wheel fails in seconds instead
  of a long broken source build, then the normal --user/bare attempts, and
  (3) prints the pip failure reason ON THE CONSOLE (it only went to
  install.log before - field report: the picker locked with no reason
  visible). The picker's lock text no longer claims the package installs
  "during the certificate step" - it names the failed startup install,
  install.log, and exactly what to run.

## 0.9.275

### Fixed
- field report: Outlook was LOCKED in the add-in picker because the
  'cryptography' package was missing, with the install deferred to the
  certificate step - the user rule is dependencies STARTUP, never mid-flow.
  install.py now ensures the runtime dependencies at startup (Windows:
  cryptography via pip --user, idempotent, no admin; skipped for dry-runs
  and macOS where openssl covers it), so the picker unlocks itself. The
  picker's lock survives only when the startup install genuinely failed,
  with the run-it-yourself command named. Battery: present=no-op, missing=
  one --user pip attempt + verification, per-process idempotence, dry-run
  and macOS skips, pip-failure keeps the lock.

## 0.9.274

### Changed
- the installer's read path is cached (~2 s): the menus called
  _config_settings FOUR times per redraw, each with its own loopback HTTP
  attempt. On machines where a refused-or-filtered loopback connect waits
  the socket timeout, that multiplied into a visibly slow transport menu;
  the cache now collapses a draw to one call (and arrow-key redraws to
  zero), with the per-call timeout tightened 1.5 s -> 0.6 s (a local server
  answers in milliseconds). Writes and the settings verification invalidate
  the cache, so no stale reads. Battery: cold draw ~130 ms -> cached draw
  ~0 ms, post-write invalidation verified.

## 0.9.273

### Fixed
- field report (first real Windows box): choosing a transport in the menu on
  a machine where install.py had never run crashed with FileNotFoundError -
  _persist_file wrote ~/.officekick/config.json without the config folder
  existing. The writer now creates the folder (a fresh-machine transport
  switch stores its mode cleanly, before any deploy). Battery-verified:
  no config dir -> set https -> file created + mode stored; a second set
  merges ports. Also on this box (all shipped earlier): py install default
  (the manager rejects 'latest' as a tag), manual-first clone instructions
  for cmd + PowerShell (%USERPROFILE% vs $env:USERPROFILE), the bootstrap
  script's self-identifying version line + empty-ToolsDir guard.

## 0.9.238
### Added (clarify: placement)
- The clarify round now also covers PLACEMENT ambiguity: the model may ask
  where the result belongs (open slide vs a named one vs a new slide after
  the named one; DocKick: where in the document; MailKick: which of the
  multi-selected items). When the ASK names a specific slide and the reply's
  ops aim at a DIFFERENT one (live case: 'turn slide 124 into an
  infographic' with slide 125 open landed on 125), the pane re-asks ONCE
  with the reason (WRONG TARGET) and re-checks before rendering the gate;
  a model that stays put runs as before. After a clarify pick, the re-ask
  carries a slide reminder when the request named a slide.

## 0.9.237
### Changed (all add-ins)
- The clarify question now renders in its own slot RIGHT below the status
  line and above the run's activity trail (the box previously landed in the
  response zone, below the activity log). The pick shows 'Your answer: ...'
  in the box and the Q&A pair files into Previous requests from the slot;
  session/New-session cleanup paths clear the slot too.

## 0.9.236
### Added (all add-ins)
- The clarify round no longer depends on the model volunteering: the word
  'infographic' (slide/doc, with no image/picture word in the ask) now ALWAYS
  surfaces the image-vs-shapes(-vs-both / vs-table) question before anything
  is built. Other ambiguous words stay model-decided via the prompt rule.
### Fixed (SlideKick)
- Textbox ladder hardening for the 16.114 mac bridge: variant 3 now creates
  the rectangle through the PROVEN geometric path (creation + styling in one
  run) and writes the text in an ISOLATED second run against the created
  shape's id - a textFrame rejection can no longer take the created shape
  with it; a failed text write reports the shape + the reason in the apply
  line instead of failing the op.

## 0.9.235
### Added (all add-ins)
- Clarification round: when a request could plausibly deliver a DIFFERENT
  form of the result (the "infographic" case: an AI-generated image vs.
  editable shapes + text vs. both), the model can now ask the user once
  instead of guessing. It replies {"kind":"clarify","question":...,
  "options":[2-4 short options]}; the pane shows the question with clickable
  option buttons plus a free-text answer row; the pick re-asks immediately
  (the choice rides inside the instruction so every later lookup/repair
  re-ask carries it) and the model delivers the final result without asking
  again. Internal action, never gated; max 2 clarify rounds per chain; the
  Q&A pair lands in Previous requests and in the session history. Rules
  added to all three system prompts (MailKick's version says a clarify reply
  replaces - not supplements - the ops/reply round).
## 0.9.271

### Added (repository tooling)
- tools/setup-windows.ps1: a PowerShell dependency installer for Windows -
  installs Python through the official PYTHON INSTALL MANAGER downloaded
  from python.org (Add-AppxPackage -AppInstallerFile ...; never winget for
  Python; direct FTP-index MSIX fallback; then `py install latest`) and
  Git for Windows (winget, git-for-windows release fallback),
  then offers an interactive git clone of the repository into
  %USERPROFILE%\Tools\officekick (an existing clone updates with
  git pull --ff-only). Per-user, no admin, unattended flags (-Yes,
  -NoClone, -SkipPython, -SkipGit, -ToolsDir, -CloneUrl). Docs:
  docs/setup-windows.md (running the script + the execution policy
  explanation - Process-scope Bypass one-off, CurrentUser RemoteSigned as
  the standing setting, Unblock-File for the downloaded-file mark).

### Added
- TARGET CONFIRMATION LADDER (user rule: when it is not clear WHERE to
  insert/write/change/delete/move, confirm - AI first, then the user):
  (1) an unnamed target (targeting verbs, no slide number, the plan writes
  exactly one existing slide) gets ONE AI confirmation round - the model
  sees the original prompt and either re-confirms deliberately (the plan
  runs), asks the user itself ({"kind":"clarify"} renders), or corrects;
  (2) when the insert-ask guard drops EVERY action (the model kept the
  wrong aim after its confirmation round), the pane asks the USER:
  "Where should the work land?" with interactive options - stick to the
  named slide, override with the plan's aim (that choice then skips the
  guard for the re-run), or cancel - plus free text, and the pick
  auto-reruns the ask with the answer embedded. The AI round sees the
  original conversation; one confirmation round per run.

### Fixed
- the stage-1 continuation hung the run (the return skipped renderParsed);
  it now renders the confirmed/clarify result like the other re-ask paths.

## 0.9.270

### Fixed
- the target guard is now a HARD WRITE GUARANTEE, not a nudge. Field proof:
  the cut-off plan that lost its creates re-aimed and the pane auto-applied
  20 actions to the user's real slide even after the guard's re-ask. An
  insert-after ask now stashes its constraint at parse time; validateOps
  DROPS any op that writes an existing slide outside the named one and the
  slides that batch itself creates (visible invalid rows: "target guard:
  this ask inserts slides after slide N - this action wrote to existing
  slide M, which this batch does not create; dropped to protect your
  content"). The constraint lives for the whole run (survives the
  re-ask/repair/recovery re-validations - each round re-checks and re-asks
  the model with the guard reasons), resets at every run start, and is
  bypassed exactly where it should be: ops for the named slide and ops on
  the slides the batch creates. Verified in the sim with a stubborn
  wrong-aimed reply: guard re-ask -> still wrong -> all 2 actions dropped
  with reasons -> nothing applied, no completed actions.

## 0.9.269

### Fixed
- field-run x2: an insert-after ask whose plan lost its create_slide ops
  (cut-off reply, recovery merge) made the target guard re-ask with "re-emit
  ... aimed at slide N" - which TAUGHT the model to drop the inserts and
  paint the user's real slide (20 auto-applied actions on the named slide
  in the worst case). The guard's re-ask nudge is now insert-aware: for
  after/insert asks it demands create_slide ops after the named slide, aims
  the new-slide ops at the slides the model creates, and restricts writes
  beyond the named slide. The generic nudge stays for non-insert asks.
  Verified with a captured /api/chat body: the nudge carries the insert
  intent.

## 0.9.268

### Fixed
- the target guard understood "slide N" but not "insert slides AFTER N": an
  ask like "make a visualization of slide 138, insert slides after it" got
  its correct plan (creates at 139/140) rejected as a wrong-target and the
  model was re-asked to aim AT the named slide - so the run painted the
  image and the infographic onto the user's real slide while the two new
  slides stayed empty. The guard now honors after-insert semantics: when
  the ask uses after/insert/new-slide wording, the plan carries create_slide
  ops, and every target sits in named+1..named+creates within the virtual
  deck, the plan passes instead of being re-asked. Genuine misses (the open
  slide aimed at instead of the named one, no creates) still re-ask.
- the find-based format_text rows in that run failed with "VR & ED"/
  "Consider using not found" because the model numbered ITS OWN newly
  created textboxes wrongly (shape 6 = its "Text Reasoning" box, not the
  VR & ED panel copy - apply-time guard, correct behavior).

## 0.9.267

### Fixed
- field report: "it keeps asking for the approval, even when I click the
  checkbox" - the session-scoped approval memory was IN-PAGE ONLY, and the
  Office webview reloads the pane page between runs (heavy deck edits:
  slides created, backgrounds applied), so every reload silently wiped the
  remembered approvals and the resumed session asked again. The approvals
  are now persisted per session id (localStorage, app-qualified key) and
  restored on the next boot - via the auto-resume path (the commit stamps
  the resume marker's session id when the run has not saved yet) or, on a
  plain reload, the boot-time restore. New/resumed/different sessions and
  session resets still wipe the stored entry (session-scoped stays true).
  Verified in the sim incl. a REAL page reload between runs: run 2 renders
  the remembered types as auto rows with no remember boxes.

## 0.9.266

### Changed
- richer visual vocabulary for designs: create_shape accepts a curated
  PowerPoint geometric-glyph set well beyond the basic 8 - chevron,
  right/down/up/leftArrow, donut, plus, hexagon, star5/star8, teardrop,
  parallelogram, pentagon, octagon, cube, can, moon, lightningBolt, sun,
  heart, wave, doubleWave, arc, blockArc, pie, foldedCorner, halfFrame,
  smileyFace - plus the model-friendly aliases circle/oval (ellipse) and
  rect (rectangle). Unknown kinds still ladder down to a rectangle at
  apply time with the substitution note.
- MODERN VISUAL DESIGN contract added to the ask rules: whenever the ask
  creates or redesigns an infographic/diagram/flow/timeline/dashboard, the
  model composes in layers from the deck's own theme palette (background
  statement -> structure + glyph accents -> text last), builds ICONS as
  picture-filled circles/squares via set_shape_image (flat vector icon,
  one accent color, soft padding; one image set per slide, max 4), uses
  asymmetry + oversized scale contrast, and is steered away from the 80s
  tells (1px outlines everywhere, saturated primary fills, equal centered
  boxes). Honest limits named: no shadows/blur in the API (depth from
  transparency layers only), roundedRectangle renders sharp on Mac builds.

### Fixed
- a validateOps regression from the vocabulary edit: the alias resolver
  dropped the kind variable declaration (the pane is strict mode) - the
  first create_shape batch died with "kindS is not defined". E2E-proven
  the circle/chevron/star5/donut/rightArrow batch validates and applies.

## 0.9.265

### Fixed
- image placement is deterministic now: the direct setSelectedDataAsync
  insert always lands on the slide the user has OPEN (no slide-targeting
  exists in the PowerPoint JS API), so an op naming a different slide
  (typical: a slide the batch just created) went to the wrong slide. The
  apply now goes straight to the picture-filled-shape path for off-slide
  targets: the image is inserted as a picture-filled rectangle ON the
  requested slide via fill.setImage, and the row reports
  "inserted on slide N as a picture-filled rectangle ...". Same-slide and
  simulator runs keep the direct insert (with the size ladder when the
  host rejects the payload).

### Fixed (validator)
- a create_shape textbox earlier in the SAME batch was invisible to the
  validator's text-state: a later format_text with find "2 · Wire" against
  the freshly created textbox read the pre-batch snapshot and
  false-rejected "not found in that shape's current text". The batch's own
  textbox text is now registered in the text state (when the slide's shape
  count is detailed) and the verified sim flow shows the find-format
  validating and applying. (Find mismatches on shapes the model numbered
  wrongly still fail at apply time with the self-diagnosing reason - that
  guard is the failsafe, unchanged.)

## 0.9.264

### Fixed
- field evidence from 16.115: roundedRectangle fails BOTH creation forms
  (native options + bare-create-with-properties) with InvalidArgument on
  the same slide where rectangle/line/triangle/textboxes all apply - the
  build rejects the KIND. create_shape now has a third variant for
  geometric kinds: a plain rectangle at the same geometry, and the row
  reports "this PowerPoint build rejected '<kind>' - used a plain
  rectangle".
- the image-insert failure was NOT size (both the raw 113 KB and the JPEG
  153 KB payloads failed with the same host error). addImage attempt 3
  inserts the picture as a picture-filled rectangle via fill.setImage
  (the same primitive as set_shape_image, PowerPointApi 1.10) at the
  requested geometry; the row reports "[inserted as a picture-filled
  rectangle ... - the direct image insert failed on this host (...)]".

## 0.9.263

### Fixed
- SlideKick image inserts ladder like the shape/textbox paths: attempt 1 =
  the raw base64 (data: prefixes stripped defensively); attempt 2 = the
  same image re-encoded IN-HOST via canvas (JPEG 0.85, white-backed) and
  re-inserted at the requested geometry. PowerPoint-Mac rejects some
  payloads with a bare host error ("An internal error has occurred") -
  oversized generated PNGs are the usual trigger. Success reports the
  conversion ("image inserted ... [inserted as JPEG M KB (the host
  rejected the original N KB payload)]"); a double failure carries both
  payload sizes.

## 0.9.262

### Fixed
- the gate's "don't ask again for <type> this session" was effectively
  "until the pane reloads": the remembered approvals were only ever SET,
  never cleared. They are now session-scoped for real - a new chat session
  (chat-tab New + History-tab New session) and any resumeSession clear
  SK.sessionApprovals, so the new session asks again. Verified in the sim:
  same-session ask 2 renders the type as auto (no approval needed); after a
  new session the gate asks again as expected.
- the remember checkbox only committed when the batch's apply ALSO approved
  at least one action of that type - ticking it alone did nothing, silently.
  It now always commits, and the gate result shows "remembered for this
  session: later '<type>' actions run without asking".
- behavior note (already true, now verified): a remembered type in a MIXED
  batch renders as an auto row (no approval tick needed); when the pane's
  auto-apply toggle is on and every action is remembered/auto, the proposal
  box is skipped entirely.

## 0.9.261

### Fixed
- SlideKick on macOS: create_shape with a geometric kind could die with a
  bare "InvalidArgument" at sync (roundedRectangle threw on 16.115 while
  addTextBox/addGeometricShape("rectangle") with the IDENTICAL options
  object succeeded). Non-textbox kinds now ladder like the 0.9.233 textbox
  repair: v1 native addGeometricShape(kind, options); v2 the docs-sample
  form - create bare, then write left/top/width/height as shape properties.
  The retry pass announcement stays one pass.
- format_text (and find-based variants) targeting a shape that does not
  exist on the slide (typically planned for a create_shape that failed in
  the same batch) now fail with a self-diagnosing reason and the slide's
  real shape count instead of PowerPoint's bare GeneralException.

## 0.9.260

### Changed
- the user must ACCEPT the local HTTPS certificate in one more OS-level
  dialog: the installer now says so up front on the consent screen and
  prints a heads-up right before the dialog appears (macOS: the keychain
  confirmation for the leaf/CA trust import - click Trust; Windows: the
  Root-store certificate dialog - click Yes). docs/certificates.md and
  docs/installer.md document the dialog step.

## 0.9.259

### Fixed
- the About tab's "Add-in identity (this manifest)" showed "?" on Word: the
  pane looked up the manifest version under the key "wordkick" while the
  server publishes the identities as dockick/mailkick/slidekick. DocKick now
  reads its own key.

## 0.9.258

### Changed
- the pane footer status line now distinguishes the transports: https keeps
  the lock + green; plain HTTP loses the lock and renders the line BLUE in
  the light theme (#0078d4) / CYAN in dark (#22d3ee). Class-driven
  (`#statusbar.http`) and applied in all three panes (boot line, heartbeat
  repaint, and the version-mismatch refresh).
- user-verified on Windows: Office's WebView2 rejects a leaf-only trust even
  from Trusted People (the add-in error), and the Root-store import is the
  working path - docs/certificates.md now says so directly.

## 0.9.256

### Fixed
- the leaf-trust consent default followed the transport NEED: with the https
  transport stored, Enter means trust (the pane requires it); declining still
  never touches the store and the warning now spells out the consequence and
  the re-run command. A declined tree caused the "Trusted People is empty"
  / blocked-add-in confusion on Windows.

## 0.9.257

### Fixed
- the Windows leaf import is now PowerShell-FIRST, verified, and the flow
  runs the decisive check: Import-Certificate (returns the thumbprint; we
  compare it) with a certutil fallback, store-listing verification, and a
  REAL https test request through the WINDOWS TLS stack (SChannel, PowerShell
  Invoke-WebRequest). 200 = Office's webview will accept the leaf; anything
  else is shown as "the Windows TLS handler REJECTS the leaf" with the
  Root-store option offered right there. The Root thumbprint is recorded in
  the cert metadata and removed again by the uninstall.

### Note
- python's own https probe uses OpenSSL with its own trust store - it never
  consulted the Windows certificate stores, which is partly why the earlier
  flow could not confirm what Office would see.

## 0.9.255

### Added
- Windows leaf trust: Office's WebView2 does not always honour a leaf-only
  trust from the Trusted People store. After the (silent, by design)
  TrustedPeople import the installer now offers an explicit escape hatch:
  also register the leaf in the CurrentUser ROOT store (the leaf is CA:FALSE,
  so the Root entry trusts exactly that one certificate; Windows shows its
  own confirmation dialog - the prompt the user looked for). The uninstall
  removes the Root entry again (tracked by thumbprint in the cert metadata).
- the Trusted People import is verified by thumbprint before the flow
  continues.

### Changed
- status line alignment is uniform: one helper (_tagline) gives ok/warn/step
  tags (and the manual dry-run lines) the same message column on both
  platforms - the uninstall output no longer mixes 1/2/3-space paddings.
- uninstall grouping: blank lines around the certificate question block and
  between the two removal questions.

## 0.9.252b

### Changed
- the menus now DEFAULT to what is already selected: the Transport menu's
  cursor starts on the current mode (auto/http/https), the certificate-mode
  submenu's cursor starts on the current certMode, and a bare Enter in the
  typed fallback re-selects that pre-selected row instead of leaving the
  menu.

## 0.9.252

### Fixed
- the START menu's transport row could show one mode while the Transport menu
  showed another ("https" in the menu, "auto" inside): the label was built
  once (stale after in-menu switches) and the two views read DIFFERENT
  sources of truth (the running server's store vs the config file). Now:
  the config FILE is the single source of truth for installer-owned settings
  (transport mode/ports, certMode, certHost) - the setters always write the
  file first and mirror to a running server best-effort, and on any
  divergence the file wins with the mismatch logged (the server re-reads at
  its next restart). The main-menu row re-computes its value every loop.

## 0.9.251

### Fixed
- Windows honours the https TRANSPORT: the Windows branches (prerequisite
  text, [3/6] certificates, tls_phase) hard-coded "no TLS on Windows (plain
  http)" and ignored the stored https mode - an https install warned about a
  missing certificate at step 4 AFTER saying "no TLS needed" at step 3, with
  Outlook unusable for the chosen transport. https on Windows now behaves
  like mac: prerequisites (cryptography/ssl) are checked and displayed,
  Outlook locks while they are unmet, and the certificates phase runs the
  leaf flow (the generator self-installs cryptography).
- the cryptography auto-install: pip --user falls back to a plain install,
  600 s timeout, import-caches invalidated - the first attempt had failed
  silently (its rc never surfaced).

## 0.9.253b (Windows readers + installer whitespace)

### Fixed
- the leaf-consent screen on Windows showed "SHA-256 fingerprint: ?" and
  "expires: unreadable" for a healthy certificate: every certificate READER
  on disk was openssl-based while the Windows leaf is generated by the
  cryptography package (openssl does not exist there). Fingerprints, name/
  expiry and the "is the pair still valid" check now read the certificate
  with cryptography on Windows, and the trust check consults the CURRENT
  USER Trusted People store by thumbprint (so the https mode can actually
  pass its own certificate check there).

### Changed
- installer whitespace + wording: the add-in picker's prerequisite block is
  grouped with blank lines (status / empty line / techniques / empty line /
  the two dim hints), technique lists carry short lines (the auto-install
  parenthetical dropped), the enable note is three short sentences, and the
  picker header is shortened.

## 0.9.249

### Fixed
- Windows: the leaf generator self-installs the `cryptography` package
  (`python -m pip install --user cryptography`, no admin) instead of sending
  the user away mid-flow.
- start_server self-heals a MISSING deploy folder (the WinError 267 crash:
  the transport menu's restart pointed Popen's cwd at a nonexistent
  install dir after an uninstall) by deploying the server first.
- deploy_files takes an explicit add-ins argument (None = the stored
  selection, set() = server-only).

## 0.9.248 (store-deletion hardening, found by the real install->uninstall audit)

### Fixed
- certificate fingerprints were computed from binary DER output through the
  text-capturing runner and produced the sha1 of an EMPTY string - trust
  removal then deleted NOTHING (the removal announced success while the
  keychain item stayed). Fingerprints come from openssl -fingerprint now.
- `security remove-trusted-cert -k` does not exist (the trust removal silently
  failed every time) - removed; when the cert file is gone the USER-domain
  trust record is dropped by rebuilding the exported trust plist without our
  sha1 hash key.
- `security find-certificate ... ~/Library/Keychains/login.keychain-db` (full
  path) silently returns NOTHING - the keychain NAME is required. All keychain
  identification/removal uses the name (with a no-arg fallback) now.
- the render/sideload renewal routed by PLATFORM, not by the stored certMode:
  on mac a leaf-mode install whose pair was missing fell back to CA mode and
  generated + IMPORTED a loopback CA behind a leaf-mode choice. Routing and
  the leaf-reuse guard (meta can no longer "reuse" a nonexistent file) are
  fixed; found by the live audit.
- cert metadata (fingerprints) now lives in the CONFIG dir so an uninstall can
  identify + remove a leaf trust entry even after the certificate files are
  gone; a stray duplicate certMode PUT whitelist that would 400 "byo" is
  removed.

### Verified
- real install (leaf: generated, consent, trusted, sha-verified in the
  keychain) -> uninstall (leaf item + trust record + files + meta gone,
  config kept, certMode reset to leaf, sessions intact) -> reinstall for a
  working leaf-only setup. Zero CA artifacts left after any of it.

## 0.9.247

### Changed
- the DEFAULT certificate mode is now self-signed leaf (no CA): fresh
  installs, any config without a stored mode, and the server-side default all
  resolve to leaf. A deliberately stored ca/byo keeps winning while it is
  still stored.
- an uninstall RESETS the certificate mode to self-signed leaf when the
  config is kept - regardless of what was used before.
- the CA-byo PUT whitelist bug: a stray first certMode block still had the
  pre-BYO whitelist and would have rejected certMode=byo with 400 (removed).

### Fixed
- certificate-store deletion is now robust even when the certificate FILES
  are already gone (manual cleanup, earlier partial uninstall): the cert
  metadata (fingerprints) moved into the config dir so the uninstall can
  identify and remove a lingering LEAF trust entry via the stored sha1, and
  the loopback CA is swept by its (ours-only) common name when ca.crt is
  missing. Verified with a real install -> uninstall cycle plus keychain and
  trust-settings audits.

## 0.9.246

### Fixed
- MailKick date anchoring: asking for a reminder/meeting "today" could land on
  the email's own date or a remembered one (the November 2 instead of November
  6 incident). Three layers:
  1. the model-facing email context carries an explicit TODAY line (weekday +
     date), with the rule that relative words anchor to it, never to the
     email's date;
  2. the prompt rule for new_meeting says the same and forbids past dates;
  3. a deterministic WRONG DATE guard in validateOps re-asks with a
     self-diagnosing reason (dates in the past are always rejected; an
     explicit today/tonight ask must land on today), dbg date.mismatch.
- the all-invalid gate collapse used to hard-code "the paragraphs/targets
  they reference do not match this email" - it now names the ACTUAL reject
  reasons (up to 3, then a count).
- verified E2E (mail simulator + canned mocks): wrong-date reply is rejected
  with the exact date mismatch spelled out; the corrected reply renders the
  meeting gate row dated today.

## 0.9.245

### Added
- Transport options menu (main menu 4) row 6 = certificate mode submenu
  (self-signed leaf / loopback CA / bring-your-own, current mode marked).
  Switching stores the mode; with a working pair it offers an immediate
  conversion (pair wiped, three-way question preselected, consent, server
  restart). Main-menu row 4 shows the current certificate mode.
- the Outlook-enable ask names where the certificate-mode choice happens
  (before step 3 of 6 - certificates) and its default (self-signed leaf);
  the ask itself is a proper block (header + option list + one-line question).

## 0.9.244

### Fixed
- the installer no longer silently re-runs a stored certificate mode when the
  certificates were removed (e.g. an uninstall that kept the config): with NO
  usable pair on disk the three-way question always returns, PRE-SELECTING the
  stored mode (Enter re-runs it) instead of running it silently. Update runs
  with a working pair keep skipping silently.
- the three-way question labels now state the trust effect of each choice:
  the leaf is imported and trusted (NO CA is created), while choosing the CA
  explicitly says the CA is imported + trusted in your keychain.

## 0.9.243

### Changed
- Session files are organized per add-in: `~/.officekick/sessions/<slidekick|
  wordkick|mailkick>/<sid>.json`. Legacy flat files migrate on boot by their
  stored `app` field; app-less (legacy/shared) docs stay at the session root
  and remain visible in every pane. Reads fall back to the root location, so
  sessions written by older builds still resolve. Per-app deletion/backup is
  now just the folder.

## 0.9.242

### Added
- ONE certificate question in the installer (replaces the CA-vs-BYO ask):
  self-signed leaf (directly trusted, renewed consent per re-trust) / local CA +
  CA-signed certificate (mac; one consent, silent renewals) / bring-your-own
  PEM pair. Asked only when no usable pair exists; updates keep the selected
  mode. A stored certMode never re-prompts (leaf/CA/BYO each run directly).
- BYO hostname: optional prompt during BYO setup, stored as the new `certHost`
  server setting (validated DNS name). ALL https URLs in every manifest ride
  that hostname; the server stays loopback-only, so a custom name must resolve
  to 127.0.0.1 (e.g. an /etc/hosts entry). The reminder is printed at setup.
- `certMode` accepts `byo`; `certHost` setting (PUT /api/settings, validated).
- `_api_set`: settings persist through the running server only when it
  CONFIRMS the stored value (older builds silently ignored unknown keys),
  with a verified config-file fallback.
- install.py `--cert-check`: mode-aware diagnostics (older CA-mode leaves get
  informational notes instead of leaf-mode failures; BYO hosts check SAN).

### Fixed
- install.py duplicated function definitions (splice waves) cleaned; every
  top-level function is defined exactly once (`deploy_files` restored).
- BYO copies guard against source == target (SameFileError).

### Changed
- installer `tls_phase` BYO branch validates the pair (SAN, pair match,
  expiry; keychain-trust result is a note - the user attests trust).
- `render_manifest(addin_dir, port, https_port, host)` and
  `validate_existing_tls(crt, host)` are hostname-aware.

## 0.9.242 (uninstall overhaul, later this release)

- uninstall now asks, in order, with BOTH defaults on KEEP (nothing is
  destroyed by an uninstall by default):
  - certificates: remove the local certificates OfficeKick installed (CA/leaf
    trust entry in the user keychain or Windows TrustedPeople, plus the files
    in ~/.officekick/certs); a BYO certificate's own trust entry is never
    touched (only our file copies are deleted, with a warning).
  - config: wipe ~/.officekick (config.json, API keys, profiles, sessions) -
    keeping it means the next install.py reuses the whole setup.
- flags for automation: --remove-certs, --purge-config (pre-existing).
- CA removal helper: user-scope trust settings + login-keychain items matched
  by sha1 against our CA (never by name when the CA file is present), so
  unrelated certificates are untouched.
- WEF-manifest purge now honors --dry-run.
## 0.9.241
### Added (self-signed leaf certificate mode, no CA)
- New server setting `certMode` (`ca` - the existing classic mode - or
  `self_signed_leaf`); PUT-validated, pub()'d, persisted across launches.
- Leaf mode: a unique self-signed leaf per installation (CN=localhost +
  OU=officekick-<install id> for identification, SAN DNS:localhost +
  IP:127.0.0.1 + IP:::1, critical basicConstraints CA:FALSE,
  digitalSignature key usage, serverAuth EKU, SHA-256, RSA 3072; configurable
  validity with renewal). No CA key/certificate is generated in this mode.
- Trust: macOS - per-certificate SSL trust of the leaf in the user login
  keychain (user scope); Windows - the leaf goes to the CURRENT USER
  Trusted People store via certutil (never the Root store; no admin).
  Generation on Windows uses the `cryptography` package; when absent the
  installer prints the exact command and refuses to weaken anything.
- First-run CONSENT block (exact wording, SHA-256 fingerprint + expiry) with
  an explicit yes; declining leaves the trust store untouched with a clear
  https-will-not-work warning. Renewal (expiry/invalidity) requires renewed
  consent; the previous trust is removed on replacement (by fingerprint) and
  during uninstallation - only OUR certificate, identified via
  certs/cert_meta.json (never relying on CN=localhost alone).
- Lifecycle: reuse of a valid leaf (no regeneration per launch), expiry and
  invalidity checks on every install, 30-day renewal window, duplicate
  trust-store entries avoided via fingerprint bookkeeping.
- python3 install.py --cert-check: SAN, CA:FALSE, EKU, key match, validity,
  localhost TLS handshake, platform trust acceptance - and the explicit
  "store trust is NOT Office compatibility" verdict with the documented
  Office verification procedure (docs/certificates.md).
- The Office/WebView2/WKWebView acceptance of per-certificate trust is
  documented as REQUIRES OFFICE TESTING; nothing claims compatibility from
  store-trust success alone.
- tests: 31 (leaf lifecycle suite + certMode round-trip).

## 0.9.240
### Fixed (Windows secret protection)
- `os.chmod(0600)` is a no-op stub on Windows (only the read-only bit), so
  config.json and any certificate files were effectively unprotected there.
  New `_protect_secret()`: chmod 0o600 on macOS/POSIX; on Windows NTFS ACL
  hardening via icacls (`/inheritance:r` to strip inherited ACEs, then
  `/grant:r <current user>:F` so only the owning user keeps access). Applied
  to config.json (API keys), the certs directory and every private key
  (server.key/ca.key). Order-safe with `relax_windows_config_acl` (grants
  the same user only).

### Added (configurable listener ports + menu renumbering)
- Settings `transportHttpPort` (default 8765) and `transportHttpsPort`
  (default 8766): whitelisted, range-validated (1024-65535, must differ),
  pub()'d; the server binds the http listener from the setting (an explicit
  `--port` flag keeps precedence) and the TLS listener from the https port.
- The Transport menu (main menu 4) now also offers "set the http/https
  port": prompts, validates, persists, re-renders the manifests with the new
  URLs and restarts the server. `--transport` keeps working for the mode.
- The installer resolves the current ports at startup (explicit --port wins)
  and uses them in the banner, status, server menu, and every https URL.
- Main menu renumbering: 1 Install / 2 Update / 3 Server - blank - 4
  Transport options - blank - 0 exit. Uninstall moved into the Server
  submenu (option 5, own group) so the main menu stays options/action/back.

## 0.9.239
### Added (transport mode - http/https/auto, persisted)
- New server setting `transportMode`: `auto` (the classic behaviour - serve
  https://localhost:8766 whenever a usable loopback certificate exists), `http`
  (plain http even when certificates exist), `https` (require the TLS listener;
  the installer creates/renews the certificate). Settable via Settings PUT
  (whitelisted + validated; invalid values are refused) and from the new
  installer menu 5) Transport or `--transport auto|http|https`.
- The installer's manifest render respects the stored mode: `http` renders
  plain-http manifests (MailKick is skipped with a note - Outlook requires
  https), `https` creates/renews the loopback pair even without the Outlook
  add-in. Updates keep the stored mode (they never flip it).
- The certificates phase prints CA/certificate details on the console after
  it runs (subject CN, expiry, days left, the file paths under
  ~/.officekick/certs) for reused, reissued and freshly generated pairs; the
  pair check runs on EVERY install/update whenever TLS is in use. When only
  the leaf is bad/expiring (< 30 days) the leaf is re-issued from the SAME
  already-trusted CA (no new trust import); a fresh CA is generated only when
  the CA itself is unusable.

### Changed (installer UX - install.py redesigned)
- Menu-driven installer: banner (version, Created by Corelan Consulting bv,
  repo link, detected platform/install dir/server + add-in status), main
  menu (Install / Update / Server submenu / Uninstall), and an add-in
  selection step: Word + PowerPoint ON by default, Outlook OFF by default.
- Outlook cannot be enabled unless the https prerequisites check passes
  (openssl, the macOS security tool, python ssl); enabling shows a clear
  warning that a self-signed loopback CA + localhost certificate will be
  created and imported into the keychains, then asks to confirm.
- Non-interactive `--install` now installs Word + PowerPoint; Outlook only
  with the new `--mailkick` flag, with --mailkick-cert/--mailkick-key, or
  when Outlook was already installed (updates preserve it). --skip-mailkick
  still forces it out.
- Structured output: one step line per phase (preparing / deploying /
  certificates / registering with Office / autostart / starting the
  server); the full detail trail (file lists, tool output, decisions,
  popup essays) goes to install.log next to install.py (gitignored).
- Automation flags unchanged: --update/--start/--stop/--restart/--status/
  --port/--ribbon/--dry-run all keep their behavior; --status gained the
  add-in summary line. Fixed: a latent NameError in the old mailkick skip
  path (wef_dir was not a function).

### Docs
- README/architecture/add-in docs now describe BOTH transports everywhere
  (http://localhost:8765 always on; one trusted https://localhost:8766
  listener for all three add-ins on macOS once a loopback certificate
  exists); the old 'MailKick-only exception' framing in README and the
  'default plain-http' phrasing are gone; simulator URLs note the https
  twin.

### Fixed (MailKick)
- The response chain never escaped the repair ladder on a clean parse: every
  MailKick run re-asked the model at least once (silently doubling cost and
  latency on every reply since the 0.9.217 repair wave). Added the missing
  "clean parse returns immediately" escape.

## 0.9.234
### Fixed (SlideKick)
- Real-Office textbox creation: PowerPoint 16.114 (mac) rejected
  `shapes.addTextBox(text, options)` with RichApi InvalidArgument at sync for
  EVERY textbox op of an infographic run while identical-args
  addGeometricShape calls succeeded in the same batch. create_shape textbox
  ops now ladder across three creation paths (each in its own run): native
  addTextBox(text+options), the docs-sample form (text only, geometry as
  shape properties), and an emulated no-fill/no-outline rectangle with the
  text written through the text frame (1.5). Failed attempts log
  `office.createShape.attempt`, the served path logs
  `office.createShape.fallback`; other shape kinds keep the single attempt.

## 0.9.233
### Added (MailKick)
- New setting "create a new session when opening a new email" (Settings >
  General > Sessions, off by default): when a different email is opened, the
  pane starts a clean session for it instead of continuing the open
  conversation. Real Outlook detects the switch through the mailbox
  SelectedItemsChanged event (Mailbox 1.13); every run and the refresh
  button re-check as a fallback. A boot-resumed session that belongs to a
  different email is dropped. Sessions with no requests are never saved -
  and the previous conversation stays in History (resumable manually).
- The simulator now serves three distinct emails ("open email" buttons in
  the Sim email panel), each with its own itemId, to exercise the setting
  E2E.
## 0.9.232
### Changed (MailKick)
- Button layout: Run / Cancel / timer on the first row; the quick-reply
  buttons (reply +, reply −, propose) live on their own row BELOW it.
## 0.9.231
### Changed (MailKick)
- The "propose" quick action no longer prescribes 2-3 replies: the model
  decides itself how many and which reply proposals would genuinely help
  (a question, a positive/negative answer, a decision, or an alternative
  it judges better - if only one reply is natural, it proposes one).
  Each proposal stays its own draft_reply action in the approval list.
## 0.9.230
### Added (MailKick)
- New quick action "propose" next to reply +/reply -: the AI analyzes the
  open conversation and proposes the 2-3 most natural replies (a
  question, a positive/negative answer, a decision - different in kind,
  not just wording). Each proposal arrives as its own draft_reply action
  so they compare side-by-side in the approval list. Nothing is sent -
  the usual gate applies.
### Changed (MailKick)
- The button row layout is stable: Cancel reserves its space while
  hidden and the timer has a fixed slot, so the Run/reply row no longer
  jumps when a run starts.
## 0.9.229
### Changed
- GUI declutter: the Web search steps sheet's three-sentence explainer is
  now two short lines (order + the free-DuckDuckGo default) with a
  pointer to the docs. The full semantics (fall-through behavior,
  per-engine key requirements, "model lookup" steps per connection type)
  moved to docs/settings.md > "Web search steps".
## 0.9.228
### Changed
- Image quality: every image prompt is now enriched at the single
  server-side /api/image choke point (all add-ins) with a fill-the-canvas
  directive: content should fill the whole image, no large whitespace
  margins, tight full-bleed composition. Skipped when the prompt already
  mentions whitespace so refinements are never doubled up.
## 0.9.227
### Fixed
- Anchored inserts no longer split sentences: every anchor resolution
  (insert_image, insert_paragraphs, move_text/copy_text blocks,
  move_image) expands the matched range to the WHOLE paragraph that
  contains it before inserting "After" - the inserted content lands after
  the paragraph's last line, never mid-sentence.
## 0.9.226
### Added (DocKick)
- Images are now VISIBLE to the model: the document outline marks image
  paragraphs as [image xN], and the context exposes the total inline
  picture count.
- move_image gains image:N addressing - relocate the Nth inline picture
  directly (image-only paragraphs are deleted after the move; images
  sharing a paragraph with text are COPIED with the original kept and
  reported). The model can now actually perform "move the images to the
  right place" instead of only moving caption text.
## 0.9.225
### Fixed
- THINKING BUDGET EXHAUSTION (root cause of the empty-replies / silent
  "Done with 0 actions" on Ollama cloud models): glm-5.3-flash:cloud puts
  its chain-of-thought in the "thinking" field, which is counted against
  num_predict. With the pane's maxTokens 8192 the thinking consumed the
  entire budget (done_reason "length", content empty), so the server
  returned an empty reply while direct raw curls succeeded. The ollama
  path now detects exactly that signature (length + empty content +
  non-empty thinking) and retries ONCE with think:false (verified live:
  valid move_text ops, done_reason stop). Reproduced with the exact
  logged request on the box: with options = empty + 8192 eval; with
  think:false = 7884 eval + content.
## 0.9.224
### Changed (DocKick)
- A "the model proposed 0 actions" run is no longer a bare "Done": the
  pane now shows the model's actual reply text (what it chose to say
  instead of acting) plus a one-click "Ask again (same request)" button,
  so a silent success with no visible result is readable and repeatable.
## 0.9.223
### Fixed
- move_text/copy_text still crashed with the style-proxy error on the
  box even with the collection loaded (0.9.222): hardened BOTH directions
  - each block paragraph is individually re-loaded ("text, style") in
  _blockSlice, and every style read/write in _insertBlockAfter is
  try-guarded, so a single un-materialized property can no longer abort
  the whole move (style simply falls back to keeping the anchor's
  formatting on that paragraph) and the op completes.
## 0.9.222
### Fixed
- move_text/copy_text applied but failed at runtime for all 6 actions
  ("The property 'style' is not available... call the load method"):
  _pasteBlock loaded paragraphs as "text" only while _insertBlockAfter
  reads each paragraph's style to preserve it. loads now "text, style"
  (both _pasteBlock and moveImage); audited all remaining load sites -
  no other unloaded-property reads on that path.
## 0.9.221
### Changed (all panes)
- Re-ask failures now log their FULL stack trace into the pane's
  activity log ("RE-ASK FAILED: ..." / "RE-ASK RENDER FAILED: ...")
  instead of only a status line, so the actual broken statement is
  visible and fixable in one round-trip instead of guesswork.
## 0.9.220
### Added (DocKick / Word)
- move_image {find, place}: relocates an inline image (or every image in
  the same paragraph) to sit right after `place` (an anchor text), or at
  the end. find = text inside/near the image's paragraph. Implementation:
  read the picture sources with getBase64ImageSrc (WordApi 1.6), insert
  at the target, then delete the source paragraph; anchor-not-found
  falls back to the end and says so. Registered in OPS_META + validated +
  previewed; schema line added so the model can use it.
## 0.9.219
### Fixed
- THE reason "All 6 action(s) were rejected": move_text and copy_text
  were missing from OPS_META, so OP_TYPES rejected them as "unknown op
  type" while the model prompt ADVERTISED them - the model could only
  ever emit invalid ops. Registered in OPS_META (20 types now); audited
  that every op advertised in the model schema exists in OP_TYPES for
  all three panes (Docs 20, Slides 23, Mail 7).
## 0.9.218
### Fixed
- DocKick still crashed the validation re-ask with "instruction is not
  defined" (0.9.217 fixed Slide/Mail only - the Word block lives in
  renderOpsPreview without an instruction param too). The recursion now
  uses SK.lastInstruction in all three panes; audited by regex that no
  bare reference remains.
## 0.9.217
### Changed (MailKick)
- The same "don't reject fast" treatment as 0.9.216 for MailKick: when
  every proposed action fails validation, the pane re-asks the model with
  the exact rejection reasons (up to 2 rounds) before the manual "Ask
  again" fallback. Scope fix for the recursion (SK.lastInstruction) so
  the re-ask works from the ops-preview context.
## 0.9.216
### Changed (DocKick + SlideKick)
- "Don't reject fast": when EVERY proposed action fails validation, the
  pane now re-asks the model up to 2 extra rounds WITH the exact
  per-op rejection reasons included in the prompt, before falling back to
  the manual "Ask again" box. Status shows "re-asking the model with the
  reasons (round N/2)…". This applies to Word documents and slide decks
  alike; the syntax-repair rounds (clean JSON) are unchanged and still
  run first.
## 0.9.215
### Fixed
- "None of the N action(s) survived validation" no longer guesses ("the
  paragraphs/targets they reference do not match this document"): the
  box now lists EACH rejected op with its ACTUAL validation reason (and
  its place target) so the fix is obvious instead of hidden. The status
  line after this path now ends the busy spinner ("Rejected - see the
  reasons above") instead of leaving it running.
## 0.9.214
### Fixed
- anchor-based inserts crashed with "The property 'items' is not
  available... call the load method": the Office-JS RangeCollection from
  body.search must be loaded ("items") before sync/reading .items. All
  three anchor sites fixed: insert_paragraphs, insert_image, and the
  move_text/copy_text block paste. Retries re-failed by design, so a
  whole batch of anchored inserts produced only errors; now they insert.
## 0.9.213
### Added (SlideKick / PowerPoint)
- copy_shape {slide, shape, to, left/top/width/height}: duplicates a shape
  (picture, text box, table - full fidelity via the Office-JS shape
  copy/paste pair, PowerPointApi 1.13) onto slide `to` (or an omitted
  target = the same slide), with optional re-placement.
- move_shape_to (same fields): CUT+paste - relocates the shape and removes
  the original; the Word-block cut/copy idiom applied to PowerPoint
  objects. Approvals: copy = auto, move = confirm.
- Model schema + preview strings updated; validation reuses the existing
  slide/shape context checks. move_shape (reposition/resize) unchanged.
## 0.9.212
### Added (DocKick / Word)
- Cut/paste-grade block ops: move_text now accepts `until` (move from the
  paragraph containing `find` through the one containing `until`) or
  `count` (N consecutive paragraphs); new op copy_text = the same
  find/until/count/place block semantics without deleting the original
  (paste-copy). Styles preserved; anchor-not-found falls back to the end
  and says so. Both default to user confirmation like other destructive
  ops. PowerPoint is a different object model: shape repositioning
  (move_shape) already exists there; text-flow cut/paste does not apply.
## 0.9.211
### Added
- insert_paragraphs accepts the same `place` vocabulary as insert_image
  ("cursor" at the selection, "end", or an ANCHOR text to insert after
  the first body match - end-fallback reported). This is exactly what
  text/image placeholders need: the model can now insert placeholder
  paragraphs at the right positions instead of only appending.
- New op move_text {find, place}: relocates the paragraph containing
  `find` to sit right after the `place` anchor (copy -> insert after ->
  delete source, style preserved; end-fallback). That covers the
  "move them to the right position yourself" ask.
- Placement guidance updated so the model uses anchors for placeholders.
### Note
- The 19:0x failures in the user session: the model was asked to place
  placeholders mid-document while every text op could only append - it
  had no workable op set and collapsed into empty/unclear JSON.
## 0.9.210
### Fixed
- "Cannot set properties of undefined (setting 'place')": the 0.9.209
  insert_image validator copied op.place into `out` before `out` was
  declared. Order corrected (0.9.209 shipped the crash).
### Changed
- Chat working memory raised in all three panes: the model now sees the
  last 16 exchanges (32 entries, was 6) with up to 10k characters per
  stored message (was 8k). Resume notes no longer hardcode the old count.
## 0.9.209
### Added
- The model now decides where insert_image lands: the op accepts
  "place": "cursor" (caret/selection), "end" (document end), or an ANCHOR
  text - the image inserts immediately after the first body match
  (Word-JS body.search -> range "After"; if the anchor is not found the
  image goes to the end and the chat line says so). The chat line reports
  the actual placement and whether the model or the pane setting chose
  it. The pane's "insert at cursor" checkbox remains the default for ops
  WITHOUT a place value. Schema/placement guidance updated so the model
  always sets place when it can.
## 0.9.208
### Fixed
- DocKick image insertion defaulted to the END of the document: the
  "insert at cursor" checkbox defaulted to OFF, and WordDeck.addImage sent
  every generated image through the body "End" fallback. At-cursor
  (caret/selection) insert is now the DEFAULT (checkbox pre-ticked +
  addImage treats a missing flag as at-cursor); unticking still appends
  at the end.
## 0.9.207
### Fixed
- About tab "Add-in identity (this manifest)" showed "?" whenever the
  server ran with a per-addin static dir instead of the install root:
  /api/config read <static_dir>/manifest/*.xml only. The lookup now also
  tries <static_dir parent>/manifest, so the identity version is found
  regardless of how the server was launched.
## 0.9.206
### Fixed
- Update notice is SHORT (one line + the button): "Click 'Pull update
  from git now' below, or run install.py --update ... and reload the
  pane." The old pane gave a paragraph.
- The 'Pull update from git now' button now exists in ALL three panes:
  it was only ever in the SlideKick task pane, so Word/Outlook panes
  showed only text even on builds that had the button for months - the
  reason an 0.9.200 install still could not click-pull from those panes.
### Note
- The button still requires the installer-recorded clone probe
  (repoPath -> git_update_probe) to mark pullAvailable; without it the
  notice says to run install.py --update.
## 0.9.205
### Changed
- The installer's user-facing output is scheme-truthful now: the transport
  banner, the "add-ins talk to a local server" note and the simulator URL
  print BOTH listeners (with each pane's actual scheme) whenever a working
  loopback certificate exists, and plain http when it does not. New
  helpers: tls_active() (validated cert present), pane_listener_line()
  and the simulator base switch.
## 0.9.204
### Changed
- When a loopback certificate is present AND working (expiry, localhost
  SAN and trust validated), the installer activates it for Word and
  PowerPoint as well: all three manifests point at the shared
  https://localhost:8766 listener. Without a usable certificate the
  add-ins stay on http://localhost:8765 (MailKick on macOS is skipped as
  before; on Windows MailKick without a cert renders http - nothing can
  listen on 8766 there).
### Fixed
- render_and_sideload used tls_pair before assigning it (UnboundLocal
  error in the step label) and only checked file existence, not
  usability; the mailkick manifest template hardcoded port 8765 so the
  pane URL could end up as https://localhost:8765 (nothing listens
  there). Both replaced by a validated, order-safe tls-pair check and a
  __PORT__ normalisation.
## 0.9.203
### Changed
- Option 1 (auto TLS) is now reuse-first: when a certificate is already on
  disk it is validated (unexpired, localhost SAN, chain trusted via
  security verify-cert) and REUSED - no new CA, no regeneration, no
  re-import. Only a failed check triggers a regeneration, with the reason
  printed. Verified on this Mac: the existing cert reuses; an expired cert
  is detected and regenerated.
### Docs
- docs/mailkick.md Option 1 describes the reuse-first behavior.
## 0.9.202
### Fixed
- Mode-vs-https correction (verified by Corelan on the real machine): the
  MailKick pane installs from BOTH New and classic macOS Outlook - https
  + a trusted certificate was the actual gate, not the Outlook mode. The
  installer's mode prompt, README, and docs/mailkick.md (requirements +
  the old "switch to classic" troubleshooting entry) now say so; the
  stale New-Outlook limitation claims are gone.
## 0.9.201
### Changed
- Boot self-update throttle: 12 h -> 5 min. Urgent updates need a restart,
  not a wait; the 5 min window still keeps pathological restart loops from
  hammering GitHub (failed checks are never stamped and retry on the next
  boot, as before).
## 0.9.200
### Fixed
- The update notice no longer tells an OLD pane to "click the update
  notice": the click-to-pull button exists only in newer panes. The
  message now says the pane cannot update itself while it IS the old
  build, and points at install.py --update for that first hop.
## 0.9.199
### Added
- MailKick-for-Outlook install modes on macOS, chosen interactively during
  install (and non-interactively via flags):
  1) the default self-signed route (per-user loopback CA generated +
     imported into the trusted keychains);
  2) bring your own already-trusted certificate: `--mailkick-cert` +
     `--mailkick-key` - the installer validates PEM/unencrypted key,
     subjectAltName DNS:localhost, serverAuth EKU and macOS trust
     (security verify-cert) and installs nothing if a check fails;
  3) `--skip-mailkick`: no https, no certificate, no keychain import;
     any previously staged MailKick manifest is removed from Outlook.
  Prerequisite failures keep skipping MailKick for Outlook automatically;
  SlideKick/DocKick are unaffected in every mode.
### Docs
- docs/mailkick.md: "Your three options" - what the default self-signed
  route does (and why the generated CA must be imported into the trusted
  keychains), the exact compliance list for a user-provided certificate,
  and the skip behavior; the README install section now summarises the
  three routes.
## 0.9.198
### Docs
- README: the transport story now covers both schemes - Word/PowerPoint on
  http://localhost:8765, MailKick on macOS over https://localhost:8766
  (loopback TLS, per-user CA) - with a per-add-in port/scheme table, how to
  read the pane's bottom line (the lock + port) and how to check a manifest
  on disk. The installer's transport banner states the same split.
## 0.9.197
### Changed
- Scope correction: only MailKick needs https URLs (Outlook's manifest
  validation). SlideKick and DocKick manifest icons are back on the local
  http server (offline-safe, fully local as before); their pane URLs were
  always http and stay that way. The TLS sibling listener + CA remain
  macOS-only and are used exclusively by MailKick's manifest.
## 0.9.196
### Changed
- **MailKick for Outlook is now prerequisite-gated on macOS.** The
  installer documents the requirements (classic Outlook mode - New Outlook
  cannot run desktop task-pane add-ins; openssl for the per-user loopback
  CA + certificate; the one-time CA trust in the login keychain; port
  8766). When openssl is missing or the CA trust is declined, MailKick is
  SKIPPED for Outlook with a clear message instead of staging a manifest
  that cannot run; SlideKick and DocKick install normally and the
  prerequisites are re-checked on every install.
### Docs
- docs/mailkick.md: a "Requirements for Outlook (macOS) - read before
  installing" section listing what is needed, why, how to inspect/revoke
  the certificate, and how to retry later; the README points at it.
## 0.9.195
### Fixed
- New Outlook rejected the loopback certificate ("not signed by a valid
  security certificate") - macOS Certificate Transparency policy deems any
  cert issued after ~2019 without SCTs untrusted for strict webviews, even
  under a user-trusted root. The server leaf certificate is now issued with
  a backdated notBefore (Jan 2018 -> Jan 2038) via `openssl ca`, which
  falls under the cutoff and verifies successfully; verify-cert passes and
  the pane loads over https://localhost:8766.
## 0.9.194
### Fixed
- MailKick pane died at boot ("parsed is not defined", caught by the new
  pane-error log on the first real https://localhost:8766 load): a stray
  orphaned statement left at the top level of mailpane.js by an earlier
  splice. Removed; the file was already deployed with the fix.
## 0.9.193
### Fixed
- **MailKick was rejected by Outlook's Add-from-File AND the wef-folder
  pickup: this Outlook build validates EVERY manifest URL as https,
  including the pane URL** ("The value of the SourceLocation attribute ...
  contains an unsupported URL"). The server now serves the same pane + APIs
  additionally over **loopback TLS on port 8766** (same process; the cert
  pair lives in ~/.officekick/certs). The macOS installer generates a
  per-user "OfficeKick Loopback CA" + a localhost certificate with
  openssl, trusts the CA in your login keychain (no admin rights; macOS may
  show a confirmation), and MailKick's manifest points at
  https://localhost:8766. SlideKick/DocKick keep their plain-http route.
### Docs
- mailkick.md: the https listener section + what the certificates are for.
## 0.9.192
### Fixed
- Outlook's Add-from-File dialog validates every manifest URL as https and
  rejected the install with "The value of the IconUrl attribute ... contains
  an unsupported URL". The manifest templates now point
  IconUrl/HighResolutionIconUrl and the ribbon icon resources at the
  project's public GitHub assets (fetched once at install, cached by
  Office); the task-pane/pane URLs stay on the local server. A re-render is
  picked up by the next install; the already-deployed MailKick manifest on
  a machine can be patched in place without quitting Office.
## 0.9.191
### Docs
- MailKick troubleshooting for a likely silent cause of "the add-in no
  longer works": Outlook running in "New Outlook" mode cannot run desktop
  task-pane add-ins (the Add-from-File dialog fails even with a correct
  manifest) - switch New Outlook off and relaunch. Also documents where
  Get Add-ins lives (behind the three-dot toolbar button on macOS).
### Changed
- The installer's WEF-clear note now mentions the New Outlook check too.
## 0.9.190
### Changed
- The installer's add-in cache/WEF clear now warns that macOS Outlook holds
  MailKick as a manually-added add-in: after a clear, re-add it once via
  Get Add-ins > My add-ins > Custom Add-ins > Add from File (the wef
  manifest keeps being re-staged automatically).
## 0.9.189
### Fixed
- **Reading a slide quoted cut text**: in large decks (>40 slides) the
  outline only carries a 900-char preview per shape, and a model asked to
  READ slide 60 answered by quoting that preview - bullet text ended with
  an ellipsis. The prompt now requires the model to reply read_slide FIRST
  for any read/quote/show request, and the pane itself fetches the FULL
  slide text before the first ask when the instruction reads/quotes slides
  (the shared 2-read budget; errors fall through to a normal run).
- **Long answers were truncated in History and resume**: every saved
  exchange was capped at 2000 chars, so a multi-thousand-character reply
  came back chopped on reopen and lost most of its content in later model
  context. Raised to 8000 per message in all three add-ins.
## 0.9.188
### Docs
- Full documentation review pass: README shows the new installer menu
  (1/2/3/x/0); SlideKick's action table carries every current op with its
  default policy (shapes, duplicate/clear/picture-fill, formatting), the
  web-lookups section documents the always-before-ask lookup for content
  builds, sessions document the current-session double-confirm removal,
  and the dev-mode command uses the correct config PATH
  (~/.officekick/config.json). DocKick: repaired a mangled intro, the
  missing create-TOC row, the typography law (Normal body / headings /
  28pt cap), the always-lookup note, current image-model wording and a
  stale limitation. MailKick: repaired a mangled intro and documents its
  reactive-only lookup. settings.md: the current defaults list (incl. the
  0.9.175 actions) and the General note that a Check-now result clears on
  re-open. architecture.md: http-only diagram, current endpoint list,
  per-pane op lists, sessions behavior, the rewritten web-lookups and
  updates sections (boot self-update included).
## 0.9.187
### Changed
- The installer's interactive menu is numbered 1 / 2 / 3 (Install / Update /
  Start-restart), Uninstall is `x` and `0` exits instead of the old
  1/2/3-uninstall + "2b" arrangement. README menu wording follows.
## 0.9.186
### Added
- **Content-writing runs always check the web first** (standing rule: the
  lookup channel is internal + free): when a SlideKick instruction asks for
  deck CONTENT (write/create/build slides about ...) or a DocKick
  instruction asks for document content (essay/article/report/letter/plan
  ...), the pane runs the search ladder + its keyless DuckDuckGo fallback
  BEFORE the first model ask and feeds the fetched results into it, so
  facts get verified and cited. It uses the same 3-lookup budget (2
  reactive rounds remain for the model) and is skipped for pure edit /
  reword / styled requests. MailKick keeps the reactive lookup (drafts are
  anchored to the thread).
### Fixed
- **DocKick essays came out in huge type**: the model had styled running
  text with title/heading styles or giant font sizes. Now guarded
  mechanically + taught in the prompt: body paragraphs = Normal style
  (title = ONE title paragraph, heading1/2 = headings), font.size > 28pt
  rejected with an explanation, and a batch where EVERY inserted paragraph
  (4+) is title/heading1 gets sent back to the model with the style law.
## 0.9.185
### Changed
- The profile editor's "show all models" is now a CHECKBOX (image slot +
  web slot): unchecked (default) lists only image- / web-capable models,
  checking it refreshes the select with the full list. State persists
  across profile edits. (Was a toggling button.)
## 0.9.184
### Fixed
- The bottom status bar (and the About tab) refreshed their version only
  at pane boot: when the server updated under an open pane, "Check now"
  truthfully said up to date while the bottom indicator still showed the
  older build. The 5 s server heartbeat now detects the served version
  change and repaints the banner + About tab without a pane reload.
## 0.9.183
### Added
- **Pane error log (timestamped)**: every uncaught task-pane error -
  window.onerror / unhandledrejection, including the ones that kill the
  boot wiring like the 0.9.177 regression - is POSTed to the local server
  and appended to ~/.officekick/logs/pane-errors-YYYY-MM-DD.log (one JSON
  line: timestamp, kind, message, stack, pane path, server version, UA).
  Always on, not behind the debug toggle; a 4 s dedupe avoids storming.
  Files clean up with the other logs (logCleanDays).
- **install.py surfaces it**: --status / --start / install+update output
  print a warning with the file path and the last entry when panes crashed
  TODAY, so a broken pane on the desktop points straight at its paper
  trail.
### Fixed
- **Deleting the CURRENT session**: the History tab's Delete button now
  double-confirms ("Delete CURRENT session?" -> "Really remove the
  CURRENT session?" + a status explanation) and, when confirmed, wipes
  the session's traces from the Sessions tab (chat output, previous
  requests, session label/resume key) and resets to a fresh session.
  Deleting a non-current session keeps its single confirm.
## 0.9.181
### Fixed
- **Panes dead at "starting..." with empty Connections/Profiles and the
  default font (0.9.177 regression)**: the new "Update the server
  automatically before it starts" checkbox was populated where the config
  variable was out of scope; the ReferenceError aborted the event wiring
  mid-boot on every platform. Population now happens inside the config
  callback where it exists.
- The Windows start-hint went null (unreachable return inside the macOS
  branch) - the hint text is back for every platform.
## 0.9.180
### Fixed
- The per-manifest result lines ("SlideKick manifest written: ...") used a
  one-space [ok] prefix while everything else uses three - they align with
  the rest of the installer output now.
## 0.9.179
### Fixed
- Installer summary: the per-app lines ("SlideKick installed for
  PowerPoint" / "DocKick installed for Word" / "MailKick installed for
  Outlook") no longer pad the names to a fixed column - they align with
  every other status line now.
## 0.9.178
### Added
- `install.py --restart` - stop the local server, then start it again (the
  same stop/start the menu's "2b" option ran).
### Fixed
- `install.py --start` now says so when the running server is an older
  build than the checkout ("run install.py --update").
## 0.9.177
### Added
- **Update before serving**: at boot the server compares the released
  VERSION with the running one (once per 12 h) and, when a newer release
  exists, updates the install BEFORE it starts serving - a git clone is
  pulled (`git pull --ff-only`), any other install downloads the GitHub
  repo zip (codeload). server.py + the add-in static folders are
  re-deployed and the process hands off to the OS-native restart helper,
  so the running server is the new release by the time it answers.
  Office and manifests are never touched by the boot update - run
  `install.py --update` when the manifests changed. Failures never block
  the boot: offline or a failed pull just serves the current code and
  retries on the next start.
- The new setting is Settings > General: "Update the server automatically
  before it starts" (default on), stored as `serverAutoUpdate`.
  `OFFICEKICK_NO_BOOT_UPDATE=1` also skips it for one run.
## 0.9.176
### Docs
- README "Use it": current tab names (the prompt sheet is the Sessions
  tab), both optional profile slots, the Search step, and direct links
  into the matching sections of docs/settings.md; "Check now" points at
  Settings > General.
## 0.9.175
### Fixed
- Saving the approval policy failed with "Unknown action type:
  duplicate_slide": the server's action whitelist was missing ops the
  policy tables render. Added duplicate_slide (auto), copy slide over
  (confirm), clear slide (confirm), picture-fill shape (confirm), create
  table of contents (auto), reject tracked changes (confirm), set header
  / set footer (confirm). Removed two duplicated literal entries.
  Existing stored policies keep their values; the new actions adopt
  their default modes. De-duping also re-lost "format_selection" and
  "add_hyperlink" from the default table - restored (confirm / auto).
## 0.9.174
### Fixed
- Activity log label for Model lookup steps said the model name twice
  (":online:online") - the step's model already carries the :online
  suffix; it is shown once now.
### Added
- Search steps: drag & drop reordering with a grip handle (the dot
  column before each row); a drop indicator line shows where the step
  lands. Save search steps persists the new order.
### Docs
- New docs/settings.md: every setting explained - what it means and how
  to configure it (General, Approval policy, Connections, Profiles,
  Search, the activity log, where config/logs live).
- The three add-in docs now point at it and keep only their own
  specifics; the Search tab and web-lookup behavior are documented.
## 0.9.173
### Fixed
- The server's VERSION stamp rides every release again: the 0.9.172
  wave shipped pane changes only and left /api/health (and the About
  tab) at 0.9.171.
- Update-check freshness: the probe result is cached for 2 minutes
  (was 10), so a just-pushed release is not reported as "up to date"
  from a stale cache entry.
## 0.9.172
### Fixed
- Validation: "find text not found" now checks what EARLIER ops in the
  same batch left in the shape/paragraph. A batch like replace_text ->
  format_text (find: a fragment of the NEW text) used to fail both
  format ops although the text exists by apply time. Trackers live at
  validator scope; replace/edit_paragraph update them. Word documents
  carry the same fix (edit_paragraph / replace_text / format find).
## 0.9.171
### Added
- read_page continuations: a fetched page carries only its first 8000
  chars. The pane's follow-up now says exactly how much of the page was
  shown and teaches the model to ask for the NEXT part
  ({"kind":"read_page","url":"...","from":<n>}). /api/fetch-url serves
  later slices (offset=from), returns {offset,total,more,next}, and
  continuations do not use the 4-page-read budget (up to 6 per run).
  All three add-ins.
## 0.9.170
### Fixed
- format_text / format_text-substring on a shape with no text (table,
  chart, picture) now fail with a clear message naming the shape type,
  instead of a bare PowerPoint "InvalidArgument" (retrying it never
  helps). The model prompt says the same rule.
- Validation: format/move ops may now target shapes the SAME batch has
  just added to an existing slide (create_table/create_shape/
  create_image) - a run failed all of them with "shape 4 not in the
  context list for slide 113" although the shapes existed after the
  batch created them.
### Changed
- Search steps: a step that cannot run is skipped and reported ("no API
  key set", "no model pinned", "no native web search") - when every
  step is unconfigured the free DuckDuckGo engine still runs, so a
  keyless Brave + modelless OpenRouter sequence no longer dead-ends.
  The activity log and the give-up reply list skipped steps.
- Activity box shows a fixed 8-row window (newest on top, older lines
  scroll below).
- Simulator parity: sim deck failures carry the friendly-error marker,
  so a failing op reports and retries like a real host instead of
  aborting the remaining batch.
## 0.9.169
### Added
- Run activity log: while a run works, a box under the status line lists
  what the add-in is doing and has done, most recent on top - "Sending
  request to <connection> (model)", "Web search via <engine>", "Model
  lookup via <connection> (model <model>:online)", plus the status trail
  (lookups, page reads, repairs, applies). Cleared on every Run press.
  All three add-ins.
## 0.9.168
### Changed
- Settings > General: the retired "Web search (all add-ins)" group is gone
  (leftover engine/key controls from the pre-Search-tab era). Web search
  lives only in Settings > Search: the steps sequence + the per-engine
  API keys.
## 0.9.167
### Changed
- Search tab: the key-state labels are colored - "key set" green,
  "no key" red (both themes); typing keeps the neutral state.
- One "+ add step" button. Every step row carries its own type select
  (Search engine / Model lookup); switching the type refreshes the row's
  selects and normalizes the fields. Fixed: switching the type used to
  rebuild only the editor and wiped the row's numbering, type select and
  remove button.
- Intro copy updated to the 0.9.166 capability rules (OpenRouter step
  skips itself when no model is pinned).
## 0.9.166
### Fixed
- ":online is not a valid model ID": a Model lookup step on OpenRouter
  with "(connection default)" (no model pinned) sent a bare ":online".
  Now the step skips itself (no search model = no search) and the editor
  hint says to pin a model. Also: provider HTTP 4xx errors (message-carried
  status) are no longer retried 5x10 s - only 429/5xx/network retry.
## 0.9.165
### Changed
- Settings -> Search wording + keys: "Search engine" / "Model lookup" step
  types, and per-engine API keys live in the Search tab (Brave / Tavily /
  Firecrawl, password fields, "type to replace, blank to clear", masked
  server-side as before). The engine dropdown steps all key their engines
  from these fields.
## 0.9.164
### Fixed
- "try again, you should have the ability to do a web search now" ended in
  "(model returned an empty action list)": a bare web_lookup intent that
  reached renderParsed after a no-results round ended the run silently.
  Now: a bare intent (no web results yet) triggers the lookup from the
  runner too; intents that carry ops keep them (stash pendingWebOps); an
  empty action list AFTER a lookup applies the stashed ops; and the empty
  box now says when a lookup ran but the model chose no deck change.
## 0.9.163
### Changed
- ONE search sequence for all add-ins (Settings -> Search): no more named
  profiles, list, or "use this profile" - set the steps, save, done. The
  old named profiles + selection migrate to the shared step list on
  startup; no steps = the free DuckDuckGo engine.
### Added
- Internal page reads: after a lookup the model may reply with
  {"kind":"read_page","url":"..."} to get a specific page's FULL text
  (free, never gated, up to 4 reads per run; the server's SSRF-guarded
  fetcher does the reading). The auto-gather still includes the top
  result pages; read_page = the explicit follow-up the model asks for.
## 0.9.162
### Fixed
- {"kind":"web_lookup"} intents never parsed (tryParse had no web_lookup
  branch) so the repair chain answered with a format nudge instead of
  running the lookup - now: the intent parses, the lookup runs (profile
  ladder; internal, never gated), and the model is re-asked with the
  results (verified live: the search fired, the model answered with the
  found sources and skipped the title change rather than guessing).
- up to 3 lookup rounds per run; a give-up reply names the steps that ran.
- trigger regex widened (can't confirm/would need live/won't guess...).
- buildSearchLadder no longer references the out-of-scope connOf helper.
## 0.9.161
### Changed
- Web lookups are now a fully internal, free action: the model can ask for
  {"kind":"web_lookup"} in up to 3 rounds per run (each round's results get
  re-asked into the conversation and the model decides the next move -
  more lookups, ops, or a reply). The lookup never appears in the approval
  gate and never needs approval; the refusal-trigger lookup is unchanged.
- System prompt states the free + repeatable nature explicitly.
## 0.9.160
### Fixed
- Search-profile editor flow: steps attach to a profile before its first
  save (the draft carries over when "Save search profile" runs; the new
  profile gets selected automatically), the online step's connection
  defaults are written back at render time (no more empty-connection
  steps), an online step's connection change rebuilds the row (the model
  list refills), and the profile ladder no longer inherits the legacy
  single-engine setting (the Search tab owns the sequence; the old
  engine/key settings keep working as its migration source).
## 0.9.159
### Added
- Settings -> "Search" tab: search PROFILES. A profile = a name + an
  ordered sequence of lookup steps ("server engine" - DuckDuckGo/Brave/
  Tavily/Firecrawl, or "online lookup" bound to one of your connections -
  OpenRouter attachments carry :online, OpenAI-like need a search-preview
  model, Ollama/llamacpp steps skip themselves). Steps fall through on no
  results or failure; exhausted = the lookup gives up. "Use this profile"
  selects it; no profile = the free keyless DuckDuckGo engine (the
  original default). CRUD via /api/search-profiles (+ /delete); keys per
  engine (searchKeys) migrate from the old single webSearchKey. The
  per-profile "Web search model" slot (0.9.157) stays as the first
  model-side attempt when a profile's online step has no model pinned.
## 0.9.158
### Added
- Firecrawl as a search engine option (Settings -> General -> Web search,
  takes an API key, POST /v1/search with the query). The free DuckDuckGo
  channel stays the keyless default; page reading still runs through the
  server's own fetcher.
## 0.9.157
### Added
- Profiles gained a 4th slot: "Web search model (optional)". Empty (the
  default) = the automatic lookup ladder; when set, that connection+model
  runs FIRST for lookups (on OpenRouter the :online plugin attaches to the
  chosen model; on other providers the model must search natively, e.g.
  OpenAI's gpt-4o-mini-search-preview).
- The web slot's model picker mirrors the image slot: "web-capable only"
  <-> "show all models" toggle (server capability=web filter). Detection:
  OpenRouter = every model qualifies (:online is attachable); Ollama = an
  empty shortlist (no native web search there) so the full list shows; the
  rest = name hints (search-preview, sonar/perplexity, web, online...).
  The Profiles table has a Web model column; renames carry the slot.
## 0.9.156
### Added
- Settings -> General -> "Web search (all add-ins)": pick the search
  engine - Automatic (provider lookup first, then the free engine) /
  DuckDuckGo (free, keyless) / Brave Search / Tavily (the last two take an
  API key, stored only in the local config, never echoed to the panes or
  shown again). The chosen engine runs first in the lookup ladder;
  provider-native channels (OpenRouter :online, OpenAI search-preview
  models, the Anthropic web_search server tool) = the documented natives.
  Free DuckDuckGo = always available, no key.
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
