# Changelog
## [0.4.3] - 2026-10-30
### Fixed
- move_slide shipped without its OPS_META whitelist entry ("unknown op type
  'move_slide'" in the gate); the entry is in and the op is E2E-verified in
  the simulator (move last slide -> position 2 reorders correctly).

## [0.5.3] - 2026-10-30
### Fixed
- Settings sub-tab headers: readable in Office dark mode (lime) and a darker green in light mode.

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
- `tests/` is untracked and gitignored (local-only); docs/outlook.md rewritten
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
- **MailKick** - the Outlook add-in (docs/outlook.md): email context, draft_reply /
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
