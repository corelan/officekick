# Changelog
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
