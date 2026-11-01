# Corelan SlideKick - the PowerPoint add-in

SlideKick is an AI assistant task pane for **Microsoft PowerPoint** (macOS and
Windows). Type what you want - `reword the current slide`, `make slides about
threat hunting`, `design this deck`, `add a table of the findings` - and
SlideKick reads your deck, routes the request to the model you chose, and
applies the changes **after** the per-action approval policy allows it.

Install + uninstall basics: see the top-level [README](../README.md). This doc
is about how SlideKick works.

## Requirements (PowerPoint specifics)

- Microsoft PowerPoint from Microsoft 365. Slide editing needs PowerPointApi
  ~1.8 (Mac 16.96 / 2504+, Windows 2504+ builds); perpetual/LTSC editions run
  the pane read-only (read + advice work, edits are refused with a clear
  message).
- An AI provider reachable from your machine (cloud API or local
  Ollama/llama-server). Configure it in the pane's Connections tab; profiles +
  Home-lab providers (Ollama/llama.cpp on a private IP) additionally need the
  macOS **Local Network** grant for the server - see the README's "macOS
  permissions" section (Errno 65 / "No route to host" = that grant missing).
  policy are shared with DocKick ([docs/dockick.md](dockick.md)).

## Using it inside PowerPoint

- **macOS**: open a deck, then **Home tab -> Add-ins** (dropdown) ->
  **Developer Add-ins** -> **Corelan SlideKick**. The pane opens on the right;
  the Add-ins entry stays available every session.
- **Ribbon (opt-in)**: install with `python3 install.py --ribbon` to keep the
  manifest's add-in commands: an **AI** group with the Corelan logo
  (16/32/80 px icons) and an "Open SlideKick" button on the Home ribbon.
  - Word keeps the same group for DocKick.
  - Outlook shows the group on the message read/compose tabs.
  - Default is list-mode (no commands).
  - Retested on a clean state (0.9.29): the 16.114.3 Office-Mac build accepts
    the commands manifest, renders nothing and drops the Add-ins entry - a
    build limitation for developer file-sideloads, not a cache artifact.
    Restore with a plain install re-run (it bumps the manifest Version).
  Default = list mode: on the Office-Mac build we tested, a command manifest
  REPLACED the Add-ins-list entry and then rendered neither the group nor a
  list item (and earlier 16.114 builds behaved the same) - the add-in was
  unreachable until the installer was re-run without `--ribbon`.
- **Windows**: quit and reopen PowerPoint, then **Insert -> My Add-ins
  (dropdown) -> SHARED FOLDER -> Corelan SlideKick -> Add** once.

## Typed instructions, routes and approvals

- The **route** decides which model runs: `auto` sends slide-related requests
  (reword / create / make slides / notes / tables...) to the **content model**
  and everything else to the **other-actions model**; force one with the Route
  dropdown.
- With *include deck context* checked, the pane sends a compact outline (slides,
  shape indices with texts **and geometry**, selection) with your request, so
  instructions like "reword the current slide" work without naming slide numbers.
- The model returns a structured action list. SlideKick **validates** every
  action against the outline (indices, shape references) and applies it only
  after the approval policy is satisfied:

  | Action | What it does | Default policy |
  |---|---|---|
  | create slide | append a new slide (title + bullets + notes when enabled) | auto |
  | edit slide | replace title / body / notes of an existing slide | auto |
  | replace text | reword one shape in place (the "reword" path); with a `find` fragment only that fragment is replaced - the rest of the shape's text (and its bullet list) stays untouched. For long lists always use `find` (the item's current words) rather than rewriting the whole shape | auto |
  | set notes | write speaker notes | auto |
  | create table | adds a table (optionally on a new title slide; position/size in points) | auto |
  | move/resize shape | sets left/top/width/height (points) of one shape | auto |
  | apply layout | switch a slide's layout (design change, text kept) | confirm |
  | delete slide | remove a slide | confirm |
  | reply | conversational answer - never touches the deck | n/a |

  Change any of these in **Settings**. `confirm` shows a checkbox per action
  (unchecked by default) with a *don't ask again this session* option; the
  preview also has a **Select all** button to approve everything pending at
  once. `deny` blocks the action with a visible notice. Every proposed action
  row also has a red **X** to remove it from the list (second click confirms). The **auto-apply**
  checkbox skips the review step entirely for runs where every action is `auto`
  by policy - if any action needs approval (or is blocked/invalid), the gated
  preview appears exactly as without it.

- **Conversation memory**: follow-ups include your recent turns - the model
  sees the last 6 exchanges of the session, so "use the 2nd title you
  suggested" just works. **New session** starts fresh.
- **Sessions & the Previous-requests box**: every exchange is auto-saved in a
  local, AI-named session (resumable from the **History** tab). The box below
  the prompt lists every run of the (resumed) session, newest first: the
  timestamp first, the request text below it, and a +/- box to expand the full
  request and the response/summary. Clicking a stored request refills an empty
  input box. While a run works your typed text stays visible (locked input);
  it is cleared when the result prints, with "Request: ..." echoed above it -
  a failed run re-enables the input and keeps your text. Older saves titled
  just "Session" get an AI-generated title on resume (stored in the session
  file).
- **Speaker notes are opt-in (off by default)**: SlideKick only generates or
  writes speaker notes when "Generate speaker notes" is enabled in Settings.
  When off, the model is told not to produce notes and any notes actions it
  returns anyway are skipped with a visible notice.
- Notes that cannot be written automatically on older builds fall back to a
  copy-paste box in the pane.

### Merging another deck

The "Merge deck…" button (next to Refresh deck) appends ALL slides of another
.pptx to the end of the current deck (PowerPointApi 1.2,
`presentation.insertSlidesFromBase64`, source formatting kept). The slides
land after the last slide; the deck then refreshes automatically.

### Styling shapes

`style_shape` gives the model a direct lever over existing shapes (PowerPointApi
1.4): a solid fill color and/or an outline color for ONE shape, e.g.
"make the callout on slide 4 dark blue with an orange border". Works on any
shape in the outline, including textboxes and the shapes created by
`create_shape`.

## Design on demand

Use the **Design this deck** chip (or type e.g. "design this deck",
"rebalance slide 3", "align everything to the left"). SlideKick sends the slide
geometry (shape positions/sizes in points) and the model returns
`move_shape` / `apply_layout` / `create_table` ops through the same approval
gate. Note: Microsoft's built-in *Designer* (Design Ideas) **cannot be invoked
by add-ins** - this is SlideKick's own AI design pass, and its edits stay fully
editable in PowerPoint.

With today's Office.js the add-in can create: slides (with layout), text,
tables (PowerPointApi 1.8 `addTable`) and edit shape geometry. Charts have no
PowerPoint.js creation API and native SmartArt is not reachable from add-ins.

## Reading slide content (read_slide)

In large decks (more than 40 slides) the deck outline the model sees is
compact: slide titles plus a short text preview - with one exception: the
SELECTED slide(s) are shown in full, every shape with its complete text.
When the model needs the exact content of another slide (to answer
"explain slide 70" or to rewrite it), it asks SlideKick itself by replying
`{"kind":"read_slide","slides":[70]}` (1-based outline numbers, up to 12
slides). SlideKick reads those slides from the deck - in real PowerPoint
and in the simulator, plain text boxes included; speaker notes are not part
of read_slide - and re-asks the model with the complete text, which it then
uses to answer. Up to 2 read rounds per run keep a model from looping.
The read itself needs no approval: it only reads, it never writes.

## Supported requests in real life

- `reword the current slide` - replaces wording, keeps meaning and slide count
- `make 4 slides about threat hunting` - appends 4 new slides (Office.js can
  only append; drag them into place if needed)
- `tighten the bullets on this slide`, `turn this text into a table` (reply or
  a real table)
- `add speaker notes to slides 2-5` (with notes enabled)
- `suggest 3 design improvements for the current slide` - via the other model
- `set the layout of slide 3 to title only`
- `explain slide 70` - in a large deck the model asks read_slide for that
  slide and answers from its exact text instead of guessing from the title

## Pictures (AI image generation)

Type e.g. `add a picture of a clean data center rack` or `insert a picture of ...`.
SlideKick/DocKick asks the model for a **create_image/insert_image** op, fetches the
actual pixels from a provider with an OpenAI-style images API, and inserts them
(op default = confirm, change per-op in Settings).

- Image models live in the profile: the optional third slot **Image model**
  (connection + model - e.g. OpenAI with `gpt-image-1`).
- The Image model list shows **image-capable only** (via the provider's own
  metadata on OpenRouter; name heuristics elsewhere). Toggle "show all models"
  if your image model isn't detected; a custom model id is always available.
- Generation support: openai and generic-openai servers that implement the
  images route; openrouter works only if the endpoint implements it;
  anthropic / ollama / llama.cpp cannot generate images.
- Insertion (PowerPoint): lands on the slide you currently have open (Office.js
  limitation); Word: at the end of the document.
- If the profile has no image model, the run says so instead of failing dark.

## Settings reference

Open the gear tab: five sub-tabs - **General**, **Approval policy**,
**Connections**, **Profiles**, **Inspector**. Every value lives in
`~/.officekick/config.json`, is re-read on each pane open (settings survive
close/reopen and app restarts) and only ever stored locally.

### General

- **Check for updates automatically at startup** (default on): once per boot
  the pane reads the released VERSION from GitHub and shows "Update available"
  with the update steps when a newer release exists (see the README section
  "How to update"). Token-free (plain HTTPS against the public repo); nothing
  is ever downloaded or installed automatically. "Check now" runs the same
  check on demand.
- **Enable debug logging** (default off): appends AI interactions and
  document/deck/mail traces to `~/.officekick/logs/debug-YYYY-MM-DD.log` (one
  file per day). The log contains text content and full provider payloads -
  keep it off for sensitive work.


- **Auto-clean up history after X days** (Sessions section, enabled by
  default, 120 days): when opening any of the add-ins, sessions whose
  NEWEST message is older than the number of days are deleted from
  `~/.officekick/sessions/` (the sweep runs at most once per hour). Disable it
  or change the days with **Save session cleanup**; the change applies at the
  next add-in open.

### Approval policy

- **Behaviour - "Auto apply approved actions (SlideKick)"**: per-add-in since
  0.9.0 (one value for SlideKick, one for DocKick, one for MailKick; the old
  global value keeps working as the default until an add-in gets its own).
  When ON, every run ticks the run's "auto-apply" box; when OFF it stays
  manual.
- **Approval policy per action**: one block per action - name + "what it
  does" on top, the three modes below:
  `auto` = apply immediately (still shown in the summary with results),
  `confirm` = waits for the checkbox + Apply in the gate,
  `deny` = blocked, shown with a notice.
  After the list (behind a divider) a per-action "don't ask again for 'X'
  this session" option may appear - ticking it auto-approves THAT action type
  for the REST OF THE PANE SESSION only (it resets when the pane reopens; it
  does not change the stored policy). Actions the policy denies or rejects
  are listed but marked, and never applied. **Save policy** persists the
  whole table.
- The gate itself: proposed actions appear with the description below the
  action tag, a per-action checkbox (right) and a red x to remove an action
  from the batch (two clicks: "sure?" confirm, auto-revert after ~3.5 s). A
  divider separates the list from the "don't ask again" checkboxes.

- **Generate speaker notes when creating or editing slides** (PowerPoint
  section, default off): when off, notes are never generated or written on
  slides - notes actions from the model are skipped with a visible notice and
  the model is told not to produce them.

## Simulator & development (PowerPoint)

```bash
# serve straight from a repo checkout (dev mode)
python3 core/server.py --config ~/.officekick --static-dir "$(pwd)" --port 8765
open http://127.0.0.1:8765/slidekick/static/simulator.html   # the exact task pane + in-memory deck
python3 tools/mock_provider.py --flavor openai --port 8901    # mock provider for the UI
npx -y office-addin-manifest validate "$HOME/Library/Application Support/OfficeKick/manifest/slidekick.xml"
python3 tools/make_icons.py                         # rebuild icons from logo_corelan.png
```

The task pane runs unmodified in the simulator via `sim.js` (the same Deck
interface the Office.js adapter implements), so prompt -> JSON ops ->
validation -> approvals -> apply is fully testable in a browser.

Validator note: the official validator passes fully in `--https` mode. In the
default plain-http mode it reports "insecure URL" errors for the
`http://localhost` pane URLs - expected: the validator has no localhost
exception, while Office itself accepts http://localhost (the default mode's
design).

## Troubleshooting (PowerPoint)

- **Add-in doesn't show**:
  1. Quit all Office apps and reopen PowerPoint.
  2. Check the health endpoint: http://127.0.0.1:8765/api/health.
  3. Check `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/slidekick.xml` exists.
  4. Re-run the installer (it clears Office's add-in caches and the WEF
     registration store - only while PowerPoint is closed).
- **"Cannot reach the server" in the pane**:
  1. Open the About tab: "Pane build" shows the running build.
  2. Check http://127.0.0.1:8765/api/health.
  3. Check the log: tail -50 ~/Library/Logs/OfficeKick/server.log
  4. Re-run the installer (it restarts the server).
- **"lacks the text-editing APIs" / read-only pane** - update Microsoft 365
  (PowerPointApi 1.8: Mac 16.96+, Windows 2504+).
- **401/403 from a provider** - fix the key on the Connections tab (a blank key
  field keeps the stored key).
- **Model answers arrive as plain text** - the model ignored the JSON contract;
  with debug logging on, the daily log (`~/.officekick/logs/debug-YYYY-MM-DD.log`)
  shows exactly what it sent.
- **Certificates** - re-run the installer with `--https`; to reset TLS delete
  `~/.officekick/certs/` and re-run.
- **Wrong port** - re-run the installer with `--port`/`-Port`; it re-renders
  manifests and restarts the server.

## License

Licensed by **Corelan Consulting bv** - see [LICENSE](../LICENSE) (free to use
with attribution, use at your own risk, no sales/commercialization,
modifications must reference the original author).