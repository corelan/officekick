# OfficeKick settings reference

What every setting means, where it lives and how to configure it. The three
add-ins share ONE config:

- All values live in `~/.officekick/config.json`
  (`%USERPROFILE%\.officekick\config.json` on Windows; the file is chmod 600).
- The pane re-reads everything each time it opens. Settings survive
  close/reopen, Office restarts and updates.
- Only the local server stores them. The pane never receives API keys back.

## Contents

- [Settings > General](#settings--general)
- [Settings > Approval policy](#settings--approval-policy)
- [Settings > Connections](#settings--connections)
- [Settings > Profiles](#settings--profiles)
- [Settings > Search](#settings--search)
- [Inspector](#inspector-tab)
- [The run activity log](#the-run-activity-log)
- [Where things live on disk](#where-things-live-on-disk)

## Settings > General

### Custom system prompt

- An extra instruction block that rides along with every request.
- The pane shows the EFFECTIVE prompt: your text when set, otherwise the
  built-in default. The built-in default itself is never editable here -
  clear the field and Save to return to it.
- The default has three blocks (`### Content (accuracy)`, `###
  Writing style`, `### Formatting`): stick to facts actually present,
  plain active voice, and a monospace rule for code/command lines.
- Saving: type your text, press **Save custom prompt**. A green
  "✓ saved" flash confirms it. Maximum 4000 characters.
- Keeping the `###` block headers is recommended: the models follow them.

### Enable debug logging

- Checkbox (default off). Saves instantly when toggled.
- Writes one log file per day to `~/.officekick/logs/debug-YYYY-MM-DD.log`.
- The log contains request/response payloads, provider calls, page fetches
  and the text of your document/deck/email. Keep it off for sensitive work.
- The daily access log next to it only lists paths the server served.

### Sessions

- **Auto-clean up history after N days** (default: on, 120 days): when any
  add-in opens, sessions whose newest message is older than N days are
  deleted from `~/.officekick/sessions/`. The sweep runs at most once per
  hour.
- Disable it or change the days, then use **Save session cleanup**. The
  change applies at the next add-in open.

### Display

- **Pane text size**: 75 / 85 / 90 / 100 / 110 / 120 / 130 percent. Scales
  the whole pane. Applies immediately, saves automatically, and is shared
  by all three add-ins.

### About / updates (same tab)

- **Check for updates automatically at startup** (default on): once per
  boot, the pane reads the released VERSION - from your local OfficeKick
  clone when one was recorded at install (it runs `git fetch`), otherwise
  from GitHub. Nothing is downloaded or installed automatically. The probe
  result is cached for two minutes.
- **Check now**: the same check on demand.
- **Update the server automatically before it starts** (default on): when
  the server boots and the released code is newer, it updates itself first
  (git clone: `git pull --ff-only`; otherwise it downloads the GitHub repo
  zip) and restarts through the OS-native start mechanism, so the running
  server IS the new release. Server code + pane files only - run
  `install.py --update` when manifests changed. Failures never block the
  boot; the next start retries (checked at most once per 12 h).
- **Update available**: shows the pull/installer steps:
  1. Reload the pane first (About > **Reload add-in**, or close/reopen the
     add-in) - the new UI is live immediately.
  2. Then run the installer (`python3 install.py --update`, Windows
     `py -3 install.py --update`) to deploy manifests and server code.
- **Reload add-in**: hard-reloads the pane and picks up new build files
  without reopening Office. Your active session auto-resumes.
- The result note under **Check now** clears itself the next time you open
  the General tab, so old "up to date" messages never linger.

## Settings > Approval policy

Approval policy decides, per ACTION type, what happens when a run proposes
that action:

- `auto` - applied immediately (still reported in the results summary).
- `confirm` - the gate waits: a checkbox per action (unchecked default),
  **Select all**, and **Apply**. Only ticked actions run.
- `deny` - the action is blocked and shown with a notice.

Configure:

1. Each action lists its name + what it does on top, then the three modes.
2. Pick a mode per action and press **Save policy**.
3. Defaults (first boot): create/edit slides, replace text, notes, tables,
   move/resize shapes, create shapes, duplicate slide, insert paragraphs,
   edit paragraph, insert comment, add hyperlink, set alt-text, create
   table of contents = `auto`; layout switching, deletions, move slide,
   copy-slide-over, clear slide, background/image/picture-fill/style
   actions, format text/selection, header/footer + tracked-changes
   decisions, rewrite document = `confirm`. MailKick: `draft reply` /
   `forward` / `new meeting` / `set categories` / `save draft` `auto`,
   `compose body` `confirm` - MailKick NEVER SENDS EMAIL, it only opens
   windows you review and send.

Details:

- A per-action **"don't ask again for 'X' this session"** option sits behind
  a divider below the list. It auto-approves THAT action type for the rest
  of the pane session only. It resets when the pane reopens and never
  changes the stored policy.
- Every proposed action row has a red **×** to remove it from the batch
  (two clicks: "sure?" confirm, auto-revert after ~3.5 s).
- Failed actions stay visible in the results and are retried ONCE after a
  short pause, automatically.
- **Auto apply approved actions** (global) plus per-app switches
  (**SlideKick / DocKick / MailKick**): when ON for an app, that app's
  per-run "auto-apply" checkbox pre-ticks and every run whose actions are
  all `auto` skips the gate entirely. When any action needs approval (or is
  invalid), the gate appears as usual.
- The per-run **auto-apply** checkbox near the prompt does the same for a
  single run.

## Settings > Connections

A connection = one AI provider endpoint.

- **Fields**: Name, Type, Base URL, API key (optional for local servers).
- **Types**: openai, generic-openai (OpenAI-compatible endpoints of any
  kind), anthropic, openrouter, ollama, llamacpp.
- **The key is write-only**: type to replace, leave blank to KEEP the
  stored key. After saving, the field clears and the value is never shown
  again - the pane only receives a boolean ("key set").
- **Test connection** fetches the model list through the server and shows
  it as the success message.
- After a save the form clears to "New connection" defaults, ready for
  the next entry.
- All connections are shared by the three add-ins (configure once).

## Settings > Profiles

A profile = which model runs for what. Routes:

- **content** - slide/document/email editing and generation.
- **other** - advice, explanations, brainstorming.
- **Image model (optional)** - generates pictures; empty = no image
  generation. The model list shows image-capable models only; tick the
  **show all models** checkbox to load the full list (can be long, >420
  on OpenRouter). A custom model id field is available.
- **Web search model (optional)** - a model used when a lookup needs one
  (see [Search](#settings--search)); empty = the automatic channel.

The first profile in the list is the default; the pane remembers the
profile you used last. Profiles show per-slot columns; **Edit** re-fills
the form, changes save back into the profile.

## Settings > Search

Web lookups are INTERNAL actions: the model asks for them freely, they
never need approval, and the results feed the next step of the run.
CONTENT-CREATION requests (write an essay/article/report in DocKick,
build slides about a topic in SlideKick) ALWAYS trigger a lookup before
the model is asked: the pane fetches the results itself and hands them to
the model with the run, so facts are verified and cited. Pure edit /
reword / restyle requests do not look anything up.

### The steps sequence

- One ordered list shared by all three add-ins.
- The top step runs first; a step that finds nothing (or fails) falls to
  the next; when all steps fail the lookup gives up with a reply that
  names the steps that ran (and those skipped as unconfigured).
- **No steps at all = the free keyless DuckDuckGo engine.**
- Steps cannot be empty when saving: at least one step, every step
  configured.
- Reorder with the **grip handle** (the dot column before each row): drag
  it; the blue line shows where the step will land. Press **Save search
  steps** to persist.

### Step kinds

1. **Search engine** - the SERVER searches (no model involved):
   `duckduckgo` (free, no key), `brave`, `tavily`, `firecrawl` (these
   three need their own API key under "Engine API keys" below). A keyed
   engine without its key skips itself with a stated reason.
2. **Model lookup** - a CONNECTION's model searches for you:
   - **OpenRouter**: pin a model - `:online` attaches to it automatically.
     No model pinned = the step skips itself (`:online` alone is not a
     valid model id).
   - **OpenAI-style**: pin a `search-preview` model; without one the step
     skips itself.
   - **Ollama / llama.cpp**: no native search channel - the step skips
     itself and says so.
   - The optional model dropdown lists web-capable models - "show all
     models" bypasses it.

The status/hints next to the dropdowns state each rule; the same rules
drive the run.

### Engine API keys

- One password field each for Brave, Tavily and Firecrawl
  ("type to replace, blank to clear").
- The state label next to each field: **"key set" in green**, **"no key"
  in red**; while you type it shows "replacing key" (uncolored).
- After **Save API keys** the fields clear; labels refresh from the
  config. Keys are stored server-side only and never round-trip to the
  pane.
- Legacy setting: the old single "web search key" migrates into this
  store automatically.

### The per-profile web model

When a profile's **Web search model** slot is set, it runs FIRST,
independent of the steps list - then the steps follow.

### Budgets (per run)

- Up to 3 web lookup rounds.
- Up to 4 page reads (distinct pages the model explicitly asks for).
- Up to 6 page-text continuations (reading further parts of a long page;
  these do not consume the page-read budget).

## Inspector tab

The request/response log for troubleshooting: every call the pane made,
with expandable payloads (what was sent, what came back). Enable debug
logging for the server-side view too (`~/.officekick/logs/`).

## The run activity log

While a run works, a box under the status line lists what it is doing and
has done - most recent on top, one line each (e.g. "Sending request to
<connection> (model)", "Web search via <engine>", "Model lookup via
<connection> (model …:online)", page reads, repairs, applies). The box
shows an 8-row window; older entries scroll below. It resets on every
**Run** press and closes nothing when the run ends (you can still read
what happened).

## Where things live on disk

| What | Where |
|---|---|
| Config (+ all API keys, chmod 600) | `~/.officekick/config.json` |
| Sessions | `~/.officekick/sessions/` |
| Debug + access logs (when enabled) | `~/.officekick/logs/` |
| Pane error log (always on) | `~/.officekick/logs/pane-errors-YYYY-MM-DD.log` |
| Installed server | `~/Library/Application Support/OfficeKick/` (macOS) / `%LOCALAPPDATA%\OfficeKick` (Windows) |
| Add-in manifests per app | each Office container's `Data/Documents/wef/` (macOS) |

Related endpoints (same machine only): `http://127.0.0.1:8765/api/health`
(server up + version), `POST /api/server-restart`, `POST /api/repo-update`
(the pane's pull button).
