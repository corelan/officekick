# SlideKick architecture

```
        Microsoft PowerPoint for Mac (host)
+--------------------------------------------------+
|  SlideKick task pane (WKWebView, Office.js)      |
|    taskpane.html / .css / .js                    |
|    - connections / profiles / settings UI        |
|    - prompt builder + route (content | other)    |
|    - ops validation + approval gate              |
|    - OfficeDeck: PowerPoint.run(...)             |
|         slides.add / getSelectedSlides /         |
|         shapes textRange.text / applyLayout /    |
|         slides.removeAt / notesPage              |
+---------------------|----------------------------+
                      | same-origin http://127.0.0.1:8765
                      v
+--------------------------------------------------+
|  slidekick_server.py (py3 stdlib, threads)       |
|    static files (no-store)                       |
|    /api/config  /api/connections  /api/profiles  |
|    /api/settings (approvals, debug, default)     |
|    /api/models   /api/chat      /api/test-conn   |
|    /api/debug   -> daily append log              |
|    Store: ~/.slidekick/config.json (0600)        |
|    adapters: openai | openrouter | generic|2xURL |
|              anthropic(messages) | ollama(chat)  |
|              llamacpp (v1/chat/completions)      |
+---------------------|----------------------------+
                      | https/http
                      v
              OpenAI / Anthropic / OpenRouter /
              OpenAI-compatible / Ollama / llama-server
```

Key design points:

- **One Deck interface, two implementations.** `applyOps()` (ops engine, validation,
  index shifting for deletes) is shared; `OfficeDeck` talks to Office.js,
  `SimDeck` keeps an in-memory deck and renders it in the browser. This is what makes
  the browser simulator an honest end-to-end test of the pipeline.
- **Append-only slide creation.** Office.js `SlideCollection.add()` always appends;
  mid-deck insert is not supported by the API, so create ops report their requested
  position and the UI notes the append.
- **Approval policy lives server-side in config, is enforced client-side** (the server
  never touches the deck).
- **Reasoning handling**: strip tags server-side (single source of truth) + again
  client-side; structured reasoning fields dropped at the adapter level; a hidden
  repair round re-asks for clean JSON when parsing fails.
- **Keys never reach the pane**: `/api/config` reports `hasKey` only; chat/model calls
  attach the stored key server-side.
- **Sideload path on macOS** (documented by Microsoft): copy the classic XML manifest
  to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`; the add-in
  appears under Home -> Add-ins after restarting PowerPoint. On Windows the same
  manifest is registered as a per-user trusted catalog
  (`HKCU\...\Office\16.0\Wef\TrustedCatalogs`) - see `install.ps1`.