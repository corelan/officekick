# Changelog

## 0.1.0 - initial release

- Office task-pane add-in for PowerPoint on macOS (Office.js), sideloaded via the
  classic XML manifest.
- Local companion server (Python 3 stdlib, 127.0.0.1 only): serves the UI, stores
  connections/profiles/approval policy, proxies AI calls.
- Provider types: openai, anthropic, openrouter, generic-openai (OpenAI-compatible),
  ollama, llama.cpp (llama-server).
- Connections (Name / Base URL / optional API Key) and profiles (content model +
  other-actions model, each picked as connection -> model list).
- Free-text instruction box (`reword the current slide`, `make 4 slides about X`, ...)
  with a structured slide-ops engine (create/edit slide, replace text, notes,
  apply layout, delete slide).
- Per-action approval policy: auto / always-confirm / deny, editable in Settings;
  delete-slide and apply-layout default to always-confirm.
- Browser simulator (`simulator.html`) for end-to-end testing without PowerPoint.
- Installer (`install.sh`): detects PowerPoint, deploys files + manifest
  (~ update/re-run supported), installs a LaunchAgent for the local server.