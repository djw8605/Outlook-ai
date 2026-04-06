# Outlook AI Assistant

An AI-powered Outlook add-in that brings the power of OpenAI (or any OpenAI-compatible
endpoint such as [Ollama](https://ollama.com)) directly into your email workflow.

## Features

| Action | Description |
|--------|-------------|
| **Summarize** | Condenses the current email thread into bullet points and action items |
| **Reply** | Drafts a professional reply to the selected email |
| **Improve** | Rewrites a compose draft to be clearer and more professional |

All text is processed through an OpenAI-compatible `/v1/chat/completions` endpoint.
Settings (API key, base URL, model) are stored in Office roaming settings so they
persist across machines.

---

## Prerequisites

- **Node.js 18+** — <https://nodejs.org>
- **Microsoft 365** account with Outlook (desktop or web)
- An **OpenAI API key** — <https://platform.openai.com/api-keys>  
  _or_ a local [Ollama](https://ollama.com) installation (no API key needed)

---

## Installation

```bash
cd outlook-ai-assistant
npm install
```

---

## Development

### Start the dev server

```bash
npm start
```

This starts a **webpack-dev-server on `https://localhost:3000`** with hot-reload.
Outlook add-ins require HTTPS even locally; webpack is pre-configured for this.

> **First run only:** install the development certificates so your browser trusts
> the self-signed cert:
>
> ```bash
> npx office-addin-dev-certs install
> ```

### Production build

```bash
npm run build
```

Output goes to `dist/`.

---

## Sideloading the manifest in Outlook for Mac

1. Open **Outlook for Mac**.
2. Click **Get Add-ins** (or **Store**) in the Home ribbon.
3. Choose **My Add-ins → + Add a custom add-in → Add from file…**
4. Select `outlook-ai-assistant/manifest.xml`.
5. The **AI Email Assistant** button will appear in the ribbon when reading or
   composing an email.

> Alternatively, sideload via the Microsoft 365 admin center for org-wide deployment.

---

## Switching from OpenAI to Ollama

1. [Install Ollama](https://ollama.com/download) and pull a model, e.g.  
   `ollama pull llama3`
2. Start the Ollama server (runs on `http://localhost:11434` by default).
3. In the task pane, open **⚙ Settings** and update:
   | Field | Value |
   |-------|-------|
   | API Key | any non-empty string (e.g. `ollama`) |
   | Base URL | `http://localhost:11434/v1` |
   | Model | `llama3` (or whichever model you pulled) |
4. Click outside the inputs to save — the settings persist automatically.

---

## Project Structure

```
outlook-ai-assistant/
├── manifest.xml             # Office add-in manifest
├── package.json
├── tsconfig.json
├── webpack.config.js
├── .env.example             # Environment variable reference
├── assets/
│   └── icon-32.png          # Ribbon icon
└── src/
    ├── taskpane/
    │   ├── taskpane.html    # Task pane UI
    │   ├── taskpane.ts      # Task pane logic
    │   └── taskpane.css     # Styles
    ├── commands/
    │   └── commands.ts      # Ribbon command handlers
    └── lib/
        ├── ai-client.ts     # OpenAI-compatible API client
        └── outlook-helper.ts # Office.js email helpers
```

---

## Configuration Reference

| Setting | Default | Description |
|---------|---------|-------------|
| API Key | _(none)_ | Your OpenAI secret key or any string for Ollama |
| Base URL | `https://api.openai.com/v1` | Any OpenAI-compatible `/v1` endpoint |
| Model | `gpt-4o` | Model identifier (must be available on your endpoint) |

Settings are saved to **Office roaming settings** — they follow you across devices
in the same Microsoft 365 tenant.
