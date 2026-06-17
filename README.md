# Word AI Chat

A Microsoft Word task-pane add-in that brings AI chat and agentic document editing to Word. Connect any **OpenAI-compatible** or **Anthropic-compatible** endpoint (OpenAI, Anthropic, LiteLLM, vLLM, Ollama gateways, private proxies, etc.).

## Status

**Phase 2 complete.** Agent mode runs a tool loop against the document with edit preview and approval.

| Phase | Scope | Status |
|-------|-------|--------|
| 0 | Scaffold, settings, streaming chat | Done |
| 1 | Document context (selection, outline), quick actions | Done |
| 2 | Agent loop + document edit tools | Done |
| 3 | Advanced editing (styles, tables, undo) | Planned |
| 4 | Polish, Word on the web, distribution | Planned |

## Features (today)

- Task-pane UI on the Word **Home** tab (**AI Chat** button)
- Provider settings: type, base URL, API key, model, temperature, max tokens
- Connection test against your configured endpoint
- **Chat mode** — streaming responses with optional document context
- **Agent mode** — multi-step tool loop with document read/write tools
- **Document tools:** `get_selection`, `get_document_text`, `insert_text`, `replace_text`
- **Edit preview** — before/after diff with **Apply** / **Reject** (default)
- Optional **auto-apply edits** in Settings
- Agent step trace (collapsible) per assistant message
- Document context modes: Selection, Outline, or None
- Quick actions: **Summarize**, **Improve**, **Explain**
- API keys stored locally on the device (separate from other settings)

## System requirements

| Platform | Minimum | Recommended |
|----------|---------|-------------|
| **Word on Windows** | Word **2016** + Microsoft Edge installed | **Microsoft 365** Word |
| **Word on Mac** | Word **15.18**+ (2016-era) | **Microsoft 365** Word for Mac |
| **Word on the web** | Modern browser (Edge, Chrome, Firefox, Safari) | Microsoft 365 account |

**Not supported:** Office 2013 and earlier. The add-in uses ribbon commands, a modern web stack (React 19), and the Word JavaScript API — all of which require Word 2016 or newer.

**Development:** Node.js 20+

## Prerequisites

- **Node.js** 20+
- **Microsoft Word 2016+** or **Microsoft 365 Word** (Desktop on Windows or Mac) for full add-in testing
- An OpenAI- or Anthropic-compatible API endpoint and API key

## Quick start

```bash
# Install dependencies
npm install

# Start the HTTPS dev server (required by Office add-ins)
npm run dev
```

The dev server runs at **https://localhost:3000**.

### Sideload in Word (Desktop)

```bash
# Starts dev server and registers the manifest with Word
npm start
```

Then in Word: **Home → AI Chat** to open the task pane.

To unregister:

```bash
npm stop
```

### Browser-only UI test

Open [https://localhost:3000/taskpane.html](https://localhost:3000/taskpane.html) in a browser. Chat and settings work, but Word document context APIs are unavailable outside Word.

## Using chat and agent modes

### Chat mode

Streaming Q&A with optional document context attached to the system prompt. No document tools are called.

### Agent mode (default)

The agent can call tools to read and edit the document:

| Tool | Description |
|------|-------------|
| `get_selection` | Read the current selection |
| `get_document_text` | Read a chunk of body text (`start`, `max_chars`) |
| `insert_text` | Insert text at selection or end of document |
| `replace_text` | Replace the current selection |

**Edit flow (default):** mutation tools stage a before/after preview. Click **Apply** to write to Word or **Reject** to discard.

Enable **Auto-apply document edits** in Settings to skip the preview.

### Example: rewrite selected text

1. Select a paragraph in Word
2. Open **AI Chat** → mode **Agent**
3. Send: `Rewrite this in a formal tone`
4. Review the agent steps and edit preview
5. Click **Apply**

### Document context bar

1. Choose **Context**: Selection, Outline, or None
2. Click **Refresh** to re-read the document before sending
3. Use **Quick actions** (Summarize / Improve / Explain) with Selection or Outline context

## Configuration

On first launch (or when settings are incomplete), the add-in opens **Settings**.

1. Choose **OpenAI-compatible** or **Anthropic-compatible**
2. Set **Base URL** — e.g. `https://api.openai.com/v1` or your gateway URL
3. Enter **API key** and **model** name
4. Click **Save settings**
5. Click **Test connection** to verify the endpoint responds
6. Switch to **Chat** and send a message

### Example endpoints

| Provider type | Example base URL |
|---------------|------------------|
| OpenAI-compatible | `https://api.openai.com/v1` |
| OpenAI-compatible (gateway) | `https://your-gateway.example/v1` |
| Anthropic-compatible | `https://api.anthropic.com/v1` |

## Scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Start Vite dev server on port 3000 (HTTPS) |
| `npm start` | Sideload add-in in Word + start dev server |
| `npm stop` | Unregister sideloaded add-in |
| `npm run build` | Typecheck and build production bundle to `dist/` |
| `npm run preview` | Preview production build |
| `npm run typecheck` | Run TypeScript without emitting |
| `npm run validate` | Validate `manifest.xml` |

## Project structure

```
msword-aichat/
├── manifest.xml              # Word add-in manifest (sideload / store)
├── taskpane.html             # Task pane entry HTML
├── commands.html             # Ribbon commands stub (required by manifest)
├── public/assets/            # Add-in icons
└── src/
    ├── agent/                # Orchestrator, tools, system prompt
    ├── llm/                  # OpenAI & Anthropic-compatible providers
    ├── word/                 # Document context + Office.js operations
    ├── settings/             # Zustand store + defaults
    ├── hooks/                # useChat, useDocumentContext
    ├── types/                # Shared TypeScript types
    └── taskpane/             # React UI (App, Chat, Settings)
```

## Architecture (high level)

```
Word (Office.js)  ←→  Task Pane UI (React)
         ↓                    ↓
   Document tools        Agent orchestrator
   + context                  ↓
                          LLM Provider  →  Your API endpoint
```

- **UI:** React 19, Fluent UI v9, Zustand for settings
- **Build:** Vite 7 with `@vitejs/plugin-basic-ssl` (HTTPS on port 3000)
- **LLM:** Thin `fetch` + SSE adapters; no SDK dependency

## Security notes

- API keys are stored in `localStorage` on the local machine only
- Document context and tool reads are sent to your configured LLM endpoint
- Edits are staged for approval by default; only **Apply** writes to the document
- Context is truncated at ~12,000 characters to limit prompt size
- Keep selection stable before clicking **Apply** on a `replace_text` edit
- Custom endpoints must be reachable from the add-in runtime (watch CORS if not using a same-origin proxy)
- Use HTTPS endpoints in production; the dev server uses a self-signed certificate

## Roadmap

See [AGENTS.md](./AGENTS.md) for the full phased plan and implementation guide for contributors and coding agents.

**Next up (Phase 3):** styles, formatting, tables, search, and undo snapshots.

## Troubleshooting

**Add-in does not load**

- Confirm `npm run dev` is running and reachable at `https://localhost:3000`
- Accept the self-signed certificate warning in your browser once
- Re-run `npm start` to re-register the manifest
- Requires **Word 2016+** or **Microsoft 365 Word**

**Connection test fails**

- Verify base URL, API key, and model name
- Check gateway CORS headers if calling a remote proxy from the task pane
- Try the same request with `curl` against your endpoint

**Context shows "No text selected"**

- Highlight text in the document, then click the refresh button in the context bar

**Manifest validation**

```bash
npm run validate
```

## License

Private project — license TBD.