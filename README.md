# Word AI Chat

A Microsoft Word task-pane add-in that brings AI chat and agentic document editing to Word. Connect any **OpenAI-compatible** or **Anthropic-compatible** endpoint (OpenAI, Anthropic, LiteLLM, vLLM, Ollama gateways, private proxies, etc.).

## Features

- Task-pane UI on the Word **Home** tab (**AI Chat** button)
- Provider settings: type, base URL, API key, model, temperature, max tokens
- Connection test against your configured endpoint
- **Chat mode** — streaming responses with optional document context
- **Agent mode** — multi-step tool loop with document read/write tools
- **Document tools (10):** `get_selection`, `get_document_text`, `search_document`, `insert_text`, `replace_text`, `delete_range`, `apply_style`, `format_range`, `insert_table`, `insert_comment`
- **Slash commands:** `/fix`, `/table`, `/toc`, `/summarize`, `/formal`, `/comment`
- **Custom instructions** — persistent persona/rules in Settings
- **Per-document conversations** — chat history keyed by document URL (toggle in Settings)
- **Review mode** — agent prefers `insert_comment` over body edits
- **Stable apply** — mutations capture a bookmarked range so Apply works after selection changes
- **Edit preview** — before/after diff with **Apply** / **Reject** (default)
- **Undo** — revert the last applied edit from the preview panel
- Optional **auto-apply edits** in Settings
- Agent step trace (collapsible) per assistant message
- Document context modes: Selection, Outline, or None
- Quick actions: **Summarize**, **Improve**, **Explain**
- API keys stored locally on the device (separate from other settings)
- **First-run onboarding** — provider setup, connection test, mode selection
- **Fetch models** — load model list from `/models` when your gateway supports it
- **Error actions** — Retry failed messages or copy error details
- **Production package** — `manifest.prod.xml` + `npm run package` for org catalog deployment
- Opt-in telemetry hooks (local debug only; no remote collection yet)

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
| `replace_text` | Replace the current selection (bookmark-captured range) |
| `delete_range` | Delete the current selection |
| `search_document` | Find text occurrences with positions |
| `apply_style` | Apply Normal, Heading1–3, Title, or Subtitle |
| `format_range` | Bold, italic, or font size on selection |
| `insert_table` | Insert a table (up to 20×10) after selection |
| `insert_comment` | Add a Word review comment on the selection (immediate) |

**Edit flow (default):** mutation tools stage a before/after preview. Click **Apply** to write to Word, **Reject** to discard, or **Undo** after applying.

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

On first launch, a **5-step onboarding wizard** walks you through setup. You can revisit **Settings** anytime.

1. Choose **OpenAI-compatible** or **Anthropic-compatible**
2. Set **Base URL** — e.g. `https://api.openai.com/v1` or your gateway URL
3. Enter **API key** and **model** name (or click **Fetch models** if `/models` is available)
4. Click **Save settings** and **Test connection**
5. Pick default **Chat** or **Agent** mode, then **Get started**

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
| `npm run validate:prod` | Validate `manifest.prod.xml` template |
| `npm run package` | Build and assemble `package/` for deployment |
| `npm run package -- https://host/path` | Package with production manifest URLs filled in |
| `npm run proxy` | Start local CORS proxy on port 8787 (see below) |
| `npm run smoke` | Run automated smoke test (build, validate, dev server, contracts) |

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
- Edits use internal bookmarks (`msword_aichat_*`) for stable apply; orphan bookmarks may remain in the document
- **Undo** for `insert_table` removes the last table in the document (best-effort)
- Custom endpoints must be reachable from the add-in runtime (watch CORS if not using a same-origin proxy)
- Use HTTPS endpoints in production; the dev server uses a self-signed certificate

## Distribution

Build and create a deployable folder:

```bash
npm run package -- https://addins.yourcompany.com/msword-aichat
```

Upload the contents of `package/` to your HTTPS origin. Use `manifest.xml` from that folder for sideloading or Microsoft 365 admin center deployment.

For local development, keep using root `manifest.xml` (localhost URLs).

### CORS dev proxy (enterprise / local gateways)

If your LLM gateway blocks browser origins from the task pane, run the included proxy:

```bash
TARGET_URL=https://api.openai.com npm run proxy
```

Then set **Base URL** to `http://localhost:8787/v1` in Settings. The proxy adds CORS headers and forwards requests. Use only for development — production should allow the add-in origin or use an org-managed proxy.

## Word on the web QA checklist

Manual smoke test in Word on the web before org-wide rollout:

| Area | Test |
|------|------|
| Load | Task pane opens from Home → AI Chat |
| Settings | Save provider, test connection, fetch models |
| Chat | Stream a response with Selection context |
| Agent | Run a read-only tool (`get_selection`) |
| Edits | `replace_text` preview → Apply → Undo |
| Styles | `apply_style` on a heading |
| Tables | `insert_table` with 2×2 data |
| Search | `search_document` returns matches |
| Errors | Retry and Copy on a failed request |

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

MIT — see [LICENSE](./LICENSE).