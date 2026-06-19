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
- **Fetch models** — load model list from `/models` when your gateway supports it
- **Error actions** — Retry failed messages or copy error details
- **Production package** — `manifest.prod.xml` + `npm run package` for org catalog deployment
- **Slash command hints** — type `/` in the message box for autocomplete (`/fix`, `/table`, …)

## System requirements

| Platform | Minimum | Recommended |
|----------|---------|-------------|
| **Word on Windows** | **Office 2016+** on `ie11-rewrite` branch; **Office 2019+** / **M365** on `main` | **Microsoft 365** Word |
| **Word on Mac** | Word **15.18**+ (2016-era) | **Microsoft 365** Word for Mac |
| **Word on the web** | Modern browser (Edge, Chrome, Firefox, Safari) | Microsoft 365 account |

**Branches:** `main` targets **Office 2019+ / M365** (React 19 + Vite). **`ie11-rewrite`** targets **Word 2016 Windows** (IE11 task pane, Webpack + React 16 + ES5). Office 2013 and earlier are not supported.

**Development:** Node.js 20+

## Prerequisites

- **Node.js** 20+
- **Word 2016+** on Windows (`ie11-rewrite`), or **Office 2019+** / **M365** (`main`), or **Word on the web**
- An OpenAI- or Anthropic-compatible API endpoint and API key

## Quick start

```bash
# Install dependencies
npm install

# Windows: trust localhost dev cert (once, elevated PowerShell)
npm run certs

# Start the HTTPS dev server (required by Office add-ins)
npm run dev
```

The dev server runs at **https://localhost:3000**.

### Word 2016 (IE11) — `ie11-rewrite` branch

Office 2016 on Windows uses **IE11** for task panes. Check out **`ie11-rewrite`**:

```bash
git checkout ie11-rewrite
npm install
npm run certs    # once, elevated — trusted localhost HTTPS
npm run dev      # Webpack dev server → https://localhost:3000
npm start        # sideload in Word (close Word first)
```

- **Webpack** + **React 16** + **Fluent UI v8** + **ES5** bundle (no Vite ESM)
- Direct **Settings** and **Chat** tabs — no onboarding wizard, no telemetry
- Chat streaming (XHR SSE), agent mode, edit preview, slash commands, per-document persistence

Legacy UI lives under `src/taskpane/components.legacy/`; `src/taskpane/components/` is the modern reference on `main`.

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

Open **Settings** from the task-pane header (the add-in opens Settings automatically until configured).

1. Choose **OpenAI-compatible** or **Anthropic-compatible**
2. Set **Base URL** — e.g. `https://api.openai.com/v1` or your gateway URL
3. Enter **API key** and **model** name (or click **Fetch models** if `/models` is available)
4. Click **Save settings** and **Test connection**
5. Switch to **Chat**, pick **Chat** or **Agent** mode, and send a message

### Example endpoints

| Provider type | Example base URL |
|---------------|------------------|
| OpenAI-compatible | `https://api.openai.com/v1` |
| OpenAI-compatible (gateway) | `https://your-gateway.example/v1` |
| Anthropic-compatible | `https://api.anthropic.com/v1` |

## Scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Start dev server on port 3000 (HTTPS) — Webpack on `ie11-rewrite`, Vite on `main` |
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
| `npm run certs` | Install trusted localhost dev certs (Windows — run terminal as Administrator) |
| `npm run certs:verify` | Verify dev certificates are installed |

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
- **Build:** Vite 7 with HTTPS on port 3000 (`office-addin-dev-certs` on Windows/Mac)
- **LLM:** Thin `fetch` + SSE adapters; no SDK dependency

## Security notes

- API keys are stored in `localStorage` on the local machine only
- Document context and tool reads are sent to your configured LLM endpoint
- Edits are staged for approval by default; only **Apply** writes to the document
- Context is truncated at ~12,000 characters to limit prompt size
- Edits use internal bookmarks (`msword_aichat_*`) for stable apply; orphan bookmarks may remain in the document
- **Undo** for `insert_table` removes the last table in the document (best-effort)
- Custom endpoints must be reachable from the add-in runtime (watch CORS if not using a same-origin proxy)
- Use HTTPS endpoints in production; local dev uses trusted localhost certs after `npm run certs`

## Distribution

Build and create a deployable folder:

```bash
npm run package -- https://addins.yourcompany.com/msword-aichat
```

Upload the contents of `package/` to your HTTPS origin. Use `manifest.xml` from that folder for sideloading or Microsoft 365 admin center deployment.

For local development, keep using root `manifest.xml` (localhost URLs).

### CORS and local routers

The task pane runs at `https://localhost:3000`. Calls to a **different origin** (your local AI router, a remote gateway, etc.) trigger a browser **CORS preflight** (`OPTIONS` before `GET`/`POST`). If the router returns **405** on `OPTIONS`, you will see **Failed to fetch** in Settings or chat.

**Option A — fix your router (recommended for local dev)**

Respond to `OPTIONS` on `/v1/*` with **204** and headers such as:

```
Access-Control-Allow-Origin: *
Access-Control-Allow-Methods: GET, POST, OPTIONS
Access-Control-Allow-Headers: Authorization, Content-Type, x-api-key, anthropic-version, Accept
```

**Option B — dev CORS proxy**

```bash
# PowerShell — point at your router's base (no /v1 suffix in TARGET_URL host path)
$env:TARGET_URL="http://localhost:8080"
npm run proxy
```

Then set **Base URL** to `http://localhost:8787/v1` in Settings. The proxy answers `OPTIONS` and forwards traffic. Development only — production should allow the add-in origin or use an org-managed proxy.

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

**Blank white task pane (no UI)**

- **Word 2016 Windows:** use branch **`ie11-rewrite`** (Webpack ES5 bundle). `main` uses Vite + React 19, which IE11 cannot run.
- **Office 2019+ / M365:** confirm [WebView2](https://developer.microsoft.com/microsoft-edge/webview2/) is installed.
- Restart `npm run dev` after `npm run certs`; close Word before `npm start`.

**Red “Uncaught runtime errors” overlay (dev only)**

Webpack dev-server HMR can trigger this on IE11. `ie11-rewrite` disables the overlay — restart `npm run dev` if you still see it.

**Fluent dropdown shows white screen (IE11)**

Use native `<select>` controls in the legacy UI (`IeSelect`). Avoid Fluent `Dropdown` / `CommandBar` overflow menus in the task pane on Word 2016.

**Certificate error / "content is blocked" (Word task pane)**

Word uses its own embedded browser and does **not** trust Vite's default self-signed certificate. Install Microsoft's dev certs once:

1. Open **PowerShell or Command Prompt as Administrator**
2. In the project folder: `npm install` then `npm run certs`
3. Stop and restart `npm run dev`
4. Close Word completely, then run `npm start` again
5. Open `https://localhost:3000/taskpane.html` in **Edge** (or IE on older setups) and confirm it loads without a cert warning

`npm run certs:verify` should report the certificate is valid. The dev server prefers `office-addin-dev-certs` over the basic-ssl fallback.

**Add-in does not load**

- Confirm `npm run dev` is running and reachable at `https://localhost:3000`
- Complete the certificate steps above on Windows before sideloading
- Re-run `npm start` to re-register the manifest
- Requires **Word 2016+** or **Microsoft 365 Word**

**Connection test fails / “Failed to fetch”**

- Verify base URL, API key, and model name
- Check router logs for `OPTIONS` returning **405** — see [CORS and local routers](#cors-and-local-routers) above
- Try the same request with `curl` against your endpoint

**Context shows "No text selected"**

- Highlight text in the document, then click the refresh button in the context bar

**Manifest validation**

```bash
npm run validate
```

## License

MIT — see [LICENSE](./LICENSE).