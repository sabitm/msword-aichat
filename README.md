# Word AI Chat

A Microsoft Word task-pane add-in that brings AI chat and agentic document editing to Word. Connect any **OpenAI-compatible** or **Anthropic-compatible** endpoint (OpenAI, Anthropic, LiteLLM, vLLM, Ollama gateways, private proxies, etc.).

Built for **Word 2016 Windows** (IE11 task pane) and also runs on **Office 2019+**, **Microsoft 365**, and **Word on the web**.

## Features

- Task-pane UI on the Word **Home** tab (**AI Chat** button)
- Provider settings: type, base URL, API key, model, temperature, max tokens
- Connection test against your configured endpoint
- **Chat mode** — streaming responses with optional document context
- **Agent mode** — multi-step tool loop with document read/write tools
- **Agent tools (12):** `get_selection`, `get_document_text`, `search_document`, `list_tables`, `insert_text`, `replace_text`, `delete_range`, `apply_style`, `format_range`, `insert_table`, `update_table`, `insert_comment`
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
- **Slash command hints** — type `/` in the message box for autocomplete
- **Production package** — `manifest.prod.xml` + `npm run package` for org catalog deployment

## System requirements

| Platform | Minimum | Notes |
|----------|---------|-------|
| **Word 2016 Windows** | Office 2016 desktop | **Primary target** — IE11 task pane, ES5 Webpack bundle |
| **Word 2019+ / M365 Windows** | WebView2 | Regression-tested; same ES5 bundle runs on modern WebView |
| **Word on Mac** | Word 15.18+ | Best-effort |
| **Word on the web** | Modern browser | Regression-tested |

Office 2013 and earlier are not supported.

**Development:** Node.js 20+

## Quick start

```bash
npm install

# Windows: trust localhost dev cert (once, elevated PowerShell)
npm run certs

# Start the HTTPS dev server (required by Office add-ins)
npm run dev
```

The dev server runs at **https://localhost:3000** (Webpack + React 16 + Fluent UI v8 + ES5).

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
| `get_selection` | Read the current selection (`inTable` hint when cursor is inside a table) |
| `get_document_text` | Read a chunk of body text (`start`, `max_chars`) |
| `search_document` | Find text occurrences with match text and estimated positions |
| `list_tables` | List document tables with index, size, values, and preview |
| `insert_text` | Insert text at selection or end of document |
| `replace_text` | Replace selected **plain text** (not tables — use `update_table`) |
| `delete_range` | Delete the current selection |
| `apply_style` | Apply Normal, Heading1–3, Title, or Subtitle |
| `format_range` | Bold, italic, or font size on selection |
| `insert_table` | Insert a new table (up to 20×10) after selection or document end |
| `update_table` | Replace an existing table **in place** via full `cells` grid (same rows/columns) |
| `insert_comment` | Add a Word review comment on the selection (immediate) |

**Edit flow (default):** mutation tools stage a before/after preview. Click **Apply** to write to Word, **Reject** to discard, or **Undo** after applying.

Enable **Auto-apply document edits** in Settings to skip the preview.

### Example: rewrite selected text

1. Select a paragraph in Word
2. Open **AI Chat** → mode **Agent**
3. Send: `Rewrite this in a formal tone`
4. Review the agent steps and edit preview
5. Click **Apply**

### Example: create or edit a table

**New table** — agent calls `insert_table` once with `rows`, `columns`, and the full `cells` 2D array (including headers). Shortcut: `/table …` in the message box.

**Edit existing table in place** — agent calls `list_tables` (when multiple tables exist), then `update_table` with the same dimensions and a full replacement `cells` grid. Do **not** use `replace_text` on a table selection; Word will error.

Example follow-up after a months/days table exists:

```
edit the table column from number of days into the number of public holidays in usa
```

Expected flow: `list_tables` → `update_table` with `table_index`, matching `rows`/`columns`, and updated `cells`.

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
| `npm run dev` | Webpack dev server on port 3000 (HTTPS) |
| `npm start` | Sideload add-in in Word + start dev server |
| `npm stop` | Unregister sideloaded add-in |
| `npm run build` | Production ES5 bundle → `dist/` |
| `npm run typecheck` | TypeScript check (no emit) |
| `npm run validate` | Validate `manifest.xml` |
| `npm run validate:prod` | Validate `manifest.prod.xml` template |
| `npm run package` | Build and assemble `package/` for deployment |
| `npm run package -- https://host/path` | Package with production manifest URLs filled in |
| `npm run proxy` | Start local CORS proxy on port 8787 (see below) |
| `npm run smoke` | Automated smoke test (build, validate, package, dev server) |
| `npm run certs` | Install trusted localhost dev certs (Windows — run terminal as Administrator) |
| `npm run certs:verify` | Verify dev certificates are installed |

## Project structure

```
msword-aichat/
├── manifest.xml              # Word add-in manifest (sideload / dev)
├── manifest.prod.xml         # Production manifest template
├── taskpane.template.html    # Webpack HTML template
├── webpack.config.cjs        # ES5 build (IE11)
├── commands.html             # Ribbon commands stub (required by manifest)
├── public/assets/            # Add-in icons
└── src/
    ├── agent/                # Orchestrator, tools, system prompt
    ├── llm/                  # OpenAI & Anthropic-compatible providers + XHR SSE
    ├── word/                 # Document context + Office.js operations
    ├── settings/             # defaults + legacy pub/sub store
    ├── hooks/                # useChat.legacy, useDocumentContext.legacy
    ├── types/                # Shared TypeScript types
    └── taskpane/
        ├── App.legacy.tsx
        ├── main.legacy.tsx
        └── components.legacy/  # Fluent UI v8 task-pane UI
```

## Architecture (high level)

```
Word (Office.js)  ←→  Task Pane UI (React 16 + Fluent v8)
         ↓                    ↓
   Document tools        Agent orchestrator
   + context                  ↓
                          LLM Provider  →  Your API endpoint
```

- **UI:** React 16.14, Fluent UI v8, custom pub/sub settings store
- **Build:** Webpack 5, Babel ES5 (`targets: { ie: 11 }`), single `taskpane.*.bundle.js`
- **Streaming:** XHR-based SSE reader for IE11 chat mode
- **LLM:** Thin `fetch` + SSE adapters; no SDK dependency

## Security notes

- API keys are stored in `localStorage` on the local machine only
- Document context and tool reads are sent to your configured LLM endpoint
- Edits are staged for approval by default; only **Apply** writes to the document
- Context is truncated at ~12,000 characters to limit prompt size
- Edits use internal bookmarks (`msword_aichat_*`) for stable apply; orphan bookmarks may remain in the document
- **Undo** for `insert_table` removes the last table in the document (best-effort)
- **Undo** for `update_table` restores the previous cell contents at the same table index
- **`replace_text` on table selections** is blocked; use `update_table` for table edits
- Custom endpoints must be reachable from the add-in runtime (watch CORS if not using a same-origin proxy)
- Use HTTPS endpoints in production; local dev uses trusted localhost certs after `npm run certs`

## Distribution

Build and create a deployable folder:

```bash
npm run package -- https://addins.yourcompany.com/msword-aichat
```

Upload the contents of `package/` to your HTTPS origin (bundle `.js` must sit next to `taskpane.html`). Use `manifest.xml` from that folder for sideloading or Microsoft 365 admin center deployment.

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

## Word 2016 Windows QA matrix (P0 sign-off)

Manual smoke test on **Word 2016 desktop** before org-wide rollout:

| # | Test | Pass |
|---|------|------|
| 1 | Task pane loads UI (not blank) | |
| 2 | Settings save + reload | |
| 3 | Test connection | |
| 4 | Chat stream (XHR SSE) | |
| 5 | Selection context + refresh | |
| 6 | Agent `get_selection` | |
| 7 | `replace_text` preview → Apply → Undo | |
| 8 | `insert_comment` | |
| 9 | `insert_table` with filled cells | |
| 10 | `update_table` in place (no duplicate table) | |
| 11 | Slash `/fix` and `/table` | |
| 12 | Conversation persists per document | |
| 13 | Retry / copy on error | |

## Word on the web QA checklist

Regression test in Word on the web after Word 2016 sign-off:

| Area | Test |
|------|------|
| Load | Task pane opens from Home → AI Chat |
| Settings | Save provider, test connection, fetch models |
| Chat | Stream a response with Selection context |
| Agent | Run a read-only tool (`get_selection`) |
| Edits | `replace_text` preview → Apply → Undo |
| Styles | `apply_style` on a heading |
| Tables | `insert_table` with 2×2 data; `update_table` changes header/values in place |
| Search | `search_document` returns matches |
| Errors | Retry and Copy on a failed request |

## Troubleshooting

**Blank white task pane (no UI)**

- Confirm you are on the `ie11-rewrite` branch with Webpack dev server running (`npm run dev`).
- Restart `npm run dev` after `npm run certs`; close Word before `npm start`.
- On Office 2019+ / M365, confirm [WebView2](https://developer.microsoft.com/microsoft-edge/webview2/) is installed if the pane is blank on modern Word.

**Red “Uncaught runtime errors” overlay (dev only)**

Webpack dev-server HMR can trigger this on IE11. HMR and the error overlay are disabled — restart `npm run dev` if you still see it.

**Fluent dropdown shows white screen (IE11)**

The legacy UI uses native `<select>` controls (`IeSelect`). Avoid Fluent `Dropdown` / `CommandBar` overflow menus in the task pane on Word 2016.

**Certificate error / "content is blocked" (Word task pane)**

Word uses its own embedded browser and does **not** trust self-signed certificates by default. Install Microsoft's dev certs once:

1. Open **PowerShell or Command Prompt as Administrator**
2. In the project folder: `npm install` then `npm run certs`
3. Stop and restart `npm run dev`
4. Close Word completely, then run `npm start` again
5. Open `https://localhost:3000/taskpane.html` in **Edge** and confirm it loads without a cert warning

`npm run certs:verify` should report the certificate is valid.

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

**Table edit fails or creates a duplicate table**

- The agent should use `update_table`, not `insert_table` or `replace_text`, when changing an existing table.
- If the document has multiple tables, the agent should call `list_tables` and pass the correct `table_index`.
- Selecting the whole table and using `replace_text` is blocked (Word `GeneralException`); click inside the table and ask for an in-place update, or start a **New chat** so the agent picks up current tool guidance.

**Manifest validation**

```bash
npm run validate
```

## License

MIT — see [LICENSE](./LICENSE).