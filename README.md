# Word AI Chat

A Microsoft Word task-pane add-in that brings AI chat — and eventually agentic document editing — to Word. Connect any **OpenAI-compatible** or **Anthropic-compatible** endpoint (OpenAI, Anthropic, LiteLLM, vLLM, Ollama gateways, private proxies, etc.).

## Status

**Phase 0 complete.** The add-in loads in Word, persists provider settings locally, tests connectivity, and streams chat responses.

| Phase | Scope | Status |
|-------|-------|--------|
| 0 | Scaffold, settings, streaming chat | Done |
| 1 | Document context (selection, outline), quick actions | Planned |
| 2 | Agent loop + document edit tools | Planned |
| 3 | Advanced editing (styles, tables, undo) | Planned |
| 4 | Polish, Word on the web, distribution | Planned |

## Features (today)

- Task-pane UI on the Word **Home** tab (**AI Chat** button)
- Provider settings: type, base URL, API key, model, temperature, max tokens
- Connection test against your configured endpoint
- Streaming chat with conversation history (in-memory per session)
- API keys stored locally on the device (separate from other settings)

## Prerequisites

- **Node.js** 20+
- **Microsoft Word** (Desktop on Windows or Mac) for full add-in testing
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

Open [https://localhost:3000/taskpane.html](https://localhost:3000/taskpane.html) in a browser. Chat and settings work, but Office.js document APIs are unavailable outside Word.

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
    ├── llm/                  # OpenAI & Anthropic-compatible providers
    ├── settings/             # Zustand store + defaults
    ├── hooks/                # useChat streaming hook
    ├── types/                # Shared TypeScript types
    └── taskpane/             # React UI (App, Chat, Settings)
```

## Architecture (high level)

```
Word (Office.js)  ←→  Task Pane UI (React)
                           ↓
                     useChat hook
                           ↓
                     LLM Provider Adapter  →  Your API endpoint
```

- **UI:** React 19, Fluent UI v9, Zustand for settings
- **Build:** Vite 7 with `@vitejs/plugin-basic-ssl` (HTTPS on port 3000)
- **LLM:** Thin `fetch` + SSE adapters; no SDK dependency

## Security notes

- API keys are stored in `localStorage` on the local machine only
- Document text is **not** sent to the LLM yet (Phase 1 adds opt-in context)
- Custom endpoints must be reachable from the add-in runtime (watch CORS if not using a same-origin proxy)
- Use HTTPS endpoints in production; the dev server uses a self-signed certificate

## Roadmap

See [AGENTS.md](./AGENTS.md) for the full phased plan and implementation guide for contributors and coding agents.

**Next up (Phase 1):** read Word selection and document outline, show context in the chat UI, and add quick actions (summarize, improve, explain).

## Troubleshooting

**Add-in does not load**

- Confirm `npm run dev` is running and reachable at `https://localhost:3000`
- Accept the self-signed certificate warning in your browser once
- Re-run `npm start` to re-register the manifest

**Connection test fails**

- Verify base URL, API key, and model name
- Check gateway CORS headers if calling a remote proxy from the task pane
- Try the same request with `curl` against your endpoint

**Manifest validation**

```bash
npm run validate
```

## License

Private project — license TBD.