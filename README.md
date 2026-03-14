# office-agent

Office Agent is a local-first tool for turning natural-language prompts into PowerPoint `.pptx`
files. The project includes:

- a Python CLI and HTTP API for generation and rendering
- a React + TypeScript frontend for workflow visualization, theme switching, and bilingual UI

## Features

- OpenAI-compatible and Ollama providers
- Two-stage generation flow: prompt -> structured `PresentationSpec` -> `.pptx`
- Ten supported slide types:
  - `title`
  - `section`
  - `bullets`
  - `two_column`
  - `image`
  - `timeline`
  - `quote`
  - `comparison`
  - `summary`
  - `table`
- Theme system with preset switching and custom overrides
- Theme JSON export/import in the frontend
- Workflow visualization for input, LLM parsing, rendering, and final Office step
- Session-only API key entry in the frontend
- English and Simplified Chinese UI switching

## Theme Model

The shared theme schema supports:

```json
{
  "preset": "default | executive | editorial",
  "custom": {
    "primary_color": "#RRGGBB",
    "accent_color": "#RRGGBB",
    "background_color": "#RRGGBB",
    "heading_font": "Font Name",
    "body_font": "Font Name",
    "cover_style": "band | centered | minimal"
  }
}
```

Preset themes can be used directly, or overridden with custom values.

## Backend Setup

Install Python dependencies with `uv`:

```bash
uv sync
```

Run the CLI:

```bash
uv run office-agent generate --prompt "Create a project update deck" --out output.pptx --theme executive
```

Run the API server:

```bash
uv run office-agent-api
```

Available API routes:

- `GET /api/health`
- `POST /api/specs`
- `POST /api/presentations`

## Frontend Setup

Install frontend dependencies:

```bash
cd frontend
npm install
```

Run the Vite dev server:

```bash
npm run dev -- --port 5174
```

The frontend proxies `/api/*` requests to `http://127.0.0.1:8000`, so run the Python API server
at the same time during development.

## Environment Variables

- `OFFICE_AGENT_PROVIDER`
- `OFFICE_AGENT_MODEL`
- `OPENAI_API_KEY`
- `OPENAI_BASE_URL`
- `OLLAMA_BASE_URL`

The frontend does not persist API keys. If the user enters an API key in the browser, it is only
sent with that request and is not stored on disk.

## Documentation

- English: `README.md`
- Simplified Chinese: `README.zh-CN.md`

## Tests

Run backend tests:

```bash
uv run python -m unittest discover -s tests
```

Run frontend tests:

```bash
cd frontend
npm test
```
