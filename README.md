# PowerPoint Notes Summarizer

A minimal web app for extracting PowerPoint speaker notes, summarizing them with an AI provider, comparing before/after notes, and downloading a copy of the deck with cleaner notes.

## Run Locally

```bash
python3 app.py
```

The server prints the local URL it is using. By default it binds to `127.0.0.1:8000` and automatically tries nearby ports if one is busy.

## Docker Deployment

Copy the example environment file and add the provider keys you want available on the server:

```bash
cp .env.example .env
```

For Docker, set these values in `.env`:

```bash
HOST=0.0.0.0
PORT=8000
DATA_DIR=/data
WORK_DIR=/tmp/powerpoint-notes-summarizer
ALLOW_SAVED_API_KEYS=1

GEMINI_API_KEY=
GOOGLE_API_KEY=
OPENAI_API_KEY=
OPENROUTER_API_KEY=
```

Build and run:

```bash
docker compose up --build
```

Open:

```text
http://localhost:8000
```

The compose file mounts a persistent Docker volume at `/data`. Saved API keys are stored server-side in that volume at `/data/provider-keys.json` with restrictive file permissions. The file is not committed to git.

### Codex Login In Docker

The Docker image installs Node.js, npm, LibreOffice, and the Codex CLI package so the `OpenAI Login` option can work inside the container. Codex login state is kept under `/data` through the persistent volume.

To use OpenAI Login in Docker, start the app, choose `OpenAI Login`, then complete the OAuth popup. If you prefer API keys, Gemini, OpenAI API, and OpenRouter API work without Codex login.

### Docker Notes

- LibreOffice is included so `.ppt` files can be converted to `.pptx`.
- Server-configured keys can be provided with `.env`, Docker secrets, or your hosting provider's environment-variable UI.
- If `ALLOW_SAVED_API_KEYS=0`, users can still paste one-off keys for a run, but the app will not save them server-side.
- For a public deployment, put the app behind authentication or rate limiting before enabling server-owned API keys for anonymous users.

## AI Provider Setup

The app includes a setup guide at:

```text
/provider-setup.html
```

Supported providers:

- Gemini using Gemini Flash.
- OpenAI Login using local OAuth through `codex exec`.
- OpenAI API using the Responses API.
- OpenRouter API using OpenRouter chat completions.

API keys are resolved in this order:

1. A one-off key entered for the current run.
2. A key saved on the server from the app UI.
3. Environment variables such as `GEMINI_API_KEY`, `OPENAI_API_KEY`, or `OPENROUTER_API_KEY`.

## What It Does

- Accepts `.pptx` files directly.
- Accepts `.ppt` files when LibreOffice or `soffice` is installed.
- Reads presentation slide order and speaker notes.
- Lets you choose the AI provider and model.
- Lets you edit the summarization prompt.
- Shows real progress through extraction, AI call, parsing, and deck reconstruction.
- Writes summarized notes back into the deck.
- Returns a `.pptx` download, or a `.zip` when multiple decks are uploaded.

## Tests

```bash
python3 -m unittest discover -s tests
```
