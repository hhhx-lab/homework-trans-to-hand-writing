<div align="center">

# Homework Handwriting Workbench

**A polished web application for turning homework documents into previewable, export-ready handwriting deliverables.**

![Web App](https://img.shields.io/badge/Web%20App-Vue%203-2563eb)
![Backend](https://img.shields.io/badge/Backend-FastAPI-059669)
![PDF](https://img.shields.io/badge/PDF-MinerU%20Source-7c3aed)
![Output](https://img.shields.io/badge/Output-PDF%20%7C%20DOCX-c2410c)
![QA](https://img.shields.io/badge/QA-Formula%20Aware-111827)

`./start-dev.sh`

</div>

> A document-to-handwriting cockpit for assignments, lecture notes, formula-heavy drafts, and printable study materials.

Homework Handwriting Workbench is built for a very specific workflow: take structured source material, normalize the text and math, preview it as handwriting, and export a clean PDF or Word document without losing the shape of the original content. It combines a Vue workbench, a FastAPI task backend, a MinerU-powered PDF extraction path, and handwriting rendering controls for fonts, paper, margins, line spacing, ink, correction marks, and full-document preview.

It is not a landing page. The first screen is the working surface.

## The 30-Second Version

```bash
./start-dev.sh
```

```text
Frontend: http://localhost:8080/
Backend:  http://127.0.0.1:5005
tmux:     handwriting-web
```

Use `./start-dev.sh --attach` to watch both panes, and `./start-dev.sh --stop` to stop the local stack.

## Why It Feels Different

| Usual path | This project |
|---|---|
| Copy text into a handwriting toy and hope formulas survive | Extract, normalize, preview, and export through a formula-aware pipeline |
| Check only the first generated page | Generate full preview pages and flip through the result before export |
| Silent PDF fallback paths that hide extraction problems | PDF source extraction is MinerU-driven and startup-validated |
| One scrolling page of controls and preview fighting for space | Left settings and right preview use independent scrolling |
| Custom font confusion from handwriting photos | The app expects real `.ttf` font files and documents the required coverage |

## Capability Matrix

| Capability | What it does | Why it matters |
|---|---|---|
| Source ingestion | Accepts pasted text plus PDF, Word, Markdown, TXT, and RTF uploads | Keeps the workflow flexible without forcing a single input format |
| MinerU PDF extraction | Sends PDFs through the configured MinerU service and normalizes the returned Markdown | Keeps PDF extraction source-grounded and avoids silent local fallback behavior |
| Math cleanup | Repairs common extraction glitches around decimals, small `c`, brackets, inequalities, `p`, `n`, and matrix fragments | Reduces the floating-character and broken-formula artifacts that make handwritten output look fake |
| Editable DOCX draft | Generates a standard Word proofing draft from Markdown | Lets you verify formulas and text before committing to handwriting export |
| Task queue | Uses async generation tasks, WebSocket updates, polling, and persisted results | Prevents long documents from freezing the UI |
| Full preview | Generates all handwriting pages and exposes page navigation | Lets you inspect layout page by page instead of trusting page one |
| Custom fonts | Loads bundled and host-mounted `.ttf` handwriting fonts | Makes handwriting style configurable without changing code |
| Private deployment | Provides local dev scripts and Docker Compose wiring | Supports repeatable local and private server deployment |

## Workflow

```text
Upload or paste source
  |
  v
Extract source text
  |-- PDF -> MinerU -> Markdown
  |-- DOCX/DOC/RTF -> Pandoc -> Markdown
  `-- MD/TXT -> Markdown
  |
  v
Normalize Markdown and math
  |
  v
Generate editable DOCX draft or handwriting preview
  |
  v
Flip through full preview pages
  |
  v
Export PDF or DOCX deliverable
```

## Previewed User Flow

| Screen area | What it is for |
|---|---|
| Left panel | Upload source, edit extracted Markdown, choose font, tune paper, margins, line spacing, randomness, ink, and correction marks |
| Right panel | Preview handwriting output, generate a full preview, flip pages, and inspect the deliverable before export |
| Single preview | Fast one-page check for style and spacing |
| Full preview | Multi-page generation path for final inspection |
| Export controls | Produce PDF or Word output from the current configuration |

## Quick Start

### 1. Prepare the backend environment

Use a project-local environment. Do not use `sudo pip`, and do not mix system Python, Homebrew Python, and the project environment.

```bash
cd backend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

The launcher uses `backend/.venv/bin/python` by default. If you manage Python with Conda or another isolated environment, point the script at it:

```bash
BACKEND_PYTHON=/path/to/python ./start-dev.sh
```

### 2. Install frontend dependencies

```bash
cd frontend
npm install
```

### 3. Configure `.env`

Create a `.env` file at the repository root. PDF upload requires MinerU configuration.

```dotenv
MINERU_BASE_URL=https://mineru.example.com/api/v4
MINERU_API_TOKEN=replace-with-your-token
MINERU_PUBLIC_BASE_URL=https://handwriting.example.com
MINERU_MODEL_VERSION=vlm
```

### 4. Start both services

```bash
./start-dev.sh
```

The script creates or restarts the `handwriting-web` tmux session, starts the backend on port `5005`, starts the frontend on port `8080`, and refuses to continue if those ports are already owned by unrelated processes.

## Manual Development Commands

Backend:

```bash
cd backend
.venv/bin/python app.py
```

Frontend:

```bash
cd frontend
npm run serve
```

Frontend build:

```bash
cd frontend
npm run build
```

Backend unit tests:

```bash
backend/.venv/bin/python -m unittest backend.tests.test_unified_handwriting_pipeline
```

## Environment Variables

| Variable | Required | Purpose |
|---|---:|---|
| `MINERU_BASE_URL` | PDF required | Base URL for the MinerU wrapper API, usually ending in `/api/v4` |
| `MINERU_API_TOKEN` | PDF required | Bearer token for MinerU requests |
| `MINERU_API_KEY` | Alternative | Alternative token name used when `MINERU_API_TOKEN` is not set |
| `MINERU_PUBLIC_BASE_URL` | PDF required | Public or private-network URL where MinerU can fetch staged PDFs from this backend |
| `MINERU_MODEL_VERSION` | Optional | MinerU model version, defaults to `vlm` |
| `MINERU_TRUST_ENV` | Optional | Set to `0` to bypass proxy environment variables for MinerU calls |
| `MINERU_BIND_HOST` | Optional | Bind outbound MinerU requests to a specific local network interface address |
| `MINERU_STARTUP_TIMEOUT_SECONDS` | Optional | Startup probe timeout, defaults to `10` |
| `MINERU_POLL_INTERVAL_SECONDS` | Optional | MinerU task polling interval, defaults to `2` |
| `MINERU_TIMEOUT_SECONDS` | Optional | MinerU extraction timeout, defaults to `600` |
| `FONT_ASSETS_DIR` | Optional | Host or container directory for custom font assets |
| `FONT_ASSETS_BUNDLED_DIR` | Optional | Bundled font directory used as an additional font source |
| `BACKEND_PYTHON` | Optional | Python executable used by `start-dev.sh` |
| `HANDWRITING_TMUX_SESSION` | Optional | tmux session name used by `start-dev.sh` |

Important MinerU rule: `MINERU_PUBLIC_BASE_URL` must be reachable by the MinerU service. `localhost`, `127.0.0.1`, and `::1` are rejected for this value because a remote or containerized MinerU process cannot fetch staged PDFs from those loopback addresses.

## API Surface

| Endpoint | Method | Purpose |
|---|---|---|
| `/api/generate_handwriting` | POST | Submit a handwriting generation task |
| `/api/generate_handwriting/ws/{task_id}` | WebSocket | Receive task status updates |
| `/api/generate_handwriting/task/{task_id}` | GET | Poll task status |
| `/api/generate_handwriting/task/{task_id}/result` | GET | Download completed task output |
| `/api/handwriting/extract_source` | POST | Extract uploaded PDF, Word, Markdown, TXT, or RTF into normalized Markdown |
| `/api/handwriting/mineru_files/{file_id}` | GET | Serve staged PDFs to MinerU |
| `/api/handwriting/markdown_docx` | POST | Generate an editable Word proofing draft from Markdown |
| `/api/generate_handwritten_document` | POST | Generate handwritten PDF or DOCX through the document pipeline |
| `/api/fonts_info` | GET | List available handwriting fonts |
| `/api/textfileprocess` | POST | Legacy text processing endpoint |
| `/api/imagefileprocess` | POST | Legacy image processing endpoint |

## Repository Layout

```text
.
|-- README.md
|-- start-dev.sh
|-- docker-compose.yml
|-- backend/
|   |-- app.py
|   |-- handwriting_document.py
|   |-- handwriting_markdown_renderer.py
|   |-- markdown_math.py
|   |-- mineru_adapter.py
|   |-- source_extract.py
|   |-- task_store.py
|   |-- requirements.txt
|   `-- tests/
|-- frontend/
|   |-- package.json
|   |-- vue.config.js
|   `-- src/
|-- ttf_files/
|-- mysql/
`-- serverless/
```

## Custom Font Notes

The upload control expects a real `.ttf` font file, not a handwriting photo.

| Need | Recommendation |
|---|---|
| English, numbers, simple formulas | Include digits, uppercase and lowercase letters, punctuation, brackets, and common math symbols |
| Chinese assignments | Include common Chinese characters, Chinese punctuation, digits, English letters, brackets, and math symbols |
| Better formula rendering | Make sure the font can display `p`, `n`, `c`, decimals, brackets, inequalities, plus/minus signs, and common Greek symbols |
| Only have handwriting photos | Build a `.ttf` with a font creation tool first, then upload the font file |

## Formula And PDF Quality Rules

This project is intentionally strict about document extraction and formula rendering.

| Rule | Behavior |
|---|---|
| PDF source | PDF extraction comes from MinerU and is validated at startup |
| Math normalization | Extracted Markdown is repaired before rendering |
| TeX residue | Internal TeX controls are stripped from the handwriting path |
| Matrix handling | Known malformed matrix blocks from MinerU are repaired before layout |
| Final inspection | Use full preview and DOCX draft generation before exporting a final PDF |

## Docker Compose

```bash
docker compose up --build -d
```

Default ports:

| Service | Port |
|---|---|
| Frontend | `2345:80` |
| Backend | `127.0.0.1:5005:5005` |

Font files placed in `ttf_files/` are mounted into the backend container.

## Troubleshooting

| Symptom | What to check |
|---|---|
| Backend refuses to start | Confirm `.env` has the MinerU variables required for PDF extraction |
| MinerU startup probe times out | Confirm MinerU is running, the base URL is correct, and the machine can reach it without a broken proxy/TUN route |
| MinerU cannot fetch the PDF | Confirm `MINERU_PUBLIC_BASE_URL` points to a URL reachable from the MinerU service |
| Private-network requests go through a proxy | Set `MINERU_TRUST_ENV=0`; if needed, also set `MINERU_BIND_HOST` to the correct local interface address |
| Script says a port is occupied | Stop the unrelated listener, or stop this project with `./start-dev.sh --stop` before restarting |
| Full preview shows stale content | Re-extract or re-upload the source, then regenerate full preview from the latest editor content |
| Custom font has missing glyphs | Use a broader `.ttf` that covers the target language, symbols, brackets, and math notation |
| Matrices or formulas still look wrong | Re-extract with the latest backend, generate the DOCX draft, and inspect the normalized Markdown before export |

## Boundaries

- Use this tool to format, proofread, preview, and render your own coursework or study materials. Always follow your course rules and verify mathematical correctness before submission.
- Do not commit `.env`, tokens, logs, generated PDFs, local databases, build artifacts, or debug outputs.
- Keep secrets in environment variables, not in source code.
- Use an isolated Python environment. Avoid `sudo pip` and avoid mixing Homebrew or system Python with the project runtime.
- Generated handwriting output should be visually inspected page by page before it is treated as a final deliverable.

## Star This If

Star this repo if you want a local-first handwriting workbench that treats document extraction, formula cleanup, preview, and final PDF export as one serious workflow instead of a pile of disconnected scripts.
