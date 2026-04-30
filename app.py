from __future__ import annotations

import json
import mimetypes
import os
import re
import select
import base64
import binascii
import io
import secrets
import shutil
import subprocess
import tempfile
import threading
import time
import zipfile
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import unquote, urlparse
from urllib.request import Request, urlopen

from pptx_notes import (
    PowerPointError,
    dump_session,
    inspect_pptx,
    load_session,
    prepare_powerpoint,
    write_summarized_notes,
)


def _load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line.removeprefix("export ").strip()
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'\"")
        if key and key not in os.environ:
            os.environ[key] = value


ROOT = Path(__file__).resolve().parent
_load_env_file(ROOT / ".env")
STATIC_DIR = ROOT / "static"
def _path_from_env(name: str, default: Path) -> Path:
    value = os.environ.get(name, "").strip()
    return Path(value).expanduser() if value else default


DATA_DIR = _path_from_env("DATA_DIR", ROOT / "data")
PROVIDER_KEYS_PATH = DATA_DIR / "provider-keys.json"
WORK_DIR = _path_from_env("WORK_DIR", Path(tempfile.gettempdir()) / "powerpoint-notes-summarizer")
MAX_UPLOAD_BYTES = 80 * 1024 * 1024
MAX_UPLOAD_FILES = 5
ALLOW_SAVED_API_KEYS = os.environ.get("ALLOW_SAVED_API_KEYS", "1").strip().lower() not in {"0", "false", "no"}
CODEX_LOGIN_SESSION: dict = {}
SUMMARIZE_JOBS: dict[str, dict] = {}
SUMMARIZE_JOBS_LOCK = threading.Lock()
SUMMARIZE_JOB_RETENTION_SECONDS = 60 * 60
PROVIDER_KEYS_LOCK = threading.RLock()

STANDARD_PROMPT = """Transform the provided detailed slide notes into concise bullet-point reminders for live presentation delivery.

INSTRUCTIONS:

- Convert verbose explanations into short, memorable bullet points
- Focus on key topics, concepts, and transitions the presenter needs to remember
- Preserve technical terms, product names, and specific numbers/statistics
- Preserve the original slide order and structure
- Include any instructor notes or presenter directions in [brackets]
- Maintain logical flow between points
- Remove filler words and redundant explanations


BULLET POINT GUIDELINES:


- Do not include slide numbers, slide headings, or section titles in the final notes
- DO NOT use any formatting in the output text, this text will be copied and pasted in a .txt file, avoid rich text. Do bullet point with a "- "
- Maximum 5-10 words per bullet when possible
- Use action verbs and keywords as memory triggers
- Include specific examples, names, or statistics that must be mentioned
- Note key transitions between sections
- Highlight any demonstrations, clicks, or interactive elements
- Keep audience engagement cues (questions, discussion points)

GOAL: Create speaking notes that allow the presenter to glance quickly and speak naturally while covering all essential content without reading verbatim."""

PROVIDERS = {
    "gemini": {
        "label": "Google Gemini",
        "shortLabel": "Gemini",
        "requiresKey": True,
        "keyLabel": "Gemini API key",
        "envKey": "GEMINI_API_KEY",
        "altEnvKeys": ["GOOGLE_API_KEY"],
        "speedLabel": "free",
        "defaultModel": "gemini-flash-latest",
        "models": ["gemini-flash-latest"],
        "keyUrl": "https://aistudio.google.com/app/apikey",
        "docsUrl": "https://ai.google.dev/gemini-api/docs/api-key?hl=en",
        "setupSummary": "Create a free-friendly Gemini key in Google AI Studio.",
    },
    "codex": {
        "label": "OpenAI Codex login",
        "shortLabel": "OpenAI Login",
        "requiresKey": False,
        "keyLabel": "",
        "envKey": "",
        "altEnvKeys": [],
        "speedLabel": "slower",
        "defaultModel": "gpt-5.4-mini",
        "models": ["gpt-5.4-mini", "gpt-5.4", "gpt-5.5", "gpt-5.2", "gpt-5.3-codex"],
        "keyUrl": "",
        "docsUrl": "https://developers.openai.com/codex/cli",
        "setupSummary": "Sign in locally with Codex OAuth when running this app on your own machine.",
    },
    "openai": {
        "label": "OpenAI API key",
        "shortLabel": "OpenAI API",
        "requiresKey": True,
        "keyLabel": "OpenAI API key",
        "envKey": "OPENAI_API_KEY",
        "altEnvKeys": [],
        "speedLabel": "faster",
        "defaultModel": "gpt-5.4-mini",
        "models": ["gpt-5.4-mini", "gpt-5.4", "gpt-5.5", "gpt-4.1-mini", "gpt-4.1"],
        "keyUrl": "https://platform.openai.com/api-keys",
        "docsUrl": "https://platform.openai.com/docs/quickstart",
        "setupSummary": "Use an OpenAI API key for high-quality summaries without OAuth.",
    },
    "openrouter": {
        "label": "OpenRouter API key",
        "shortLabel": "OpenRouter API",
        "requiresKey": True,
        "keyLabel": "OpenRouter API key",
        "envKey": "OPENROUTER_API_KEY",
        "altEnvKeys": [],
        "speedLabel": "faster",
        "defaultModel": "openai/gpt-5.2",
        "models": ["openai/gpt-5.2", "openai/gpt-5.1", "openai/gpt-5", "openai/gpt-4.1"],
        "keyUrl": "https://openrouter.ai/settings/keys",
        "docsUrl": "https://openrouter.ai/docs/api-reference/overview",
        "setupSummary": "Use one OpenRouter key to route summaries through many model providers.",
    },
}
DEFAULT_PROVIDER = "codex"

MODEL_DESCRIPTIONS = {
    "gemini-flash-latest": "Free Gemini Flash route",
    "gpt-5.4-mini": "Fast, lower cost, good for most decks",
    "gpt-5.4": "Higher quality, balanced speed",
    "gpt-5.5": "Highest quality, slower",
    "gpt-5.2": "Strong quality, dependable reasoning",
    "gpt-5.3-codex": "Codex-tuned, useful with OpenAI login",
    "gpt-4.1-mini": "Fast legacy option",
    "gpt-4.1": "High quality legacy option",
    "openai/gpt-5.2": "High quality through OpenRouter",
    "openai/gpt-5.1": "High quality through OpenRouter",
    "openai/gpt-5": "High quality through OpenRouter",
    "openai/gpt-4.1": "High quality legacy route",
}


class AppHandler(BaseHTTPRequestHandler):
    server_version = "PowerPointNotesSummarizer/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path.startswith("/api/download/"):
            self._download(parsed.path.removeprefix("/api/download/"))
            return
        if parsed.path.startswith("/api/summarize/progress/"):
            self._summarize_progress(parsed.path.removeprefix("/api/summarize/progress/"))
            return
        if parsed.path == "/api/provider-status":
            self._provider_status()
            return
        if parsed.path == "/api/codex-oauth/status":
            self._json(codex_oauth_status())
            return
        self._serve_static(parsed.path)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/analyze":
                self._analyze()
            elif parsed.path == "/api/summarize":
                self._summarize()
            elif parsed.path == "/api/summarize/start":
                self._start_summarize_job()
            elif parsed.path == "/api/provider-key":
                self._save_provider_key()
            elif parsed.path == "/api/codex-oauth/start":
                self._start_codex_oauth()
            elif parsed.path == "/api/cancel":
                self._cancel()
            else:
                self._json({"error": "Not found."}, HTTPStatus.NOT_FOUND)
        except PowerPointError as exc:
            self._json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except ValueError as exc:
            self._json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self._json({"error": f"Unexpected server error: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.end_headers()

    def do_DELETE(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path.startswith("/api/provider-key/"):
                self._delete_provider_key(parsed.path.removeprefix("/api/provider-key/"))
            else:
                self._json({"error": "Not found."}, HTTPStatus.NOT_FOUND)
        except ValueError as exc:
            self._json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self._json({"error": f"Unexpected server error: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def log_message(self, format: str, *args: object) -> None:
        return

    def _analyze(self) -> None:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            raise ValueError("Upload a PowerPoint file.")
        if content_length > MAX_UPLOAD_BYTES:
            raise ValueError("Upload up to 5 files, 80 MB total.")

        body = self.rfile.read(content_length)
        fields = _parse_multipart(self.headers.get("Content-Type", ""), body)
        uploads = [item for item in fields.get("file", []) if item.get("content")]
        if not uploads:
            raise ValueError("Upload a PowerPoint file.")
        if len(uploads) > MAX_UPLOAD_FILES:
            raise ValueError(f"Upload up to {MAX_UPLOAD_FILES} PowerPoint files at once.")

        session_id = secrets.token_urlsafe(16)
        session_dir = WORK_DIR / session_id
        session_dir.mkdir(parents=True, exist_ok=True)

        presentations = []
        for index, upload in enumerate(uploads, start=1):
            filename = _safe_filename(upload.get("filename") or f"presentation-{index}.pptx")
            suffix = Path(filename).suffix.lower()
            if suffix not in {".pptx", ".ppt"}:
                raise ValueError("Only .pptx and .ppt files are supported.")

            deck_dir = session_dir / f"presentation-{index}"
            deck_dir.mkdir(parents=True, exist_ok=True)
            input_path = deck_dir / filename
            input_path.write_bytes(upload["content"])

            pptx_path = prepare_powerpoint(input_path, deck_dir)
            inspection = inspect_pptx(pptx_path)
            output_filename = f"{Path(filename).stem}-summarized-notes.pptx"
            presentations.append(
                {
                    "id": str(index),
                    "filename": filename,
                    "input_path": str(input_path),
                    "pptx_path": str(pptx_path),
                    "output_path": str(deck_dir / output_filename),
                    "output_filename": output_filename,
                    "inspection": inspection,
                }
            )

        session = {
            "id": session_id,
            "presentations": presentations,
        }
        dump_session(session_dir / "session.json", session)
        public_presentations = [_public_presentation(presentation) for presentation in presentations]

        self._json(
            {
                "sessionId": session_id,
                "presentationCount": len(presentations),
                "filename": presentations[0]["filename"] if len(presentations) == 1 else f"{len(presentations)} presentations",
                "slideCount": sum(item["slideCount"] for item in public_presentations),
                "noteCount": sum(item["noteCount"] for item in public_presentations),
                "noteSlideCount": sum(item["noteSlideCount"] for item in public_presentations),
                "writableNoteCount": sum(item["writableNoteCount"] for item in public_presentations),
                "presentations": public_presentations,
                "prompt": STANDARD_PROMPT,
                "provider": DEFAULT_PROVIDER,
                "providers": _public_providers(),
                "allowSavedApiKeys": ALLOW_SAVED_API_KEYS,
                "limits": {"maxFiles": MAX_UPLOAD_FILES, "maxTotalMb": MAX_UPLOAD_BYTES // (1024 * 1024)},
            }
        )

    def _summarize(self) -> None:
        payload = self._read_json()
        self._json(_summarize_payload(payload))

    def _start_summarize_job(self) -> None:
        _cleanup_summarize_jobs()
        payload = self._read_json()
        job_id = secrets.token_urlsafe(16)
        _update_summarize_job(
            job_id,
            status="queued",
            percent=0,
            title="Queued",
            detail="Waiting to start",
        )
        thread = threading.Thread(target=_run_summarize_job, args=(job_id, payload), daemon=True)
        thread.start()
        self._json({"jobId": job_id})

    def _summarize_progress(self, job_id: str) -> None:
        try:
            safe_job_id = _safe_id(job_id)
        except ValueError:
            self._json({"error": "Summarization job not found."}, HTTPStatus.NOT_FOUND)
            return
        with SUMMARIZE_JOBS_LOCK:
            job = SUMMARIZE_JOBS.get(safe_job_id)
            public_job = dict(job) if job else None
        if not public_job:
            self._json({"error": "Summarization job not found."}, HTTPStatus.NOT_FOUND)
            return
        self._json(public_job)

    def _start_codex_oauth(self) -> None:
        self._json(start_codex_oauth())

    def _provider_status(self) -> None:
        self._json(_provider_status_payload())

    def _save_provider_key(self) -> None:
        payload = self._read_json()
        provider_id = (payload.get("provider") or "").strip()
        api_key = (payload.get("apiKey") or "").strip()
        self._json(_save_provider_api_key(provider_id, api_key))

    def _delete_provider_key(self, provider_id: str) -> None:
        self._json(_delete_provider_api_key(unquote(provider_id)))

    def _cancel(self) -> None:
        payload = self._read_json()
        session_id = _safe_id(payload.get("sessionId", ""))
        if session_id:
            shutil.rmtree(WORK_DIR / session_id, ignore_errors=True)
        self._json({"ok": True})

    def _download(self, download_path: str) -> None:
        parts = [part for part in unquote(download_path).split("/") if part]
        if not parts or len(parts) > 2:
            self._json({"error": "Download not found."}, HTTPStatus.NOT_FOUND)
            return
        session_id = _safe_id(parts[0])
        presentation_id = _safe_id(parts[1]) if len(parts) == 2 else ""
        session_path = WORK_DIR / session_id / "session.json"
        if not session_path.exists():
            self._json({"error": "Download not found."}, HTTPStatus.NOT_FOUND)
            return
        session = load_session(session_path)
        if not presentation_id and len(_session_presentations(session)) > 1:
            self._download_zip(session)
            return

        presentation = _presentation_for_request(session, presentation_id or None)
        output_path = Path(presentation["output_path"])
        if not output_path.exists():
            self._json({"error": "The summarized PowerPoint has not been created yet."}, HTTPStatus.NOT_FOUND)
            return

        data = output_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        self.send_header("Content-Disposition", f'attachment; filename="{presentation["output_filename"]}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _download_zip(self, session: dict) -> None:
        archive_buffer = io.BytesIO()
        written_names: set[str] = set()
        added_count = 0
        with zipfile.ZipFile(archive_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for presentation in _session_presentations(session):
                output_path = Path(presentation["output_path"])
                if not output_path.exists():
                    continue
                archive_name = presentation["output_filename"]
                if archive_name in written_names:
                    stem = Path(archive_name).stem
                    archive_name = f"{stem}-{presentation['id']}.pptx"
                written_names.add(archive_name)
                archive.write(output_path, archive_name)
                added_count += 1

        if not added_count:
            self._json({"error": "No summarized PowerPoints have been created yet."}, HTTPStatus.NOT_FOUND)
            return

        data = archive_buffer.getvalue()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/zip")
        self.send_header("Content-Disposition", 'attachment; filename="summarized-presentations.zip"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _serve_static(self, path: str) -> None:
        if path in {"", "/"}:
            target = STATIC_DIR / "index.html"
        else:
            target = (STATIC_DIR / path.lstrip("/")).resolve()
            if STATIC_DIR.resolve() not in target.parents:
                self._json({"error": "Not found."}, HTTPStatus.NOT_FOUND)
                return

        if not target.exists() or not target.is_file():
            self._json({"error": "Not found."}, HTTPStatus.NOT_FOUND)
            return

        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        data = target.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _read_json(self) -> dict:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            return {}
        raw = self.rfile.read(content_length)
        return json.loads(raw.decode("utf-8"))

    def _json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def _run_summarize_job(job_id: str, payload: dict) -> None:
    def progress(percent: int, title: str, detail: str) -> None:
        _update_summarize_job(
            job_id,
            status="running",
            percent=percent,
            title=title,
            detail=detail,
        )

    try:
        result = _summarize_payload(payload, progress)
        _update_summarize_job(
            job_id,
            status="complete",
            percent=100,
            title="Complete",
            detail="Deck ready to download",
            result=result,
        )
    except Exception as exc:
        _update_summarize_job(
            job_id,
            status="error",
            title="Summarization failed",
            detail=str(exc),
            error=str(exc),
        )


def _update_summarize_job(job_id: str, **updates: object) -> None:
    with SUMMARIZE_JOBS_LOCK:
        job = SUMMARIZE_JOBS.setdefault(job_id, {"jobId": job_id})
        job.update(updates)
        job["jobId"] = job_id
        job["updatedAt"] = time.time()


def _cleanup_summarize_jobs() -> None:
    cutoff = time.time() - SUMMARIZE_JOB_RETENTION_SECONDS
    with SUMMARIZE_JOBS_LOCK:
        expired = [
            job_id
            for job_id, job in SUMMARIZE_JOBS.items()
            if job.get("status") in {"complete", "error"} and float(job.get("updatedAt", 0)) < cutoff
        ]
        for job_id in expired:
            SUMMARIZE_JOBS.pop(job_id, None)


def _emit_progress(progress: object, percent: int, title: str, detail: str) -> None:
    if callable(progress):
        progress(percent, title, detail)


def _summarize_payload(payload: dict, progress: object = None) -> dict:
    _emit_progress(progress, 4, "Loading presentation", "Opening upload session")
    session_id = _safe_id(payload.get("sessionId", ""))
    presentation_id = _safe_id(str(payload.get("presentationId", "") or ""))
    provider = (payload.get("provider") or DEFAULT_PROVIDER).strip()
    provider_config = PROVIDERS.get(provider)
    if not provider_config:
        raise ValueError("Choose a valid inference provider.")
    provider_label = provider_config["shortLabel"]
    model = (payload.get("model") or provider_config["defaultModel"]).strip()
    prompt = (payload.get("prompt") or STANDARD_PROMPT).strip()
    api_key = (payload.get("apiKey") or "").strip()
    if not model:
        raise ValueError("Choose a model.")
    if not prompt:
        raise ValueError("Add a summarization prompt.")

    session_dir = WORK_DIR / session_id
    session_path = session_dir / "session.json"
    if not session_path.exists():
        raise ValueError("This upload session is no longer available.")
    session = load_session(session_path)
    presentation = _presentation_for_request(session, presentation_id or None)
    _emit_progress(progress, 10, "Loading notes", f"Reading extracted notes from {presentation['filename']}")
    slides = presentation["inspection"]["slides"]
    slides_with_notes = [slide for slide in slides if slide.get("notes", "").strip() and slide.get("can_write")]
    if not slides_with_notes:
        raise ValueError("No writable speaker notes were found in this presentation.")

    _emit_progress(progress, 18, "Preparing prompt", f"{len(slides_with_notes)} note slides ready")
    notes_payload = _format_notes_for_model(slides_with_notes)
    model_input = _build_model_input(prompt, notes_payload, [slide["number"] for slide in slides_with_notes])
    _emit_progress(progress, 32, f"Sending to {provider_label}", f"Waiting for {model}")
    model_text = _call_inference_provider(provider, model, model_input, api_key)
    _emit_progress(progress, 66, f"{provider_label} responded", "Parsing returned notes")
    parsed = _parse_slide_sections(model_text, [slide["number"] for slide in slides_with_notes])

    _emit_progress(progress, 74, "Matching slides", "Mapping AI notes back to slide numbers")
    summaries_by_slide: dict[int, str] = {}
    warnings: list[str] = []
    for slide in slides_with_notes:
        number = slide["number"]
        summary = parsed.get(number)
        if summary:
            summaries_by_slide[number] = summary
        else:
            warnings.append(f"Slide {number} was not returned by the model and was left unchanged.")

    if not summaries_by_slide:
        raise ValueError("The model response could not be matched back to slide numbers.")

    _emit_progress(progress, 86, "Rebuilding deck", "Writing summarized notes into PowerPoint")
    write_result = write_summarized_notes(
        Path(presentation["pptx_path"]),
        Path(presentation["output_path"]),
        slides,
        summaries_by_slide,
    )
    warnings.extend(write_result["warnings"])
    summary = {
        "model": model,
        "updated_count": write_result["updated_count"],
        "warnings": warnings,
    }
    presentation["summary"] = summary
    if "presentations" not in session:
        session["summary"] = summary

    _emit_progress(progress, 94, "Saving results", "Preparing download and comparison")
    dump_session(session_path, session)

    comparison = []
    for slide in slides_with_notes:
        number = slide["number"]
        comparison.append(
            {
                "number": number,
                "originalNotes": slide.get("notes", ""),
                "summarizedNotes": summaries_by_slide.get(number, ""),
                "changed": number in summaries_by_slide,
            }
        )

    return {
        "sessionId": session_id,
        "presentationId": presentation["id"],
        "filename": presentation["filename"],
        "provider": provider,
        "updatedCount": write_result["updated_count"],
        "warnings": warnings,
        "downloadUrl": f"/api/download/{session_id}/{presentation['id']}",
        "comparison": comparison,
    }


def _call_inference_provider(provider: str, model: str, model_input: str, api_key: str) -> str:
    if provider == "codex":
        if not _codex_login_is_connected():
            raise ValueError("Connect your OpenAI account with OpenAI login before summarizing.")
        return _call_codex(model, model_input)
    if provider == "openai":
        return _call_openai_api(model, model_input, api_key)
    if provider == "openrouter":
        return _call_openrouter(model, model_input, api_key)
    if provider == "gemini":
        return _call_gemini_api(model, model_input, api_key)
    raise ValueError("Choose a valid inference provider.")


def _call_codex(model: str, model_input: str) -> str:
    codex_bin = os.environ.get("CODEX_BIN") or shutil.which("codex")
    app_bundle_codex = Path("/Applications/Codex.app/Contents/Resources/codex")
    if not codex_bin and app_bundle_codex.exists():
        codex_bin = str(app_bundle_codex)
    if not codex_bin:
        raise ValueError("Codex CLI was not found. Install Codex or set CODEX_BIN.")

    output_handle = tempfile.NamedTemporaryFile("w", suffix=".txt", prefix="codex-summary-", delete=False)
    output_path = Path(output_handle.name)
    output_handle.close()

    command = [
        codex_bin,
        "exec",
        "--ephemeral",
        "--skip-git-repo-check",
        "--ignore-rules",
        "--sandbox",
        "read-only",
        "--color",
        "never",
        "--output-last-message",
        str(output_path),
        "--model",
        model,
        "-",
    ]

    try:
        result = subprocess.run(
            command,
            input=model_input,
            text=True,
            capture_output=True,
            cwd=ROOT,
            timeout=600,
            check=False,
        )
        output_text = output_path.read_text(encoding="utf-8").strip() if output_path.exists() else ""
        if result.returncode != 0:
            detail = output_text or result.stderr.strip() or result.stdout.strip() or "Codex exited without details."
            raise ValueError(f"Codex inference failed: {detail}")
        if output_text:
            return output_text
        fallback = result.stdout.strip()
        if fallback:
            return fallback
        raise ValueError("Codex returned no summary text.")
    except subprocess.TimeoutExpired as exc:
        raise ValueError("Codex inference timed out.") from exc
    finally:
        output_path.unlink(missing_ok=True)


def start_codex_oauth() -> dict:
    if _codex_login_is_connected():
        return _connected_codex_payload()

    existing = CODEX_LOGIN_SESSION.get("process")
    if existing and existing.poll() is None:
        return {
            "status": "pending",
            "authUrl": CODEX_LOGIN_SESSION.get("auth_url"),
            "userCode": CODEX_LOGIN_SESSION.get("user_code"),
        }

    codex_bin = os.environ.get("CODEX_BIN") or shutil.which("codex")
    app_bundle_codex = Path("/Applications/Codex.app/Contents/Resources/codex")
    if not codex_bin and app_bundle_codex.exists():
        codex_bin = str(app_bundle_codex)
    if not codex_bin:
        raise ValueError("Codex CLI was not found. Install Codex or set CODEX_BIN.")

    process = subprocess.Popen(
        [codex_bin, "login", "--device-auth"],
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )

    output = _read_available_process_output(process, timeout_seconds=10)
    auth_url = _first_match(r"https://auth\.openai\.com/\S+", output)
    user_code = _first_match(r"\b[A-Z0-9]{4,5}-[A-Z0-9]{4,8}\b", output)

    if not auth_url or not user_code:
        process.terminate()
        detail = _clean_terminal_text(output).strip() or "Codex did not return an OAuth device code."
        raise ValueError(detail)

    CODEX_LOGIN_SESSION.clear()
    CODEX_LOGIN_SESSION.update(
        {
            "process": process,
            "auth_url": auth_url,
            "user_code": user_code,
            "output": output,
        }
    )

    return {"status": "pending", "authUrl": auth_url, "userCode": user_code}


def codex_oauth_status() -> dict:
    process = CODEX_LOGIN_SESSION.get("process")
    if process and process.poll() is None:
        return {
            "status": "pending",
            "authUrl": CODEX_LOGIN_SESSION.get("auth_url"),
            "userCode": CODEX_LOGIN_SESSION.get("user_code"),
        }

    if process:
        output = CODEX_LOGIN_SESSION.get("output", "")
        output += _read_available_process_output(process, timeout_seconds=0)
        return_code = process.poll()
        CODEX_LOGIN_SESSION.clear()
        if return_code == 0:
            return _connected_codex_payload()
        detail = _clean_terminal_text(output).strip() or "OpenAI account connection did not complete."
        return {"status": "error", "error": detail}

    if _codex_login_is_connected():
        return _connected_codex_payload()
    return {"status": "disconnected"}


def _connected_codex_payload() -> dict:
    return {"status": "connected", "account": _codex_account_info()}


def _codex_login_is_connected() -> bool:
    connected, _ = _codex_login_status()
    return connected


def _codex_login_status() -> tuple[bool, str]:
    codex_bin = os.environ.get("CODEX_BIN") or shutil.which("codex")
    app_bundle_codex = Path("/Applications/Codex.app/Contents/Resources/codex")
    if not codex_bin and app_bundle_codex.exists():
        codex_bin = str(app_bundle_codex)
    if not codex_bin:
        return False, ""
    try:
        result = subprocess.run(
            [codex_bin, "login", "status"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False, ""
    status_text = _clean_login_status_text(result.stdout + result.stderr)
    return result.returncode == 0 and "Logged in" in status_text, status_text


def _codex_account_info() -> dict:
    connected, status_text = _codex_login_status()
    info: dict[str, object] = {}
    if status_text:
        info["statusText"] = status_text
    if not connected:
        return info

    auth_path = Path.home() / ".codex" / "auth.json"
    try:
        auth_data = json.loads(auth_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return info

    auth_mode = auth_data.get("auth_mode")
    if isinstance(auth_mode, str) and auth_mode:
        info["authMode"] = _humanize_compact_label(auth_mode)
    last_refresh = auth_data.get("last_refresh")
    if isinstance(last_refresh, str) and last_refresh:
        info["lastRefresh"] = last_refresh

    tokens = auth_data.get("tokens")
    if not isinstance(tokens, dict):
        tokens = {}
    id_claims = _decode_jwt_claims(tokens.get("id_token"))
    access_claims = _decode_jwt_claims(tokens.get("access_token"))
    auth_claims = _merge_dicts(
        id_claims.get("https://api.openai.com/auth"),
        access_claims.get("https://api.openai.com/auth"),
    )
    profile_claims = _merge_dicts(
        id_claims.get("https://api.openai.com/profile"),
        access_claims.get("https://api.openai.com/profile"),
    )

    for key, value in {
        "name": id_claims.get("name"),
        "email": profile_claims.get("email") or id_claims.get("email"),
        "accountId": tokens.get("account_id") or auth_claims.get("chatgpt_account_id"),
        "plan": _humanize_compact_label(auth_claims.get("chatgpt_plan_type")),
        "subscriptionActiveUntil": auth_claims.get("chatgpt_subscription_active_until"),
        "subscriptionLastChecked": auth_claims.get("chatgpt_subscription_last_checked"),
        "authProvider": _humanize_compact_label(id_claims.get("auth_provider")),
    }.items():
        if isinstance(value, str) and value:
            info[key] = value

    email_verified = profile_claims.get("email_verified", id_claims.get("email_verified"))
    if isinstance(email_verified, bool):
        info["emailVerified"] = email_verified

    organization = _default_codex_organization(auth_claims.get("organizations"))
    if organization:
        info["organization"] = organization

    return info


def _merge_dicts(*values: object) -> dict:
    merged: dict = {}
    for value in values:
        if isinstance(value, dict):
            merged.update(value)
    return merged


def _decode_jwt_claims(value: object) -> dict:
    if not isinstance(value, str) or value.count(".") < 2:
        return {}
    try:
        payload = value.split(".")[1]
        payload += "=" * ((4 - len(payload) % 4) % 4)
        decoded = base64.urlsafe_b64decode(payload.encode("utf-8"))
        claims = json.loads(decoded.decode("utf-8"))
    except (binascii.Error, UnicodeDecodeError, ValueError, json.JSONDecodeError):
        return {}
    return claims if isinstance(claims, dict) else {}


def _default_codex_organization(value: object) -> dict | None:
    if not isinstance(value, list):
        return None
    organizations = [item for item in value if isinstance(item, dict)]
    if not organizations:
        return None
    organization = next((item for item in organizations if item.get("is_default")), organizations[0])
    title = organization.get("title")
    role = organization.get("role")
    if not isinstance(title, str) or not title:
        return None
    result = {"title": title}
    if isinstance(role, str) and role:
        result["role"] = _humanize_compact_label(role)
    return result


def _humanize_compact_label(value: object) -> str | None:
    if not isinstance(value, str) or not value:
        return None
    known = {
        "chatgpt": "ChatGPT",
        "pro": "Pro",
        "plus": "Plus",
        "team": "Team",
        "enterprise": "Enterprise",
        "free": "Free",
        "apple": "Apple",
        "google": "Google",
        "microsoft": "Microsoft",
    }
    normalized = value.strip().lower()
    if normalized in known:
        return known[normalized]
    return " ".join(part.capitalize() for part in re.split(r"[_\s-]+", value.strip()) if part)


def _clean_login_status_text(text: str) -> str:
    clean = _clean_terminal_text(text)
    lines = [
        line.strip()
        for line in clean.splitlines()
        if line.strip() and not line.strip().startswith("WARNING:")
    ]
    return "\n".join(lines)


def _read_available_process_output(process: subprocess.Popen, timeout_seconds: float) -> str:
    if process.stdout is None:
        return ""

    chunks: list[str] = []
    deadline = time.monotonic() + timeout_seconds
    while True:
        remaining = deadline - time.monotonic()
        if timeout_seconds > 0 and remaining <= 0:
            break
        wait_time = max(0, min(0.2, remaining)) if timeout_seconds > 0 else 0
        readable, _, _ = select.select([process.stdout], [], [], wait_time)
        if not readable:
            if timeout_seconds == 0:
                break
            continue
        line = process.stdout.readline()
        if not line:
            break
        chunks.append(line)
        combined = "".join(chunks)
        if "https://auth.openai.com/" in combined and _first_match(r"\b[A-Z0-9]{4,5}-[A-Z0-9]{4,8}\b", combined):
            break
    return "".join(chunks)


def _first_match(pattern: str, text: str) -> str | None:
    clean = _clean_terminal_text(text)
    match = re.search(pattern, clean)
    return match.group(0) if match else None


def _clean_terminal_text(text: str) -> str:
    return re.sub(r"\x1b\[[0-9;]*m", "", text)


def _call_openai_api(model: str, model_input: str, api_key: str) -> str:
    key = _resolve_provider_api_key("openai", api_key)
    data = _post_json(
        "https://api.openai.com/v1/responses",
        key,
        {"model": model, "input": model_input},
    )
    text = data.get("output_text")
    if text:
        return text.strip()

    fragments: list[str] = []
    for item in data.get("output", []):
        for content in item.get("content", []):
            if content.get("type") in {"output_text", "text"} and content.get("text"):
                fragments.append(content["text"])
    if fragments:
        return "\n".join(fragments).strip()
    raise ValueError("OpenAI returned no summary text.")


def _call_openrouter(model: str, model_input: str, api_key: str) -> str:
    key = _resolve_provider_api_key("openrouter", api_key)
    data = _post_json(
        "https://openrouter.ai/api/v1/chat/completions",
        key,
        {
            "model": model,
            "messages": [{"role": "user", "content": model_input}],
        },
        {
            "HTTP-Referer": "http://127.0.0.1",
            "X-OpenRouter-Title": "PowerPoint Notes Summarizer",
        },
    )
    choices = data.get("choices") or []
    if not choices:
        raise ValueError("OpenRouter returned no choices.")
    message = choices[0].get("message") or {}
    content = message.get("content") or choices[0].get("text")
    if isinstance(content, list):
        content = "\n".join(part.get("text", "") for part in content if isinstance(part, dict))
    if content:
        return str(content).strip()
    raise ValueError("OpenRouter returned no summary text.")


def _call_gemini_api(model: str, model_input: str, api_key: str) -> str:
    key = _resolve_provider_api_key("gemini", api_key)
    data = _post_json_with_headers(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        {"contents": [{"parts": [{"text": model_input}]}]},
        {
            "Content-Type": "application/json",
            "X-goog-api-key": key,
        },
    )
    fragments: list[str] = []
    for candidate in data.get("candidates", []):
        content = candidate.get("content") or {}
        for part in content.get("parts", []):
            text = part.get("text")
            if text:
                fragments.append(text)
    if fragments:
        return "\n".join(fragments).strip()
    raise ValueError("Gemini returned no summary text.")


def _post_json(url: str, bearer_token: str, payload: dict, extra_headers: dict | None = None) -> dict:
    headers = {
        "Authorization": f"Bearer {bearer_token}",
        "Content-Type": "application/json",
    }
    if extra_headers:
        headers.update(extra_headers)
    return _post_json_with_headers(url, payload, headers)


def _post_json_with_headers(url: str, payload: dict, headers: dict) -> dict:
    request = Request(url, data=json.dumps(payload).encode("utf-8"), method="POST", headers=headers)
    try:
        with urlopen(request, timeout=180) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise ValueError(f"Inference request failed: {detail}") from exc
    except URLError as exc:
        raise ValueError(f"Inference request failed: {exc.reason}") from exc


def _resolve_provider_api_key(provider_id: str, request_key: str = "") -> str:
    provider = PROVIDERS.get(provider_id)
    if not provider or not provider.get("requiresKey"):
        return request_key.strip()

    key = request_key.strip()
    if key:
        return key

    key = _saved_provider_key(provider_id)
    if key:
        return key

    key = _env_provider_key(provider_id)
    if key:
        return key

    env_names = _provider_env_names(provider)
    hint = f"Enter a {provider['keyLabel']}"
    if env_names:
        hint += f" or set {' or '.join(env_names)}"
    hint += "."
    raise ValueError(hint)


def _provider_env_names(provider: dict) -> list[str]:
    names = []
    env_key = provider.get("envKey")
    if env_key:
        names.append(env_key)
    names.extend(provider.get("altEnvKeys") or [])
    return names


def _env_provider_key(provider_id: str) -> str:
    provider = PROVIDERS.get(provider_id)
    if not provider:
        return ""
    for env_name in _provider_env_names(provider):
        value = os.environ.get(env_name, "").strip()
        if value:
            return value
    return ""


def _saved_provider_key(provider_id: str) -> str:
    if not ALLOW_SAVED_API_KEYS:
        return ""
    return str(_load_provider_keys().get(provider_id, "")).strip()


def _load_provider_keys() -> dict[str, str]:
    with PROVIDER_KEYS_LOCK:
        if not PROVIDER_KEYS_PATH.exists():
            return {}
        try:
            data = json.loads(PROVIDER_KEYS_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}
        if not isinstance(data, dict):
            return {}
        return {str(key): str(value) for key, value in data.items() if isinstance(value, str) and value.strip()}


def _write_provider_keys(keys: dict[str, str]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = PROVIDER_KEYS_PATH.with_suffix(".tmp")
    temp_path.write_text(json.dumps(keys, indent=2, sort_keys=True), encoding="utf-8")
    os.chmod(temp_path, 0o600)
    temp_path.replace(PROVIDER_KEYS_PATH)
    try:
        os.chmod(PROVIDER_KEYS_PATH, 0o600)
    except OSError:
        pass


def _save_provider_api_key(provider_id: str, api_key: str) -> dict:
    if not ALLOW_SAVED_API_KEYS:
        raise ValueError("Saving API keys is disabled on this deployment.")
    provider = PROVIDERS.get(provider_id)
    if not provider or not provider.get("requiresKey"):
        raise ValueError("Choose a provider that uses an API key.")
    if not api_key:
        raise ValueError("Paste an API key before saving.")

    with PROVIDER_KEYS_LOCK:
        keys = _load_provider_keys()
        keys[provider_id] = api_key
        _write_provider_keys(keys)
    return _provider_status_payload()


def _delete_provider_api_key(provider_id: str) -> dict:
    if not ALLOW_SAVED_API_KEYS:
        raise ValueError("Saving API keys is disabled on this deployment.")
    provider = PROVIDERS.get(provider_id)
    if not provider or not provider.get("requiresKey"):
        raise ValueError("Choose a provider that uses an API key.")

    with PROVIDER_KEYS_LOCK:
        keys = _load_provider_keys()
        keys.pop(provider_id, None)
        _write_provider_keys(keys)
    return _provider_status_payload()


def _provider_status_payload() -> dict:
    return {
        "allowSavedApiKeys": ALLOW_SAVED_API_KEYS,
        "providers": _public_providers(),
        "statuses": {provider_id: _provider_key_status(provider_id) for provider_id in PROVIDERS},
    }


def _provider_key_status(provider_id: str) -> dict:
    provider = PROVIDERS[provider_id]
    if not provider.get("requiresKey"):
        return {
            "configured": True,
            "source": "login",
            "label": "Uses login",
            "canSave": False,
        }

    saved = bool(_saved_provider_key(provider_id))
    env = bool(_env_provider_key(provider_id))
    source = "saved" if saved else "environment" if env else "missing"
    label = "Saved on this server" if saved else "Configured by environment" if env else "Needs setup"
    return {
        "configured": saved or env,
        "source": source,
        "label": label,
        "saved": saved,
        "environment": env,
        "canSave": ALLOW_SAVED_API_KEYS,
    }


def _build_model_input(prompt: str, notes_payload: str, slide_numbers: list[int]) -> str:
    slide_list = ", ".join(str(number) for number in slide_numbers)
    return f"""You are being used as a text summarization engine inside a local web app.
Do not run commands, inspect files, ask follow-up questions, or explain your process.
Return only the requested slide-note text.

{prompt}

Return exactly one section for each slide number listed here: {slide_list}.
Do not add commentary before or after the slide sections.
Keep each section title exactly as "## Slide [number]".
Use those section titles only so the app can match output back to slides; do not repeat slide numbers or headings inside the note text.

ORIGINAL SLIDE NOTES:

{notes_payload}
"""


def _format_notes_for_model(slides: list[dict]) -> str:
    sections = []
    for slide in slides:
        sections.append(f"## Slide {slide['number']}\n{slide.get('notes', '').strip()}")
    return "\n\n".join(sections)


def _parse_slide_sections(text: str, expected_numbers: list[int]) -> dict[int, str]:
    expected = set(expected_numbers)
    pattern = re.compile(r"(?im)^\s*(?:#{1,6}\s*)?Slide\s+(\d+)\s*:?\s*$")
    matches = list(pattern.finditer(text))

    if not matches and len(expected_numbers) == 1:
        number = expected_numbers[0]
        return {number: _normalize_summary(number, text)}

    sections: dict[int, str] = {}
    for index, match in enumerate(matches):
        number = int(match.group(1))
        if number not in expected:
            continue
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        body = text[start:end].strip()
        sections[number] = _normalize_summary(number, body)
    return sections


def _normalize_summary(number: int, body: str) -> str:
    lines = [line.rstrip() for line in body.strip().splitlines() if line.strip()]
    while lines and _is_slide_heading(lines[0]):
        lines.pop(0)
    return "\n".join(lines)


def _is_slide_heading(line: str) -> bool:
    return bool(re.fullmatch(r"\s*(?:#{1,6}\s*)?Slide\s+\d+\s*:?\s*", line, flags=re.IGNORECASE))


def _session_presentations(session: dict) -> list[dict]:
    presentations = session.get("presentations")
    if isinstance(presentations, list):
        return presentations
    return [
        {
            "id": "1",
            "filename": session["filename"],
            "input_path": session["input_path"],
            "pptx_path": session["pptx_path"],
            "output_path": session["output_path"],
            "output_filename": session["output_filename"],
            "inspection": session["inspection"],
            "summary": session.get("summary"),
        }
    ]


def _presentation_for_request(session: dict, presentation_id: str | None) -> dict:
    presentations = _session_presentations(session)
    if not presentations:
        raise ValueError("This upload session has no presentations.")
    if not presentation_id:
        return presentations[0]
    for presentation in presentations:
        if presentation.get("id") == presentation_id:
            return presentation
    raise ValueError("Presentation not found.")


def _parse_multipart(content_type: str, body: bytes) -> dict[str, list[dict]]:
    if "multipart/form-data" not in content_type:
        raise ValueError("Expected a multipart upload.")
    message = BytesParser(policy=policy.default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )
    fields: dict[str, list[dict]] = {}
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        fields.setdefault(name, []).append(
            {
                "filename": part.get_filename(),
                "content": part.get_payload(decode=True),
                "content_type": part.get_content_type(),
            }
        )
    return fields


def _public_presentation(presentation: dict) -> dict:
    inspection = presentation["inspection"]
    return {
        "id": presentation["id"],
        "filename": presentation["filename"],
        "slideCount": inspection["slide_count"],
        "noteCount": inspection["note_count"],
        "noteSlideCount": inspection["note_slide_count"],
        "writableNoteCount": inspection["writable_note_count"],
        "slides": _public_slides(inspection["slides"]),
    }


def _public_slides(slides: list[dict]) -> list[dict]:
    return [
        {
            "number": slide["number"],
            "hasNotes": bool(slide.get("notes", "").strip()),
            "canWrite": slide.get("can_write", False),
            "originalNotes": slide.get("notes", ""),
        }
        for slide in slides
    ]


def _public_providers() -> list[dict]:
    return [
        {
            "id": provider_id,
            "label": config["label"],
            "shortLabel": config["shortLabel"],
            "requiresKey": config["requiresKey"],
            "keyLabel": config["keyLabel"],
            "envKey": config["envKey"],
            "envKeys": _provider_env_names(config),
            "speedLabel": config["speedLabel"],
            "defaultModel": config["defaultModel"],
            "models": config["models"],
            "modelDescriptions": {model: MODEL_DESCRIPTIONS.get(model, "") for model in config["models"]},
            "keyUrl": config["keyUrl"],
            "docsUrl": config["docsUrl"],
            "setupSummary": config["setupSummary"],
            "keyStatus": _provider_key_status(provider_id),
        }
        for provider_id, config in PROVIDERS.items()
    ]


def _safe_filename(filename: str) -> str:
    name = Path(filename).name
    name = re.sub(r"[^A-Za-z0-9._ -]+", "-", name).strip(". ")
    return name or "presentation.pptx"


def _safe_id(value: str) -> str:
    if not value:
        return ""
    if not re.fullmatch(r"[A-Za-z0-9_-]+", value):
        raise ValueError("Invalid session id.")
    return value


def _make_server(host: str, preferred_port: int) -> tuple[ThreadingHTTPServer, int]:
    last_error: OSError | None = None
    for port in range(preferred_port, preferred_port + 50):
        try:
            server = ThreadingHTTPServer((host, port), AppHandler)
            return server, port
        except OSError as exc:
            last_error = exc
            continue
    detail = f" Last error: {last_error}" if last_error else ""
    raise RuntimeError(f"No available local port was found.{detail}")


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    host = os.environ.get("HOST", "127.0.0.1")
    preferred_port = int(os.environ.get("PORT", "8000"))
    server, port = _make_server(host, preferred_port)
    print(f"PowerPoint Notes Summarizer running at http://{host}:{port}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.", flush=True)
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
