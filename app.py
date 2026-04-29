from __future__ import annotations

import json
import mimetypes
import os
import re
import secrets
import shutil
import subprocess
import tempfile
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


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
WORK_DIR = Path(tempfile.gettempdir()) / "powerpoint-notes-summarizer"
MAX_UPLOAD_BYTES = 80 * 1024 * 1024

STANDARD_PROMPT = """Transform the provided detailed slide notes into concise bullet-point reminders for live presentation delivery.

INSTRUCTIONS:

- Convert verbose explanations into short, memorable bullet points
- Focus on key topics, concepts, and transitions the presenter needs to remember
- Preserve technical terms, product names, and specific numbers/statistics
- Keep the original slide numbering and structure
- Include any instructor notes or presenter directions in [brackets]
- Maintain logical flow between points
- Remove filler words and redundant explanations


BULLET POINT GUIDELINES:


- Keep the titles for each slide as "## Slide [number]" without any further text after the number, followed by the summarized slide notes in the next line
- DO NOT use any formatting in the output text, this text will be copied and pasted in a .txt file, avoid rich text. Do bullet point with a "- "
- Maximum 5-10 words per bullet when possible
- Use action verbs and keywords as memory triggers
- Include specific examples, names, or statistics that must be mentioned
- Note key transitions between sections
- Highlight any demonstrations, clicks, or interactive elements
- Keep audience engagement cues (questions, discussion points)

GOAL: Create speaking notes that allow the presenter to glance quickly and speak naturally while covering all essential content without reading verbatim."""

PROVIDERS = {
    "codex": {
        "label": "OpenAI Codex login",
        "shortLabel": "Codex OAuth",
        "requiresKey": False,
        "keyLabel": "",
        "envKey": "",
        "defaultModel": "gpt-5.4-mini",
        "models": ["gpt-5.4-mini", "gpt-5.4", "gpt-5.5", "gpt-5.2", "gpt-5.3-codex"],
    },
    "openai": {
        "label": "OpenAI API key",
        "shortLabel": "OpenAI Key",
        "requiresKey": True,
        "keyLabel": "OpenAI API key",
        "envKey": "OPENAI_API_KEY",
        "defaultModel": "gpt-5.4-mini",
        "models": ["gpt-5.4-mini", "gpt-5.4", "gpt-5.5", "gpt-4.1-mini", "gpt-4.1"],
    },
    "openrouter": {
        "label": "OpenRouter API key",
        "shortLabel": "OpenRouter Key",
        "requiresKey": True,
        "keyLabel": "OpenRouter API key",
        "envKey": "OPENROUTER_API_KEY",
        "defaultModel": "openai/gpt-5.2",
        "models": ["openai/gpt-5.2", "openai/gpt-5.1", "openai/gpt-5", "openai/gpt-4.1"],
    },
}
DEFAULT_PROVIDER = "codex"


class AppHandler(BaseHTTPRequestHandler):
    server_version = "PowerPointNotesSummarizer/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path.startswith("/api/download/"):
            self._download(parsed.path.rsplit("/", 1)[-1])
            return
        self._serve_static(parsed.path)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/analyze":
                self._analyze()
            elif parsed.path == "/api/summarize":
                self._summarize()
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

    def log_message(self, format: str, *args: object) -> None:
        return

    def _analyze(self) -> None:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            raise ValueError("Upload a PowerPoint file.")
        if content_length > MAX_UPLOAD_BYTES:
            raise ValueError("The uploaded file is larger than 80 MB.")

        body = self.rfile.read(content_length)
        fields = _parse_multipart(self.headers.get("Content-Type", ""), body)
        upload = fields.get("file")
        if not upload or not upload.get("content"):
            raise ValueError("Upload a PowerPoint file.")

        filename = _safe_filename(upload.get("filename") or "presentation.pptx")
        suffix = Path(filename).suffix.lower()
        if suffix not in {".pptx", ".ppt"}:
            raise ValueError("Only .pptx and .ppt files are supported.")

        session_id = secrets.token_urlsafe(16)
        session_dir = WORK_DIR / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        input_path = session_dir / filename
        input_path.write_bytes(upload["content"])

        pptx_path = prepare_powerpoint(input_path, session_dir)
        inspection = inspect_pptx(pptx_path)
        output_filename = f"{Path(filename).stem}-summarized-notes.pptx"
        session = {
            "id": session_id,
            "filename": filename,
            "input_path": str(input_path),
            "pptx_path": str(pptx_path),
            "output_path": str(session_dir / output_filename),
            "output_filename": output_filename,
            "inspection": inspection,
        }
        dump_session(session_dir / "session.json", session)

        self._json(
            {
                "sessionId": session_id,
                "filename": filename,
                "slideCount": inspection["slide_count"],
                "noteCount": inspection["note_count"],
                "noteSlideCount": inspection["note_slide_count"],
                "writableNoteCount": inspection["writable_note_count"],
                "slides": _public_slides(inspection["slides"]),
                "prompt": STANDARD_PROMPT,
                "provider": DEFAULT_PROVIDER,
                "providers": _public_providers(),
            }
        )

    def _summarize(self) -> None:
        payload = self._read_json()
        session_id = _safe_id(payload.get("sessionId", ""))
        provider = (payload.get("provider") or DEFAULT_PROVIDER).strip()
        provider_config = PROVIDERS.get(provider)
        if not provider_config:
            raise ValueError("Choose a valid inference provider.")
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
        slides = session["inspection"]["slides"]
        slides_with_notes = [slide for slide in slides if slide.get("notes", "").strip() and slide.get("can_write")]
        if not slides_with_notes:
            raise ValueError("No writable speaker notes were found in this deck.")

        notes_payload = _format_notes_for_model(slides_with_notes)
        model_input = _build_model_input(prompt, notes_payload, [slide["number"] for slide in slides_with_notes])
        model_text = _call_inference_provider(provider, model, model_input, api_key)
        parsed = _parse_slide_sections(model_text, [slide["number"] for slide in slides_with_notes])

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

        write_result = write_summarized_notes(
            Path(session["pptx_path"]),
            Path(session["output_path"]),
            slides,
            summaries_by_slide,
        )
        warnings.extend(write_result["warnings"])
        session["summary"] = {
            "model": model,
            "updated_count": write_result["updated_count"],
            "warnings": warnings,
        }
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

        self._json(
            {
                "sessionId": session_id,
                "filename": session["filename"],
                "provider": provider,
                "updatedCount": write_result["updated_count"],
                "warnings": warnings,
                "downloadUrl": f"/api/download/{session_id}",
                "comparison": comparison,
            }
        )

    def _cancel(self) -> None:
        payload = self._read_json()
        session_id = _safe_id(payload.get("sessionId", ""))
        if session_id:
            shutil.rmtree(WORK_DIR / session_id, ignore_errors=True)
        self._json({"ok": True})

    def _download(self, session_id: str) -> None:
        session_id = _safe_id(unquote(session_id))
        session_path = WORK_DIR / session_id / "session.json"
        if not session_path.exists():
            self._json({"error": "Download not found."}, HTTPStatus.NOT_FOUND)
            return
        session = load_session(session_path)
        output_path = Path(session["output_path"])
        if not output_path.exists():
            self._json({"error": "The summarized PowerPoint has not been created yet."}, HTTPStatus.NOT_FOUND)
            return

        data = output_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        self.send_header("Content-Disposition", f'attachment; filename="{session["output_filename"]}"')
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


def _call_inference_provider(provider: str, model: str, model_input: str, api_key: str) -> str:
    if provider == "codex":
        return _call_codex(model, model_input)
    if provider == "openai":
        return _call_openai_api(model, model_input, api_key)
    if provider == "openrouter":
        return _call_openrouter(model, model_input, api_key)
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
        "--ask-for-approval",
        "never",
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


def _call_openai_api(model: str, model_input: str, api_key: str) -> str:
    key = api_key or os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        raise ValueError("Enter an OpenAI API key or set OPENAI_API_KEY.")
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
    key = api_key or os.environ.get("OPENROUTER_API_KEY", "").strip()
    if not key:
        raise ValueError("Enter an OpenRouter API key or set OPENROUTER_API_KEY.")
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


def _post_json(url: str, bearer_token: str, payload: dict, extra_headers: dict | None = None) -> dict:
    headers = {
        "Authorization": f"Bearer {bearer_token}",
        "Content-Type": "application/json",
    }
    if extra_headers:
        headers.update(extra_headers)
    request = Request(url, data=json.dumps(payload).encode("utf-8"), method="POST", headers=headers)
    try:
        with urlopen(request, timeout=180) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise ValueError(f"Inference request failed: {detail}") from exc
    except URLError as exc:
        raise ValueError(f"Inference request failed: {exc.reason}") from exc


def _build_model_input(prompt: str, notes_payload: str, slide_numbers: list[int]) -> str:
    slide_list = ", ".join(str(number) for number in slide_numbers)
    return f"""You are being used as a text summarization engine inside a local web app.
Do not run commands, inspect files, ask follow-up questions, or explain your process.
Return only the requested slide-note text.

{prompt}

Return exactly one section for each slide number listed here: {slide_list}.
Do not add commentary before or after the slide sections.
Keep each section title exactly as "## Slide [number]".

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
    clean_body = "\n".join(line.rstrip() for line in body.strip().splitlines() if line.strip())
    if clean_body:
        return f"## Slide {number}\n{clean_body}"
    return f"## Slide {number}"


def _parse_multipart(content_type: str, body: bytes) -> dict[str, dict]:
    if "multipart/form-data" not in content_type:
        raise ValueError("Expected a multipart upload.")
    message = BytesParser(policy=policy.default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )
    fields: dict[str, dict] = {}
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        fields[name] = {
            "filename": part.get_filename(),
            "content": part.get_payload(decode=True),
            "content_type": part.get_content_type(),
        }
    return fields


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
            "defaultModel": config["defaultModel"],
            "models": config["models"],
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


def _make_server(preferred_port: int) -> tuple[ThreadingHTTPServer, int]:
    last_error: OSError | None = None
    for port in range(preferred_port, preferred_port + 50):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), AppHandler)
            return server, port
        except OSError as exc:
            last_error = exc
            continue
    detail = f" Last error: {last_error}" if last_error else ""
    raise RuntimeError(f"No available local port was found.{detail}")


def main() -> None:
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    preferred_port = int(os.environ.get("PORT", "8000"))
    server, port = _make_server(preferred_port)
    print(f"PowerPoint Notes Summarizer running at http://127.0.0.1:{port}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.", flush=True)
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
