import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app
from app import (
    _call_gemini_api,
    _delete_provider_api_key,
    _parse_multipart,
    _parse_slide_sections,
    _provider_status_payload,
    _public_providers,
    _resolve_provider_api_key,
    _save_provider_api_key,
)


class AppParsingTest(unittest.TestCase):
    def test_slide_headers_are_not_written_into_summaries(self):
        parsed = _parse_slide_sections(
            """## Slide 4
- Introduce roadmap
- Ask opening question

## Slide 5
Slide 5
- Show demo
- Pause for questions
""",
            [4, 5],
        )

        self.assertEqual(parsed[4], "- Introduce roadmap\n- Ask opening question")
        self.assertEqual(parsed[5], "- Show demo\n- Pause for questions")

    def test_single_slide_fallback_strips_header(self):
        parsed = _parse_slide_sections(
            """## Slide 1
- Summarize main point
- Close with next step
""",
            [1],
        )

        self.assertEqual(parsed[1], "- Summarize main point\n- Close with next step")

    def test_multipart_upload_keeps_repeated_file_fields(self):
        boundary = "deck-boundary"
        body = (
            f"--{boundary}\r\n"
            'Content-Disposition: form-data; name="file"; filename="one.pptx"\r\n'
            "Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation\r\n\r\n"
            "first\r\n"
            f"--{boundary}\r\n"
            'Content-Disposition: form-data; name="file"; filename="two.pptx"\r\n'
            "Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation\r\n\r\n"
            "second\r\n"
            f"--{boundary}--\r\n"
        ).encode("utf-8")

        fields = _parse_multipart(f"multipart/form-data; boundary={boundary}", body)

        self.assertEqual([item["filename"] for item in fields["file"]], ["one.pptx", "two.pptx"])
        self.assertEqual(fields["file"][1]["content"], b"second")

    def test_provider_order_and_gemini_metadata(self):
        providers = _public_providers()

        self.assertEqual([provider["id"] for provider in providers], ["gemini", "codex", "openai", "openrouter"])
        self.assertEqual([provider["shortLabel"] for provider in providers], ["Gemini", "OpenAI Login", "OpenAI API", "OpenRouter API"])
        self.assertEqual(providers[0]["speedLabel"], "free")
        self.assertEqual(providers[0]["envKey"], "GEMINI_API_KEY")
        self.assertIn("GOOGLE_API_KEY", providers[0]["envKeys"])

    @patch("app._post_json_with_headers")
    def test_gemini_call_uses_api_key_header_and_extracts_text(self, post_json):
        post_json.return_value = {
            "candidates": [
                {
                    "content": {
                        "parts": [
                            {"text": "## Slide 1\n- Test summary"},
                        ]
                    }
                }
            ]
        }

        result = _call_gemini_api("gemini-flash-latest", "Summarize this", "local-key")

        self.assertEqual(result, "## Slide 1\n- Test summary")
        post_json.assert_called_once_with(
            "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent",
            {"contents": [{"parts": [{"text": "Summarize this"}]}]},
            {
                "Content-Type": "application/json",
                "X-goog-api-key": "local-key",
            },
        )

    def test_saved_provider_key_is_used_without_echoing_secret_in_status(self):
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(app, "DATA_DIR", Path(tmp)), patch.object(
                app, "PROVIDER_KEYS_PATH", Path(tmp) / "provider-keys.json"
            ), patch.object(app, "ALLOW_SAVED_API_KEYS", True), patch.dict("os.environ", {}, clear=True):
                _save_provider_api_key("gemini", "saved-gemini-key")

                self.assertEqual(_resolve_provider_api_key("gemini"), "saved-gemini-key")
                payload = _provider_status_payload()
                gemini_status = payload["statuses"]["gemini"]
                self.assertTrue(gemini_status["configured"])
                self.assertEqual(gemini_status["source"], "saved")
                self.assertNotIn("saved-gemini-key", str(payload))

                _delete_provider_api_key("gemini")
                self.assertFalse(_provider_status_payload()["statuses"]["gemini"]["configured"])

    def test_environment_provider_key_is_used_when_no_saved_key_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(app, "PROVIDER_KEYS_PATH", Path(tmp) / "provider-keys.json"), patch.dict(
                "os.environ", {"GOOGLE_API_KEY": "env-gemini-key"}, clear=True
            ):
                self.assertEqual(_resolve_provider_api_key("gemini"), "env-gemini-key")


if __name__ == "__main__":
    unittest.main()
