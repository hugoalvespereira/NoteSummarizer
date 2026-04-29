import unittest

from app import _parse_multipart, _parse_slide_sections


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


if __name__ == "__main__":
    unittest.main()
