import tempfile
import unittest
import zipfile
from pathlib import Path

from pptx_notes import inspect_pptx, write_summarized_notes


class PptxNotesTest(unittest.TestCase):
    def test_extract_and_rewrite_notes(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "sample.pptx"
            output = Path(tmp) / "output.pptx"
            self._write_sample_pptx(source)

            inspection = inspect_pptx(source)
            self.assertEqual(inspection["slide_count"], 1)
            self.assertEqual(inspection["note_count"], 1)
            self.assertIn("Introduce the roadmap", inspection["slides"][0]["notes"])

            result = write_summarized_notes(
                source,
                output,
                inspection["slides"],
                {1: "## Slide 1\n- Introduce roadmap\n- Ask opening question"},
            )
            self.assertEqual(result["updated_count"], 1)

            updated = inspect_pptx(output)
            self.assertIn("## Slide 1", updated["slides"][0]["notes"])
            self.assertIn("- Ask opening question", updated["slides"][0]["notes"])

            with zipfile.ZipFile(source, "r") as original_deck, zipfile.ZipFile(output, "r") as updated_deck:
                self.assertEqual(
                    original_deck.read("ppt/slides/slide1.xml"),
                    updated_deck.read("ppt/slides/slide1.xml"),
                )
                updated_notes_xml = updated_deck.read("ppt/notesSlides/notesSlide1.xml")
                self.assertIn(b'typeface="Aptos"', updated_notes_xml)
                self.assertIn(b'sz="2400"', updated_notes_xml)
                self.assertNotIn(b'sz="1200"', updated_notes_xml)

    def _write_sample_pptx(self, path: Path):
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as deck:
            deck.writestr("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>""")
            deck.writestr(
                "ppt/presentation.xml",
                """<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>""",
            )
            deck.writestr(
                "ppt/_rels/presentation.xml.rels",
                """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>""",
            )
            deck.writestr(
                "ppt/slides/slide1.xml",
                """<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:bodyPr/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="3200"><a:latin typeface="Aptos Display"/></a:rPr>
              <a:t>Visible slide text keeps its font.</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>""",
            )
            deck.writestr(
                "ppt/slides/_rels/slide1.xml.rels",
                """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>""",
            )
            deck.writestr(
                "ppt/notesSlides/notesSlide1.xml",
                """<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:pPr marL="342900"><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Aptos"/></a:rPr><a:t>Introduce the roadmap for the product demo.</a:t></a:r></a:p>
          <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Aptos"/></a:rPr><a:t>Ask the audience what workflows feel slow today.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>""",
            )


if __name__ == "__main__":
    unittest.main()
