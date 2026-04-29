from __future__ import annotations

import json
import os
import posixpath
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import unquote
from xml.etree import ElementTree as ET


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"

NS = {"p": P_NS, "a": A_NS, "r": R_NS, "rel": REL_NS}

ET.register_namespace("p", P_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("r", R_NS)


@dataclass
class SlideNote:
    number: int
    slide_path: str
    notes_path: str | None
    notes: str
    has_notes_slide: bool
    can_write: bool

    @property
    def has_text(self) -> bool:
        return bool(self.notes.strip())

    def to_public_dict(self) -> dict:
        data = asdict(self)
        data["has_text"] = self.has_text
        return data


class PowerPointError(RuntimeError):
    pass


def prepare_powerpoint(input_path: Path, work_dir: Path) -> Path:
    suffix = input_path.suffix.lower()
    if suffix == ".pptx":
        return input_path
    if suffix == ".ppt":
        return convert_ppt_to_pptx(input_path, work_dir)
    raise PowerPointError("Only .pptx and .ppt files are supported.")


def convert_ppt_to_pptx(input_path: Path, work_dir: Path) -> Path:
    soffice = _find_soffice()
    if not soffice:
        raise PowerPointError(
            ".ppt support requires LibreOffice or soffice on PATH. Convert this file to .pptx first, or install LibreOffice."
        )

    before = set(work_dir.glob("*.pptx"))
    command = [
        soffice,
        "--headless",
        "--convert-to",
        "pptx",
        "--outdir",
        str(work_dir),
        str(input_path),
    ]
    result = subprocess.run(
        command,
        cwd=work_dir,
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    after = set(work_dir.glob("*.pptx"))
    created = sorted(after - before, key=lambda p: p.stat().st_mtime, reverse=True)
    expected = work_dir / f"{input_path.stem}.pptx"
    if expected.exists():
        return expected
    if created:
        return created[0]

    detail = (result.stderr or result.stdout or "LibreOffice conversion failed.").strip()
    raise PowerPointError(detail)


def inspect_pptx(pptx_path: Path) -> dict:
    try:
        with zipfile.ZipFile(pptx_path, "r") as zf:
            slide_paths = _presentation_slide_paths(zf)
            slides = []
            for index, slide_path in enumerate(slide_paths, start=1):
                notes_path = _notes_path_for_slide(zf, slide_path)
                notes = ""
                can_write = False
                if notes_path and notes_path in zf.namelist():
                    notes, can_write = _extract_notes_text(zf.read(notes_path))
                slides.append(
                    SlideNote(
                        number=index,
                        slide_path=slide_path,
                        notes_path=notes_path,
                        notes=notes,
                        has_notes_slide=bool(notes_path and notes_path in zf.namelist()),
                        can_write=can_write,
                    )
                )
    except zipfile.BadZipFile as exc:
        raise PowerPointError("This file is not a readable .pptx package.") from exc
    except ET.ParseError as exc:
        raise PowerPointError("The PowerPoint XML could not be parsed.") from exc

    note_slide_count = sum(1 for slide in slides if slide.has_notes_slide)
    note_count = sum(1 for slide in slides if slide.has_text)
    writable_note_count = sum(1 for slide in slides if slide.can_write and slide.has_text)

    return {
        "slide_count": len(slides),
        "note_slide_count": note_slide_count,
        "note_count": note_count,
        "writable_note_count": writable_note_count,
        "slides": [slide.to_public_dict() for slide in slides],
    }


def write_summarized_notes(
    pptx_path: Path,
    output_path: Path,
    slide_dicts: Iterable[dict],
    summaries_by_slide: dict[int, str],
) -> dict:
    slide_notes = [SlideNote(**_slide_dict_for_dataclass(slide)) for slide in slide_dicts]
    notes_paths = {
        slide.notes_path: summaries_by_slide[slide.number]
        for slide in slide_notes
        if slide.notes_path and slide.number in summaries_by_slide
    }
    updated_paths: set[str] = set()
    warnings: list[str] = []

    with zipfile.ZipFile(pptx_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in notes_paths:
                    replacement, changed = _replace_notes_text(data, notes_paths[item.filename])
                    if changed:
                        data = replacement
                        updated_paths.add(item.filename)
                    else:
                        warnings.append(f"Could not rewrite notes XML at {item.filename}.")
                zout.writestr(item, data)

    missing_paths = set(notes_paths) - updated_paths
    for path in sorted(missing_paths):
        warnings.append(f"No writable notes body was found at {path}.")

    return {"updated_count": len(updated_paths), "warnings": warnings}


def dump_session(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def load_session(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _slide_dict_for_dataclass(slide: dict) -> dict:
    return {
        "number": slide["number"],
        "slide_path": slide["slide_path"],
        "notes_path": slide.get("notes_path"),
        "notes": slide.get("notes", ""),
        "has_notes_slide": slide.get("has_notes_slide", False),
        "can_write": slide.get("can_write", False),
    }


def _find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    mac_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac_path.exists():
        return str(mac_path)
    return None


def _presentation_slide_paths(zf: zipfile.ZipFile) -> list[str]:
    names = set(zf.namelist())
    if "ppt/presentation.xml" in names and "ppt/_rels/presentation.xml.rels" in names:
        rels = _relationships(zf, "ppt/_rels/presentation.xml.rels")
        rel_by_id = {rel["Id"]: rel for rel in rels}
        root = ET.fromstring(zf.read("ppt/presentation.xml"))
        slide_paths: list[str] = []
        for slide_id in root.findall(".//p:sldId", NS):
            rel_id = slide_id.attrib.get(f"{{{R_NS}}}id")
            rel = rel_by_id.get(rel_id or "")
            if not rel or rel.get("TargetMode") == "External":
                continue
            path = _normalize_target("ppt/presentation.xml", rel["Target"])
            if path in names:
                slide_paths.append(path)
        if slide_paths:
            return slide_paths

    return sorted(
        [name for name in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)],
        key=_natural_key,
    )


def _notes_path_for_slide(zf: zipfile.ZipFile, slide_path: str) -> str | None:
    rels_path = _rels_path_for_part(slide_path)
    for rel in _relationships(zf, rels_path):
        if rel.get("TargetMode") == "External":
            continue
        if rel.get("Type", "").endswith("/notesSlide"):
            return _normalize_target(slide_path, rel["Target"])
    return None


def _relationships(zf: zipfile.ZipFile, rels_path: str) -> list[dict[str, str]]:
    if rels_path not in zf.namelist():
        return []
    root = ET.fromstring(zf.read(rels_path))
    relationships = []
    for rel in root:
        if _local_name(rel.tag) != "Relationship":
            continue
        relationships.append(dict(rel.attrib))
    return relationships


def _rels_path_for_part(part_path: str) -> str:
    directory = posixpath.dirname(part_path)
    filename = posixpath.basename(part_path)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def _normalize_target(source_part: str, target: str) -> str:
    target = unquote(target)
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    source_dir = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(source_dir, target))


def _extract_notes_text(xml_bytes: bytes) -> tuple[str, bool]:
    root = ET.fromstring(xml_bytes)
    tx_body = _find_notes_tx_body(root)
    if tx_body is None:
        return "", False

    paragraphs: list[str] = []
    for paragraph in tx_body.findall("a:p", NS):
        fragments = [node.text or "" for node in paragraph.findall(".//a:t", NS)]
        text = "".join(fragments).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs).strip(), True


def _replace_notes_text(xml_bytes: bytes, notes: str) -> tuple[bytes, bool]:
    root = ET.fromstring(xml_bytes)
    tx_body = _find_notes_tx_body(root)
    if tx_body is None:
        return xml_bytes, False

    for child in list(tx_body):
        if child.tag == f"{{{A_NS}}}p":
            tx_body.remove(child)

    lines = [line.rstrip() for line in notes.strip().splitlines() if line.strip()]
    if not lines:
        lines = [""]

    for line in lines:
        tx_body.append(_paragraph(line))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), True


def _find_notes_tx_body(root: ET.Element) -> ET.Element | None:
    body_candidates: list[ET.Element] = []
    fallback_candidates: list[ET.Element] = []
    for shape in root.findall(".//p:sp", NS):
        tx_body = shape.find("p:txBody", NS)
        if tx_body is None:
            continue
        placeholder = shape.find("./p:nvSpPr/p:nvPr/p:ph", NS)
        if placeholder is not None:
            placeholder_type = placeholder.attrib.get("type")
            placeholder_idx = placeholder.attrib.get("idx")
            if placeholder_type == "body" or placeholder_idx == "1":
                body_candidates.append(tx_body)
            else:
                fallback_candidates.append(tx_body)
        else:
            fallback_candidates.append(tx_body)

    if body_candidates:
        return body_candidates[0]

    for candidate in fallback_candidates:
        text = " ".join(node.text or "" for node in candidate.findall(".//a:t", NS)).strip()
        if text:
            return candidate
    return fallback_candidates[0] if fallback_candidates else None


def _paragraph(text: str) -> ET.Element:
    paragraph = ET.Element(f"{{{A_NS}}}p")
    run = ET.SubElement(paragraph, f"{{{A_NS}}}r")
    ET.SubElement(run, f"{{{A_NS}}}rPr", {"lang": "en-US", "sz": "1200"})
    text_node = ET.SubElement(run, f"{{{A_NS}}}t")
    text_node.text = text
    if text != text.strip() or "  " in text:
        text_node.set(f"{{{XML_NS}}}space", "preserve")
    return paragraph


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _natural_key(value: str) -> list[int | str]:
    return [int(part) if part.isdigit() else part for part in re.split(r"(\d+)", value)]
