"""Local Office/PDF extraction and Copilot inventory prompt generation.

DOCX and PPTX are parsed using the Python standard library. PDF uses pypdf,
or PyMuPDF if pypdf is unavailable. This module never makes network requests.
"""

import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
MAX_FILE_BYTES = 120 * 1024 * 1024
MAX_XML_PART_BYTES = 40 * 1024 * 1024


def _tag(prefix, name):
    return "{%s}%s" % (NS[prefix], name)


def _clean(s):
    return re.sub(r"\r\n?", "\n", s or "").strip()


def _read_xml(zf, name):
    if name not in zf.namelist():
        return None
    if zf.getinfo(name).file_size > MAX_XML_PART_BYTES:
        raise ValueError("XML part exceeds the size limit: " + name)
    return ET.fromstring(zf.read(name))


def _check_file(path, suffixes):
    path = Path(path)
    if not path.is_file() or path.suffix.lower() not in suffixes:
        raise ValueError("Select a valid " + "/".join(sorted(suffixes)) + " file: " + str(path))
    if path.stat().st_size > MAX_FILE_BYTES:
        raise ValueError("The file exceeds the 120 MB local parsing limit: " + path.name)
    return path


def _word_text(element):
    pieces = []
    for item in element.iter():
        if item.tag == _tag("w", "t"):
            pieces.append(item.text or "")
        elif item.tag == _tag("w", "tab"):
            pieces.append("\t")
        elif item.tag in (_tag("w", "br"), _tag("w", "cr")):
            pieces.append("\n")
    return _clean("".join(pieces))


def _word_style(p):
    e = p.find("w:pPr/w:pStyle", NS)
    return e.get(_tag("w", "val"), "") if e is not None else ""


def _word_table_rows(tbl):
    rows = []
    # Rows can be wrapped by a content control (w:sdt).
    for tr in tbl.findall(".//w:tr", NS):
        cells = []
        for tc in tr.findall("w:tc", NS):
            paragraphs = [_word_text(p) for p in tc.findall(".//w:p", NS)]
            cells.append("\n".join(p for p in paragraphs if p))
        if any(cells):
            rows.append(cells)
    return rows


def _word_blocks(container, part, blocks, headings, counter):
    for child in container:
        if child.tag == _tag("w", "p"):
            value = _word_text(child)
            if not value:
                continue
            style = _word_style(child)
            match = re.match(r"heading\s*([1-9])$", style, re.I)
            if match and part == "body":
                level = int(match.group(1))
                headings[level] = value
                for n in list(headings):
                    if n > level:
                        del headings[n]
            counter[0] += 1
            block = {
                "source_ref": "BRD-B%05d" % counter[0],
                "part": part,
                "kind": "paragraph",
                "text": value,
                "heading_path": [headings[n] for n in sorted(headings)],
            }
            if style:
                block["style"] = style
            blocks.append(block)
        elif child.tag == _tag("w", "tbl"):
            counter[0] += 1
            table_ref = "BRD-T%05d" % counter[0]
            for row_no, cells in enumerate(_word_table_rows(child), 1):
                blocks.append({
                    "source_ref": table_ref + "-R%04d" % row_no,
                    "part": part,
                    "kind": "table_row",
                    "row_number": row_no,
                    "cells": cells,
                    "heading_path": [headings[n] for n in sorted(headings)],
                })
        elif child.tag not in (_tag("w", "sectPr"),):
            # Includes w:sdt (content controls) and custom XML wrappers.
            _word_blocks(child, part, blocks, headings, counter)


def extract_docx(path):
    path = _check_file(path, {".docx"})
    blocks, warnings, counter = [], [], [0]
    try:
        with zipfile.ZipFile(path) as zf:
            document = _read_xml(zf, "word/document.xml")
            if document is None:
                raise ValueError("DOCX is missing word/document.xml")
            body = document.find("w:body", NS)
            if body is None:
                raise ValueError("DOCX body is missing")
            _word_blocks(body, "body", blocks, {}, counter)
            optional = [n for n in zf.namelist() if re.match(
                r"word/(header\d+|footer\d+|footnotes|endnotes)\.xml$", n)]
            for name in sorted(optional):
                xml = _read_xml(zf, name)
                if xml is not None:
                    _word_blocks(xml, Path(name).stem, blocks, {}, counter)
            image_count = sum(n.startswith("word/media/") and not n.endswith("/") for n in zf.namelist())
            if image_count:
                warnings.append("%d embedded image(s) were not OCR processed; review diagrams and screenshots manually." % image_count)
    except zipfile.BadZipFile as exc:
        raise ValueError("The DOCX file is invalid or damaged") from exc
    if not blocks:
        raise ValueError("No text found in the BRD DOCX; check whether it contains only images.")
    return _payload(path, "BRD", "docx", blocks, warnings)


def extract_pdf(path):
    path = _check_file(path, {".pdf"})
    blocks, warnings = [], []
    pages = []
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        if reader.is_encrypted:
            raise ValueError("Encrypted PDF: provide an unlocked copy for local extraction.")
        for page in reader.pages:
            pages.append(page.extract_text() or "")
    except ImportError:
        try:
            import fitz
        except ImportError as exc:
            raise RuntimeError("PDF parsing needs pypdf or PyMuPDF in this WinPython environment. See README.txt.") from exc
        try:
            with fitz.open(str(path)) as document:
                if document.needs_pass:
                    raise ValueError("Encrypted PDF: provide an unlocked copy for local extraction.")
                pages = [page.get_text("text") for page in document]
        except ValueError:
            raise
    for page_no, page_text in enumerate(pages, 1):
        # Page references remain stable; blank lines separate source blocks.
        paragraphs = [s.strip() for s in re.split(r"\n\s*\n", page_text) if s.strip()]
        for block_no, para in enumerate(paragraphs, 1):
            blocks.append({
                "source_ref": "BRD-P%04d-B%04d" % (page_no, block_no),
                "page_number": page_no,
                "kind": "page_text",
                "text": para,
            })
    if any(not p.strip() for p in pages):
        warnings.append("Some PDF pages have no selectable text. Check scans, screenshots, and diagrams; OCR is not included.")
    if not blocks:
        raise ValueError("No selectable text found in the BRD PDF; a scanned PDF requires approved OCR first.")
    return _payload(path, "BRD", "pdf", blocks, warnings, page_count=len(pages))


def _drawing_text(element):
    # Explicit paragraphs preserve boundaries and reading order within a shape.
    values = []
    for paragraph in element.findall(".//a:p", NS):
        pieces = []
        for item in paragraph.iter():
            if item.tag == _tag("a", "t"):
                pieces.append(item.text or "")
            elif item.tag == _tag("a", "br"):
                pieces.append("\n")
        value = _clean("".join(pieces))
        if value:
            values.append(value)
    return "\n".join(values)


def _rels(zf, part_path):
    directory, basename = posixpath.split(part_path)
    rel_path = posixpath.join(directory, "_rels", basename + ".rels")
    root = _read_xml(zf, rel_path)
    if root is None:
        return {}
    result = {}
    for rel in root.findall("rel:Relationship", NS):
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        result[rel.get("Id")] = posixpath.normpath(
            target.lstrip("/") if target.startswith("/") else posixpath.join(directory, target))
    return result


def _slide_paths(zf):
    presentation = _read_xml(zf, "ppt/presentation.xml")
    if presentation is None:
        raise ValueError("PPTX is missing ppt/presentation.xml")
    rels = _rels(zf, "ppt/presentation.xml")
    ids = presentation.findall(".//p:sldId", NS)
    slides = [rels[e.get(_tag("r", "id"))] for e in ids if e.get(_tag("r", "id")) in rels]
    if not slides:
        raise ValueError("Could not resolve the presentation slide order")
    return slides


def _slide_shape_blocks(zf, node, slide_no, slide_part, rels, blocks, warnings, is_notes=False):
    count = 0
    for shape in node:
        local = shape.tag.rsplit("}", 1)[-1]
        if local == "grpSp":
            count += _slide_shape_blocks(zf, shape, slide_no, slide_part, rels, blocks, warnings, is_notes)
            continue
        if local not in ("sp", "cxnSp", "graphicFrame", "pic"):
            continue
        count += 1
        prefix = "SOL-S%04d-%s%03d" % (slide_no, "N" if is_notes else "E", len(blocks) + 1)
        if local == "pic":
            warnings.append("Slide %d contains an image; image text and diagrams are not OCR processed." % slide_no)
            continue
        if is_notes:
            placeholder = shape.find(".//p:ph", NS)
            if placeholder is not None and placeholder.get("type") in ("sldNum", "dt", "ftr", "hdr", "sldImg"):
                continue
        table = shape.find(".//a:tbl", NS) if local == "graphicFrame" else None
        if table is not None:
            for row_no, tr in enumerate(table.findall("a:tr", NS), 1):
                cells = [_drawing_text(tc) for tc in tr.findall("a:tc", NS)]
                if any(cells):
                    blocks.append({
                        "source_ref": prefix + "-R%03d" % row_no,
                        "slide_number": slide_no,
                        "kind": "table_row",
                        "row_number": row_no,
                        "cells": cells,
                    })
            continue
        value = _drawing_text(shape)
        if value:
            placeholder = shape.find(".//p:ph", NS)
            ph_type = placeholder.get("type", "") if placeholder is not None else ""
            blocks.append({
                "source_ref": prefix,
                "slide_number": slide_no,
                "kind": "speaker_notes" if is_notes else "shape_text",
                "text": value,
                **({"placeholder_type": ph_type} if ph_type else {}),
            })
        if local == "graphicFrame":
            # Native chart labels/data and SmartArt diagram data live in related parts.
            relation_ids = {
                value for item in shape.iter() for key, value in item.attrib.items()
                if key.startswith("{" + NS["r"] + "}") and value in rels
            }
            for rid in sorted(relation_ids):
                target = rels.get(rid)
                if not target or not target.endswith(".xml") or target not in zf.namelist():
                    continue
                if not (target.startswith("ppt/charts/") or target.startswith("ppt/diagrams/")):
                    continue
                related = _read_xml(zf, target)
                if related is None:
                    continue
                if target.startswith("ppt/charts/"):
                    values = [x.text or "" for x in related.iter() if x.tag in (_tag("c", "v"), _tag("a", "t"))]
                else:
                    values = [x.text or "" for x in related.iter() if x.tag == _tag("a", "t")]
                if values:
                    blocks.append({
                        "source_ref": prefix + "-DATA",
                        "slide_number": slide_no,
                        "kind": "chart_or_diagram_data",
                        "text": " | ".join(values),
                    })
    return count


def extract_pptx(path):
    path = _check_file(path, {".pptx"})
    blocks, warnings = [], []
    try:
        with zipfile.ZipFile(path) as zf:
            slide_paths = _slide_paths(zf)
            for slide_no, slide_part in enumerate(slide_paths, 1):
                slide = _read_xml(zf, slide_part)
                if slide is None:
                    warnings.append("Slide %d could not be read." % slide_no)
                    continue
                start = len(blocks)
                rels = _rels(zf, slide_part)
                tree = slide.find(".//p:spTree", NS)
                if tree is not None:
                    _slide_shape_blocks(zf, tree, slide_no, slide_part, rels, blocks, warnings)
                title = next((b["text"].split("\n", 1)[0] for b in blocks[start:]
                              if b.get("placeholder_type") in ("title", "ctrTitle")), "")
                if not title:
                    title = next((b["text"].split("\n", 1)[0] for b in blocks[start:] if b.get("kind") == "shape_text"), "")
                for block in blocks[start:]:
                    block["slide_title"] = title
                notes_parts = [v for v in rels.values() if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", v)]
                for notes_part in notes_parts:
                    notes = _read_xml(zf, notes_part)
                    if notes is not None:
                        note_tree = notes.find(".//p:spTree", NS)
                        if note_tree is not None:
                            _slide_shape_blocks(zf, note_tree, slide_no, notes_part,
                                                _rels(zf, notes_part), blocks, warnings, is_notes=True)
                            for block in blocks[start:]:
                                block.setdefault("slide_title", title)
                if len(blocks) == start:
                    warnings.append("Slide %d has no extractable text; review it visually." % slide_no)
    except zipfile.BadZipFile as exc:
        raise ValueError("The PPTX file is invalid or damaged") from exc
    if not blocks:
        raise ValueError("No text found in the PPTX; image-only slides require approved OCR or manual review.")
    # Image warning appears once per image. Compact the message for readability.
    warnings = list(dict.fromkeys(warnings))
    return _payload(path, "Solution", "pptx", blocks, warnings, slide_count=len(slide_paths))


def _payload(path, source_type, fmt, blocks, warnings, **counts):
    return {
        "schema_version": 1,
        "source_type": source_type,
        "source_file": path.name,
        "format": fmt,
        "block_count": len(blocks),
        **counts,
        "warnings": warnings,
        "blocks": blocks,
    }


def extract_brd(path):
    ext = Path(path).suffix.lower()
    if ext == ".docx":
        return extract_docx(path)
    if ext == ".pdf":
        return extract_pdf(path)
    raise ValueError("BRD must be a .docx or .pdf file. Convert legacy .doc files first.")


