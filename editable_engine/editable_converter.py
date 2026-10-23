"""Offline editable PDF -> DOCX conversion helper.

The C++ UI invokes this executable.  Text remains Word text, detected tables
become Word tables, raster images remain separate pictures, and every PDF page
gets its own Word section/page size.
"""

from __future__ import annotations

import argparse
from collections import Counter
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import traceback
import unicodedata
from pathlib import Path

import pymupdf
from pdf2docx import Converter
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table
from docx.text.run import Run


BASE_FONT = "Microsoft YaHei"
MARGIN_PT = 32.0
KNOWN_PUA_DIGITS = {
    # Adobe Illustrator + NotoSansSC subset used by cn-cb-tariff-zh.pdf.
    # The embedded glyphs draw 0..9, while the broken ToUnicode map exposes
    # them as U+F6B1..U+F6BA.
    **{0xF6B1 + digit: str(digit) for digit in range(10)},
}


def repair_pdf_glyph_text(text: str) -> tuple[str, int]:
    """Recover common PDF subset-font characters mapped into Unicode PUA.

    A frequent Chinese PDF producer stores printable ASCII glyphs at U+F020
    through U+F07E. PDF viewers can draw those embedded glyphs, but Word does
    not have the subset font and displays squares. Convert that mapping back to
    real Unicode before the DOCX is saved.
    """
    repaired = []
    count = 0
    for char in text or "":
        code = ord(char)
        if code in KNOWN_PUA_DIGITS:
            repaired.append(KNOWN_PUA_DIGITS[code])
            count += 1
        elif 0xF020 <= code <= 0xF07E:
            repaired.append(chr(code - 0xF000))
            count += 1
        else:
            normalized = unicodedata.normalize("NFKC", char)
            repaired.append(normalized)
            if normalized != char:
                count += 1
    return "".join(repaired), count


def suspicious_glyph_count(text: str) -> int:
    return sum(
        char in {"\ufffd", "\u25a1"} or 0xE000 <= ord(char) <= 0xF8FF
        for char in text or ""
    )


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    tr_pr.append(repeat)


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_geometry(table, widths_pt: list[float]) -> None:
    """Write deterministic fixed table/grid/cell widths."""
    widths_dxa = [max(240, int(width * 20)) for width in widths_pt]
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_pr = table._tbl.tblPr

    tbl_w = tbl_pr.first_child_found_in("w:tblW")
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths_dxa)))
    tbl_w.set(qn("w:type"), "dxa")

    layout = tbl_pr.first_child_found_in("w:tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    indent = tbl_pr.first_child_found_in("w:tblInd")
    if indent is None:
        indent = OxmlElement("w:tblInd")
        tbl_pr.append(indent)
    indent.set(qn("w:w"), "120")
    indent.set(qn("w:type"), "dxa")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths_dxa:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)

    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            width = widths_dxa[min(idx, len(widths_dxa) - 1)]
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.first_child_found_in("w:tcW")
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            set_cell_margins(cell)


def configure_styles(document: Document) -> None:
    # Source-derived formatting is the main authority.  These are the explicit
    # compact-reference fallbacks used when a PDF span has no usable metadata.
    normal = document.styles["Normal"]
    normal.font.name = BASE_FONT
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), BASE_FONT)
    normal.font.size = Pt(11)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.25


def configure_section(section, width_pt: float, height_pt: float) -> str:
    landscape = width_pt > height_pt
    section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    section.page_width = Pt(width_pt)
    section.page_height = Pt(height_pt)
    section.top_margin = Pt(MARGIN_PT)
    section.bottom_margin = Pt(MARGIN_PT)
    section.left_margin = Pt(MARGIN_PT)
    section.right_margin = Pt(MARGIN_PT)
    section.header_distance = Pt(14)
    section.footer_distance = Pt(14)
    return "landscape" if landscape else "portrait"


def intersection_ratio(a, b) -> float:
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    x0, y0 = max(ax0, bx0), max(ay0, by0)
    x1, y1 = min(ax1, bx1), min(ay1, by1)
    if x1 <= x0 or y1 <= y0:
        return 0.0
    area = max(1.0, (ax1 - ax0) * (ay1 - ay0))
    return ((x1 - x0) * (y1 - y0)) / area


def font_color(color_value: int) -> RGBColor:
    color_value = int(color_value or 0)
    return RGBColor((color_value >> 16) & 255, (color_value >> 8) & 255, color_value & 255)


def add_text_block(document: Document, block: dict, page_width: float) -> bool:
    lines = block.get("lines") or []
    if not lines:
        return False
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(3)
    paragraph.paragraph_format.line_spacing = 1.05
    x0, _, x1, _ = block["bbox"]
    center = (x0 + x1) / 2
    if abs(center - page_width / 2) < page_width * 0.08 and (x1 - x0) < page_width * 0.88:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif x0 > page_width * 0.55:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    wrote = False
    for line_index, line in enumerate(lines):
        for span in line.get("spans") or []:
            text = span.get("text", "")
            if not text:
                continue
            run = paragraph.add_run(text)
            font_name = span.get("font") or BASE_FONT
            run.font.name = font_name
            if run._element.rPr is not None:
                run._element.rPr.rFonts.set(qn("w:eastAsia"), BASE_FONT)
            run.font.size = Pt(min(48.0, max(6.5, float(span.get("size") or 11))))
            lower_name = font_name.lower()
            flags = int(span.get("flags") or 0)
            run.bold = "bold" in lower_name or bool(flags & 16)
            run.italic = "italic" in lower_name or "oblique" in lower_name or bool(flags & 2)
            run.font.color.rgb = font_color(span.get("color", 0))
            wrote = True
        if line_index + 1 < len(lines):
            paragraph.add_run().add_break()
    if not wrote:
        paragraph._element.getparent().remove(paragraph._element)
    return wrote


def table_grid_x_edges(table_obj) -> list[float]:
    values = [float(table_obj.bbox[0]), float(table_obj.bbox[2])]
    for row in table_obj.rows:
        for cell in row.cells:
            if cell is not None:
                values.extend((float(cell[0]), float(cell[2])))
    values.sort()
    edges = []
    for value in values:
        if not edges or abs(value - edges[-1]) > 0.5:
            edges.append(value)
    expected = int(table_obj.col_count) + 1
    if len(edges) == expected:
        return edges
    # The detector may omit all internal edges for a fully merged table. In
    # that case retain its declared column count with equal logical columns.
    left, right = float(table_obj.bbox[0]), float(table_obj.bbox[2])
    return [left + (right - left) * index / max(1, expected - 1) for index in range(expected)]


def text_in_rect(raw: dict, rect) -> str:
    """Extract characters by geometry without trusting broken table merges."""
    x0, y0, x1, y1 = rect
    found = []
    for block in raw.get("blocks") or []:
        if int(block.get("type", -1)) != 0:
            continue
        for line in block.get("lines") or []:
            chars = []
            for span in line.get("spans") or []:
                for char in span.get("chars") or []:
                    box = char.get("bbox") or (0, 0, 0, 0)
                    center_x = (float(box[0]) + float(box[2])) / 2
                    center_y = (float(box[1]) + float(box[3])) / 2
                    if x0 <= center_x < x1 and y0 <= center_y < y1:
                        chars.append((float(box[0]), char.get("c", "")))
            if chars:
                chars.sort(key=lambda item: item[0])
                line_box = line.get("bbox") or (0, 0, 0, 0)
                found.append((float(line_box[1]), float(line_box[0]), "".join(char for _, char in chars)))
    found.sort(key=lambda item: (round(item[0], 1), item[1]))
    return "\n".join(text for _, _, text in found).strip()


def extract_table_matrix(page, table_obj) -> list[list[str]]:
    x_edges = table_grid_x_edges(table_obj)
    raw = page.get_text("rawdict", sort=False)
    matrix = []
    for row in table_obj.rows:
        boxes = [cell for cell in row.cells if cell is not None]
        if not boxes:
            matrix.append([""] * int(table_obj.col_count))
            continue
        y0 = min(float(cell[1]) for cell in boxes)
        y1 = max(float(cell[3]) for cell in boxes)
        matrix.append([
            text_in_rect(raw, (x_edges[index], y0, x_edges[index + 1], y1))
            for index in range(int(table_obj.col_count))
        ])
    return matrix


def drawing_fill_for_rect(drawings: list[dict], rect) -> tuple[int, int, int] | None:
    x0, y0, x1, y1 = rect
    area = max(1.0, (x1 - x0) * (y1 - y0))
    best = None
    best_ratio = 0.0
    for drawing in drawings:
        fill = drawing.get("fill")
        box = drawing.get("rect")
        if not fill or box is None:
            continue
        ix0, iy0 = max(x0, float(box.x0)), max(y0, float(box.y0))
        ix1, iy1 = min(x1, float(box.x1)), min(y1, float(box.y1))
        if ix1 <= ix0 or iy1 <= iy0:
            continue
        ratio = ((ix1 - ix0) * (iy1 - iy0)) / area
        if ratio > best_ratio:
            best_ratio = ratio
            best = tuple(max(0, min(255, round(float(value) * 255))) for value in fill[:3])
    return best if best_ratio >= 0.45 else None


def table_column_widths(table_obj, usable_width: float) -> list[float]:
    count = max(1, int(table_obj.col_count))
    try:
        edges = table_grid_x_edges(table_obj)
        raw = [max(1.0, edges[index + 1] - edges[index]) for index in range(count)]
        total = sum(raw)
        return [usable_width * value / total for value in raw]
    except Exception:
        return [usable_width / count] * count


def add_table(document: Document, page, table_obj, usable_width: float) -> bool:
    matrix = extract_table_matrix(page, table_obj)
    if not matrix or not any(any((cell or "").strip() for cell in row) for row in matrix):
        return False
    cols = max(int(table_obj.col_count), max(len(row) for row in matrix))
    word_table = document.add_table(rows=len(matrix), cols=cols)
    word_table.style = "Table Grid"
    widths = table_column_widths(table_obj, usable_width)
    if len(widths) < cols:
        widths.extend([usable_width / cols] * (cols - len(widths)))
    set_table_geometry(word_table, widths[:cols])
    x_edges = table_grid_x_edges(table_obj)
    drawings = page.get_drawings()

    for row_index, row in enumerate(matrix):
        source_boxes = [cell for cell in table_obj.rows[row_index].cells if cell is not None]
        source_y0 = min(float(cell[1]) for cell in source_boxes) if source_boxes else float(table_obj.bbox[1])
        source_y1 = max(float(cell[3]) for cell in source_boxes) if source_boxes else float(table_obj.bbox[3])
        for col_index in range(cols):
            cell = word_table.cell(row_index, col_index)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            value = "" if col_index >= len(row) or row[col_index] is None else str(row[col_index]).strip()
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            run = paragraph.add_run(value)
            run.font.name = BASE_FONT
            run._element.rPr.rFonts.set(qn("w:eastAsia"), BASE_FONT)
            run.font.size = Pt(9.5)
            source_rect = (x_edges[col_index], source_y0, x_edges[col_index + 1], source_y1)
            fill = drawing_fill_for_rect(drawings, source_rect)
            if fill:
                shade_cell(cell, "{:02X}{:02X}{:02X}".format(*fill))
                luminance = 0.2126 * fill[0] + 0.7152 * fill[1] + 0.0722 * fill[2]
                if luminance < 190:
                    run.font.color.rgb = RGBColor(255, 255, 255)
            if row_index == 0:
                run.bold = True
                if not fill:
                    shade_cell(cell, "E8EEF5")
    set_repeat_table_header(word_table.rows[0])
    document.add_paragraph().paragraph_format.space_after = Pt(2)
    return True


def add_image_block(document: Document, block: dict, usable_width: float, usable_height: float) -> bool:
    image = block.get("image")
    if not image:
        return False
    width_px = max(1, int(block.get("width") or 1))
    height_px = max(1, int(block.get("height") or 1))
    width = min(usable_width, max(36.0, float(block["bbox"][2] - block["bbox"][0])))
    height = width * height_px / width_px
    if height > usable_height * 0.72:
        height = usable_height * 0.72
        width = height * width_px / height_px
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(4)
    try:
        paragraph.add_run().add_picture(io.BytesIO(image), width=Pt(width), height=Pt(height))
        return True
    except Exception:
        paragraph._element.getparent().remove(paragraph._element)
        return False


def add_page_fallback(document: Document, page, usable_width: float, usable_height: float) -> None:
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5), alpha=False)
    data = pixmap.tobytes("png")
    ratio = pixmap.height / max(1, pixmap.width)
    width = usable_width
    height = width * ratio
    if height > usable_height:
        height = usable_height
        width = height / ratio
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.add_run().add_picture(io.BytesIO(data), width=Pt(width), height=Pt(height))


EMU_PER_PT = 12700
# A scan re-encoded as PNG is several times larger than the JPEG it came from.
# Only sizeable images are worth re-compressing; icons and line art are left
# alone, where PNG is both smaller and sharper.
IMAGE_RECOMPRESS_MIN_BYTES = 300 * 1024
IMAGE_RECOMPRESS_QUALITY = 82
IMAGE_RECOMPRESS_MIN_GAIN = 0.25


def section_usable_box(section) -> tuple[float, float]:
    """Width and height in points that content can actually occupy."""
    width = float(section.page_width.pt) - float(section.left_margin.pt) - float(section.right_margin.pt)
    height = float(section.page_height.pt) - float(section.top_margin.pt) - float(section.bottom_margin.pt)
    return max(36.0, width), max(36.0, height)


def body_blocks_by_section(document):
    """Yield ``(block_element, index, section)`` for every top level body block.

    A section break lives in the last paragraph of the section it closes, so
    walking the body in order is the only way to tell which page geometry a
    given picture will be laid out on.
    """
    sections = document.sections
    if not sections:
        return
    index = 0
    for child in document.element.body:
        if child.tag == qn("w:sectPr"):
            continue
        position = min(index, len(sections) - 1)
        yield child, position, sections[position]
        if child.tag == qn("w:p"):
            properties = child.find(qn("w:pPr"))
            if properties is not None and properties.find(qn("w:sectPr")) is not None:
                index += 1


def fit_images_to_page(document) -> int:
    """Shrink pictures that do not fit between the margins.

    The layout engine sizes a full page scan to the whole paper size, margins
    included. Word then pushes the picture past the bottom margin and every
    source page turns into two Word pages. Scaling the picture down to the
    usable box keeps one page per page.
    """
    resized = 0
    for block, _, section in body_blocks_by_section(document):
        max_width, max_height = section_usable_box(section)
        for drawing in block.iter(qn("w:drawing")):
            for extent in drawing.iter(qn("wp:extent")):
                try:
                    cx, cy = int(extent.get("cx")), int(extent.get("cy"))
                except (TypeError, ValueError):
                    continue
                width, height = cx / EMU_PER_PT, cy / EMU_PER_PT
                if width <= 0 or height <= 0:
                    continue
                scale = min(1.0, max_width / width, max_height / height)
                if scale > 0.999:
                    continue
                new_cx, new_cy = int(cx * scale), int(cy * scale)
                extent.set("cx", str(new_cx))
                extent.set("cy", str(new_cy))
                for ext in drawing.iter(qn("a:ext")):
                    ext.set("cx", str(new_cx))
                    ext.set("cy", str(new_cy))
                resized += 1
    return resized


def jpeg_version(data: bytes) -> bytes | None:
    """Re-encode one image as JPEG, or return None when that is a bad idea."""
    try:
        pixmap = pymupdf.Pixmap(data)
    except Exception:
        return None
    try:
        # Transparency cannot survive JPEG, and CMYK/greyscale separations are
        # left to the original encoder.
        if pixmap.alpha or pixmap.n not in (1, 3):
            return None
        return pixmap.tobytes("jpeg", jpg_quality=IMAGE_RECOMPRESS_QUALITY)
    except Exception:
        return None


def shrink_docx_images(output_path: Path) -> dict:
    """Store large photographic images as JPEG instead of PNG.

    pdf2docx re-encodes every embedded picture as PNG, which inflates a scanned
    document several times over -- a 1 MB scan came back as a 5.4 MB DOCX. The
    pixels are unchanged apart from JPEG quantisation; layout and text are not
    touched at all.
    """
    report = {"recompressed_images": 0, "bytes_saved": 0}
    with zipfile.ZipFile(output_path) as archive:
        names = archive.namelist()
        candidates = {}
        for name in names:
            if not name.startswith("word/media/") or not name.lower().endswith(".png"):
                continue
            if archive.getinfo(name).file_size < IMAGE_RECOMPRESS_MIN_BYTES:
                continue
            original = archive.read(name)
            converted = jpeg_version(original)
            if converted and len(converted) < len(original) * (1.0 - IMAGE_RECOMPRESS_MIN_GAIN):
                candidates[name] = (converted, len(original))
        if not candidates:
            return report
        renames = {name: name[: -len(".png")] + ".jpeg" for name in candidates}
        payload = {}
        for name in names:
            if name in candidates:
                continue
            payload[name] = archive.read(name)

    def patch_xml(data: bytes) -> bytes:
        text = data.decode("utf-8")
        for name, renamed in renames.items():
            text = text.replace(Path(name).name, Path(renamed).name)
        return text.encode("utf-8")

    content_types = payload.get("[Content_Types].xml", b"").decode("utf-8")
    if 'Extension="jpeg"' not in content_types:
        content_types = content_types.replace(
            "</Types>",
            '<Default Extension="jpeg" ContentType="image/jpeg"/></Types>',
            1,
        )
    payload["[Content_Types].xml"] = content_types.encode("utf-8")
    for name in list(payload):
        if name.endswith(".rels"):
            payload[name] = patch_xml(payload[name])

    temporary = output_path.with_suffix(".jpeg-pass.docx")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in payload.items():
            archive.writestr(name, data)
        for name, (data, original_size) in candidates.items():
            # Images are already compressed; storing them again wastes time.
            archive.writestr(renames[name], data, zipfile.ZIP_STORED)
            report["recompressed_images"] += 1
            report["bytes_saved"] += original_size - len(data)

    try:
        probe = Document(temporary)
        if len(probe.inline_shapes) != len(Document(output_path).inline_shapes):
            raise ValueError("inline shape count changed")
    except Exception:
        # Never hand back a file that cannot be reopened.
        temporary.unlink(missing_ok=True)
        return {"recompressed_images": 0, "bytes_saved": 0}

    temporary.replace(output_path)
    return report


def set_run_font(run, font_name: str = BASE_FONT) -> None:
    run.font.name = font_name
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    for name in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.set(qn(f"w:{name}"), font_name)


def iter_document_parts(document):
    """Yield every XML tree of the document that can hold visible text.

    ``document.paragraphs`` and ``document.tables`` only expose top level body
    content.  The layout engine regularly puts text into nested tables, text
    boxes and page headers/footers, so every scan in this module walks the XML
    instead.  Missing those places used to make perfectly converted numbers
    look lost and aborted the whole conversion.
    """
    yield document.element.body
    visited = []
    for section in document.sections:
        for name in (
            "header",
            "footer",
            "first_page_header",
            "first_page_footer",
            "even_page_header",
            "even_page_footer",
        ):
            part = getattr(section, name, None)
            if part is None:
                continue
            try:
                if part.is_linked_to_previous:
                    continue
            except Exception:
                pass
            element = part._element
            if any(element is known for known in visited):
                continue
            visited.append(element)
            yield element


def closest_ancestor(element, tag: str):
    parent = element.getparent()
    while parent is not None:
        if parent.tag == tag:
            return parent
        parent = parent.getparent()
    return None


def iter_document_runs(document):
    """Yield ``(run, inside_table)`` for every run of the document."""
    for part in iter_document_parts(document):
        for element in part.iter(qn("w:r")):
            yield Run(element, None), closest_ancestor(element, qn("w:tbl")) is not None


def iter_document_tables(document):
    """Yield every table, including tables nested inside other cells."""
    for part in iter_document_parts(document):
        for element in part.iter(qn("w:tbl")):
            yield Table(element, None)


# Separators must survive the scan: two cells reading "1" and "2" around a tab
# must not be read back as the number "12".
TEXT_NODES = {
    qn("w:t"): None,
    qn("w:tab"): "\t",
    qn("w:br"): "\n",
    qn("w:cr"): "\n",
    qn("w:noBreakHyphen"): "-",
}


def document_text(document) -> str:
    """Return the full visible text, nested and floating content included."""
    parts = []
    for part in iter_document_parts(document):
        for paragraph in part.iter(qn("w:p")):
            pieces = []
            for node in paragraph.iter(*TEXT_NODES):
                # A text box nests complete paragraphs inside another
                # paragraph. Only take the text owned by this paragraph so a
                # value is never counted twice.
                if closest_ancestor(node, qn("w:p")) is not paragraph:
                    continue
                replacement = TEXT_NODES[node.tag]
                pieces.append(node.text or "" if replacement is None else replacement)
            if pieces:
                parts.append("".join(pieces))
    return "\n".join(parts)


def compact_wide_tables(document: Document) -> None:
    """Keep dense bank/price tables close to the PDF's compact geometry."""
    for table in iter_document_tables(document):
        column_count = max((len(row.cells) for row in table.rows), default=0)
        if column_count < 5:
            continue
        target_size = 6.0 if column_count >= 7 else 7.0
        table.autofit = False
        for row in table.rows:
            tr_pr = row._tr.get_or_add_trPr()
            for tag in ("w:trHeight", "w:cantSplit"):
                for node in list(tr_pr.findall(qn(tag))):
                    tr_pr.remove(node)
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                set_cell_margins(cell, top=20, start=35, bottom=20, end=35)
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.keep_together = False
                    paragraph.paragraph_format.keep_with_next = False
                    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    paragraph.paragraph_format.line_spacing = 1.0
                    for run in paragraph.runs:
                        set_run_font(run)
                        if run.font.size is None or run.font.size.pt > target_size:
                            run.font.size = Pt(target_size)


def repair_and_normalize_docx(output_path: Path) -> tuple[dict, str]:
    """Repair broken glyphs everywhere and return the report plus final text."""
    document = Document(output_path)
    repaired_glyphs = 0
    for run, inside_table in iter_document_runs(document):
        repaired, count = repair_pdf_glyph_text(run.text)
        if count:
            run.text = repaired
            repaired_glyphs += count
            set_run_font(run)
        elif inside_table:
            # PDF subset fonts are not installed on the target PC. A real
            # system font prevents otherwise valid digits and Chinese glyphs
            # from becoming square placeholders.
            set_run_font(run)
    compact_wide_tables(document)
    resized_images = fit_images_to_page(document)
    document.save(output_path)
    # The text is taken from the in-memory document; re-opening a 180 page
    # file only to read it back again doubled the conversion time.
    text, stats = docx_text_and_stats(document)
    report = dict(stats)
    report.update({
        "repaired_glyphs": repaired_glyphs,
        "suspicious_glyphs": suspicious_glyph_count(text),
        "resized_images": resized_images,
    })
    return report, text


def convert_legacy(input_path: Path, output_path: Path) -> dict:
    pdf = pymupdf.open(input_path)
    if pdf.page_count == 0:
        raise ValueError("PDF has no pages")

    document = Document()
    configure_styles(document)
    # Content is added page by page. Some templates expose an empty seed
    # paragraph while others expose none, so handle both forms.
    if document.paragraphs and not document.paragraphs[0].text:
        seed = document.paragraphs[0]._element
        seed.getparent().remove(seed)

    report = {
        "pages": pdf.page_count,
        "text_blocks": 0,
        "tables": 0,
        "images": 0,
        "fallback_pages": [],
        "orientations": [],
    }

    for page_index, page in enumerate(pdf):
        section = document.sections[0] if page_index == 0 else document.add_section(WD_SECTION.NEW_PAGE)
        page_width, page_height = float(page.rect.width), float(page.rect.height)
        report["orientations"].append(configure_section(section, page_width, page_height))
        usable_width = max(72.0, page_width - MARGIN_PT * 2)
        usable_height = max(72.0, page_height - MARGIN_PT * 2)

        try:
            detected_tables = list(page.find_tables().tables)
        except Exception:
            detected_tables = []
        table_boxes = [tuple(table.bbox) for table in detected_tables]
        page_dict = page.get_text("dict", sort=True)
        events = []
        for table in detected_tables:
            events.append((float(table.bbox[1]), 0, "table", table))
        for block in page_dict.get("blocks") or []:
            bbox = tuple(block.get("bbox") or (0, 0, 0, 0))
            if any(intersection_ratio(bbox, table_bbox) > 0.25 for table_bbox in table_boxes):
                continue
            block_type = int(block.get("type", -1))
            if block_type == 0:
                events.append((float(bbox[1]), 1, "text", block))
            elif block_type == 1:
                events.append((float(bbox[1]), 2, "image", block))
        events.sort(key=lambda event: (event[0], event[1]))

        page_content = 0
        for _, _, kind, value in events:
            if kind == "table" and add_table(document, page, value, usable_width):
                report["tables"] += 1
                page_content += 1
            elif kind == "text" and add_text_block(document, value, page_width):
                report["text_blocks"] += 1
                page_content += 1
            elif kind == "image" and add_image_block(document, value, usable_width, usable_height):
                report["images"] += 1
                page_content += 1

        # Image-only/scanned pages cannot be made editable without OCR. Preserve
        # their appearance instead of silently emitting a blank page.
        if page_content == 0:
            add_page_fallback(document, page, usable_width, usable_height)
            report["fallback_pages"].append(page_index + 1)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    pdf.close()
    return report


NUMBER_PATTERN = re.compile(
    r"(?<![0-9A-Za-z_])(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|"
    r"[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?)(?![0-9A-Za-z_])"
)


# Converting a long manual can drop a handful of numbers that live in drawn
# figures, watermarks or artefacts of the page layout analysis. Refusing such a
# document outright left the user with nothing at all, which is worse than an
# accurate document plus a warning. Only a real loss of numeric content (a
# broken subset font, a page range that produced no text) still fails.
NUMERIC_TOLERANCE_RATIO = 0.01
NUMERIC_TOLERANCE_FLOOR = 5
ABSENT_TOLERANCE_RATIO = 0.02
ABSENT_TOLERANCE_FLOOR = 3
GLYPH_TOLERANCE_RATIO = 0.002
GLYPH_TOLERANCE_FLOOR = 4


def numeric_tolerance(total_tokens: int) -> int:
    return max(NUMERIC_TOLERANCE_FLOOR, int(total_tokens * NUMERIC_TOLERANCE_RATIO))


def absent_tolerance(distinct_values: int) -> int:
    return max(ABSENT_TOLERANCE_FLOOR, int(distinct_values * ABSENT_TOLERANCE_RATIO))


def glyph_tolerance(text_length: int) -> int:
    return max(GLYPH_TOLERANCE_FLOOR, int(text_length * GLYPH_TOLERANCE_RATIO))


def numeric_tokens(text: str) -> list[str]:
    return [match.group(0) for match in NUMBER_PATTERN.finditer(text or "")]


def normalize_number(token: str) -> str:
    """Compare values instead of spelling.

    ``10,000`` in the PDF and ``10000`` in the DOCX are the same number, and
    full width digits have to match their ASCII form.
    """
    text = unicodedata.normalize("NFKC", token or "")
    return text.replace(",", "").replace(" ", "").lstrip("+")


# Windows recognises printed Chinese well, but reports every decimal point as
# a full width comma, every hyphen as the character 一, and separates CJK
# characters with spaces. Those substitutions are systematic, so they can be
# undone with confidence -- measured on a real quotation, all 26 prices and all
# 26 part numbers came back correct.
OCR_RENDER_DPI = 300
OCR_LANGUAGE = "zh-Hans-CN"
OCR_TEXT_SIZE_PT = 1.0


def repair_ocr_text(text: str) -> str:
    """Undo the systematic punctuation damage of the Windows OCR engine."""
    text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", text)
    text = re.sub(r"(?<=[0-9A-Za-z])\s+(?=[，．。·、,.\u4e00-\u9fff])", "", text)
    text = re.sub(r"(?<=[，．。·、,.\u4e00-\u9fff])\s+(?=[0-9A-Za-z])", "", text)
    # 240 ， 00 -> 240.00
    text = re.sub(r"(?<=\d)[，。．·、](?=\d{1,2}(?!\d))", ".", text)
    # CZCT7001 一 01 -> CZCT7001-01
    text = re.sub(r"(?<=[0-9A-Za-z])一(?=[0-9A-Za-z])", "-", text)
    return text.strip()


def ocr_page_data(tool: str, page, work_dir: Path, index, clip=None) -> tuple[dict, str]:
    """Recognise one page, returning its lines in reading order.

    Windows OCR supplies the page layout; the bundled offline model verifies
    ruled-table cells and recognises rectified photographs.
    """
    image_path = work_dir / f"ocr-page-{index}.png"
    json_path = work_dir / f"ocr-page-{index}.json"
    page.get_pixmap(dpi=OCR_RENDER_DPI, clip=clip).save(image_path)
    # Remove long grid strokes before recognition: narrow cells otherwise lose
    # digits such as 11, and WinRT can read a border as part of the number.
    from scan_tables import detect_tables, extract_stamps, stamp_free_image
    import cv2
    import numpy as np
    image = cv2.imdecode(np.fromfile(str(image_path), dtype=np.uint8), cv2.IMREAD_COLOR)
    tables = detect_tables(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
    photo = {}
    if not tables:
        from photo_tables import rectify_photo_table
        flattened = rectify_photo_table(cv2.cvtColor(image,cv2.COLOR_BGR2RGB),photo)
        if flattened is not None:
            image = cv2.cvtColor(flattened,cv2.COLOR_RGB2BGR)
            cv2.imencode('.png',image)[1].tofile(str(image_path))
            tables = detect_tables(flattened)
    recognised_image = image_path
    if tables:
        rgb = cv2.cvtColor(image,cv2.COLOR_BGR2RGB)
        stamps = extract_stamps(rgb)
        clean = cv2.cvtColor(stamp_free_image(rgb,stamps),cv2.COLOR_RGB2BGR)
        mask = cv2.bitwise_or(tables[0]["horizontal"], tables[0]["vertical"])
        # Only remove ruled strokes inside confirmed table bounds.
        bounded = np.zeros_like(mask)
        for table in tables:
            x0, y0, x1, y1 = table["bbox"]
            bounded[max(0,y0-3):y1+4, max(0,x0-3):x1+4] = mask[max(0,y0-3):y1+4, max(0,x0-3):x1+4]
        clean[cv2.dilate(bounded, np.ones((3, 3), np.uint8)) > 0] = 255
        recognised_image = work_dir / f"ocr-clean-{index}.png"
        cv2.imencode(".png", clean)[1].tofile(str(recognised_image))
    try:
        subprocess.run(
            [tool, "--ocr", str(recognised_image), str(json_path), "--ocr-language", OCR_LANGUAGE],
            check=False,
            capture_output=True,
            timeout=300,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        return {}, f"调用 OCR 失败：{exc}"
    if not json_path.exists():
        return {}, "OCR 没有返回结果"
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        return {}, f"OCR 结果无法解析：{exc}"
    if not payload.get("ok"):
        return {}, str(payload.get("error") or "OCR 不可用")
    if photo.get('rows') and tables:
        payload['photo_rectified'] = True
        payload['photo_rows'] = photo['rows']
        payload['photo_columns'] = photo['columns']
        payload['photo_side_artwork'] = photo.get('side_artwork',[])
    # Re-read empty numeric cells independently. Windows OCR occasionally drops
    # a short value (notably 11/16) in an otherwise perfectly recognised row.
    from scan_tables import cell_groups, word_box, words_text
    if tables:
        all_words = [w for line in payload.get("lines", []) for w in line.get("words", [])]
        retry_count = 0
        for table in tables:
            xs, ys = table["xs"], table["ys"]
            groups = set(cell_groups(table))
            for col in range(len(xs)-1):
                header = [w for w in all_words if xs[col] <= sum(word_box(w)[::2])/2 <= xs[col+1]
                          and ys[0] <= (float(w['y'])+float(w['h'])/2) <= ys[1]]
                label = words_text(header, repair_ocr_text)
                if not any(name in label for name in ("序号", "数量", "单价", "金额", "合计")):
                    continue
                for row in range(1, len(ys)-1):
                    if (row, col, row, col) not in groups or retry_count >= 20:
                        continue
                    present = [w for w in all_words if xs[col] <= float(w['x'])+float(w['w'])/2 <= xs[col+1]
                               and ys[row] <= float(w['y'])+float(w['h'])/2 <= ys[row+1]]
                    if present:
                        continue
                    retry_count += 1
                    x0, y0, x1, y1 = xs[col]+7, ys[row]+7, xs[col+1]-7, ys[row+1]-7
                    crop = clean[y0:y1, x0:x1]
                    crop = cv2.resize(crop, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                    crop = cv2.copyMakeBorder(crop, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=(255,255,255))
                    crop_path = work_dir / f"cell-{index}-{row}-{col}.png"
                    result_path = crop_path.with_suffix('.json')
                    cv2.imencode('.png', crop)[1].tofile(str(crop_path))
                    try:
                        subprocess.run([tool, '--ocr', str(crop_path), str(result_path), '--ocr-language', 'en-US'],
                                       capture_output=True, timeout=30, check=False)
                        cell_payload = json.loads(result_path.read_text(encoding='utf-8'))
                        for line in cell_payload.get('lines', []):
                            if not re.fullmatch(r'[\d\s.,，．]+', str(line.get('text', ''))):
                                continue
                            for word in line.get('words', []):
                                word['x'] = x0 + (float(word['x'])-20)/2
                                word['y'] = y0 + (float(word['y'])-20)/2
                                word['w'] = float(word['w'])/2
                                word['h'] = float(word['h'])/2
                            payload['lines'].append(line)
                    except (OSError, ValueError, subprocess.SubprocessError):
                        pass
        # A seal can cover the strokes of 月/日. Re-read only malformed date
        # lines at a larger scale, accepting a replacement only when the OCR
        # explicitly returns a valid date with the same recognised year.
        from datetime import date
        for stamp_index,stamp in enumerate(stamps):
            x0,y0,x1,y1 = stamp['bbox']
            cy = (y0+y1)/2
            if any(t['ys'][0] <= cy <= t['ys'][-1] for t in tables):
                continue
            suspicious = [line for line in payload['lines']
                          if re.search(r'\d{4}\s*年',str(line.get('text','')))
                          and any(y0 <= float(w['y']) <= y1 for w in line.get('words',[]))
                          and not re.fullmatch(r'\d{4}年\d{1,2}月\d{1,2}日',
                                               words_text(line.get('words',[]),repair_ocr_text))]
            if not suspicious:
                continue
            padding = max(40,round((x1-x0)*.35))
            left,top = max(0,x0-padding),max(0,y0-40)
            right,bottom = min(clean.shape[1],x1+padding),min(clean.shape[0],y1+40)
            crop = cv2.resize(clean[top:bottom,left:right],None,fx=2,fy=2,interpolation=cv2.INTER_CUBIC)
            crop_path = work_dir/f'stamp-date-{index}-{stamp_index}.png'
            result_path = crop_path.with_suffix('.json')
            cv2.imencode('.png',crop)[1].tofile(str(crop_path))
            try:
                subprocess.run([tool,'--ocr',str(crop_path),str(result_path),'--ocr-language',OCR_LANGUAGE],
                               capture_output=True,timeout=30,check=False)
                retry = json.loads(result_path.read_text(encoding='utf-8'))
                for line in retry.get('lines',[]):
                    value = words_text(line.get('words',[]),repair_ocr_text)
                    match = re.fullmatch(r'(\d{4})年(\d{1,2})月(\d{1,2})日',value)
                    if not match:
                        continue
                    date(*(int(part) for part in match.groups()))
                    original = next((old for old in suspicious if match[1] in str(old.get('text',''))),None)
                    if original is None:
                        continue
                    for word in line.get('words',[]):
                        word['x'] = left+float(word['x'])/2
                        word['y'] = top+float(word['y'])/2
                        word['w'] = float(word['w'])/2
                        word['h'] = float(word['h'])/2
                    payload['lines'].remove(original)
                    suspicious.remove(original)
                    payload['lines'].append(line)
            except (OSError,ValueError,subprocess.SubprocessError):
                pass
    if payload.get('photo_rectified'):
        from photo_ocr import recognize_grid
        photo_result=recognize_grid(clean,tables[0],payload,repair_ocr_text)
        for key in ('photo_rectified','photo_rows','photo_columns','photo_side_artwork'):
            photo_result[key]=payload[key]
        payload=photo_result
    elif tables:
        from cell_ocr import verify_cells
        try:
            payload = verify_cells(clean,tables,payload,repair_ocr_text,stamps=stamps)
        except Exception as exc:
            # Preserve the first reading, but never silently report a verified
            # result when the independent recognizer was unavailable.
            payload['cell_ocr_error'] = f'单元格复核失败：{type(exc).__name__}: {exc}'
    return payload, ""


def collect_cell_ocr_report(report: dict, payload: dict, page: int) -> None:
    report['cell_ocr_checked'] = report.get('cell_ocr_checked',0)+payload.get('cell_ocr_checked',0)
    for key in ('cell_ocr_changes','cell_ocr_issues'):
        report.setdefault(key,[]).extend(dict(item,page=page) for item in payload.get(key,[]))
    if payload.get('cell_ocr_error'):
        report.setdefault('cell_ocr_errors',[]).append(payload['cell_ocr_error'])


def ocr_lines(payload: dict) -> list[str]:
    # Sort by the first word of each line: the engine returns lines in its own
    # order, which jumbles columns together.
    entries = []
    for line in payload.get("lines") or []:
        words = line.get("words") or []
        if not words:
            continue
        top = min(float(word.get("y", 0.0)) for word in words)
        left = min(float(word.get("x", 0.0)) for word in words)
        text = repair_ocr_text(str(line.get("text") or ""))
        if text:
            entries.append((round(top / 12.0), left, text))
    entries.sort()
    return [text for _, _, text in entries]


def ocr_page_lines(tool: str, page, work_dir: Path, index: int) -> tuple[list[str], str]:
    payload, error = ocr_page_data(tool, page, work_dir, index)
    return ocr_lines(payload), error


def float_picture_over_page(paragraph_element, width_pt: float, height_pt: float) -> bool:
    """Turn an inline page picture into one pinned over the whole sheet.

    The scan then covers the paper exactly as it did in the PDF, and the
    recognised text sits underneath it: the page looks untouched, but the words
    can be searched, selected and copied, and deleting the picture reveals them.
    """
    for drawing in paragraph_element.iter(qn("w:drawing")):
        inline = drawing.find(qn("wp:inline"))
        if inline is None:
            continue
        graphic = inline.find(qn("a:graphic"))
        if graphic is None:
            continue
        cx, cy = int(width_pt * EMU_PER_PT), int(height_pt * EMU_PER_PT)
        anchor = parse_xml(
            f"<wp:anchor {nsdecls('wp', 'a', 'r')} distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\""
            " simplePos=\"0\" relativeHeight=\"2\" behindDoc=\"0\" locked=\"0\""
            " layoutInCell=\"1\" allowOverlap=\"1\">"
            "<wp:simplePos x=\"0\" y=\"0\"/>"
            "<wp:positionH relativeFrom=\"page\"><wp:posOffset>0</wp:posOffset></wp:positionH>"
            "<wp:positionV relativeFrom=\"page\"><wp:posOffset>0</wp:posOffset></wp:positionV>"
            f"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>"
            "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>"
            "<wp:wrapNone/>"
            "<wp:docPr id=\"900\" name=\"scanned page\"/>"
            "<wp:cNvGraphicFramePr/>"
            "</wp:anchor>"
        )
        for extent in graphic.iter(qn("a:ext")):
            extent.set("cx", str(cx))
            extent.set("cy", str(cy))
        anchor.append(graphic)
        drawing.remove(inline)
        drawing.append(anchor)
        return True
    return False


def add_ocr_paragraph(body, before, text: str) -> None:
    paragraph = parse_xml(
        f"<w:p {nsdecls('w')}><w:pPr>"
        f"<w:spacing w:before=\"0\" w:after=\"0\" w:line=\"{int(OCR_TEXT_SIZE_PT * 20)}\" w:lineRule=\"exact\"/>"
        f"<w:rPr><w:sz w:val=\"{int(OCR_TEXT_SIZE_PT * 2)}\"/></w:rPr>"
        "</w:pPr><w:r><w:rPr>"
        f"<w:rFonts w:ascii=\"{BASE_FONT}\" w:hAnsi=\"{BASE_FONT}\" w:eastAsia=\"{BASE_FONT}\"/>"
        f"<w:sz w:val=\"{int(OCR_TEXT_SIZE_PT * 2)}\"/>"
        "</w:rPr><w:t xml:space=\"preserve\"></w:t></w:r></w:p>"
    )
    paragraph.iter(qn("w:t")).__next__().text = text
    before.addprevious(paragraph)


def add_ocr_text_layer(output_path: Path, input_path: Path, pages: list[int], tool: str,
                       work_dir: Path) -> dict:
    """Reconstruct ruled scan tables, retaining searchable fallback for other scans."""
    from scan_tables import rebuild_page
    import cv2
    import numpy as np

    report = {"ocr_pages": [], "ocr_lines": 0, "ocr_error": "",
              "ocr_table_pages": [], "ocr_tables": 0, "ocr_cell_words": 0}
    document = Document(output_path)
    blocks = list(body_blocks_by_section(document))
    sections = document.sections
    source = pymupdf.open(input_path)
    try:
        if len(sections) != source.page_count:
            report["ocr_error"] = "页面与节的数量不一致，跳过文字层"
            return report
        by_section: dict[int, list] = {}
        for block, position, _ in blocks:
            by_section.setdefault(position, []).append(block)

        for page_number in pages:
            index = page_number - 1
            if index not in by_section:
                continue
            payload, error = ocr_page_data(tool, source[index], work_dir, page_number)
            lines = ocr_lines(payload)
            if error:
                report["ocr_error"] = error
                break
            if not lines:
                continue
            image_path = work_dir / f"ocr-page-{page_number}.png"
            image = cv2.cvtColor(cv2.imdecode(np.fromfile(str(image_path), dtype=np.uint8),
                                             cv2.IMREAD_COLOR), cv2.COLOR_BGR2RGB)
            fragment = Document()
            configure_styles(fragment)
            if payload.get('photo_rectified'):
                # Phone-photo pixel dimensions are not physical paper points.
                # Use readable A4 geometry for the reconstructed table.
                configure_section(sections[index],595.,842.)
            width, height = section_usable_box(sections[index])
            tables, cell_words = rebuild_page(fragment, image, payload, width, height,
                                              repair_ocr_text, set_table_geometry, set_run_font)
            if tables:
                target = by_section[index][0]
                # Image bands belong to the fragment's package. Rebind their
                # relationships before moving the XML into the output package.
                for blip in fragment.element.body.iter(qn("a:blip")):
                    rid = blip.get(qn("r:embed"))
                    if rid:
                        part = fragment.part.related_parts[rid]
                        new_rid, _ = document.part.get_or_add_image(io.BytesIO(part.blob))
                        blip.set(qn("r:embed"), new_rid)
                for child in list(fragment.element.body):
                    if child.tag != qn("w:sectPr"):
                        target.addprevious(child)
                for block in by_section[index]:
                    properties = block.find(qn("w:pPr"))
                    if properties is not None and properties.find(qn("w:sectPr")) is not None:
                        # The closing section break belongs to this source page.
                        for child in list(block):
                            if child is not properties:
                                block.remove(child)
                        # pdf2docx may leave an explicit page break on the holder.
                    else:
                        document.element.body.remove(block)
                report["ocr_table_pages"].append(page_number)
                report["ocr_tables"] += tables
                report["ocr_cell_words"] += cell_words
                report["ocr_pages"].append(page_number)
                report["ocr_lines"] += len(lines)
                collect_cell_ocr_report(report,payload,page_number)
                if payload.get('photo_rectified'):
                    report.setdefault('photo_rectified_pages',[]).append(page_number)
                    report['photo_ocr_engine']=payload.get('photo_ocr_engine','')
                    report.setdefault('photo_unresolved_cells',[]).extend(
                        [page_number,row+1,col+1] for row,col in payload.get('photo_unresolved_cells',[]))
                continue
            rect = source[index].rect
            holder = None
            for block in by_section[index]:
                if block.find(qn("w:r")) is not None or block.iter(qn("w:drawing")):
                    if float_picture_over_page(block, float(rect.width), float(rect.height)):
                        holder = block
                        break
            target = holder if holder is not None else by_section[index][0]
            for line in lines:
                add_ocr_paragraph(document.element.body, target, line)
            report["ocr_pages"].append(page_number)
            report["ocr_lines"] += len(lines)
        if report["ocr_pages"]:
            used = {element.get(qn("r:embed")) for element in document.element.body.iter(qn("a:blip"))}
            for rid, relation in list(document.part.rels.items()):
                if relation.reltype.endswith('/image') and rid not in used:
                    document.part.drop_rel(rid)
            normalize_drawing_ids(document)
            document.save(output_path)
    finally:
        source.close()
    return report


def add_embedded_image_tables(output_path: Path, input_path: Path, excluded_pages: list[int],
                              tool: str, work_dir: Path) -> dict:
    """Replace inline raster tables on otherwise text-bearing PDF pages."""
    from scan_tables import detect_tables, rebuild_page
    import numpy as np

    report = {"embedded_table_pages": [], "embedded_tables": 0}
    document = Document(output_path)
    source = pymupdf.open(input_path)
    try:
        if len(document.sections) != len(source):
            return report
        blocks = list(body_blocks_by_section(document))
        for index, page in enumerate(source):
            if index + 1 in excluded_pages:
                continue
            candidates = []
            for block, position, section in blocks:
                if position != index or block.tag != qn('w:p'):
                    continue
                # Avoid removing prose that shares a paragraph with a picture.
                if any((node.text or '').strip() for node in block.iter(qn('w:t'))):
                    continue
                for drawing in block.iter(qn('wp:inline')):
                    extent = drawing.find(qn('wp:extent'))
                    if extent is not None:
                        candidates.append((block, drawing, extent, section))
            used = set()
            seen_rects = set()
            for image_index, info in enumerate(page.get_image_info()):
                rect = pymupdf.Rect(info['bbox']) & page.rect
                key = tuple(round(v, 1) for v in rect)
                if key in seen_rects or rect.width < 90 or rect.height < 35:
                    continue
                seen_rects.add(key)
                matches = []
                for candidate in candidates:
                    block, drawing, extent, section = candidate
                    if id(block) in used:
                        continue
                    w = int(extent.get('cx')) / EMU_PER_PT
                    h = int(extent.get('cy')) / EMU_PER_PT
                    difference = abs((w / h) / (rect.width / rect.height) - 1)
                    if difference < .08:
                        matches.append((difference, candidate))
                if not matches:
                    continue
                pix = page.get_pixmap(dpi=OCR_RENDER_DPI, clip=rect)
                image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                if not detect_tables(image):
                    continue
                token = f'embedded-{index+1}-{image_index}'
                payload, error = ocr_page_data(tool, page, work_dir, token, clip=rect)
                if error:
                    report['embedded_ocr_error'] = error
                    continue
                _, (block, _, extent, section) = min(matches, key=lambda item: item[0])
                fragment = Document()
                configure_styles(fragment)
                usable_width, usable_height = section_usable_box(section)
                width = min(usable_width, int(extent.get('cx')) / EMU_PER_PT)
                height = min(usable_height, int(extent.get('cy')) / EMU_PER_PT)
                count, _ = rebuild_page(fragment, image, payload, width, height,
                                         repair_ocr_text, set_table_geometry, set_run_font)
                if not count:
                    continue
                collect_cell_ocr_report(report,payload,index+1)
                for blip in fragment.element.body.iter(qn('a:blip')):
                    rid = blip.get(qn('r:embed'))
                    if rid:
                        new_rid, _ = document.part.get_or_add_image(io.BytesIO(fragment.part.related_parts[rid].blob))
                        blip.set(qn('r:embed'), new_rid)
                for child in list(fragment.element.body):
                    if child.tag != qn('w:sectPr'):
                        block.addprevious(child)
                # Preserve the section properties, but remove the original picture.
                for child in list(block):
                    if child.tag != qn('w:pPr'):
                        block.remove(child)
                used.add(id(block))
                report['embedded_tables'] += count
                if index + 1 not in report['embedded_table_pages']:
                    report['embedded_table_pages'].append(index + 1)
        if report['embedded_tables']:
            used_rids = {node.get(qn('r:embed')) for node in document.element.iter(qn('a:blip'))}
            for rid, relation in list(document.part.rels.items()):
                if relation.reltype.endswith('/image') and rid not in used_rids:
                    document.part.drop_rel(rid)
            normalize_drawing_ids(document)
            document.save(output_path)
    finally:
        source.close()
    return report


def normalize_drawing_ids(document) -> None:
    """Fragment-local picture IDs must be unique in the assembled Word body."""
    stamp = 0
    for index,node in enumerate(document.element.iter(qn('wp:docPr')),1):
        node.set('id',str(index))
        if node.get('name','').startswith('PDF stamp '):
            stamp += 1
            node.set('name',f'PDF stamp {stamp}')


def image_only_pages(input_path: Path) -> list[int]:
    """Page numbers that carry no text layer at all.

    Text-based quality gates cannot distinguish scans from empty pages, so
    these pages require a separate OCR / table reconstruction pass.
    """
    source = pymupdf.open(input_path)
    try:
        pages = []
        for index, page in enumerate(source):
            if page.get_text("text").strip():
                continue
            if page.get_images() or page.get_drawings():
                pages.append(index + 1)
        return pages
    finally:
        source.close()


def describe_pages(pages: list[int], limit: int = 20) -> str:
    shown = "、".join(str(page) for page in pages[:limit])
    return shown + ("…" if len(pages) > limit else "")


def source_numeric_tokens(input_path: Path) -> list[str]:
    source = pymupdf.open(input_path)
    try:
        extracted = "\n".join(page.get_text("text", sort=True) for page in source)
        repaired, _ = repair_pdf_glyph_text(extracted)
        return numeric_tokens(repaired)
    finally:
        source.close()


def docx_text_and_stats(source) -> tuple[str, dict]:
    """Accept a path or an already opened document to avoid re-parsing it."""
    document = source if hasattr(source, "element") else Document(source)
    orientations = [
        "landscape" if section.page_width > section.page_height else "portrait"
        for section in document.sections
    ]
    return document_text(document), {
        "paragraphs": len(document.paragraphs),
        "tables": len(document.tables),
        "images": len(document.inline_shapes),
        "sections": len(document.sections),
        "stamp_objects": sum(node.get('name','').startswith('PDF stamp ')
                             for node in document.element.iter(qn('wp:docPr'))),
        "orientations": orientations,
    }


def missing_numeric_tokens(source_tokens: list[str], output_text: str) -> list[str]:
    """Values whose number of occurrences dropped, one entry per lost copy."""
    missing = Counter(normalize_number(token) for token in source_tokens)
    missing -= Counter(normalize_number(token) for token in numeric_tokens(output_text))
    return sorted(missing.elements())


def absent_numeric_values(source_tokens: list[str], output_text: str) -> list[str]:
    """Values that vanished completely.

    Word merges cells and repeats header rows, so an exact occurrence count can
    never be trusted for a pass/fail decision.  A value that is nowhere in the
    document, however, really was lost, and that is what aborts a conversion.
    """
    produced = {normalize_number(token) for token in numeric_tokens(output_text)}
    return sorted({normalize_number(token) for token in source_tokens} - produced)


def enforce_page_sections(input_path: Path, output_path: Path) -> None:
    """Force every output section to use the corresponding PDF page geometry."""
    source = pymupdf.open(input_path)
    document = Document(output_path)
    try:
        if len(document.sections) != source.page_count:
            return
        for index, section in enumerate(document.sections):
            rect = source[index].rect
            configure_section(section, float(rect.width), float(rect.height))
        document.save(output_path)
    finally:
        source.close()


def convert_layout(input_path: Path, output_path: Path) -> tuple[dict, str]:
    converter = Converter(str(input_path))
    try:
        converter.convert(str(output_path))
    finally:
        converter.close()
    enforce_page_sections(input_path, output_path)
    report, text = repair_and_normalize_docx(output_path)
    report.update({"engine": "pdf2docx-layout", "fallback_pages": []})
    return report, text


def evaluate_candidate(text: str, source_numbers: list[str]) -> tuple[list[str], list[str], int]:
    """Score one result: values gone, occurrences lost, unreadable glyphs kept."""
    return (
        absent_numeric_values(source_numbers, text),
        missing_numeric_tokens(source_numbers, text),
        suspicious_glyph_count(text),
    )


def convert(input_path: Path, output_path: Path, ocr_tool: str = "") -> dict:
    """Prefer the layout engine and only fall back when its result is unusable.

    The layout engine keeps the original page design, the structured engine
    rebuilds the page from PDF text objects.  Both are judged by how many
    source numbers went missing and how many unreadable glyphs remain.  A small
    loss is reported as a warning next to the finished document; only a loss
    beyond the tolerance aborts the conversion.
    """
    source_numbers = source_numeric_tokens(input_path)
    scanned_pages = image_only_pages(input_path)
    distinct_numbers = {normalize_number(token) for token in source_numbers}
    allowed_missing = numeric_tolerance(len(source_numbers))
    allowed_absent = absent_tolerance(len(distinct_numbers))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="pdf-to-word-") as temp_dir:
        temp = Path(temp_dir)
        layout_path = temp / "layout.docx"
        legacy_path = temp / "legacy.docx"
        layout_report = None
        layout_text = ""
        layout_error = None
        try:
            layout_report, layout_text = convert_layout(input_path, layout_path)
        except Exception as exc:
            layout_error = str(exc)

        if layout_report is None:
            layout_absent = sorted(distinct_numbers)
            layout_missing, layout_suspicious = list(source_numbers), 0
        else:
            layout_absent, layout_missing, layout_suspicious = evaluate_candidate(
                layout_text, source_numbers
            )

        chosen_path = layout_path
        chosen_report = layout_report
        chosen_text = layout_text
        chosen_absent = layout_absent
        chosen_missing = layout_missing
        chosen_suspicious = layout_suspicious

        layout_acceptable = (
            layout_report is not None
            and len(layout_absent) <= allowed_absent
            and layout_suspicious <= glyph_tolerance(len(layout_text))
        )
        # Running the second engine on an already good document only doubled
        # the waiting time on long files, so it is a real fallback now.
        if not layout_acceptable:
            try:
                legacy_report = convert_legacy(input_path, legacy_path)
                repair_report, legacy_text = repair_and_normalize_docx(legacy_path)
                legacy_report.update(repair_report)
                legacy_absent, legacy_missing, legacy_suspicious = evaluate_candidate(
                    legacy_text, source_numbers
                )
            except Exception as exc:
                # A failing fallback must not throw away a usable first result.
                if layout_report is None:
                    raise ValueError(
                        "版式转换失败：" + (layout_error or "未知错误") + f"；备用引擎同样失败：{exc}"
                    ) from exc
                legacy_report = None
            if legacy_report is not None and (layout_report is None or (
                len(legacy_absent), legacy_suspicious, len(legacy_missing)
            ) < (len(layout_absent), layout_suspicious, len(layout_missing))):
                chosen_path = legacy_path
                chosen_report = legacy_report
                chosen_report["engine"] = "structured-fallback"
                chosen_text = legacy_text
                chosen_absent = legacy_absent
                chosen_missing = legacy_missing
                chosen_suspicious = legacy_suspicious

        if chosen_report is None:
            raise ValueError("版式转换失败：" + (layout_error or "未知错误"))

        allowed_glyphs = glyph_tolerance(len(chosen_text))
        if len(chosen_absent) > allowed_absent:
            preview = "、".join(chosen_absent[:12])
            raise ValueError(
                f"数字完整性校验未通过：源 PDF 共 {len(distinct_numbers)} 个不同数值，"
                f"其中 {len(chosen_absent)} 个在 Word 中完全找不到（允许 {allowed_absent} 个）。"
                f"缺失示例：{preview}。"
                "该 PDF 可能使用了损坏的子集字体，或是没有文字层的扫描件。"
            )
        if chosen_suspicious > allowed_glyphs:
            raise ValueError(
                f"字体编码修复未通过：仍有 {chosen_suspicious} 个私有区或方块字符"
                f"（允许 {allowed_glyphs} 个）。"
            )

        if scanned_pages and ocr_tool:
            chosen_report.update(
                add_ocr_text_layer(chosen_path, input_path, scanned_pages, ocr_tool, temp)
            )
        if ocr_tool:
            embedded_report = add_embedded_image_tables(chosen_path, input_path, scanned_pages, ocr_tool, temp)
            chosen_report['cell_ocr_checked'] = chosen_report.get('cell_ocr_checked',0)+embedded_report.pop('cell_ocr_checked',0)
            for key in ('cell_ocr_changes','cell_ocr_issues','cell_ocr_errors'):
                chosen_report.setdefault(key,[]).extend(embedded_report.pop(key,[]))
            chosen_report.update(embedded_report)
        chosen_report.update(shrink_docx_images(chosen_path))
        _, final_stats = docx_text_and_stats(chosen_path)
        chosen_report.update(final_stats)
        shutil.copy2(chosen_path, output_path)

        warnings = []
        if chosen_report.get('photo_rectified_pages'):
            warnings.append(
                f"第 {describe_pages(chosen_report['photo_rectified_pages'])} 页已校正拍照表格的折皱和阴影，"
                "并用内置离线 OCR 重建表格；请复核原件中模糊或手写遮挡的字符。"
            )
        if chosen_report.get('photo_unresolved_cells'):
            cells=chosen_report['photo_unresolved_cells']
            warnings.append(f"拍照表格有 {len(cells)} 个单元格需人工复核（页、行、列）："+
                            '、'.join('/'.join(map(str,cell)) for cell in cells[:12]))
        if chosen_report.get('embedded_tables'):
            warnings.append(
                f"第 {describe_pages(chosen_report['embedded_table_pages'])} 页的图片已重建为 "
                f"{chosen_report['embedded_tables']} 个可编辑 Word 表格，请复核 OCR 文字和数值。"
            )
        if scanned_pages:
            page_count = chosen_report.get("sections") or len(scanned_pages)
            scope = "整份文件" if len(scanned_pages) >= page_count else f"第 {describe_pages(scanned_pages)} 页"
            recognised = chosen_report.get("ocr_pages") or []
            table_pages = chosen_report.get("ocr_table_pages") or []
            if table_pages:
                warnings.append(
                    f"第 {describe_pages(table_pages)} 页已从扫描图片重建 "
                    f"{chosen_report['ocr_tables']} 个可编辑 Word 表格。"
                    "OCR 识别可能有误，请对照原件复核金额、图号和合并单元格。"
                )
            remaining = [page for page in recognised if page not in table_pages]
            if remaining:
                warnings.append(
                    f"第 {describe_pages(remaining)} 页没有文字层（扫描件），已用 Windows OCR 识别出 "
                    f"{chosen_report.get('ocr_lines', 0)} 行文字并放在扫描图下方："
                    "可以搜索和复制，删除页面上的图片即可看到并编辑这些文字。"
                    "识别可能有误，金额和图号请对照原件核对。"
                )
            unrecognised = [page for page in scanned_pages if page not in recognised]
            if unrecognised:
                detail = chosen_report.get("ocr_error") or "本版本未启用 OCR"
                failed_scope = "整份文件" if len(unrecognised) == page_count else f"第 {describe_pages(unrecognised)} 页"
                warnings.append(
                    f"{failed_scope}没有文字层（扫描件或图片页），只能按页面图片保留："
                    f"共 {len(unrecognised)} 页的文字无法在 Word 中编辑或搜索（{detail}）。"
                )
        for error in chosen_report.get('cell_ocr_errors',[]):
            warnings.append(error)
        cell_issues = chosen_report.get('cell_ocr_issues',[])
        if cell_issues:
            locations = '、'.join(f"第{x['page']}页"+(x['region'] if x.get('region') else f"表{x['table']}第{x['row']}行第{x['column']}列") for x in cell_issues[:12])
            warnings.append(f'有 {len(cell_issues)} 处单元格识别或金额关系需要人工复核：{locations}。保留实际识别值，未用算式补写。')
        if chosen_absent:
            warnings.append(
                f"有 {len(chosen_absent)} 个数值在 Word 中完全找不到："
                + "、".join(chosen_absent[:12])
                + "。请对照 PDF 复核这些数值。"
            )
        if len(chosen_missing) > allowed_missing:
            preview = "、".join(chosen_missing[:12])
            warnings.append(
                f"有 {len(chosen_missing)} 处数字的出现次数少于源文件"
                f"（源文件共 {len(source_numbers)} 处，容差 {allowed_missing} 处）：{preview}。"
                "合并单元格和重复表头会影响计数，通常无需处理。"
            )
        if chosen_suspicious:
            warnings.append(
                f"有 {chosen_suspicious} 个字符仍无法确定编码，请检查对应位置。"
            )
        if chosen_report.get("fallback_pages"):
            pages = chosen_report["fallback_pages"]
            warnings.append(
                f"有 {len(pages)} 页没有文字层，已按页面图片保留（页码："
                + ", ".join(str(page) for page in pages[:20])
                + "）。"
            )

        chosen_report["numbers_checked"] = len(source_numbers)
        chosen_report["distinct_numbers"] = len(distinct_numbers)
        chosen_report["image_only_pages"] = scanned_pages
        chosen_report["number_mismatches"] = len(chosen_missing)
        chosen_report["absent_numbers"] = chosen_absent[:50]
        chosen_report["missing_numbers"] = chosen_missing[:50]
        chosen_report["numeric_tolerance"] = allowed_missing
        chosen_report["absent_tolerance"] = allowed_absent
        chosen_report["warnings"] = warnings
        return chosen_report


def main() -> int:
    parser = argparse.ArgumentParser(description="Offline editable PDF to DOCX converter")
    parser.add_argument("input")
    parser.add_argument("output")
    parser.add_argument("--report")
    parser.add_argument("--error-file")
    parser.add_argument("--warning-file")
    # The C++ shell passes its own path here; it recognises scanned pages with
    # the OCR engine built into Windows.
    parser.add_argument("--ocr-tool", default="")
    args = parser.parse_args()
    try:
        report = convert(Path(args.input), Path(args.output), args.ocr_tool)
        if args.report:
            Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        if args.warning_file:
            # The UI shows this next to the success message; an empty file
            # means the document converted without any reservation.
            Path(args.warning_file).write_text(
                "\n".join(report.get("warnings") or []), encoding="utf-8"
            )
        print(json.dumps(report, ensure_ascii=False))
        return 0
    except Exception as exc:
        if args.error_file:
            Path(args.error_file).write_text(str(exc), encoding="utf-8")
        print(f"Conversion failed: {exc}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
