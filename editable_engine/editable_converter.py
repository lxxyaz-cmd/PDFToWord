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
import sys
import tempfile
import traceback
import unicodedata
from pathlib import Path

import pymupdf
from pdf2docx import Converter
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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
    document.save(output_path)
    # The text is taken from the in-memory document; re-opening a 180 page
    # file only to read it back again doubled the conversion time.
    text, stats = docx_text_and_stats(document)
    report = dict(stats)
    report.update({
        "repaired_glyphs": repaired_glyphs,
        "suspicious_glyphs": suspicious_glyph_count(text),
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


def convert(input_path: Path, output_path: Path) -> dict:
    """Prefer the layout engine and only fall back when its result is unusable.

    The layout engine keeps the original page design, the structured engine
    rebuilds the page from PDF text objects.  Both are judged by how many
    source numbers went missing and how many unreadable glyphs remain.  A small
    loss is reported as a warning next to the finished document; only a loss
    beyond the tolerance aborts the conversion.
    """
    source_numbers = source_numeric_tokens(input_path)
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

        shutil.copy2(chosen_path, output_path)

        warnings = []
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
    args = parser.parse_args()
    try:
        report = convert(Path(args.input), Path(args.output))
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
