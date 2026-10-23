# -*- coding: utf-8 -*-
"""Offline checks for the DOCX scanning and numeric integrity logic.

The conversion engines (``pymupdf``/``pdf2docx``) are only present in the build
environment, so they are stubbed out here: every function under test works on
an already produced Word document and needs ``python-docx`` alone.

    python tests/test_numeric_check.py
"""

from __future__ import annotations

import base64
import contextlib
import sys
import types
from pathlib import Path

# The smallest valid PNG; the picture only has to exist, not to look like
# anything, for the geometry checks.
SMALL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)

def ensure_module(name: str) -> bool:
    """Import the real module, or install a stub and report that it is fake."""
    try:
        __import__(name)
        return True
    except ImportError:
        stub = types.ModuleType(name)
        stub.open = lambda *args, **kwargs: None  # type: ignore[attr-defined]
        stub.Converter = object  # type: ignore[attr-defined]
        sys.modules[name] = stub
        return False


HAS_PYMUPDF = ensure_module("pymupdf")
ensure_module("pdf2docx")

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "editable_engine"))

from docx import Document  # noqa: E402
from docx.oxml import parse_xml  # noqa: E402
from docx.oxml.ns import nsdecls, qn  # noqa: E402
from docx.shared import Pt  # noqa: E402

import editable_converter as engine  # noqa: E402


FAILURES: list[str] = []


def check(condition: bool, label: str) -> None:
    print(("PASS  " if condition else "FAIL  ") + label)
    if not condition:
        FAILURES.append(label)


def build_document():
    """A document that hides numbers where the old scan never looked."""
    document = Document()
    document.add_paragraph("合计 10,000 元")

    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "章节 5.2"
    nested = table.cell(0, 1).add_table(rows=1, cols=1)
    nested.cell(0, 0).text = "深处的 2.1"

    # A text box: a complete paragraph nested inside another paragraph.
    box = parse_xml(
        "<w:p %s><w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\">"
        "<v:textbox><w:txbxContent><w:p><w:r><w:t>图注 84</w:t></w:r></w:p>"
        "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>" % nsdecls("w")
    )
    document.element.body.append(box)

    header = document.sections[0].header
    header.is_linked_to_previous = False
    header.paragraphs[0].text = "第 6 页"
    return document


def test_document_text_finds_hidden_numbers():
    document = build_document()
    text = engine.document_text(document)
    for value in ("10,000", "5.2", "2.1", "84", "6"):
        check(value in text, f"document_text 包含 {value}")

    tokens = engine.numeric_tokens(text)
    check(tokens.count("84") == 1, "文本框内容只统计一次（无重复计数）")
    check(tokens.count("2.1") == 1, "嵌套表格内容只统计一次")


def test_missing_numeric_tokens_normalizes():
    document = build_document()
    text = engine.document_text(document)
    # "10000" in the PDF and "10,000" in the DOCX are the same value.
    missing = engine.missing_numeric_tokens(["10000", "5.2", "2.1", "84", "6"], text)
    check(missing == [], f"千分位与全角写法可以匹配，实际缺失={missing}")

    missing = engine.missing_numeric_tokens(["10000", "999"], text)
    check(missing == ["999"], f"真正缺失的数字仍被报告，实际={missing}")

    full_width = engine.normalize_number("１２３４")
    check(full_width == "1234", f"全角数字归一化，实际={full_width}")


def test_absent_values_versus_occurrence_counts():
    document = Document()
    document.add_paragraph("单价 12 元")
    text = engine.document_text(document)
    # Word merges cells and repeats header rows, so a lower occurrence count is
    # not proof of data loss; a value nobody can find in the file is.
    check(
        engine.missing_numeric_tokens(["12", "12"], text) == ["12"],
        "重复出现次数减少会被记录为提示",
    )
    check(
        engine.absent_numeric_values(["12", "12"], text) == [],
        "只要数值仍然存在就不算丢失",
    )
    check(
        engine.absent_numeric_values(["12", "77"], text) == ["77"],
        "彻底消失的数值仍被判定为丢失",
    )
    check(engine.absent_tolerance(0) == 3, "少量数值使用最小丢失容差")
    check(engine.absent_tolerance(500) == 10, "大量数值按 2% 计算丢失容差")


def test_separators_do_not_glue_numbers():
    document = Document()
    paragraph = document.add_paragraph()
    paragraph.add_run("1")
    paragraph.add_run().add_tab()
    paragraph.add_run("2")
    tokens = engine.numeric_tokens(engine.document_text(document))
    check(tokens == ["1", "2"], f"制表符不会把 1 和 2 拼成 12，实际={tokens}")


def test_tolerance_thresholds():
    check(engine.numeric_tolerance(0) == 5, "空文档使用最小容差")
    check(engine.numeric_tolerance(3412) == 34, "长文档按 1% 计算容差")
    check(engine.numeric_tolerance(44) == 5, "短文档不会因为四舍五入失去容差")
    check(engine.glyph_tolerance(100) == 4, "少量文本使用最小字形容差")
    check(engine.glyph_tolerance(100000) == 200, "长文本按 0.2% 计算字形容差")


def test_repair_walks_nested_content(tmp_path: Path):
    document = Document()
    table = document.add_table(rows=1, cols=1)
    nested = table.cell(0, 0).add_table(rows=1, cols=1)
    # U+F6B1..U+F6BA are the broken subset-font digits 0..9.
    nested.cell(0, 0).paragraphs[0].add_run(chr(0xF6B2) + chr(0xF6B3))
    path = tmp_path / "nested.docx"
    document.save(path)

    report, text = engine.repair_and_normalize_docx(path)
    check(report["repaired_glyphs"] == 2, f"嵌套表格内的私有区数字被修复，实际={report['repaired_glyphs']}")
    check("12" in text, f"修复后的文本为 12，实际文本={text!r}")
    check(report["suspicious_glyphs"] == 0, "修复后不再残留私有区字符")


def test_compact_wide_tables_handles_nested(tmp_path: Path):
    document = Document()
    table = document.add_table(rows=2, cols=6)
    for index, cell in enumerate(table.rows[0].cells):
        cell.text = f"列{index}"
    nested = table.cell(1, 0).add_table(rows=1, cols=7)
    nested.cell(0, 0).text = "嵌套宽表"
    engine.compact_wide_tables(document)
    path = tmp_path / "wide.docx"
    document.save(path)
    check(path.exists(), "宽表压缩支持嵌套表格且不抛异常")


PATCHED = (
    "source_numeric_tokens",
    "image_only_pages",
    "convert_layout",
    "convert_legacy",
    "repair_and_normalize_docx",
)


@contextlib.contextmanager
def stub_engines(source_numbers, layout_body, legacy_body=None, scanned_pages=()):
    """Replace the PDF side of the pipeline with fixed text."""

    def make(body):
        def run(input_path, output_path):
            document = Document()
            document.add_paragraph(body)
            document.save(output_path)
            report = {"engine": "stub", "fallback_pages": []}
            return report, engine.document_text(document)

        return run

    saved = {name: getattr(engine, name) for name in PATCHED}
    engine.source_numeric_tokens = lambda input_path: list(source_numbers)
    engine.image_only_pages = lambda input_path: list(scanned_pages)
    engine.convert_layout = make(layout_body)
    engine.convert_legacy = lambda input_path, output_path: (
        make(legacy_body if legacy_body is not None else layout_body)(input_path, output_path)[0]
    )
    engine.repair_and_normalize_docx = lambda path: (
        {"repaired_glyphs": 0, "suspicious_glyphs": 0},
        engine.document_text(Document(path)),
    )
    try:
        yield
    finally:
        for name, value in saved.items():
            setattr(engine, name, value)


def test_convert_delivers_document_with_warning(tmp_path: Path):
    # Every value is present, only one repeated occurrence is gone.
    with stub_engines(["7"] * 8 + [str(n) for n in range(100, 160)],
                      "7 " + " ".join(str(n) for n in range(100, 160))):
        output = tmp_path / "warned.docx"
        report = engine.convert(Path("source.pdf"), output)
        check(output.exists(), "少量数字对不上时仍然输出 Word 文件")
        check(report["absent_numbers"] == [], "所有数值都还在，不报缺失")
        check(report["number_mismatches"] == 7, f"出现次数差异被记录，实际={report['number_mismatches']}")
        check(
            len(report["warnings"]) == 1 and "出现次数" in report["warnings"][0],
            f"超出计数容差时给出提示而不是失败，实际={report['warnings']}",
        )


def test_scanned_pdf_warns_that_text_is_not_editable(tmp_path: Path):
    # A scan has no text at all, so every numeric and glyph gate passes with a
    # perfect score. The user still has to be told the Word file is pictures.
    with stub_engines([], "", scanned_pages=[1, 2]):
        output = tmp_path / "scanned.docx"
        report = engine.convert(Path("scan.pdf"), output)
        check(output.exists(), "扫描件仍然输出 Word 文件")
        check(report["image_only_pages"] == [1, 2], f"记录无文字层的页码，实际={report.get('image_only_pages')}")
        warning = " ".join(report["warnings"])
        check("没有文字层" in warning, f"提示说明没有文字层，实际={report['warnings']}")
        check("OCR" in warning, f"提示说明本版本不含 OCR，实际={report['warnings']}")


def test_convert_still_fails_on_real_data_loss(tmp_path: Path):
    # The subset-font case: the numbers never arrive in the document.
    with stub_engines([str(n) for n in range(100, 160)], "文本没有任何数字"):
        output = tmp_path / "failed.docx"
        try:
            engine.convert(Path("source.pdf"), output)
            check(False, "数值大面积丢失时应当报错")
        except ValueError as exc:
            check("数字完整性校验未通过" in str(exc), f"报错信息说明原因，实际={exc}")
            check(not output.exists(), "校验失败时不写出残缺的 Word 文件")


def page_sized_picture_document(tmp_path: Path):
    """A4 page holding one picture sized to the whole sheet, margins ignored."""
    document = Document()
    section = document.sections[0]
    engine.configure_section(section, 595.0, 841.0)
    image = tmp_path / "page.png"
    image.write_bytes(SMALL_PNG)
    document.add_paragraph().add_run().add_picture(
        str(image), width=Pt(595), height=Pt(841)
    )
    return document


def test_fit_images_to_page(tmp_path: Path):
    document = page_sized_picture_document(tmp_path)
    usable_width, usable_height = engine.section_usable_box(document.sections[0])
    check(
        (round(usable_width), round(usable_height)) == (531, 777),
        f"可用区按页面尺寸减边距计算，实际={usable_width:.0f}x{usable_height:.0f}",
    )

    before = document.inline_shapes[0]
    check(
        before.height.pt > usable_height,
        f"修复前图片高于可用区（{before.height.pt:.0f} > {usable_height:.0f}），会溢出到下一页",
    )

    resized = engine.fit_images_to_page(document)
    shape = document.inline_shapes[0]
    check(resized == 1, f"超出可用区的图片被缩放，实际处理 {resized} 张")
    check(
        shape.width.pt <= usable_width + 0.5 and shape.height.pt <= usable_height + 0.5,
        f"缩放后落在可用区内，实际={shape.width.pt:.0f}x{shape.height.pt:.0f}",
    )
    check(
        abs(shape.width.pt / shape.height.pt - 595 / 841) < 0.01,
        "缩放保持原始宽高比",
    )

    # A picture that already fits must not be touched.
    check(engine.fit_images_to_page(document) == 0, "再次运行不会重复缩放")


def test_shrink_docx_images(tmp_path: Path):
    if not HAS_PYMUPDF:
        print("SKIP  JPEG 重编码（本环境没有 pymupdf）")
        return
    import pymupdf

    # A photo-like page: smooth gradient plus dither, the shape of a scan.
    width = height = 600
    samples = bytearray()
    for y in range(height):
        for x in range(width):
            base = (x * 255 // width + y * 255 // height) // 2
            samples += bytes(((base + x % 7) % 256, (base + y % 5) % 256, base))
    pixmap = pymupdf.Pixmap(pymupdf.csRGB, width, height, bytes(samples), False)
    picture = tmp_path / "scan.png"
    picture.write_bytes(pixmap.tobytes("png"))

    document = Document()
    document.add_paragraph().add_run().add_picture(str(picture), width=Pt(400))
    output = tmp_path / "scan.docx"
    document.save(output)
    before = output.stat().st_size

    saved_min = engine.IMAGE_RECOMPRESS_MIN_BYTES
    try:
        # Exercise the rewrite regardless of how well this fixture compresses.
        engine.IMAGE_RECOMPRESS_MIN_BYTES = 1024
        report = engine.shrink_docx_images(output)
    finally:
        engine.IMAGE_RECOMPRESS_MIN_BYTES = saved_min

    check(report["recompressed_images"] == 1, f"图片被重编码，实际={report}")
    check(report["bytes_saved"] > 0, f"确实变小了，实际省下 {report['bytes_saved']} 字节")
    check(output.stat().st_size < before, "docx 整体体积下降")

    import zipfile

    with zipfile.ZipFile(output) as archive:
        media = [n for n in archive.namelist() if n.startswith("word/media/")]
        content_types = archive.read("[Content_Types].xml").decode("utf-8")
    check(media == ["word/media/image1.jpeg"], f"媒体文件改名为 .jpeg，实际={media}")
    check('Extension="jpeg"' in content_types, "[Content_Types].xml 登记了 jpeg")

    reopened = Document(output)
    check(len(reopened.inline_shapes) == 1, "重写后的 docx 仍可打开且图片还在")


def test_repair_ocr_text():
    # Exactly what the Windows engine returns for a scanned quotation.
    check(engine.repair_ocr_text("240 ， 00") == "240.00", "全角逗号还原为小数点")
    check(engine.repair_ocr_text("130 ． 00") == "130.00", "全角句号还原为小数点")
    check(engine.repair_ocr_text("960 · 00") == "960.00", "间隔号还原为小数点")
    code = engine.repair_ocr_text("CZCT7001 一 01 一 30 一 001")
    check(code == "CZCT7001-01-30-001", f"汉字一还原为连字符，实际={code}")
    name = engine.repair_ocr_text("天 津 鑫 地 机 械")
    check(name == "天津鑫地机械", f"去掉汉字之间的空格，实际={name}")
    # A real "一" between Chinese characters must survive.
    check(engine.repair_ocr_text("第 一 页") == "第一页", "中文里的「一」不受影响")


def fake_ocr_tool(tmp_path: Path) -> str:
    """A stand-in for the C++ shell that writes the JSON it would produce."""
    script = tmp_path / "fake_ocr.py"
    script.write_text(
        "import json, sys\n"
        "out = sys.argv[sys.argv.index('--ocr') + 2]\n"
        "json.dump({'ok': True, 'width': 2480, 'height': 3505, 'lines': ["
        "{'text': '240 \\uff0c 00', 'words': [{'t': '240', 'x': 100.0, 'y': 500.0}]},"
        "{'text': '\\u5929 \\u6d25 \\u946b \\u5730', 'words': [{'t': '\\u5929', 'x': 90.0, 'y': 100.0}]}"
        "]}, open(out, 'w', encoding='utf-8'))\n",
        encoding="utf-8",
    )
    shim = tmp_path / "fake_ocr.cmd"
    shim.write_text(f'@echo off\r\n"{sys.executable}" "{script}" %*\r\n', encoding="ascii")
    return str(shim)


def test_ocr_text_layer(tmp_path: Path):
    if not HAS_PYMUPDF:
        print("SKIP  OCR 文字层（本环境没有 pymupdf）")
        return
    import pymupdf

    # A one page PDF with no text layer, like a scan.
    source = tmp_path / "scan.pdf"
    pdf = pymupdf.open()
    page = pdf.new_page(width=595, height=841)
    page.insert_image(page.rect, pixmap=pymupdf.Pixmap(pymupdf.csRGB, 40, 56, bytes(40 * 56 * 3), False))
    pdf.save(source)
    pdf.close()

    # The document the layout engine would have produced for it.
    picture = Document()
    engine.configure_section(picture.sections[0], 595.0, 841.0)
    image = tmp_path / "page.png"
    image.write_bytes(SMALL_PNG)
    picture.add_paragraph().add_run().add_picture(str(image), width=Pt(595), height=Pt(841))
    output = tmp_path / "scan.docx"
    picture.save(output)

    report = engine.add_ocr_text_layer(output, source, [1], fake_ocr_tool(tmp_path), tmp_path)
    check(report["ocr_error"] == "", f"OCR 调用无错误，实际={report['ocr_error']}")
    check(report["ocr_pages"] == [1], f"记录识别过的页码，实际={report['ocr_pages']}")
    check(report["ocr_lines"] == 2, f"记录识别行数，实际={report['ocr_lines']}")

    rebuilt = Document(output)
    text = engine.document_text(rebuilt)
    body = rebuilt.element.body
    check(len(list(body.iter(qn("wp:anchor")))) == 1, "页面图片改为锚定在纸张上的浮动图片")
    check(len(list(body.iter(qn("wp:inline")))) == 0, "不再有内联图片")
    check("240.00" in text, f"文字层写入修复后的金额，实际={text!r}")
    check("天津鑫地" in text, f"文字层写入中文，实际={text!r}")
    # Reading order comes from the word boxes, not from the engine's own order.
    check(text.index("天津鑫地") < text.index("240.00"), "按词框坐标重排阅读顺序")


def test_ocr_failure_keeps_the_document(tmp_path: Path):
    if not HAS_PYMUPDF:
        print("SKIP  OCR 失败处理（本环境没有 pymupdf）")
        return
    import pymupdf

    source = tmp_path / "scan2.pdf"
    pdf = pymupdf.open()
    pdf.new_page(width=595, height=841)
    pdf.save(source)
    pdf.close()

    document = Document()
    engine.configure_section(document.sections[0], 595.0, 841.0)
    document.add_paragraph("原有内容")
    output = tmp_path / "scan2.docx"
    document.save(output)

    report = engine.add_ocr_text_layer(output, source, [1], str(tmp_path / "missing.cmd"), tmp_path)
    check(report["ocr_pages"] == [], "OCR 工具不存在时不声称识别过")
    check(bool(report["ocr_error"]), f"记录失败原因，实际={report['ocr_error']!r}")
    check("原有内容" in engine.document_text(Document(output)), "失败时文档保持原样")


def main() -> int:
    import tempfile

    with tempfile.TemporaryDirectory() as temp:
        tmp_path = Path(temp)
        test_document_text_finds_hidden_numbers()
        test_missing_numeric_tokens_normalizes()
        test_absent_values_versus_occurrence_counts()
        test_separators_do_not_glue_numbers()
        test_tolerance_thresholds()
        test_repair_walks_nested_content(tmp_path)
        test_compact_wide_tables_handles_nested(tmp_path)
        test_fit_images_to_page(tmp_path)
        test_shrink_docx_images(tmp_path)
        test_convert_delivers_document_with_warning(tmp_path)
        test_scanned_pdf_warns_that_text_is_not_editable(tmp_path)
        test_convert_still_fails_on_real_data_loss(tmp_path)
        test_repair_ocr_text()
        test_ocr_text_layer(tmp_path)
        test_ocr_failure_keeps_the_document(tmp_path)

    if FAILURES:
        print(f"\n{len(FAILURES)} 项失败")
        return 1
    print("\n全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
