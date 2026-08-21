# -*- coding: utf-8 -*-
"""Offline checks for the DOCX scanning and numeric integrity logic.

The conversion engines (``pymupdf``/``pdf2docx``) are only present in the build
environment, so they are stubbed out here: every function under test works on
an already produced Word document and needs ``python-docx`` alone.

    python tests/test_numeric_check.py
"""

from __future__ import annotations

import sys
import types
from pathlib import Path

for missing in ("pymupdf", "pdf2docx"):
    if missing not in sys.modules:
        stub = types.ModuleType(missing)
        stub.open = lambda *args, **kwargs: None  # type: ignore[attr-defined]
        stub.Converter = object  # type: ignore[attr-defined]
        sys.modules[missing] = stub

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "editable_engine"))

from docx import Document  # noqa: E402
from docx.oxml import parse_xml  # noqa: E402
from docx.oxml.ns import nsdecls  # noqa: E402

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


def stub_engines(source_numbers, layout_body, legacy_body=None):
    """Replace the PDF side of the pipeline with fixed text."""

    def make(body):
        def run(input_path, output_path):
            document = Document()
            document.add_paragraph(body)
            document.save(output_path)
            report = {"engine": "stub", "fallback_pages": []}
            return report, engine.document_text(document)

        return run

    engine.source_numeric_tokens = lambda input_path: list(source_numbers)
    engine.convert_layout = make(layout_body)
    engine.convert_legacy = lambda input_path, output_path: (
        make(legacy_body if legacy_body is not None else layout_body)(input_path, output_path)[0]
    )
    engine.repair_and_normalize_docx = lambda path: (
        {"repaired_glyphs": 0, "suspicious_glyphs": 0},
        engine.document_text(Document(path)),
    )


def test_convert_delivers_document_with_warning(tmp_path: Path):
    original = (engine.source_numeric_tokens, engine.convert_layout,
                engine.convert_legacy, engine.repair_and_normalize_docx)
    try:
        # Every value is present, only one repeated occurrence is gone.
        stub_engines(["7"] * 8 + [str(n) for n in range(100, 160)], "7 " + " ".join(str(n) for n in range(100, 160)))
        output = tmp_path / "warned.docx"
        report = engine.convert(Path("source.pdf"), output)
        check(output.exists(), "少量数字对不上时仍然输出 Word 文件")
        check(report["absent_numbers"] == [], "所有数值都还在，不报缺失")
        check(report["number_mismatches"] == 7, f"出现次数差异被记录，实际={report['number_mismatches']}")
        check(
            len(report["warnings"]) == 1 and "出现次数" in report["warnings"][0],
            f"超出计数容差时给出提示而不是失败，实际={report['warnings']}",
        )
    finally:
        (engine.source_numeric_tokens, engine.convert_layout,
         engine.convert_legacy, engine.repair_and_normalize_docx) = original


def test_convert_still_fails_on_real_data_loss(tmp_path: Path):
    original = (engine.source_numeric_tokens, engine.convert_layout,
                engine.convert_legacy, engine.repair_and_normalize_docx)
    try:
        # The subset-font case: the numbers never arrive in the document.
        stub_engines([str(n) for n in range(100, 160)], "文本没有任何数字")
        output = tmp_path / "failed.docx"
        try:
            engine.convert(Path("source.pdf"), output)
            check(False, "数值大面积丢失时应当报错")
        except ValueError as exc:
            check("数字完整性校验未通过" in str(exc), f"报错信息说明原因，实际={exc}")
            check(not output.exists(), "校验失败时不写出残缺的 Word 文件")
    finally:
        (engine.source_numeric_tokens, engine.convert_layout,
         engine.convert_legacy, engine.repair_and_normalize_docx) = original


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
        test_convert_delivers_document_with_warning(tmp_path)
        test_convert_still_fails_on_real_data_loss(tmp_path)

    if FAILURES:
        print(f"\n{len(FAILURES)} 项失败")
        return 1
    print("\n全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
