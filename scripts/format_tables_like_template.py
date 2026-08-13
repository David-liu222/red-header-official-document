#!/usr/bin/env python3
"""Center DOCX tables, disable text wrapping, force solid borders, and clean captions.

The document package is copied directly. By default only w:tblPr/w:jc is
added or updated, floating table positioning (w:tblpPr) is removed to prevent
surrounding text from wrapping vertically beside a table, and table border line
types are normalized to w:val="single". With --repeat-first-row, the sole extra
pagination change is w:tblHeader on each table's original first row, so the
header repeats when that table naturally continues to a new page.

The script also removes consecutive empty paragraphs immediately before a table
when those paragraphs contain no text, page break, section properties, drawing,
object, or field content. This fixes the common WPS/Word conversion artifact
where a table caption and the table are separated by a large blank gap.
"""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path


TBL_PR = re.compile(rb"(<w:tblPr(?:\s[^>]*)?>)(.*?)(</w:tblPr>)", re.DOTALL)
TABLE_OPEN = re.compile(rb"<w:tbl(?:\s[^>]*)?>")
P_OPEN = re.compile(rb"<w:p(?:\s[^>]*)?>")
TBL_BORDERS = re.compile(rb"(<w:tblBorders(?:\s[^>]*)?>)(.*?)(</w:tblBorders>)", re.DOTALL)
TC_BORDERS = re.compile(rb"(<w:tcBorders(?:\s[^>]*)?>)(.*?)(</w:tcBorders>)", re.DOTALL)
JC = re.compile(rb"<w:jc\b[^>]*/>")
TBL_IND = re.compile(rb"<w:tblInd\b[^>]*/>")
TBL_POSITION = re.compile(rb"<w:tblpPr\b[^>]*/>")
VAL_ATTR = re.compile(rb'\sw:val="[^"]*"')
TABLE_BORDER_NAMES = (b"top", b"left", b"bottom", b"right", b"insideH", b"insideV")
ALL_BORDER_NAMES = TABLE_BORDER_NAMES + (b"tl2br", b"tr2bl")
FIRST_ROW = re.compile(
    rb"(<w:tblPr(?:\s[^>]*)?>.*?</w:tblPr>.*?<w:tr(?:\s[^>]*)?>)(.*?)(</w:tr>)",
    re.DOTALL,
)
TR_PR = re.compile(rb"(<w:trPr(?:\s[^>]*)?>.*?)(</w:trPr>)", re.DOTALL)
CONTROL_IN_EMPTY_PARAGRAPH = re.compile(
    rb"<w:(?:t|br|sectPr|drawing|pict|object|fldSimple|instrText|footnoteReference|endnoteReference)\b",
    re.DOTALL,
)


def is_safe_empty_paragraph_before_table(paragraph: bytes) -> bool:
    """Return True only for inert empty paragraphs safe to remove before tables."""
    if CONTROL_IN_EMPTY_PARAGRAPH.search(paragraph):
        return False
    text_values = re.findall(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", paragraph, re.DOTALL)
    return all(value.strip() == b"" for value in text_values)


def remove_empty_paragraphs_before_tables(document_xml: bytes) -> tuple[bytes, int]:
    """Remove consecutive inert empty paragraphs directly before each table."""
    removed = 0
    position = 0
    while True:
        table_match = TABLE_OPEN.search(document_xml, position)
        if table_match is None:
            break
        table_start = table_match.start()
        previous_end = document_xml.rfind(b"</w:p>", 0, table_start)
        if previous_end < 0:
            position = table_match.end()
            continue
        previous_end += len(b"</w:p>")
        if document_xml[previous_end:table_start].strip():
            position = table_match.end()
            continue
        previous_paragraphs = list(P_OPEN.finditer(document_xml, 0, previous_end))
        if not previous_paragraphs:
            position = table_match.end()
            continue
        previous_start = previous_paragraphs[-1].start()
        paragraph = document_xml[previous_start:previous_end]
        if is_safe_empty_paragraph_before_table(paragraph):
            document_xml = document_xml[:previous_start] + document_xml[previous_end:]
            removed += 1
            position = previous_start
            continue
        position = table_match.end()
    return document_xml, removed


def default_border(name: bytes) -> bytes:
    return b'<w:' + name + b' w:val="single" w:sz="4" w:space="0" w:color="auto"/>'


def force_border_element_single(element: bytes) -> bytes:
    """Change only the border line type, preserving width/color/space/etc."""
    if VAL_ATTR.search(element):
        return VAL_ATTR.sub(b' w:val="single"', element, count=1)
    return element[:-2] + b' w:val="single"/>'


def force_border_children_single(content: bytes) -> bytes:
    for name in ALL_BORDER_NAMES:
        pattern = re.compile(rb"<w:" + name + rb"\b[^>]*/>")
        content = pattern.sub(lambda item: force_border_element_single(item.group(0)), content)
    return content


def force_tbl_borders(content: bytes) -> bytes:
    """Ensure every table has printable solid outer and inner grid borders."""
    def update_container(match: re.Match[bytes]) -> bytes:
        opening, inner, closing = match.groups()
        inner = force_border_children_single(inner)
        for name in TABLE_BORDER_NAMES:
            if not re.search(rb"<w:" + name + rb"\b", inner):
                inner += default_border(name)
        return opening + inner + closing

    if TBL_BORDERS.search(content):
        return TBL_BORDERS.sub(update_container, content, count=1)
    return content + b"<w:tblBorders>" + b"".join(default_border(name) for name in TABLE_BORDER_NAMES) + b"</w:tblBorders>"


def force_cell_borders(content: bytes) -> bytes:
    """Normalize existing cell and diagonal borders without adding cell layout."""
    return TC_BORDERS.sub(
        lambda match: match.group(1) + force_border_children_single(match.group(2)) + match.group(3),
        content,
    )


def center_table_properties(match: re.Match[bytes], preserve_floating: bool) -> bytes:
    opening, content, closing = match.groups()
    content = force_tbl_borders(content)
    # Floating tables make Word/WPS wrap later paragraphs in the narrow space
    # beside the table. In red-header official documents this creates the
    # observed "文字跑到表格左侧/竖排" defect, so remove only the outer floating
    # positioning by default. Table rows, cells, text, widths, merges and
    # diagonal borders remain untouched.
    if preserve_floating and TBL_POSITION.search(content):
        return opening + content + closing
    content = TBL_POSITION.sub(b"", content)

    centered = b'<w:jc w:val="center"/>'
    # For ordinary inline tables, tblInd can oppose w:jc in WPS. Remove only
    # that external left indent; never touch rows, cells, or their paragraphs.
    content = TBL_IND.sub(b"", content)
    if JC.search(content):
        content = JC.sub(centered, content, count=1)
    else:
        content += centered
    return opening + content + closing


def mark_first_row_as_repeat_header(match: re.Match[bytes]) -> bytes:
    """Add only the necessary repeat-header metadata, never cell content."""
    opening, content, closing = match.groups()
    if b"<w:tblHeader" in content:
        return match.group(0)
    marker = b'<w:tblHeader w:val="true"/>'
    if TR_PR.search(content):
        content = TR_PR.sub(
            lambda item: item.group(1) + marker + item.group(2), content, count=1
        )
    else:
        content = b"<w:trPr>" + marker + b"</w:trPr>" + content
    return opening + content + closing


def apply(input_path: Path, output_path: Path, repeat_first_row: bool, preserve_floating: bool) -> tuple[int, int, int, int]:
    """Copy a DOCX while changing only explicitly requested table metadata."""
    with zipfile.ZipFile(input_path, "r") as source:
        document_xml = source.read("word/document.xml")
        document_xml, pretable_blanks_removed = remove_empty_paragraphs_before_tables(document_xml)
        floating_removed = 0 if preserve_floating else len(TBL_POSITION.findall(document_xml))
        updated_xml, centered = TBL_PR.subn(
            lambda match: center_table_properties(match, preserve_floating),
            document_xml,
        )
        updated_xml = force_cell_borders(updated_xml)
        if centered == 0:
            raise ValueError("未在 word/document.xml 中找到表格。")
        repeated = 0
        if repeat_first_row:
            updated_xml, repeated = FIRST_ROW.subn(mark_first_row_as_repeat_header, updated_xml)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w") as target:
            for info in source.infolist():
                data = updated_xml if info.filename == "word/document.xml" else source.read(info.filename)
                copied = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                copied.comment = info.comment
                copied.extra = info.extra
                copied.internal_attr = info.internal_attr
                copied.external_attr = info.external_attr
                copied.create_system = info.create_system
                copied.create_version = info.create_version
                copied.extract_version = info.extract_version
                copied.flag_bits = info.flag_bits
                copied.compress_type = info.compress_type
                target.writestr(copied, data, compress_type=info.compress_type)
    return centered, repeated, floating_removed, pretable_blanks_removed


def main() -> None:
    parser = argparse.ArgumentParser(description="表格整体居中；默认取消浮动/文字环绕；清理表格前空段；强制所有表格边框为single实线；可选设置原首行为续页表头")
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--repeat-first-row", action="store_true", help="仅为续页重复原首行表头")
    parser.add_argument("--preserve-floating", action="store_true", help="保留 w:tblpPr 浮动定位；仅在正式模板明确要求时使用")
    args = parser.parse_args()
    if args.input.resolve() == args.output.resolve():
        raise SystemExit("禁止覆盖输入文件，请指定新输出路径。")
    centered, repeated, floating_removed, pretable_blanks_removed = apply(args.input, args.output, args.repeat_first_row, args.preserve_floating)
    floating_msg = "已保留浮动定位" if args.preserve_floating else f"已取消浮动/文字环绕定位：{floating_removed} 处"
    print(f"已处理表格：{centered} 个；{floating_msg}；已清理表格前空段：{pretable_blanks_removed} 个；已强制表格边框为 single 实线；已设置续页表头：{repeated} 个")


if __name__ == "__main__":
    main()
