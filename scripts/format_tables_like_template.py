#!/usr/bin/env python3
"""Center ordinary DOCX tables and force printable table borders to solid lines.

The document package is copied directly. By default only w:tblPr/w:jc is
added or updated, and table border line types are normalized to w:val="single".
With --repeat-first-row, the sole extra pagination change is w:tblHeader on each
table's original first row, so the header repeats when that table naturally
continues to a new page.
"""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path


TBL_PR = re.compile(rb"(<w:tblPr(?:\s[^>]*)?>)(.*?)(</w:tblPr>)", re.DOTALL)
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


def center_table_properties(match: re.Match[bytes]) -> bytes:
    opening, content, closing = match.groups()
    content = force_tbl_borders(content)
    # A floating table can have separately positioned VML lines or drawings
    # (for example a diagonal header line). Its tblpPr/jc/tblInd form one
    # coordinate system with those objects, so changing the table position
    # would visually detach them. Preserve the whole outer positioning block.
    if TBL_POSITION.search(content):
        return opening + content + closing

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


def apply(input_path: Path, output_path: Path, repeat_first_row: bool) -> tuple[int, int]:
    """Copy a DOCX while changing only explicitly requested table metadata."""
    with zipfile.ZipFile(input_path, "r") as source:
        document_xml = source.read("word/document.xml")
        updated_xml, centered = TBL_PR.subn(center_table_properties, document_xml)
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
    return centered, repeated


def main() -> None:
    parser = argparse.ArgumentParser(description="普通嵌入式表格整体居中；浮动表格保持原定位；强制所有表格边框为single实线；可选设置原首行为续页表头")
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--repeat-first-row", action="store_true", help="仅为续页重复原首行表头")
    args = parser.parse_args()
    if args.input.resolve() == args.output.resolve():
        raise SystemExit("禁止覆盖输入文件，请指定新输出路径。")
    centered, repeated = apply(args.input, args.output, args.repeat_first_row)
    print(f"已处理表格：{centered} 个（浮动表格已保留原定位）；已强制表格边框为 single 实线；已设置续页表头：{repeated} 个")


if __name__ == "__main__":
    main()
