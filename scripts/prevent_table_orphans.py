#!/usr/bin/env python3
"""Move a table's preceding caption to the next page without touching the table.

Use only after rendering shows a complete table would fit on the next page but
is being split at the bottom of the current page. This edits the paragraph
immediately before the selected table; table XML remains byte-for-byte intact.
"""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path


TABLE_OPEN = re.compile(rb"<w:tbl(?:\s[^>]*)?>")
P_OPEN = re.compile(rb"<w:p(?:\s[^>]*)?>")
PPR = re.compile(rb"(<w:pPr(?:\s[^>]*)?>.*?)(</w:pPr>)", re.DOTALL)
PAGE_BREAK = re.compile(rb"<w:pageBreakBefore(?:\s[^>]*)?/>")


def add_page_break_before_table(document_xml: bytes, table_index: int) -> bytes:
    tables = list(TABLE_OPEN.finditer(document_xml))
    if not 1 <= table_index <= len(tables):
        raise ValueError(f"表格序号应在 1 至 {len(tables)} 之间。")
    table_start = tables[table_index - 1].start()
    previous_end = document_xml.rfind(b"</w:p>", 0, table_start)
    preceding_paragraphs = list(P_OPEN.finditer(document_xml, 0, previous_end))
    if not preceding_paragraphs or previous_end < 0:
        raise ValueError("未找到该表格前的标题或说明段落。")
    previous_start = preceding_paragraphs[-1].start()
    previous_end += len(b"</w:p>")
    paragraph = document_xml[previous_start:previous_end]
    if b'<w:pageBreakBefore w:val="true"' in paragraph or b"<w:pageBreakBefore/>" in paragraph:
        return document_xml
    marker = b"<w:pageBreakBefore/>"
    if PPR.search(paragraph):
        def update_properties(item: re.Match[bytes]) -> bytes:
            properties = item.group(1)
            if PAGE_BREAK.search(properties):
                properties = PAGE_BREAK.sub(marker, properties, count=1)
            else:
                properties += marker
            return properties + item.group(2)
        paragraph = PPR.sub(update_properties, paragraph, count=1)
    else:
        opening = P_OPEN.search(paragraph)
        if opening is None:
            raise ValueError("表格前的段落格式无效。")
        paragraph = (
            paragraph[:opening.end()]
            + b"<w:pPr>"
            + marker
            + b"</w:pPr>"
            + paragraph[opening.end():]
        )
    return document_xml[:previous_start] + paragraph + document_xml[previous_end:]


def apply(input_path: Path, output_path: Path, table_index: int) -> None:
    with zipfile.ZipFile(input_path, "r") as source:
        original = source.read("word/document.xml")
        updated = add_page_break_before_table(original, table_index)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w") as target:
            for info in source.infolist():
                data = updated if info.filename == "word/document.xml" else source.read(info.filename)
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


def main() -> None:
    parser = argparse.ArgumentParser(description="将表格前标题与表格一起移到下一页")
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--before-table-index", type=int, required=True)
    args = parser.parse_args()
    if args.input.resolve() == args.output.resolve():
        raise SystemExit("禁止覆盖输入文件，请指定新输出路径。")
    apply(args.input, args.output, args.before_table_index)
    print(f"已将第 {args.before_table_index} 张表格前的标题移到下一页。")


if __name__ == "__main__":
    main()
