#!/usr/bin/env python3
"""Read-only structural validator for a red-header official-document DOCX.

This intentionally reports evidence and does not rewrite the document. It is a
first-pass check; red-header color/line placement and visual balance still need
rendered PDF/PNG inspection.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Mm, Pt
except ImportError as exc:  # pragma: no cover - environment-dependent
    raise SystemExit("缺少 python-docx，请在工作环境安装后重试。") from exc


def east_asia_font(run):
    """Return the East Asian font name when it is explicitly present."""
    rpr = run._element.rPr
    if rpr is None or rpr.rFonts is None:
        return None
    return rpr.rFonts.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia")


def mm(value):
    return round(value / Mm(1), 2)


def pt(value):
    return None if value is None else round(value / Pt(1), 2)


def nonempty_paragraphs(doc):
    return [(idx + 1, p) for idx, p in enumerate(doc.paragraphs) if p.text.strip()]


def has_paragraph_border(paragraph):
    return bool(paragraph._p.xpath("./w:pPr/w:pBdr"))


def main():
    parser = argparse.ArgumentParser(description="检查套红头 DOCX 的常见固定规则（只读）")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--title", help="正文标题，用于校验文件名和标题段落")
    parser.add_argument("--expected-company", default="内蒙古准格尔旗力量煤业有限公司")
    parser.add_argument("--expected-mine", help="需要校验时传入矿名；不传则不强制要求矿名")
    parser.add_argument("--body-bold", action="store_true", help="按规则批注强制检查正文加粗；默认按66号定稿检查正文不加粗")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    args = parser.parse_args()

    if args.docx.suffix.lower() != ".docx":
        raise SystemExit("只接受 .docx；旧 .doc 请先在副本上转换。")
    doc = Document(str(args.docx))
    checks = []

    def add(rule, status, evidence, severity="info"):
        checks.append({"rule": rule, "status": status, "evidence": evidence, "severity": severity})

    for idx, section in enumerate(doc.sections, 1):
        margins = {
            "top_mm": mm(section.top_margin),
            "bottom_mm": mm(section.bottom_margin),
            "left_mm": mm(section.left_margin),
            "right_mm": mm(section.right_margin),
        }
        ok = all(21 <= value <= 25 for value in margins.values())
        add(f"第{idx}节页边距 23mm±2mm", "pass" if ok else "review", margins, "error" if not ok else "info")

    paragraphs = nonempty_paragraphs(doc)
    if not paragraphs:
        add("存在正文段落", "fail", "未找到非空段落", "error")
    else:
        add("存在正文段落", "pass", f"共 {len(paragraphs)} 个非空段落")

    title_hits = []
    if args.title:
        title_hits = [(idx, p) for idx, p in paragraphs if p.text.strip() == args.title]
        add("正文标题与传入标题一致", "pass" if title_hits else "review", [idx for idx, _ in title_hits], "error" if not title_hits else "info")
        name_ok = args.docx.stem == args.title
        add("文件名与标题一致", "pass" if name_ok else "review", {"filename": args.docx.stem, "title": args.title}, "error" if not name_ok else "info")
        title_issues = []
        for idx, paragraph in title_hits:
            for run in paragraph.runs:
                if not run.text.strip():
                    continue
                font_name = east_asia_font(run) or run.font.name
                size = pt(run.font.size)
                if font_name and "黑体" not in font_name and "SimHei" not in font_name:
                    title_issues.append({"paragraph": idx, "font": font_name, "text": run.text[:20]})
                if size is not None and abs(size - 22) > 0.5:
                    title_issues.append({"paragraph": idx, "size_pt": size, "text": run.text[:20]})
                if run.bold is True:
                    title_issues.append({"paragraph": idx, "bold": True, "text": run.text[:20]})
        add("标题黑体二号且不加粗", "pass" if title_hits and not title_issues else "review", title_issues[:20] if title_issues else [idx for idx, _ in title_hits], "error" if title_issues else "info")
        title_spacing_issues = []
        for idx, paragraph in title_hits:
            before = pt(paragraph.paragraph_format.space_before)
            after = pt(paragraph.paragraph_format.space_after)
            if before not in (None, 0.0):
                title_spacing_issues.append({"paragraph": idx, "space_before_pt": before})
            if after not in (None, 0.0):
                title_spacing_issues.append({"paragraph": idx, "space_after_pt": after})
        add("标题段前段后均为0", "pass" if title_hits and not title_spacing_issues else "review", title_spacing_issues[:20] if title_spacing_issues else [idx for idx, _ in title_hits], "error" if title_spacing_issues else "info")
        title_line_spacing_issues = []
        for idx, paragraph in title_hits:
            spacing = paragraph._p.pPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing") if paragraph._p.pPr is not None else None
            if spacing is not None:
                for attribute in ("before", "after", "beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
                    value = spacing.get(f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{attribute}")
                    if value is not None:
                        title_line_spacing_issues.append({"paragraph": idx, attribute: value})
        add("标题段前段后继承模板0行显示", "pass" if title_hits and not title_line_spacing_issues else "review", title_line_spacing_issues[:20] if title_line_spacing_issues else [idx for idx, _ in title_hits], "error" if title_line_spacing_issues else "info")
        title_style_issues = []
        for idx, paragraph in title_hits:
            if paragraph.style.name != "Normal":
                title_style_issues.append({"paragraph": idx, "style": paragraph.style.name, "expected": "Normal"})
        add("标题使用正文文本Normal样式", "pass" if title_hits and not title_style_issues else "review", title_style_issues[:20] if title_style_issues else [idx for idx, _ in title_hits], "error" if title_style_issues else "info")

    recipient_paragraph = next(
        (
            p
            for idx, p in paragraphs
            if not (args.title and p.text.strip() == args.title)
            and not has_paragraph_border(p)
            and not any(run.font.color and run.font.color.rgb for run in p.runs)
        ),
        None,
    )
    body_candidates = [(idx, p) for idx, p in paragraphs if not (args.title and p.text.strip() == args.title)]
    font_issues = []
    recipient_font_issues = []
    paragraph_issues = []
    semicolon_hits = []
    for idx, paragraph in body_candidates:
        # Red-head paragraphs are intentionally not body text.
        if has_paragraph_border(paragraph) or any(run.font.color and run.font.color.rgb for run in paragraph.runs):
            continue
        if paragraph.text.rstrip().endswith("；"):
            semicolon_hits.append(idx)
        pf = paragraph.paragraph_format
        before = pt(pf.space_before)
        after = pt(pf.space_after)
        if before not in (None, 0.0) or after not in (None, 0.0):
            paragraph_issues.append({"paragraph": idx, "space_before_pt": before, "space_after_pt": after})
        indent = pt(pf.first_line_indent)
        if indent is not None and not 26 <= indent <= 30:
            paragraph_issues.append({"paragraph": idx, "first_line_indent_pt": indent})
        if pf.line_spacing is not None and isinstance(pf.line_spacing, Pt):
            line_pt = pt(pf.line_spacing)
            if not 25 <= line_pt <= 29:
                paragraph_issues.append({"paragraph": idx, "line_spacing_pt": line_pt})
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            font_name = east_asia_font(run) or run.font.name
            size = pt(run.font.size)
            if font_name and "仿宋" not in font_name:
                font_issues.append({"paragraph": idx, "font": font_name, "text": run.text[:20]})
            if size is not None and abs(size - 14) > 0.5:
                font_issues.append({"paragraph": idx, "size_pt": size, "text": run.text[:20]})
            expected_bold = True if paragraph is recipient_paragraph else args.body_bold
            target = recipient_font_issues if paragraph is recipient_paragraph else font_issues
            if run.bold is not expected_bold:
                target.append({"paragraph": idx, "bold": run.bold, "expected_bold": expected_bold, "text": run.text[:20]})

    body_label = "正文仿宋四号且加粗" if args.body_bold else "正文仿宋四号、按66号定稿不加粗"
    add(body_label, "pass" if not font_issues else "review", font_issues[:20], "error" if font_issues else "info")
    add("主送单位仿宋四号且加粗", "pass" if not recipient_font_issues else "review", recipient_font_issues[:20], "error" if recipient_font_issues else "info")
    add("正文段前段后为0、固定行距在27±2磅内", "pass" if not paragraph_issues else "review", paragraph_issues[:20], "error" if paragraph_issues else "info")
    add("正文条目不以分号结尾", "pass" if not semicolon_hits else "review", semicolon_hits, "error" if semicolon_hits else "info")

    full_text = "\n".join(p.text for _, p in paragraphs)
    expected_values = [("公司名", args.expected_company)]
    if args.expected_mine:
        expected_values.append(("矿名", args.expected_mine))
    for label, value in expected_values:
        add(f"包含{label}", "pass" if value in full_text else "review", value, "error" if value not in full_text else "info")

    if not args.title:
        add("标题黑体二号且不加粗", "manual", "请传入 --title 后自动检查标题段落")
    add("红头颜色、红线、右下角落款、页码和上长下短", "manual", "需渲染 PDF/PNG 目测确认")
    add("序号层级、书名号间顿号、附件空行、结束语和落款分行", "manual", "需结合具体文稿语义确认")

    report = {"file": str(args.docx), "checks": checks, "summary": {"total": len(checks), "review": sum(c["status"] in {"review", "manual"} for c in checks), "failed": sum(c["status"] == "fail" for c in checks)}}
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"文件：{args.docx}")
        for check in checks:
            print(f"[{check['status']}] {check['rule']}：{check['evidence']}")
        print(f"汇总：{report['summary']}")


if __name__ == "__main__":
    main()
