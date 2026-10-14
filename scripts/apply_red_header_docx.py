#!/usr/bin/env python3
"""Apply the current red-header rules to an existing DOCX.

The source is read-only. The output is a new DOCX. This is intentionally
conservative: it adds the red-header block and normalizes formatting without
rewriting the document's substantive text.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from red_header_common import (attachment_start, closing, independent_heading, safe_blank, spacing_issues,
                               ppr_layers, bold_snapshot, mapped_bold, effective_bold, text_runs, furniture, tail_layout,
                               external_document_number)


RED = "FF0000"
BODY_FONT = "仿宋"
TITLE_FONT = "黑体"
# The approved 66号 .doc has no direct rFonts on the red-header runs. Its
# effective CJK paragraph default exports as 宋体, so preserve that template
# behavior instead of substituting a conventional but unproven font.
HEADER_FONT = "宋体"
LATIN_FONT = "Times New Roman"
BODY_LINE_PT = 28
RED_LINE_WIDTH_PT = 504
# The approved 80号 DOCX uses paragraph DISTRIBUTE for both red-header lines,
# with no run-level tracking. This also gives the two lines matching edges.
SIGNER_LABEL_TRACKING_TWIPS = 140
SIGNER_RIGHT_INDENT_PT = 68
DEFAULT_NUMBER_LINE_PT = 27
PUBLIC_STYLE = "红头公文基准"


def public_style(doc):
    """An independent style for public paragraphs; never edit protected table styles."""
    existing = PUBLIC_STYLE in doc.styles
    style = doc.styles[PUBLIC_STYLE] if existing else doc.styles.add_style(PUBLIC_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    element = deepcopy(style._element) if existing else style._element
    for node in list(element.findall(qn('w:basedOn'))):
        element.remove(node)
    ppr = element.get_or_add_pPr()
    for tag in ("ind", "spacing", "snapToGrid"):
        for node in list(ppr.findall(qn('w:' + tag))):
            ppr.remove(node)
    ind = OxmlElement('w:ind')
    for key in ('left', 'right', 'start', 'end', 'firstLine', 'firstLineChars'):
        ind.set(qn('w:' + key), '0')
    ppr.append(ind)
    spacing = OxmlElement('w:spacing')
    for key, value in {'before': '0', 'after': '0', 'beforeLines': '0', 'afterLines': '0',
                       'beforeAutospacing': '0', 'afterAutospacing': '0', 'line': '560', 'lineRule': 'exact'}.items():
        spacing.set(qn('w:' + key), value)
    ppr.append(spacing)
    snap = OxmlElement('w:snapToGrid')
    snap.set(qn('w:val'), '0')
    ppr.append(snap)
    def definition(node):
        # Ignore unused namespace declarations added by the enclosing styles part.
        return (node.tag, tuple(sorted(node.attrib.items())), (node.text or '').strip(),
                tuple(definition(child) for child in node))
    if existing and definition(element) != definition(style._element):
        raise ValueError("原稿已有不同定义的“红头公文基准”样式，不能全局改写同名样式；请在副本上明确处理样式冲突。")
    return style


def use_public_style(paragraph):
    # Materialize inherited automatic numbering before changing paragraph style.
    inherited_num = next((p.find(qn('w:numPr')) for p in ppr_layers(paragraph) if p.find(qn('w:numPr')) is not None), None)
    if inherited_num is not None and paragraph._p.get_or_add_pPr().find(qn('w:numPr')) is None:
        paragraph._p.pPr.append(deepcopy(inherited_num))
    paragraph.style = PUBLIC_STYLE


def reset_indent(paragraph, chars=0):
    ppr = paragraph._p.get_or_add_pPr()
    for node in list(ppr.findall(qn('w:ind'))):
        ppr.remove(node)
    ind = OxmlElement('w:ind')
    for key in ('left', 'right', 'start', 'end'):
        ind.set(qn('w:' + key), '0')
    ind.set(qn('w:firstLine'), str(round(chars * 280)))
    ind.set(qn('w:firstLineChars'), str(round(chars * 100)))
    ppr.append(ind)


def disable_grid(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    node = ppr.find(qn('w:snapToGrid'))
    if node is None:
        node = OxmlElement('w:snapToGrid')
        ppr.append(node)
    node.set(qn('w:val'), '0')


def clear_old_spacing(paragraph):
    spacing = paragraph._p.get_or_add_pPr().find(qn('w:spacing'))
    if spacing is not None:
        for key in ('beforeLines', 'afterLines', 'beforeAutospacing', 'afterAutospacing'):
            spacing.attrib.pop(qn('w:' + key), None)


def ensure_one_blank_before(anchor):
    previous = anchor._p.getprevious()
    blanks = []
    from docx.text.paragraph import Paragraph
    while previous is not None and previous.tag == qn('w:p'):
        candidate = Paragraph(previous, anchor._parent)
        if not safe_blank(candidate):
            break
        blanks.append(previous)
        previous = previous.getprevious()
    if not blanks:
        blank = anchor.insert_paragraph_before()
        use_public_style(blank)
        set_spacing(blank, indent_pt=None)
        blank.paragraph_format.keep_with_next = False
    else:
        for node in blanks[1:]:
            node.getparent().remove(node)


def normalize_blank_separators(doc, signature_start):
    paragraphs = list(doc.paragraphs)
    for i, p in enumerate(paragraphs):
        if attachment_start(p.text):
            previous = i - 1
            while previous >= 0 and safe_blank(paragraphs[previous]):
                previous -= 1
            if previous >= 0 and closing(paragraphs[previous].text):
                ensure_one_blank_before(p)
        elif closing(p.text):
            previous = i - 1
            while previous >= 0 and safe_blank(paragraphs[previous]):
                previous -= 1
            # Only a clearly numbered attachment block; unknown intervening content stays intact.
            if previous >= 0 and re.match(r'^\d+[.．]', paragraphs[previous].text.strip()):
                if any(attachment_start(x.text) for x in paragraphs[:previous + 1]):
                    ensure_one_blank_before(p)
    if signature_start is not None:
        ensure_one_blank_before(signature_start)


def restore_source_bold(paragraph, snapshot):
    """Retain source emphasis by character range after title/date reflow."""
    from docx.text.run import Run
    mask = mapped_bold(snapshot, paragraph.text)
    offset = 0
    for run in text_runs(paragraph):
        text = run.text
        desired = [effective_bold(run) or flag for flag in mask[offset:offset + len(text)]]
        offset += len(text)
        if not text or not any(desired):
            continue
        groups = []
        for char, weight in zip(text, desired):
            if groups and groups[-1][1] == weight:
                groups[-1][0] += char
            else:
                groups.append([char, weight])
        if len(groups) > 1 and any(child.tag not in (qn('w:rPr'), qn('w:t'), qn('w:br'), qn('w:tab'), qn('w:cr')) for child in run._r):
            raise ValueError("原稿局部加粗涉及对象/域，须在副本上明确保留范围。")
        for value, weight in groups:
            element = deepcopy(run._r)
            new = Run(element, paragraph)
            if len(groups) > 1:
                new.text = value
            new.bold = weight
            node = new._r.get_or_add_rPr().find(qn('w:bCs'))
            if node is None:
                node = OxmlElement('w:bCs')
                new._r.rPr.append(node)
            node.set(qn('w:val'), '1' if weight else '0')
            run._r.addprevious(element)
        run._r.getparent().remove(run._r)


def trim_post_signature_blank_pages(doc):
    anchor, blanks, breaks, applicable = tail_layout(doc)
    if not applicable:
        return
    # The first terminal section break describes the preceding real content.
    # Carry its settings forward instead of inheriting those of an empty section.
    section = next((n.find(qn('w:pPr')).find(qn('w:sectPr')) for n in blanks
                    if n.find(qn('w:pPr')) is not None and n.find(qn('w:pPr')).find(qn('w:sectPr')) is not None), None)
    if section is not None:
        body = doc._element.body
        old = body.find(qn('w:sectPr'))
        if old is not None:
            body.remove(old)
        body.append(deepcopy(section))
    for node in breaks + blanks:
        node.getparent().remove(node)


def enable_title_auto_spacing(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    for tag in ("autoSpaceDN", "autoSpaceDE"):
        element = ppr.find(qn(f"w:{tag}"))
        if element is None:
            element = OxmlElement(f"w:{tag}")
            ppr.append(element)
        element.set(qn("w:val"), "1")


def format_issuing_number(paragraph, section):
    """Reproduce the approved mixed-font number and wider red rule."""
    for run in text_runs(paragraph):
        set_run(run, font=HEADER_FONT, size=14, bold=False, color="000000")
        fonts = run._r.get_or_add_rPr().rFonts
        for key in ("ascii", "hAnsi", "cs"):
            fonts.set(qn(f"w:{key}"), LATIN_FONT)
    ppr = paragraph._p.get_or_add_pPr()
    for element in list(ppr.findall(qn("w:ind"))):
        ppr.remove(element)
    body_width = (section.page_width - section.left_margin - section.right_margin) / Pt(1)
    extension = (RED_LINE_WIDTH_PT - body_width) / 2
    paragraph.paragraph_format.left_indent = Pt(-extension)
    paragraph.paragraph_format.right_indent = Pt(-extension)
    ind = paragraph._p.get_or_add_pPr().find(qn('w:ind'))
    ind.set(qn('w:start'), str(round(-extension * 20)))
    ind.set(qn('w:end'), str(round(-extension * 20)))
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_bottom_border(paragraph)


def canonical_title(text: str) -> str:
    return text.replace("\n", "").replace("\r", "").strip()


def official_filename(title: str, issuing_number: str) -> str:
    number = issuing_number.strip() or "〔发文字号待补〕"
    return f"{number}：{canonical_title(title)}.docx"


def distribute_red_header(paragraph):
    """Match the approved DOCX: distribute each line across the same text area."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.DISTRIBUTE
    ppr = paragraph._p.get_or_add_pPr()
    for element in list(ppr.findall(qn("w:ind"))):
        ppr.remove(element)
    for run in text_runs(paragraph):
        run.text = run.text.replace(" ", "").replace("\u3000", "")
        rpr = run._r.get_or_add_rPr()
        for old in list(rpr.findall(qn("w:spacing"))):
            rpr.remove(old)


def set_title_lines(paragraph, lines: list[str]):
    if not lines or any(not line.strip() for line in lines):
        raise ValueError("标题行不能为空。")
    if canonical_title("".join(lines)) != canonical_title(paragraph.text):
        raise ValueError("标题换行不得改变原标题文字。")
    replace_text(paragraph, "\n".join(lines))
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.keep_together = True
    for run in text_runs(paragraph):
        set_run(run, font=TITLE_FONT, size=22, bold=False)
    set_single_spacing(paragraph)
    clear_title_before_after_spacing(paragraph)
    enable_title_auto_spacing(paragraph)


def set_east_asia_font(run, name: str):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme"):
        rfonts.attrib.pop(qn(f"w:{key}"), None)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{key}"), name)


def set_run(run, *, font: str, size: float, bold: bool | None, color: str | None = None):
    set_east_asia_font(run, font)
    run.font.size = Pt(size)
    run.bold = bold
    rpr = run._r.get_or_add_rPr()
    for tag, value in (("szCs", str(round(size * 2))), ("bCs", "1" if bold else "0")):
        node = rpr.find(qn('w:' + tag))
        if node is None:
            node = OxmlElement('w:' + tag)
            rpr.append(node)
        node.set(qn('w:val'), value)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def clear_paragraph(paragraph):
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def remove_paragraph(paragraph):
    paragraph._p.getparent().remove(paragraph._p)


def replace_text(paragraph, text: str):
    if paragraph._p.xpath('.//w:drawing | .//w:pict | .//w:object | .//w:fldChar | .//w:instrText | .//w:hyperlink'):
        raise ValueError("文字调整涉及对象/域/超链接，须在副本上明确处理。")
    # Keep comment ranges, references and bookmarks while replacing visible text.
    runs = paragraph.runs
    text_run = next((r for r in runs if r._r.xpath('./w:t')), None)
    if text_run is None:
        text_run = paragraph.add_run()
    for r in paragraph.runs:
        for node in list(r._r):
            if node.tag in (qn('w:t'), qn('w:br'), qn('w:tab')):
                if node.tag == qn('w:br') and node.get(qn('w:type')) == 'page':
                    raise ValueError("文字调整会覆盖手动分页，须明确处理。")
                r._r.remove(node)
    for index, line in enumerate(text.split('\n')):
        if index:
            text_run._r.append(OxmlElement('w:br'))
        node = OxmlElement('w:t')
        node.set(qn('xml:space'), 'preserve')
        node.text = line
        text_run._r.append(node)
    return text_run


def set_spacing(paragraph, *, indent_pt: float | None):
    clear_old_spacing(paragraph)
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(BODY_LINE_PT)
    reset_indent(paragraph, chars=0 if indent_pt is None else indent_pt / 14)
    disable_grid(paragraph)


def set_single_spacing(paragraph):
    clear_old_spacing(paragraph)
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.line_spacing = 1.0
    disable_grid(paragraph)


def clear_title_before_after_spacing(paragraph):
    """Match the approved template's inherited 0-line before/after spacing.

    WPS displays a title with no direct before/after spacing as “0 行”. Writing
    point-based zero or explicit beforeLines/afterLines zero instead displays
    “0 磅” or “自动”, which does not match the supplied template.
    """
    ppr = paragraph._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is not None:
        for attribute in ("before", "after", "beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
            spacing.attrib.pop(qn(f"w:{attribute}"), None)


def set_exact_spacing(paragraph, points: float):
    clear_old_spacing(paragraph)
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(points)
    disable_grid(paragraph)


def set_bottom_border(paragraph, color=RED, size="16"):
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        ppr.append(pbdr)
    bottom = pbdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        pbdr.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)


def remove_paragraph_border(paragraph, side="bottom"):
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        return
    border = borders.find(qn(f"w:{side}"))
    if border is not None:
        borders.remove(border)
    if len(borders) == 0:
        ppr.remove(borders)


def set_paragraph_border(paragraph, side, *, color="000000", size="8", space="1"):
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    border = borders.find(qn(f"w:{side}"))
    if border is None:
        border = OxmlElement(f"w:{side}")
        borders.append(border)
    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), size)
    border.set(qn("w:space"), space)
    border.set(qn("w:color"), color)


def add_before(anchor, text, *, size, font, bold, color=None, border=False, line_spacing=None):
    paragraph = anchor._parent.add_paragraph()
    anchor._p.addprevious(paragraph._p)
    use_public_style(paragraph)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if line_spacing is None:
        set_single_spacing(paragraph)
    else:
        set_exact_spacing(paragraph, line_spacing)
    pf.keep_with_next = True
    run = paragraph.add_run(text)
    set_run(run, font=font, size=size, bold=bold, color=color)
    if border:
        set_bottom_border(paragraph)
    return paragraph


def add_internal_signatory(anchor, section, signer="", *,
                           label_tracking_twips=SIGNER_LABEL_TRACKING_TWIPS,
                           right_indent_pt=SIGNER_RIGHT_INDENT_PT):
    paragraph = add_before(
        anchor,
        "",
        size=14,
        font=HEADER_FONT,
        bold=False,
        line_spacing=20,
    )
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.paragraph_format.left_indent = Pt(0)
    paragraph.paragraph_format.right_indent = Pt(right_indent_pt)
    paragraph._p.get_or_add_pPr().find(qn('w:ind')).set(qn('w:end'), str(round(right_indent_pt * 20)))
    label = paragraph.runs[0]
    label.text = "签发"
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:val"), str(label_tracking_twips))
    label._r.get_or_add_rPr().append(spacing)
    label_tail = paragraph.add_run("人：")
    set_run(label_tail, font=HEADER_FONT, size=14, bold=False)
    if signer.strip():
        name = paragraph.add_run(signer.strip())
        set_run(name, font=HEADER_FONT, size=14, bold=False)
    return paragraph


def add_internal_red_rule(anchor, section):
    """Keep the wide red rule independent of the signatory's right indent."""
    paragraph = add_before(
        anchor,
        "",
        size=10,
        font=BODY_FONT,
        bold=False,
        line_spacing=2,
    )
    body_width = (section.page_width - section.left_margin - section.right_margin) / Pt(1)
    extension = (RED_LINE_WIDTH_PT - body_width) / 2
    paragraph.paragraph_format.left_indent = Pt(-extension)
    paragraph.paragraph_format.right_indent = Pt(-extension)
    ind = paragraph._p.get_or_add_pPr().find(qn('w:ind'))
    ind.set(qn('w:start'), str(round(-extension * 20)))
    ind.set(qn('w:end'), str(round(-extension * 20)))
    paragraph.paragraph_format.space_after = Pt(3)
    set_bottom_border(paragraph)
    return paragraph


def add_internal_distribution_block(doc, *, title, copy_to, issue_unit, print_count):
    if any(p.text.strip().startswith("主题：") for p in doc.paragraphs):
        raise ValueError("文档已存在内发印发栏，禁止重复添加。")
    section = doc.sections[-1]
    usable_width = section.page_width - section.left_margin - section.right_margin
    values = [
        f"主题：{canonical_title(title)}",
        f"抄送：{copy_to.strip().rstrip('：:')}",
    ]
    block = []
    for text in values:
        paragraph = doc.add_paragraph()
        use_public_style(paragraph)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.first_line_indent = Pt(0)
        paragraph.paragraph_format.keep_with_next = True
        paragraph.paragraph_format.keep_together = True
        set_exact_spacing(paragraph, 17)
        set_run(paragraph.add_run(text), font=BODY_FONT, size=12, bold=False)
        block.append(paragraph)
    paragraph = doc.add_paragraph()
    use_public_style(paragraph)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = Pt(0)
    paragraph.paragraph_format.keep_together = True
    set_exact_spacing(paragraph, 17)
    paragraph.paragraph_format.tab_stops.add_tab_stop(usable_width, WD_TAB_ALIGNMENT.RIGHT)
    set_run(paragraph.add_run(f"发文：{issue_unit.strip()}"), font=BODY_FONT, size=12, bold=False)
    count_text = str(print_count) if print_count is not None else "〔印数待确认〕"
    set_run(paragraph.add_run(f"\t共印{count_text}份"), font=BODY_FONT, size=12, bold=False)
    block.append(paragraph)
    set_paragraph_border(block[0], "top", size="8", space="2")
    set_paragraph_border(block[-1], "bottom", size="8", space="2")
    return block


def normalize_signature(paragraph, section=None, company_text=None):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ppr = paragraph._p.get_or_add_pPr()
    for element in list(ppr.findall(qn("w:ind"))):
        ppr.remove(element)
    for element in list(ppr.findall(qn("w:tabs"))):
        ppr.remove(element)
    set_spacing(paragraph, indent_pt=None)
    if section is not None and company_text:
        body_width = (section.page_width - section.left_margin - section.right_margin) / Pt(1)
        # A reconstruction starting point, not a claim about the source Word's
        # original indent. Actual glyph widths and long names need visual QA.
        estimated_width = sum(7 if c.isascii() else 14 for c in company_text.strip())
        if estimated_width + 6 < body_width:
            paragraph.paragraph_format.left_indent = Pt(body_width - 3 - estimated_width)
            paragraph.paragraph_format.right_indent = Pt(3)
            ind = paragraph._p.get_or_add_pPr().find(qn('w:ind'))
            ind.set(qn('w:start'), str(round((body_width - 3 - estimated_width) * 20)))
            ind.set(qn('w:end'), '60')
    for run in text_runs(paragraph):
        set_run(run, font=BODY_FONT, size=14, bold=False)


def add_page_number(section):
    footer = section.footer
    paragraph = footer.paragraphs[0]
    clear_paragraph(paragraph)
    use_public_style(paragraph)
    reset_indent(paragraph)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_single_spacing(paragraph)
    run = paragraph.add_run()
    set_run(run, font=LATIN_FONT, size=9, bold=False)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, placeholder, end])


def apply(
    input_path: Path,
    output_path: Path,
    *,
    title: str,
    issuing_number: str | None = None,
    title_continuations: list[str] | None = None,
    no_recipient: bool = False,
    body_bold: bool = False,
    appointment_notice: bool = False,
    no_title: bool = False,
    title_lines: list[str] | None = None,
    sign_type: str = "external",
    signer: str = "",
    copy_to: str | None = None,
    issue_unit: str | None = None,
    print_count: int | None = None,
    inspection_material: bool = False,
    blank_document_number: bool = False,
    blank_number_line_pt: float = DEFAULT_NUMBER_LINE_PT,
    signer_label_tracking_twips: int = SIGNER_LABEL_TRACKING_TWIPS,
    signer_right_indent_pt: float = SIGNER_RIGHT_INDENT_PT,
):
    if appointment_notice and not inspection_material:
        raise ValueError("任职通知原内容分支仅用于已确认的迎检资料。")
    if no_title and not appointment_notice:
        raise ValueError("无正文标题分支须同时声明任职类迎检通知。")
    if no_title and (title_lines or title_continuations):
        raise ValueError("原稿无正文标题时不得另行添加标题或标题续行。")
    if blank_document_number and (issuing_number or '').strip() not in ("", "〔发文字号待补〕"):
        raise ValueError("已提供正式文号，不能同时要求正文文号留空。")
    if blank_number_line_pt <= 0:
        raise ValueError("空文号段行距必须大于0。")
    if signer_label_tracking_twips < 0 or signer_right_indent_pt < 0:
        raise ValueError("签发栏字距和右缩进不得为负数。")
    doc = Document(str(input_path))
    if not issuing_number:
        existing_numbers = [] if blank_document_number else list(dict.fromkeys(
            re.sub(r'\s+', '', p.text) for p in doc.paragraphs
            if re.fullmatch(r'蒙力煤[字号发]*[〔（(]\d{4}[〕）)]\d+号', re.sub(r'\s+', '', p.text))
        ))
        if len(existing_numbers) > 1:
            raise ValueError('原稿存在多个不同文号，须明确传入--issuing-number，不自动选取。')
        issuing_number = existing_numbers[0] if existing_numbers else '〔发文字号待补〕'
    issuing_number = external_document_number(issuing_number, sign_type=sign_type, inspection_material=inspection_material)
    full_title = canonical_title(title + "".join(title_continuations or []))
    output_path = output_path.with_name(official_filename(full_title, issuing_number))
    if input_path.resolve() == output_path.resolve():
        raise ValueError("禁止覆盖源文件，请使用不同输出目录。")
    source_bolds = {p._p: bold_snapshot(p) for p in doc.paragraphs}
    public_style(doc)
    # Normal must supply inherited zero spacing for titles. It may be repaired
    # globally only when no protected table depends on it; otherwise fail safely.
    if not doc.tables:
        normal = doc.styles['Normal']._element.get_or_add_pPr()
        spacing = normal.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            normal.append(spacing)
        for key in ('before', 'after', 'beforeLines', 'afterLines', 'beforeAutospacing', 'afterAutospacing'):
            spacing.set(qn('w:' + key), '0')
    if not doc.paragraphs:
        raise ValueError("文档没有正文段落，无法安全套用红头。")

    original_paragraphs = list(doc.paragraphs)
    title_paragraph = next((p for p in original_paragraphs if p.text.strip() and not furniture(p.text)), None) if no_title else next((p for p in original_paragraphs if p.text.strip() == title), None)
    if title_paragraph is None:
        raise ValueError(f"未找到标题段落：{title}")
    title_paragraphs = [] if no_title else [title_paragraph]
    for continuation in title_continuations or []:
        paragraph = next((p for p in original_paragraphs if p.text.strip() == continuation), None)
        if paragraph is None:
            raise ValueError(f"未找到标题续行段落：{continuation}")
        title_paragraphs.append(paragraph)
    if title_lines and len(title_paragraphs) > 1:
        snapshots = [source_bolds[p._p] for p in title_paragraphs]
        source_bolds[title_paragraph._p] = (''.join(s[0] for s in snapshots), [b for s in snapshots for b in s[1]])

    # Replace only recognized leading furniture. Unknown leading content or
    # embedded objects must be handled explicitly, not silently destroyed.
    for paragraph in original_paragraphs:
        if paragraph is title_paragraph:
            break
        compact = re.sub(r"\s+", "", paragraph.text)
        known = compact in ("内蒙古准格尔旗", "力量煤业有限公司文件", "〔发文字号待补〕") or bool(
            re.fullmatch(r"蒙力煤[字号发]*[〔（(]\d{4}[〕）)]\d+号(?:签发人[:：].*)?", compact)
        )
        signer_line = bool(re.match(r"签发人[:：]", compact))
        if paragraph._p.xpath('.//w:drawing | .//w:pict | .//w:object | ./w:pPr/w:sectPr'):
            raise ValueError("标题前有图形/节设置，须在副本上人工识别旧红头后再套用。")
        if paragraph._p.xpath('.//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference'):
            raise ValueError("原红头包含批注，清理会破坏锚点；须在副本上明确处理后再套用。")
        if not compact or known or (inspection_material and signer_line):
            remove_paragraph(paragraph)
        else:
            raise ValueError(f"标题前有未处理内容，禁止重复套红头：{paragraph.text}")

    if inspection_material:
        # Only strip recognized trailing distribution furniture after a dated
        # signature; never scan-delete body paragraphs containing these words.
        remaining = list(doc.paragraphs)
        date_indices = [i for i, p in enumerate(remaining) if re.fullmatch(
            r"\d{4}年\d{1,2}月\d{1,2}日", re.sub(r"\s+", "", p.text)
        )]
        if date_indices:
            for paragraph in remaining[date_indices[-1] + 1:]:
                compact = re.sub(r"\s+", "", paragraph.text)
                if re.match(r"(?:主题|抄报|抄送|发文|印发)[:：]|共印\d+份", compact):
                    if paragraph._p.xpath('.//w:drawing | .//w:pict | .//w:object'):
                        raise ValueError("印发栏包含图形，需人工确认后删除。")
                    remove_paragraph(paragraph)
                elif not compact and paragraph._p.xpath('./w:pPr/w:pBdr'):
                    if paragraph._p.xpath('.//w:drawing | .//w:pict | .//w:object | ./w:pPr/w:sectPr'):
                        raise ValueError("印发栏横线包含其他对象，需人工确认。")
                    remove_paragraph(paragraph)
                elif compact:
                    break
        # Do not claim removal succeeded for unsupported table/textbox layouts.
        visible_text = "".join(doc._element.body.xpath('.//w:t/text()'))
        if re.search(r"签\s*发\s*人\s*[:：]|(?:主题|抄报|抄送|发文|印发)\s*[:：]|共印\s*\d+\s*份", visible_text):
            raise ValueError("仍检测到签发/印发栏（可能在表格或文本框中），需人工识别处理后重试。")
    trim_post_signature_blank_pages(doc)
    # Keep the same wrappers so title/recipient identity comparisons remain
    # valid after cleanup; doc.paragraphs would create new Paragraph wrappers.
    original_paragraphs = [p for p in original_paragraphs if p._p.getparent() is not None]

    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(23)
    section.bottom_margin = Mm(23)
    section.left_margin = Mm(23)
    section.right_margin = Mm(23)
    section.header_distance = Mm(0)
    section.footer_distance = Mm(15)
    add_page_number(section)

    # Add red-head paragraphs before the original title, in reverse insertion order.
    for header_text in ("内蒙古准格尔旗", "力量煤业有限公司文件"):
        header = add_before(title_paragraph, header_text, size=45, font=HEADER_FONT, bold=True, color=RED, line_spacing=65)
        distribute_red_header(header)
    code_text = "" if blank_document_number else (issuing_number or "〔发文字号待补〕")
    number_line_pt = blank_number_line_pt if blank_document_number else DEFAULT_NUMBER_LINE_PT
    number_paragraph = add_before(title_paragraph, code_text, size=14, font=HEADER_FONT, bold=False, color=None, border=True, line_spacing=number_line_pt)
    format_issuing_number(number_paragraph, section)
    if sign_type == "internal" and not inspection_material:
        if not copy_to or not issue_unit:
            raise ValueError("内发文必须提供 copy_to（抄送）和 issue_unit（发文单位）。")
        remove_paragraph_border(number_paragraph, "bottom")
        add_internal_signatory(title_paragraph, section, signer,
                               label_tracking_twips=signer_label_tracking_twips,
                               right_indent_pt=signer_right_indent_pt)
        add_internal_red_rule(title_paragraph, section)
    else:
        add_before(title_paragraph, "", size=10, font=BODY_FONT, bold=False, line_spacing=10)

    # Title: 黑体二号, normal weight, centered; inherit template 0-line
    # before/after spacing so WPS shows “0 行”.
    # The approved 24-number template uses Normal (正文文本) rather than the
    # source document's Heading 3 style. Heading 3 makes WPS show 自动 here.
    for paragraph in title_paragraphs:
        paragraph.style = doc.styles["Normal"]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_single_spacing(paragraph)
        clear_title_before_after_spacing(paragraph)
        enable_title_auto_spacing(paragraph)
        reset_indent(paragraph)
        if spacing_issues(paragraph, single=True):
            raise ValueError("标题Normal样式继承的间距不符合0行/单倍要求；不得全局改样式而影响业务表格，请在参考副本中校准。")
        for run in text_runs(paragraph):
            set_run(run, font=TITLE_FONT, size=22, bold=False)

    recipient = None if no_recipient else next(
        (p for p in original_paragraphs if p not in title_paragraphs and p.text.strip()),
        None,
    )
    if appointment_notice and recipient is not None and not recipient.text.strip().endswith(('：', ':')):
        recipient = None
    company_text = "内蒙古准格尔旗力量煤业有限公司"
    signature_start = next((idx for idx, p in enumerate(original_paragraphs)
                            if p.text.strip() == company_text), None)
    signature_paragraphs = set()
    if signature_start is not None:
        for p in original_paragraphs[signature_start:]:
            value = re.sub(r"\s+", "", p.text)
            if not value:
                continue
            if value in (company_text, "大饭铺煤矿") or re.fullmatch(r"\d{4}年\d{1,2}月\d{1,2}日", value):
                signature_paragraphs.add(p)
            else:
                break
    for paragraph in original_paragraphs:
        text = paragraph.text.strip()
        if paragraph in title_paragraphs:
            continue
        use_public_style(paragraph)
        if paragraph in signature_paragraphs:
            text = re.sub(r"\s+", "", text)
            date = re.fullmatch(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
            if date:
                text = f"{date[1]}年{int(date[2])}月{int(date[3])}日"
            if paragraph.text != text:
                replace_text(paragraph, text)
            normalize_signature(paragraph, section, company_text)
            paragraph.paragraph_format.keep_with_next = not bool(re.fullmatch(r"\d{4}年\d{1,2}月\d{1,2}日", text))
            continue
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        indent = None if paragraph is recipient else 28
        set_spacing(paragraph, indent_pt=indent)
        is_heading = independent_heading(text) and not appointment_notice
        paragraph.paragraph_format.keep_with_next = is_heading
        for run in text_runs(paragraph):
            set_run(run, font=BODY_FONT, size=14, bold=is_heading or body_bold, color="000000")

    # Main recipient is top-aligned; keep its text intact and remove accidental leading spaces.
    if recipient is not None:
        text = recipient.text.strip()
        if text != recipient.text:
            replace_text(recipient, text)
        recipient.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_spacing(recipient, indent_pt=None)
        for run in recipient.runs:
            set_run(run, font=BODY_FONT, size=14, bold=True)

    if title_lines:
        for continuation in title_paragraphs[1:]:
            if continuation._p.xpath('.//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference'):
                raise ValueError("标题续行含批注，须先明确保留范围，不能直接删除该段。")
            remove_paragraph(continuation)
        replace_text(title_paragraph, full_title)
        set_title_lines(title_paragraph, title_lines)
    normalize_blank_separators(doc, original_paragraphs[signature_start] if signature_start is not None else None)
    if sign_type == "internal" and not inspection_material:
        add_internal_distribution_block(
            doc,
            title=full_title,
            copy_to=copy_to,
            issue_unit=issue_unit,
            print_count=print_count,
        )
    trim_post_signature_blank_pages(doc)
    for paragraph in doc.paragraphs:
        if paragraph._p in source_bolds:
            restore_source_bold(paragraph, source_bolds[paragraph._p])
    # Newline formatting in WPS also reads the paragraph marker's rPr.
    for paragraph in doc.paragraphs:
        first = next((r for r in paragraph.runs if r.text.strip()), None)
        if first is not None and first._r.rPr is not None:
            ppr = paragraph._p.get_or_add_pPr()
            old = ppr.find(qn('w:rPr'))
            if old is not None:
                ppr.remove(old)
            ppr.append(deepcopy(first._r.rPr))
    doc.core_properties.title = full_title
    doc.core_properties.subject = "套红头公文"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def main():
    parser = argparse.ArgumentParser(description="为现有 DOCX 加套红头并按固定规则排版")
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--title", required=True)
    parser.add_argument("--title-continuation", action="append", default=[], help="标题的续行文字，可重复传入")
    parser.add_argument("--title-line", action="append", help="输出标题每行文字，按上长下短设置，可重复传入；必须渲染复核")
    parser.add_argument("--no-recipient", action="store_true", help="原稿无主送单位时使用，避免把首段正文误当主送单位")
    parser.add_argument("--body-bold", action="store_true", help="仅用户明确要求普通正文加粗时使用；一级独立小标题默认加粗")
    parser.add_argument("--appointment-notice", action="store_true", help="任职类迎检通知：按原内容排版，不要求一、序号或新标题")
    parser.add_argument("--no-title", action="store_true", help="该任职类迎检原稿无正文标题；--title仅用于文件名，不插入正文")
    parser.add_argument("--issuing-number", help="已确认文号优先；省略时读取原稿唯一文号，无文号时标待补，多文号时停止")
    parser.add_argument("--blank-document-number", action="store_true", help="正文文号行按用户要求留空；文件名仍标注发文字号待补")
    parser.add_argument("--blank-number-line-pt", type=float, default=DEFAULT_NUMBER_LINE_PT,
                        help="空文号段行距，默认27磅；本次2023内发样本使用19磅")
    parser.add_argument("--signer-label-tracking-twips", type=int, default=SIGNER_LABEL_TRACKING_TWIPS,
                        help="按同类型正式样本指定签发栏‘签发’字距，默认140")
    parser.add_argument("--signer-right-indent-pt", type=float, default=SIGNER_RIGHT_INDENT_PT,
                        help="按同类型正式样本指定签发栏右缩进，默认68磅")
    parser.add_argument("--sign-type", choices=("external", "internal"), default="external")
    parser.add_argument("--inspection-material", action="store_true", help="迎检内发文：不设签发人及末尾印发栏，不要求抄送/印数")
    parser.add_argument("--signer", default="", help="内发文签发人姓名；未知时保留空白签发栏")
    parser.add_argument("--copy-to", help="内发文印发栏中的抄送范围")
    parser.add_argument("--issue-unit", help="内发文印发栏中的发文单位")
    parser.add_argument("--print-count", type=int, help="内发文印发份数；未提供时保留〔印数待确认〕")
    args = parser.parse_args()
    if args.input.resolve() == args.output.resolve():
        raise SystemExit("禁止覆盖源文件，请指定新的输出路径。")
    output = apply(
        args.input,
        args.output,
        title=args.title,
        issuing_number=args.issuing_number,
        blank_document_number=args.blank_document_number,
        blank_number_line_pt=args.blank_number_line_pt,
        signer_label_tracking_twips=args.signer_label_tracking_twips,
        signer_right_indent_pt=args.signer_right_indent_pt,
        title_continuations=args.title_continuation,
        no_recipient=args.no_recipient,
        body_bold=args.body_bold,
        appointment_notice=args.appointment_notice,
        no_title=args.no_title,
        title_lines=args.title_line,
        sign_type=args.sign_type,
        signer=args.signer,
        copy_to=args.copy_to,
        issue_unit=args.issue_unit,
        print_count=args.print_count,
        inspection_material=args.inspection_material,
    )
    print(output)


if __name__ == "__main__":
    main()
