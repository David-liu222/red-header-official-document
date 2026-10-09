"""Shared paragraph classification and effective OOXML property readers."""
from __future__ import annotations

import re
from docx.oxml.ns import qn

CHINESE_NUMBER = r"[一二三四五六七八九十百]+"
CLOSINGS = ("特此请示", "特此报告", "特此通知", "特此函复", "特此复函", "妥否，请批示", "妥否请批示")


def independent_heading(text):
    text = text.strip()
    match = re.match(rf"^(?:{CHINESE_NUMBER}、|第{CHINESE_NUMBER}章)\s*(.+)$", text)
    return bool(match and len(match[1]) <= 50 and not re.search(r"[，。；：！？,.;:!?]", match[1]))


def numbered_level(text):
    text = text.lstrip()
    patterns = ((0, rf"^第{CHINESE_NUMBER}章"), (1, rf"^{CHINESE_NUMBER}、"),
                (2, rf"^[（(]{CHINESE_NUMBER}[）)]"), (3, r"^\d+[.．](?!\d)"),
                (4, r"^[（(]\d+[）)]"), (5, r"^[①②③④⑤⑥⑦⑧⑨⑩]"))
    return next((level for level, pattern in patterns if re.match(pattern, text)), None)


def closing(text):
    return text.strip().rstrip("。.;；！!？?") in CLOSINGS


def attachment_start(text):
    return bool(re.match(r"^附件\s*[:：]", text.strip()))


def safe_blank(paragraph):
    if paragraph.text.strip():
        return False
    return not paragraph._p.xpath(
        './/w:drawing | .//w:pict | .//w:object | .//w:br | .//w:fldChar | '
        './/w:instrText | .//w:bookmarkStart | .//w:bookmarkEnd | '
        './/w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference | '
        './w:pPr/w:sectPr | ./w:pPr/w:pageBreakBefore[not(@w:val="0" or @w:val="false")]')


def style_chain(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style._element
        style = style.base_style


def ppr_layers(paragraph):
    if paragraph._p.pPr is not None:
        yield paragraph._p.pPr
    for style in style_chain(paragraph.style):
        ppr = style.find(qn('w:pPr'))
        if ppr is not None:
            yield ppr
    yield from paragraph.part.package.main_document_part.styles.element.xpath('./w:docDefaults/w:pPrDefault/w:pPr')


def effective_ppr(paragraph, tag):
    attrs = {}
    for layer in ppr_layers(paragraph):
        node = layer.find(qn('w:' + tag))
        if node is not None:
            for key, value in node.attrib.items():
                attrs.setdefault(key.rsplit('}', 1)[-1], value)
            if not node.attrib:
                attrs.setdefault('val', '1')
    return attrs


def run_layers(run):
    if run._r.rPr is not None:
        yield run._r.rPr
    if run.style is not None:
        for style in style_chain(run.style):
            rpr = style.find(qn('w:rPr'))
            if rpr is not None:
                yield rpr
    for style in style_chain(run._parent.style):
        rpr = style.find(qn('w:rPr'))
        if rpr is not None:
            yield rpr
    yield from run.part.package.main_document_part.styles.element.xpath('./w:docDefaults/w:rPrDefault/w:rPr')


def effective_run(run, tag):
    attrs = {}
    for layer in run_layers(run):
        node = layer.find(qn('w:' + tag))
        if node is not None:
            if tag == 'rFonts':
                # A higher-level literal font masks a lower-level theme choice
                # for the same script, rather than inheriting both simultaneously.
                for literal, theme in (('ascii', 'asciiTheme'), ('hAnsi', 'hAnsiTheme'),
                                       ('eastAsia', 'eastAsiaTheme'), ('cs', 'cstheme')):
                    if literal in attrs or theme in attrs:
                        continue
                    for key in (literal, theme):
                        value = node.get(qn('w:' + key))
                        if value is not None:
                            attrs[key] = value
                continue
            for key, value in node.attrib.items():
                attrs.setdefault(key.rsplit('}', 1)[-1], value)
            if not node.attrib:
                attrs.setdefault('val', '1')
    return attrs


def on(value):
    return value not in (None, '0', 'false', 'off')


def spacing_issues(paragraph, *, line_pt=None, single=False):
    attrs = effective_ppr(paragraph, 'spacing')
    issues = {}
    for key in ('before', 'after', 'beforeLines', 'afterLines'):
        if float(attrs.get(key, '0')) != 0:
            issues[key] = attrs[key]
    for key in ('beforeAutospacing', 'afterAutospacing'):
        if on(attrs.get(key)):
            issues[key] = attrs[key]
    if single and (attrs.get('lineRule', 'auto') != 'auto' or int(attrs.get('line', '240')) != 240):
        issues['line'] = attrs
    if line_pt is not None and (attrs.get('lineRule') != 'exact' or
                               abs(float(attrs.get('line', '0')) / 20 - line_pt) > 0.05):
        issues['line'] = attrs
    if line_pt is not None and on(effective_ppr(paragraph, 'snapToGrid').get('val')):
        issues['snapToGrid'] = 'enabled: actual leading requires native review'
    return issues


def indent_issues(paragraph, *, chars):
    attrs = effective_ppr(paragraph, 'ind')
    issues = {}
    for key in ('left', 'right', 'start', 'end', 'leftChars', 'rightChars', 'startChars', 'endChars', 'hanging', 'hangingChars'):
        if float(attrs.get(key, '0')) != 0:
            issues[key] = attrs[key]
    actual = float(attrs['firstLineChars']) / 100 if 'firstLineChars' in attrs else float(attrs.get('firstLine', '0')) / 280
    if abs(actual - chars) > 0.05:
        issues['first_line_chars'] = actual
    return issues


def run_format_issues(run, *, font, size, bold, color=None):
    issues = {}
    fonts = effective_run(run, 'rFonts')
    actual_font = fonts.get('eastAsia') or fonts.get('hAnsi') or fonts.get('ascii')
    aliases = {'宋体': ('宋体', 'SimSun'), '黑体': ('黑体', 'SimHei'), '仿宋': ('仿宋', 'FangSong')}
    if actual_font not in aliases.get(font, (font,)):
        issues['font'] = actual_font
    if any(key.lower().endswith('theme') for key in fonts):
        issues['theme_font'] = fonts
    actual_size = float(effective_run(run, 'sz').get('val', '0')) / 2
    if abs(actual_size - size) > 0.05:
        issues['size_pt'] = actual_size
    actual_bold = on(effective_run(run, 'b').get('val'))
    if actual_bold != bold:
        issues['bold'] = actual_bold
    cs_size = effective_run(run, 'szCs').get('val')
    cs_bold = effective_run(run, 'bCs').get('val')
    if cs_size is not None and abs(float(cs_size) / 2 - size) > 0.05:
        issues['complex_size'] = cs_size
    if cs_bold is not None and on(cs_bold) != bold:
        issues['complex_bold'] = cs_bold
    if color and effective_run(run, 'color').get('val') != color:
        issues['color'] = effective_run(run, 'color').get('val')
    return issues
