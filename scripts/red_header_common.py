"""Shared paragraph classification and effective OOXML property readers."""
from __future__ import annotations

import re
from difflib import SequenceMatcher
from docx.oxml.ns import qn
from docx.text.run import Run

CHINESE_NUMBER = r"[一二三四五六七八九十百]+"
CLOSINGS = ("特此请示", "特此报告", "特此通知", "特此函复", "特此复函", "妥否，请批示", "妥否请批示")


def external_document_number(value, *, sign_type, inspection_material=False):
    """Apply the external prefix without inventing a year or serial number."""
    if sign_type != 'external' or inspection_material or not value or value == '〔发文字号待补〕':
        return value
    compact = re.sub(r'\s+', '', value)
    match = re.fullmatch(r'蒙力煤[字号发]*([〔（(]\d{4}[〕）)]\d+号)', compact)
    if not match:
        raise ValueError('外发文号须为蒙力煤发〔年度〕序号号；无法识别年度和序号时须明确，不自动编造。')
    return '蒙力煤发' + match[1]


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


def effective_bold(run):
    """Bold is a toggle in styles and an absolute value in direct formatting."""
    state = False
    defaults = run.part.package.main_document_part.styles.element.xpath('./w:docDefaults/w:rPrDefault/w:rPr/w:b')
    if defaults:
        state = on(defaults[0].get(qn('w:val'), '1'))
    chains = (list(style_chain(run._parent.style)), list(style_chain(run.style)))
    for chain in chains:
        for style in reversed(chain):
            node = style.find(qn('w:rPr'))
            node = node.find(qn('w:b')) if node is not None else None
            if node is not None and on(node.get(qn('w:val'), '1')):
                state = not state
    if run._r.rPr is not None:
        direct = run._r.rPr.find(qn('w:b'))
        if direct is not None:
            state = on(direct.get(qn('w:val'), '1'))
    return state


def text_runs(paragraph):
    return [Run(node, paragraph) for node in paragraph._p.xpath('./w:r | ./w:hyperlink/w:r')]


def bold_snapshot(paragraph):
    runs = text_runs(paragraph)
    return (''.join(r.text for r in runs), [effective_bold(r) for r in runs for _ in r.text])


def compact_snapshot(snapshot):
    text, mask = snapshot
    positions = [i for i, char in enumerate(text) if not char.isspace()]
    value = ''.join(text[i] for i in positions)
    date_match = re.fullmatch(r'\d{4}年(\d{1,2})月(\d{1,2})日', value)
    if date_match:
        drop = {date_match.start(group) for group in (1, 2) if date_match[group].startswith('0')}
        positions = [p for i, p in enumerate(positions) if i not in drop]
        value = ''.join(text[i] for i in positions)
    return value, [mask[i] for i in positions], positions


def mapped_bold(snapshot, target):
    old, weights, _ = compact_snapshot(snapshot)
    new, _, positions = compact_snapshot((target, [False] * len(target)))
    result = [False] * len(target)
    for block in SequenceMatcher(None, old, new, autojunk=False).get_matching_blocks():
        for i in range(block.size):
            result[positions[block.b + i]] = weights[block.a + i]
    return result


def furniture(text):
    text = re.sub(r'\s+', '', text)
    return text in ('内蒙古准格尔旗', '力量煤业有限公司文件', '〔发文字号待补〕') or bool(re.fullmatch(
        r'蒙力煤[字号发]*[〔（(]\d{4}[〕）)]\d+号(?:签发人[:：].*)?', text)) or bool(re.match(
        r'签发人[:：]|(?:主题|抄报|抄送|发文|印发)[:：]', text))


def source_bold_map(source, output):
    """Match source character ranges across reflow and normalized signature dates."""
    def stream(doc):
        text, weights, locations = '', [], []
        for p in doc.paragraphs:
            if furniture(p.text):
                continue
            value, bold, positions = compact_snapshot(bold_snapshot(p))
            text += value
            weights.extend(bold)
            locations.extend((p._p, i) for i in positions)
        return text, weights, locations
    old, old_weights, _ = stream(source)
    new, _, locations = stream(output)
    maps = {p._p: [False] * len(p.text) for p in output.paragraphs}
    matched, dropped = set(), []
    for block in SequenceMatcher(None, old, new, autojunk=False).get_matching_blocks():
        for i in range(block.size):
            matched.add(block.a + i)
            if old_weights[block.a + i]:
                paragraph, position = locations[block.b + i]
                maps[paragraph][position] = True
    for i, weight in enumerate(old_weights):
        if weight and i not in matched:
            dropped.append({'source_character': i, 'text': old[max(0, i - 8):i + 9]})
    return maps, dropped, old == new


def tail_layout(doc):
    """Inspect the blank suffix after the last real paragraph, including attachments."""
    from docx.text.paragraph import Paragraph
    nodes = list(doc._element.body)
    if nodes and nodes[-1].tag == qn('w:sectPr'):
        nodes.pop()
    blanks = []
    while nodes and nodes[-1].tag == qn('w:p'):
        p = Paragraph(nodes[-1], doc._body)
        if p.text.strip() or p._p.xpath('.//w:drawing | .//w:pict | .//w:object | .//w:fldChar | .//w:instrText | '
                                       './/w:bookmarkStart | .//w:bookmarkEnd | .//w:commentRangeStart | '
                                       './/w:commentRangeEnd | .//w:commentReference | .//w:sdt | .//w:ins | .//w:del'):
            break
        blanks.insert(0, nodes.pop())
    anchor = Paragraph(nodes[-1], doc._body) if nodes and nodes[-1].tag == qn('w:p') else None
    value = re.sub(r'\s+', '', anchor.text) if anchor else ''
    applicable = bool(anchor and (value or anchor._p.xpath('.//w:drawing | .//w:pict | .//w:object')))
    breaks = []
    if applicable:
        for child in reversed(list(anchor._p)):
            if child.tag != qn('w:r'):
                break
            for node in reversed(list(child)):
                if node.tag == qn('w:rPr'):
                    continue
                if node.tag == qn('w:t') and not (node.text or '').strip():
                    continue
                if node.tag in (qn('w:br'), qn('w:cr'), qn('w:lastRenderedPageBreak')):
                    breaks.append(node)
                    continue
                return anchor, blanks, breaks, applicable
    return anchor, blanks, breaks, applicable


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


def recipient_layout_issues(paragraph):
    issues = indent_issues(paragraph, chars=0)
    attrs = effective_ppr(paragraph, 'ind')
    for key in ('firstLine', 'firstLineChars', 'hanging', 'hangingChars'):
        if float(attrs.get(key, '0')) != 0:
            issues[key] = attrs[key]
    alignment = effective_ppr(paragraph, 'jc').get('val', 'left')
    if alignment not in ('left', 'start'):
        issues['alignment'] = alignment
    if re.search(r'(^|[\r\n])[^\S\r\n]+', paragraph.text):
        issues['manual_leading_whitespace'] = True
    return issues


def rendered_page_content(folder, section):
    """Check all numbered renderer PNGs, excluding page margins/header/footer."""
    from PIL import Image, ImageOps
    pages = sorted(((int(m[1]), p) for p in folder.glob('page-*.png')
                    if (m := re.fullmatch(r'page-(\d+)\.png', p.name))), key=lambda item: item[0])
    if not pages or [n for n, _ in pages] != list(range(1, len(pages) + 1)):
        raise ValueError('须提供完整且连续的page-1.png至末页渲染图，不能用空目录或缺页截图验收。')
    results = []
    for number, path in pages:
        with Image.open(path) as img:
            width, height = img.size
            box = (round(width * section.left_margin / section.page_width),
                   round(height * section.top_margin / section.page_height),
                   round(width * (1 - section.right_margin / section.page_width)),
                   round(height * (1 - section.bottom_margin / section.page_height)))
            if box[0] >= box[2] or box[1] >= box[3]:
                raise ValueError('页面或页边距无有效正文区域，须检查实际渲染。')
            histogram = ImageOps.grayscale(img.convert('RGB').crop(box)).histogram()
            ink = sum(histogram[:220])
            results.append({'page': number, 'body_ink_pixels': ink, 'blank_body': ink < 12})
    return results


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
    actual_bold = effective_bold(run)
    if bold is not None and actual_bold != bold:
        issues['bold'] = actual_bold
    cs_size = effective_run(run, 'szCs').get('val')
    cs_bold = effective_run(run, 'bCs').get('val')
    if cs_size is not None and abs(float(cs_size) / 2 - size) > 0.05:
        issues['complex_size'] = cs_size
    if cs_bold is not None and on(cs_bold) != (actual_bold if bold is None else bold):
        issues['complex_bold'] = cs_bold
    if color and effective_run(run, 'color').get('val') != color:
        issues['color'] = effective_run(run, 'color').get('val')
    return issues
