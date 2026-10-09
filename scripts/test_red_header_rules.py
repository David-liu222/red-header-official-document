#!/usr/bin/env python3
"""Synthetic regression checks; never load business documents or employee data."""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from apply_red_header_docx import apply, set_paragraph_border
from extract_review_comments import extract
from validate_outputs import visible_files


TITLE = "关于测试事项的通知"
NUMBER = "蒙力煤字〔2026〕1号"
COMPANY = "内蒙古准格尔旗力量煤业有限公司"


class HeaderRulesTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def source(self, decorated=False):
        doc = Document()
        if decorated:
            for text in ("内蒙古准格尔旗", "力量煤业有限公司文件", NUMBER, "签 发 人："):
                doc.add_paragraph(text)
        doc.add_paragraph(TITLE)
        doc.add_paragraph("公司各部门：")
        doc.add_paragraph("一、按已确认要求办理。")
        doc.add_paragraph(COMPANY)
        date = doc.add_paragraph("2026 年 09 月 01 日")
        date.paragraph_format.left_indent = Pt(20)
        if decorated:
            first = doc.add_paragraph("主题：" + TITLE)
            set_paragraph_border(first, "top")
            doc.add_paragraph("抄报：测试范围")
            last = doc.add_paragraph("发文：测试单位\t共印1份")
            set_paragraph_border(last, "bottom")
        path = self.root / "source.docx"
        doc.save(path)
        return path

    def generate(self, decorated=False, **kwargs):
        source = self.source(decorated)
        original = source.read_bytes()
        issuing_number = kwargs.pop("issuing_number", NUMBER)
        output = apply(source, self.root / "output" / "out.docx", title=TITLE,
                       issuing_number=issuing_number, **kwargs)
        self.assertEqual(source.read_bytes(), original)
        return output, Document(output)

    def test_inspection_removes_existing_furniture(self):
        output, doc = self.generate(decorated=True, sign_type="internal", inspection_material=True)
        texts = [p.text for p in doc.paragraphs]
        self.assertEqual(texts.count("内蒙古准格尔旗"), 1)
        self.assertEqual(texts.count(NUMBER), 1)
        self.assertFalse(any(any(word in t for word in ("签 发 人", "主题：", "抄报：", "共印")) for t in texts))
        self.assertIn("一、按已确认要求办理。", texts)
        self.assertFalse(doc._element.body.xpath('.//w:pBdr/*[@w:color="000000"]'))
        self.check_alignment(doc)
        result = subprocess.run([sys.executable, str(Path(__file__).with_name("validate_red_header_docx.py")),
                                 str(output), "--title", TITLE, "--sign-type", "internal",
                                 "--inspection-material", "--json"], check=True, capture_output=True, text=True)
        checks = json.loads(result.stdout)["checks"]
        self.assertFalse([c for c in checks if c["severity"] == "error"], checks)

    def check_alignment(self, doc):
        for p in doc.paragraphs[:2]:
            self.assertEqual(p.alignment, WD_ALIGN_PARAGRAPH.DISTRIBUTE)
            self.assertEqual(p.paragraph_format.line_spacing, Pt(65))
            self.assertFalse(p._p.xpath('./w:pPr/w:ind'))
            for r in p.runs:
                self.assertEqual(r.font.name, "宋体")
                self.assertEqual(r.font.size, Pt(45))
                self.assertIsNone(r._r.rPr.find(qn("w:spacing")))
        self.assertEqual(doc.paragraphs[0].text, "内蒙古准格尔旗")
        self.assertEqual(doc.paragraphs[1].text, "力量煤业有限公司文件")
        company = next(p for p in doc.paragraphs if p.text == COMPANY)
        date = next(p for p in doc.paragraphs if p.text == "2026年9月1日")
        self.assertEqual(company.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(date.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(company.paragraph_format.left_indent, date.paragraph_format.left_indent)
        self.assertEqual(company.paragraph_format.right_indent, date.paragraph_format.right_indent)

    def test_execution_internal_retains_furniture(self):
        _, doc = self.generate(sign_type="internal", copy_to="测试范围", issue_unit="测试单位", print_count=1)
        text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("签发人：", text)
        self.assertIn("抄送：测试范围", text)
        self.assertIn("共印1份", text)
        self.check_alignment(doc)

    def test_missing_print_count_stays_unconfirmed(self):
        _, doc = self.generate(sign_type="internal", copy_to="测试范围", issue_unit="测试单位")
        self.assertIn("共印〔印数待确认〕份", "\n".join(p.text for p in doc.paragraphs))

    def test_signer_sample_override_generates_and_validates(self):
        output, doc = self.generate(sign_type="internal", copy_to="测试范围", issue_unit="测试单位",
                                    signer_label_tracking_twips=100, signer_right_indent_pt=80)
        signer = next(p for p in doc.paragraphs if p.text.startswith("签发人："))
        self.assertEqual(signer.paragraph_format.right_indent, Pt(80))
        spacing = signer.runs[0]._r.rPr.find(qn("w:spacing"))
        self.assertEqual(int(spacing.get(qn("w:val"))), 100)
        result = subprocess.run([sys.executable, str(Path(__file__).with_name("validate_red_header_docx.py")),
                                 str(output), "--title", TITLE, "--sign-type", "internal",
                                 "--signer-label-tracking-twips", "100", "--signer-right-indent-pt", "80", "--json"],
                                check=True, capture_output=True, text=True)
        checks = json.loads(result.stdout)["checks"]
        self.assertTrue(any(c["rule"] == "内发文签发栏及红线位置按同类样本设置"
                            and c["status"] == "pass" for c in checks))

    def test_inspection_needs_no_distribution_fields(self):
        self.generate(sign_type="internal", inspection_material=True)

    def test_execution_internal_still_requires_distribution_fields(self):
        with self.assertRaisesRegex(ValueError, "必须提供"):
            self.generate(sign_type="internal")

    def test_validator_rejects_prior_center_plus_tracking(self):
        output, doc = self.generate(inspection_material=True)
        doc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.paragraphs[1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:val"), "450")
        doc.paragraphs[0].runs[0]._r.rPr.append(spacing)
        doc.save(output)
        result = subprocess.run([sys.executable, str(Path(__file__).with_name("validate_red_header_docx.py")),
                                 str(output), "--title", TITLE, "--inspection-material", "--json"],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 1)
        self.assertGreater(json.loads(result.stdout)['summary']['failed'], 0)
        checks = json.loads(result.stdout)["checks"]
        self.assertTrue(any(c["rule"] == "红头两行按定稿分散对齐" and c["severity"] == "error" for c in checks))
        self.assertTrue(any(c["rule"] == "红头原文完整且无手工字距和缩进"
                            and c["severity"] == "error" for c in checks))

    def test_validator_rejects_manual_header_tracking(self):
        output, doc = self.generate(sign_type="internal", inspection_material=True)
        first_run = doc.paragraphs[0].runs[0]
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:val"), "450")
        first_run._r.rPr.append(spacing)
        doc.save(output)
        result = subprocess.run([sys.executable, str(Path(__file__).with_name("validate_red_header_docx.py")),
                                 str(output), "--title", TITLE, "--sign-type", "internal", "--inspection-material", "--json"],
                                 capture_output=True, text=True)
        self.assertEqual(result.returncode, 1)
        checks = json.loads(result.stdout)["checks"]
        self.assertTrue(any(c["rule"] == "红头原文完整且无手工字距和缩进"
                            and c["severity"] == "error" for c in checks))

    def test_external_uses_approved_distributed_header(self):
        output, doc = self.generate(sign_type="external")
        self.check_alignment(doc)
        result = subprocess.run([sys.executable, str(Path(__file__).with_name("validate_red_header_docx.py")),
                                 str(output), "--title", TITLE, "--sign-type", "external", "--json"],
                                check=True, capture_output=True, text=True)
        checks = json.loads(result.stdout)["checks"]
        self.assertTrue(any(c["rule"] == "红头原文完整且无手工字距和缩进"
                            and c["status"] == "pass" for c in checks))

    def test_requested_blank_number_line_is_empty(self):
        output, doc = self.generate(sign_type="internal", issuing_number="",
                                    blank_document_number=True, blank_number_line_pt=19,
                                    copy_to="测试范围", issue_unit="测试单位")
        self.assertTrue(output.name.startswith("〔发文字号待补〕："))
        self.assertEqual(doc.paragraphs[2].text, "")
        self.assertEqual(doc.paragraphs[2].paragraph_format.line_spacing, Pt(19))
        result = subprocess.run([sys.executable, str(Path(__file__).with_name("validate_red_header_docx.py")),
                                 str(output), "--title", TITLE, "--sign-type", "internal",
                                 "--blank-document-number", "--blank-number-line-pt", "19", "--json"],
                                check=True, capture_output=True, text=True)
        checks = json.loads(result.stdout)["checks"]
        self.assertTrue(any(c["rule"] == "文号正文按要求留空且保留设定行距"
                            and c["status"] == "pass" for c in checks))
        self.assertFalse(any(c["rule"] == "文号宋体中文、Times New Roman数字、四号不加粗"
                             for c in checks))

    def test_blank_number_default_keeps_standard_line_height(self):
        _, doc = self.generate(sign_type="external", issuing_number="", blank_document_number=True)
        self.assertEqual(doc.paragraphs[2].text, "")
        self.assertEqual(doc.paragraphs[2].paragraph_format.line_spacing, Pt(27))

    def custom_source(self, body=None, no_recipient=False, table=False):
        doc = Document()
        doc.styles['Normal'].paragraph_format.space_before = Pt(0)
        doc.styles['Normal'].paragraph_format.space_after = Pt(0)
        doc.add_paragraph(TITLE)
        if not no_recipient:
            doc.add_paragraph('公司各部门：')
        for text in body or ['一、工作要求', '（一）按既有要求办理。']:
            doc.add_paragraph(text)
        if table:
            tbl = doc.add_table(rows=2, cols=2)
            tbl.cell(0, 0).text = '项目'
            tbl.cell(0, 1).text = '说明'
            tbl.cell(1, 0).text = '测试'
            tbl.cell(1, 1).text = '保留原表内容'
            borders = OxmlElement('w:tblBorders')
            for name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                border = OxmlElement('w:' + name)
                border.set(qn('w:val'), 'single')
                borders.append(border)
            tbl._tbl.tblPr.append(borders)
        doc.add_paragraph(COMPANY)
        doc.add_paragraph('2026年9月1日')
        path = self.root / 'custom.docx'
        doc.save(path)
        return path

    def apply_custom(self, source, **kwargs):
        return apply(source, self.root / 'custom-output/out.docx', title=TITLE, issuing_number=NUMBER, **kwargs)

    def validate(self, output, *extra):
        result = subprocess.run([sys.executable, str(Path(__file__).with_name('validate_red_header_docx.py')),
                                 str(output), '--title', TITLE, '--json', *extra], capture_output=True, text=True)
        self.assertFalse(result.stderr, result.stderr)
        return result.returncode, json.loads(result.stdout)

    def failures(self, report):
        return [c for c in report['checks'] if c['status'] == 'fail']

    def test_heading_and_narrative_have_distinct_weights(self):
        output = self.apply_custom(self.custom_source())
        doc = Document(output)
        heading = next(p for p in doc.paragraphs if p.text == '一、工作要求')
        body = next(p for p in doc.paragraphs if p.text.startswith('（一）'))
        self.assertTrue(all(r.bold for r in heading.runs if r.text.strip()))
        self.assertTrue(all(r.bold is False for r in body.runs if r.text.strip()))
        code, report = self.validate(output)
        self.assertEqual(code, 0, self.failures(report))
        heading.runs[0].bold = False
        doc.save(output)
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any(c['rule'] == '一级独立小标题仿宋四号且加粗' for c in self.failures(report)))

    def test_no_recipient_does_not_reclassify_first_body(self):
        output = self.apply_custom(self.custom_source(no_recipient=True), no_recipient=True)
        code, report = self.validate(output, '--no-recipient')
        self.assertEqual(code, 0, self.failures(report))
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any(c['rule'] == '主送单位可明确识别' for c in self.failures(report)))

    def test_hierarchy_skip_is_a_hard_failure(self):
        output = self.apply_custom(self.custom_source(['一、工作要求', '（一）具体如下：', '（1）登记完成情况。']))
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any(c['rule'] == '显式序号层级无跳级' for c in self.failures(report)))
        self.assertGreater(report['summary']['failed'], 0)
        self.assertFalse(report['summary']['automatic_ok'])

    def test_valid_hierarchy_and_intro_colon_are_accepted(self):
        output = self.apply_custom(self.custom_source(['一、工作要求', '（一）具体如下：', '1.按既有要求办理。', '（1）登记完成情况。']))
        code, report = self.validate(output)
        self.assertEqual(code, 0, self.failures(report))

    def test_book_title_separator_and_missing_period_are_detected(self):
        output = self.apply_custom(self.custom_source(['一、工作要求', '（一）对照《甲》、《乙》执行；']))
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        rules = [c['rule'] for c in self.failures(report)]
        self.assertIn('相邻书名号之间无顿号', rules)
        self.assertIn('完整序号条目用句号结束', rules)

    def test_closing_and_attachment_separators_are_normalized(self):
        output = self.apply_custom(self.custom_source(['一、工作要求', '（一）按既有要求办理。', '特此通知', '附件：1.测试清单', '2.测试说明']))
        code, report = self.validate(output)
        self.assertEqual(code, 0, self.failures(report))
        doc = Document(output)
        attachment = next(i for i, p in enumerate(doc.paragraphs) if p.text.startswith('附件'))
        self.assertEqual(doc.paragraphs[attachment - 1].text, '')
        doc.paragraphs[attachment - 1]._p.getparent().remove(doc.paragraphs[attachment - 1]._p)
        doc.save(output)
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any('恰好空一行' in c['rule'] for c in self.failures(report)))

    def test_closing_period_is_reported_without_rewriting_text(self):
        output = self.apply_custom(self.custom_source(['一、工作要求', '（一）按既有要求办理。', '特此通知。']))
        self.assertIn('特此通知。', [p.text for p in Document(output).paragraphs])
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any(c['rule'] == '已识别结束语不带结束标点' for c in self.failures(report)))

    def test_inherited_header_indent_is_detected(self):
        output = self.apply_custom(self.custom_source())
        doc = Document(output)
        ind = doc.styles['红头公文基准']._element.pPr.find(qn('w:ind'))
        ind.set(qn('w:firstLineChars'), '200')
        doc.save(output)
        self.assertFalse(doc.paragraphs[0]._p.xpath('./w:pPr/w:ind'))
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any('红头有效缩进' in c['rule'] for c in self.failures(report)))

    def test_character_unit_spacing_and_grid_cannot_pass_as_zero(self):
        output = self.apply_custom(self.custom_source())
        doc = Document(output)
        paragraph = next(p for p in doc.paragraphs if p.text.startswith('（一）'))
        paragraph._p.pPr.find(qn('w:spacing')).set(qn('w:beforeLines'), '50')
        paragraph._p.pPr.find(qn('w:snapToGrid')).set(qn('w:val'), '1')
        doc.save(output)
        self.assertEqual(paragraph.paragraph_format.space_before, Pt(0))
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any('正文段前段后' in c['rule'] for c in self.failures(report)))

    def test_title_inherited_spacing_is_checked(self):
        output = self.apply_custom(self.custom_source())
        doc = Document(output)
        doc.styles['Normal'].paragraph_format.space_after = Pt(8)
        doc.save(output)
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any('标题继承后的' in c['rule'] for c in self.failures(report)))

    def test_complex_script_weight_conflict_is_detected(self):
        output = self.apply_custom(self.custom_source())
        doc = Document(output)
        paragraph = next(p for p in doc.paragraphs if p.text.startswith('（一）'))
        paragraph.runs[0]._r.rPr.find(qn('w:bCs')).set(qn('w:val'), '1')
        doc.save(output)
        self.assertIs(paragraph.runs[0].bold, False)
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any(c['rule'] == '普通正文仿宋四号且不加粗' for c in self.failures(report)))

    def test_lowercase_complex_script_theme_is_detected(self):
        output = self.apply_custom(self.custom_source())
        doc = Document(output)
        body = next(p for p in doc.paragraphs if p.text.startswith('（一）'))
        body.runs[0]._r.get_or_add_rPr().rFonts.set(qn('w:cstheme'), 'minorBidi')
        doc.save(output)
        code, report = self.validate(output)
        self.assertEqual(code, 1)
        self.assertTrue(any('theme_font' in str(c) for c in self.failures(report)))

    def test_leading_zero_and_invalid_date_are_detected(self):
        output = self.apply_custom(self.custom_source())
        for value in ('2026年09月01日', '2026年2月30日'):
            doc = Document(output)
            p = next(p for p in doc.paragraphs if '年' in p.text and p.text.endswith('日'))
            p.runs[0].text = value
            doc.save(output)
            code, report = self.validate(output)
            self.assertEqual(code, 1)
            self.assertTrue(any('落款日期为有效' in c['rule'] for c in self.failures(report)))

    def test_company_and_mine_cannot_share_paragraph(self):
        source = self.custom_source()
        doc = Document(source)
        p = next(p for p in doc.paragraphs if p.text == COMPANY)
        p.add_run('大饭铺煤矿')
        doc.save(source)
        output = self.apply_custom(source)
        code, report = self.validate(output, '--expected-mine', '大饭铺煤矿')
        self.assertEqual(code, 1)
        self.assertTrue(any(c['rule'] == '公司与矿名在独立段落分行' for c in self.failures(report)))

    def test_table_xml_and_source_styles_are_preserved(self):
        source = self.custom_source(table=True)
        doc = Document(source)
        doc.styles['Normal'].paragraph_format.first_line_indent = Pt(28)
        doc.save(source)
        table_before = Document(source).tables[0]._tbl.xml
        normal_before = Document(source).styles['Normal']._element.xml
        source_before = source.read_bytes()
        output = self.apply_custom(source)
        self.assertEqual(source_before, source.read_bytes())
        generated = Document(output)
        self.assertEqual(generated.tables[0]._tbl.xml, table_before)
        self.assertEqual(generated.styles['Normal']._element.xml, normal_before)
        code, report = self.validate(output)
        self.assertEqual(code, 0, self.failures(report))

    def test_nonzero_normal_spacing_with_table_stops_without_output(self):
        source = self.custom_source(table=True)
        doc = Document(source)
        doc.styles['Normal'].paragraph_format.space_after = Pt(8)
        doc.save(source)
        before = source.read_bytes()
        with self.assertRaisesRegex(ValueError, 'Normal样式继承'):
            self.apply_custom(source)
        self.assertEqual(source.read_bytes(), before)
        self.assertFalse((self.root / 'custom-output').exists())

    def test_conflicting_public_style_is_not_rewritten(self):
        source = self.custom_source(table=True)
        doc = Document(source)
        from docx.enum.style import WD_STYLE_TYPE
        style = doc.styles.add_style('红头公文基准', WD_STYLE_TYPE.PARAGRAPH)
        style.paragraph_format.space_after = Pt(8)
        doc.tables[0].cell(0, 0).paragraphs[0].style = style
        doc.save(source)
        before = source.read_bytes()
        with self.assertRaisesRegex(ValueError, '不能全局改写同名样式'):
            self.apply_custom(source)
        self.assertEqual(source.read_bytes(), before)
        self.assertFalse((self.root / 'custom-output').exists())

    def test_reapplying_canonical_public_style_preserves_table(self):
        output = self.apply_custom(self.custom_source(table=True))
        doc = Document(output)
        table_before = doc.tables[0]._tbl.xml
        style_before = doc.styles['红头公文基准']._element.xml
        second = apply(output, self.root / 'second-output/out.docx', title=TITLE, issuing_number=NUMBER)
        generated = Document(second)
        self.assertEqual(generated.tables[0]._tbl.xml, table_before)
        self.assertEqual(generated.styles['红头公文基准']._element.xml, style_before)
        code, report = self.validate(second)
        self.assertEqual(code, 0, self.failures(report))

    def test_comments_keep_precise_anchors_during_date_formatting(self):
        source = self.custom_source()
        doc = Document(source)
        company = next(p for p in doc.paragraphs if p.text == COMPANY)
        date = doc.paragraphs[-1]
        date.runs[0].text = '2026 年 09 月 01 日'
        doc.add_comment(company.runs, text='落款前空一行', author='测试审阅人')
        doc.add_comment(date.runs, text='日期不加前导0', author='测试审阅人')
        doc.save(source)
        original = source.read_bytes()
        output = self.apply_custom(source)
        record = extract(output)
        self.assertEqual(record['comment_count'], 2)
        by_text = {c['text']: c for c in record['comments']}
        self.assertEqual(by_text['日期不加前导0']['anchors'][0]['selected_text'], '2026年9月1日')
        self.assertEqual(by_text['落款前空一行']['anchors'][0]['selected_text'], COMPANY)
        self.assertEqual(source.read_bytes(), original)

    def test_comment_on_old_header_prevents_silent_anchor_deletion(self):
        source = self.source(decorated=True)
        doc = Document(source)
        doc.add_comment(doc.paragraphs[0].runs, text='核对红头', author='测试审阅人')
        doc.save(source)
        before = source.read_bytes()
        with self.assertRaisesRegex(ValueError, '原红头包含批注'):
            self.apply_custom(source)
        self.assertEqual(source.read_bytes(), before)

    def test_bookmarked_empty_paragraph_is_preserved(self):
        source = self.custom_source()
        doc = Document(source)
        company = next(p for p in doc.paragraphs if p.text == COMPANY)
        blank = company.insert_paragraph_before()
        start, end = OxmlElement('w:bookmarkStart'), OxmlElement('w:bookmarkEnd')
        start.set(qn('w:id'), '42'); start.set(qn('w:name'), 'test_anchor')
        end.set(qn('w:id'), '42')
        blank._p.extend([start, end])
        doc.save(source)
        output = self.apply_custom(source)
        self.assertTrue(Document(output)._element.body.xpath('.//w:bookmarkStart[@w:id="42"]'))
        _, report = self.validate(output)
        self.assertTrue(any(c['rule'] == '带标记或对象空段的实际分隔间距' and c['status'] == 'manual' for c in report['checks']))

    def test_manual_pending_does_not_mean_formal_ready(self):
        output = self.apply_custom(self.custom_source())
        code, report = self.validate(output, '--require-manual-review')
        self.assertEqual(code, 2)
        self.assertEqual(report['summary']['failed'], 0)
        self.assertTrue(report['summary']['automatic_ok'])
        self.assertFalse(report['summary']['formal_ready'])
        self.assertGreater(report['summary']['pending_manual'], 0)

    def test_hidden_work_and_wps_lock_files_are_not_deliverables(self):
        folder = self.root / 'delivery'; folder.mkdir()
        work = folder / '.work'; work.mkdir()
        (work / 'source.docx').write_bytes(b'test')
        (work / 'preview.pdf').write_bytes(b'test')
        (folder / '~$lock.docx').write_bytes(b'test')
        self.assertEqual(visible_files(folder), [])
        (folder / 'deliverable.docx').write_bytes(b'test')
        self.assertEqual([p.name for p in visible_files(folder)], ['deliverable.docx'])


if __name__ == "__main__":
    unittest.main()
