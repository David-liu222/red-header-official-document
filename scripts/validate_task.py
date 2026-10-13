#!/usr/bin/env python3
"""Validate a web office-assistant task for red-header document generation.

The runner treats this as advisory input for Codex. It should be strict about
malformed task JSON, but ordinary missing document fields are warnings because
the skill can still produce a review draft.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


def value(data: dict, *keys: str) -> str:
    for key in keys:
        current = data
        ok = True
        for part in key.split("."):
            if not isinstance(current, dict) or part not in current:
                ok = False
                break
            current = current[part]
        if ok and str(current or "").strip():
            return str(current).strip()
    return ""


def main() -> int:
    if len(sys.argv) != 2:
        print(json.dumps({"status": "failed", "errors": ["usage: validate_task.py task.json"]}, ensure_ascii=False, indent=2))
        return 2

    path = Path(sys.argv[1])
    try:
        task = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        print(json.dumps({"status": "failed", "errors": [f"task.json 无法读取或解析：{exc}"]}, ensure_ascii=False, indent=2))
        return 2

    errors: list[str] = []
    warnings: list[str] = []
    notes: list[str] = []
    payload = task.get("payload") if isinstance(task.get("payload"), dict) else {}
    red = task.get("red_head_document") if isinstance(task.get("red_head_document"), dict) else {}
    source_files = task.get("source_files") if isinstance(task.get("source_files"), list) else []

    if task.get("task_type") != "office_red_head_document":
        errors.append("task_type 应为 office_red_head_document")
    if task.get("skill_name") != "red-header-official-document":
        errors.append("skill_name 应为 red-header-official-document")
    if not value(task, "output_folder"):
        errors.append("缺少 output_folder")

    fields = {
        "文件标题": value(red, "title") or value(payload, "title"),
        "主送单位": value(red, "recipient") or value(payload, "recipient"),
        "发文字号": value(red, "document_no") or value(payload, "documentNo"),
        "落款单位": value(red, "issue_unit") or value(payload, "issueUnit"),
        "成文日期": value(red, "document_date") or value(payload, "documentDate"),
        "正文": value(payload, "body"),
    }
    if not fields["文件标题"]:
        warnings.append("文件标题缺失：可从原稿识别，不确定时列为待确认")
    if not fields["主送单位"]:
        warnings.append("主送单位缺失：不得凭空判断内外发")
    if not fields["发文字号"]:
        warnings.append("发文字号字段未填写：先核查原稿已有文号；唯一且无冲突时沿用，原稿无文号且没有用户确认时才生成待补文号审阅稿")
    if not fields["落款单位"]:
        warnings.append("落款单位缺失：可按红头单位临时处理并列为待确认")
    if not fields["成文日期"]:
        warnings.append("成文日期缺失：不得自动使用当前日期冒充正式日期")
    if not fields["正文"] and not source_files:
        errors.append("缺少正文或上传原稿")

    sign_type = value(payload, "signType") or value(red, "sign_type") or "auto"
    if sign_type not in {"auto", "external", "internal"}:
        warnings.append(f"发文类型取值异常：{sign_type}")
    inspection = payload.get("inspectionMaterial", red.get("inspection_material", False))
    if not isinstance(inspection, bool):
        errors.append("inspectionMaterial 必须为 JSON 布尔值，不接受字符串代替")
        inspection = False
    if inspection:
        notes.append("迎检内发文：生成和校验均须传 --inspection-material；无签发人/印发栏，不要求抄送及印数")
        if sign_type == "external":
            warnings.append("迎检标记与外发选择冲突：须确认文件类型，不静默当作普通内发文")
    appointment = payload.get('appointmentNotice', red.get('appointment_notice', False))
    no_title = payload.get('noTitle', red.get('no_title', False))
    for name, flag in [('appointmentNotice', appointment), ('noTitle', no_title)]:
        if not isinstance(flag, bool):
            errors.append(f'{name} 必须为 JSON 布尔值')
    if appointment is True:
        if not inspection:
            errors.append('任职通知原内容分支须先确认迎检用途')
        notes.append('任职迎检通知保留原内容和原有加粗，不强制一、序号或新增标题；校验传 --source 原稿副本')
    if no_title is True:
        if appointment is not True:
            errors.append('noTitle 只用于任职类迎检原稿无标题分支')
        notes.append('原稿无正文标题：--title仅作文件名标签；生成和校验均传 --no-title')
    output_mode = value(payload, "outputMode") or value(red, "output_mode") or "docxOnly"
    if output_mode == "docxPdf":
        warnings.append("输出模式 docxPdf 已废弃：按 docxOnly 执行，不产出 PDF")
    if output_mode not in {"docxPdf", "docxOnly", "legacyDoc"}:
        warnings.append(f"输出模式取值异常：{output_mode}")

    extracted = [item for item in source_files if value(item, "extracted_text_path")]
    if source_files and not extracted:
        warnings.append("已上传文件，但没有可读取的提取文本；需要从原文件或 OCR 继续识别")
    if source_files:
        notes.append(f"上传材料 {len(source_files)} 个，可读取文本 {len(extracted)} 个")
    if value(payload, "body") and source_files:
        notes.append("网页正文和上传材料同时存在：应以原稿为主体，网页正文作补充说明；冲突写入报告")

    report = {
        "status": "failed" if errors else "ok",
        "errors": errors,
        "warnings": warnings,
        "notes": notes,
        "fields": {
            key: ("已填写" if text else "缺失") for key, text in fields.items()
        },
        "source_file_count": len(source_files),
        "sign_type": sign_type,
        "inspection_material": inspection,
        "appointment_notice": appointment,
        "no_title": no_title,
        "output_mode": output_mode,
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
