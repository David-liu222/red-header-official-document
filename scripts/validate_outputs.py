#!/usr/bin/env python3
"""Validate red-header document output folder.

This script is intentionally contract-level: it checks that user-visible files
exist. Deep formatting validation should still be done with
validate_red_header_docx.py and rendered review.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


def visible_files(output_dir: Path) -> list[Path]:
    ignored = {
        "skill-run-metadata.json",
        "task-validation.json",
        "output-validation.json",
        "codex-last-message.md",
    }
    return sorted(
        path for path in output_dir.rglob("*")
        if path.is_file() and not path.name.startswith(".") and path.name not in ignored
    )


def main() -> int:
    if len(sys.argv) != 2:
        print(json.dumps({"ok": False, "errors": ["usage: validate_outputs.py output_dir"]}, ensure_ascii=False, indent=2))
        return 2
    output_dir = Path(sys.argv[1])
    files = visible_files(output_dir)
    names = [path.name for path in files]
    docx_files = [name for name in names if name.lower().endswith(".docx")]
    report_files = [
        name for name in names
        if ("验收" in name or "报告" in name or "检查" in name) and name.lower().endswith((".docx", ".md", ".pdf", ".txt"))
    ]
    errors: list[str] = []
    warnings: list[str] = []

    if not docx_files:
        errors.append("缺少可下载的套红头 DOCX")
    if not report_files:
        errors.append("缺少验收报告或检查报告")
    if any(path.stat().st_size == 0 for path in files):
        errors.append("存在 0 字节输出文件")
    if not any(name.lower().endswith(".pdf") for name in names):
        warnings.append("未输出 PDF 预览；如 outputMode=docxPdf，应在报告中说明原因")

    result = {
        "ok": not errors,
        "errors": errors,
        "warnings": warnings,
        "file_count": len(files),
        "files": names,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
