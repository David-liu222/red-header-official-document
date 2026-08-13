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
        and not any(part.startswith("render") for part in path.relative_to(output_dir).parts[:-1])
    )


def main() -> int:
    if len(sys.argv) != 2:
        print(json.dumps({"ok": False, "errors": ["usage: validate_outputs.py output_dir"]}, ensure_ascii=False, indent=2))
        return 2
    output_dir = Path(sys.argv[1])
    files = visible_files(output_dir)
    names = [path.name for path in files]
    docx_files = [name for name in names if name.lower().endswith(".docx")]
    errors: list[str] = []
    warnings: list[str] = []

    if not docx_files:
        errors.append("缺少可下载的套红头 DOCX")
    if any(path.stat().st_size == 0 for path in files):
        errors.append("存在 0 字节输出文件")
    if any(name.lower().endswith(".pdf") for name in names):
        errors.append("默认只产出 Word 文件，输出目录不应包含 PDF")

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
