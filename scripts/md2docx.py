#!/usr/bin/env python3
"""
md2docx.py — 极简 Markdown -> docx 转换（用于 AI 草稿快速转 Word）。

支持：标题(#/##/###)、段落、无序列表(-/*)、有序列表(1.)、**粗体**、*斜体*。
文献标记 [@...] 与交叉引用标记 [#id] / [图:id] 原样保留，供后续脚本处理。

用法：
    python md2docx.py draft.md -o draft.docx
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document

_INLINE = re.compile(r"(\*\*.+?\*\*|\*[^*\s][^*]*?\*)")


def add_inline_runs(paragraph, text: str):
    pos = 0
    for m in _INLINE.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        seg = m.group(0)
        if seg.startswith("**") and seg.endswith("**") and len(seg) >= 4:
            r = paragraph.add_run(seg[2:-2])
            r.bold = True
        elif seg.startswith("*") and seg.endswith("*") and len(seg) >= 2:
            r = paragraph.add_run(seg[1:-1])
            r.italic = True
        else:
            paragraph.add_run(seg)
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def convert(md_text: str) -> Document:
    doc = Document()
    lines = md_text.splitlines()
    for line in lines:
        stripped = line.rstrip()
        if not stripped.strip():
            continue
        if stripped.startswith("### "):
            doc.add_heading(stripped[4:].strip(), level=3)
        elif stripped.startswith("## "):
            doc.add_heading(stripped[3:].strip(), level=2)
        elif stripped.startswith("# "):
            doc.add_heading(stripped[2:].strip(), level=1)
        elif re.match(r"^\s*[-*]\s+", stripped):
            p = doc.add_paragraph(style="List Bullet")
            add_inline_runs(p, re.sub(r"^\s*[-*]\s+", "", stripped))
        elif re.match(r"^\s*\d+[.)]\s+", stripped):
            p = doc.add_paragraph(style="List Number")
            add_inline_runs(p, re.sub(r"^\s*\d+[.)]\s+", "", stripped))
        else:
            p = doc.add_paragraph()
            add_inline_runs(p, stripped)
    return doc


def main():
    args = sys.argv[1:]
    input_path = None
    output_path = None
    i = 0
    while i < len(args):
        a = args[i]
        if a in ("-o", "--out"):
            i += 1
            output_path = args[i]
        elif input_path is None:
            input_path = a
        i += 1
    if not input_path:
        print(__doc__)
        return
    input_path = Path(input_path)
    output_path = Path(output_path) if output_path else input_path.with_suffix(".docx")
    doc = convert(input_path.read_text(encoding="utf-8"))
    doc.save(str(output_path))
    print(f"写入 {output_path}")


if __name__ == "__main__":
    main()
