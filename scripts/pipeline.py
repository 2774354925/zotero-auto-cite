#!/usr/bin/env python3
"""
pipeline.py — 一条命令完成「拉取文献 -> 导入 Zotero -> 文章引用 -> 图表交叉引用」。

输入：Markdown 草稿 或 已有 docx（正文里用 [@doi:…] / [@pmid:…] / [@arxiv:…] / [@key:…]
      标注引用，用 [图:id] 题注 + [#id] 交叉引用）。
输出：
    <name>.docx          Zotero 活动引用版（桌面 Word + Zotero 插件打开后 Refresh 渲染）
    <name>_预览.docx     纯文本预览版（任意编辑器可读，无域代码）

用法：
    python pipeline.py draft.md -o 论文.docx
    python pipeline.py draft.md -o 论文.docx --style gb7714 --collection "AI导入"
    python pipeline.py --ris 知网导出.txt                # 仅批量导入知网 RIS 文献

样式别名（--style 可接友好名或完整 Zotero style URL）：
    apa / nature / vancouver / ieee / chicago-author-date / gb7714
    国标 GB/T 7714 = http://www.zotero.org/styles/china-national-standard-gb-t-7714-2015-numeric

需要：Zotero 已运行（本地 API 已开启），首次写库会弹授权框（选 Always Allow）。
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

STYLES = {
    "apa": "http://www.zotero.org/styles/apa",
    "nature": "http://www.zotero.org/styles/nature",
    "vancouver": "http://www.zotero.org/styles/vancouver",
    "ieee": "http://www.zotero.org/styles/ieee",
    "chicago-author-date": "http://www.zotero.org/styles/chicago-author-date",
    "gb7714": "http://www.zotero.org/styles/china-national-standard-gb-t-7714-2015-numeric",
}
DEFAULT_COLLECTION = "AI导入"


def main():
    args = sys.argv[1:]
    input_path = None
    output_path = None
    style = STYLES["apa"]
    ris_path = None
    collection = DEFAULT_COLLECTION
    no_preview = False
    i = 0
    while i < len(args):
        a = args[i]
        if a in ("-o", "--out"):
            i += 1
            output_path = args[i]
        elif a == "--style":
            i += 1
            style = args[i]
        elif a == "--collection":
            i += 1
            collection = args[i]
        elif a == "--ris":
            i += 1
            ris_path = args[i]
        elif a == "--no-preview":
            no_preview = True
        elif a in ("-h", "--help"):
            print(__doc__)
            return
        elif input_path is None and not a.startswith("-"):
            input_path = a
        i += 1

    style = STYLES.get(style.lower(), style)

    # 仅批量导入知网 RIS
    if ris_path and not input_path:
        from zotero_client import ZoteroClient
        z = ZoteroClient()
        ckey = z.ensure_collection(collection)
        created = z.import_ris(Path(ris_path).read_text(encoding="utf-8-sig"), ckey)
        print(f"已导入 {len(created)} 条到 Zotero 分类「{collection}」：")
        for c in created:
            data = c.get("data") or c
            print(f"  {c.get('key')}  {data.get('title', '')[:60]}")
        return

    if not input_path:
        print(__doc__)
        return

    input_path = Path(input_path)
    output_path = Path(output_path) if output_path else input_path.with_name(input_path.stem + "_final.docx")

    work_dir = Path(tempfile.mkdtemp(prefix="zotero_pipe_"))
    docx_input = input_path if input_path.suffix.lower() == ".docx" else None

    # 1) Markdown -> docx
    if docx_input is None:
        from md2docx import convert
        md_docx = work_dir / "draft.docx"
        convert(input_path.read_text(encoding="utf-8")).save(str(md_docx))
        docx_input = md_docx
        print(f"[1/3] Markdown -> {docx_input}")

    # 2) 引用注入（field 版）
    from zotero_client import ZoteroClient
    from inject_citations import make_resolver, inject_docx
    print("[2/3] 拉取文献并注入 Zotero 引用（首次写库请点 Always Allow）...")
    z = ZoteroClient()
    collection_key = z.ensure_collection(collection)
    print(f"      文献将归入 Zotero 分类「{collection}」")
    resolver = make_resolver(collection_key=collection_key)
    cited_docx = work_dir / "cited.docx"
    n = inject_docx(str(docx_input), str(cited_docx), resolver, style)
    print(f"      注入 {n} 处引用")

    # 3) 图表交叉引用（field 版）
    from crossrefs import process_docx
    print("[3/3] 处理图/表交叉引用...")
    result = process_docx(str(cited_docx), str(output_path))
    print(f"      题注 {result['captions']} 处，交叉引用 {result['refs']} 处")

    # 4) 纯文本预览版（任意编辑器可读，无域代码）
    preview_path = None
    if not no_preview:
        from inject_citations import inject_docx as inject_plain
        prev_tmp = work_dir / "preview_raw.docx"
        inject_plain(str(docx_input), str(prev_tmp), resolver, style, plain=True)
        preview_path = output_path.with_name(output_path.stem + "_预览.docx")
        process_docx(str(prev_tmp), str(preview_path))
        print(f"      预览版 -> {preview_path.name}")

    print(f"\n完成 -> {output_path}")
    print(f"样式：{style}")
    print("正式引用：用桌面 Word 打开上面 docx，点 Zotero 选项卡 Refresh 渲染引用与文末 References；")
    if preview_path:
        print("想直接阅读：打开「_预览.docx」（任何编辑器都能看，无域代码）。")


if __name__ == "__main__":
    main()
