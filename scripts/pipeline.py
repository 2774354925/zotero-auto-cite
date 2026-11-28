#!/usr/bin/env python3
"""
pipeline.py — 一条命令完成「拉取文献 -> 导入 Zotero -> 文章引用 -> 图表交叉引用」。

输入：Markdown 草稿 或 已有 docx（正文里用 [@doi:…] / [@pmid:…] / [@arxiv:…] / [@key:…]
      标注引用，用 [图:id] 题注 + [#id] 交叉引用）。
默认输出 WPS 安全版：引文、参考文献、图题和交叉引用都直接显示，不依赖
Word 宏或 Zotero 域刷新。需要 Word 原生活引用时加 `--target word`。

用法：
    python pipeline.py draft.md -o 论文.docx
    python pipeline.py draft.md -o 论文.docx --style gb7714 --collection "AI导入"
    python pipeline.py draft.md -o 论文.docx --target word
    python pipeline.py draft.md -o 论文.docx --target both
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
TARGETS = {"wps", "wps-live", "word", "both"}


def _docx_path(path: Path) -> Path:
    return path if path.suffix.lower() == ".docx" else path.with_suffix(".docx")


def _render_wps(docx_input: Path, output_path: Path, resolver, style: str,
                work_dir: Path) -> int:
    from inject_citations import inject_docx
    from crossrefs import process_docx

    cited_tmp = work_dir / "wps_cited.docx"
    n = inject_docx(str(docx_input), str(cited_tmp), resolver, style, plain=True)
    result = process_docx(str(cited_tmp), str(output_path), plain=True)
    print(f"      WPS 固定引文 {n} 处；题注 {result['captions']} 处，"
          f"交叉引用 {result['refs']} 处")
    return n


def _render_word(docx_input: Path, output_path: Path, resolver, style: str,
                 work_dir: Path, with_preview: bool) -> None:
    from inject_citations import inject_docx
    from crossrefs import process_docx

    cited_tmp = work_dir / "word_cited.docx"
    n = inject_docx(str(docx_input), str(cited_tmp), resolver, style)
    result = process_docx(str(cited_tmp), str(output_path))
    print(f"      Word 活动引用 {n} 处；题注 {result['captions']} 处，"
          f"交叉引用 {result['refs']} 处")

    if with_preview:
        preview_tmp = work_dir / "word_preview_raw.docx"
        preview_path = output_path.with_name(output_path.stem + "_预览.docx")
        inject_docx(str(docx_input), str(preview_tmp), resolver, style, plain=True)
        process_docx(str(preview_tmp), str(preview_path), plain=True)
        print(f"      可读预览 -> {preview_path.name}")


def _render_wps_live(docx_input: Path, output_path: Path, resolver, style: str,
                     work_dir: Path) -> int:
    """WPS 活动域版：保留 ZOTERO_ITEM/BIBL，但不写 WPS 会露出的 PREF 域。"""
    from inject_citations import inject_docx
    from crossrefs import process_docx

    cited_tmp = work_dir / "wps_live_cited.docx"
    n = inject_docx(
        str(docx_input), str(cited_tmp), resolver, style,
        include_pref=False,
    )
    result = process_docx(str(cited_tmp), str(output_path))
    print(f"      WPS 活动引用 {n} 处；题注 {result['captions']} 处，"
          f"交叉引用 {result['refs']} 处")
    return n


def main():
    args = sys.argv[1:]
    input_path = None
    output_path = None
    style = STYLES["apa"]
    ris_path = None
    collection = DEFAULT_COLLECTION
    no_preview = False
    target = "wps"
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
        elif a in ("--target", "--engine"):
            i += 1
            target = args[i].lower()
        elif a == "--wps":
            target = "wps"
        elif a == "--wps-live":
            target = "wps-live"
        elif a == "--word":
            target = "word"
        elif a in ("-h", "--help"):
            print(__doc__)
            return
        elif input_path is None and not a.startswith("-"):
            input_path = a
        i += 1

    style = STYLES.get(style.lower(), style)
    if target not in TARGETS:
        raise SystemExit("--target 必须是 wps、wps-live、word 或 both")

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
    output_path = _docx_path(
        Path(output_path) if output_path else input_path.with_name(input_path.stem + "_final.docx")
    )

    work_dir = Path(tempfile.mkdtemp(prefix="zotero_pipe_"))
    docx_input = input_path if input_path.suffix.lower() == ".docx" else None

    # 1) Markdown -> docx
    if docx_input is None:
        from md2docx import convert
        md_docx = work_dir / "draft.docx"
        convert(input_path.read_text(encoding="utf-8")).save(str(md_docx))
        docx_input = md_docx
        print(f"[1/3] Markdown -> {docx_input}")

    # 2) 文献解析并写入 Zotero
    from zotero_client import ZoteroClient
    from inject_citations import make_resolver
    print("[2/3] 拉取文献并写入 Zotero（首次写库请点 Always Allow）...")
    z = ZoteroClient()
    collection_key = z.ensure_collection(collection)
    print(f"      文献将归入 Zotero 分类「{collection}」")
    resolver = make_resolver(collection_key=collection_key)
    print("[3/3] 生成文档...")
    if target == "wps":
        _render_wps(docx_input, output_path, resolver, style, work_dir)
    elif target == "wps-live":
        _render_wps_live(docx_input, output_path, resolver, style, work_dir)
    elif target == "word":
        _render_word(
            docx_input, output_path, resolver, style, work_dir,
            with_preview=not no_preview,
        )
    else:
        _render_wps(docx_input, output_path, resolver, style, work_dir)
        word_path = output_path.with_name(output_path.stem + "_Zotero活动引用.docx")
        _render_word(
            docx_input, word_path, resolver, style, work_dir,
            with_preview=False,
        )

    print(f"\n完成 -> {output_path}")
    print(f"引用样式：{style}")
    if target in ("wps", "both"):
        print("WPS 可直接打开上述文件；引文、参考文献和图题均已固化为可见文本。")
    if target == "wps-live":
        print("WPS 活动域版已生成；全选应能识别 Zotero 引用域。")
    if target == "word":
        print("用桌面 Word 打开，点 Zotero 选项卡中的 Refresh 渲染活动引用。")


if __name__ == "__main__":
    main()
