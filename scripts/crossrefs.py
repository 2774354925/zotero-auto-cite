#!/usr/bin/env python3
"""
crossrefs.py — 给 docx 自动加「图/表题注」与「交叉引用」（编号自动更新）。

用法约定：
- 题注：在某个段落开头写  [图:fig1] 或 [表:tab1]，后面接题注文字，例如：
        [图:workflow] 本研究的技术路线
        [表:cohort] 入组患者基线特征
- 正文引用：写  [#fig1] 或  [#tab1]，例如：
        见图[#workflow]、如(表[#cohort])所示

脚本会：
1. 把题注标记替换成「标签 + SEQ 域（自动编号）+ 书签」；
2. 把正文 [#id] 替换成指向该书签的 REF 域（编号自动更新）。

在 Word 中 Ctrl+A 再按 F9（或打印预览）即可刷新所有编号。

用法：
    python crossrefs.py draft.docx -o out.docx
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from lxml import etree as ET

_xml_fromstring = ET.fromstring

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ET.register_namespace("w", W_NS)

CAPTION_RE = re.compile(r"\[(图|表):([A-Za-z0-9_\-]+)\]\s*")
REF_RE = re.compile(r"\[#([A-Za-z0-9_\-]+)\]")
SEQ_NAME = {"图": "Figure", "表": "Table"}
LABEL = {"图": "图", "表": "表"}


def _make_run(rpr: ET.Element | None) -> ET.Element:
    r = ET.Element(W + "r")
    if rpr is not None:
        r.append(_clone(rpr))
    return r


def _clone(el: ET.Element) -> ET.Element:
    new = ET.Element(el.tag, el.attrib)
    new.text = el.text
    new.tail = el.tail
    for c in el:
        new.append(_clone(c))
    return new


def _complex_field(instr: str, result_text: str, rpr: ET.Element | None) -> list[ET.Element]:
    runs = []
    r = _make_run(rpr)
    fc = ET.SubElement(r, W + "fldChar")
    fc.set(W + "fldCharType", "begin")
    runs.append(r)
    r = _make_run(rpr)
    it = ET.SubElement(r, W + "instrText")
    it.set(XML_SPACE, "preserve")
    it.text = f" {instr} "
    runs.append(r)
    r = _make_run(rpr)
    fc = ET.SubElement(r, W + "fldChar")
    fc.set(W + "fldCharType", "separate")
    runs.append(r)
    r = _make_run(rpr)
    t = ET.SubElement(r, W + "t")
    t.set(XML_SPACE, "preserve")
    t.text = result_text
    runs.append(r)
    r = _make_run(rpr)
    fc = ET.SubElement(r, W + "fldChar")
    fc.set(W + "fldCharType", "end")
    runs.append(r)
    return runs


def _bookmark_start(bid: int, name: str) -> ET.Element:
    bs = ET.Element(W + "bookmarkStart")
    bs.set(W + "id", str(bid))
    bs.set(W + "name", name)
    return bs


def _bookmark_end(bid: int) -> ET.Element:
    be = ET.Element(W + "bookmarkEnd")
    be.set(W + "id", str(bid))
    return be


def _para_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iter(W + "t"))


def _para_first_rpr(p: ET.Element) -> ET.Element | None:
    for r in p.iter(W + "r"):
        rpr = r.find(W + "rPr")
        if rpr is not None:
            return rpr
    return None


def _rebuild_para(p: ET.Element, new_children: list[ET.Element]):
    for child in list(p):
        p.remove(child)
    for c in new_children:
        p.append(c)


def process_docx(input_path: str, output_path: str, plain: bool = False) -> dict:
    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")

    for prefix, uri in set(re.findall(r'xmlns:([\w.-]+)="([^"]+)"', doc_xml)):
        if prefix != "xml":
            ET.register_namespace(prefix, uri)

    root = _xml_fromstring(doc_xml.encode("utf-8"))
    paras = list(root.iter(W + "p"))

    # 第一遍：登记题注，分配编号与书签
    counter = {"图": 0, "表": 0}
    bookmark = {}
    bm_id = 1000
    for p in paras:
        text = _para_text(p)
        m = CAPTION_RE.search(text)
        if not m:
            continue
        kind, pid = m.group(1), m.group(2)
        if pid in bookmark:
            continue
        counter[kind] += 1
        name = f"_Ref_{pid}"
        bookmark[pid] = {"kind": kind, "num": counter[kind], "name": name}

    # 第二遍：替换题注标记与正文引用
    n_caption = n_ref = 0
    for p in paras:
        text = _para_text(p)
        m = CAPTION_RE.search(text)
        if m:
            kind, pid = m.group(1), m.group(2)
            info = bookmark.get(pid)
            if info:
                _replace_caption(p, kind, info, m.end(), plain=plain)
                n_caption += 1
                continue
        # 正文引用
        if "[#" not in text:
            continue
        n_ref += _replace_refs(p, bookmark, plain=plain)

    new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8")
    m_orig = re.search(r"<w:document\b[^>]*>", doc_xml)
    m_new = re.search(r"<w:document\b[^>]*>", new_xml)
    if m_orig and m_new:
        declared = dict(re.findall(r'xmlns:([\w.-]+)="([^"]+)"', m_new.group(0)))
        missing = [(p_, u) for p_, u in re.findall(r'xmlns:([\w.-]+)="([^"]+)"', m_orig.group(0))
                   if p_ not in declared]
        if missing:
            insertion = "".join(f' xmlns:{p_}="{u}"' for p_, u in missing)
            new_xml = new_xml[:m_new.start()] + m_new.group(0)[:-1] + insertion + ">" + new_xml[m_new.end():]

    with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            if item == "word/document.xml":
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item))
    return {"captions": n_caption, "refs": n_ref, "bookmarks": bookmark}


def _replace_caption(p: ET.Element, kind: str, info: dict, marker_end: int,
                     plain: bool = False):
    """把 [图:id] 替换成 标签 + SEQ 域 + 书签，剩余文本作为题注。"""
    rpr = _para_first_rpr(p)
    # 收集段落文本段
    segs = _linearize(p)
    # marker 在文本中的偏移
    full = "".join(s["text"] for s in segs if s["kind"] == "text")
    # 定位 marker 结束后的文本（marker_end 是相对 full 的偏移）
    new_children: list[ET.Element] = []

    # 先输出标签（图 / 表）
    lab_r = _make_run(rpr)
    t = ET.SubElement(lab_r, W + "t")
    t.set(XML_SPACE, "preserve")
    t.text = LABEL[kind]
    new_children.append(lab_r)

    if plain:
        num_r = _make_run(rpr)
        t = ET.SubElement(num_r, W + "t")
        t.set(XML_SPACE, "preserve")
        t.text = str(info["num"])
        new_children.append(num_r)
    else:
        # 书签 + SEQ 域
        new_children.append(_bookmark_start(_next_bid(), info["name"]))
        seq_instr = f"SEQ {SEQ_NAME[kind]} \\* ARABIC"
        for e in _complex_field(seq_instr, str(info["num"]), rpr):
            new_children.append(e)
        new_children.append(_bookmark_end(_next_bid()))

    # 题注文本前的空格（渲染为「图1 本研究…」）
    sp_r = _make_run(rpr)
    t = ET.SubElement(sp_r, W + "t")
    t.set(XML_SPACE, "preserve")
    t.text = " "
    new_children.append(sp_r)

    # 剩余题注文本（截掉 marker）
    _emit_remaining(new_children, segs, marker_end, rpr)
    _rebuild_para(p, new_children)


_bid_counter = [2000]
def _next_bid() -> int:
    _bid_counter[0] += 1
    return _bid_counter[0]


def _linearize(p: ET.Element) -> list[dict]:
    segs = []
    for child in list(p):
        if child.tag == W + "r":
            for sub in list(child):
                if sub.tag == W + "t":
                    segs.append({"kind": "text", "run": child, "text": sub.text or ""})
                    break
            else:
                segs.append({"kind": "other", "elem": child})
        else:
            segs.append({"kind": "other", "elem": child})
    return segs


def _emit_remaining(new_children, segs, marker_end, rpr):
    """从 marker_end 之后开始输出剩余文本段。"""
    cur = 0
    for s in segs:
        if s["kind"] == "other":
            new_children.append(_clone(s["elem"]))
            continue
        start = cur
        end = cur + len(s["text"])
        if end <= marker_end:
            cur = end
            continue
        part = s["text"][max(0, marker_end - start):]
        if part:
            r = _make_run(rpr if rpr is not None else s["run"].find(W + "rPr"))
            t = ET.SubElement(r, W + "t")
            t.set(XML_SPACE, "preserve")
            t.text = part
            new_children.append(r)
        cur = end


def _replace_refs(p: ET.Element, bookmark: dict, plain: bool = False) -> int:
    """把段落里的 [#id] 替换成 REF 域。"""
    segs = _linearize(p)
    full = "".join(s["text"] for s in segs if s["kind"] == "text")
    refs = [(m.start(), m.end(), m.group(1)) for m in REF_RE.finditer(full)]
    if not refs:
        return 0

    offsets = []
    cur = 0
    for s in segs:
        if s["kind"] == "text":
            offsets.append((cur, cur + len(s["text"]), s))
            cur += len(s["text"])
        else:
            offsets.append((cur, cur, s))

    new_children: list[ET.Element] = []
    n = 0
    r_idx = 0

    def emit_text(text, run_template):
        if not text:
            return
        r = ET.Element(W + "r")
        rpr = run_template.find(W + "rPr") if run_template is not None else None
        if rpr is not None:
            r.append(_clone(rpr))
        t = ET.SubElement(r, W + "t")
        t.set(XML_SPACE, "preserve")
        t.text = text
        new_children.append(r)

    for ostart, oend, seg in offsets:
        if seg["kind"] == "other":
            new_children.append(_clone(seg["elem"]))
            continue
        text = seg["text"]
        run = seg["run"]
        cursor = ostart
        while r_idx < len(refs) and refs[r_idx][1] <= ostart:
            r_idx += 1
        while r_idx < len(refs) and refs[r_idx][0] < oend:
            rs, re_, pid = refs[r_idx]
            if rs > cursor:
                emit_text(text[cursor - ostart:min(rs, oend) - ostart], run)
                cursor = min(rs, oend)
            if re_ <= oend:
                info = bookmark.get(pid)
                if info:
                    rpr = run.find(W + "rPr")
                    if plain:
                        emit_text(str(info["num"]), run)
                    else:
                        instr = f"REF {info['name']} \\h"
                        for e in _complex_field(instr, str(info["num"]), rpr):
                            new_children.append(e)
                    n += 1
                else:
                    emit_text(text[rs - ostart:re_ - ostart], run)
                cursor = re_
                r_idx += 1
            else:
                if rs > cursor:
                    emit_text(text[cursor - ostart:oend - ostart], run)
                cursor = oend
                break
        if cursor < oend:
            emit_text(text[cursor - ostart:], run)

    _rebuild_para(p, new_children)
    return n


def main():
    args = sys.argv[1:]
    input_path = None
    output_path = None
    plain = True
    i = 0
    while i < len(args):
        a = args[i]
        if a in ("-o", "--out"):
            i += 1
            output_path = args[i]
        elif a == "--plain":
            plain = True
        elif a == "--field":
            plain = False
        elif a in ("-h", "--help"):
            print(__doc__)
            return
        elif input_path is None:
            input_path = a
        i += 1
    if not input_path:
        print(__doc__)
        return
    input_path = Path(input_path)
    output_path = Path(output_path) if output_path else input_path.with_name(input_path.stem + "_crossref.docx")
    result = process_docx(str(input_path), str(output_path), plain=plain)
    print(f"题注 {result['captions']} 处，交叉引用 {result['refs']} 处")
    print(f"写入 {output_path}")
    if plain:
        print("已写入 WPS 可直接显示的静态题注和交叉引用。")
    else:
        print("提示：Word 中 Ctrl+A 然后 F9 刷新编号。")


if __name__ == "__main__":
    main()
