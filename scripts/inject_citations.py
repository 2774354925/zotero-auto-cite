#!/usr/bin/env python3
"""
inject_citations.py — 把 docx 中的文献标记替换成 Zotero 活动引用域代码，
并在文末自动追加 References 参考文献域；也可输出「纯文本预览版」（无域代码）。

支持的正文标记（写在 Word 正文里即可）：
    [@doi:10.1038/s41586-020-2649-2]
    [@pmid:32214244]
    [@arxiv:1706.03762]
    [@key:ABCD1234]                     # 直接引用库里已有条目
    [@doi:10.xxxx/yyyy, p. 12]          # 带定位符（页码等）

两种模式：
- field（默认）：替换为 ADDIN ZOTERO_ITEM CSL_CITATION 域 + 文末 ADDIN ZOTERO_BIBL 域，
  供桌面 Word + Zotero 插件 Refresh 渲染；引用与库保持双向关联。
- plain（--plain）：替换为 (Author, Year) 纯文本 + 文末纯文本参考文献表，供任意编辑器阅读。

用法：
    python inject_citations.py draft.docx -o out.docx
    python inject_citations.py draft.docx -o out.docx --style "http://www.zotero.org/styles/apa"
    python inject_citations.py draft.docx -o out.docx --plain
"""
from __future__ import annotations

import hashlib
import json
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

MARKER_RE = re.compile(r"\[@(doi|pmid|arxiv|key):([^\],\s]+)(?:,\s*([^\]]*))?\]", re.I)
SCHEMA_URL = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
NUMERIC_STYLE_MARKERS = ("numeric", "vancouver", "ieee", "nature", "gb-t-7714")


# ---------------- 纯函数 ----------------
def csl_item_id(key: str) -> int:
    """条目 key -> 稳定的 CSL 数字 id（引文与参考文献域必须一致）。"""
    return int(hashlib.md5(key.encode("utf-8")).hexdigest()[:8], 16)


def csl_clean(csl: dict) -> dict:
    return {k: v for k, v in csl.items()
            if k not in ("id", "citation-key", "library", "citekey")}


def csl_authors(csl: dict) -> list[str]:
    authors = csl.get("author") or csl.get("editor") or []
    fams = []
    for a in authors:
        fam = (a.get("family") or a.get("literal") or "").strip()
        if fam:
            fams.append(fam)
        elif a.get("name"):
            fams.append(a["name"].strip())
    return fams


def csl_year(csl: dict) -> str:
    issued = csl.get("issued") or {}
    dp = issued.get("date-parts") or [[None]]
    y = dp[0][0] if dp and dp[0] else None
    return str(y) if y else ""


def style_is_numeric(style: str) -> bool:
    s = (style or "").lower()
    return any(m in s for m in NUMERIC_STYLE_MARKERS)


def placeholder(csl: dict, locator: str = "", num: int | None = None,
                numeric: bool = False) -> str:
    """Refresh 之前的可见占位文本：作者年样式 (Smith, 2020) / 数字样式 [1]。"""
    if numeric and num is not None:
        if locator:
            return f"[{num}, {locator}]"
        return f"[{num}]"
    authors = csl_authors(csl)
    year = csl_year(csl)
    if authors:
        if len(authors) == 1:
            base = authors[0]
        elif len(authors) == 2:
            base = f"{authors[0]} & {authors[1]}"
        else:
            base = f"{authors[0]} et al."
    else:
        t = (csl.get("title") or "").strip()
        base = (t[:28] + "…") if len(t) > 28 else (t or "?")
    if year:
        base = f"{base}, {year}"
    if locator:
        base = f"{base}, {locator}"
    return f"({base})"


def bibliography_text(csl: dict) -> str:
    """纯文本参考文献占位（近似 APA；Refresh 后由 Zotero 用所选样式重排）。"""
    authors = csl_authors(csl)
    year = csl_year(csl)
    title = (csl.get("title") or "").strip()
    container = (csl.get("container-title") or csl.get("publicationTitle")
                 or csl.get("bookTitle") or "").strip()
    volume = (csl.get("volume") or "").strip()
    issue = (csl.get("issue") or "").strip()
    pages = (csl.get("page") or "").strip()
    doi = (csl.get("DOI") or "").strip()

    s = ""
    if authors:
        if len(authors) == 1:
            s += authors[0]
        elif len(authors) <= 7:
            s += ", ".join(authors[:-1]) + ", & " + authors[-1]
        else:
            s += ", ".join(authors[:6]) + ", … " + authors[-1]
        s += " "
    if year:
        s += f"({year}). "
    if title:
        s += title.rstrip(".") + ". "
    if container:
        s += container.rstrip(".") + ". "
    if volume:
        s += f"{volume}"
        if issue:
            s += f"({issue})"
        s += ". "
    if pages:
        s += pages.rstrip(".") + ". "
    if doi:
        s += f"https://doi.org/{doi}"
    return s.strip()


def build_citation_json(item: dict, formatted: str) -> str:
    """item: {key, csl, uri} -> CSL_CITATION 的 JSON 字符串。"""
    key = item["key"]
    csl = csl_clean(item["csl"])
    numeric_id = csl_item_id(key)
    csl["id"] = numeric_id
    citation_item = {
        "id": numeric_id,
        "uris": [item["uri"]],
        "uri": [item["uri"]],
        "itemData": csl,
    }
    payload = {
        "citationID": hashlib.md5(formatted.encode("utf-8")).hexdigest()[:12],
        "properties": {
            "formattedCitation": formatted,
            "plainCitation": formatted,
            "noteIndex": 0,
        },
        "citationItems": [citation_item],
        "schema": SCHEMA_URL,
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def build_biblio_json(items: list[dict]) -> str:
    citation_items = []
    for item in items:
        csl = csl_clean(item["csl"])
        numeric_id = csl_item_id(item["key"])
        csl["id"] = numeric_id
        citation_items.append({
            "id": numeric_id,
            "uris": [item["uri"]],
            "uri": [item["uri"]],
            "itemData": csl,
        })
    payload = {
        "citationItems": citation_items,
        "properties": {"noteIndex": 0},
        "schema": SCHEMA_URL,
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


# ---------------- 域代码构造 ----------------
def _clone(el: ET.Element) -> ET.Element:
    new = ET.Element(el.tag, el.attrib)
    new.text = el.text
    new.tail = el.tail
    for c in el:
        new.append(_clone(c))
    return new


def _make_run(rpr: ET.Element | None) -> ET.Element:
    r = ET.Element(W + "r")
    if rpr is not None:
        r.append(_clone(rpr))
    return r


def field_runs(formatted: str, csl_json: str, rpr: ET.Element | None) -> list[ET.Element]:
    runs = []
    r = _make_run(rpr)
    fc = ET.SubElement(r, W + "fldChar")
    fc.set(W + "fldCharType", "begin")
    runs.append(r)
    r = _make_run(rpr)
    it = ET.SubElement(r, W + "instrText")
    it.set(XML_SPACE, "preserve")
    it.text = f" ADDIN ZOTERO_ITEM CSL_CITATION {csl_json} "
    runs.append(r)
    r = _make_run(rpr)
    fc = ET.SubElement(r, W + "fldChar")
    fc.set(W + "fldCharType", "separate")
    runs.append(r)
    r = _make_run(rpr)
    t = ET.SubElement(r, W + "t")
    t.set(XML_SPACE, "preserve")
    t.text = formatted
    runs.append(r)
    r = _make_run(rpr)
    fc = ET.SubElement(r, W + "fldChar")
    fc.set(W + "fldCharType", "end")
    runs.append(r)
    return runs


def biblio_field_runs(biblio_json: str, entries: list[str],
                      rpr: ET.Element | None) -> list[ET.Element]:
    """ADDIN ZOTERO_BIBL 域；entries 为 Refresh 前可见的参考文献占位。"""
    runs = []
    r = _make_run(rpr)
    fc = ET.SubElement(r, W + "fldChar")
    fc.set(W + "fldCharType", "begin")
    runs.append(r)
    r = _make_run(rpr)
    it = ET.SubElement(r, W + "instrText")
    it.set(XML_SPACE, "preserve")
    it.text = f" ADDIN ZOTERO_BIBL {biblio_json} "
    runs.append(r)
    r = _make_run(rpr)
    fc = ET.SubElement(r, W + "fldChar")
    fc.set(W + "fldCharType", "separate")
    runs.append(r)
    for i, entry in enumerate(entries):
        if i > 0:
            r = _make_run(rpr)
            ET.SubElement(r, W + "br")
            runs.append(r)
        r = _make_run(rpr)
        t = ET.SubElement(r, W + "t")
        t.set(XML_SPACE, "preserve")
        t.text = entry
        runs.append(r)
    r = _make_run(rpr)
    fc = ET.SubElement(r, W + "fldChar")
    fc.set(W + "fldCharType", "end")
    runs.append(r)
    return runs


def pref_runs(style: str = "http://www.zotero.org/styles/apa",
              locale: str = "en-US") -> list[ET.Element]:
    """文档级 ZOTERO_PREF_1，让 Refresh 免去手动配置 Document Preferences。"""
    session_id = hashlib.md5(b"zotero-auto-cite").hexdigest()[:13].upper()
    pref_xml = (
        '<data data-version="3" zotero-version="7.0.0">'
        f'<session id="{session_id}" />'
        f'<style id="{style}" locale="{locale}" hasBibliography="1" '
        'bibliographyStyleHasBeenSet="1" />'
        '<prefs>'
        '<pref name="fieldType" value="Field" />'
        '<pref name="automaticJournalAbbreviations" value="false" />'
        '<pref name="noteType" value="0" />'
        '</prefs></data>'
    )
    runs = []
    for piece in ("begin", "instr", "separate", "end"):
        r = ET.Element(W + "r")
        rPr = ET.SubElement(r, W + "rPr")
        ET.SubElement(rPr, W + "vanish")
        if piece == "begin":
            fc = ET.SubElement(r, W + "fldChar")
            fc.set(W + "fldCharType", "begin")
        elif piece == "instr":
            it = ET.SubElement(r, W + "instrText")
            it.set(XML_SPACE, "preserve")
            it.text = f" ADDIN ZOTERO_PREF_1 {pref_xml}"
        elif piece == "separate":
            fc = ET.SubElement(r, W + "fldChar")
            fc.set(W + "fldCharType", "separate")
        else:
            fc = ET.SubElement(r, W + "fldChar")
            fc.set(W + "fldCharType", "end")
        runs.append(r)
    return runs


# ---------------- 段落级替换 ----------------
def _linearize(p: ET.Element) -> list[dict]:
    segs = []
    for child in list(p):
        if child.tag == W + "r":
            handled = False
            for sub in list(child):
                if sub.tag == W + "t":
                    segs.append({"kind": "text", "run": child, "text": sub.text or ""})
                    handled = True
            if not handled:
                segs.append({"kind": "other", "elem": child})
        else:
            segs.append({"kind": "other", "elem": child})
    return segs


def replace_in_paragraph(p: ET.Element, resolve, state: dict) -> int:
    """resolve(kind, value) -> {'key','csl','uri'}。返回替换的引文数量。

    state: {'numeric', 'counter', 'cited', 'seen', 'plain'}
    """
    segs = _linearize(p)
    full = "".join(s["text"] for s in segs if s["kind"] == "text")
    if "[@" not in full:
        return 0

    markers = [(m.start(), m.end(), m.group(1).lower(), m.group(2).strip(),
                (m.group(3) or "").strip()) for m in MARKER_RE.finditer(full)]
    if not markers:
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
    n_replaced = 0
    m_idx = 0

    def emit_text(text, run_template):
        if not text:
            return
        r = ET.Element(W + "r")
        if run_template is not None:
            rpr = run_template.find(W + "rPr")
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
        while m_idx < len(markers) and markers[m_idx][1] <= ostart:
            m_idx += 1
        while m_idx < len(markers) and markers[m_idx][0] < oend:
            ms, me, kind, value, locator = markers[m_idx]
            if ms > cursor:
                emit_text(text[cursor - ostart:min(ms, oend) - ostart], run)
                cursor = min(ms, oend)
            if me <= oend:
                try:
                    item = resolve(kind, value)
                    key = item["key"]
                    if key not in state["seen"]:
                        state["seen"].add(key)
                        state["counter"] += 1
                        num = state["counter"]
                        state["cited"].append(dict(item, _num=num))
                    else:
                        num = next(it["_num"] for it in state["cited"] if it["key"] == key)
                    formatted = placeholder(item["csl"], locator, num=num,
                                            numeric=state["numeric"])
                    if state["plain"]:
                        emit_text(formatted, run)
                    else:
                        csl_json = build_citation_json(item, formatted)
                        rpr = run.find(W + "rPr")
                        for r_elem in field_runs(formatted, csl_json, rpr):
                            new_children.append(r_elem)
                    n_replaced += 1
                except Exception as e:
                    emit_text(text[ms - ostart:me - ostart], run)
                    print(f"  [warn] 无法解析 [@{kind}:{value}] -> {e}", file=sys.stderr)
                cursor = me
                m_idx += 1
            else:
                if ms > cursor:
                    emit_text(text[cursor - ostart:oend - ostart], run)
                cursor = oend
                break
        if cursor < oend:
            emit_text(text[cursor - ostart:], run)

    for child in list(p):
        p.remove(child)
    for c in new_children:
        p.append(c)
    return n_replaced


# ---------------- 解析器（依赖 Zotero + resolvers） ----------------
def make_resolver(collection_key: str | None = None):
    import os
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from zotero_client import ZoteroClient
    from resolvers import resolve, detect_identifier

    z = ZoteroClient()
    cache: dict[str, dict] = {}

    def _uri_from_csl(csl: dict, key: str) -> str:
        cid = csl.get("id")
        if isinstance(cid, str) and cid.startswith("http"):
            return cid
        return f"http://zotero.org/users/0/items/{key}"

    def resolver(kind: str, value: str) -> dict:
        ident = f"{kind}:{value}"
        if ident in cache:
            return cache[ident]

        if kind == "key":
            csl = z.get_item_csl(value)
            item = {"key": value, "csl": csl, "uri": _uri_from_csl(csl, value)}
        else:
            ident_full = value if kind == "doi" else f"{kind}:{value}"
            meta = resolve(ident_full)
            doi = meta.get("DOI")
            existing = z.find_by_doi(doi) if doi else None
            if existing:
                key = existing["key"]
                csl = z.get_item_csl(key)
                item = {"key": key, "csl": csl, "uri": _uri_from_csl(csl, key)}
                if collection_key:
                    z.add_to_collection(collection_key, key)
            else:
                created = z.create_item(meta, collection_key)
                key = created["key"]
                csl = z.get_item_csl(key)
                item = {"key": key, "csl": csl, "uri": _uri_from_csl(csl, key)}
        cache[ident] = item
        return item

    return resolver


# ---------------- 参考文献追加 ----------------
def _append_references(root: ET.Element, items: list[dict], plain: bool):
    body = root.find(W + "body")
    if body is None or not items:
        return
    sectpr = body.find(W + "sectPr")

    h = ET.Element(W + "p")
    hpPr = ET.SubElement(h, W + "pPr")
    ps = ET.SubElement(hpPr, W + "pStyle")
    ps.set(W + "val", "Heading1")
    hr = ET.SubElement(h, W + "r")
    ht = ET.SubElement(hr, W + "t")
    ht.set(XML_SPACE, "preserve")
    ht.text = "References"

    bp = ET.Element(W + "p")
    if plain:
        for i, item in enumerate(items):
            if i > 0:
                r = ET.SubElement(bp, W + "r")
                ET.SubElement(r, W + "br")
            r = ET.SubElement(bp, W + "r")
            t = ET.SubElement(r, W + "t")
            t.set(XML_SPACE, "preserve")
            t.text = f"{i + 1}. {bibliography_text(item['csl'])}"
    else:
        entries = [bibliography_text(item["csl"]) for item in items]
        biblio_json = build_biblio_json(items)
        for run in biblio_field_runs(biblio_json, entries, None):
            bp.append(run)

    if sectpr is not None:
        idx = list(body).index(sectpr)
        body.insert(idx, h)
        body.insert(idx + 1, bp)
    else:
        body.append(h)
        body.append(bp)


# ---------------- 主流程 ----------------
def inject_docx(input_path: str, output_path: str, resolve, style: str,
                plain: bool = False, biblio: bool = True) -> int:
    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")

    for prefix, uri in set(re.findall(r'xmlns:([\w.-]+)="([^"]+)"', doc_xml)):
        if prefix != "xml":
            ET.register_namespace(prefix, uri)

    root = _xml_fromstring(doc_xml.encode("utf-8"))
    state = {
        "numeric": style_is_numeric(style),
        "counter": 0,
        "cited": [],
        "seen": set(),
        "plain": plain,
    }
    paras = list(root.iter(W + "p"))
    n_total = 0
    for p in paras:
        n_total += replace_in_paragraph(p, resolve, state)
    print(f"替换了 {n_total} 处引文")

    if n_total:
        if not plain:
            for run in pref_runs(style):
                paras[0].insert(_first_insert_index(paras[0]), run)
        if biblio and state["cited"]:
            _append_references(root, state["cited"], plain=plain)

    new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8")

    # 补回 ET 可能丢弃的根命名空间声明
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
    return n_total


def _first_insert_index(p: ET.Element) -> int:
    idx = 0
    for i, child in enumerate(list(p)):
        if child.tag == W + "pPr":
            idx = i + 1
        else:
            break
    return idx


def main():
    args = [a for a in sys.argv[1:]]
    input_path = None
    output_path = None
    style = "http://www.zotero.org/styles/apa"
    plain = False
    biblio = True
    collection_key = None
    i = 0
    while i < len(args):
        a = args[i]
        if a in ("-o", "--out"):
            i += 1
            output_path = args[i]
        elif a == "--style":
            i += 1
            style = args[i]
        elif a == "--plain":
            plain = True
        elif a == "--no-biblio":
            biblio = False
        elif a == "--collection":
            i += 1
            collection_key = args[i]
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
    output_path = Path(output_path) if output_path else input_path.with_name(input_path.stem + "_zotero.docx")

    print(f"输入: {input_path}")
    print("解析并写入 Zotero（如未运行请先启动 Zotero，首次写入会弹授权框，请点 Always Allow）...")
    resolver = make_resolver(collection_key)
    n = inject_docx(str(input_path), str(output_path), resolver, style,
                    plain=plain, biblio=biblio)
    print(f"完成，写入 {output_path}")
    if plain:
        print("这是纯文本预览版；正式引用请用非 --plain 输出并在 Word+Zotero 中打开。")
    else:
        print("在 Word 中打开后，点 Zotero 选项卡的 Refresh 即可按样式渲染引用与文末参考文献；")
        print("之后增删文献、换样式都直接联动 Zotero。")


if __name__ == "__main__":
    main()
