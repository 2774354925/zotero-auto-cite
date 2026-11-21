# -*- coding: utf-8 -*-
"""离线自测：不依赖 Zotero 运行，验证 md2docx / crossrefs / inject_citations 的 XML 生成。"""
import sys, zipfile, tempfile, os
from pathlib import Path

SKILL_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(SKILL_DIR))

from docx import Document

def read_docx_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")

ok = True
def check(name, cond, extra=""):
    global ok
    print(("PASS " if cond else "FAIL ") + name + (("  " + extra) if extra else ""))
    if not cond:
        ok = False

tmp = Path(tempfile.mkdtemp(prefix="za_test_"))

# ---- 1. md2docx ----
import md2docx
md = "# 标题\n\n引言，见**图**[#workflow]。\n\n[图:workflow] 本研究技术路线\n\n正文 *斜体* 如(表[#cohort])所示。\n\n[表:cohort] 患者基线特征\n"
docx1 = tmp / "draft.docx"
md2docx.convert(md).save(str(docx1))
check("md2docx 生成", docx1.exists())

# ---- 2. crossrefs ----
import crossrefs
docx2 = tmp / "crossref.docx"
r = crossrefs.process_docx(str(docx1), str(docx2))
x = read_docx_xml(docx2)
check("crossrefs 题注数=2", r["captions"] == 2, str(r["captions"]))
check("crossrefs 引用数=2", r["refs"] == 2, str(r["refs"]))
check("crossrefs 含 SEQ", "SEQ" in x)
check("crossrefs 含 REF", "REF" in x)
check("crossrefs 含书签", "bookmarkStart" in x)

# ---- 3. inject_citations（假解析器）----
import inject_citations as ic

def fake_resolve(kind, value):
    csl = {
        "type": "article-journal",
        "title": "Array programming with NumPy",
        "author": [{"family": "Harris", "given": "Charles R."}],
        "issued": {"date-parts": [[2020, 9, 16]]},
        "DOI": "10.1038/s41586-020-2649-2",
        "container-title": "Nature",
        "volume": "585",
        "page": "357-362",
    }
    return {"key": "ABCD1234", "csl": csl, "uri": "http://zotero.org/users/1/items/ABCD1234"}

d = Document()
d.add_paragraph("Prior work [@doi:10.1038/s41586-020-2649-2] shows X.")
d.add_paragraph("See [@key:ABCD1234, p. 12] for details.")
docx3 = tmp / "cite_in.docx"
d.save(str(docx3))
docx4 = tmp / "cite_out.docx"
n = ic.inject_docx(str(docx3), str(docx4), fake_resolve, "http://www.zotero.org/styles/apa")
x = read_docx_xml(docx4)
check("inject 替换数=2", n == 2, str(n))
check("inject 含 ZOTERO_ITEM", "ADDIN ZOTERO_ITEM CSL_CITATION" in x)
check("inject 含 ZOTERO_PREF_1", "ADDIN ZOTERO_PREF_1" in x)
check("inject 占位文本", "(Harris, 2020)" in x)

# 确认生成的 docx 能被 python-docx 重新打开（XML 结构合法）
try:
    Document(str(docx4))
    check("输出 docx 可重新打开", True)
except Exception as e:
    check("输出 docx 可重新打开", False, str(e))

print("\n全部通过" if ok else "\n存在失败")
sys.exit(0 if ok else 1)
