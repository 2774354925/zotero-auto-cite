# zotero-auto-cite

本项目是一个面向本机 Zotero 10 的文献写作工具。它可以读取 Markdown
或 DOCX 草稿中的引用标记，自动解析 DOI、PMID 和 arXiv ID，把文献写入
本机 Zotero，再在正文中生成引文、参考文献和图/表交叉引用。

默认生成 WPS 可直接打开的静态引文版，正文、参考文献和图表编号都是
可见文本。需要 Zotero 活动引用时，也可以生成供 Microsoft Word 使用的
原生活引用版。

## 功能

- 按主题检索 PubMed，并读取题目、摘要、作者、DOI、PMID 和出版类型。
- 从 PubMed、Crossref 和 arXiv 获取文献元数据。
- 按 DOI 查重，避免重复导入 Zotero。
- 将文献自动归入指定 Zotero 分类，默认 `AI导入`。
- 支持 GB/T 7714、APA、Vancouver、Nature、IEEE、Chicago author-date。
- WPS 版直接写出可见引文和逐条参考文献。
- Word 版支持 Zotero 活动引用、Refresh、文末参考文献域。
- 支持图/表题注与正文交叉引用。
- 支持知网/RIS 文献批量导入。

## 环境

- 桌面版 Zotero 10 或更高版本，并保持运行。
- 在 Zotero 的“设置 → 高级”中勾选
  *Allow other applications on this computer to communicate with Zotero*。
- Python 3.10 或更高版本。
- 安装 Python 依赖：`python -m pip install python-docx lxml`
- WPS 用户可继续保留社区版 `Zotero.dotm`，用于手动插入/编辑引文。

## 安装

```bash
git clone https://github.com/2774354925/zotero-auto-cite.git
cd zotero-auto-cite
python -m pip install python-docx lxml
```

## 最简流程

### 1. 检索 PubMed

```bash
python scripts/pubmed_search.py "obesity prevention adults lifestyle intervention" --limit 10
```

返回结果中包含每篇文献的摘要、PMID 和可直接写进正文的引用标记：

```text
[@pmid:30077361]
```

### 2. 写 Markdown 草稿

在对应论点句后写引用标记：

```markdown
肥胖是一种由遗传、环境、行为和代谢因素共同作用的慢性疾病
[@pmid:30077361]。儿童和青少年肥胖还与多种代谢和心血管风险相关
[@pmid:38856986]。
```

图表写法：

```text
[图:workflow] 本研究技术路线
见图[#workflow]。
```

### 3. 生成 WPS 版本

```bash
python scripts/pipeline.py draft.md -o 论文.docx --style gb7714 --collection "AI导入"
```

程序会：

1. 解析 PMID/DOI/arXiv。
2. 查重并导入本机 Zotero。
3. 将条目加入 Zotero 分类。
4. 在 docx 中写出可见引文和参考文献。
5. 生成 WPS 可直接打开的最终文件。

首次写库时 Zotero 会弹授权框，选 **Always Allow**。

可以使用仓库中的示例直接体验完整流程：

```bash
python scripts/pipeline.py examples/wps_obesity_demo.md -o 肥胖科普.docx --style gb7714 --collection "AI导入"
```

## 输出模式

### WPS（默认）

```bash
python scripts/pipeline.py draft.md -o 论文.docx --target wps
```

引文样式示例：

- GB/T 7714：正文 `[1]`，文末 `[1] 作者. 标题[J]. 期刊, 年, 卷(期): 页码. DOI...`
- APA：正文 `(Smith et al., 2020)`，文末 APA 风格条目。

WPS 版不包含 `ADDIN ZOTERO_ITEM`、`ADDIN ZOTERO_BIBL`、`SEQ`、`REF` 域，因此不会因为 WPS 无法刷新 Zotero 宏而只显示几行网址或原始代码。

### Word

```bash
python scripts/pipeline.py draft.md -o 论文.docx --target word
```

用 Microsoft Word + Zotero 插件打开，点 Zotero 选项卡中的 **Refresh**。

### 同时生成

```bash
python scripts/pipeline.py draft.md -o 论文.docx --target both
```

主文件是 WPS 版，另生成 `论文_Zotero活动引用.docx`。

## 引用标记

| 标记 | 含义 |
| --- | --- |
| `[@pmid:30077361]` | 按 PubMed PMID 拉取、入库并引用 |
| `[@doi:10.1016/j.cell.2007.11.019]` | 按 DOI 拉取、入库并引用 |
| `[@arxiv:1706.03762]` | 按 arXiv ID 拉取、入库并引用 |
| `[@key:ABCD1234]` | 引用 Zotero 中已有条目 |
| `[@pmid:30077361, p. 12]` | 带页码或其他定位符 |

`[@key:ABCD1234]` 中的 `ABCD1234` 是 Zotero 条目 key 的格式示例。实际 key
可以通过导入命令的输出或 Zotero 条目的 URI 获取。

## 引用样式

| 别名 | 样式 |
| --- | --- |
| `gb7714` | GB/T 7714-2015 numeric |
| `apa` | APA 7th |
| `vancouver` | Vancouver |
| `nature` | Nature |
| `ieee` | IEEE |
| `chicago-author-date` | Chicago author-date |

也可直接传 Zotero CSL style URL。

## 知网/RIS

知网不提供稳定的公开抓取 API。建议在知网导出 EndNote/RIS：

```bash
python scripts/pipeline.py --ris 知网导出.txt --collection "AI导入"
```

导入命令会打印 Zotero 条目 key，随后可用 `[@key:条目key]` 引用。

## 分步命令

```bash
python scripts/pubmed_search.py "obesity prevention" --limit 8
python scripts/resolvers.py "10.1016/j.cell.2007.11.019"
python scripts/zotero_client.py authorize
python scripts/zotero_client.py search "肥胖"
python scripts/inject_citations.py draft.docx -o out.docx --plain
python scripts/crossrefs.py draft.docx -o out.docx --plain
python scripts/test_offline.py
```

## 故障排查

**Zotero 本地 API 连不上**

先启动 Zotero，并确认 127.0.0.1:23119 可访问。

**WPS 只显示几行原始代码**

不要使用旧的 Zotero 域版。用 `--target wps` 重新生成，输出中不应出现 `ADDIN ZOTERO_ITEM`。

**首次导入没有反应**

查看 Zotero 是否弹出授权框，选择 `Always Allow`。

**PubMed 检索失败**

确认网络或代理能访问 `eutils.ncbi.nlm.nih.gov`。

**希望 WPS 中引文自动编号联动**

WPS 官方不支持 Zotero 的 Word 域刷新。当前方案保证 WPS 文档完整可读和编号正确；新增或删除文献后需要重新运行 pipeline 生成。

## License

本项目使用 MIT License，详见 [`LICENSE`](LICENSE)。
