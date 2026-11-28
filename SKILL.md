---
name: zotero-auto-cite
description: 把「PubMed 选题检索 → 阅读摘要选文献 → 导入本机 Zotero → 在 WPS/Word 稿件中插入可见引文与参考文献」串成一条 pipeline。当用户要求 AI 撰写文章并自动引用 PubMed、DOI 或现有 Zotero 文献，批量导入 RIS，或生成带图/表题注的 docx 时使用。默认生成可直接在 WPS 打开的静态引文版，也可生成 Word 原生活引用版。
---

# Zotero 自动引用 pipeline

把「找文献 → 进 Zotero → 写进稿件」做成一条命令。面向 WPS 时默认生成完整可见的引文和参考文献，不再依赖 WPS 执行 Zotero 的 Word 域。

## 前置条件

1. **Zotero 10+** 已安装并正在运行；设置 → 高级 → 勾选 *Allow other applications on this computer to communicate with Zotero*。
2. **WPS 用户**：社区版 `Zotero.dotm` 已放入 WPS 的 `office6\startup`。本技能不修改该宏；它用于用户之后在 WPS 中手动插入/编辑引文。
3. **Python 3.10+** 且 `python -m pip install python-docx lxml`。
4. 首次写库会弹 **Zotero 授权框**，请点 **Always Allow**。

## 默认决策

- 用户未指定引用格式时：中文文章默认 `gb7714`，英文文章默认 `vancouver`。
- 用户未指定 Zotero 分类时：默认导入 `AI导入`。
- 面向 WPS：**默认 `--target wps`**，最终 docx 中的引文、参考文献、题注都直接可见。
- 用户明确使用 Microsoft Word 或要求 Zotero 活动引用时：加 `--target word`。
- 需要同时交付两种格式：加 `--target both`；主文件为 WPS 版，另生成 `_Zotero活动引用.docx`。

## 正文标记语法

文献引用（写在正文里即可）：

```
[@doi:10.1038/s41586-020-2649-2]
[@pmid:32214244]
[@arxiv:1706.03762]
[@key:ABCD1234]                      # 直接引用库里已有条目（ABCD1234 是 Zotero 条目 key）
[@doi:10.xxxx/yyyy, p. 12]           # 带定位符
```

图/表交叉引用：

```
[图:workflow] 本研究技术路线          # 题注段落（图1/表1 自动编号）
见图[#workflow] 或 (表[#cohort])       # 正文引用（编号自动更新）
```

## 工作流

### 方式 A：PubMed 主题写作（推荐）

当用户说“写一篇肥胖科普，引用 PubMed 文献”这类请求时，不要凭空编造参考文献，按下面流程执行：

1. 将主题改写为英文 PubMed 查询，优先包含人群、干预/暴露和结局。例如：

```bash
python scripts/pubmed_search.py "obesity prevention adults lifestyle intervention" --limit 10
python scripts/pubmed_search.py "obesity prevalence global review" --limit 8 --date-from 2018
```

2. 读取返回的题目、摘要、出版类型和 DOI。挑选 4–8 篇与正文主张直接相关的文献，优先指南、系统综述/Meta 分析、大型队列或 RCT；不要用标题相关但结论不支持正文的材料。
3. 在文章的对应事实句后直接写返回的 `citation_marker`，例如 `[@pmid:30077361]`。文末参考文献由脚本生成，正文不要手写文献表。
4. 运行 WPS 默认流程：

```bash
python scripts/pipeline.py draft.md -o 论文.docx --style gb7714 --collection "AI导入"
```

5. 脚本会把所选 PMID 解析、查重并写入本机 Zotero，然后生成可直接在 WPS 打开的 docx。完成后报告文件路径、引用篇数、Zotero 分类。

用户明确要求在 Microsoft Word 中使用 Zotero 活动引用时：

```bash
python scripts/pipeline.py draft.md -o 论文.docx --target word --style gb7714
```

样式别名（`--style` 可接别名或完整 Zotero style URL）：
- `apa` → `http://www.zotero.org/styles/apa`
- `gb7714` → `http://www.zotero.org/styles/china-national-standard-gb-t-7714-2015-numeric`（国标）
- `nature` / `vancouver` / `ieee` / `chicago-author-date`

WPS 默认只产出一个最终 docx：

- 正文引文是可见文本，如 `[1]` 或 `(Smith et al., 2020)`。
- 参考文献逐条成段显示，不再放在巨型 Word 域里。
- 图/表题注和交叉引用显示为静态编号，不依赖 WPS 执行 `SEQ/REF` 域。
- 文献条目已经真实导入 Zotero；之后仍可在 WPS 中用社区宏手动插入新的 Zotero 引文。

`--target word` 产出 Word 活动引用版，另生成 `_预览.docx`。活动引用版需要 MS Word + Zotero 插件，并在 Word 中执行 Refresh。

可选参数：`--style`、`--collection`、`--target`、`--no-preview`。

### 方式 B：分步（便于排查）

```
python scripts/pubmed_search.py "obesity prevention adults" --limit 8  # PubMed 检索 + 摘要
python scripts/resolvers.py "10.1038/s41586-020-2649-2"      # 仅测试标识符解析
python scripts/zotero_client.py authorize                     # 单独触发授权
python scripts/zotero_client.py search "关键词"                # 搜库
python scripts/inject_citations.py draft.docx -o out.docx --plain  # WPS 可见引用
python scripts/crossrefs.py draft.docx -o out.docx --plain    # WPS 静态图表引用
```

## 知网 / 中文文献（重要）

知网无公开 API，无法全自动抓取。可靠做法：在知网勾选文献 → **导出文献 → EndNote(RIS)** → 保存为 `.txt`，然后：

```
python scripts/pipeline.py --ris 知网导出.txt
```

会把整批导入 Zotero（含标题/作者/期刊/年卷期/DOI 等）。之后在正文用 `[@key:ABCD1234]` 引用（`ABCD1234` 为条目 key 的格式示例，实际 key 由导入命令打印），或按标题搜库找到 key 再引用。

## 关键脚本

- `resolvers.py` — DOI(Crossref)/PMID(PubMed)/arXiv → Zotero 条目 JSON（零依赖）。
- `zotero_client.py` — 本地 API 读/写/授权/查重/RIS 导入（零依赖）。
- `pubmed_search.py` — PubMed 主题检索、摘要和结构化元数据。
- `inject_citations.py` — `[@…]` → WPS 可见引文，或 Word 活动引用域。
- `crossrefs.py` — WPS 静态题注/交叉引用，或 Word 的 SEQ/REF 域。
- `md2docx.py` — 极简 Markdown→docx（标题/段落/列表/粗斜体）。
- `pipeline.py` — 一键编排。

## 注意事项 / 排查

- **Zotero 必须运行**且本地 API 已开启，否则脚本报「目标计算机积极拒绝」(127.0.0.1:23119)。
- **WPS 默认模式不是 Zotero 活动引用**。这是因为 WPS 官方不支持 Zotero Word 域刷新；强行保留域可能只显示原始代码。若必须活动引用，请使用 MS Word + `--target word`。
- 生成的引用默认先写完整 `itemData`，即使库 URI 判定偏差也能正常渲染；若 Refresh 后提示 orphaned（未关联库），见 `references/zotero-local-api.md` 的 URI 说明。
- `md2docx.py` 只支持常见 Markdown（标题/段落/列表/粗体/斜体），复杂排版可先以 docx 输入跑 pipeline。
- 写请求细节（Server-ID、授权、Write-Token）见 `references/zotero-local-api.md`。
