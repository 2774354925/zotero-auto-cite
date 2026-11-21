---
name: zotero-auto-cite
description: 把「拉取文献 → 导入本机 Zotero → 在 Word 中生成可被 Zotero 识别的活动引用 → 图/表自动题注与交叉引用」串成一条 pipeline。当用户需要 AI 撰写论文/稿件并自动引用文献、按 DOI/PMID/arXiv 标注引用并自动入库 Zotero、批量导入知网(RIS/EndNote)文献、或生成带自动编号的图/表交叉引用 docx 时使用。适用于 Word(docx) 交付、依赖本机 Zotero 桌面端的场景。
---

# Zotero 自动引用 pipeline

把「找文献 → 进 Zotero → 插引用」做成一条命令，用户只需：开头给一次文献（或让 AI 直接按 DOI 写），中间在对话框里选一次引用格式，结尾在 Word 里点一次 Refresh。

## 前置条件（一次性）

1. **Zotero 10+** 已安装并**正在运行**；设置 → 高级 → 勾选 *Allow other applications on this computer to communicate with Zotero*。
2. **Zotero Word 插件**已装（Zotero 安装时通常自动带）。
3. **Python 3.10+** 且 `pip install python-docx`（脚本只用标准库 + python-docx）。
4. 首次写库会弹 **Zotero 授权框**，请点 **Always Allow**（否则每次写都要点）。

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

### 方式 A：一条命令（推荐）

AI 先写 Markdown 草稿（带上面的标记）。**开始前用对话框（interaction_ask）问用户两件事**：① 引用格式（APA / 国标 GB/T 7714 / Nature / Vancouver / IEEE / Chicago author-date）；② 导入到 Zotero 哪个分类（默认 `AI导入`，会自动在「我的文库」下新建子文件夹）。然后：

```
python scripts/pipeline.py draft.md -o 论文.docx --style <样式别名> --collection <分类名>
```

或输入已是 docx：

```
python scripts/pipeline.py draft.docx -o 论文.docx --style gb7714
```

样式别名（`--style` 可接别名或完整 Zotero style URL）：
- `apa` → `http://www.zotero.org/styles/apa`
- `gb7714` → `http://www.zotero.org/styles/china-national-standard-gb-t-7714-2015-numeric`（国标）
- `nature` / `vancouver` / `ieee` / `chicago-author-date`

产出两个 docx：
- `<name>.docx` —— **Zotero 活动引用版**：已内置引用域 + 文末 `References` 参考文献域 + 文档偏好。用**桌面 Word（装了 Zotero 插件）**打开，点 Zotero 选项卡 **Refresh** 即渲染引用和参考文献表；之后增删文献、换样式都联动 Zotero。**不要**在在线预览/飞书文档里打开看这个文件（会显示原始域代码，属正常现象）。
- `<name>_预览.docx` —— **纯文本预览版**：无域代码，任何编辑器都能直接读（引用为 `(作者, 年)` 或 `[1]`，文末是纯文本参考文献表），用于快速阅读或发给不装 Zotero 的人。

可选参数：`--style`、`--collection`、`--no-preview`（不生成预览版）。

### 方式 B：分步（便于排查）

```
python scripts/resolvers.py "10.1038/s41586-020-2649-2"      # 仅测试标识符解析
python scripts/zotero_client.py authorize                     # 单独触发授权
python scripts/zotero_client.py search "关键词"                # 搜库
python scripts/inject_citations.py draft.docx -o out.docx     # 仅做引用注入
python scripts/crossrefs.py draft.docx -o out.docx            # 仅做图表交叉引用
```

## 知网 / 中文文献（重要）

知网无公开 API，无法全自动抓取。可靠做法：在知网勾选文献 → **导出文献 → EndNote(RIS)** → 保存为 `.txt`，然后：

```
python scripts/pipeline.py --ris 知网导出.txt
```

会把整批导入 Zotero（含标题/作者/期刊/年卷期/DOI 等）。之后在正文用 `[@key:XXXX]`（key 由导入命令打印）或按标题搜库找到 key 再引用。

## 关键脚本

- `resolvers.py` — DOI(Crossref)/PMID(PubMed)/arXiv → Zotero 条目 JSON（零依赖）。
- `zotero_client.py` — 本地 API 读/写/授权/查重/RIS 导入（零依赖）。
- `inject_citations.py` — `[@…]` → `ADDIN ZOTERO_ITEM CSL_CITATION` 活动引用域代码。
- `crossrefs.py` — 图/表 SEQ 题注 + REF 交叉引用域代码。
- `md2docx.py` — 极简 Markdown→docx（标题/段落/列表/粗斜体）。
- `pipeline.py` — 一键编排。

## 注意事项 / 排查

- **Zotero 必须运行**且本地 API 已开启，否则脚本报「目标计算机积极拒绝」(127.0.0.1:23119)。
- 引用域代码依赖 Zotero Word 插件；未装插件时引用显示为灰底域但无法刷新。
- 生成的引用默认先写完整 `itemData`，即使库 URI 判定偏差也能正常渲染；若 Refresh 后提示 orphaned（未关联库），见 `references/zotero-local-api.md` 的 URI 说明。
- `md2docx.py` 只支持常见 Markdown（标题/段落/列表/粗体/斜体），复杂排版请先在 Word 里排版，再以 docx 输入跑 pipeline。
- 写请求细节（Server-ID、授权、Write-Token）见 `references/zotero-local-api.md`。
