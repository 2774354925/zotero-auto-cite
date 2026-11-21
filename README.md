# zotero-auto-cite

把「找文献 → 导入本机 Zotero → 插入可刷新的活动引用 → 图/表自动题注与交叉引用」串成一条命令的本地自动化工具。

让 AI Agent（或你自己）写完论文草稿后，**文献自动入库、自动引用、自动编号**，而不是给你一堆「死文本」让你手动搬进 Zotero 再一篇篇插进 Word。

---

## 特性

- **多标识符解析**：`DOI`（Crossref）、`PMID`（PubMed）、`arXiv` 自动拉取元数据（标题/作者/期刊/年卷期/摘要等）。
- **自动入库 Zotero**：写入本机 Zotero 本地 API（Zotero 10+），自动按 DOI 查重，自动归入指定分类（默认「AI导入」）。
- **活动引用**：生成 Zotero 官方 `ADDIN ZOTERO_ITEM CSL_CITATION` 域代码，在 Word 中点 **Refresh** 即渲染，可增删、可换样式、与库双向联动。
- **参考文献表自动生成**：文末自动插入 `ADDIN ZOTERO_BIBL` 参考文献域，无需手动 Add/Edit Bibliography。
- **图/表交叉引用**：`[图:id]` / `[表:id]` 自动编号，`[#id]` 自动交叉引用；编号由 `SEQ`/`REF` 域驱动，`Ctrl+A` + `F9` 一键更新。
- **知网中文文献**：支持「知网导出 EndNote(RIS) → 一条命令批量导入」。
- **纯文本预览版**：额外生成无域代码的 `_预览.docx`，任何编辑器都能直接读。
- **零重依赖**：仅标准库 + `python-docx`，无 pandoc / Better BibTeX / citeproc 依赖。

---

## 环境要求

- **Zotero 10+**（桌面版），并在「设置 → 高级」勾选 *Allow other applications on this computer to communicate with Zotero*。
- **Zotero Word 插件**（Zotero 安装时通常自动安装）。
- **Python 3.10+**，执行 `pip install python-docx`。

> 系统不限；已在 Windows 10/11 + Zotero 10 上实测通过。macOS/Linux 原理一致（本地 API 端口同为 `127.0.0.1:23119`），路径按系统调整即可。

---

## 安装

```bash
git clone <你的仓库地址>
cd zotero-auto-cite
pip install python-docx
```

无需其他依赖。

---

## 快速开始

**第 1 步：写草稿，标注引用。**

正文里这样标：

```text
[@doi:10.1038/s41586-020-2649-2]
[@pmid:32214244]
[@arxiv:1706.03762]
[@key:ABCD1234]                 # 引用库里已有条目（ABCD1234 为条目 key）
[@doi:10.xxxx/yyyy, p. 12]      # 带页码定位
```

图/表这样标：

```text
[图:workflow] 本研究的技术路线
见图[#workflow]、如表[#cohort]所示
[表:cohort] 入组患者基线特征
```

**第 2 步：运行一条命令。**

```bash
python scripts/pipeline.py draft.md -o 论文.docx --style gb7714 --collection "AI导入"
```

首次写库会弹 Zotero 授权框，请选 **Always Allow（总是允许）**。

**第 3 步：桌面 Word 打开 → Zotero 选项卡 → Refresh。** 完成。

仓库内 `examples/` 提供了两个可直接运行的示例：

```bash
python scripts/pipeline.py examples/hipsc_opinion.md -o out.docx --style vancouver
python scripts/pipeline.py examples/crossref_demo.md -o out.docx
```

---

## 正文标记语法

| 标记 | 含义 |
| --- | --- |
| `[@doi:10.xxxx/yyyy]` | 按 DOI 拉取并引用 |
| `[@pmid:32214244]` | 按 PubMed ID 拉取并引用 |
| `[@arxiv:1706.03762]` | 按 arXiv ID 拉取并引用 |
| `[@key:ABCD1234]` | 直接引用库里已有条目 |
| `[@doi:…, p. 12]` | 带定位符（页码/章节） |
| `[图:fig1] 题注` | 图题注（自动编号 图1…） |
| `[表:tab1] 题注` | 表题注（自动编号 表1…） |
| `[#fig1]` | 正文交叉引用（指向对应题注） |

---

## 命令参考

### `pipeline.py` —— 一键编排（推荐）

```bash
python scripts/pipeline.py <draft.md|draft.docx> -o <out.docx> [选项]
```

| 选项 | 说明 |
| --- | --- |
| `--style` | 引用样式（别名或完整 Zotero style URL），见下表 |
| `--collection` | 导入到 Zotero 的分类名，默认 `AI导入`（自动创建） |
| `--no-preview` | 不生成纯文本预览版 |
| `--ris 知网导出.txt` | 仅批量导入知网 RIS（不处理正文） |

样式别名：

| 别名 | 对应样式 |
| --- | --- |
| `apa` | APA 7th |
| `gb7714` | 国标 GB/T 7714（numeric） |
| `nature` | Nature |
| `vancouver` | Vancouver |
| `ieee` | IEEE |
| `chicago-author-date` | Chicago author-date |

也可直接传完整 URL，如 `--style "http://www.zotero.org/styles/apa"`。

### 分步脚本（便于排查）

```bash
python scripts/resolvers.py "10.1038/s41586-020-2649-2"   # 测试标识符解析
python scripts/zotero_client.py authorize                  # 单独触发授权
python scripts/zotero_client.py search "关键词"             # 搜索库
python scripts/inject_citations.py draft.docx -o out.docx  # 仅注入引用
python scripts/crossrefs.py draft.docx -o out.docx         # 仅图/表交叉引用
python scripts/md2docx.py draft.md -o draft.docx           # 仅 Markdown→docx
python scripts/test_offline.py                             # 离线自测（不依赖 Zotero）
```

---

## 输出文件

- `<name>.docx` —— **Zotero 活动引用版**：内置引用域 + 文末 References 域 + 文档偏好。用桌面 Word（装了 Zotero 插件）打开，点 **Refresh** 渲染。
- `<name>_预览.docx` —— **纯文本预览版**：无域代码，任何编辑器都能直接读。

---

## 目录结构

```
zotero-auto-cite/
├── README.md
├── LICENSE
├── .gitignore
├── SKILL.md                        # 豆包技能元数据（供豆包用户直接安装）
├── references/
│   └── zotero-local-api.md         # Zotero 本地 API 踩坑记录（Server-ID / 授权 / 版本头）
├── scripts/
│   ├── pipeline.py                 # 一键编排
│   ├── resolvers.py                # DOI/PMID/arXiv → Zotero 条目（零依赖）
│   ├── zotero_client.py            # 本地 API 读写/授权/查重/分类/RIS 导入（零依赖）
│   ├── inject_citations.py         # [@…] → ZOTERO 活动引用域 + 参考文献域
│   ├── crossrefs.py                # 图/表 SEQ 题注 + REF 交叉引用
│   ├── md2docx.py                  # 极简 Markdown → docx
│   └── test_offline.py             # 离线自测
└── examples/
    ├── hipsc_opinion.md            # 带 4 篇文献引用的观点短文
    └── crossref_demo.md            # 图/表交叉引用演示
```

---

## 工作原理

1. **解析**：`resolvers.py` 调用 Crossref / PubMed / arXiv 的开放接口，把标识符解析成 Zotero 条目结构。
2. **入库**：`zotero_client.py` 通过 Zotero 本地 API（`http://127.0.0.1:23119/api/`）写入条目；写请求带 `Zotero-Server-ID` + `Zotero-API-Key` + `If-Unmodified-Since-Version`，401 时自动重新授权重试。
3. **注入**：`inject_citations.py` 把正文标记替换为 Word 字段域代码，并注入文档级 `ZOTERO_PREF_1`（样式/文档偏好）与文末 `ZOTERO_BIBL`。
4. **交叉引用**：`crossrefs.py` 生成 `SEQ`（自动编号）+ `REF`（交叉引用）+ 书签。

细节见 `references/zotero-local-api.md`。

---

## FAQ

**Q：为什么在在线预览/飞书里打开 docx 是一堆 `ADDIN ZOTERO_…` 代码？**
域代码是「活动引用」的必要载体，只有桌面 Word + Zotero 插件能折叠渲染。想直接读内容，用 `_预览.docx`。

**Q：WPS 能用吗？**
Zotero 官方不原生支持 WPS。社区做法是把 `Zotero.dotm` 复制进 WPS 启动目录并开启宏（见 Zotero 中文社区《在 WPS 中插入参考文献》）。本工具生成的 docx 用的是同一套域代码，WPS 里 Zotero 插件接好后即可使用；但因用户环境（Zotero/WPS 版本、VBA 组件、宏安全级别）差异大，**不承诺 WPS 开箱即用**。

**Q：首次授权点错了（选了 Allow 而不是 Always Allow）？**
本工具内置「写请求 401 → 重新授权重试」，单次授权也能跑完，只是每次写都要点一次。建议在 Zotero 里重新触发并选 Always Allow。

**Q：知网能全自动抓吗？**
不能。知网无公开 API。请用「知网 → 导出文献 → EndNote(RIS)」保存为 `.txt`，再 `python scripts/pipeline.py --ris 导出.txt` 批量入库，然后用 `[@key:XXXX]` 引用。

**Q：会不会重复导入同一篇文献？**
会先按 DOI 查重，已存在的直接复用并归入分类，不重复建条目。

---

## 开发与贡献

```bash
python scripts/test_offline.py    # 不依赖 Zotero 的离线回归测试
```

欢迎提交 issue / PR。已知改进方向：更友好的交互、更多内置样式、WPS 环境探测、知网抓取（如政策允许）、导出到 .ris/.bib。

---

## License

见 `LICENSE` 文件（默认 MIT，发布前请按需调整）。
