# Zotero 10 本地 API 写操作要点（排错参考）

来源：Zotero 官方文档（本地 API / 写请求 / Zotero 10 for Developers）。

## 基础

- 基址 `http://127.0.0.1:23119/api/`，仅 API v3。
- 读请求免鉴权；写请求（Zotero 10+）需本地 API key。
- Zotero 未开本地 API 时返回 `403`；端口没开=「连接被拒绝」。

## 安全头（Zotero 10 起）

- 请求必须带 `Host: localhost/127.0.0.1/[::1]`，否则 `400`。
- 浏览器样式的请求（`User-Agent` 以 `Mozilla/` 开头，或带 `Origin` 头）会被丢弃，
  除非带 `Zotero-Allowed-Request: true` 头。本客户端始终带该头。

## Server-ID 与版本

- 每个响应带 `Zotero-Server-ID` 头；从任意响应（`GET /api/` 即可）读取并缓存。
- 写请求（含 `POST /api/local/authorize`）**必须**带 `Zotero-Server-ID` 请求头，否则 `428`。
- 版本为本地版本，与 Web API 版本无关；写对象可用 `If-Unmodified-Since-Version` 做前置条件。

## 授权流程（Authorizing Writes）

```
POST /api/local/authorize
Content-Type: application/json
Zotero-Server-ID: <serverID>
{"appName": "zotero-auto-cite"}
```

- Zotero 弹框：Allow / Always Allow / Deny。
- 通过返回 `{"key": "<32位>", "remember": bool}`；Always Allow 时 `remember=true` 可复用，
  否则 key 一次性（首个写请求消费后需重新授权）。
- 拒绝返回 `403 {"denied": true}`。
- 每分钟最多 5 次弹框请求，超限 `429` + `Retry-After`。
- 后续写请求带 `Zotero-API-Key: <key>`（或 `Authorization: Bearer <key>`）。
- 无 key/错误 key 返回 `401` + `WWW-Authenticate: Zotero-API-Key realm="Zotero Local API"`。
- 用户可在 设置→高级 里「Clear Write Authorizations」撤销所有已记住授权。

## 创建条目

```
GET  /users/0/items/new?itemType=journalArticle   # 空模板
POST /users/0/items                                # body = [模板+字段]
```

- 需要 `Zotero-Write-Token`（内存缓存，重启即忘）或 `If-Unmodified-Since-Version` 头。
- 响应 `{"successful": {"0": {...含 key...}}, "failed": {...}}`。

## 导出 CSL-JSON（用于引文域代码的 itemData）

```
GET /users/0/items?itemKey=<KEY>&format=csljson
```

返回 CSL JSON（`type`、`author`、`issued`、`container-title` 等）。

## 引文域代码结构（ADDIN ZOTERO_ITEM）

复数字段由 fldChar begin/instrText/separate/可见文本/end 组成：

- `instrText` = ` ADDIN ZOTERO_ITEM CSL_CITATION {json} `，其中 json：
  - `citationID`：任意唯一串
  - `properties.formattedCitation/plainCitation`：可见占位文本（Refresh 后覆盖）
  - `citationItems[].id`：文档内唯一数字 id
  - `citationItems[].uris`：`http://zotero.org/users/<libID>/items/<key>`
  - `citationItems[].itemData`：完整 CSL JSON（`id` 与 citationItems[].id 一致）
- 文档级 `ADDIN ZOTERO_PREF_1 <data>...` 用于免去手动 Document Preferences。

## library URI 判定

域代码 uris 需指向正确库，否则条目「orphaned」（仍能渲染，但无法从库刷新元数据）。
**正确来源是 CSL 导出的 `id` 字段**：未同步到 zotero.org 的本地库，其值为
`http://zotero.org/users/local/<本地库ID>/items/<key>`（本地库 ID 是一串哈希，非数字，
且不等于 Server-ID；已实测）。`inject_citations.py` 的 `make_resolver()` 直接复用 CSL 的 `id`，
**不要**改用条目 JSON 的 `library.id`（对本地库固定为 0）。
