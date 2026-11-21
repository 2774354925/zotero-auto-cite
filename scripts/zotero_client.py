#!/usr/bin/env python3
"""
zotero_client.py — 访问本机 Zotero 本地 API（http://127.0.0.1:23119/api）。

Zotero 10+ 特性：
- 读请求免鉴权；
- 写请求需先 POST /api/local/authorize 弹出授权框（Allow / Always Allow / Deny），
  拿到 32 位本地 API key 后作为 Zotero-API-Key 头随写请求发送；
- 写请求需携带 Zotero-Server-ID 头（从任意响应读取并缓存）；
- 浏览器样式的请求会被丢弃，需带 Zotero-Allowed-Request 头。

零第三方依赖（标准库 urllib）。可独立测试：
    python zotero_client.py authorize
    python zotero_client.py search "cancer"
    python zotero_client.py create "10.1038/s41586-020-2649-2"
"""
from __future__ import annotations

import json
import sys
import urllib.error
import urllib.request
from typing import Optional

BASE = "http://127.0.0.1:23119/api"
APP_NAME = "zotero-auto-cite"
UA = "zotero-auto-cite/1.0"

# RIS TY -> Zotero itemType（覆盖 CNKI / 知网导出常见类型）
RIS_TYPE_MAP = {
    "JOUR": "journalArticle",
    "JFULL": "journalArticle",
    "BOOK": "book",
    "CHAP": "bookSection",
    "CONF": "conferencePaper",
    "CPAPER": "conferencePaper",
    "THES": "thesis",
    "RPRT": "report",
    "ELEC": "webpage",
    "NEWS": "newspaperArticle",
    "MGZN": "magazineArticle",
    "PAT": "patent",
    "STD": "standard",
    "GEN": "document",
    "COMP": "computerProgram",
    "DATA": "dataset",
}

_RIS_FIELD_MAP = {
    "TI": "title",
    "T1": "title",
    "JO": "publicationTitle",
    "JF": "publicationTitle",
    "JA": "publicationTitle",
    "T2": "publicationTitle",
    "VL": "volume",
    "IS": "issue",
    "SP": "pages_start",
    "EP": "pages_end",
    "PY": "date",
    "Y1": "date",
    "DA": "date",
    "DO": "DOI",
    "SN": "ISSN",
    "AB": "abstractNote",
    "N1": "note",
    "UR": "url",
    "PB": "publisher",
    "CY": "place",
    "ET": "edition",
}


class ZoteroError(Exception):
    def __init__(self, message: str, status: Optional[int] = None):
        super().__init__(message)
        self.status = status


class ZoteroClient:
    def __init__(self, base: str = BASE, app_name: str = APP_NAME):
        self.base = base.rstrip("/")
        self.app_name = app_name
        self._server_id: Optional[str] = None
        self._api_key: Optional[str] = None
        self._last_version: Optional[int] = None

    # ---- 底层请求 ----
    def _req(self, method: str, path: str, data: Optional[dict] = None,
             extra_headers: Optional[dict] = None, expect_json: bool = True):
        url = f"{self.base}{path}"
        headers = {
            "User-Agent": UA,
            "Zotero-Allowed-Request": "true",
            "Accept": "application/json",
        }
        if self._server_id:
            headers["Zotero-Server-ID"] = self._server_id
        if self._api_key:
            headers["Zotero-API-Key"] = self._api_key
        if extra_headers:
            headers.update(extra_headers)

        body = None
        if data is not None:
            body = json.dumps(data, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"

        req = urllib.request.Request(url, data=body, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                self._remember_headers(resp)
                raw = resp.read().decode("utf-8")
                if expect_json and raw:
                    return json.loads(raw)
                return raw
        except urllib.error.HTTPError as e:
            detail = ""
            try:
                detail = e.read().decode("utf-8")
            except Exception:
                pass
            raise ZoteroError(f"{method} {path} -> HTTP {e.code}: {detail}", status=e.code) from e

    def _remember_headers(self, resp):
        sid = resp.headers.get("Zotero-Server-ID")
        if sid and not self._server_id:
            self._server_id = sid
        lmv = resp.headers.get("Last-Modified-Version")
        if lmv:
            try:
                self._last_version = int(lmv)
            except ValueError:
                pass

    # ---- 授权（写操作前置）----
    def authorize(self, force: bool = False) -> str:
        """弹出 Zotero 授权框，返回本地 API key（用户选 Always Allow 则可复用）。"""
        if self._api_key and not force:
            return self._api_key
        # 先读一次任意响应拿到 server-id
        if not self._server_id:
            self._req("GET", "/", expect_json=False)
        resp = self._req("POST", "/local/authorize", data={"appName": self.app_name})
        key = resp.get("key")
        if not key:
            raise ZoteroError(f"未获得授权 key: {resp}")
        self._api_key = key
        return key

    # ---- 读 ----
    def search(self, query: str, limit: int = 10, qmode: str = "everything") -> list[dict]:
        import urllib.parse
        q = urllib.parse.quote(query)
        return self._req("GET", f"/users/0/items?q={q}&qmode={qmode}&limit={limit}")

    def get_item(self, key: str) -> dict:
        return self._req("GET", f"/users/0/items/{key}")

    def get_item_csl(self, key: str) -> dict:
        """返回该条目的 CSL-JSON（用于生成引文域代码）。"""
        import urllib.parse
        q = urllib.parse.quote(key)
        resp = self._req("GET", f"/users/0/items?itemKey={q}&format=csljson")
        if isinstance(resp, list):
            if not resp:
                raise ZoteroError(f"未找到条目 {key}")
            return resp[0]
        return resp

    def library_uri(self) -> str:
        """探测本机用户库 URI 前缀（用于域代码 uris）。"""
        items = self._req("GET", "/users/0/items?limit=1")
        if items:
            it = items[0] if isinstance(items, list) else items
            lib = it.get("library") or {}
            lib_id = lib.get("id")
            if lib_id is not None:
                return f"http://zotero.org/users/{lib_id}"
        return "http://zotero.org/users/0"

    # ---- 写 ----
    def _ensure_write_headers(self) -> dict:
        if not self._server_id:
            self._req("GET", "/", expect_json=False)
        if not self._api_key:
            self.authorize()
        return {}

    def create_item(self, item: dict, collection_key: Optional[str] = None) -> dict:
        """创建单个条目，返回含 key 的完整条目。"""
        self._ensure_write_headers()
        # 本地 API 不提供 /items/new 模板接口，直接用完整字段 POST；
        # 写请求需 If-Unmodified-Since-Version 或 Zotero-Write-Token，这里用库版本。
        if self._last_version is None:
            self._req("GET", "/users/0/items?limit=1")
        payload = [dict(item)]
        if collection_key:
            payload[0]["collections"] = [collection_key]
        headers = {}
        if self._last_version is not None:
            headers["If-Unmodified-Since-Version"] = str(self._last_version)
        try:
            resp = self._req("POST", "/users/0/items", data=payload, extra_headers=headers)
        except ZoteroError as e:
            if e.status != 401:
                raise
            # 单次授权 key 被消费：重新授权并重试一次
            self.authorize(force=True)
            resp = self._req("POST", "/users/0/items", data=payload, extra_headers=headers)
        created = None
        for field in ("successful", "success"):
            bucket = resp.get(field) or {}
            if isinstance(bucket, dict) and "0" in bucket:
                created = bucket["0"]
                break
        if created is None:
            raise ZoteroError(f"创建条目失败: {json.dumps(resp, ensure_ascii=False)[:400]}")
        return created

    def find_by_doi(self, doi: str) -> Optional[dict]:
        """按 DOI 查重：找到即返回条目，否则 None。"""
        try:
            hits = self.search(doi, limit=5)
        except ZoteroError:
            return None
        for h in hits:
            if str(h.get("data", {}).get("DOI", "")).lower() == doi.lower():
                return h
        return None

    # ---- 分类文件夹 ----
    def list_collections(self) -> list[dict]:
        try:
            return self._req("GET", "/users/0/collections?limit=100")
        except ZoteroError:
            return []

    def find_collection(self, name: str) -> Optional[dict]:
        want = (name or "").strip()
        for c in self.list_collections():
            if ((c.get("data", {}).get("name")) or "").strip() == want:
                return c
        return None

    def create_collection(self, name: str, parent_key: Optional[str] = None) -> dict:
        """在「我的文库」根目录（parent_key=None）或指定父分类下新建子文件夹。"""
        self._ensure_write_headers()
        if self._last_version is None:
            self._req("GET", "/users/0/collections?limit=1")
        payload = [{"name": (name or "").strip(),
                    "parentCollection": parent_key if parent_key else False}]
        headers = {}
        if self._last_version is not None:
            headers["If-Unmodified-Since-Version"] = str(self._last_version)
        try:
            resp = self._req("POST", "/users/0/collections", data=payload, extra_headers=headers)
        except ZoteroError as e:
            if e.status != 401:
                raise
            self.authorize(force=True)
            resp = self._req("POST", "/users/0/collections", data=payload, extra_headers=headers)
        created = None
        for field in ("successful", "success"):
            bucket = resp.get(field) or {}
            if isinstance(bucket, dict) and "0" in bucket:
                created = bucket["0"]
                break
        if created is None:
            # 本地 API 有时直接返回对象/数组
            if isinstance(resp, dict) and resp.get("key"):
                created = resp
            elif isinstance(resp, list) and resp:
                created = resp[0]
        if created is None:
            raise ZoteroError(f"创建分类失败: {json.dumps(resp, ensure_ascii=False)[:400]}")
        return created

    def ensure_collection(self, name: str, parent_key: Optional[str] = None) -> str:
        """查找或新建分类，返回其 key（幂等）。"""
        found = self.find_collection(name)
        if found:
            return found["key"]
        return self.create_collection(name, parent_key)["key"]

    def add_to_collection(self, collection_key: str, item_key: str) -> dict:
        """把已有条目加入分类（保留原有分类，追加目标分类）。"""
        self._ensure_write_headers()
        item = self.get_item(item_key)
        cur = list(item.get("data", {}).get("collections") or [])
        if collection_key in cur:
            return item
        cur.append(collection_key)
        headers = {}
        if self._last_version is not None:
            headers["If-Unmodified-Since-Version"] = str(self._last_version)
        payload = {"collections": cur}
        try:
            return self._req("PATCH", f"/users/0/items/{item_key}", data=payload, extra_headers=headers)
        except ZoteroError as e:
            if e.status != 401:
                raise
            self.authorize(force=True)
            return self._req("PATCH", f"/users/0/items/{item_key}", data=payload, extra_headers=headers)

    def import_ris(self, ris_text: str, collection_key: Optional[str] = None) -> list[dict]:
        """解析 RIS（含知网导出的 EndNote/RIS 文本）并批量创建，返回创建后的条目列表。"""
        records = self._parse_ris(ris_text)
        created = []
        for rec in records:
            created.append(self.create_item(rec, collection_key))
        return created

    @staticmethod
    def _parse_ris(ris_text: str) -> list[dict]:
        records: list[dict] = []
        cur: Optional[dict] = None
        for raw in ris_text.splitlines():
            line = raw.rstrip("\r")
            if not line.strip():
                continue
            tag = line[:2].upper()
            value = line[6:] if len(line) > 6 else ""
            value = value.strip()
            if tag == "TY":
                cur = {"itemType": RIS_TYPE_MAP.get(value, "document"),
                       "creators": [], "_pages_start": None, "_pages_end": None}
            elif cur is None:
                continue
            elif tag == "AU":
                cur["creators"].append(_ris_author(value))
            elif tag == "ER":
                if cur is not None:
                    _finalize_ris(cur)
                    records.append(cur)
                cur = None
            elif tag in _RIS_FIELD_MAP:
                key = _RIS_FIELD_MAP[tag]
                if key in ("pages_start", "pages_end"):
                    cur["_" + key] = value
                else:
                    cur[key] = value
        if cur is not None:
            _finalize_ris(cur)
            records.append(cur)
        return [r for r in records if r.get("title")]


def _ris_author(value: str) -> dict:
    value = value.strip()
    if "," in value:
        last, first = value.split(",", 1)
        return {"creatorType": "author", "lastName": last.strip(), "firstName": first.strip()}
    parts = value.split()
    if len(parts) == 1:
        return {"creatorType": "author", "name": parts[0]}
    return {"creatorType": "author", "lastName": parts[0], "firstName": " ".join(parts[1:])}


def _finalize_ris(cur: dict):
    sp, ep = cur.pop("_pages_start", None), cur.pop("_pages_end", None)
    if sp or ep:
        cur["pages"] = f"{sp or ''}-{ep or ''}"
    if "date" in cur and isinstance(cur["date"], str):
        # 只取年份，兼容 '2020' / '2020/01/01' / '2020-01'
        import re
        m = re.search(r"\b(\d{4})\b", cur["date"])
        if m:
            cur["date"] = m.group(1)
    if not cur.get("creators"):
        cur.pop("creators", None)


def _main():
    client = ZoteroClient()
    if len(sys.argv) < 2:
        print("用法: python zotero_client.py <authorize|search|create|ris> [参数]")
        return
    cmd = sys.argv[1]
    if cmd == "authorize":
        key = client.authorize(force=True)
        print(f"已授权，key={key}")
    elif cmd == "search":
        q = sys.argv[2] if len(sys.argv) > 2 else ""
        for it in client.search(q):
            print(f"{it.get('key')}\t{it.get('data', {}).get('title', '')[:60]}")
    elif cmd == "create":
        ident = sys.argv[2]
        sys.path.insert(0, __import__("os").path.dirname(__file__))
        from resolvers import resolve
        created = client.create_item(resolve(ident))
        print(json.dumps({"key": created.get("key"), "title": created.get("data", {}).get("title")},
                         ensure_ascii=False))
    elif cmd == "ris":
        path = sys.argv[2]
        with open(path, encoding="utf-8") as f:
            created = client.import_ris(f.read())
        print(f"导入 {len(created)} 条")


if __name__ == "__main__":
    _main()
