#!/usr/bin/env python3
"""
Search PubMed and return structured metadata plus abstracts for evidence-based
drafting. The output contains PMIDs that can be cited with [@pmid:...].

Examples:
    python pubmed_search.py "obesity prevention adults" --limit 8
    python pubmed_search.py "GLP-1 obesity randomized trial" --limit 6 --date-from 2023
"""
from __future__ import annotations

import json
import re
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from typing import Optional

USER_AGENT = "zotero-auto-cite/1.1 (local research assistant)"
EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"


def _fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=45) as resp:
        return resp.read()


def _clean_text(element: Optional[ET.Element]) -> str:
    if element is None:
        return ""
    return " ".join("".join(element.itertext()).split())


def _inner_text(element: Optional[ET.Element]) -> str:
    if element is None:
        return ""
    return "".join(element.itertext()).strip()


def _first_text(parent: ET.Element, *paths: str) -> str:
    for path in paths:
        value = _clean_text(parent.find(path))
        if value:
            return value
    return ""


def _parse_year(article: ET.Element) -> str:
    medline = article.find("MedlineCitation")
    if medline is None:
        return ""
    value = _first_text(
        medline,
        "Article/Journal/JournalIssue/PubDate/Year",
        "Article/ArticleDate/Year",
        "DateCompleted/Year",
        "DateRevised/Year",
    )
    if value:
        return value
    medline_date = _first_text(medline, "Article/Journal/JournalIssue/PubDate/MedlineDate")
    match = re.search(r"\b(19|20)\d{2}\b", medline_date)
    return match.group(0) if match else ""


def _parse_authors(article: ET.Element) -> list[dict]:
    creators = []
    for author in article.findall(".//AuthorList/Author"):
        collective = _first_text(author, "CollectiveName")
        if collective:
            creators.append({"creatorType": "author", "name": collective})
            continue
        last = _first_text(author, "LastName")
        first = _first_text(author, "ForeName", "Initials")
        if last:
            creators.append({
                "creatorType": "author",
                "lastName": last,
                "firstName": first,
            })
    return creators


def _parse_publication_types(article: ET.Element) -> list[str]:
    return [
        text for text in (
            _clean_text(node)
            for node in article.findall(".//PublicationTypeList/PublicationType")
        ) if text
    ]


def _parse_keywords(article: ET.Element) -> list[str]:
    return [
        text for text in (
            _clean_text(node)
            for node in article.findall(".//KeywordList/Keyword")
        ) if text
    ]


def parse_pubmed_xml(xml_text: str | bytes) -> list[dict]:
    """Parse PubMed efetch XML into stable JSON-serialisable records."""
    if isinstance(xml_text, bytes):
        xml_text = xml_text.decode("utf-8", errors="replace")
    root = ET.fromstring(xml_text)
    records = []
    for pubmed_article in root.findall(".//PubmedArticle"):
        citation = pubmed_article.find("MedlineCitation")
        article = citation.find("Article") if citation is not None else None
        if citation is None or article is None:
            continue

        pmid = _first_text(citation, "PMID", "PMID/PMID")
        if not pmid:
            continue
        title = _first_text(article, "ArticleTitle").rstrip(".")
        journal = _first_text(
            article, "Journal/Title", "Journal/ISOAbbreviation"
        )
        journal_abbrev = _first_text(article, "Journal/ISOAbbreviation")
        abstract_nodes = article.findall("Abstract/AbstractText")
        abstract_parts = []
        for node in abstract_nodes:
            label = (node.get("Label") or "").strip()
            text = _clean_text(node)
            if text:
                abstract_parts.append(f"{label}: {text}" if label else text)
        abstract = "\n".join(abstract_parts)

        doi = ""
        for node in pubmed_article.findall(
                "PubmedData/ArticleIdList/ArticleId"):
            if (node.get("IdType") or "").lower() == "doi":
                doi = _clean_text(node)
                break

        authors = _parse_authors(article)
        year = _parse_year(pubmed_article)
        volume = _first_text(article, "Journal/JournalIssue/Volume")
        issue = _first_text(article, "Journal/JournalIssue/Issue")
        pages = _first_text(article, "Pagination/MedlinePgn")
        publication_types = _parse_publication_types(article)
        keywords = _parse_keywords(citation)

        zotero_item: dict = {
            "itemType": "journalArticle",
            "title": title,
            "creators": authors,
            "publicationTitle": journal,
            "extra": f"PMID: {pmid}",
            "url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
        }
        if abstract:
            zotero_item["abstractNote"] = abstract
        if year:
            zotero_item["date"] = year
        if volume:
            zotero_item["volume"] = volume
        if issue:
            zotero_item["issue"] = issue
        if pages:
            zotero_item["pages"] = pages
        if doi:
            zotero_item["DOI"] = doi

        records.append({
            "pmid": pmid,
            "title": title,
            "abstract": abstract,
            "authors": [
                " ".join(filter(None, (a.get("firstName"), a.get("lastName"))))
                or a.get("name", "")
                for a in authors
            ],
            "journal": journal,
            "journal_abbrev": journal_abbrev,
            "year": year,
            "volume": volume,
            "issue": issue,
            "pages": pages,
            "doi": doi,
            "url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
            "publication_types": publication_types,
            "keywords": keywords,
            "citation_marker": f"[@pmid:{pmid}]",
            "zotero_item": zotero_item,
        })
    return records


def search_pubmed(query: str, limit: int = 8, date_from: str = "",
                  date_to: str = "", sort: str = "relevance") -> dict:
    if not query.strip():
        raise ValueError("PubMed 查询不能为空")
    limit = max(1, min(int(limit), 100))
    params = {
        "db": "pubmed",
        "term": query,
        "retmode": "json",
        "retmax": str(limit),
        "sort": sort,
    }
    if date_from or date_to:
        params.update({
            "datetype": "pdat",
            "mindate": date_from or "1900",
            "maxdate": date_to or "3000",
        })
    search_url = f"{EUTILS}/esearch.fcgi?" + urllib.parse.urlencode(params)
    payload = json.loads(_fetch(search_url).decode("utf-8"))
    result = payload.get("esearchresult") or {}
    ids = result.get("idlist") or []
    if not ids:
        return {
            "query": query,
            "total_found": int(result.get("count") or 0),
            "returned": 0,
            "results": [],
        }

    fetch_params = {
        "db": "pubmed",
        "id": ",".join(ids),
        "retmode": "xml",
        "rettype": "abstract",
    }
    fetch_url = f"{EUTILS}/efetch.fcgi?" + urllib.parse.urlencode(fetch_params)
    records = parse_pubmed_xml(_fetch(fetch_url))
    by_pmid = {r["pmid"]: r for r in records}
    ordered = [by_pmid[pmid] for pmid in ids if pmid in by_pmid]
    return {
        "query": query,
        "total_found": int(result.get("count") or 0),
        "returned": len(ordered),
        "results": ordered,
    }


def main():
    args = sys.argv[1:]
    query = ""
    limit = 8
    date_from = ""
    date_to = ""
    sort = "relevance"
    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "--limit":
            i += 1
            limit = int(args[i])
        elif arg == "--date-from":
            i += 1
            date_from = args[i]
        elif arg == "--date-to":
            i += 1
            date_to = args[i]
        elif arg == "--sort":
            i += 1
            sort = args[i]
        elif arg in ("-h", "--help"):
            print(__doc__)
            return
        elif not query:
            query = arg
        i += 1
    if not query:
        print(__doc__)
        return
    try:
        result = search_pubmed(query, limit, date_from, date_to, sort)
    except Exception as exc:
        print(json.dumps({
            "error": str(exc),
            "hint": "检查网络、代理和 PubMed E-utilities 是否可访问。",
        }, ensure_ascii=False, indent=2))
        raise SystemExit(1)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
