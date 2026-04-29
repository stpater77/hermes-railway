from __future__ import annotations

import json
import re
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime
from html import unescape
from typing import Any


KNOWN_FEEDS: dict[str, str] = {
    "fox": "https://moxie.foxnews.com/google-publisher/latest.xml",
    "foxnews": "https://moxie.foxnews.com/google-publisher/latest.xml",
    "fox news": "https://moxie.foxnews.com/google-publisher/latest.xml",
    "reuters": "https://www.reutersagency.com/feed/?best-topics=top-news&post_type=best",
    "ap": "https://apnews.com/hub/ap-top-news?output=rss",
    "associated press": "https://apnews.com/hub/ap-top-news?output=rss",
    "bbc": "https://feeds.bbci.co.uk/news/rss.xml",
    "bbc news": "https://feeds.bbci.co.uk/news/rss.xml",
    "npr": "https://feeds.npr.org/1001/rss.xml",
    "cnbc": "https://www.cnbc.com/id/100003114/device/rss/rss.html",
    "cnn": "http://rss.cnn.com/rss/cnn_topstories.rss",
    "nytimes": "https://rss.nytimes.com/services/xml/rss/nyt/HomePage.xml",
    "new york times": "https://rss.nytimes.com/services/xml/rss/nyt/HomePage.xml",
    "wsj": "https://feeds.a.dj.com/rss/RSSWorldNews.xml",
    "wall street journal": "https://feeds.a.dj.com/rss/RSSWorldNews.xml",
    "techcrunch": "https://techcrunch.com/feed/",
    "the verge": "https://www.theverge.com/rss/index.xml",
}


def _clean_text(value: str | None) -> str:
    if not value:
        return ""
    value = unescape(value)
    value = re.sub(r"<[^>]+>", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def _format_dt(value: str | None) -> str:
    if not value:
        return ""
    try:
        return parsedate_to_datetime(value).isoformat()
    except Exception:
        return value.strip()


def _fetch_url(url: str, timeout: int = 20) -> bytes:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "HermesRSS/1.0 (+https://github.com/stpater77/hermes-railway)",
            "Accept": "application/rss+xml, application/atom+xml, application/xml, text/xml, */*",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _parse_rss(root: ET.Element, limit: int) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    channel = root.find("channel")
    rss_items = channel.findall("item") if channel is not None else root.findall(".//item")

    for item in rss_items[:limit]:
        title = _clean_text(item.findtext("title"))
        link = _clean_text(item.findtext("link"))
        published = _format_dt(item.findtext("pubDate") or item.findtext("published"))
        summary = _clean_text(item.findtext("description"))

        items.append({
            "title": title,
            "url": link,
            "published": published,
            "summary": summary,
        })

    return items


def _parse_atom(root: ET.Element, limit: int) -> list[dict[str, Any]]:
    ns = {"atom": "http://www.w3.org/2005/Atom"}
    entries = root.findall("atom:entry", ns) or root.findall(".//{http://www.w3.org/2005/Atom}entry")
    items: list[dict[str, Any]] = []

    for entry in entries[:limit]:
        title = _clean_text(entry.findtext("atom:title", default="", namespaces=ns))
        link = ""

        for link_el in entry.findall("atom:link", ns):
            href = link_el.attrib.get("href")
            rel = link_el.attrib.get("rel", "alternate")
            if href and rel == "alternate":
                link = href
                break
        if not link:
            first_link = entry.find("atom:link", ns)
            link = first_link.attrib.get("href", "") if first_link is not None else ""

        published = _clean_text(
            entry.findtext("atom:published", default="", namespaces=ns)
            or entry.findtext("atom:updated", default="", namespaces=ns)
        )
        summary = _clean_text(
            entry.findtext("atom:summary", default="", namespaces=ns)
            or entry.findtext("atom:content", default="", namespaces=ns)
        )

        items.append({
            "title": title,
            "url": link,
            "published": published,
            "summary": summary,
        })

    return items


def get_feed_headlines(feed_url: str, limit: int = 5) -> list[dict[str, Any]]:
    """Fetch headlines from an RSS or Atom feed URL."""
    if not feed_url:
        raise ValueError("feed_url is required")

    limit = max(1, min(int(limit or 5), 25))
    data = _fetch_url(feed_url)
    root = ET.fromstring(data)

    tag = root.tag.lower()
    if tag.endswith("rss") or root.find("channel") is not None:
        return _parse_rss(root, limit)
    if tag.endswith("feed"):
        return _parse_atom(root, limit)

    raise ValueError(f"Unsupported feed format: {root.tag}")


def get_headlines_for_source(source: str, limit: int = 5) -> list[dict[str, Any]]:
    """Fetch headlines from a known source alias such as fox, reuters, ap, bbc, cnbc."""
    key = (source or "").strip().lower()
    if not key:
        raise ValueError("source is required")

    feed_url = KNOWN_FEEDS.get(key)
    if not feed_url:
        available = ", ".join(sorted(KNOWN_FEEDS))
        raise ValueError(f"Unknown RSS source: {source}. Available sources: {available}")

    return get_feed_headlines(feed_url, limit=limit)


def list_known_feeds() -> dict[str, str]:
    """Return known RSS source aliases and feed URLs."""
    return dict(sorted(KNOWN_FEEDS.items()))


def headline_lines(source: str = "", feed_url: str = "", limit: int = 5) -> str:
    """Return RSS headlines as a clean numbered list with URLs."""
    if feed_url:
        items = get_feed_headlines(feed_url, limit=limit)
    else:
        items = get_headlines_for_source(source, limit=limit)

    lines: list[str] = []
    for i, item in enumerate(items, 1):
        title = item.get("title") or "(no title)"
        url = item.get("url") or ""
        published = item.get("published") or ""
        suffix = f" — {url}" if url else ""
        if published:
            suffix += f" ({published})"
        lines.append(f"{i}. {title}{suffix}")
    return "\n".join(lines)


RSS_NEWS_ACTIONS = {
    "list_known_feeds": list_known_feeds,
    "get_feed_headlines": get_feed_headlines,
    "get_headlines_for_source": get_headlines_for_source,
    "headline_lines": headline_lines,
}


def rss_news(action: str, **kwargs) -> str:
    """Dispatch RSS/news feed actions and return JSON-safe text."""
    if not action:
        raise ValueError("action is required")

    action = action.strip()
    if action not in RSS_NEWS_ACTIONS:
        available = ", ".join(sorted(RSS_NEWS_ACTIONS))
        raise ValueError(f"Unknown RSS news action: {action}. Available actions: {available}")

    result = RSS_NEWS_ACTIONS[action](**kwargs)

    return json.dumps({
        "action": action,
        "result": result,
    }, ensure_ascii=False, default=str)


RSS_NEWS_SCHEMA = {
    "name": "rss_news",
    "description": (
        "Fetch low-token RSS/Atom news headlines from known feeds or a provided feed URL. "
        "Use this when the user explicitly asks for RSS, feeds, strict news headlines, "
        "or low-token headline retrieval. For general web research, use web_search/web_extract."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "action": {
                "type": "string",
                "description": "RSS/news action to run.",
                "enum": sorted(RSS_NEWS_ACTIONS.keys()),
            },
            "source": {
                "type": "string",
                "description": "Known source alias, e.g. fox, reuters, ap, bbc, cnbc.",
            },
            "feed_url": {
                "type": "string",
                "description": "Direct RSS or Atom feed URL.",
            },
            "limit": {
                "type": "integer",
                "description": "Maximum number of feed items to return.",
            },
        },
        "required": ["action"],
    },
}


def check_rss_news_requirements() -> bool:
    return True


from tools.registry import registry

registry.register(
    name="rss_news",
    emoji="📰",
    toolset="rss_news",
    schema=RSS_NEWS_SCHEMA,
    handler=lambda args, **kw: rss_news(**args),
    check_fn=check_rss_news_requirements,
)
