"""Microsoft Graph helper functions for Hermes."""

from __future__ import annotations

import json
import urllib.parse
import urllib.request
from typing import Any

from agent.microsoft_oauth import get_access_token


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def _request(method: str, path: str, payload: dict[str, Any] | None = None, params: dict[str, Any] | None = None) -> dict[str, Any]:
    token = get_access_token()

    url = GRAPH_BASE_URL + path
    if params:
        url += "?" + urllib.parse.urlencode(params)

    data = None
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"

    req = urllib.request.Request(url, data=data, headers=headers, method=method)

    with urllib.request.urlopen(req, timeout=30) as resp:
        raw = resp.read().decode("utf-8")
        return json.loads(raw) if raw else {}


def graph_get(path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
    return _request("GET", path, params=params)


def graph_post(path: str, payload: dict[str, Any]) -> dict[str, Any]:
    return _request("POST", path, payload=payload)


def list_recent_messages(limit: int = 10) -> list[dict[str, Any]]:
    data = graph_get(
        "/me/messages",
        {
            "$top": max(1, min(int(limit), 50)),
            "$select": "id,subject,from,receivedDateTime,webLink",
            "$orderby": "receivedDateTime DESC",
        },
    )
    return data.get("value", [])


def get_message(message_id: str) -> dict[str, Any]:
    if not message_id:
        raise ValueError("message_id is required")

    return graph_get(
        f"/me/messages/{urllib.parse.quote(message_id)}",
        {
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,webLink",
        },
    )


def search_messages(query: str, limit: int = 10) -> list[dict[str, Any]]:
    if not query or not query.strip():
        raise ValueError("query is required")

    data = graph_get(
        "/me/messages",
        {
            "$top": max(1, min(int(limit), 50)),
            "$search": f'"{query.strip()}"',
            "$select": "id,subject,from,receivedDateTime,bodyPreview,webLink",
        },
    )
    return data.get("value", [])


def _recipient_list(to: str | list[str]) -> list[dict[str, dict[str, str]]]:
    if isinstance(to, str):
        recipients = [addr.strip() for addr in to.split(",") if addr.strip()]
    else:
        recipients = [str(addr).strip() for addr in to if str(addr).strip()]

    if not recipients:
        raise ValueError("At least one recipient is required")

    return [{"emailAddress": {"address": address}} for address in recipients]


def send_message(to: str | list[str], subject: str, body: str, save_to_sent_items: bool = True) -> dict[str, Any]:
    if not subject:
        raise ValueError("subject is required")
    if not body:
        raise ValueError("body is required")

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": _recipient_list(to),
        },
        "saveToSentItems": save_to_sent_items,
    }

    return graph_post("/me/sendMail", payload)


def create_draft(to: str | list[str], subject: str, body: str) -> dict[str, Any]:
    if not subject:
        raise ValueError("subject is required")
    if not body:
        raise ValueError("body is required")

    payload = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": _recipient_list(to),
    }

    return graph_post("/me/messages", payload)


def summarizable_message_text(message_id: str) -> str:
    msg = get_message(message_id)

    sender = (msg.get("from") or {}).get("emailAddress", {})
    body = msg.get("body") or {}

    parts = [
        f"Subject: {msg.get('subject') or ''}",
        f"From: {sender.get('name') or ''} <{sender.get('address') or ''}>",
        f"Received: {msg.get('receivedDateTime') or ''}",
        "",
        body.get("content") or msg.get("bodyPreview") or "",
    ]

    return "\n".join(parts).strip()


def _html_to_text(html: str) -> str:
    import re
    from html import unescape

    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", html or "")
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p\s*>", "\n", text)
    text = re.sub(r"(?i)</div\s*>", "\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def summarizable_message_text(message_id: str) -> str:
    msg = get_message(message_id)

    sender = (msg.get("from") or {}).get("emailAddress", {})
    body = msg.get("body") or {}
    raw_body = body.get("content") or msg.get("bodyPreview") or ""

    content_type = (body.get("contentType") or "").lower()
    clean_body = _html_to_text(raw_body) if content_type == "html" else raw_body

    parts = [
        f"Subject: {msg.get('subject') or ''}",
        f"From: {sender.get('name') or ''} <{sender.get('address') or ''}>",
        f"Received: {msg.get('receivedDateTime') or ''}",
        "",
        clean_body,
    ]

    return "\n".join(parts).strip()


def list_calendar_events(limit: int = 10) -> list[dict[str, Any]]:
    data = graph_get(
        "/me/events",
        {
            "$top": max(1, min(int(limit), 50)),
            "$select": "id,subject,start,end,location,organizer,attendees,webLink",
            "$orderby": "start/dateTime",
        },
    )
    return data.get("value", [])


def list_upcoming_calendar_events(limit: int = 10, days_ahead: int = 30) -> list[dict[str, Any]]:
    from datetime import datetime, timedelta, timezone

    now = datetime.now(timezone.utc)
    end = now + timedelta(days=max(1, int(days_ahead)))

    data = graph_get(
        "/me/calendarView",
        {
            "startDateTime": now.isoformat(),
            "endDateTime": end.isoformat(),
            "$top": max(1, min(int(limit), 50)),
            "$select": "id,subject,start,end,location,organizer,attendees,webLink",
            "$orderby": "start/dateTime",
        },
    )
    return data.get("value", [])


def create_calendar_event(
    subject: str,
    start_datetime: str,
    end_datetime: str,
    timezone: str = "America/New_York",
    location: str = "",
    body: str = "",
    attendees: list[str] | None = None,
) -> dict[str, Any]:
    if not subject:
        raise ValueError("subject is required")
    if not start_datetime:
        raise ValueError("start_datetime is required")
    if not end_datetime:
        raise ValueError("end_datetime is required")

    payload = {
        "subject": subject,
        "body": {
            "contentType": "Text",
            "content": body or "",
        },
        "start": {
            "dateTime": start_datetime,
            "timeZone": timezone,
        },
        "end": {
            "dateTime": end_datetime,
            "timeZone": timezone,
        },
        "location": {
            "displayName": location or "",
        },
        "attendees": [
            {
                "emailAddress": {
                    "address": email,
                },
                "type": "required",
            }
            for email in (attendees or [])
        ],
    }

    return graph_post("/me/events", payload)


def list_onedrive_root(limit: int = 25) -> list[dict[str, Any]]:
    data = graph_get(
        "/me/drive/root/children",
        {
            "$top": max(1, min(int(limit), 100)),
            "$select": "id,name,size,webUrl,file,folder,lastModifiedDateTime",
        },
    )
    return data.get("value", [])


def graph_put_bytes(path: str, content: bytes, content_type: str = "application/octet-stream") -> dict[str, Any]:
    token = get_access_token()

    url = GRAPH_BASE_URL + path

    req = urllib.request.Request(
        url,
        data=content,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "Content-Type": content_type,
        },
        method="PUT",
    )

    with urllib.request.urlopen(req, timeout=30) as resp:
        raw = resp.read().decode("utf-8")
        return json.loads(raw) if raw else {}


def upload_text_file_to_onedrive(filename: str, content: str) -> dict[str, Any]:
    if not filename:
        raise ValueError("filename is required")

    safe_name = filename.strip().replace("/", "-")
    upload_path = f"/me/drive/root:/{urllib.parse.quote(safe_name)}:/content"

    return graph_put_bytes(
        upload_path,
        content.encode("utf-8"),
        "text/plain; charset=utf-8",
    )


def list_todo_lists() -> list[dict[str, Any]]:
    data = graph_get(
        "/me/todo/lists",
        {
            "$select": "id,displayName,isOwner,isShared,isDefaultList",
        },
    )
    return data.get("value", [])


def create_word_docx_and_upload(filename: str, title: str, body: str) -> dict[str, Any]:
    from io import BytesIO
    from docx import Document

    if not filename:
        raise ValueError("filename is required")
    if not filename.lower().endswith(".docx"):
        filename = filename + ".docx"

    doc = Document()

    if title:
        doc.add_heading(title, level=1)

    for paragraph in (body or "").split("\n\n"):
        clean = paragraph.strip()
        if clean:
            doc.add_paragraph(clean)

    buffer = BytesIO()
    doc.save(buffer)

    safe_name = filename.strip().replace("/", "-")
    upload_path = f"/me/drive/root:/{urllib.parse.quote(safe_name)}:/content"

    return graph_put_bytes(
        upload_path,
        buffer.getvalue(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
