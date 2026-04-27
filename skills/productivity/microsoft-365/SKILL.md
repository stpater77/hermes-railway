---
name: microsoft-365
description: Microsoft 365 integration for Hermes covering Outlook email, Outlook Calendar, OneDrive, and Word document creation through Microsoft Graph OAuth. Use when the user asks Hermes to read, search, summarize, draft, or send Outlook email; create or read calendar events; read or save OneDrive files; or create and save Word documents.
---

# Microsoft 365

Microsoft 365 integration for Hermes through Microsoft Graph and Hermes-managed OAuth.

## Scripts

- scripts/microsoft_graph.py

## Current Capabilities

- Outlook Mail: list, search, read, summarize-ready text extraction, draft, send
- Outlook Calendar: list events, list upcoming events, read event details, create events, update events, and delete events
- OneDrive: list root files/folders, upload text files
- Word: create and upload .docx documents
- Microsoft To Do: list task lists, create tasks, list tasks, update tasks, and delete tasks

## Key Functions

Use these functions from scripts/microsoft_graph.py:

- list_recent_messages(limit)
- search_messages(query, limit)
- get_message(message_id)
- summarizable_message_text(message_id)
- create_draft(to, subject, body)
- send_message(to, subject, body)
- list_upcoming_calendar_events(limit, days_ahead)
- get_calendar_event(event_id)
- create_calendar_event(subject, start_datetime, end_datetime, timezone, location, body, attendees)
- update_calendar_event(event_id, subject, start_datetime, end_datetime, timezone, location, body, attendees)
- delete_calendar_event(event_id)
- list_onedrive_root(limit)
- upload_text_file_to_onedrive(filename, content)
- create_word_docx_and_upload(filename, title, body)

## Safety Rules

1. Never send an email without explicit user confirmation.
2. Never create, update, or delete a calendar event without explicit user confirmation.
3. When composing email, create a draft first unless the user explicitly says to send now.
4. Use summarizable_message_text(message_id) before summarizing email content.
5. Confirm target filenames before saving files to OneDrive unless the user gave an exact filename.
6. Do not expose tokens, client IDs, tenant IDs, or raw OAuth payloads.
7. Do not assign or delegate Microsoft To Do tasks unless assignment support is explicitly added and tested.

## Auth State

Required local files/config:

- ~/.hermes/hermes-agent/agent/microsoft_oauth.py
- ~/.hermes/auth/microsoft_oauth.json
- ~/.hermes/.env

Required env values in ~/.hermes/.env:

- HERMES_MICROSOFT_CLIENT_ID
- HERMES_MICROSOFT_TENANT_ID

## Deferred

- Microsoft To Do assignment/delegation workflows
- Microsoft Teams
- SharePoint / Sites
- Advanced attachment workflows
- Rich HTML email composition
- Threaded replies
