"""
Fetch Gmail threads (your outreach plus every reply) and cache them.

    python fetch_threads.py

Needs email_analyzer.py (your existing Gmail fetch file) in the same folder,
because it reuses its login and cleanup helpers.

Output: emails_cache.json, in the same format rank_companies.py already reads.
"""

import json
from datetime import datetime, timezone

from email_analyzer import (
    _extract_body,
    company_from_address,
    gmail_service,
    strip_quoted_reply,
)

from rank_companies import PROFILE   # search wording depends on MODE (sponsor or job)

# Gmail search used to find threads. Loosen it (e.g. "in:sent newer_than:180d") to pull
# more mail; rank_companies.py still filters out anything that isn't relevant.
QUERY = PROFILE["query"]


def fetch_sent_threads(query=QUERY, max_threads=150):
    svc = gmail_service()

    thread_ids, page_token = [], None
    while len(thread_ids) < max_threads:
        resp = svc.users().threads().list(
            userId="me", q=query,
            maxResults=min(100, max_threads - len(thread_ids)),
            pageToken=page_token,
        ).execute(num_retries=5)
        thread_ids += [t["id"] for t in resp.get("threads", [])]
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    emails = []
    for tid in thread_ids:
        thread = svc.users().threads().get(
            userId="me", id=tid, format="full"
        ).execute(num_retries=5)

        for msg in thread.get("messages", []):
            headers = {h["name"].lower(): h["value"] for h in msg["payload"]["headers"]}
            sent_at = datetime.fromtimestamp(int(msg["internalDate"]) / 1000, timezone.utc)
            emails.append({
                "id": msg["id"],
                "thread_id": msg["threadId"],
                "subject": headers.get("subject", ""),
                "date": sent_at.isoformat(),
                "labels": msg.get("labelIds", []),       # "SENT" marks your own messages
                **company_from_address(headers.get("from", "")),
                "to": f"{headers.get('to', '')}, {headers.get('cc', '')}",
                "body": strip_quoted_reply(_extract_body(msg["payload"])),
            })
    return emails


if __name__ == "__main__":
    emails = fetch_sent_threads(max_threads=150)   # try 5 first for a quick test

    with open("emails_cache.json", "w", encoding="utf-8") as f:
        json.dump(emails, f, indent=2)
    n_threads = len({e["thread_id"] for e in emails})
    print(f"Saved {len(emails)} messages from {n_threads} threads to emails_cache.json")