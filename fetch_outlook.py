"""
Fetch mail from Outlook.com / Microsoft 365 with the Microsoft Graph API
and cache it in the same format as fetch_threads.py.

    pip install msal requests
    python fetch_outlook.py

One-time setup (free, about 5 minutes):
  1. Go to entra.microsoft.com -> App registrations -> New registration
  2. Supported account types: "Accounts in any organizational directory and
     personal Microsoft accounts"
  3. Leave Redirect URI blank (device code flow doesn't need one)
  4. Authentication -> Advanced settings -> Allow public client flows = Yes -> Save
  5. Copy the Application (client) ID and set it:
         Windows cmd:   set OUTLOOK_CLIENT_ID=your-id-here
         PowerShell:    $env:OUTLOOK_CLIENT_ID="your-id-here"
  (You don't need to add API permissions by hand: the first sign-in asks you to
   consent to read-only mail access.)

Needs email_analyzer.py in the same folder (reuses its text-cleanup helpers).
"""

import json
import os
import time
from datetime import datetime, timedelta, timezone

import msal
import requests

from email_analyzer import company_from_address, strip_quoted_reply

CLIENT_ID = "INSERT_ID_HERE"
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["User.Read", "Mail.Read"]                                # read-only
GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN_FILE = "outlook_token.json"                         # keep private, like token.json


def get_token():
    if not CLIENT_ID:
        raise SystemExit("Set OUTLOOK_CLIENT_ID first (see the setup steps at the top of this file).")
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_FILE):
        cache.deserialize(open(TOKEN_FILE, encoding="utf-8").read())
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result:                                  # first run, or the saved token expired
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Could not start sign-in: {flow}")
        print(flow["message"])                      # "go to microsoft.com/devicelogin, enter code..."
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", result))

    if cache.has_state_changed:
        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(cache.serialize())
    return result["access_token"]


def graph_get(url, token, params=None):
    """GET with retry when Graph says slow down (429) or is briefly unavailable."""
    headers = {"Authorization": f"Bearer {token}",
               "Prefer": 'outlook.body-content-type="text"'}      # plain-text bodies
    for attempt in range(6):
        r = requests.get(url, headers=headers, params=params, timeout=30)
        if r.status_code in (429, 503, 504):
            time.sleep(int(r.headers.get("Retry-After", 2 ** attempt)))
            continue
        r.raise_for_status()
        return r.json()
    r.raise_for_status()


def _iso(s):
    return s.replace("Z", "+00:00")


def fetch_outlook_mail(max_emails=150, days=180):
    token = get_token()
    me = graph_get(f"{GRAPH}/me", token, {"$select": "mail,userPrincipalName"})
    my_addr = (me.get("mail") or me.get("userPrincipalName") or "").lower()

    since = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url = f"{GRAPH}/me/messages"            # includes Sent Items
    params = {
        "$select": "id,conversationId,subject,from,toRecipients,ccRecipients,"
                   "receivedDateTime,sentDateTime,body,isDraft",
        "$filter": f"receivedDateTime ge {since}",
        "$orderby": "receivedDateTime desc",
        "$top": 50,
    }

    emails = []
    while url and len(emails) < max_emails:
        page = graph_get(url, token, params)
        params = None                       # the nextLink already carries the query
        for m in page.get("value", []):
            if m.get("isDraft"):
                continue
            frm = (m.get("from") or {}).get("emailAddress", {})
            addr, name = (frm.get("address") or "").lower(), frm.get("name") or ""
            mine = addr == my_addr
            recips = (m.get("toRecipients") or []) + (m.get("ccRecipients") or [])
            emails.append({
                "id": m["id"],
                "thread_id": m.get("conversationId"),       # Outlook's version of a thread
                "subject": m.get("subject") or "",
                "date": _iso(m.get("sentDateTime") or m["receivedDateTime"]),
                "labels": ["SENT"] if mine else [],
                **company_from_address(f"{name} <{addr}>"),
                "to": ", ".join(r["emailAddress"]["address"] for r in recips),
                "body": strip_quoted_reply((m.get("body") or {}).get("content", "")),
            })
        url = page.get("@odata.nextLink")
    return emails[:max_emails]


if __name__ == "__main__":
    emails = fetch_outlook_mail(max_emails=150, days=180)   # try max_emails=20 first

    with open("emails_cache.json", "w", encoding="utf-8") as f:
        json.dump(emails, f, indent=2)
    n_threads = len({e["thread_id"] for e in emails})
    print(f"Saved {len(emails)} messages from {n_threads} conversations to emails_cache.json")
    print("Set MY_EMAIL in rank_companies.py to this Outlook address before ranking.")