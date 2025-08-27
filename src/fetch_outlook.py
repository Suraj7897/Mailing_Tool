#!/usr/bin/env python3
"""
Outlook last-N-days email export using Microsoft Graph + MSAL (device code).
- Filters by time window and optional single keyword (client-side filter)
- Traverses a folder path like: Inbox/Invoices/2025
- Extracts dates from email body (NLP + regex)
- Writes results to Excel (fresh file each run; uses atomic replace)
"""
import os
import sys
import time
import re
import argparse
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional

import requests
import msal
import pandas as pd
from bs4 import BeautifulSoup
from dateutil import tz

import dateparser  # for parse fallback
try:
    from dateparser.search import search_dates
except Exception:
    search_dates = None

IST = tz.gettz("Asia/Kolkata")
CACHE_FILE = "msal_cache.bin"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

DATE_PATTERNS = [
    r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b',
    r'\b\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4}\b',
    r'\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*,\s*\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b',
]

def read_env():
    env = {}
    if os.path.exists(".env"):
        with open(".env","r", encoding="utf-8") as f:
            for line in f:
                line=line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    k,v = line.split("=",1)
                    env[k.strip()] = v.strip()
    return env

def build_app(client_id: str, tenant: str) -> msal.PublicClientApplication:
    authority = f"https://login.microsoftonline.com/{tenant}"
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        try:
            cache.deserialize(open(CACHE_FILE,"r").read())
        except Exception:
            pass
    app = msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)
    return app

def save_cache(app: msal.PublicClientApplication):
    try:
        cache = app.token_cache
        if cache.has_state_changed:
            with open(CACHE_FILE, "w") as f:
                f.write(cache.serialize())
    except Exception:
        pass

def acquire_token(app: msal.PublicClientApplication, scopes: List[str]) -> str:
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise RuntimeError("Failed to create device flow. Check your CLIENT_ID and tenant.")
    print(f"\n>>> Open {flow['verification_uri']} and enter code: {flow['user_code']}")
    print(">>> Waiting for sign-in...")
    result = app.acquire_token_by_device_flow(flow)
    save_cache(app)
    if "access_token" not in result:
        raise RuntimeError(f"Token error: {result.get('error_description', result)}")
    return result["access_token"]

def strip_html_to_text(html: str) -> str:
    soup = BeautifulSoup(html or "", "html.parser")
    return soup.get_text(" ", strip=True)

def backoff_sleep(attempt: int):
    delay = min(60, (2 ** attempt))
    time.sleep(delay)

def graph_get(url: str, token: str, params: Optional[Dict]=None) -> Dict:
    headers = {"Authorization": f"Bearer {token}"}
    attempt = 0
    while True:
        resp = requests.get(url, headers=headers, params=params)
        if resp.status_code == 429 or resp.status_code >= 500:
            attempt += 1
            if attempt > 5:
                resp.raise_for_status()
            backoff_sleep(attempt)
            continue
        if resp.status_code != 200:
            raise RuntimeError(f"Graph error {resp.status_code}: {resp.text}")
        return resp.json()

def find_folder_id(token: str, folder_path: str) -> str:
    parts = [p for p in folder_path.split("/") if p]
    if not parts:
        raise ValueError("Folder path is empty.")
    url = f"{GRAPH_BASE}/me/mailFolders?$top=200"
    data = graph_get(url, token)
    current = None
    for f in data.get("value", []):
        if f.get("displayName") == parts[0]:
            current = f
            break
    if current is None:
        # try known folder id path
        well_known = parts[0]
        test = f"{GRAPH_BASE}/me/mailFolders('{well_known}')"
        try:
            current = graph_get(test, token)
        except Exception:
            raise RuntimeError(f"Top-level folder '{parts[0]}' not found. Check spelling.")
    for name in parts[1:]:
        url = f"{GRAPH_BASE}/me/mailFolders('{current['id']}')/childFolders?$top=200"
        data = graph_get(url, token)
        nxt = None
        for f in data.get("value", []):
            if f.get("displayName") == name:
                nxt = f
                break
        if nxt is None:
            raise RuntimeError(f"Subfolder '{name}' not found under '{current.get('displayName')}'.")
        current = nxt
    return current["id"]

def iso_utc(dt_obj: datetime) -> str:
    return dt_obj.replace(tzinfo=timezone.utc).isoformat().replace("+00:00","Z")

def extract_dates(text: str) -> List[str]:
    hits = set()
    if search_dates:
        try:
            res = search_dates(
                text or "",
                languages=['en'],
                settings={'RETURN_AS_TIMEZONE_AWARE': False, 'PREFER_DATES_FROM': 'past'}
            ) or []
            for _, d in res:
                try:
                    hits.add(d.date().isoformat())
                except Exception:
                    pass
        except Exception:
            # fall through to regex
            pass
    for pat in DATE_PATTERNS:
        for m in re.findall(pat, text or "", flags=re.IGNORECASE):
            d = dateparser.parse(m)
            if d:
                hits.add(d.date().isoformat())
    return sorted(hits)

def _matches_keyword(subject: str, content: str, sender_name: str, sender_email: str, keyword: str) -> bool:
    if not keyword:
        return True
    k = keyword.lower()
    return (k in (subject or "").lower()
            or k in (content or "").lower()
            or k in (sender_name or "").lower()
            or k in (sender_email or "").lower())

def collect_messages(token: str, folder_id: str, days: int, keyword: Optional[str]) -> List[Dict]:
    since_utc = datetime.utcnow() - timedelta(days=days)
    filter_since = f"receivedDateTime ge {iso_utc(since_utc)}"
    params = {
        "$top": "50",
        "$select": "id,subject,from,receivedDateTime,webLink,bodyPreview,body"
    }
    if filter_since:
        params["$filter"] = filter_since

    url = f"{GRAPH_BASE}/me/mailFolders('{folder_id}')/messages"
    rows = []
    total_fetched = 0
    matched = 0
    while True:
        data = graph_get(url, token, params=params)
        for m in data.get("value", []):
            total_fetched += 1
            subject = m.get("subject","") or ""
            received = m.get("receivedDateTime")
            if received:
                dt_obj = datetime.fromisoformat(received.replace("Z","+00:00"))
                dt_ist = dt_obj.astimezone(IST).strftime("%Y-%m-%d %H:%M")
            else:
                dt_ist = ""
            content = ""
            body = m.get("body")
            if isinstance(body, dict):
                content = body.get("content","") or ""
                if (body.get("contentType") or "").lower() == "html":
                    content = strip_html_to_text(content)
            else:
                content = m.get("bodyPreview","") or ""

            sender_name = ""
            sender_email = ""
            frm = m.get("from",{}).get("emailAddress",{})
            if isinstance(frm, dict):
                sender_name = frm.get("name","") or ""
                sender_email = frm.get("address","") or ""

            # Client-side keyword match (case-insensitive substring)
            if not _matches_keyword(subject, content, sender_name, sender_email, keyword):
                continue

            found_dates = ", ".join(extract_dates(content))
            rows.append({
                "Subject": subject,
                "Received (IST)": dt_ist,
                "Extracted Dates": found_dates,
                "Sender Name": sender_name,
                "Sender Email": sender_email,
                "Link": m.get("webLink","")
            })
            matched += 1

        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        url, params = next_link, None
    # print small summary to help user debug
    print(f"Fetched {total_fetched} messages (since last {days} days).")
    if keyword:
        print(f"Keyword filter: '{keyword}' -> {matched} matching messages saved.")
    else:
        print(f"No keyword provided -> {matched} messages saved.")
    return rows

def main():
    env = read_env()
    parser = argparse.ArgumentParser(description="Export Outlook emails for last N days with single-keyword search to Excel.")
    parser.add_argument("--client-id", default=env.get("CLIENT_ID"), help="Azure App (client) ID")
    parser.add_argument("--tenant", default=env.get("TENANT","common"), help="Tenant id or 'common'")
    parser.add_argument("--folder", default=env.get("FOLDER_PATH","Inbox"), help="Folder path e.g. Inbox/Invoices")
    parser.add_argument("--keywords", default=env.get("KEYWORDS",""), help="Single keyword (no commas). Optional.")
    parser.add_argument("--days", type=int, default=int(env.get("DAYS", "7")), help="Lookback days (default 7)")
    parser.add_argument("--out", default=env.get("OUTPUT_XLSX","outlook_last7days.xlsx"), help="Output Excel filename")
    args = parser.parse_args()

    if not args.client_id:
        print("ERROR: CLIENT_ID missing. Set it in .env or pass --client-id.")
        sys.exit(1)

    # enforce single keyword (no commas)
    keyword = args.keywords.strip() if args.keywords else ""
    if keyword and "," in keyword:
        print("ERROR: Only one keyword allowed; remove commas. Example: --keywords invoice")
        sys.exit(1)

    app = build_app(args.client_id, args.tenant)
    scopes = ["User.Read", "Mail.Read"]
    token = acquire_token(app, scopes)

    folder_id = find_folder_id(token, args.folder)

    rows = collect_messages(token, folder_id, args.days, keyword or None)

    df = pd.DataFrame(rows, columns=[
        "Subject", "Received (IST)", "Extracted Dates", "Sender Name", "Sender Email", "Link"
    ])

    # write to a temporary file and atomically replace the destination
    out_path = args.out
    tmp_path = out_path + ".tmp.xlsx"
    try:
        df.to_excel(tmp_path, index=False)
        # attempt atomic replace
        os.replace(tmp_path, out_path)
    except PermissionError:
        # most likely Excel has the file open
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        print(f"\nERROR: Cannot write to '{out_path}'. It may be open in Excel. Please close the file and re-run the script.")
        sys.exit(1)
    except Exception as e:
        # cleanup temp
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        print("ERROR writing output:", str(e))
        sys.exit(1)

    print(f"Wrote {len(df)} rows to {out_path}")

if __name__ == "__main__":
    main()