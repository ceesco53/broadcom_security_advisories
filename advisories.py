import csv
import sys
from datetime import datetime, timedelta
from typing import Dict, List, Any
import re
import requests

# -------------------- Config --------------------
API_URL = "https://support.broadcom.com/web/ecx/security-advisory/-/securityadvisory/getSecurityAdvisoryList"
CSV_OUT = "broadcom_security_advisories.csv"
PAGE_SIZE = 50

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json;charset=UTF-8",
}

# -------------------- Helpers --------------------
def parse_updated(ts: str) -> datetime:
    ts = ts.strip()
    fmts = [
        "%Y-%m-%d %H:%M:%S.%f",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ]
    for fmt in fmts:
        try:
            return datetime.strptime(ts, fmt)
        except ValueError:
            pass
    raise ValueError(f"Unrecognized updated timestamp: {ts}")

def within_last_14_days(updated_dt: datetime, now_local: datetime) -> bool:
    return (now_local - updated_dt) <= timedelta(days=14)

def fmt_pub_date(dt: datetime) -> str:
    try:
        return dt.strftime("%-d %B %Y")
    except ValueError:
        return dt.strftime("%#d %B %Y")

# NEW: keep full prefixed notification code (e.g., VTDSA-2025-12345)
def get_full_notification_id(item: Dict[str, Any]) -> str:
    """
    Prefer any API field that already contains the prefixed code (VTDSA-YYYY-NNNN).
    If not present, attempt to extract from notificationUrl.
    Fallback to raw notificationId as string.
    """
    candidates = [
        "notificationCode",
        "notificationNo",
        "notificationNumber",
        "notificationIdentifier",
        "notificationIdStr",
        "notification_id",  # just in case
    ]
    for key in candidates:
        val = item.get(key)
        if val and isinstance(val, str) and re.search(r"[A-Z]{3,}-\d{4}-\d+", val):
            return val.strip()

    url = (item.get("notificationUrl") or "").strip()
    if url:
        m = re.search(r"/([A-Z]{3,}-\d{4}-\d+)(?:[/?#]|$)", url)
        if m:
            return m.group(1)

    # Last resort: whatever the API exposes as notificationId
    return str(item.get("notificationId", "")).strip()

# -------------------- Fetch all pages --------------------
def fetch_page(page_number: int, page_size: int = PAGE_SIZE) -> Dict[str, Any]:
    payload = {
        "pageNumber": page_number,
        "pageSize": page_size,
        "searchVal": "",
        "segment": "VT",  # from your HAR; adjust if you want others
        "sortInfo": {"column": "", "order": ""},
    }
    r = requests.post(API_URL, json=payload, headers=HEADERS, timeout=30)
    r.raise_for_status()
    data = r.json()
    if not data.get("success", False):
        raise RuntimeError(f"API returned success=false: {data}")
    return data["data"]

def fetch_all_rows() -> List[Dict[str, Any]]:
    page_num = 0
    rows: List[Dict[str, Any]] = []
    while True:
        data = fetch_page(page_num)
        batch = data.get("list", []) or []
        rows.extend(batch)
        page_info = data.get("pageInfo") or {}
        total = page_info.get("totalCount") or len(rows)
        if len(rows) >= total or not batch:
            break
        page_num += 1
    return rows

# -------------------- Transform & Write --------------------
def main():
    now_local = datetime.now()
    all_rows = fetch_all_rows()

    out_records: List[Dict[str, str]] = []
    for item in all_rows:
        link = (item.get("notificationUrl") or "").strip()
        updated_raw = (item.get("updated") or "").strip()
        if not updated_raw:
            continue
        try:
            updated_dt = parse_updated(updated_raw)
        except ValueError:
            continue
        if not within_last_14_days(updated_dt, now_local):
            continue

        # Use the full VTDSA-style identifier where available
        notification_id = get_full_notification_id(item)

        severity = (item.get("severity") or "").strip()
        title = (item.get("title") or "").strip()
        pub_date = fmt_pub_date(updated_dt)

        out_records.append({
            "CVE ID": notification_id,   # ("Notification Id") with VTDSA-… preserved
            "RATING": severity,          # ("Severity")
            "COMMENTS": title,           # ("Title")
            "Link": link,                # href for the Notification Id
            "Pub date": pub_date,        # (today – Updated) as date string; API gives absolute Updated
        })

    fieldnames = ["CVE ID", "RATING", "COMMENTS", "Link", "Pub date"]
    with open(CSV_OUT, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for rec in out_records:
            w.writerow(rec)

    print(f"Wrote {len(out_records)} rows to {CSV_OUT}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        sys.stderr.write(f"Error: {e}\n")
        sys.exit(1)