#!/usr/bin/env python3
"""
OJCommerce Email Monitor — Hostinger Sent Folder Watcher

Connects to your Hostinger email via IMAP, watches the Sent folder,
and automatically updates the dashboard whenever you send an email
to a known prospect — no manual logging needed.

Usage:
  # Check Sent folder once and exit
  python3 email_monitor.py --check

  # Watch continuously (checks every 5 minutes)
  python3 email_monitor.py --watch

  # Watch with custom interval (e.g. every 2 minutes)
  python3 email_monitor.py --watch --interval 120
"""

import argparse
import csv
import imaplib
import email
import json
import sys
import time
from datetime import datetime
from email.header import decode_header
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import config
from dashboard import generate_dashboard


# ─── Helpers ─────────────────────────────────────────────────

def load_prospect_emails():
    """Return dict of {email_address: site_name} from prospects.csv."""
    mapping = {}
    with open(config.PROSPECTS_CSV, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            email_addr = row.get("contact_email", "").strip().lower()
            if email_addr:
                mapping[email_addr] = row["site_name"]
    return mapping


def load_seen_ids():
    """Load set of already-processed IMAP message IDs."""
    if config.MONITOR_STATE_FILE.exists():
        try:
            return set(json.loads(config.MONITOR_STATE_FILE.read_text()))
        except Exception:
            return set()
    return set()


def save_seen_ids(seen_ids):
    """Persist seen message IDs to disk."""
    config.MONITOR_STATE_FILE.write_text(json.dumps(list(seen_ids)))


def log_sent_auto(site_name, to_email, subject):
    """Append auto-detected send to sent_log.csv."""
    log_path = config.SENT_LOG_CSV
    write_header = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if write_header:
            writer.writerow(["timestamp", "site_name", "email", "template", "status", "error"])
        writer.writerow([
            datetime.now().isoformat(timespec="seconds"),
            site_name,
            to_email,
            "hostinger-sent",
            "sent",
            "",
        ])


def update_prospect_status(site_name, new_status):
    """Update status column for a prospect in prospects.csv."""
    rows = []
    fieldnames = None
    with open(config.PROSPECTS_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        for row in reader:
            if row["site_name"] == site_name:
                row["status"] = new_status
            rows.append(row)
    with open(config.PROSPECTS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def decode_str(value):
    """Decode an email header value to a plain string."""
    if value is None:
        return ""
    parts = decode_header(value)
    result = ""
    for part, charset in parts:
        if isinstance(part, bytes):
            result += part.decode(charset or "utf-8", errors="replace")
        else:
            result += part
    return result


def extract_to_addresses(msg):
    """Extract all To: email addresses from an email message."""
    to_header = msg.get("To", "")
    addresses = []
    for part in to_header.split(","):
        part = part.strip()
        # Extract email from "Name <email>" format
        if "<" in part and ">" in part:
            addr = part[part.index("<") + 1 : part.index(">")].strip().lower()
        else:
            addr = part.lower().strip()
        if addr:
            addresses.append(addr)
    return addresses


def extract_from_address(msg):
    """Extract the From: email address from an email message."""
    from_header = msg.get("From", "")
    if "<" in from_header and ">" in from_header:
        return from_header[from_header.index("<") + 1 : from_header.index(">")].strip().lower()
    return from_header.strip().lower()


def extract_email_body(msg):
    """Extract plain text body from a MIME email message."""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", ""))
            if content_type == "text/plain" and "attachment" not in disposition:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    return payload.decode(charset, errors="replace").strip()
        # Fallback: try text/html if no plain text found
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", ""))
            if content_type == "text/html" and "attachment" not in disposition:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    html_text = payload.decode(charset, errors="replace")
                    # Basic HTML stripping
                    import re
                    text = re.sub(r'<br\s*/?>', '\n', html_text)
                    text = re.sub(r'<[^>]+>', '', text)
                    text = re.sub(r'&nbsp;', ' ', text)
                    text = re.sub(r'&amp;', '&', text)
                    text = re.sub(r'&lt;', '<', text)
                    text = re.sub(r'&gt;', '>', text)
                    return text.strip()
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="replace").strip()
            # Strip HTML if content type is text/html
            if msg.get_content_type() == "text/html":
                import re
                text = re.sub(r'<br\s*/?>', '\n', text)
                text = re.sub(r'<[^>]+>', '', text)
                text = re.sub(r'&nbsp;', ' ', text)
                text = re.sub(r'&amp;', '&', text)
                text = re.sub(r'&lt;', '<', text)
                text = re.sub(r'&gt;', '>', text)
                text = re.sub(r'&#39;', "'", text)
                text = text.strip()
            return text
    return ""


def log_reply_auto(site_name, from_email, subject, body=""):
    """Append auto-detected reply to sent_log.csv and replies_log.csv."""
    # Log to sent_log.csv (existing behavior)
    log_path = config.SENT_LOG_CSV
    write_header = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if write_header:
            writer.writerow(["timestamp", "site_name", "email", "template", "status", "error"])
        writer.writerow([
            datetime.now().isoformat(timespec="seconds"),
            site_name,
            from_email,
            "reply-detected",
            "replied",
            "",
        ])

    # Log to replies_log.csv with full body
    replies_path = config.REPLIES_LOG_CSV
    write_header = not replies_path.exists()
    with open(replies_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if write_header:
            writer.writerow(["timestamp", "site_name", "from_email", "subject", "body"])
        writer.writerow([
            datetime.now().isoformat(timespec="seconds"),
            site_name,
            from_email,
            subject,
            body,
        ])


# ─── IMAP Connection ─────────────────────────────────────────

def connect_imap():
    """Connect and authenticate to Hostinger IMAP. Returns imap object."""
    if not config.IMAP_USERNAME or not config.IMAP_PASSWORD:
        print("ERROR: IMAP credentials not set. Add IMAP_USERNAME and IMAP_PASSWORD to .env")
        sys.exit(1)

    print(f"  Connecting to {config.IMAP_HOST}:{config.IMAP_PORT}...")
    imap = imaplib.IMAP4_SSL(config.IMAP_HOST, config.IMAP_PORT)
    imap.login(config.IMAP_USERNAME, config.IMAP_PASSWORD)
    print(f"  Logged in as {config.IMAP_USERNAME}")
    return imap


def find_sent_folder(imap):
    """Try to find the correct Sent folder name."""
    candidates = [
        config.IMAP_SENT_FOLDER,
        "Sent",
        "Sent Items",
        "Sent Messages",
        "INBOX.Sent",
        "[Gmail]/Sent Mail",
    ]
    _, folders = imap.list()
    available = []
    for f in folders:
        decoded = f.decode() if isinstance(f, bytes) else f
        available.append(decoded)

    for candidate in candidates:
        status, _ = imap.select(f'"{candidate}"', readonly=True)
        if status == "OK":
            print(f"  Found Sent folder: {candidate}")
            return candidate

    # Print available folders to help debug
    print("  Could not find Sent folder. Available folders:")
    for f in available:
        print(f"    {f}")
    print(f"  Set IMAP_SENT_FOLDER in .env to one of the above.")
    sys.exit(1)


# ─── Core Check Logic ────────────────────────────────────────

def check_sent_folder(imap, sent_folder, prospect_emails, seen_ids):
    """
    Scan the Sent folder for emails to known prospects.
    Returns list of newly detected sends: [(site_name, to_email, subject), ...]
    """
    imap.select(f'"{sent_folder}"', readonly=True)

    # Search all emails in Sent folder
    _, message_ids = imap.search(None, "ALL")
    all_ids = message_ids[0].split()

    if not all_ids:
        return []

    # Only look at the last 100 sent emails for performance
    recent_ids = all_ids[-100:]
    new_detections = []

    for msg_id in recent_ids:
        msg_id_str = msg_id.decode()

        if msg_id_str in seen_ids:
            continue

        # Fetch headers only (faster than full message)
        _, msg_data = imap.fetch(msg_id, "(BODY[HEADER.FIELDS (TO SUBJECT DATE)])")
        if not msg_data or not msg_data[0]:
            continue

        raw_header = msg_data[0][1]
        msg = email.message_from_bytes(raw_header)

        to_addresses = extract_to_addresses(msg)
        subject = decode_str(msg.get("Subject", ""))

        # Check if any To: address matches a known prospect
        for addr in to_addresses:
            if addr in prospect_emails:
                site_name = prospect_emails[addr]
                new_detections.append((site_name, addr, subject, msg_id_str))
                break

        seen_ids.add(msg_id_str)

    return new_detections


def check_inbox_replies(imap, prospect_emails, seen_ids):
    """
    Scan the Inbox for replies from known prospect email addresses.
    Also checks if the From domain matches a prospect domain (for auto-replies
    from different addresses like noreply@ or submissions@).
    Returns list of newly detected replies: [(site_name, from_email, subject, msg_id_str), ...]
    """
    imap.select("INBOX", readonly=True)

    _, message_ids = imap.search(None, "ALL")
    all_ids = message_ids[0].split()

    if not all_ids:
        return []

    # Only look at the last 200 inbox emails for performance
    recent_ids = all_ids[-200:]
    new_replies = []

    # Build domain-to-site mapping for matching auto-replies from different addresses
    domain_to_site = {}
    # Also build site_url-based keyword matching for Zendesk/third-party replies
    site_keywords = {}
    for addr, site_name in prospect_emails.items():
        domain = addr.split("@")[-1] if "@" in addr else ""
        if domain:
            domain_to_site[domain] = site_name
            # Extract base domain keyword (e.g. "thisoldhouse" from "thisoldhouse.com")
            base = domain.split(".")[0]
            if len(base) > 3:  # skip short generic domains
                site_keywords[base] = site_name

    for msg_id in recent_ids:
        msg_id_str = "inbox-" + msg_id.decode()

        if msg_id_str in seen_ids:
            continue

        # Fetch headers first for matching
        _, msg_data = imap.fetch(msg_id, "(BODY[HEADER.FIELDS (FROM SUBJECT DATE)])")
        if not msg_data or not msg_data[0]:
            seen_ids.add(msg_id_str)
            continue

        raw_header = msg_data[0][1]
        msg = email.message_from_bytes(raw_header)

        from_addr = extract_from_address(msg)
        subject = decode_str(msg.get("Subject", ""))
        from_domain = from_addr.split("@")[-1] if "@" in from_addr else ""

        site_name = None
        # Check 1: exact email match
        if from_addr in prospect_emails:
            site_name = prospect_emails[from_addr]
        # Check 2: domain match (catches auto-replies from noreply@, submissions@, etc.)
        elif from_domain in domain_to_site:
            site_name = domain_to_site[from_domain]
        # Check 3: keyword match in from_domain (catches Zendesk/third-party auto-replies
        # e.g. support@thisoldhouse.zendesk.com matches "thisoldhouse")
        else:
            for keyword, sname in site_keywords.items():
                if keyword in from_domain:
                    site_name = sname
                    break

        if site_name:
            # Fetch full message to get body content
            _, full_data = imap.fetch(msg_id, "(RFC822)")
            body = ""
            if full_data and full_data[0]:
                full_msg = email.message_from_bytes(full_data[0][1])
                body = extract_email_body(full_msg)
            new_replies.append((site_name, from_addr, subject, msg_id_str, body))

        seen_ids.add(msg_id_str)

    return new_replies


def process_detections(detections):
    """Log and update status for each newly detected send."""
    if not detections:
        return 0

    for site_name, to_email, subject, _ in detections:
        log_sent_auto(site_name, to_email, subject)
        update_prospect_status(site_name, "Pitched")
        print(f"  AUTO-DETECTED SEND  {site_name} ({to_email})")
        print(f"    Subject: {subject}")
        print()

    # Regenerate dashboard
    dash_path = generate_dashboard()
    print(f"  Dashboard updated: {dash_path}")
    return len(detections)


def process_replies(replies):
    """Log and update status for each newly detected reply."""
    if not replies:
        return 0

    for site_name, from_email, subject, _, body in replies:
        log_reply_auto(site_name, from_email, subject, body)
        update_prospect_status(site_name, "Replied")
        print(f"  REPLY DETECTED  {site_name} ({from_email})")
        print(f"    Subject: {subject}")
        if body:
            preview = body[:100].replace("\n", " ")
            print(f"    Body preview: {preview}...")
        print()

    # Regenerate dashboard
    dash_path = generate_dashboard()
    print(f"  Dashboard updated: {dash_path}")
    return len(replies)


# ─── Main ─────────────────────────────────────────────────────

def run_check(verbose=True):
    """Run a single check of Sent folder and Inbox."""
    prospect_emails = load_prospect_emails()
    seen_ids = load_seen_ids()

    if verbose:
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Checking Sent folder + Inbox...")
        print(f"  Watching {len(prospect_emails)} prospect emails")

    imap = connect_imap()

    # Check Sent folder for outgoing emails
    sent_folder = find_sent_folder(imap)
    detections = check_sent_folder(imap, sent_folder, prospect_emails, seen_ids)

    # Check Inbox for replies from prospects
    if verbose:
        print(f"  Scanning Inbox for replies...")
    replies = check_inbox_replies(imap, prospect_emails, seen_ids)

    imap.logout()

    save_seen_ids(seen_ids)
    sent_count = process_detections(detections)
    reply_count = process_replies(replies)
    total = sent_count + reply_count

    if verbose:
        if total == 0:
            print("  No new activity detected.")
        else:
            if sent_count:
                print(f"  {sent_count} new email(s) sent.")
            if reply_count:
                print(f"  {reply_count} new reply(ies) detected!")

    return total


def main():
    parser = argparse.ArgumentParser(
        description="OJCommerce Email Monitor — Hostinger Sent Folder Watcher",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true",
                      help="Check Sent folder once and exit")
    mode.add_argument("--watch", action="store_true",
                      help="Watch continuously (polls every N seconds)")

    parser.add_argument("--interval", type=int, default=config.IMAP_POLL_INTERVAL,
                        help=f"Polling interval in seconds (default: {config.IMAP_POLL_INTERVAL})")
    args = parser.parse_args()

    if args.check:
        run_check()
        return

    if args.watch:
        interval = args.interval

        # Save PID for dashboard detection
        import os
        pid_file = config.BASE_DIR / ".monitor_pid"
        pid_file.write_text(str(os.getpid()))

        print(f"\nOJCommerce Email Monitor — watching every {interval}s")
        print(f"Monitoring: {config.IMAP_USERNAME}")
        print("Press Ctrl+C to stop.\n")

        try:
            while True:
                try:
                    run_check(verbose=True)
                except imaplib.IMAP4.error as e:
                    print(f"  IMAP error: {e} — will retry next cycle")
                except Exception as e:
                    print(f"  Unexpected error: {e} — will retry next cycle")

                print(f"  Next check in {interval}s...\n")
                time.sleep(interval)
        finally:
            # Clean up PID file on exit
            if pid_file.exists():
                pid_file.unlink()


if __name__ == "__main__":
    main()
