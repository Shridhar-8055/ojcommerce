"""
OJCommerce Email Outreach — Configuration
Loads SMTP credentials from environment variables (.env file).
"""

import os
from pathlib import Path

# Load .env file if python-dotenv is available
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass  # python-dotenv not installed; rely on actual env vars

# --- SMTP Settings ---
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USERNAME = os.getenv("SMTP_USERNAME", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_USE_TLS = os.getenv("SMTP_USE_TLS", "true").lower() == "true"

# --- Sender Info ---
SENDER_NAME = os.getenv("SENDER_NAME", "Shamique — OJCommerce")
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
SENDER_SIGNATURE = os.getenv("SENDER_SIGNATURE", """Best regards,
Shamique
OJCommerce — Premium Home Furniture & Decor
https://ojcommerce.com""")

# --- Rate Limiting ---
DELAY_BETWEEN_EMAILS = int(os.getenv("DELAY_BETWEEN_EMAILS", "60"))  # seconds
MAX_EMAILS_PER_SESSION = int(os.getenv("MAX_EMAILS_PER_SESSION", "10"))

# --- IMAP Settings (for Hostinger Sent folder monitoring) ---
IMAP_HOST = os.getenv("IMAP_HOST", "imap.hostinger.com")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
IMAP_USERNAME = os.getenv("IMAP_USERNAME", "")
IMAP_PASSWORD = os.getenv("IMAP_PASSWORD", "")
IMAP_SENT_FOLDER = os.getenv("IMAP_SENT_FOLDER", "Sent")
IMAP_POLL_INTERVAL = int(os.getenv("IMAP_POLL_INTERVAL", "300"))  # seconds (5 min)

# --- File Paths ---
BASE_DIR = Path(__file__).parent
PROSPECTS_CSV = BASE_DIR / "prospects.csv"
SENT_LOG_CSV = BASE_DIR / "sent_log.csv"
REPLIES_LOG_CSV = BASE_DIR / "replies_log.csv"
MONITOR_STATE_FILE = BASE_DIR / ".monitor_seen_ids"
