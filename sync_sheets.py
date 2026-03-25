"""
sync_sheets.py
--------------
Syncs all tabs from a source Google Spreadsheet (viewer access only)
to a destination spreadsheet (full access) using OAuth2 + Refresh Token.

Runs every 24 hours via GitHub Actions. On failure, sends an alert email
via Gmail SMTP with the full error traceback.
"""

import os
import sys
import json
import time
import smtplib
import logging
import traceback
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Configuration (all values come from environment variables / GitHub Secrets)
# ---------------------------------------------------------------------------
SOURCE_SPREADSHEET_ID = os.environ["SOURCE_SPREADSHEET_ID"]
DEST_SPREADSHEET_ID = os.environ["DEST_SPREADSHEET_ID"]

# OAuth credentials (stored as JSON strings in GitHub Secrets)
CREDENTIALS_JSON = os.environ["GOOGLE_CREDENTIALS_JSON"]   # contents of credentials.json
TOKEN_JSON = os.environ["GOOGLE_TOKEN_JSON"]                # contents of token.json

# Gmail SMTP alert settings
SMTP_SENDER = os.environ["SMTP_SENDER"]          # e.g. yourname@gmail.com
SMTP_PASSWORD = os.environ["SMTP_APP_PASSWORD"]  # Gmail App Password (16 chars)
SMTP_RECIPIENT = os.environ["SMTP_RECIPIENT"]    # alert destination address

# Rate-limit guard: seconds to wait between tab syncs
SLEEP_BETWEEN_TABS = int(os.environ.get("SLEEP_BETWEEN_TABS", "5"))

# Google API scopes required
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]


# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------
def build_credentials() -> Credentials:
    """
    Build and (if necessary) refresh OAuth2 credentials from the JSON blobs
    stored in GitHub Secrets.
    """
    token_data = json.loads(TOKEN_JSON)
    cred_data = json.loads(CREDENTIALS_JSON)

    # The installed-app client_id / client_secret live inside credentials.json
    client_config = cred_data.get("installed") or cred_data.get("web")

    creds = Credentials(
        token=token_data.get("token"),
        refresh_token=token_data["refresh_token"],
        token_uri=token_data.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=token_data.get("client_id") or client_config["client_id"],
        client_secret=token_data.get("client_secret") or client_config["client_secret"],
        scopes=token_data.get("scopes", SCOPES),
    )

    # Refresh if the access token is expired
    if not creds.valid:
        log.info("Access token expired – refreshing via refresh_token …")
        creds.refresh(Request())
        log.info("Token refreshed successfully.")

    return creds


# ---------------------------------------------------------------------------
# Email alert
# ---------------------------------------------------------------------------
def send_error_email(error_message: str) -> None:
    """Send a plain-text email with the error traceback via Gmail SMTP."""
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = "⚠️ Google Sheets Sync Failed"
        msg["From"] = SMTP_SENDER
        msg["To"] = SMTP_RECIPIENT

        body = (
            "The automated Google Sheets synchronisation job has FAILED.\n\n"
            "=== Error Details ===\n\n"
            f"{error_message}\n\n"
            "Please check the GitHub Actions logs for the full run output."
        )
        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(SMTP_SENDER, SMTP_PASSWORD)
            server.sendmail(SMTP_SENDER, SMTP_RECIPIENT, msg.as_string())

        log.info("Alert email sent to %s", SMTP_RECIPIENT)
    except Exception:  # noqa: BLE001
        # Don't let the mailer hide the original error
        log.error("Failed to send alert email:\n%s", traceback.format_exc())


# ---------------------------------------------------------------------------
# Sheet helpers
# ---------------------------------------------------------------------------
def get_or_create_worksheet(
    dest_spreadsheet: gspread.Spreadsheet, title: str
) -> gspread.Worksheet:
    """
    Return the destination worksheet with *title*, creating it if absent.
    """
    existing_titles = [ws.title for ws in dest_spreadsheet.worksheets()]
    if title in existing_titles:
        return dest_spreadsheet.worksheet(title)

    log.info("  Tab '%s' not found in destination – creating it …", title)
    return dest_spreadsheet.add_worksheet(title=title, rows=5000, cols=50)


def sync_worksheet(
    src_ws: gspread.Worksheet,
    dest_ws: gspread.Worksheet,
) -> None:
    """
    Fetch all values from *src_ws* and overwrite *dest_ws* (values only).
    Uses a single batch call to stay within API quotas.
    """
    log.info("  Fetching data from source tab '%s' …", src_ws.title)
    data: list[list] = src_ws.get_all_values()

    if not data:
        log.info("  Source tab is empty – clearing destination and skipping.")
        dest_ws.clear()
        return

    row_count = len(data)
    col_count = max(len(row) for row in data)
    log.info("  %d rows × %d cols – writing to destination …", row_count, col_count)

    # Clear destination first so stale rows/columns are removed
    dest_ws.clear()

    # Resize destination to fit (avoids "exceeds grid limits" errors)
    dest_ws.resize(rows=max(row_count, 1), cols=max(col_count, 1))

    # Single batch update – much faster and kinder to rate limits than
    # row-by-row append or multiple range calls.
    dest_ws.update(data, value_input_option="RAW")

    log.info("  ✓ Tab '%s' synced (%d rows).", src_ws.title, row_count)


# ---------------------------------------------------------------------------
# Main sync routine
# ---------------------------------------------------------------------------
def run_sync() -> None:
    log.info("=== Google Sheets Sync Started ===")

    creds = build_credentials()
    client = gspread.authorize(creds)

    log.info("Opening source spreadsheet: %s", SOURCE_SPREADSHEET_ID)
    source_ss = client.open_by_key(SOURCE_SPREADSHEET_ID)

    log.info("Opening destination spreadsheet: %s", DEST_SPREADSHEET_ID)
    dest_ss = client.open_by_key(DEST_SPREADSHEET_ID)

    source_worksheets = source_ss.worksheets()
    log.info("Found %d tab(s) in source spreadsheet.", len(source_worksheets))

    for idx, src_ws in enumerate(source_worksheets, start=1):
        log.info("[%d/%d] Processing tab: '%s'", idx, len(source_worksheets), src_ws.title)

        dest_ws = get_or_create_worksheet(dest_ss, src_ws.title)
        sync_worksheet(src_ws, dest_ws)

        # Brief pause to respect Google Sheets API rate limits (100 req / 100 s)
        if idx < len(source_worksheets):
            log.info("  Sleeping %ds before next tab …", SLEEP_BETWEEN_TABS)
            time.sleep(SLEEP_BETWEEN_TABS)

    log.info("=== Sync Completed Successfully ===")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    try:
        run_sync()
    except Exception:  # noqa: BLE001
        error_details = traceback.format_exc()
        log.error("SYNC FAILED:\n%s", error_details)
        send_error_email(error_details)
        sys.exit(1)
