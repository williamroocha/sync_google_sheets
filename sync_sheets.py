"""
sync_sheets.py  (v2 — values + formatting + hyperlinks + filters)
-----------------------------------------------------------------
Syncs all tabs from a source Google Spreadsheet (viewer access only)
to a destination spreadsheet (full access) using OAuth2 + Refresh Token.

What is copied per tab
──────────────────────
  ① Cell values          — via values.update (FORMULA render) so that
                           =HYPERLINK() and all formulas transfer verbatim.
  ② Cell formatting      — background colour, fonts, bold/italic, borders,
                           alignment, number formats, merged cells.
                           Done with the "temp-sheet" trick:
                             • sheets.copyTo() copies the source tab INTO
                               the destination file as a hidden temp sheet.
                             • CopyPasteRequest (PASTE_FORMAT) moves the
                               formatting from the temp sheet to the real
                               destination tab in one batchUpdate RPC.
                             • The temp sheet is deleted.
  ③ Data validation      — same temp-sheet pass, PASTE_DATA_VALIDATION.
  ④ Column widths /
     row heights         — read from source gridData (includeGridData=True),
                           applied via updateDimensionProperties batchUpdate.
  ⑤ Frozen panes         — updateSheetProperties batchUpdate.
  ⑥ Basic filter         — clearBasicFilter → setBasicFilter batchUpdate.
  ⑦ Filter views         — delete all on destination, recreate from source.

Rate-limit strategy for ~24 000 rows / 8 tabs
──────────────────────────────────────────────
  • Each sub-step is a *single* API call (or a chunked batchUpdate).
  • A configurable sleep (default 6 s) between tabs keeps the script
    well inside Google's 100-requests / 100-seconds quota.
  • Dimension updates are chunked at 200 per batchUpdate to stay
    under the 10 MB request body limit.
"""

from __future__ import annotations

import os
import sys
import json
import time
import smtplib
import logging
import traceback
from typing import Any
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from googleapiclient.discovery import build as google_build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# Configuration  (all values injected via environment variables / GH Secrets)
# ─────────────────────────────────────────────────────────────────────────────
SOURCE_SPREADSHEET_ID = os.environ["SOURCE_SPREADSHEET_ID"]
DEST_SPREADSHEET_ID   = os.environ["DEST_SPREADSHEET_ID"]

CREDENTIALS_JSON = os.environ["GOOGLE_CREDENTIALS_JSON"]   # full credentials.json
TOKEN_JSON       = os.environ["GOOGLE_TOKEN_JSON"]          # full token.json

SMTP_SENDER    = os.environ["SMTP_SENDER"]
SMTP_PASSWORD  = os.environ["SMTP_APP_PASSWORD"]
SMTP_RECIPIENT = os.environ["SMTP_RECIPIENT"]

SLEEP_BETWEEN_TABS = int(os.environ.get("SLEEP_BETWEEN_TABS", "6"))

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ─────────────────────────────────────────────────────────────────────────────
# Authentication
# ─────────────────────────────────────────────────────────────────────────────

def build_credentials() -> Credentials:
    """Reconstruct OAuth2 credentials from Secrets and refresh if needed."""
    token_data = json.loads(TOKEN_JSON)
    cred_data  = json.loads(CREDENTIALS_JSON)
    client_cfg = cred_data.get("installed") or cred_data.get("web")

    creds = Credentials(
        token=token_data.get("token"),
        refresh_token=token_data["refresh_token"],
        token_uri=token_data.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=token_data.get("client_id") or client_cfg["client_id"],
        client_secret=token_data.get("client_secret") or client_cfg["client_secret"],
        scopes=token_data.get("scopes", SCOPES),
    )
    if not creds.valid:
        log.info("Access token expired — refreshing …")
        creds.refresh(Request())
        log.info("Token refreshed successfully.")
    return creds


# ─────────────────────────────────────────────────────────────────────────────
# Email alert on failure
# ─────────────────────────────────────────────────────────────────────────────

def send_error_email(error_message: str) -> None:
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = "⚠️ Google Sheets Sync Failed"
        msg["From"]    = SMTP_SENDER
        msg["To"]      = SMTP_RECIPIENT
        body = (
            "The automated Google Sheets sync job has FAILED.\n\n"
            "=== Error Details ===\n\n"
            f"{error_message}\n\n"
            "Check the GitHub Actions logs for the full run output."
        )
        msg.attach(MIMEText(body, "plain"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
            srv.login(SMTP_SENDER, SMTP_PASSWORD)
            srv.sendmail(SMTP_SENDER, SMTP_RECIPIENT, msg.as_string())
        log.info("Alert email sent to %s", SMTP_RECIPIENT)
    except Exception:
        log.error("Could not send alert email:\n%s", traceback.format_exc())


# ─────────────────────────────────────────────────────────────────────────────
# Sheets API helpers
# ─────────────────────────────────────────────────────────────────────────────

def sheets_service(creds: Credentials):
    return google_build("sheets", "v4", credentials=creds, cache_discovery=False)


def full_grid(sheet_id: int, row_count: int, col_count: int) -> dict:
    """Return a GridRange covering the entire sheet."""
    return {
        "sheetId":          sheet_id,
        "startRowIndex":    0,
        "endRowIndex":      max(row_count, 1),
        "startColumnIndex": 0,
        "endColumnIndex":   max(col_count, 1),
    }


def get_metadata(service, spreadsheet_id: str, include_grid_data: bool = False,
                 ranges: list[str] | None = None) -> dict:
    kwargs: dict[str, Any] = {
        "spreadsheetId":   spreadsheet_id,
        "includeGridData": include_grid_data,
    }
    if ranges:
        kwargs["ranges"] = ranges
    return service.spreadsheets().get(**kwargs).execute()


def sheet_props(metadata: dict, sheet_id: int) -> dict | None:
    for s in metadata.get("sheets", []):
        if s["properties"]["sheetId"] == sheet_id:
            return s
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Step ① — Values  (preserves =HYPERLINK() and all formulas)
# ─────────────────────────────────────────────────────────────────────────────

def sync_values(service, src_id: str, title: str, dest_id: str) -> int:
    """
    Read all values from the source tab (FORMULA mode — returns raw formulas
    including =HYPERLINK(url, label)) and write them to the destination tab.
    Returns the number of rows written.
    """
    log.info("    ① Syncing values …")
    sheet_range = f"'{title}'"

    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=src_id,
            range=sheet_range,
            valueRenderOption="FORMULA",
            dateTimeRenderOption="FORMATTED_STRING",
        )
        .execute()
    )
    values: list[list] = result.get("values", [])

    # Always clear destination first (removes stale data beyond new extent)
    service.spreadsheets().values().clear(
        spreadsheetId=dest_id, range=sheet_range, body={}
    ).execute()

    if not values:
        log.info("    Source tab is empty — destination cleared.")
        return 0

    service.spreadsheets().values().update(
        spreadsheetId=dest_id,
        range=sheet_range,
        valueInputOption="USER_ENTERED",   # re-evaluates formulas / hyperlinks
        body={"values": values},
    ).execute()
    log.info("    Wrote %d rows.", len(values))
    return len(values)


# ─────────────────────────────────────────────────────────────────────────────
# Step ② + ③ — Formatting + data validation  (temp-sheet copy method)
# ─────────────────────────────────────────────────────────────────────────────

def copy_format_and_validation(
    service,
    src_id: str,
    src_sheet_id: int,
    dest_id: str,
    dest_sheet_id: int,
    row_count: int,
    col_count: int,
) -> None:
    """
    Cross-spreadsheet formatting copy:
      1. sheets.copyTo() — duplicates the source sheet (with ALL formatting)
         into the destination file as a temporary sheet.
      2. CopyPasteRequest (PASTE_FORMAT) — copies formatting from temp → dest.
      3. CopyPasteRequest (PASTE_DATA_VALIDATION) — copies validation rules.
      4. deleteSheet — removes the temporary sheet.

    This is the only method supported by the public Sheets API for
    cross-file formatting that does NOT iterate cell by cell.
    """
    log.info("    ② Copying formatting (temp-sheet method) …")

    # 1. Copy source sheet into destination file
    resp = (
        service.spreadsheets()
        .sheets()
        .copyTo(
            spreadsheetId=src_id,
            sheetId=src_sheet_id,
            body={"destinationSpreadsheetId": dest_id},
        )
        .execute()
    )
    temp_id    = resp["sheetId"]
    temp_title = resp["title"]
    log.info("    Temp sheet '%s' (id=%d) created in destination.", temp_title, temp_id)

    src_grid  = full_grid(temp_id,      row_count, col_count)
    dest_grid = full_grid(dest_sheet_id, row_count, col_count)

    # 2+3. Paste format + validation in one batchUpdate
    service.spreadsheets().batchUpdate(
        spreadsheetId=dest_id,
        body={
            "requests": [
                {
                    "copyPaste": {
                        "source":           src_grid,
                        "destination":      dest_grid,
                        "pasteType":        "PASTE_FORMAT",
                        "pasteOrientation": "NORMAL",
                    }
                },
                {
                    "copyPaste": {
                        "source":           src_grid,
                        "destination":      dest_grid,
                        "pasteType":        "PASTE_DATA_VALIDATION",
                        "pasteOrientation": "NORMAL",
                    }
                },
            ]
        },
    ).execute()
    log.info("    Format + validation applied.")

    # 4. Delete temp sheet
    service.spreadsheets().batchUpdate(
        spreadsheetId=dest_id,
        body={"requests": [{"deleteSheet": {"sheetId": temp_id}}]},
    ).execute()
    log.info("    Temp sheet deleted.")


# ─────────────────────────────────────────────────────────────────────────────
# Step ④ — Column widths + row heights
# ─────────────────────────────────────────────────────────────────────────────

def sync_dimensions(
    service,
    src_id: str,
    src_sheet_id: int,
    dest_id: str,
    dest_sheet_id: int,
    title: str,
) -> None:
    """Read per-column and per-row pixel sizes from the source and apply them
    to the destination in chunked batchUpdate calls."""
    log.info("    ④ Syncing column widths and row heights …")

    # includeGridData=True returns columnMetadata / rowMetadata with pixelSize
    meta = get_metadata(service, src_id, include_grid_data=True,
                        ranges=[f"'{title}'"])

    src_block = next(
        (s for s in meta["sheets"] if s["properties"]["sheetId"] == src_sheet_id),
        None,
    )
    if not src_block:
        log.warning("    Could not locate source sheet block — skipping dimensions.")
        return

    grid_data = src_block.get("data", [{}])[0] if src_block.get("data") else {}
    requests: list[dict] = []

    for idx, col in enumerate(grid_data.get("columnMetadata", [])):
        if col.get("pixelSize"):
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId":    dest_sheet_id,
                        "dimension":  "COLUMNS",
                        "startIndex": idx,
                        "endIndex":   idx + 1,
                    },
                    "properties": {"pixelSize": col["pixelSize"]},
                    "fields": "pixelSize",
                }
            })

    for idx, row in enumerate(grid_data.get("rowMetadata", [])):
        if row.get("pixelSize"):
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId":    dest_sheet_id,
                        "dimension":  "ROWS",
                        "startIndex": idx,
                        "endIndex":   idx + 1,
                    },
                    "properties": {"pixelSize": row["pixelSize"]},
                    "fields": "pixelSize",
                }
            })

    if not requests:
        log.info("    No custom column/row sizes found.")
        return

    chunk = 200   # stay well under the 10 MB batchUpdate body limit
    for i in range(0, len(requests), chunk):
        service.spreadsheets().batchUpdate(
            spreadsheetId=dest_id,
            body={"requests": requests[i : i + chunk]},
        ).execute()
    log.info("    Applied %d dimension updates.", len(requests))


# ─────────────────────────────────────────────────────────────────────────────
# Step ⑤ — Frozen panes
# ─────────────────────────────────────────────────────────────────────────────

def sync_frozen_panes(
    service,
    src_sheet_block: dict,
    dest_id: str,
    dest_sheet_id: int,
) -> None:
    gp = src_sheet_block["properties"].get("gridProperties", {})
    frozen_rows = gp.get("frozenRowCount", 0)
    frozen_cols = gp.get("frozenColumnCount", 0)
    if not frozen_rows and not frozen_cols:
        return
    log.info("    ⑤ Setting frozen panes: %d row(s), %d col(s).", frozen_rows, frozen_cols)
    service.spreadsheets().batchUpdate(
        spreadsheetId=dest_id,
        body={
            "requests": [{
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": dest_sheet_id,
                        "gridProperties": {
                            "frozenRowCount":    frozen_rows,
                            "frozenColumnCount": frozen_cols,
                        },
                    },
                    "fields": (
                        "gridProperties.frozenRowCount,"
                        "gridProperties.frozenColumnCount"
                    ),
                }
            }]
        },
    ).execute()


# ─────────────────────────────────────────────────────────────────────────────
# Step ⑥ — Basic filter
# ─────────────────────────────────────────────────────────────────────────────

def sync_basic_filter(
    service,
    src_sheet_block: dict,
    dest_id: str,
    dest_sheet_id: int,
    row_count: int,
    col_count: int,
) -> None:
    src_filter = src_sheet_block.get("basicFilter")
    requests: list[dict] = [{"clearBasicFilter": {"sheetId": dest_sheet_id}}]

    if src_filter:
        log.info("    ⑥ Re-applying basic filter …")
        new_filter: dict[str, Any] = {
            "range": full_grid(dest_sheet_id, row_count, col_count),
        }
        if "sortSpecs"   in src_filter: new_filter["sortSpecs"]   = src_filter["sortSpecs"]
        if "filterSpecs" in src_filter: new_filter["filterSpecs"] = src_filter["filterSpecs"]
        requests.append({"setBasicFilter": {"filter": new_filter}})
    else:
        log.info("    ⑥ No basic filter on source — cleared on destination.")

    service.spreadsheets().batchUpdate(
        spreadsheetId=dest_id,
        body={"requests": requests},
    ).execute()


# ─────────────────────────────────────────────────────────────────────────────
# Step ⑦ — Filter views
# ─────────────────────────────────────────────────────────────────────────────

def sync_filter_views(
    service,
    src_sheet_block: dict,
    dest_id: str,
    dest_sheet_id: int,
    row_count: int,
    col_count: int,
) -> None:
    # Delete all filter views currently on the destination tab
    dest_meta  = get_metadata(service, dest_id)
    dest_block = next(
        (s for s in dest_meta["sheets"] if s["properties"]["sheetId"] == dest_sheet_id),
        None,
    )
    del_reqs: list[dict] = []
    if dest_block:
        for fv in dest_block.get("filterViews", []):
            del_reqs.append({"deleteFilterView": {"filterId": fv["filterViewId"]}})
    if del_reqs:
        service.spreadsheets().batchUpdate(
            spreadsheetId=dest_id, body={"requests": del_reqs}
        ).execute()

    src_views = src_sheet_block.get("filterViews", [])
    if not src_views:
        log.info("    ⑦ No filter views on source tab.")
        return

    log.info("    ⑦ Recreating %d filter view(s) …", len(src_views))
    add_reqs: list[dict] = []
    for fv in src_views:
        new_fv: dict[str, Any] = {
            "title": fv.get("title", "Filter View"),
            "range": full_grid(dest_sheet_id, row_count, col_count),
        }
        if "sortSpecs"   in fv: new_fv["sortSpecs"]   = fv["sortSpecs"]
        if "filterSpecs" in fv: new_fv["filterSpecs"] = fv["filterSpecs"]
        add_reqs.append({"addFilterView": {"filter": new_fv}})

    service.spreadsheets().batchUpdate(
        spreadsheetId=dest_id, body={"requests": add_reqs}
    ).execute()
    log.info("    Filter views recreated.")


# ─────────────────────────────────────────────────────────────────────────────
# Per-tab orchestrator
# ─────────────────────────────────────────────────────────────────────────────

def sync_tab(
    service,
    src_id: str,
    src_sheet_block: dict,
    dest_id: str,
    dest_sheet_id: int,
) -> None:
    src_sheet_id = src_sheet_block["properties"]["sheetId"]
    title        = src_sheet_block["properties"]["title"]
    gp           = src_sheet_block["properties"].get("gridProperties", {})
    row_count    = gp.get("rowCount", 1000)
    col_count    = gp.get("columnCount", 26)

    # ① Values (preserves hyperlink formulas)
    rows_written = sync_values(service, src_id, title, dest_id)
    if rows_written == 0:
        return   # empty tab — nothing visual to copy

    # ② + ③ Formatting + data validation
    copy_format_and_validation(
        service, src_id, src_sheet_id,
        dest_id, dest_sheet_id,
        row_count, col_count,
    )

    # ④ Column widths + row heights
    sync_dimensions(service, src_id, src_sheet_id, dest_id, dest_sheet_id, title)

    # ⑤ Frozen panes
    sync_frozen_panes(service, src_sheet_block, dest_id, dest_sheet_id)

    # ⑥ Basic filter
    sync_basic_filter(service, src_sheet_block, dest_id, dest_sheet_id,
                      row_count, col_count)

    # ⑦ Filter views
    sync_filter_views(service, src_sheet_block, dest_id, dest_sheet_id,
                      row_count, col_count)

    log.info("  ✓ Tab '%s' fully synced (values + formatting + filters).", title)


# ─────────────────────────────────────────────────────────────────────────────
# Destination sheet management
# ─────────────────────────────────────────────────────────────────────────────

def ensure_dest_sheet(
    service,
    dest_id: str,
    dest_meta: dict,
    src_sheet_block: dict,
) -> int:
    """Return sheetId of the matching destination tab, creating it if needed."""
    title = src_sheet_block["properties"]["title"]
    for s in dest_meta["sheets"]:
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]

    log.info("  Creating new tab '%s' in destination …", title)
    gp = src_sheet_block["properties"].get("gridProperties", {})
    resp = service.spreadsheets().batchUpdate(
        spreadsheetId=dest_id,
        body={
            "requests": [{
                "addSheet": {
                    "properties": {
                        "title": title,
                        "gridProperties": {
                            "rowCount":    max(gp.get("rowCount", 1000), 1),
                            "columnCount": max(gp.get("columnCount", 26), 1),
                        },
                    }
                }
            }]
        },
    ).execute()
    new_id = resp["replies"][0]["addSheet"]["properties"]["sheetId"]
    log.info("  Tab '%s' created with sheetId=%d", title, new_id)
    return new_id


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def run_sync() -> None:
    log.info("=== Google Sheets Sync v2 (values + formatting + filters) ===")

    creds   = build_credentials()
    service = sheets_service(creds)

    log.info("Reading source metadata …")
    src_meta = get_metadata(service, SOURCE_SPREADSHEET_ID)

    log.info("Reading destination metadata …")
    dest_meta = get_metadata(service, DEST_SPREADSHEET_ID)

    src_sheets = src_meta["sheets"]
    log.info("Source contains %d tab(s).", len(src_sheets))

    for idx, src_block in enumerate(src_sheets, start=1):
        title = src_block["properties"]["title"]
        log.info("[%d/%d] ══ Processing tab: '%s'", idx, len(src_sheets), title)

        # Ensure the destination tab exists; refresh dest_meta afterwards
        dest_sheet_id = ensure_dest_sheet(service, DEST_SPREADSHEET_ID, dest_meta, src_block)
        dest_meta     = get_metadata(service, DEST_SPREADSHEET_ID)   # keep IDs fresh

        sync_tab(
            service,
            src_id=SOURCE_SPREADSHEET_ID,
            src_sheet_block=src_block,
            dest_id=DEST_SPREADSHEET_ID,
            dest_sheet_id=dest_sheet_id,
        )

        if idx < len(src_sheets):
            log.info("  Sleeping %ds before next tab …", SLEEP_BETWEEN_TABS)
            time.sleep(SLEEP_BETWEEN_TABS)

    log.info("=== Sync Completed Successfully ===")


if __name__ == "__main__":
    try:
        run_sync()
    except Exception:
        err = traceback.format_exc()
        log.error("SYNC FAILED:\n%s", err)
        send_error_email(err)
        sys.exit(1)
