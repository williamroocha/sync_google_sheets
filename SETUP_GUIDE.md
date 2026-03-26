# Google Sheets Auto-Sync — Complete Setup Guide

## Project Structure

```
your-repo/
├── sync_sheets.py          # Main sync script
├── generate_token.py       # One-time local token generator
├── requirements.txt
└── .github/
    └── workflows/
        └── main.yml        # GitHub Actions workflow
```

---

## Step 1 — Create a Google Cloud Project and OAuth Credentials

This is the one-time setup that lets you authenticate as *yourself* (no Service Account needed).

1. Go to [https://console.cloud.google.com](https://console.cloud.google.com) and sign in with your **personal** Google account (the one that has access to both spreadsheets).

2. Click the project dropdown (top left) → **New Project**. Name it anything, e.g. `sheets-sync`.

3. In the left sidebar, navigate to **APIs & Services → Library**.

4. Search for and **Enable** both of these APIs:
   - `Google Sheets API`
   - `Google Drive API`

5. Navigate to **APIs & Services → OAuth consent screen**.
   - Choose **External** user type → **Create**.
   - Fill in the **App name** (e.g. "Sheets Sync"), your **User support email**, and **Developer contact email**.
   - Click **Save and Continue** through the Scopes and Test Users screens (no changes needed for now).
   - On the Summary page, click **Back to Dashboard**.

6. Navigate to **APIs & Services → Credentials**.
   - Click **+ Create Credentials → OAuth client ID**.
   - Application type: **Desktop app**.
   - Name: anything (e.g. "sheets-sync-desktop").
   - Click **Create**.

7. In the dialog that appears, click **Download JSON**. Save the file as **`credentials.json`** in your project folder.

> ⚠️ Do **not** commit `credentials.json` to Git. Add it to `.gitignore`.

---

## Step 2 — Generate token.json Locally (One-Time Only)

This step exchanges your `credentials.json` for a long-lived `refresh_token` that the GitHub Actions runner will use to authenticate on your behalf.

### Prerequisites

```bash
pip install google-auth-oauthlib
```

### Run the helper script

Make sure `credentials.json` is in the same folder as `generate_token.py`, then run:

```bash
python generate_token.py
```

- A browser window will open.
- Log in with your **personal Google account** (the one with access to the spreadsheets).
- Click **Allow** on the permissions screen.
- The script writes `token.json` to the current directory and prints a success message.

> ⚠️ The `refresh_token` field in `token.json` is the long-lived credential. Keep it secret — treat it like a password.

---

## Step 3 — Add Secrets to GitHub

All sensitive values are stored as **GitHub Actions Secrets** so they never appear in your code or logs.

1. Open your GitHub repository → **Settings → Secrets and variables → Actions**.

2. Click **New repository secret** for each of the following:

| Secret Name | Value |
|---|---|
| `SOURCE_SPREADSHEET_ID` | The ID from the source sheet URL: `spreadsheets/d/`**`<THIS_PART>`**`/edit` |
| `DEST_SPREADSHEET_ID` | Same pattern for your personal copy |
| `GOOGLE_CREDENTIALS_JSON` | Paste the **entire contents** of `credentials.json` |
| `GOOGLE_TOKEN_JSON` | Paste the **entire contents** of `token.json` |
| `SMTP_SENDER` | Your Gmail address (e.g. `you@gmail.com`) |
| `SMTP_APP_PASSWORD` | Your Gmail App Password (see Step 4) |
| `SMTP_RECIPIENT` | Email address that receives failure alerts |

> **Tip:** To paste a multi-line JSON file as a single secret, open the file in a text editor, select all, copy, and paste directly into the GitHub secret field — no modifications needed.

---

## Step 4 — Set Up a Gmail App Password

Gmail requires an App Password (not your regular password) for SMTP access from scripts.

1. Go to your Google Account: [https://myaccount.google.com](https://myaccount.google.com).
2. Navigate to **Security**.
3. Under "How you sign in to Google", ensure **2-Step Verification** is **On**. (App Passwords require 2FA.)
4. Search for **App Passwords** in the search bar at the top of your account settings (or go to [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)).
5. Enter an app name (e.g. "Sheets Sync Alert") and click **Create**.
6. Copy the 16-character password that appears — you will only see it once.
7. Paste it as the `SMTP_APP_PASSWORD` GitHub Secret.

---

## Step 5 — Push to GitHub and Test

1. Commit and push all files to your repository:

```bash
git add sync_sheets.py generate_token.py requirements.txt .github/
git commit -m "feat: add 24h Google Sheets sync automation"
git push
```

2. Trigger a manual run to verify everything works **before** waiting for the daily schedule:
   - GitHub → **Actions** tab → **Sync Google Sheets (Daily)** → **Run workflow** → **Run workflow**.

3. Watch the logs in real time. A green checkmark means the sync completed successfully.

---

## Troubleshooting

| Symptom | Likely Cause | Fix |
|---|---|---|
| `RefreshError: invalid_grant` | `refresh_token` is expired or revoked | Re-run `generate_token.py` locally and update `GOOGLE_TOKEN_JSON` |
| `403 Forbidden` on source sheet | Your account lost viewer access | Confirm access in Google Sheets directly |
| `429 Too Many Requests` | Rate limit hit | Increase `SLEEP_BETWEEN_TABS` (e.g. to `10`) |
| `SMTPAuthenticationError` | Wrong App Password or 2FA not enabled | Re-generate the App Password |
| Tabs missing in destination | Tab name mismatch or creation failed | Check the Action logs for the specific tab error |

---

## Refreshing the token.json (if it ever expires)

Google refresh tokens for Desktop apps do not expire unless:
- You revoke access in your Google Account settings.
- Your OAuth consent screen app is deleted.
- Your account password changes and you haven't enabled "keep me signed in".

If you ever need to regenerate it, simply re-run `generate_token.py` locally and update the `GOOGLE_TOKEN_JSON` secret on GitHub.

---

## Security Checklist

- [ ] `credentials.json` and `token.json` are in `.gitignore` — never committed.
- [ ] All secrets are stored in GitHub Secrets, not hard-coded.
- [ ] Gmail App Password is stored only in GitHub Secrets.
- [ ] OAuth consent screen is set to **External** but only your personal account is an authorized test user.
