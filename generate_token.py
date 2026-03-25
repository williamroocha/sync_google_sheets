"""
generate_token.py  —  Run this ONCE locally to produce token.json.
-----------------------------------------------------------------------
Prerequisites:
  pip install google-auth-oauthlib

Usage:
  python generate_token.py

It will open a browser window, ask you to log in with your personal
Google account, and write token.json to the current directory.
Copy the contents of that file into the GOOGLE_TOKEN_JSON GitHub Secret.
"""

import json
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

def main():
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
    creds = flow.run_local_server(port=0)

    token_data = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": list(creds.scopes),
    }

    with open("token.json", "w") as f:
        json.dump(token_data, f, indent=2)

    print("\n✅  token.json has been created successfully.")
    print("   Copy its contents into the GOOGLE_TOKEN_JSON GitHub Secret.\n")

if __name__ == "__main__":
    main()
