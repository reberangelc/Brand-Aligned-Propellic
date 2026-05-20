#!/usr/bin/env python3
"""
Run once to authorize Google Drive access for Propellic skills.
Saves token to ~/.propellic/token.json.
"""
import os

SCOPES      = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/documents']
CREDS_PATH  = os.path.expanduser('~/.propellic/credentials.json')
TOKEN_PATH  = os.path.expanduser('~/.propellic/token.json')

from google_auth_oauthlib.flow import InstalledAppFlow

flow = InstalledAppFlow.from_client_secrets_file(CREDS_PATH, SCOPES)
creds = flow.run_local_server(port=0)

with open(TOKEN_PATH, 'w') as f:
    f.write(creds.to_json())

print(f"\n✓ Authorization complete. Token saved to {TOKEN_PATH}")
