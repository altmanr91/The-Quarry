import json
import os
import sys
from datetime import date
from pathlib import Path

import requests
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook

from comps_writer import append_articles
from contacts_writer import upsert_contacts

load_dotenv()

HANDOFF_URL   = 'https://altmanr91.github.io/CRE-News-Reader/articles_handoff.json'
COMPS_FILE    = Path('comps.xlsx')
CONTACTS_FILE = Path('contacts.xlsx')
ONEDRIVE_DIR  = 'Documents/AI Tools/The Quarry'
TOKEN_URL     = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token'
GRAPH_URL     = 'https://graph.microsoft.com/v1.0'


def _fetch_handoff() -> dict:
    resp = requests.get(HANDOFF_URL, timeout=30)
    resp.raise_for_status()
    return resp.json()


def _load_or_create(path: Path) -> Workbook:
    if path.exists():
        return load_workbook(path)
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def _get_access_token(client_id: str, client_secret: str, refresh_token: str) -> str:
    resp = requests.post(TOKEN_URL, data={
        'grant_type':    'refresh_token',
        'refresh_token': refresh_token,
        'client_id':     client_id,
        'client_secret': client_secret,
        'scope':         'offline_access Files.ReadWrite',
    })
    resp.raise_for_status()
    return resp.json()['access_token']


def _download_from_onedrive(comps_path: Path, contacts_path: Path) -> None:
    client_id     = os.getenv('ONEDRIVE_CLIENT_ID')
    client_secret = os.getenv('ONEDRIVE_CLIENT_SECRET')
    refresh_token = os.getenv('ONEDRIVE_REFRESH_TOKEN')

    if not all([client_id, client_secret, refresh_token]):
        print('  [onedrive] Credentials not set — skipping download')
        return

    try:
        access_token = _get_access_token(client_id, client_secret, refresh_token)
    except Exception as e:
        print(f'  [onedrive] Could not get access token — skipping download: {e}')
        return

    headers = {'Authorization': f'Bearer {access_token}'}

    for path in (comps_path, contacts_path):
        url = f'{GRAPH_URL}/me/drive/root:/{ONEDRIVE_DIR}/{path.name}:/content'
        resp = requests.get(url, headers=headers)
        if resp.status_code == 404:
            print(f'  [onedrive] {path.name} not found — will create fresh')
            continue
        resp.raise_for_status()
        path.write_bytes(resp.content)
        print(f'  [onedrive] Downloaded {path.name}')


def _upload_to_onedrive(comps_path: Path, contacts_path: Path) -> None:
    client_id     = os.getenv('ONEDRIVE_CLIENT_ID')
    client_secret = os.getenv('ONEDRIVE_CLIENT_SECRET')
    refresh_token = os.getenv('ONEDRIVE_REFRESH_TOKEN')

    if not all([client_id, client_secret, refresh_token]):
        print('  [onedrive] Credentials not set — skipping upload')
        return

    access_token = _get_access_token(client_id, client_secret, refresh_token)
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type':  'application/octet-stream',
    }

    for path in (comps_path, contacts_path):
        with open(path, 'rb') as f:
            data = f.read()
        upload_url = f'{GRAPH_URL}/me/drive/root:/{ONEDRIVE_DIR}/{path.name}:/content'
        resp = requests.put(upload_url, headers=headers, data=data)
        resp.raise_for_status()
        print(f'  [onedrive] Uploaded {path.name}')


def main() -> None:
    print('Fetching handoff JSON...')
    try:
        handoff = _fetch_handoff()
    except Exception as e:
        print(f'ERROR: Could not fetch handoff: {e}')
        sys.exit(1)

    date_str = handoff.get('date') or date.today().isoformat()
    articles = handoff.get('articles') or []
    print(f'  {len(articles)} articles for {date_str}')

    print('Downloading workbooks from OneDrive...')
    _download_from_onedrive(COMPS_FILE, CONTACTS_FILE)

    print('Loading workbooks...')
    comps_wb    = _load_or_create(COMPS_FILE)
    contacts_wb = _load_or_create(CONTACTS_FILE)

    print('Writing comps...')
    counts = append_articles(date_str, articles, comps_wb)
    print(f'  Sales: {counts["sales"]}, Leases: {counts["leases"]}, Loans: {counts["loans"]}')

    print('Upserting contacts...')
    new_contacts = upsert_contacts(date_str, articles, contacts_wb)
    print(f'  New contacts: {new_contacts}')

    print('Saving workbooks...')
    comps_wb.save(COMPS_FILE)
    contacts_wb.save(CONTACTS_FILE)

    print('Uploading to OneDrive...')
    _upload_to_onedrive(COMPS_FILE, CONTACTS_FILE)

    print('Done.')


if __name__ == '__main__':
    main()
