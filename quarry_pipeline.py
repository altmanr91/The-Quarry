import json
import os
import smtplib
import sys
from datetime import date
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import requests
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook

from comps_writer import append_articles
from contacts_writer import upsert_contacts

load_dotenv()

HANDOFF_URL  = 'https://altmanr91.github.io/CRE-News-Reader/articles_handoff.json'
COMPS_FILE   = Path('comps.xlsx')
CONTACTS_FILE = Path('contacts.xlsx')
RECIPIENT    = 'altmanr91@gmail.com'


def _fetch_handoff() -> dict:
    resp = requests.get(HANDOFF_URL, timeout=30)
    resp.raise_for_status()
    return resp.json()


def _load_or_create(path: Path) -> Workbook:
    if path.exists():
        return load_workbook(path)
    wb = Workbook()
    wb.remove(wb.active)   # remove default blank sheet
    return wb


def _send_email(comps_path: Path, contacts_path: Path, date_str: str, counts: dict, new_contacts: int) -> None:
    password = os.getenv('GMAIL_APP_PASSWORD')
    if not password:
        print('  [email] GMAIL_APP_PASSWORD not set — skipping email')
        return

    msg = MIMEMultipart()
    msg['From']    = RECIPIENT
    msg['To']      = RECIPIENT
    msg['Subject'] = f'The Quarry — {date_str}'

    body = (
        f'The Quarry update for {date_str}.\n\n'
        f'Comps added:  {counts["sales"]} sales, {counts["leases"]} leases, {counts["loans"]} loans\n'
        f'New contacts: {new_contacts}\n\n'
        'Both workbooks attached.'
    )
    msg.attach(MIMEText(body, 'plain'))

    for path in (comps_path, contacts_path):
        with open(path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{path.name}"')
        msg.attach(part)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(RECIPIENT, password)
        server.send_message(msg)
    print('  [email] Sent')


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

    print('Sending email...')
    _send_email(COMPS_FILE, CONTACTS_FILE, date_str, counts, new_contacts)

    print('Done.')


if __name__ == '__main__':
    main()
