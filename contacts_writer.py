import os
import anthropic
from openpyxl import Workbook

CONTACTS_COLS = ['Name', 'Title', 'Company', 'Role', 'Market',
                 'Last Seen', 'First Seen', 'Appearances', 'Notes', 'Review Flag']

C_NAME        = 1
C_TITLE       = 2
C_COMPANY     = 3
C_ROLE        = 4
C_MARKET      = 5
C_LAST_SEEN   = 6
C_FIRST_SEEN  = 7
C_APPEARANCES = 8
C_NOTES       = 9
C_REVIEW_FLAG = 10

_ai = None

def _client():
    global _ai
    if _ai is None:
        _ai = anthropic.Anthropic(api_key=os.environ.get('ANTHROPIC_API_KEY'))
    return _ai


def _get_or_create_tab(wb, name, columns):
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    ws.append(columns)
    return ws


def _load_index(ws):
    index = {}
    for row_num in range(2, ws.max_row + 1):
        name    = str(ws.cell(row_num, C_NAME).value    or '').strip().lower()
        company = str(ws.cell(row_num, C_COMPANY).value or '').strip().lower()
        if name:
            index[(name, company)] = row_num
    return index


def _cre_names(people_data, narrative):
    """
    people_data: list of (name, title, role, firm)
    Returns set of lowercased names that are CRE industry professionals.
    Falls back to all names if the API call fails.
    """
    if not people_data:
        return set()
    lines = '\n'.join(
        f'- {name} | Title: {title or "unknown"} | Role: {role or "unknown"} | Firm: {firm or "unknown"}'
        for name, title, role, firm in people_data
    )
    try:
        msg = _client().messages.create(
            model='claude-haiku-4-5-20251001',
            max_tokens=512,
            messages=[{
                'role': 'user',
                'content': (
                    f'Article context:\n{narrative}\n\n'
                    f'People mentioned:\n{lines}\n\n'
                    'Which of these are CRE (commercial real estate) industry professionals — '
                    'brokers, investors, developers, lenders, operators, asset managers, '
                    'executives at real estate firms, etc.? '
                    'Exclude politicians, government officials, athletes, celebrities, '
                    'residential real estate agents and brokers (Redfin, Compass, residential RE agents, etc.), '
                    'and anyone not working in the commercial real estate industry. '
                    'Return only the exact names of CRE professionals, one per line, no other text.'
                ),
            }],
        )
        result = [l.strip().lstrip('- ') for l in msg.content[0].text.strip().splitlines() if l.strip()]
        lookup = {n.lower() for n, _, _, _ in people_data}
        return {r.lower() for r in result if r.lower() in lookup}
    except Exception:
        return {n.lower() for n, _, _, _ in people_data}


def _note_entry(date_str, role, title, market):
    snippet = title[:60].rstrip()
    if len(title) > 60:
        snippet += '...'
    return f'{date_str}: {role} — {snippet} ({market or ""})'


def upsert_contacts(date_str: str, articles: list, wb: Workbook) -> int:
    ws = _get_or_create_tab(wb, 'Contacts', CONTACTS_COLS)
    index = _load_index(ws)
    new_count = 0

    for article in articles:
        tx        = (article.get('transaction_type') or '').lower()
        cp        = article.get('companies_people') or []
        title     = article.get('title') or ''
        market    = article.get('market') or ''
        narrative = article.get('narrative') or ''

        if not tx and not any(e.get('people') for e in cp):
            continue

        # Collect all people in this article, filter to CRE professionals
        all_people = [
            (person.get('name', '').strip(),
             person.get('title', '').strip(),
             entry.get('label', '').upper(),
             entry.get('firm_name', '').strip())
            for entry in cp
            for person in (entry.get('people') or [])
            if person.get('name', '').strip()
        ]
        cre_names = _cre_names(all_people, narrative)

        for entry in cp:
            role   = (entry.get('label') or '').upper()
            firm   = (entry.get('firm_name') or '').strip()
            people = entry.get('people') or []

            for person in people:
                name = (person.get('name') or '').strip()
                if not name or name.lower() not in cre_names:
                    continue
                person_title = (person.get('title') or '').strip()
                note = _note_entry(date_str, role, title, market)
                key  = (name.lower(), firm.lower())

                if key in index:
                    row_num = index[key]
                    flagged = False

                    old_title = str(ws.cell(row_num, C_TITLE).value  or '').strip()
                    old_role  = str(ws.cell(row_num, C_ROLE).value   or '').strip()
                    old_mkt   = str(ws.cell(row_num, C_MARKET).value or '').strip()

                    if person_title and person_title != old_title:
                        ws.cell(row_num, C_TITLE).value = person_title
                        flagged = True
                    if role and role != old_role:
                        ws.cell(row_num, C_ROLE).value = role
                        flagged = True
                    if market and market != old_mkt:
                        ws.cell(row_num, C_MARKET).value = market
                        flagged = True

                    ws.cell(row_num, C_LAST_SEEN).value = date_str
                    appearances = int(ws.cell(row_num, C_APPEARANCES).value or 0) + 1
                    ws.cell(row_num, C_APPEARANCES).value = appearances
                    existing_notes = str(ws.cell(row_num, C_NOTES).value or '')
                    ws.cell(row_num, C_NOTES).value = (
                        existing_notes + ' | ' + note if existing_notes else note
                    )
                    if flagged:
                        ws.cell(row_num, C_REVIEW_FLAG).value = 'YES'

                else:
                    name_lower = name.lower()
                    same_name_exists = any(k[0] == name_lower for k in index)
                    review_flag = 'YES' if same_name_exists else ''

                    ws.append([
                        name,
                        person_title,
                        firm,
                        role,
                        market,
                        date_str,
                        date_str,
                        1,
                        note,
                        review_flag,
                    ])
                    index[key] = ws.max_row
                    new_count += 1

    return new_count
