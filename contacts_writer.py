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


def _get_or_create_tab(wb, name, columns):
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    ws.append(columns)
    return ws


def _load_index(ws):
    """Return dict mapping (name_lower, company_lower) -> row_number."""
    index = {}
    for row_num in range(2, ws.max_row + 1):
        name    = str(ws.cell(row_num, C_NAME).value    or '').strip().lower()
        company = str(ws.cell(row_num, C_COMPANY).value or '').strip().lower()
        if name:
            index[(name, company)] = row_num
    return index


def _note_entry(date_str, role, title, market):
    snippet = title[:60].rstrip()
    if len(title) > 60:
        snippet += '...'
    return f'{date_str}: {role} — {snippet} ({market or ""})'


def upsert_contacts(date_str: str, articles: list, wb: Workbook) -> int:
    """
    Upsert contacts from articles into the Contacts tab.
    Returns count of new contacts added.
    """
    ws = _get_or_create_tab(wb, 'Contacts', CONTACTS_COLS)
    index = _load_index(ws)
    new_count = 0

    for article in articles:
        tx = (article.get('transaction_type') or '').lower()
        cp = article.get('companies_people') or []
        title  = article.get('title') or ''
        market = article.get('market') or ''

        # Skip non-transaction articles that have no people data
        if not tx and not any(e.get('people') for e in cp):
            continue

        for entry in cp:
            role   = (entry.get('label') or '').upper()
            firm   = (entry.get('firm_name') or '').strip()
            people = entry.get('people') or []

            for person in people:
                name = (person.get('name') or '').strip()
                if not name:
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
                    # Name match at a different firm → possible firm change
                    name_lower = name.lower()
                    same_name_exists = any(k[0] == name_lower for k in index)
                    review_flag = 'YES' if same_name_exists else ''

                    ws.append([
                        name,
                        person_title,
                        firm,
                        role,
                        market,
                        date_str,    # Last Seen
                        date_str,    # First Seen
                        1,           # Appearances
                        note,
                        review_flag,
                    ])
                    index[key] = ws.max_row
                    new_count += 1

    return new_count
