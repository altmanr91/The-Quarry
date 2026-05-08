from openpyxl import Workbook
from openpyxl.styles import PatternFill

YELLOW = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

FMT_DOLLARS = '$#,##0'
FMT_COMMA   = '#,##0'
FMT_PCT     = '0.0"%"'


def _fmt(ws, row, col_fmt_pairs):
    for col, fmt in col_fmt_pairs:
        ws.cell(row, col).number_format = fmt

SALES_COLS  = ['Date', 'Property Name', 'Address', 'Market', 'Property Type',
               'Size (SF)', 'Units', 'Sale Price', '$/SF', '$/Unit',
               'Year Built', 'Occupancy %', 'Buyer', 'Seller', 'Broker', 'Lender',
               'Source', 'Link', 'Notes']

LEASES_COLS = ['Date', 'Property Name', 'Address', 'Market', 'Property Type',
               'Size (SF)', 'Rent ($/SF/yr)', 'Tenants', 'Landlord', 'Tenant Rep', 'Landlord Broker',
               'Source', 'Link']

LOANS_COLS  = ['Date', 'Property Name', 'Address', 'Market', 'Property Type',
               'Size (SF)', 'Units', 'Loan Amount', 'Loan/SF', 'Loan/Unit',
               'Borrower/Sponsor', 'Lender', 'Source', 'Link', 'Notes']

SALE_TYPES  = {'sale', 'acquisition'}
LEASE_TYPES = {'lease'}
LOAN_TYPES  = {'loan', 'refinance'}


def _get_or_create_tab(wb, name, columns):
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    ws.append(columns)
    return ws


def _firms(companies_people, *labels):
    labels_up = {l.upper() for l in labels}
    return ', '.join(
        e['firm_name'] for e in companies_people
        if e.get('label', '').upper() in labels_up and e.get('firm_name')
    )


def _calc(numerator, denominator):
    if numerator and denominator and denominator > 0:
        return round(numerator / denominator, 2)
    return None


def _load_addresses(ws, addr_col_idx):
    """Return dict mapping address_lower -> [row_numbers] for all data rows."""
    result = {}
    for row_num in range(2, ws.max_row + 1):
        addr = str(ws.cell(row_num, addr_col_idx).value or '').strip().lower()
        if addr:
            result.setdefault(addr, []).append(row_num)
    return result


def _highlight_row(ws, row_num):
    for col in range(1, ws.max_column + 1):
        ws.cell(row_num, col).fill = YELLOW


def _check_duplicate(ws, addr_map, address, addr_col_idx):
    """Append new row, highlight it and any prior rows with the same address."""
    new_row = ws.max_row
    addr_lower = (address or '').strip().lower()
    if not addr_lower:
        return
    if addr_lower in addr_map:
        _highlight_row(ws, new_row)
        for prev in addr_map[addr_lower]:
            _highlight_row(ws, prev)
        addr_map[addr_lower].append(new_row)
    else:
        addr_map[addr_lower] = [new_row]


def _purge_no_basis(wb: Workbook) -> int:
    """Remove existing rows that lack a financial basis. Returns count removed."""
    removed = 0
    checks = [
        ('Sales',  8, 5),   # Sale Price, Property Type
        ('Leases', 6, None),
        ('Loans',  8, None),
    ]
    for sheet_name, basis_col, type_col in checks:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for row in range(ws.max_row, 1, -1):  # bottom-up to preserve indices
            no_basis = not ws.cell(row, basis_col).value
            no_type  = type_col and not ws.cell(row, type_col).value
            if no_basis or no_type:
                ws.delete_rows(row)
                removed += 1
    return removed


def append_articles(date_str: str, articles: list, wb: Workbook) -> dict:
    """
    Append new comp rows from articles_handoff.json to the comps workbook.
    Returns counts: {'sales': N, 'leases': N, 'loans': N}
    """
    sales_ws  = _get_or_create_tab(wb, 'Sales',  SALES_COLS)
    leases_ws = _get_or_create_tab(wb, 'Leases', LEASES_COLS)
    loans_ws  = _get_or_create_tab(wb, 'Loans',  LOANS_COLS)

    # Address column index: 3 (C) in all tabs
    sales_addrs  = _load_addresses(sales_ws,  3)
    leases_addrs = _load_addresses(leases_ws, 3)
    loans_addrs  = _load_addresses(loans_ws,  3)

    counts = {'sales': 0, 'leases': 0, 'loans': 0}

    for article in articles:
        tx = (article.get('transaction_type') or '').lower()
        if tx not in SALE_TYPES | LEASE_TYPES | LOAN_TYPES:
            continue

        dp  = article.get('data_points') or {}
        cp  = article.get('companies_people') or []
        addr = dp.get('address') or ''

        if tx in SALE_TYPES:
            if not dp.get('sale_price') or not dp.get('property_type'):
                continue
            sales_ws.append([
                date_str,
                dp.get('property_name'),
                addr,
                article.get('market'),
                dp.get('property_type'),
                dp.get('size_sf'),
                dp.get('size_units'),
                dp.get('sale_price'),
                _calc(dp.get('sale_price'), dp.get('size_sf')),
                _calc(dp.get('sale_price'), dp.get('size_units')),
                dp.get('year_built'),
                dp.get('occupancy'),
                _firms(cp, 'BUYER'),
                _firms(cp, 'SELLER'),
                _firms(cp, 'SELLER BROKER', 'BUYER BROKER'),
                _firms(cp, 'LENDER'),
                article.get('source'),
                article.get('link'),
                article.get('financing'),
            ])
            _check_duplicate(sales_ws, sales_addrs, addr, 3)
            _fmt(sales_ws, sales_ws.max_row, [
                (6, FMT_COMMA),    # Size (SF)
                (8, FMT_DOLLARS),  # Sale Price
                (9, FMT_DOLLARS),  # $/SF
                (10, FMT_DOLLARS), # $/Unit
                (12, FMT_PCT),     # Occupancy %
            ])
            counts['sales'] += 1

        elif tx in LEASE_TYPES:
            if not dp.get('size_sf'):
                continue
            tenants = ', '.join(article.get('tenants') or [])
            leases_ws.append([
                date_str,
                dp.get('property_name'),
                addr,
                article.get('market'),
                dp.get('property_type'),
                dp.get('size_sf'),
                dp.get('rental_rate'),
                tenants,
                _firms(cp, 'LANDLORD', 'OWNER'),
                _firms(cp, 'TENANT REP'),
                _firms(cp, 'LEASING AGENT'),
                article.get('source'),
                article.get('link'),
            ])
            _check_duplicate(leases_ws, leases_addrs, addr, 3)
            _fmt(leases_ws, leases_ws.max_row, [
                (6, FMT_COMMA),    # Size (SF)
                (7, FMT_DOLLARS),  # Rent ($/SF/yr)
            ])
            counts['leases'] += 1

        elif tx in LOAN_TYPES:
            if not dp.get('loan_amount'):
                continue
            loans_ws.append([
                date_str,
                dp.get('property_name'),
                addr,
                article.get('market'),
                dp.get('property_type'),
                dp.get('size_sf'),
                dp.get('size_units'),
                dp.get('loan_amount'),
                _calc(dp.get('loan_amount'), dp.get('size_sf')),
                _calc(dp.get('loan_amount'), dp.get('size_units')),
                _firms(cp, 'SPONSOR', 'DEVELOPER/SPONSOR'),
                _firms(cp, 'LENDER'),
                article.get('source'),
                article.get('link'),
                article.get('financing'),
            ])
            _check_duplicate(loans_ws, loans_addrs, addr, 3)
            _fmt(loans_ws, loans_ws.max_row, [
                (6, FMT_COMMA),    # Size (SF)
                (8, FMT_DOLLARS),  # Loan Amount
                (9, FMT_DOLLARS),  # Loan/SF
                (10, FMT_DOLLARS), # Loan/Unit
            ])
            counts['loans'] += 1

    return counts
