"""
build_dashboard.py — Parse LP Excel reports and regenerate Marketing_Dashboard.html.

Column mappings (0-indexed, verified empirically):
  appt_by_source / appt_by_subsource / appt_by_product:
    data rows  : name=0, Set=2, Sale=4, Demo=31, NetIssue=33, GrossIssue=35, Drop=37
    Total row  : Set=2, Sale=4, Demo=30, NetIssue=32, GrossIssue=34, Drop=36
  appt_by_setter:
    data rows  : name=0, Set=3, GrossIssue=5, Demo=9, Sale=11, GrossAmt=33, Drop=39
    Total row  : Set=3, GrossIssue=5, Demo=8, Sale=11, GrossAmt=32, Drop=39
  marketing_sub_source:
    sub-source rows: name=0, Raw=3, Set=5, Issue=8, Demo=11, Sold=13, GrossSale=18
    Total/GrandTotal: name=0, Raw=2, Set=5, Issue=8, Demo=11, Sold=13, GrossSale=18
    source-group header: name=0 only (col3 is empty string)
"""

import xlrd
from pathlib import Path
from datetime import date, timedelta
from collections import defaultdict

REPORTS_DIR = Path(__file__).parent.parent / "reports"
OUTPUT     = Path(__file__).parent.parent / "Marketing_Dashboard.html"

PRODUCT_LABELS = {
    'FB': 'Full Bathroom', 'WA': 'Wet Area', 'FK': 'Full Kitchen',
    'CAB': 'Cabinets', 'WAIT': 'Walk-in Tub', 'WINDR': 'Windows & Doors',
}
KITCHEN_CODES  = {'CAB', 'FK'}
BATHROOM_CODES = {'FB', 'WA', 'WAIT'}
WINDOW_CODES   = {'WINDR'}

# ── Value helpers ─────────────────────────────────────────────────────────────

def iv(v):
    if v == '' or v is None: return 0
    try: return int(float(str(v)))
    except: return 0

def fv(v):
    if v == '' or v is None: return 0.0
    try: return float(str(v))
    except: return 0.0

def gv(row, col):
    return row[col] if col < len(row) else ''

def pct(num, den, dec=1):
    if not den: return '—'
    return f'{num/den*100:.{dec}f}%'

def dollar(v):
    if v == 0: return '$0'
    return f'${v:,.0f}'

def avg_sale(gross, sales):
    if not sales: return '—'
    return f'${gross/sales:,.0f}'

def rate_class(pct_str):
    if pct_str == '—': return ''
    v = float(pct_str.rstrip('%'))
    if v >= 80: return ' class="g"'
    if v >= 60: return ' class="w"'
    return ' class="b"'

# ── Parsers ───────────────────────────────────────────────────────────────────

def load_sheet(path):
    wb = xlrd.open_workbook(str(path))
    return wb.sheets()[0]

def is_footer(name):
    return any(name.startswith(x) for x in (
        'SNS=', 'CNS=', 'C1.', '6/', 'Page', 'Appointment Stat',
        'For Appoint', 'Set Appoint', 'Market:', 'Source:', 'User:'))

def parse_cc(path):
    """Parse appt_by_source / appt_by_subsource / appt_by_product.
    Each named row is followed by an unnamed pct/dollar sub-row.
    Gross $ (Close $) lives in that sub-row: col 25 for data rows, col 24 for Total.
    """
    ws = load_sheet(path)
    all_rows = [ws.row_values(r) for r in range(ws.nrows)]
    rows, total = [], None
    r = 0
    while r < len(all_rows):
        rv = all_rows[r]
        name = str(rv[0]).strip() if rv[0] != '' else ''
        if not name or is_footer(name):
            r += 1; continue
        if name in ('Product', 'Source', 'Sub-Source', 'Setter'):
            r += 1; continue

        # Peek at next row for the gross $ (dollar sub-row has no name in col 0)
        nxt = all_rows[r + 1] if r + 1 < len(all_rows) else []

        if name == 'Total:':
            total = {
                'set': iv(gv(rv,2)), 'sale': iv(gv(rv,4)),
                'demo': iv(gv(rv,30)), 'net_issue': iv(gv(rv,32)),
                'gross_issue': iv(gv(rv,34)), 'drop': iv(gv(rv,36)),
                'gross_amt': fv(gv(nxt, 24)),   # Close $ on Total sub-row is at col 24
            }
            r += 2  # skip the sub-row we just consumed
        else:
            set_v = iv(gv(rv,2))
            if set_v > 0 or any(isinstance(rv[c], float) for c in [4,31,33,35,37] if c < len(rv)):
                rows.append({
                    'name': name,
                    'set': set_v, 'sale': iv(gv(rv,4)),
                    'demo': iv(gv(rv,31)), 'net_issue': iv(gv(rv,33)),
                    'gross_issue': iv(gv(rv,35)), 'drop': iv(gv(rv,37)),
                    'gross_amt': fv(gv(nxt, 25)),  # Close $ on data sub-row is at col 25
                })
            r += 2  # skip the sub-row
    return rows, total

def parse_setter(path):
    """Parse appt_by_setter.
    Column map (data row / Total row):
      Set=3/3, GrossIssue=5/5, Demo=9/8, Sale=10/11,
      CNS=13/14, SRS=17/18, 1Leg=19/20, CXL=22/23,
      Porched=24/25, NG=27/28, Other=29/30,
      GrossAmt=33/32, SaleCXL=34/34, NetAmt(NSLI)=36/36, Drop=39/39
    Note: Sale col shifts by 1 between data rows and Total row (merged cell artifact).
    """
    ws = load_sheet(path)
    rows, total = [], None
    for r in range(ws.nrows):
        rv = ws.row_values(r)
        name = str(rv[0]).strip() if rv[0] != '' else ''
        if not name or is_footer(name): continue
        if name == 'Setter': continue

        if name == 'Total:':
            total = {
                'set':        iv(gv(rv,3)),
                'gross_issue':iv(gv(rv,5)),
                'demo':       iv(gv(rv,8)),
                'sale':       iv(gv(rv,11)),
                'cns':        iv(gv(rv,14)),
                'one_leg':    iv(gv(rv,20)),
                'cxl':        iv(gv(rv,23)),
                'ng':         iv(gv(rv,28)),
                'gross_amt':  fv(gv(rv,32)),
                'net_amt':    fv(gv(rv,36)),
                'drop':       iv(gv(rv,39)),
            }
        elif iv(gv(rv,3)) > 0 or name:
            rows.append({
                'name':       name,
                'set':        iv(gv(rv,3)),
                'gross_issue':iv(gv(rv,5)),
                'demo':       iv(gv(rv,9)),
                'sale':       iv(gv(rv,10)),   # col10 in data rows (col11 in Total)
                'cns':        iv(gv(rv,13)),
                'one_leg':    iv(gv(rv,19)),
                'cxl':        iv(gv(rv,22)),
                'ng':         iv(gv(rv,27)),
                'gross_amt':  fv(gv(rv,33)),
                'net_amt':    fv(gv(rv,36)),   # NSLI = Net Amt
                'drop':       iv(gv(rv,39)),
            })
    return rows, total

TRACKED_PROMOTERS = {'teresa', 'andrew', 'morgan', 'jazmin'}

def parse_promoter(path):
    """Parse appt_by_promoter — same format as appt_by_setter.
    Returns (rows, total) where rows include net_issue = gross_issue - drop.
    """
    if not path.exists():
        return [], None
    ws = load_sheet(path)
    rows, total = [], None
    for r in range(ws.nrows):
        rv = ws.row_values(r)
        name = str(rv[0]).strip() if rv[0] != '' else ''
        if not name or is_footer(name): continue
        if name in ('Setter', 'Promoter'): continue
        if name == 'Total:':
            gi   = iv(gv(rv,5))
            drop = iv(gv(rv,39))
            total = {
                'set':         iv(gv(rv,3)),
                'gross_issue': gi,
                'net_issue':   gi - drop,
                'demo':        iv(gv(rv,8)),
                'sale':        iv(gv(rv,11)),
                'gross_amt':   fv(gv(rv,32)),
                'drop':        drop,
            }
        elif iv(gv(rv,3)) > 0 or name:
            gi   = iv(gv(rv,5))
            drop = iv(gv(rv,39))
            rows.append({
                'name':        name,
                'set':         iv(gv(rv,3)),
                'gross_issue': gi,
                'net_issue':   gi - drop,
                'demo':        iv(gv(rv,9)),
                'sale':        iv(gv(rv,10)),
                'gross_amt':   fv(gv(rv,33)),
                'drop':        drop,
            })
    return rows, total


def parse_dispo(path):
    """Parse Dispo Distribution report (by setter).
    Row structure: header rows, then pairs of (count row, pct row) per setter, then Totals.
    Col map: name=0, total=2(data)/1(totals), then dispo cols read dynamically from row 6.
    Returns (rows, total, dispo_cols) where dispo_cols = {col_idx: label}.
    """
    ws = load_sheet(path)
    # Read dispo column headers from row 6
    dispo_cols = {}
    if ws.nrows > 6:
        for ci, v in enumerate(ws.row_values(6)):
            s = str(v).strip()
            if s and ci > 1:
                dispo_cols[ci] = s

    rows, total = [], None
    r = 0
    while r < ws.nrows:
        rv = ws.row_values(r)
        name = str(rv[0]).strip() if rv[0] != '' else ''
        # Skip header/footer rows
        if not name or name in ('Setter',) or is_footer(name):
            r += 1
            continue
        if name.lower().startswith('total'):
            total_count = iv(gv(rv, 1))
            dispos = {dispo_cols[ci]: iv(rv[ci]) for ci in dispo_cols if ci < len(rv) and iv(rv[ci]) > 0}
            total = {'name': 'TOTAL', 'total': total_count, 'dispos': dispos}
            r += 2  # skip pct row
            continue
        # Data row: count row followed by pct row
        count = iv(gv(rv, 2))
        dispos = {dispo_cols[ci]: iv(rv[ci]) for ci in dispo_cols if ci < len(rv) and iv(rv[ci]) > 0}
        if count > 0 or dispos:
            rows.append({'name': name, 'total': count, 'dispos': dispos})
        r += 2  # skip the pct row that follows
    return rows, total, dispo_cols

def parse_marketing(path):
    """Parse marketing_sub_source. Returns (groups, grand_total).
    groups = {source_name: {'raw': int, 'subs': [{name,raw,set,issue,demo,sold,gross_sale}]}}
    """
    ws = load_sheet(path)
    groups = {}
    grand_total = None
    cur = None

    for r in range(ws.nrows):
        rv = ws.row_values(r)
        name = str(rv[0]).strip() if rv[0] != '' else ''
        if not name or is_footer(name): continue
        if name in ('Raw', '# Net Close', '% Net Close', 'Net Sale $', 'NSLI'): continue

        col3 = gv(rv, 3)
        col2 = gv(rv, 2)

        if name == 'Grand Total:':
            grand_total = {
                'raw': iv(col2), 'set': iv(gv(rv,5)),
                'issue': iv(gv(rv,8)), 'demo': iv(gv(rv,11)),
                'sold': iv(gv(rv,13)), 'gross_sale': fv(gv(rv,18)),
            }
        elif name == 'Total:':
            if cur:
                groups[cur]['raw'] = iv(col2)
        elif isinstance(col3, float) or isinstance(col3, int):
            # sub-source row
            if cur and cur in groups:
                groups[cur]['subs'].append({
                    'name': name, 'raw': iv(col3),
                    'set': iv(gv(rv,5)), 'issue': iv(gv(rv,8)),
                    'demo': iv(gv(rv,11)), 'sold': iv(gv(rv,13)),
                    'gross_sale': fv(gv(rv,18)),
                })
        elif col3 == '' and col2 == '':
            # source group header
            cur = name
            if cur not in groups:
                groups[cur] = {'raw': 0, 'subs': []}

    return groups, grand_total

# ── Load all periods ──────────────────────────────────────────────────────────

PERIODS = ['prior_week', 'prior_month', 'mtd', 'ytd']

def load_all():
    data = {}
    for period in PERIODS:
        d = REPORTS_DIR / period
        # dispo_distribution is .xls (downloaded directly); others are .xlsx
        dispo_path = d / 'dispo_distribution.xls'
        if not dispo_path.exists():
            dispo_path = d / 'dispo_distribution.xlsx'
        data[period] = {
            'source':    parse_cc(d / 'appt_by_source.xlsx'),
            'subsource': parse_cc(d / 'appt_by_subsource.xlsx'),
            'product':   parse_cc(d / 'appt_by_product.xlsx'),
            'setter':    parse_setter(d / 'appt_by_setter.xlsx'),
            'marketing': parse_marketing(d / 'marketing_sub_source.xlsx'),
            'dispo':     parse_dispo(dispo_path) if dispo_path.exists() else ([], None, {}),
            'promoter':  parse_promoter(d / 'appt_by_promoter.xlsx'),
        }
    return data

# ── Date range labels ─────────────────────────────────────────────────────────

def date_ranges():
    today = date.today()
    days_since_monday = today.weekday()
    last_sat  = today - timedelta(days=days_since_monday + 2)
    last_mon  = last_sat - timedelta(days=5)
    first_of_month = today.replace(day=1)
    pm_end   = first_of_month - timedelta(days=1)
    pm_start = pm_end.replace(day=1)
    def d_str(d, fmt_str):
        # %-d (Linux) → %#d (Windows) for zero-stripped day
        import platform
        f = fmt_str.replace('%-d', '%#d') if platform.system() == 'Windows' else fmt_str
        return d.strftime(f)
    return {
        'prior_week':  (last_mon, last_sat,
                        f'{d_str(last_mon, "%b %-d")} – {d_str(last_sat, "%b %-d")}, {last_sat.year}'),
        'prior_month': (pm_start, pm_end,
                        f'{d_str(pm_start, "%B %-d")} – {d_str(pm_end, "%-d")}, {pm_end.year}'),
        'mtd':         (today.replace(day=1), today,
                        f'{d_str(today.replace(day=1), "%B %-d")} – {d_str(today, "%-d")}, {today.year}'),
        'ytd':         (today.replace(month=1,day=1), today,
                        f'January 1 – {d_str(today, "%B %-d")}, {today.year}'),
    }

# ── HTML helpers ──────────────────────────────────────────────────────────────

def kpi(label, value, sub, source, color=''):
    cls = f'kpi-card {color}' if color else 'kpi-card'
    return (f'<div class="{cls}"><div class="kpi-source">{source}</div>'
            f'<div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>'
            f'<div class="kpi-sub">{sub}</div></div>')

def build_subsource_table(mkt_groups, mkt_grand, ss_rows, ss_total, colspan=12, ytd_source_map=None):
    """Build the combined sub-source HTML table body.
    ytd_source_map: {sub_source_name_lower: source_group_name} built from YTD marketing report
    so CC sub-sources not in this period's marketing still land in the right group.
    """
    # Index CC sub-source data by normalized name
    cc = {r['name'].strip().lower(): r for r in ss_rows}

    # Build a lookup of raw leads by sub-source name from marketing
    mkt_subs = {}  # name.lower() -> raw
    for grp in mkt_groups.values():
        for s in grp['subs']:
            mkt_subs[s['name'].strip().lower()] = s['raw']

    # Build source-group structure
    # Use marketing source groups + catch CC sub-sources not in marketing
    SKIP_COLS = colspan  # number of cols in table

    def td(v, cls=''):
        if cls: return f'<td{cls}>{v}</td>'
        return f'<td>{v}</td>'

    lines = []

    # Track which CC sub-sources we've placed
    placed = set()

    def row_for(cc_row, raw_val):
        """Render one data row. cc_row may be None if no CC data."""
        if cc_row:
            set_v = cc_row['set']
            gi    = cc_row['gross_issue']
            ni    = cc_row['net_issue']
            demo  = cc_row['demo']
            sale  = cc_row['sale']
            drop  = cc_row['drop']
            gross = cc_row.get('gross_amt', 0)
            i_pct = pct(gi, set_v)
            d_pct = pct(demo, gi) if gi else '—'
            c_pct = pct(sale, demo) if demo else '—'
        else:
            set_v = gi = ni = demo = sale = drop = gross = 0
            i_pct = d_pct = c_pct = '—'

        raw_disp = str(raw_val) if raw_val else '—'

        if colspan == 12:
            return (f'<td>&nbsp;&nbsp;{cc_row["name"] if cc_row else "—"}</td>'
                    f'<td class="mc">{raw_disp}</td>'
                    f'{td(set_v)}{td(gi)}'
                    f'<td{rate_class(i_pct)}>{i_pct}</td>'
                    f'{td(ni)}{td(demo)}'
                    f'<td{rate_class(d_pct)}>{d_pct}</td>'
                    f'{td(sale)}'
                    f'<td{rate_class(c_pct)}>{c_pct}</td>'
                    f'{td(drop)}<td>{dollar(gross)}</td>')
        else:
            # shorter colspan=11 (prior week)
            return (f'<td>&nbsp;&nbsp;{cc_row["name"] if cc_row else "—"}</td>'
                    f'<td class="mc">{raw_disp}</td>'
                    f'{td(set_v)}{td(gi)}'
                    f'<td{rate_class(i_pct)}>{i_pct}</td>'
                    f'{td(demo)}'
                    f'<td{rate_class(d_pct)}>{d_pct}</td>'
                    f'{td(sale)}'
                    f'<td{rate_class(c_pct)}>{c_pct}</td>'
                    f'{td(drop)}<td>{dollar(gross)}</td>')

    # Identify marketing source groups and their sub-sources
    for grp_name, grp in mkt_groups.items():
        grp_subs = grp['subs']
        if not grp_subs and grp['raw'] == 0:
            continue

        lines.append(f'<tr class="div-row"><td colspan="{colspan}">{grp_name.upper()}</td></tr>')

        grp_set = grp_cc_gi = grp_ni = grp_demo = grp_sale = grp_drop = grp_raw = grp_gross = 0

        for s in grp_subs:
            key = s['name'].strip().lower()
            cc_row = cc.get(key)
            if cc_row:
                placed.add(key)
            raw_v = s['raw']
            grp_raw += raw_v

            # Build display name for first column
            disp_name = s['name'].strip()
            r_cc = cc_row or {'set':0,'gross_issue':0,'net_issue':0,'demo':0,'sale':0,'drop':0,'gross_amt':0}
            set_v  = r_cc['set'];  gi = r_cc['gross_issue']
            ni     = r_cc.get('net_issue', gi); demo = r_cc['demo']
            sale   = r_cc['sale']; drop = r_cc['drop']
            gross  = r_cc.get('gross_amt', 0)
            grp_set += set_v; grp_cc_gi += gi; grp_ni += ni
            grp_demo += demo; grp_sale += sale; grp_drop += drop; grp_gross += gross

            i_pct = pct(gi, set_v) if set_v else '—'
            d_pct = pct(demo, gi) if gi else '—'
            c_pct = pct(sale, demo) if demo else '—'
            raw_d = str(raw_v) if raw_v else '—'

            if colspan == 12:
                lines.append(
                    f'<tr><td>&nbsp;&nbsp;{disp_name}</td>'
                    f'<td class="mc">{raw_d}</td>'
                    f'{td(set_v)}{td(gi)}'
                    f'<td{rate_class(i_pct)}>{i_pct}</td>'
                    f'{td(ni)}{td(demo)}'
                    f'<td{rate_class(d_pct)}>{d_pct}</td>'
                    f'{td(sale)}<td{rate_class(c_pct)}>{c_pct}</td>'
                    f'{td(drop)}<td>{dollar(gross)}</td></tr>')
            else:
                lines.append(
                    f'<tr><td>&nbsp;&nbsp;{disp_name}</td>'
                    f'<td class="mc">{raw_d}</td>'
                    f'{td(set_v)}{td(gi)}'
                    f'<td{rate_class(i_pct)}>{i_pct}</td>'
                    f'{td(demo)}'
                    f'<td{rate_class(d_pct)}>{d_pct}</td>'
                    f'{td(sale)}<td{rate_class(c_pct)}>{c_pct}</td>'
                    f'{td(drop)}<td>{dollar(gross)}</td></tr>')

        # Group total
        gi_p = pct(grp_cc_gi, grp_set) if grp_set else '—'
        dm_p = pct(grp_demo, grp_ni) if grp_ni else '—'
        cl_p = pct(grp_sale, grp_demo) if grp_demo else '—'
        raw_d = str(grp_raw) if grp_raw else '—'
        if colspan == 12:
            lines.append(
                f'<tr class="grp"><td>&nbsp;&nbsp;{grp_name} Total</td>'
                f'<td class="mc">{raw_d}</td>'
                f'{td(grp_set)}{td(grp_cc_gi)}'
                f'<td>{gi_p}</td>{td(grp_ni)}{td(grp_demo)}'
                f'<td>{dm_p}</td>{td(grp_sale)}<td>{cl_p}</td>'
                f'{td(grp_drop)}<td>{dollar(grp_gross)}</td></tr>')
        else:
            lines.append(
                f'<tr class="grp"><td>&nbsp;&nbsp;{grp_name} Total</td>'
                f'<td class="mc">{raw_d}</td>'
                f'{td(grp_set)}{td(grp_cc_gi)}'
                f'<td>{gi_p}</td>{td(grp_demo)}'
                f'<td>{dm_p}</td>{td(grp_sale)}<td>{cl_p}</td>'
                f'{td(grp_drop)}<td>{dollar(grp_gross)}</td></tr>')

    # CC sub-sources not matched to any marketing group in this period
    unplaced = [r for r in ss_rows if r['name'].strip().lower() not in placed]
    if unplaced:
        # Use YTD source map to place them in the right group; fall back to OTHER
        by_grp = defaultdict(list)
        for r in unplaced:
            grp_name = (ytd_source_map or {}).get(r['name'].strip().lower(), 'Other')
            by_grp[grp_name].append(r)

        for grp_name, rows in sorted(by_grp.items()):
            # If this group already appeared above (had marketing subs this period),
            # add these CC-only rows under the same header; otherwise open a new one.
            lines.append(f'<tr class="div-row"><td colspan="{colspan}">{grp_name.upper()} (no raw leads this period)</td></tr>')

            grp_set = grp_cc_gi = grp_ni = grp_demo = grp_sale = grp_drop = grp_gross = 0
            for r in rows:
                set_v=r['set']; gi=r['gross_issue']; ni=r.get('net_issue',gi)
                demo=r['demo']; sale=r['sale']; drop=r['drop']; gross=r.get('gross_amt',0)
                grp_set+=set_v; grp_cc_gi+=gi; grp_ni+=ni
                grp_demo+=demo; grp_sale+=sale; grp_drop+=drop; grp_gross+=gross
                i_pct=pct(gi,set_v) if set_v else '—'
                d_pct=pct(demo,gi) if gi else '—'
                c_pct=pct(sale,demo) if demo else '—'
                if colspan==12:
                    lines.append(
                        f'<tr><td>&nbsp;&nbsp;{r["name"]}</td>'
                        f'<td class="mc">—</td>'
                        f'{td(set_v)}{td(gi)}'
                        f'<td{rate_class(i_pct)}>{i_pct}</td>'
                        f'{td(ni)}{td(demo)}'
                        f'<td{rate_class(d_pct)}>{d_pct}</td>'
                        f'{td(sale)}<td{rate_class(c_pct)}>{c_pct}</td>'
                        f'{td(drop)}<td>{dollar(gross)}</td></tr>')
                else:
                    lines.append(
                        f'<tr><td>&nbsp;&nbsp;{r["name"]}</td>'
                        f'<td class="mc">—</td>'
                        f'{td(set_v)}{td(gi)}'
                        f'<td{rate_class(i_pct)}>{i_pct}</td>'
                        f'{td(demo)}'
                        f'<td{rate_class(d_pct)}>{d_pct}</td>'
                        f'{td(sale)}<td{rate_class(c_pct)}>{c_pct}</td>'
                        f'{td(drop)}<td>{dollar(gross)}</td></tr>')
            # Group subtotal
            gi_p=pct(grp_cc_gi,grp_set) if grp_set else '—'
            dm_p=pct(grp_demo,grp_ni) if grp_ni else '—'
            cl_p=pct(grp_sale,grp_demo) if grp_demo else '—'
            if colspan==12:
                lines.append(
                    f'<tr class="grp"><td>&nbsp;&nbsp;{grp_name} Total</td>'
                    f'<td class="mc">—</td>'
                    f'{td(grp_set)}{td(grp_cc_gi)}<td>{gi_p}</td>'
                    f'{td(grp_ni)}{td(grp_demo)}<td>{dm_p}</td>'
                    f'{td(grp_sale)}<td>{cl_p}</td>{td(grp_drop)}<td>{dollar(grp_gross)}</td></tr>')
            else:
                lines.append(
                    f'<tr class="grp"><td>&nbsp;&nbsp;{grp_name} Total</td>'
                    f'<td class="mc">—</td>'
                    f'{td(grp_set)}{td(grp_cc_gi)}<td>{gi_p}</td>'
                    f'{td(grp_demo)}<td>{dm_p}</td>'
                    f'{td(grp_sale)}<td>{cl_p}</td>{td(grp_drop)}<td>{dollar(grp_gross)}</td></tr>')

    # Grand total row
    if ss_total:
        t=ss_total
        total_raw = mkt_grand['raw'] if mkt_grand else '—'
        gi_p = pct(t['gross_issue'], t['set']) if t['set'] else '—'
        dm_p = pct(t['demo'], t['gross_issue']) if t['gross_issue'] else '—'
        cl_p = pct(t['sale'], t['demo']) if t['demo'] else '—'
        if colspan==12:
            lines.append(
                f'<tr class="gtot"><td>GRAND TOTAL</td>'
                f'<td>{total_raw}</td>'
                f'{td(t["set"])}{td(t["gross_issue"])}'
                f'<td>{gi_p}</td>{td(t["net_issue"])}{td(t["demo"])}'
                f'<td>{dm_p}</td>{td(t["sale"])}<td>{cl_p}</td>'
                f'{td(t["drop"])}<td>{dollar(t.get("gross_amt",0))}</td></tr>')
        else:
            lines.append(
                f'<tr class="gtot"><td>GRAND TOTAL</td>'
                f'<td>{total_raw}</td>'
                f'{td(t["set"])}{td(t["gross_issue"])}'
                f'<td>{gi_p}</td>{td(t["demo"])}'
                f'<td>{dm_p}</td>{td(t["sale"])}<td>{cl_p}</td>'
                f'{td(t["drop"])}<td>{dollar(t.get("gross_amt",0))}</td></tr>')

    return '\n'.join(lines)


def build_product_table(prod_rows, prod_total, period_label):
    """Build the product performance table."""
    BATH_CODES = {'FB', 'WA', 'WAIT', 'WAIT'}
    KIT_CODES  = {'FK', 'CAB'}
    WIN_CODES  = {'WINDR'}

    grp = {'kitchens': [], 'bathrooms': [], 'windows': [], 'other': []}
    for r in prod_rows:
        code = r['name'].upper()
        if code in KIT_CODES: grp['kitchens'].append(r)
        elif code in BATH_CODES: grp['bathrooms'].append(r)
        elif code in WIN_CODES: grp['windows'].append(r)
        else: grp['other'].append(r)

    def sum_grp(rows):
        return {k: sum(r.get(k,0) for r in rows) for k in ('set','sale','demo','net_issue','gross_issue','drop','gross_amt')}

    def prod_row(r):
        code = r['name'].upper()
        lbl  = PRODUCT_LABELS.get(code, code)
        gi   = r['gross_issue']; ni = r['net_issue']
        demo = r['demo']; sale = r['sale']
        gross = r.get('gross_amt', 0)
        ni_p = pct(ni, gi) if gi else '—'
        dm_p = pct(demo, gi) if gi else '—'
        cl_p = pct(sale, demo) if demo else '—'
        return (f'<tr><td>&nbsp;&nbsp;{code}</td><td>{lbl}</td>'
                f'<td>{gi}</td><td>{ni}</td>'
                f'<td{rate_class(ni_p)}>{ni_p}</td>'
                f'<td>{demo}</td>'
                f'<td{rate_class(dm_p)}>{dm_p}</td>'
                f'<td>{sale}</td>'
                f'<td{rate_class(cl_p) if sale>0 else ""}>{cl_p}</td>'
                f'<td>{dollar(gross)}</td></tr>')

    def grp_total_row(name, rows, css):
        s = sum_grp(rows)
        gi=s['gross_issue']; ni=s['net_issue']
        demo=s['demo']; sale=s['sale']
        gross=s.get('gross_amt',0)
        ni_p=pct(ni,gi) if gi else '—'
        dm_p=pct(demo,gi) if gi else '—'
        cl_p=pct(sale,demo) if demo else '—'
        return (f'<tr class="grp-{css}"><td>&nbsp;&nbsp;{name}</td><td></td>'
                f'<td>{gi}</td><td>{ni}</td>'
                f'<td{rate_class(ni_p)}>{ni_p}</td>'
                f'<td>{demo}</td>'
                f'<td{rate_class(dm_p)}>{dm_p}</td>'
                f'<td>{sale}</td><td>{cl_p}</td><td>{dollar(gross)}</td></tr>')

    lines = [
        '<tr class="div-kitchens"><td colspan="10">KITCHENS — CAB (Cabinets) + FK (Full Kitchen)</td></tr>',
    ]
    if grp['kitchens']:
        for r in grp['kitchens']: lines.append(prod_row(r))
        lines.append(grp_total_row('Kitchens Total', grp['kitchens'], 'kitchens'))
    else:
        lines.append('<tr><td colspan="10" style="text-align:center;color:#aaa;font-style:italic;padding:14px;">No appointments this period</td></tr>')

    lines.append('<tr class="div-bathrooms"><td colspan="10">BATHROOMS — FB (Full Bath) + WA (Wet Area) + WAIT (Walk-in Tub)</td></tr>')
    if grp['bathrooms']:
        for r in grp['bathrooms']: lines.append(prod_row(r))
        lines.append(grp_total_row('Bathrooms Total', grp['bathrooms'], 'bathrooms'))
    else:
        lines.append('<tr><td colspan="10" style="text-align:center;color:#aaa;font-style:italic;padding:14px;">No appointments this period</td></tr>')

    lines.append('<tr class="div-windows"><td colspan="10">WINDOWS &amp; DOORS — WINDR</td></tr>')
    if grp['windows']:
        for r in grp['windows']: lines.append(prod_row(r))
        lines.append(grp_total_row('W&amp;D Total', grp['windows'], 'windows'))
    else:
        lines.append('<tr><td colspan="10" style="text-align:center;color:#aaa;font-style:italic;padding:14px;">No appointments this period</td></tr>')

    if prod_total:
        t=prod_total
        gi=t['gross_issue']; ni=t['net_issue']; demo=t['demo']; sale=t['sale']
        gross=t.get('gross_amt',0)
        ni_p=pct(ni,gi) if gi else '—'
        dm_p=pct(demo,gi) if gi else '—'
        cl_p=pct(sale,demo) if demo else '—'
        lines.append(
            f'<tr class="gtot"><td>GRAND TOTAL</td><td></td>'
            f'<td>{gi}</td><td>{ni}</td><td>{ni_p}</td>'
            f'<td>{demo}</td><td>{dm_p}</td>'
            f'<td>{sale}</td><td>{cl_p}</td><td>{dollar(gross)}</td></tr>')

    return '\n'.join(lines)


def build_setter_table(setter_rows, setter_total):
    """Full Appointment Statistics by Setter table."""
    def nsli(gross_amt, gi):
        return dollar(gross_amt / gi) if gi else '—'

    lines = []
    for r in setter_rows:
        lines.append(
            f'<tr>'
            f'<td>{r["name"]}</td>'
            f'<td>{r["set"]}</td>'
            f'<td>{r["gross_issue"]}</td>'
            f'<td>{r["demo"]}</td>'
            f'<td>{r["sale"]}</td>'
            f'<td>{r["cns"]}</td>'
            f'<td>{r["one_leg"]}</td>'
            f'<td>{r["cxl"]}</td>'
            f'<td>{r["ng"]}</td>'
            f'<td>{r["drop"]}</td>'
            f'<td>{dollar(r["gross_amt"])}</td>'
            f'<td>{dollar(r["net_amt"])}</td>'
            f'<td>{nsli(r["gross_amt"], r["gross_issue"])}</td>'
            f'</tr>')
    if setter_total:
        t = setter_total
        lines.append(
            f'<tr class="gtot">'
            f'<td>TOTAL</td>'
            f'<td>{t["set"]}</td>'
            f'<td>{t["gross_issue"]}</td>'
            f'<td>{t["demo"]}</td>'
            f'<td>{t["sale"]}</td>'
            f'<td>{t["cns"]}</td>'
            f'<td>{t["one_leg"]}</td>'
            f'<td>{t["cxl"]}</td>'
            f'<td>{t["ng"]}</td>'
            f'<td>{t["drop"]}</td>'
            f'<td>{dollar(t["gross_amt"])}</td>'
            f'<td>{dollar(t["net_amt"])}</td>'
            f'<td>{nsli(t["gross_amt"], t["gross_issue"])}</td>'
            f'</tr>')
    return '\n'.join(lines)


# ── Chart data helpers ────────────────────────────────────────────────────────

def js_arr(lst):
    return '[' + ','.join(str(x) for x in lst) + ']'

def js_str_arr(lst):
    return '[' + ','.join(f"'{x}'" for x in lst) + ']'

def chart_data_from_source(src_rows):
    """Labels and set counts from appt_by_source rows for pie chart."""
    labels = [r['name'] for r in src_rows if r['set'] > 0]
    vals   = [r['set'] for r in src_rows if r['set'] > 0]
    return labels, vals

def funnel_data(mkt_grand, ss_total):
    """[raw, set, gross_issue, net_issue, demo, sale]"""
    raw = mkt_grand['raw'] if mkt_grand else 0
    t   = ss_total or {}
    return [raw, t.get('set',0), t.get('gross_issue',0),
            t.get('net_issue',0), t.get('demo',0), t.get('sale',0)]


# ── HTML template ─────────────────────────────────────────────────────────────

CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; color: #222; }
header { background: linear-gradient(135deg, #1a2e4a 0%, #2d5a8e 100%); color: white; padding: 20px 32px; display: flex; align-items: center; gap: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.3); }
header h1 { font-size: 1.5rem; font-weight: 700; }
header p { font-size: 0.83rem; opacity: 0.75; margin-top: 2px; }
.logo { font-size: 2rem; }
.data-note { background: #e8f4fd; border-left: 4px solid #4da6ff; padding: 10px 32px; font-size: 0.79rem; color: #444; }
.data-note strong { color: #1a2e4a; }
.tabs { display: flex; background: #1a2e4a; padding: 0 32px; gap: 4px; flex-wrap: wrap; }
.tab { padding: 11px 20px; cursor: pointer; color: rgba(255,255,255,0.6); font-size: 0.85rem; font-weight: 600; border-bottom: 3px solid transparent; transition: all 0.2s; }
.tab:hover { color: white; }
.tab.active { color: white; border-bottom-color: #4da6ff; }
.section { display: none; padding: 24px 32px; }
.section.active { display: block; }
.period-label { font-size: 0.78rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #4da6ff; margin-bottom: 16px; }
.kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
.kpi-card { background: white; border-radius: 10px; padding: 16px 18px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-left: 4px solid #4da6ff; }
.kpi-card.green { border-left-color: #28a745; } .kpi-card.orange { border-left-color: #fd7e14; } .kpi-card.purple { border-left-color: #6f42c1; } .kpi-card.red { border-left-color: #dc3545; } .kpi-card.teal { border-left-color: #17a2b8; }
.kpi-source { font-size: 0.65rem; color: #bbb; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }
.kpi-label { font-size: 0.73rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.8px; color: #888; }
.kpi-value { font-size: 1.9rem; font-weight: 800; margin: 4px 0 2px; color: #1a2e4a; line-height: 1; }
.kpi-sub { font-size: 0.77rem; color: #999; }
.charts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 24px; }
.chart-card { background: white; border-radius: 10px; padding: 18px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
.chart-card.full { grid-column: 1 / -1; }
.chart-title { font-size: 0.82rem; font-weight: 700; color: #444; margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
.chart-subtitle { display: block; font-size: 0.69rem; color: #999; font-weight: normal; text-transform: none; letter-spacing: 0; margin-top: 2px; }
.chart-wrap { position: relative; height: 260px; } .chart-wrap.tall { height: 340px; }
.table-card { background: white; border-radius: 10px; padding: 18px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin-bottom: 20px; overflow-x: auto; }
.table-note { font-size: 0.72rem; color: #888; margin-bottom: 10px; font-style: italic; }
table { width: 100%; border-collapse: collapse; font-size: 0.78rem; }
thead th { background: #1a2e4a; color: white; padding: 8px 9px; text-align: right; font-weight: 600; font-size: 0.7rem; white-space: nowrap; }
thead th:first-child { text-align: left; }
thead th.src-col { background: #4a2d7a; }
tbody tr { border-bottom: 1px solid #f0f0f0; } tbody tr:hover { background: #f7f9fc; }
tbody tr.grp { background: #eef3f9; font-weight: 700; }
tbody tr.gtot { background: #1a2e4a; color: white; font-weight: 800; }
.ni-cell { background: #e8f4fd; color: #1a3a5c; }
tbody tr.div-row td { background: #f0f4fa; font-size: 0.67rem; text-transform: uppercase; letter-spacing: 0.8px; color: #666; font-weight: 700; padding: 5px 9px; }
tbody tr.div-kitchens td { background: #fef9ec; color: #92400e; }
tbody tr.div-bathrooms td { background: #eff6ff; color: #1e40af; }
tbody tr.div-windows td { background: #ecfdf5; color: #065f46; }
tbody tr.grp-kitchens { background: #fef3c7; font-weight: 700; }
tbody tr.grp-bathrooms { background: #dbeafe; font-weight: 700; }
tbody tr.grp-windows { background: #d1fae5; font-weight: 700; }
tbody td { padding: 7px 9px; text-align: right; white-space: nowrap; } tbody td:first-child { text-align: left; }
.g { color: #28a745; font-weight: 600; } .w { color: #fd7e14; font-weight: 600; } .b { color: #dc3545; font-weight: 600; }
.mc { color: #6f42c1; font-style: italic; }
.period-toggle { display: flex; gap: 8px; margin-bottom: 20px; }
.ptab { padding: 7px 18px; border-radius: 20px; cursor: pointer; font-size: 0.81rem; font-weight: 600; background: #e9ecef; color: #555; border: none; transition: all 0.2s; }
.ptab.active { background: #1a2e4a; color: white; }
.ptab:hover:not(.active) { background: #d1d5db; }
.period-section { display: none; }
.period-section.active { display: block; }
.group-summary { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 20px; }
.group-card { background: white; border-radius: 10px; padding: 16px 18px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
.group-card.kitchens { border-top: 4px solid #fd7e14; }
.group-card.bathrooms { border-top: 4px solid #4da6ff; }
.group-card.windows { border-top: 4px solid #20c997; }
.group-name { font-size: 0.84rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 10px; }
.group-name.kitchens { color: #92400e; }
.group-name.bathrooms { color: #1e40af; }
.group-name.windows { color: #065f46; }
.group-stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; }
.gstat-label { font-size: 0.66rem; color: #aaa; text-transform: uppercase; }
.gstat-val { font-size: 1.15rem; font-weight: 800; color: #1a2e4a; }
.rep-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 20px; }
.rep-card { background: white; border-radius: 10px; padding: 14px 16px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-left: 4px solid #4da6ff; }
.rep-name { font-size: 0.88rem; font-weight: 700; color: #1a2e4a; margin-bottom: 8px; }
.rep-stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; }
.rstat-label { font-size: 0.64rem; color: #aaa; text-transform: uppercase; }
.rstat-val { font-size: 1rem; font-weight: 700; color: #1a2e4a; }
.rep-gross { margin-top: 8px; font-size: 0.75rem; color: #555; }
"""

JS_HELPERS = """
function showTab(id,el){document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));document.getElementById(id).classList.add('active');el.classList.add('active');window.dispatchEvent(new Event('resize'));}
function showPeriod(sectId,id,el){document.querySelectorAll('#'+sectId+' .period-section').forEach(s=>s.classList.remove('active'));document.querySelectorAll('#'+sectId+' .ptab').forEach(t=>t.classList.remove('active'));document.getElementById(id).classList.add('active');el.classList.add('active');window.dispatchEvent(new Event('resize'));}
const C=['#2d5a8e','#4da6ff','#28a745','#fd7e14','#6f42c1','#dc3545','#17a2b8','#ffc107','#20c997','#e83e8c'];
function pie(id,labels,data){var el=document.getElementById(id);if(!el)return;new Chart(el,{type:'doughnut',data:{labels,datasets:[{data,backgroundColor:C,borderWidth:2,borderColor:'#fff'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{font:{size:11},padding:10}}}}})}
function hbar(id,labels,data,colors){var el=document.getElementById(id);if(!el)return;new Chart(el,{type:'bar',data:{labels,datasets:[{data,backgroundColor:colors,borderRadius:4}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{beginAtZero:true}}}})}
function bar(id,labels,datasets){var el=document.getElementById(id);if(!el)return;new Chart(el,{type:'bar',data:{labels,datasets},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top'}},scales:{x:{ticks:{font:{size:10}}},y:{beginAtZero:true}}}})}
"""


def gen_period_section(pid, label, date_label, d, colspan, ytd_source_map=None):
    """Generate one main period section (PW/PM/MTD/YTD)."""
    mkt_groups, mkt_grand = d['marketing']
    src_rows, src_total   = d['source']
    ss_rows,  ss_total    = d['subsource']

    t   = ss_total or {}
    raw = mkt_grand['raw'] if mkt_grand else 0

    # KPI calculations
    set_v   = t.get('set', 0)
    gi      = t.get('gross_issue', 0)
    ni      = t.get('net_issue', 0)
    demo    = t.get('demo', 0)
    sale    = t.get('sale', 0)
    drop    = t.get('drop', 0)

    gi_rate  = pct(gi, set_v) if set_v else '—'
    dm_rate  = pct(demo, gi) if gi else '—'
    cl_rate  = pct(sale, demo) if demo else '—'
    dr_rate  = pct(drop, set_v) if set_v else '—'

    # Setter total for gross revenue
    _, setter_total = d['setter']
    gross = setter_total['gross_amt'] if setter_total else 0

    # Charts
    src_labels, src_sets = chart_data_from_source(src_rows)
    funnel = funnel_data(mkt_grand, ss_total)

    # Raw leads by source for pie
    raw_by_src = {}
    for grp_name, grp in mkt_groups.items():
        raw_by_src[grp_name] = grp['raw']
    raw_labels = [k for k,v in raw_by_src.items() if v > 0]
    raw_vals   = [v for v in raw_by_src.values() if v > 0]

    has_multi_src = len(src_labels) > 1

    # Sub-source table
    table_body = build_subsource_table(mkt_groups, mkt_grand, ss_rows, ss_total, colspan, ytd_source_map)

    # Table header
    if colspan == 12:
        th = ('<th>Sub-Source</th><th class="src-col">Raw*</th>'
              '<th>Set</th><th>Gross Issue</th><th>Issue Rate</th>'
              '<th>Net Issue</th><th>Demo</th><th>Demo Rate</th>'
              '<th>Sales</th><th>Close %</th><th>Drop</th><th>Gross $</th>')
    else:
        th = ('<th>Sub-Source</th><th class="src-col">Raw*</th>'
              '<th>Set</th><th>Gross Issue</th><th>Issue Rate</th>'
              '<th>Demo</th><th>Demo Rate</th>'
              '<th>Sales</th><th>Close %</th><th>Drop</th><th>Gross $</th>')

    chart_ids_js = ''
    charts_html  = ''
    if has_multi_src:
        charts_html = f'''
  <div class="charts-grid">
    <div class="chart-card"><div class="chart-title">Raw Leads by Source<span class="chart-subtitle">Marketing Sub-Source Report</span></div><div class="chart-wrap"><canvas id="{pid}-leads-pie"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">Appointments Set by Source<span class="chart-subtitle">Call Center Report</span></div><div class="chart-wrap"><canvas id="{pid}-set-pie"></canvas></div></div>
    <div class="chart-card full"><div class="chart-title">Conversion Funnel<span class="chart-subtitle">Call Center Report</span></div><div class="chart-wrap"><canvas id="{pid}-funnel"></canvas></div></div>
  </div>'''
        chart_ids_js = (
            f"pie('{pid}-leads-pie',{js_str_arr(raw_labels)},{js_arr(raw_vals)});\n"
            f"pie('{pid}-set-pie',{js_str_arr(src_labels)},{js_arr(src_sets)});\n"
            f"hbar('{pid}-funnel',['Raw Leads (Mktg)','Set (CC)','Gross Issue (CC)','Net Issue (CC)','Demos (CC)','Sales (CC)'],"
            f"{js_arr(funnel)},['#6f42c1','#2d5a8e','#4da6ff','#17a2b8','#28a745','#fd7e14']);\n"
        )
    else:
        charts_html = f'''
  <div class="charts-grid">
    <div class="chart-card full"><div class="chart-title">Conversion Funnel<span class="chart-subtitle">Call Center Report</span></div><div class="chart-wrap"><canvas id="{pid}-funnel"></canvas></div></div>
  </div>'''
        chart_ids_js = (
            f"hbar('{pid}-funnel',['Raw Leads (Mktg)','Set (CC)','Gross Issue (CC)','Net Issue (CC)','Demos (CC)','Sales (CC)'],"
            f"{js_arr(funnel)},['#6f42c1','#2d5a8e','#4da6ff','#17a2b8','#28a745','#fd7e14']);\n"
        )

    html = f'''
<!-- ===== {label.upper()} ===== -->
<div class="section{"" if pid != "pw" else " active"}" id="{pid}">
  <div class="period-label">{label} &nbsp;•&nbsp; {date_label}</div>
  <div class="kpi-grid">
    {kpi('Total Raw Leads', raw, f'Leads received this period', 'Marketing Report')}
    {kpi('Appointments Set', set_v, f'Appt date {date_label}', 'Call Center Report', 'green')}
    {kpi('Demos Run', demo, f'{dm_rate} of gross issued', 'Call Center Report', 'orange')}
    {kpi('Sales', sale, f'{dollar(gross)} gross revenue', 'Call Center Report', 'purple')}
    {kpi('Gross Issue Rate', gi_rate, f'{gi} issued / {set_v} set', 'Call Center Report', 'teal')}
    {kpi('Demo Rate', dm_rate, f'{demo} demos / {gi} gross issued', 'Call Center Report', 'green')}
    {kpi('Gross Close %', cl_rate, f'{sale} sales / {demo} demos', 'Call Center Report', 'red' if sale==0 else 'green')}
    {kpi('Drop Rate', dr_rate, f'{drop} dropped / {set_v} set', 'Call Center Report')}
  </div>
{charts_html}
  <div class="table-card">
    <div class="chart-title">{label} Results by Sub-Source</div>
    <div class="table-note">Raw = Marketing Sub-Source Report. Set / Issue / Demo / Sales = Call Center Appointment Statistics ({date_label}).</div>
    <table><thead><tr>{th}</tr></thead>
    <tbody>
{table_body}
    </tbody></table>
  </div>
</div>
'''
    return html, chart_ids_js


def gen_products_section(data_by_period, date_ranges_dict):
    """Generate the Products tab."""
    PMAP = {'prior_week':'pw','prior_month':'pm','mtd':'mtd','ytd':'ytd'}
    PLABELS = {'prior_week':'Prior Week','prior_month':'Prior Month','mtd':'Month to Date','ytd':'Year to Date'}
    sections = []
    default = 'ytd'
    for period in PERIODS:
        pid = PMAP[period]
        label = PLABELS[period]
        _, _, date_lbl = date_ranges_dict[period]
        prod_rows, prod_total = data_by_period[period]['product']
        table = build_product_table(prod_rows, prod_total, label)
        active = ' active' if period == default else ''
        sections.append(f'''
  <div class="period-section{active}" id="prod-{pid}">
    <div class="period-label">{label} &nbsp;•&nbsp; {date_lbl}</div>
    <div class="table-card">
      <div class="chart-title">Product Performance — {label}</div>
      <div class="table-note">Appointment Statistics by Product. Appointment dates {date_lbl}.</div>
      <table><thead><tr><th>Product</th><th>Description</th><th>Issued</th><th>Net Iss</th><th>Net Iss %</th><th>Demos</th><th>Demo %</th><th>Sales</th><th>Close %</th><th>Gross $</th></tr></thead>
      <tbody>{table}</tbody></table>
    </div>
  </div>''')

    btns = ' '.join(
        f'<button class="ptab{"" if p!="ytd" else " active"}" onclick="showPeriod(\'products\',\'prod-{PMAP[p]}\',this)">{PLABELS[p]}</button>'
        for p in PERIODS)
    return f'''
<!-- ===== PRODUCTS ===== -->
<div class="section" id="products">
  <div class="period-label">Product Performance &nbsp;•&nbsp; Appointment Statistics by Product</div>
  <div class="period-toggle">{btns}</div>
{''.join(sections)}
</div>
'''


def gen_salesreps_section(data_by_period, date_ranges_dict):
    """Generate the Setters tab with infographic, NSLI, and full report table."""
    PMAP   = {'prior_week':'pw','prior_month':'pm','mtd':'mtd','ytd':'ytd'}
    PLABELS= {'prior_week':'Prior Week','prior_month':'Prior Month','mtd':'Month to Date','ytd':'Year to Date'}
    sections = []
    chart_js = ''
    default = 'ytd'

    for period in PERIODS:
        pid   = PMAP[period]
        label = PLABELS[period]
        _, _, date_lbl = date_ranges_dict[period]
        setter_rows, setter_total = data_by_period[period]['setter']
        active_rows = [r for r in setter_rows if r['set'] > 0]
        dispo_rows, dispo_total, dispo_cols = data_by_period[period].get('dispo', ([], None, {}))

        # ── KPI summary ──────────────────────────────────────────────────────
        t = setter_total or {}
        gross = t.get('gross_amt', 0)
        net   = t.get('net_amt', 0)
        total_set   = t.get('set', 0)
        total_demo  = t.get('demo', 0)
        total_sale  = t.get('sale', 0)
        total_gi    = t.get('gross_issue', 0)
        # NSLI = Gross Revenue / Gross Issued Leads ($ per issued lead)
        nsli_val    = dollar(gross / total_gi) if total_gi else '—'
        nsli_color  = 'teal'
        # Demo % = Sit (Demo) / Gross Issue — matches LP report calculation
        overall_demo_pct = f'{total_demo/total_gi*100:.1f}%' if total_gi else '—'
        overall_demo_meets_target = total_gi > 0 and total_demo / total_gi >= 0.70

        kpis = (
            kpi('Appointments Set', total_set, f'{len(active_rows)} active setters', 'Appt Stats by Setter', 'green') +
            kpi('Demos Run', total_demo, f'{total_sale} sales', 'Appt Stats by Setter', 'orange') +
            kpi('Demo %', overall_demo_pct, 'Target: 70%', 'Appt Stats by Setter', 'green' if overall_demo_meets_target else 'red') +
            kpi('NSLI', nsli_val, 'Gross $ ÷ Issued Leads', 'Appt Stats by Setter', nsli_color) +
            kpi('Gross Revenue', dollar(gross), f'Net: {dollar(net)}', 'Appt Stats by Setter', 'purple') +
            kpi('CNS', t.get('cns',0), 'Customer No-Show', 'Appt Stats by Setter') +
            kpi('1-Leg Resets', t.get('one_leg',0), 'One leg reset', 'Appt Stats by Setter') +
            kpi('CXL', t.get('cxl',0), 'Cancel prior to issue', 'Appt Stats by Setter') +
            kpi('No Good (NG)', t.get('ng',0), 'No good appointments', 'Appt Stats by Setter')
        )

        # ── Chart IDs ────────────────────────────────────────────────────────
        pie_id        = f'setter-results-pie-{pid}'
        demo_chart_id = f'setter-demo-pct-{pid}'
        s_labels      = js_str_arr([r['name'].split(',')[0] for r in active_rows])
        chart_height  = max(220, len(active_rows) * 42)

        # Pie chart — overall appointment result breakdown from total row
        t_sale    = t.get('sale', 0)
        t_demo    = t.get('demo', 0)
        t_cns     = t.get('cns', 0)
        t_one_leg = t.get('one_leg', 0)
        t_cxl     = t.get('cxl', 0)
        t_ng      = t.get('ng', 0)
        t_drop    = t.get('drop', 0)
        t_demo_ns = max(0, t_demo - t_sale)   # demos run that didn't close
        pie_labels = js_str_arr(['Sale','Demo (No Sale)','CNS','1-Leg Reset','CXL','NG','Drop'])
        pie_vals   = js_arr([t_sale, t_demo_ns, t_cns, t_one_leg, t_cxl, t_ng, t_drop])
        pie_colors = js_str_arr(['#28a745','#4da6ff','#fd7e14','#6f42c1','#dc3545','#adb5bd','#17a2b8'])

        # Demo % = Sit (Demo) / Gross Issue — matches LP report
        demo_pcts   = [round(r['demo']/r['gross_issue']*100, 1) if r.get('gross_issue') else 0 for r in active_rows]
        demo_colors = js_str_arr(['#28a745' if p >= 70 else '#dc3545' for p in demo_pcts])
        demo_vals   = js_arr(demo_pcts)

        chart_js += f"""
(function(){{
  var el=document.getElementById('{pie_id}');
  if(!el)return;
  new Chart(el,{{
    type:'doughnut',
    data:{{
      labels:{pie_labels},
      datasets:[{{data:{pie_vals},backgroundColor:{pie_colors},borderWidth:2,borderColor:'#fff'}}]
    }},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'right',labels:{{font:{{size:11}},padding:12}}}}}}
    }}
  }});
}})();
(function(){{
  var el=document.getElementById('{demo_chart_id}');
  if(!el)return;
  new Chart(el,{{
    type:'bar',
    data:{{
      labels:{s_labels},
      datasets:[{{
        label:'Demo %',
        data:{demo_vals},
        backgroundColor:{demo_colors},
        borderRadius:4
      }}]
    }},
    options:{{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{{
        legend:{{display:false}},
        annotation:{{
          annotations:{{
            target:{{
              type:'line',drawTime:'afterDatasetsDraw',
              scaleID:'x',value:70,
              borderColor:'#1a2e4a',borderWidth:2,borderDash:[6,3],
              label:{{display:true,content:'Target 70%',position:'start',
                     backgroundColor:'#1a2e4a',color:'white',
                     font:{{size:11,weight:'bold'}},padding:4}}
            }}
          }}
        }}
      }},
      scales:{{
        x:{{beginAtZero:true,max:100,ticks:{{callback:function(v){{return v+'%'}}}}}},
        y:{{ticks:{{font:{{size:11}}}}}}
      }}
    }}
  }});
}})();
"""

        # ── Per-setter summary cards ─────────────────────────────────────────
        rep_cards = ''
        for r in active_rows:
            nsli_r     = dollar(r['gross_amt'] / r['gross_issue']) if r.get('gross_issue') else '—'
            demo_pct_r = f'{r["demo"]/r["gross_issue"]*100:.1f}%' if r.get('gross_issue') else '—'
            rep_cards += f'''
    <div class="rep-card">
      <div class="rep-name">{r["name"]}</div>
      <div class="rep-stats">
        <div><div class="rstat-label">Set</div><div class="rstat-val">{r["set"]}</div></div>
        <div><div class="rstat-label">Demo</div><div class="rstat-val">{r["demo"]}</div></div>
        <div><div class="rstat-label">Demo %</div><div class="rstat-val">{demo_pct_r}</div></div>
        <div><div class="rstat-label">Sale</div><div class="rstat-val">{r["sale"]}</div></div>
        <div><div class="rstat-label">CNS</div><div class="rstat-val">{r["cns"]}</div></div>
        <div><div class="rstat-label">1-Leg</div><div class="rstat-val">{r["one_leg"]}</div></div>
      </div>
      <div class="rep-gross">NSLI: {nsli_r} &nbsp;|&nbsp; CXL: {r["cxl"]} &nbsp;|&nbsp; Drop: {r["drop"]}</div>
    </div>'''

        # ── Full report table ────────────────────────────────────────────────
        table_body = build_setter_table(setter_rows, setter_total)

        # ── Dispo Distribution pie chart + table ─────────────────────────────
        dispo_pie_id = f'dispo-pie-{pid}'
        dispo_html = ''
        if dispo_total and dispo_total.get('dispos'):
            dt = dispo_total['dispos']
            grand_total = dispo_total['total'] or 1
            # Sort by count descending
            sorted_dispos = sorted(dt.items(), key=lambda x: x[1], reverse=True)
            # Pie: labels include count + pct
            def dpct(n): return round(n / grand_total * 100)
            dp_labels = js_str_arr([f'{d[0]}: {d[1]} ({dpct(d[1])}%)' for d in sorted_dispos])
            dp_vals   = js_arr([d[1] for d in sorted_dispos])
            DPAL = ['#2d5a8e','#28a745','#fd7e14','#dc3545','#6f42c1','#17a2b8',
                    '#ffc107','#4da6ff','#20c997','#e83e8c','#adb5bd','#343a40',
                    '#6610f2','#795548','#00bcd4','#ff5722','#9c27b0','#607d8b']
            dp_colors = js_str_arr([DPAL[i % len(DPAL)] for i in range(len(sorted_dispos))])

            chart_js += f"""
(function(){{
  var el=document.getElementById('{dispo_pie_id}');
  if(!el)return;
  var total={grand_total};
  new Chart(el,{{
    type:'doughnut',
    data:{{
      labels:{dp_labels},
      datasets:[{{data:{dp_vals},backgroundColor:{dp_colors},borderWidth:2,borderColor:'#fff'}}]
    }},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{
        legend:{{position:'right',labels:{{font:{{size:11}},padding:10}}}},
        tooltip:{{callbacks:{{label:function(ctx){{
          var n=ctx.parsed;
          var pct=Math.round(n/total*100);
          return ' '+ctx.label.split(':')[0]+': '+n+' ('+pct+'%)';
        }}}}}}
      }}
    }}
  }});
}})();
"""
            # Dispo table — each cell shows "N (pct%)", total row shows count + pct
            dispo_cols_sorted = [d[0] for d in sorted_dispos]
            th_dispos = ''.join(f'<th>{c}</th>' for c in dispo_cols_sorted)
            dispo_tbl_rows = ''
            for drow in dispo_rows:
                row_total = drow['total'] or 1
                def cell(c, row_tot=row_total):
                    n = drow['dispos'].get(c, 0)
                    if not n: return '—'
                    return f'{n} ({dpct(n)}%)'
                tds = ''.join(f'<td>{cell(c)}</td>' for c in dispo_cols_sorted)
                dispo_tbl_rows += f'<tr><td>{drow["name"]}</td><td>{drow["total"]}</td>{tds}</tr>'
            if dispo_total:
                def tcell(c):
                    n = dispo_total['dispos'].get(c, 0)
                    if not n: return '—'
                    return f'{n} ({dpct(n)}%)'
                tot_tds = ''.join(f'<td><strong>{tcell(c)}</strong></td>' for c in dispo_cols_sorted)
                dispo_tbl_rows += f'<tr class="gtot"><td>TOTAL</td><td>{dispo_total["total"]}</td>{tot_tds}</tr>'

            dispo_html = f'''
    <div class="charts-grid">
      <div class="chart-card full">
        <div class="chart-title">Dispo Distribution — {label}
          <span class="chart-subtitle">Source: Call Center Dispo Distribution report. All appointments including issued and not issued.</span>
        </div>
        <div class="chart-wrap" style="height:300px"><canvas id="{dispo_pie_id}"></canvas></div>
      </div>
    </div>
    <div class="table-card">
      <div class="chart-title">Dispo Distribution by Setter — {label}</div>
      <div class="table-note">Source: Call Center Dispo Distribution. Appointment dates {date_lbl}. Columns ordered most to least frequent (YTD).</div>
      <table>
        <thead><tr><th>Setter</th><th>Total</th>{th_dispos}</tr></thead>
        <tbody>{dispo_tbl_rows}</tbody>
      </table>
    </div>'''

        active = ' active' if period == default else ''
        sections.append(f'''
  <div class="period-section{active}" id="rep-{pid}">
    <div class="period-label">{label} &nbsp;•&nbsp; {date_lbl}</div>
    <div class="kpi-grid">{kpis}</div>

    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-title">Appointment Results Breakdown
          <span class="chart-subtitle">All setters combined — {label}</span>
        </div>
        <div class="chart-wrap"><canvas id="{pie_id}"></canvas></div>
      </div>
      <div class="chart-card full">
        <div class="chart-title">Demo % vs. 70% Target — by Setter
          <span class="chart-subtitle">Green = at or above target &nbsp;·&nbsp; Red = below target</span>
        </div>
        <div class="chart-wrap" style="height:{chart_height}px"><canvas id="{demo_chart_id}"></canvas></div>
      </div>
    </div>

    <div class="rep-grid">{rep_cards}</div>

    <div class="table-card">
      <div class="chart-title">Appointment Statistics by Setter — {label}</div>
      <div class="table-note">
        Source: Appointment Statistics by Setter. Appointment dates {date_lbl}.<br>
        NSLI = Gross Revenue ÷ Gross Issued Leads ($ per issued lead). CNS=Customer No-Show · 1Leg=One Leg Reset · CXL=Cancel prior to Issue · NG=No Good.
      </div>
      <table>
        <thead><tr>
          <th>Setter</th><th>Set</th><th>Issued</th><th>Demo</th><th>Sale</th>
          <th>CNS</th><th>1-Leg</th><th>CXL</th><th>NG</th><th>Drop</th>
          <th>Gross $</th><th>Net $</th><th>NSLI</th>
        </tr></thead>
        <tbody>{table_body}</tbody>
      </table>
    </div>
    {dispo_html}
  </div>''')

    btns = ' '.join(
        f'<button class="ptab{"" if p!="ytd" else " active"}" onclick="showPeriod(\'salesreps\',\'rep-{PMAP[p]}\',this)">{PLABELS[p]}</button>'
        for p in PERIODS)

    return f'''
<!-- ===== SETTERS ===== -->
<div class="section" id="salesreps">
  <div class="period-label">Setter Performance &nbsp;•&nbsp; Appointment Statistics by Setter</div>
  <div class="period-toggle">{btns}</div>
{"".join(sections)}
</div>
''', chart_js


def gen_callcenter_section(data_by_period, date_ranges_dict):
    """Generate the Call Center tab."""
    PMAP = {'prior_week':'pw','prior_month':'pm','mtd':'mtd','ytd':'ytd'}
    PLABELS = {'prior_week':'Prior Week','prior_month':'Prior Month','mtd':'Month to Date','ytd':'Year to Date'}
    sections = []
    chart_js = ''
    default = 'ytd'
    for period in PERIODS:
        pid = PMAP[period]
        label = PLABELS[period]
        _, _, date_lbl = date_ranges_dict[period]
        setter_rows, setter_total = data_by_period[period]['setter']
        prod_rows, prod_total     = data_by_period[period]['product']
        ss_rows, ss_total         = data_by_period[period]['subsource']

        t   = ss_total or {}
        set_v=t.get('set',0); gi=t.get('gross_issue',0); ni=t.get('net_issue',0)
        demo=t.get('demo',0); sale=t.get('sale',0)
        gross = setter_total['gross_amt'] if setter_total else 0

        gi_rate = pct(gi, set_v) if set_v else '—'
        dm_rate = pct(demo, gi) if gi else '—'
        cl_rate = pct(sale, demo) if demo else '—'

        active_setters = len([r for r in setter_rows if r['set'] > 0])

        setter_tbl = build_setter_table(setter_rows, setter_total)

        # Product simple table
        prod_lines = []
        for r in prod_rows:
            code = r['name'].upper()
            lbl  = PRODUCT_LABELS.get(code, code)
            prod_lines.append(f'<tr><td>{lbl} ({code})</td><td>{r["set"]}</td><td>{r["sale"]}</td><td>—</td></tr>')
        if prod_total:
            prod_lines.append(f'<tr class="gtot"><td>TOTAL</td><td>{prod_total["set"]}</td><td>{prod_total["sale"]}</td><td>{dollar(gross)}</td></tr>')
        prod_tbl = '\n'.join(prod_lines)

        # Setter set pie
        s_labels = [r['name'] for r in setter_rows if r['set'] > 0]
        s_sets   = [r['set'] for r in setter_rows if r['set'] > 0]
        # Product set pie
        p_labels = [f'{r["name"].upper()} ({PRODUCT_LABELS.get(r["name"].upper(), r["name"])})' for r in prod_rows if r['set'] > 0]
        p_sets   = [r['set'] for r in prod_rows if r['set'] > 0]

        chart_js += (
            f"pie('cc-{pid}-setter-pie',{js_str_arr(s_labels)},{js_arr(s_sets)});\n"
            f"pie('cc-{pid}-product-pie',{js_str_arr(p_labels)},{js_arr(p_sets)});\n"
        )

        active = ' active' if period == default else ''
        sections.append(f'''
  <div class="period-section{active}" id="cc-{pid}">
    <div class="period-label">{label} &nbsp;•&nbsp; {date_lbl}</div>
    <div class="kpi-grid">
      {kpi('Appointments Set', set_v, f'{active_setters} active setters', 'Appt Stats by Setter', 'green')}
      {kpi('Demos Run', demo, f'{dm_rate} of gross issued', 'Appt Stats by Sub-Source', 'orange')}
      {kpi('Sales', sale, f'{dollar(gross)} gross revenue', 'Appt Stats by Sub-Source', 'purple')}
      {kpi('Gross Issue Rate', gi_rate, f'{gi} issued / {set_v} set', 'Appt Stats by Sub-Source', 'teal')}
      {kpi('Demo Rate', dm_rate, f'{demo} demos / {gi} gross issued', 'Appt Stats by Sub-Source', 'green')}
      {kpi('Close Rate', cl_rate, f'{sale} sales / {demo} demos', 'Appt Stats by Sub-Source', 'red' if sale==0 else 'green')}
    </div>
    <div class="charts-grid">
      <div class="chart-card"><div class="chart-title">Set by Setter<span class="chart-subtitle">{label} — {set_v} total</span></div><div class="chart-wrap"><canvas id="cc-{pid}-setter-pie"></canvas></div></div>
      <div class="chart-card"><div class="chart-title">Set by Product<span class="chart-subtitle">{label} — {set_v} total</span></div><div class="chart-wrap"><canvas id="cc-{pid}-product-pie"></canvas></div></div>
    </div>
    <div class="table-card">
      <div class="chart-title">Appointment Statistics by Setter — {label}</div>
      <div class="table-note">Appt dates {date_lbl}.</div>
      <table><thead><tr><th>Setter</th><th>Set</th><th>Sales</th><th>Gross $</th><th>Rev / Set</th></tr></thead>
      <tbody>{setter_tbl}</tbody></table>
    </div>
    <div class="table-card">
      <div class="chart-title">Appointment Statistics by Product — {label}</div>
      <div class="table-note">Appt dates {date_lbl}.</div>
      <table><thead><tr><th>Product</th><th>Set</th><th>Sales</th><th>Gross $</th></tr></thead>
      <tbody>{prod_tbl}</tbody></table>
    </div>
  </div>''')

    btns = ' '.join(
        f'<button class="ptab{"" if p!="ytd" else " active"}" onclick="showPeriod(\'callcenter\',\'cc-{PMAP[p]}\',this)">{PLABELS[p]}</button>'
        for p in PERIODS)
    return f'''
<!-- ===== CALL CENTER ===== -->
<div class="section" id="callcenter">
  <div class="period-label">Call Center Performance &nbsp;•&nbsp; Appointment Statistics Reports</div>
  <div class="period-toggle">{btns}</div>
{''.join(sections)}
</div>
''', chart_js


def gen_promoter_section(data_by_period, date_ranges_dict):
    """Generate the Promoter Bonus Tracking tab."""
    PMAP   = {'prior_week':'pw','prior_month':'pm','mtd':'mtd','ytd':'ytd'}
    PLABELS= {'prior_week':'Prior Week','prior_month':'Prior Month','mtd':'Month to Date','ytd':'Year to Date'}
    COLORS = ['#2d5a8e','#28a745','#e67e22','#6f42c1']
    sections = []
    chart_js = ''
    default = 'ytd'

    # Collect all promoter names seen across all periods (preserve order, tracked first)
    all_names_seen = []
    for period in PERIODS:
        rows, _ = data_by_period[period]['promoter']
        for r in rows:
            first = r['name'].split(',')[0].strip() if ',' in r['name'] else r['name'].split()[0].strip()
            if first not in all_names_seen:
                all_names_seen.append(first)

    # YTD running totals for the summary banner
    ytd_rows, ytd_total = data_by_period['ytd']['promoter']

    for period in PERIODS:
        pid   = PMAP[period]
        label = PLABELS[period]
        _, _, date_lbl = date_ranges_dict[period]
        rows, total = data_by_period[period]['promoter']

        if not rows:
            sections.append(f'''
  <div class="period-section{"" if period != default else " active"}" id="promo-{pid}">
    <div class="period-label">{label} &nbsp;•&nbsp; {date_lbl}</div>
    <p style="padding:24px;color:#aaa;font-style:italic;">No promoter data found for this period — download the Appointment Statistics by Promoter report to populate this tab.</p>
  </div>''')
            continue

        # Build bar chart data — net issued leads per promoter
        bar_labels = [r['name'].split(',')[0].strip() if ',' in r['name'] else r['name'].split()[0].strip() for r in rows]
        bar_ni     = [r['net_issue'] for r in rows]
        bar_gi     = [r['gross_issue'] for r in rows]

        chart_js += (
            f"bar('promo-{pid}-bar',{js_str_arr(bar_labels)},"
            f"[{{label:'Net Issued',data:{js_arr(bar_ni)},backgroundColor:'#2d5a8e',borderRadius:4}},"
            f"{{label:'Gross Issued',data:{js_arr(bar_gi)},backgroundColor:'#aac4e0',borderRadius:4}}]);\n"
        )

        # KPI totals
        total_ni = total['net_issue'] if total else sum(r['net_issue'] for r in rows)
        total_gi = total['gross_issue'] if total else sum(r['gross_issue'] for r in rows)
        total_set = total['set'] if total else sum(r['set'] for r in rows)

        # Table rows
        tbl_rows = []
        for r in rows:
            first = r['name'].split(',')[0].strip() if ',' in r['name'] else r['name'].split()[0].strip()
            drop_pct = pct(r['drop'], r['gross_issue']) if r['gross_issue'] else '—'
            tbl_rows.append(
                f'<tr><td>{r["name"]}</td>'
                f'<td>{r["set"]}</td>'
                f'<td>{r["gross_issue"]}</td>'
                f'<td>{r["drop"]}</td>'
                f'<td class="ni-cell"><strong>{r["net_issue"]}</strong></td>'
                f'<td>{drop_pct}</td></tr>'
            )
        if total:
            tot_drop_pct = pct(total['drop'], total['gross_issue']) if total['gross_issue'] else '—'
            tbl_rows.append(
                f'<tr class="gtot"><td>TOTAL</td>'
                f'<td>{total_set}</td>'
                f'<td>{total_gi}</td>'
                f'<td>{total["drop"]}</td>'
                f'<td><strong>{total_ni}</strong></td>'
                f'<td>{tot_drop_pct}</td></tr>'
            )

        active = ' active' if period == default else ''
        sections.append(f'''
  <div class="period-section{active}" id="promo-{pid}">
    <div class="period-label">{label} &nbsp;•&nbsp; {date_lbl}</div>
    <div class="kpi-grid">
      {kpi('Total Appointments Set', total_set, f'{len(rows)} active promoters', 'Appt Stats by Promoter', 'green')}
      {kpi('Gross Issued Leads', total_gi, 'Leads passed to sales', 'Appt Stats by Promoter', 'teal')}
      {kpi('Net Issued Leads', total_ni, f'After {total["drop"] if total else 0} drops removed', 'Appt Stats by Promoter', 'purple')}
    </div>
    <div class="charts-grid" style="grid-template-columns:1fr;">
      <div class="chart-card">
        <div class="chart-title">Net Issued Leads by Promoter
          <span class="chart-subtitle">{label} — {total_ni} total net issued</span>
        </div>
        <div class="chart-wrap" style="height:280px;">
          <canvas id="promo-{pid}-bar"></canvas>
        </div>
      </div>
    </div>
    <div class="table-card">
      <div class="chart-title">Promoter Net Issued Lead Tally — {label}</div>
      <div class="table-note">Source: Appointment Statistics by Promoter. Net Issued = Gross Issued − Drops.</div>
      <table>
        <thead><tr>
          <th>Promoter</th><th>Appts Set</th><th>Gross Issued</th>
          <th>Drops</th><th>Net Issued ★</th><th>Drop %</th>
        </tr></thead>
        <tbody>{"".join(tbl_rows)}</tbody>
      </table>
    </div>
  </div>''')

    btns = ' '.join(
        f'<button class="ptab{"" if p!="ytd" else " active"}" onclick="showPeriod(\'promoter\',\'promo-{PMAP[p]}\',this)">{PLABELS[p]}</button>'
        for p in PERIODS)
    return f'''
<!-- ===== PROMOTER BONUS TRACKING ===== -->
<div class="section" id="promoter">
  <div class="period-label">Promoter Bonus Tracking &nbsp;•&nbsp; Net Issued Leads</div>
  <div class="period-toggle">{btns}</div>
{''.join(sections)}
</div>
''', chart_js


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    print("Loading Excel files…")
    data = load_all()
    dr   = date_ranges()

    import platform
    _dfmt = '%B %#d, %Y' if platform.system() == 'Windows' else '%B %-d, %Y'
    today_str = date.today().strftime(_dfmt)

    # Build source-group lookup from YTD marketing report (full year = most complete).
    # Used to correctly group CC sub-sources that have no raw leads in a given period.
    ytd_mkt_groups, _ = data['ytd']['marketing']
    ytd_source_map = {}
    for grp_name, grp in ytd_mkt_groups.items():
        for s in grp['subs']:
            ytd_source_map[s['name'].strip().lower()] = grp_name
    print(f"  YTD source map: {len(ytd_source_map)} sub-sources across {len(ytd_mkt_groups)} groups")

    # Main period sections
    PMAP = {'prior_week':('pw','Prior Week',12),
             'prior_month':('pm','Prior Month',12),
             'mtd':('mtd','Month to Date',12),
             'ytd':('ytd','Year to Date',12)}

    period_sections = ''
    chart_js = ''
    for period, (pid, label, colspan) in PMAP.items():
        _, _, date_lbl = dr[period]
        sec, cjs = gen_period_section(pid, label, date_lbl, data[period], colspan, ytd_source_map)
        period_sections += sec
        chart_js += cjs

    products_sec  = gen_products_section(data, dr)
    reps_sec, reps_js = gen_salesreps_section(data, dr)
    chart_js += reps_js
    cc_sec, cc_js = gen_callcenter_section(data, dr)
    chart_js += cc_js
    promo_sec, promo_js = gen_promoter_section(data, dr)
    chart_js += promo_js

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SuperFast Kitchen &amp; Bath — Marketing Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-annotation/3.0.1/chartjs-plugin-annotation.min.js"></script>
<script>window.ChartAnnotation=ChartAnnotation;</script>
<style>
{CSS}
</style>
</head>
<body>
<header>
  <div class="logo">🏠</div>
  <div>
    <h1>SuperFast Kitchen &amp; Bath — Marketing Dashboard</h1>
    <p>Generated {today_str} &nbsp;|&nbsp; Raw Leads: Marketing Sub-Source Report &nbsp;|&nbsp; Campaign Metrics: Call Center Reports</p>
  </div>
</header>
<div class="data-note">
  <strong>Data Sources:</strong> &nbsp;
  <span style="color:#6f42c1;font-weight:700">■ Raw Leads</span> = Marketing Sub-Source Report (lead entry date) &nbsp;&nbsp;
  <span style="color:#2d5a8e;font-weight:700">■ Set/Demo/Sales</span> = Call Center: Appointment Statistics by Sub-Source &nbsp;&nbsp;
  <span style="color:#28a745;font-weight:700">■ Products &amp; Setters</span> = Appt Statistics by Product / by Setter
</div>
<div class="tabs">
  <div class="tab active" onclick="showTab('pw',this)">Prior Week</div>
  <div class="tab" onclick="showTab('pm',this)">Prior Month</div>
  <div class="tab" onclick="showTab('mtd',this)">Month to Date</div>
  <div class="tab" onclick="showTab('ytd',this)">Year to Date</div>
  <div class="tab" onclick="showTab('products',this)">Products</div>
  <div class="tab" onclick="showTab('salesreps',this)">📞 Setters</div>
  <div class="tab" onclick="showTab('callcenter',this)">Call Center</div>
  <div class="tab" onclick="showTab('promoter',this)">🏆 Promoter Bonus</div>
</div>

{period_sections}
{products_sec}
{reps_sec}
{cc_sec}
{promo_sec}

<script>
{JS_HELPERS}
{chart_js}
</script>
</body>
</html>"""

    OUTPUT.write_text(html, encoding='utf-8')
    print(f"Dashboard written -> {OUTPUT}")

if __name__ == '__main__':
    run()
