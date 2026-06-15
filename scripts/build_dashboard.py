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
    """Parse appt_by_source / appt_by_subsource / appt_by_product."""
    ws = load_sheet(path)
    rows, total = [], None
    for r in range(ws.nrows):
        rv = ws.row_values(r)
        name = str(rv[0]).strip() if rv[0] != '' else ''
        if not name or is_footer(name): continue
        if name in ('Product', 'Source', 'Sub-Source', 'Setter'): continue

        if name == 'Total:':
            total = {
                'set': iv(gv(rv,2)), 'sale': iv(gv(rv,4)),
                'demo': iv(gv(rv,30)), 'net_issue': iv(gv(rv,32)),
                'gross_issue': iv(gv(rv,34)), 'drop': iv(gv(rv,36)),
            }
        else:
            set_v = iv(gv(rv,2))
            # include row if it has a numeric set value OR if it has any numeric data
            if set_v > 0 or any(isinstance(rv[c], float) for c in [4,31,33,35,37] if c < len(rv)):
                rows.append({
                    'name': name,
                    'set': set_v, 'sale': iv(gv(rv,4)),
                    'demo': iv(gv(rv,31)), 'net_issue': iv(gv(rv,33)),
                    'gross_issue': iv(gv(rv,35)), 'drop': iv(gv(rv,37)),
                })
    return rows, total

def parse_setter(path):
    """Parse appt_by_setter."""
    ws = load_sheet(path)
    rows, total = [], None
    for r in range(ws.nrows):
        rv = ws.row_values(r)
        name = str(rv[0]).strip() if rv[0] != '' else ''
        if not name or is_footer(name): continue
        if name == 'Setter': continue

        if name == 'Total:':
            total = {
                'set': iv(gv(rv,3)), 'gross_issue': iv(gv(rv,5)),
                'demo': iv(gv(rv,8)), 'sale': iv(gv(rv,11)),
                'gross_amt': fv(gv(rv,32)), 'drop': iv(gv(rv,39)),
            }
        elif iv(gv(rv,3)) > 0 or name:
            rows.append({
                'name': name,
                'set': iv(gv(rv,3)), 'gross_issue': iv(gv(rv,5)),
                'demo': iv(gv(rv,9)), 'sale': iv(gv(rv,11)),
                'gross_amt': fv(gv(rv,33)), 'drop': iv(gv(rv,39)),
            })
    return rows, total

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
        data[period] = {
            'source':    parse_cc(d / 'appt_by_source.xlsx'),
            'subsource': parse_cc(d / 'appt_by_subsource.xlsx'),
            'product':   parse_cc(d / 'appt_by_product.xlsx'),
            'setter':    parse_setter(d / 'appt_by_setter.xlsx'),
            'marketing': parse_marketing(d / 'marketing_sub_source.xlsx'),
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

def build_subsource_table(mkt_groups, mkt_grand, ss_rows, ss_total, colspan=12):
    """Build the combined sub-source HTML table body."""
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
            i_pct = pct(gi, set_v)
            d_pct = pct(demo, ni) if ni else '—'
            c_pct = pct(sale, demo) if demo else '—'
        else:
            set_v = gi = ni = demo = sale = drop = 0
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
                    f'{td(drop)}<td>$0</td>')
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
                    f'{td(drop)}<td>$0</td>')

    # Identify marketing source groups and their sub-sources
    for grp_name, grp in mkt_groups.items():
        grp_subs = grp['subs']
        if not grp_subs and grp['raw'] == 0:
            continue

        lines.append(f'<tr class="div-row"><td colspan="{colspan}">{grp_name.upper()}</td></tr>')

        grp_set = grp_cc_gi = grp_ni = grp_demo = grp_sale = grp_drop = grp_raw = 0

        for s in grp_subs:
            key = s['name'].strip().lower()
            cc_row = cc.get(key)
            if cc_row:
                placed.add(key)
            raw_v = s['raw']
            grp_raw += raw_v

            # Build display name for first column
            disp_name = s['name'].strip()
            r_cc = cc_row or {'set':0,'gross_issue':0,'net_issue':0,'demo':0,'sale':0,'drop':0}
            set_v  = r_cc['set'];  gi = r_cc['gross_issue']
            ni     = r_cc.get('net_issue', gi); demo = r_cc['demo']
            sale   = r_cc['sale']; drop = r_cc['drop']
            grp_set += set_v; grp_cc_gi += gi; grp_ni += ni
            grp_demo += demo; grp_sale += sale; grp_drop += drop

            i_pct = pct(gi, set_v) if set_v else '—'
            d_pct = pct(demo, ni) if ni else '—'
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
                    f'{td(drop)}<td>$0</td></tr>')
            else:
                lines.append(
                    f'<tr><td>&nbsp;&nbsp;{disp_name}</td>'
                    f'<td class="mc">{raw_d}</td>'
                    f'{td(set_v)}{td(gi)}'
                    f'<td{rate_class(i_pct)}>{i_pct}</td>'
                    f'{td(demo)}'
                    f'<td{rate_class(d_pct)}>{d_pct}</td>'
                    f'{td(sale)}<td{rate_class(c_pct)}>{c_pct}</td>'
                    f'{td(drop)}<td>$0</td></tr>')

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
                f'{td(grp_drop)}<td>$0</td></tr>')
        else:
            lines.append(
                f'<tr class="grp"><td>&nbsp;&nbsp;{grp_name} Total</td>'
                f'<td class="mc">{raw_d}</td>'
                f'{td(grp_set)}{td(grp_cc_gi)}'
                f'<td>{gi_p}</td>{td(grp_demo)}'
                f'<td>{dm_p}</td>{td(grp_sale)}<td>{cl_p}</td>'
                f'{td(grp_drop)}<td>$0</td></tr>')

    # CC sub-sources not matched to any marketing group
    unplaced = [r for r in ss_rows if r['name'].strip().lower() not in placed]
    if unplaced:
        lines.append(f'<tr class="div-row"><td colspan="{colspan}">OTHER</td></tr>')
        for r in unplaced:
            set_v=r['set']; gi=r['gross_issue']; ni=r.get('net_issue',gi)
            demo=r['demo']; sale=r['sale']; drop=r['drop']
            i_pct=pct(gi,set_v) if set_v else '—'
            d_pct=pct(demo,ni) if ni else '—'
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
                    f'{td(drop)}<td>$0</td></tr>')
            else:
                lines.append(
                    f'<tr><td>&nbsp;&nbsp;{r["name"]}</td>'
                    f'<td class="mc">—</td>'
                    f'{td(set_v)}{td(gi)}'
                    f'<td{rate_class(i_pct)}>{i_pct}</td>'
                    f'{td(demo)}'
                    f'<td{rate_class(d_pct)}>{d_pct}</td>'
                    f'{td(sale)}<td{rate_class(c_pct)}>{c_pct}</td>'
                    f'{td(drop)}<td>$0</td></tr>')

    # Grand total row
    if ss_total:
        t=ss_total
        total_raw = mkt_grand['raw'] if mkt_grand else '—'
        gi_p = pct(t['gross_issue'], t['set']) if t['set'] else '—'
        dm_p = pct(t['demo'], t['net_issue']) if t['net_issue'] else '—'
        cl_p = pct(t['sale'], t['demo']) if t['demo'] else '—'
        if colspan==12:
            lines.append(
                f'<tr class="gtot"><td>GRAND TOTAL</td>'
                f'<td>{total_raw}</td>'
                f'{td(t["set"])}{td(t["gross_issue"])}'
                f'<td>{gi_p}</td>{td(t["net_issue"])}{td(t["demo"])}'
                f'<td>{dm_p}</td>{td(t["sale"])}<td>{cl_p}</td>'
                f'{td(t["drop"])}<td>$0</td></tr>')
        else:
            lines.append(
                f'<tr class="gtot"><td>GRAND TOTAL</td>'
                f'<td>{total_raw}</td>'
                f'{td(t["set"])}{td(t["gross_issue"])}'
                f'<td>{gi_p}</td>{td(t["demo"])}'
                f'<td>{dm_p}</td>{td(t["sale"])}<td>{cl_p}</td>'
                f'{td(t["drop"])}<td>$0</td></tr>')

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
        return {k: sum(r.get(k,0) for r in rows) for k in ('set','sale','demo','net_issue','gross_issue','drop')}

    def prod_row(r):
        code = r['name'].upper()
        lbl  = PRODUCT_LABELS.get(code, code)
        gi   = r['gross_issue']; ni = r['net_issue']
        demo = r['demo']; sale = r['sale']
        ni_p = pct(ni, gi) if gi else '—'
        dm_p = pct(demo, ni) if ni else '—'
        cl_p = pct(sale, demo) if demo else '—'
        return (f'<tr><td>&nbsp;&nbsp;{code}</td><td>{lbl}</td>'
                f'<td>{gi}</td><td>{ni}</td>'
                f'<td{rate_class(ni_p)}>{ni_p}</td>'
                f'<td>{demo}</td>'
                f'<td{rate_class(dm_p)}>{dm_p}</td>'
                f'<td>{sale}</td>'
                f'<td{rate_class(cl_p) if sale>0 else ""}>{cl_p}</td>'
                f'<td>—</td></tr>')

    def grp_total_row(name, rows, css):
        s = sum_grp(rows)
        gi=s['gross_issue']; ni=s['net_issue']
        demo=s['demo']; sale=s['sale']
        ni_p=pct(ni,gi) if gi else '—'
        dm_p=pct(demo,ni) if ni else '—'
        cl_p=pct(sale,demo) if demo else '—'
        return (f'<tr class="grp-{css}"><td>&nbsp;&nbsp;{name}</td><td></td>'
                f'<td>{gi}</td><td>{ni}</td>'
                f'<td{rate_class(ni_p)}>{ni_p}</td>'
                f'<td>{demo}</td>'
                f'<td{rate_class(dm_p)}>{dm_p}</td>'
                f'<td>{sale}</td><td>{cl_p}</td><td>$0</td></tr>')

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
        ni_p=pct(ni,gi) if gi else '—'
        dm_p=pct(demo,ni) if ni else '—'
        cl_p=pct(sale,demo) if demo else '—'
        lines.append(
            f'<tr class="gtot"><td>GRAND TOTAL</td><td></td>'
            f'<td>{gi}</td><td>{ni}</td><td>{ni_p}</td>'
            f'<td>{demo}</td><td>{dm_p}</td>'
            f'<td>{sale}</td><td>{cl_p}</td><td>$0</td></tr>')

    return '\n'.join(lines)


def build_setter_table(setter_rows, setter_total):
    """Build setter performance table for Call Center tab."""
    lines = []
    for r in setter_rows:
        sale=r['sale']; gross=r['gross_amt']; set_v=r['set']
        rev_per_set = f'${gross/set_v:,.0f}' if set_v and gross else '—'
        lines.append(
            f'<tr><td>{r["name"]}</td><td>{set_v}</td>'
            f'<td>{sale}</td><td>{dollar(gross)}</td>'
            f'<td>{rev_per_set}</td></tr>')
    if setter_total:
        t=setter_total; gross=t['gross_amt']
        rev=f'${gross/t["set"]:,.0f}' if t['set'] and gross else '—'
        lines.append(
            f'<tr class="gtot"><td>TOTAL</td><td>{t["set"]}</td>'
            f'<td>{t["sale"]}</td><td>{dollar(gross)}</td>'
            f'<td>{rev}</td></tr>')
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


def gen_period_section(pid, label, date_label, d, colspan):
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
    dm_rate  = pct(demo, ni) if ni else '—'
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
    table_body = build_subsource_table(mkt_groups, mkt_grand, ss_rows, ss_total, colspan)

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
    {kpi('Demos Run', demo, f'{dm_rate} of net issued', 'Call Center Report', 'orange')}
    {kpi('Sales', sale, f'{dollar(gross)} gross revenue', 'Call Center Report', 'purple')}
    {kpi('Gross Issue Rate', gi_rate, f'{gi} issued / {set_v} set', 'Call Center Report', 'teal')}
    {kpi('Demo Rate', dm_rate, f'{demo} demos / {ni} net issued', 'Call Center Report', 'green')}
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
    """Generate the Sales Reps tab from appt_by_setter."""
    PMAP = {'prior_week':'pw','prior_month':'pm','mtd':'mtd','ytd':'ytd'}
    PLABELS = {'prior_week':'Prior Week','prior_month':'Prior Month','mtd':'Month to Date','ytd':'Year to Date'}
    sections = []
    default = 'ytd'
    for period in PERIODS:
        pid = PMAP[period]
        label = PLABELS[period]
        _, _, date_lbl = date_ranges_dict[period]
        setter_rows, setter_total = data_by_period[period]['setter']
        table_body = build_setter_table(setter_rows, setter_total)
        active = ' active' if period == default else ''
        sections.append(f'''
  <div class="period-section{active}" id="rep-{pid}">
    <div class="period-label">{label} &nbsp;•&nbsp; {date_lbl}</div>
    <div class="table-card">
      <div class="chart-title">Setter Performance — {label}</div>
      <div class="table-note">Source: Appointment Statistics by Setter. Dates {date_lbl}.</div>
      <table><thead><tr><th>Setter</th><th>Set</th><th>Sales</th><th>Gross $</th><th>Rev / Set</th></tr></thead>
      <tbody>{table_body}</tbody></table>
    </div>
  </div>''')

    btns = ' '.join(
        f'<button class="ptab{"" if p!="ytd" else " active"}" onclick="showPeriod(\'salesreps\',\'rep-{PMAP[p]}\',this)">{PLABELS[p]}</button>'
        for p in PERIODS)
    return f'''
<!-- ===== SETTERS / CALL CENTER ===== -->
<div class="section" id="salesreps">
  <div class="period-label">Setter Performance &nbsp;•&nbsp; Appointment Statistics by Setter</div>
  <div class="period-toggle">{btns}</div>
{''.join(sections)}
</div>
'''


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
        dm_rate = pct(demo, ni) if ni else '—'
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
      {kpi('Demos Run', demo, f'{dm_rate} of net issued', 'Appt Stats by Sub-Source', 'orange')}
      {kpi('Sales', sale, f'{dollar(gross)} gross revenue', 'Appt Stats by Sub-Source', 'purple')}
      {kpi('Gross Issue Rate', gi_rate, f'{gi} issued / {set_v} set', 'Appt Stats by Sub-Source', 'teal')}
      {kpi('Demo Rate', dm_rate, f'{demo} demos / {ni} net issued', 'Appt Stats by Sub-Source', 'green')}
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


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    print("Loading Excel files…")
    data = load_all()
    dr   = date_ranges()

    import platform
    _dfmt = '%B %#d, %Y' if platform.system() == 'Windows' else '%B %-d, %Y'
    today_str = date.today().strftime(_dfmt)

    # Main period sections
    PMAP = {'prior_week':('pw','Prior Week',12),
             'prior_month':('pm','Prior Month',12),
             'mtd':('mtd','Month to Date',12),
             'ytd':('ytd','Year to Date',12)}

    period_sections = ''
    chart_js = ''
    for period, (pid, label, colspan) in PMAP.items():
        _, _, date_lbl = dr[period]
        sec, cjs = gen_period_section(pid, label, date_lbl, data[period], colspan)
        period_sections += sec
        chart_js += cjs

    products_sec  = gen_products_section(data, dr)
    reps_sec      = gen_salesreps_section(data, dr)
    cc_sec, cc_js = gen_callcenter_section(data, dr)
    chart_js += cc_js

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SuperFast Kitchen &amp; Bath — Marketing Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
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
</div>

{period_sections}
{products_sec}
{reps_sec}
{cc_sec}

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
