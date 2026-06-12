"""
Assembly Department Daily Operational Dashboard - Flask Backend
Reads D:/Dasboard_Data.xlsx and serves data as a REST API.
Supports live reload when the Excel file changes.
"""

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
import openpyxl
import os
import threading
import time
from datetime import datetime, date
from collections import defaultdict

app = Flask(__name__, static_folder='static', static_url_path='')
CORS(app)

EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Assembly_Daily_Data.xlsx')
_data_cache = {}
_last_modified = 0

INVENTORY_EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "input summary -May'26.xlsx")
_inventory_cache = {}
_inventory_last_modified = 0

FTA_EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Case assembly - LE & HE FTA - 26-27.xlsx')
_fta_detail_cache = {}
_fta_detail_last_modified = 0

_lock = threading.Lock()

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def safe_float(v):
    try:
        f = float(v)
        return f if f == f else 0.0   # NaN check
    except:
        return 0.0

def safe_str(v):
    if v is None:
        return ''
    return str(v).strip()

def fmt_date(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    import re
    if re.match(r'^\d{4}-\d{2}-\d{2}$', s):
        return s
    for fmt in ('%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except:
            pass
    return None

def normalize_group(g):
    if g is None:
        return None
    g = str(g).strip()
    if g.upper() in ('PED-AUTO', 'PED_AUTO', 'PED-AUTO '):
        return 'PED-AUTO'
    if g.upper() == 'PED-AUTO':
        return 'PED-AUTO'
    return g.upper()

VALID_GROUPS = {'A1','A2','A3','G1','G2','G3','G4','G5','G6','G7','G8','PED','PED-AUTO'}

# ─────────────────────────────────────────────
# Data Loading
# ─────────────────────────────────────────────

def load_data():
    global _data_cache, _last_modified
    try:
        mtime = os.path.getmtime(EXCEL_PATH)
        if mtime <= _last_modified and _data_cache:
            return _data_cache
    except:
        pass

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    except Exception as e:
        print(f"Error loading Excel: {e}")
        return _data_cache

    data = {}

    # ── Sheet: Output  (Production Out – completed/accepted orders) ──
    # Columns: Order no | MODEL | ACC | ACT GR | Date
    output_rows = []
    if 'Output ' in wb.sheetnames:
        ws = wb['Output ']
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if not row or len(row) < 5:
                continue
            order_no, model, acc, group, dt = row[0], row[1], row[2], row[3], row[4]
            if group is None or dt is None:
                continue
            g = normalize_group(group)
            if g not in VALID_GROUPS:
                continue
            acc_val = safe_float(acc)
            if acc_val == 0:
                continue
            dt_str = fmt_date(dt)
            if not dt_str:
                continue
            output_rows.append({
                'order_no': safe_str(order_no),
                'model':    safe_str(model),
                'qty':      acc_val,
                'group':    g,
                'date':     dt_str,
                'month':    dt_str[:7]
            })
    data['output'] = output_rows
    print(f"Output rows loaded: {len(output_rows)}")

    # ── Sheet: Input (WIP entering assembly groups) ──
    # Columns: Group | Model no | Qty | PPC DATE | Batch/Shift | Gr-INPUT DATE | Type | Remarks
    input_rows = []
    if 'Input' in wb.sheetnames:
        ws2 = wb['Input']
        for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, values_only=True):
            if not row or len(row) < 6:
                continue
            group, model, qty, ppc_date, batch, gr_input_date = \
                row[0], row[1], row[2], row[3], row[4], row[5]
            if group is None:
                continue
            g = normalize_group(group)
            if g not in VALID_GROUPS:
                continue
            qty_val = safe_float(qty)
            if qty_val == 0:
                continue
            
            # Determine valid date using PPC DATE, fallback to Gr-INPUT DATE
            dt_str = None
            if ppc_date not in (None, ''):
                dt_str = fmt_date(ppc_date)
            if not dt_str and gr_input_date not in (None, ''):
                dt_str = fmt_date(gr_input_date)
                
            if not dt_str:
                continue
                
            shift_val = safe_float(batch) if batch not in (None, '') else 0.0
            input_rows.append({
                'model':  safe_str(model),
                'qty':    qty_val,
                'group':  g,
                'date':   dt_str,
                'month':  dt_str[:7],
                'shift':  shift_val
            })
    data['input'] = input_rows
    print(f"Input rows loaded: {len(input_rows)}")

    # ── Sheet: Groupwise Acceptance Dashboard ──
    # This is the authoritative capacity/target/actuals summary.
    # Values are in THOUSANDS (e.g. 27.1 = 27,100 units)
    capacity_rows = []
    capacity_total = {}
    if 'Groupwise Acceptance Dasboard' in wb.sheetnames:
        ws3 = wb['Groupwise Acceptance Dasboard']
        for row in ws3.iter_rows(min_row=7, max_row=ws3.max_row, values_only=True):
            if not row or len(row) < 17:
                continue
            g_raw = safe_str(row[3])
            if not g_raw or g_raw == 'Group':
                continue
            if g_raw.upper() == 'TOTAL':
                capacity_total = {
                    'capacity':          safe_float(row[4]),
                    'rate_per_day':      safe_float(row[5]),
                    'target_month':      safe_float(row[6]) * 1000,
                    'target_till_date':  safe_float(row[7]) * 1000,
                    'actuals_qty':       safe_float(row[8]) * 1000,
                    'actuals_equiv':     safe_float(row[9]) * 1000,
                    'balance_till_date': safe_float(row[10]) * 1000,
                    'balance_month':     safe_float(row[11]) * 1000,
                    'present_rate':      safe_float(row[12]),
                    'rate_required':     safe_float(row[13]),
                    'current_pct':       safe_float(row[14]),
                    'month_pct':         safe_float(row[15]),
                    'attendance':        safe_float(row[16])
                }
                continue
            g = normalize_group(g_raw)
            capacity_rows.append({
                'group':             g,
                'capacity':          safe_float(row[4]),
                'rate_per_day':      safe_float(row[5]),
                'target_month':      safe_float(row[6]) * 1000,
                'target_till_date':  safe_float(row[7]) * 1000,
                'actuals_qty':       safe_float(row[8]) * 1000,
                'actuals_equiv':     safe_float(row[9]) * 1000,
                'balance_till_date': safe_float(row[10]) * 1000,
                'balance_month':     safe_float(row[11]) * 1000,
                'present_rate':      safe_float(row[12]),
                'rate_required':     safe_float(row[13]),
                'current_pct':       safe_float(row[14]),
                'month_pct':         safe_float(row[15]),
                'attendance':        safe_float(row[16])
            })
    data['capacity'] = capacity_rows
    data['capacity_total'] = capacity_total

    # ── Sheet: FTA ──
    fta_rows = []
    if 'FTA' in wb.sheetnames:
        ws4 = wb['FTA']
        current_month = None
        current_week = 'W1'
        rows = list(ws4.iter_rows(min_row=2, max_row=ws4.max_row, values_only=True))
        for idx, row in enumerate(rows):
            if not row or len(row) < 9:
                continue
            month_val, group = row[0], row[1]
            if month_val and month_val != 'MONTH':
                if isinstance(month_val, (datetime, date)):
                    current_month = month_val.strftime('%Y-%m')
                    # Look ahead to determine week number for this block
                    if idx + 1 < len(rows):
                        next_val = rows[idx + 1][0]
                        if next_val and str(next_val).strip().startswith('W'):
                            current_week = str(next_val).strip()
                        else:
                            current_week = 'W1'
                    else:
                        current_week = 'W1'
                else:
                    s = safe_str(month_val)
                    dt_parsed = fmt_date(month_val)
                    if dt_parsed:
                        current_month = dt_parsed[:7]
                        # Look ahead
                        if idx + 1 < len(rows):
                            next_val = rows[idx + 1][0]
                            if next_val and str(next_val).strip().startswith('W'):
                                current_week = str(next_val).strip()
                            else:
                                current_week = 'W1'
                        else:
                            current_week = 'W1'
                    elif s.startswith('W'):
                        current_week = s
                    elif 'Total' in s or 'total' in s:
                        continue
            g_raw = safe_str(group)
            if not g_raw or g_raw in ('Group', 'MONTH', ''):
                continue
            g = normalize_group(g_raw)
            fta_rows.append({
                'month':    current_month,
                'week':     current_week,
                'group':    g,
                'acc':      safe_float(row[2]),
                'rej':      safe_float(row[3]),
                'acc_100':  safe_float(row[4]),
                'rej_100':  safe_float(row[5]),
                'he_acc':   safe_float(row[6]),
                'he_rej':   safe_float(row[7]),
                'fta_pct':  safe_float(row[8])
            })
    data['fta'] = fta_rows

    with _lock:
        _data_cache = data
        try:
            _last_modified = os.path.getmtime(EXCEL_PATH)
        except:
            pass

    print(f"Data loaded: {len(output_rows)} out, {len(input_rows)} in, "
          f"{len(capacity_rows)} groups, {len(fta_rows)} FTA rows")
    return data


def load_inventory_data():
    global _inventory_cache, _inventory_last_modified
    try:
        if not os.path.exists(INVENTORY_EXCEL_PATH):
            return _inventory_cache
        mtime = os.path.getmtime(INVENTORY_EXCEL_PATH)
        if mtime <= _inventory_last_modified and _inventory_cache:
            return _inventory_cache
    except:
        pass

    try:
        wb = openpyxl.load_workbook(INVENTORY_EXCEL_PATH, data_only=True)
    except Exception as e:
        print(f"Error loading Inventory Excel: {e}")
        return _inventory_cache

    inv_data = {}

    # 1. Parse 'Summary ' sheet
    summary_data = {}
    if 'Summary ' in wb.sheetnames:
        ws = wb['Summary ']
        groups = []
        for c in range(3, 17):
            g = ws.cell(3, c).value
            if g:
                groups.append((c, normalize_group(g)))
        
        for r in range(4, 20):
            lbl_raw = ws.cell(r, 2).value
            if not lbl_raw:
                continue
            lbl = str(lbl_raw).strip().replace('\xa0', ' ')
            lbl_clean = lbl.lower().replace(' ', '_').replace(':', '')
            summary_data[lbl_clean] = {}
            for col_idx, g_name in groups:
                val = safe_float(ws.cell(r, col_idx).value)
                summary_data[lbl_clean][g_name] = val
    inv_data['summary'] = summary_data

    # 2. Parse 'WIP ' sheet
    wip_log = []
    if 'WIP ' in wb.sheetnames:
        ws = wb['WIP ']
        groups_wip = []
        for c in range(3, 17):
            g = ws.cell(5, c).value
            if g:
                groups_wip.append((c, normalize_group(g)))
        
        for r in range(6, ws.max_row + 1):
            dt_val = ws.cell(r, 2).value
            dt_str = fmt_date(dt_val)
            if not dt_str:
                continue
            for col_idx, g_name in groups_wip:
                val = ws.cell(r, col_idx).value
                if val is not None:
                    wip_log.append({
                        'date': dt_str,
                        'group': g_name,
                        'qty': safe_float(val)
                    })
    inv_data['wip'] = wip_log

    # 3. Parse 'Input ' sheet
    input_log = []
    if 'Input ' in wb.sheetnames:
        ws = wb['Input ']
        for r in range(2, ws.max_row + 1):
            g_raw = ws.cell(r, 1).value
            model = ws.cell(r, 2).value
            qty = ws.cell(r, 3).value
            dt_val = ws.cell(r, 4).value
            batch = ws.cell(r, 5).value
            mg_type = ws.cell(r, 10).value
            
            if not g_raw or qty is None:
                continue
            g = normalize_group(g_raw)
            dt_str = fmt_date(dt_val)
            if not dt_str:
                continue
            
            input_log.append({
                'group': g,
                'model': safe_str(model),
                'qty': safe_float(qty),
                'date': dt_str,
                'month': dt_str[:7],
                'batch': safe_float(batch) if batch is not None else 0.0,
                'mg_type': safe_str(mg_type)
            })
    inv_data['input'] = input_log

    # 4. Parse 'Output ' sheet
    output_log = []
    if 'Output ' in wb.sheetnames:
        ws = wb['Output ']
        for r in range(2, ws.max_row + 1):
            order = ws.cell(r, 1).value
            model = ws.cell(r, 2).value
            acc = ws.cell(r, 3).value
            g_raw = ws.cell(r, 4).value
            dt_val = ws.cell(r, 5).value
            
            if not g_raw or acc is None:
                continue
            g = normalize_group(g_raw)
            dt_str = fmt_date(dt_val)
            if not dt_str:
                continue
            
            output_log.append({
                'group': g,
                'model': safe_str(model),
                'qty': safe_float(acc),
                'date': dt_str,
                'month': dt_str[:7]
            })
    inv_data['output'] = output_log

    # 5. Parse 'Returns ' sheet
    stripping_log = []
    returns_log = []
    scrap_log = []
    if 'Returns ' in wb.sheetnames:
        ws = wb['Returns ']
        groups_ret = []
        for c in range(3, 16):
            g = ws.cell(5, c).value
            if g:
                groups_ret.append((c, normalize_group(g)))
        
        for r in range(6, 38):
            dt_val = ws.cell(r, 2).value
            dt_str = fmt_date(dt_val)
            if not dt_str:
                continue
            for col_idx, g_name in groups_ret:
                val = ws.cell(r, col_idx).value
                if val is not None:
                    stripping_log.append({
                        'date': dt_str,
                        'group': g_name,
                        'qty': safe_float(val)
                    })
        
        groups_ret2 = []
        for c in range(3, 16):
            g = ws.cell(42, c).value
            if g:
                groups_ret2.append((c, normalize_group(g)))
                
        for r in range(43, 71):
            dt_val = ws.cell(r, 2).value
            dt_str = fmt_date(dt_val)
            if not dt_str:
                continue
            for col_idx, g_name in groups_ret2:
                val = ws.cell(r, col_idx).value
                if val is not None:
                    returns_log.append({
                        'date': dt_str,
                        'group': g_name,
                        'qty': safe_float(val)
                    })
                    
        groups_ret3 = []
        for c in range(3, 16):
            g = ws.cell(74, c).value
            if g:
                groups_ret3.append((c, normalize_group(g)))
                
        for r in range(75, 106):
            dt_val = ws.cell(r, 2).value
            dt_str = fmt_date(dt_val)
            if not dt_str:
                continue
            for col_idx, g_name in groups_ret3:
                val = ws.cell(r, col_idx).value
                if val is not None:
                    scrap_log.append({
                        'date': dt_str,
                        'group': g_name,
                        'qty': safe_float(val)
                    })

    inv_data['stripping_log'] = stripping_log
    inv_data['returns_log'] = returns_log
    inv_data['scrap_log'] = scrap_log

    with _lock:
        _inventory_cache = inv_data
        try:
            _inventory_last_modified = os.path.getmtime(INVENTORY_EXCEL_PATH)
        except:
            pass

    print(f"Inventory data loaded: {len(input_log)} in, {len(output_log)} out, {len(wip_log)} WIP, "
          f"{len(stripping_log)} stripping, {len(returns_log)} returns, {len(scrap_log)} scrap")
    return inv_data


# ─────────────────────────────────────────────
# File Watcher
# ─────────────────────────────────────────────

def load_fta_detail():
    """Load LE & HE FTA data from 'Case assembly - LE & HE FTA - 26-27.xlsx'.
    Returns monthly FTA% (lot-based for LE, unit-based for HE), defect breakdown,
    and stripping %."""
    global _fta_detail_cache, _fta_detail_last_modified
    try:
        if not os.path.exists(FTA_EXCEL_PATH):
            return _fta_detail_cache
        mtime = os.path.getmtime(FTA_EXCEL_PATH)
        if mtime <= _fta_detail_last_modified and _fta_detail_cache:
            return _fta_detail_cache
    except:
        pass

    try:
        wb = openpyxl.load_workbook(FTA_EXCEL_PATH, data_only=True)
    except Exception as e:
        print(f"Error loading FTA Excel: {e}")
        return _fta_detail_cache

    result = {}

    # -- LOW END sheet (sampling-based inspection) --
    # For LE, FTA = accepted lots / total lots (verdict based)
    le_monthly = defaultdict(lambda: {
        'lot_qty': 0, 'acc': 0, 'rej': 0,
        'acc_lots': 0, 'rej_lots': 0, 'total_lots': 0,
        'defects': defaultdict(float)
    })
    le_defect_cols = list(range(11, 37))  # 0-indexed columns 11-36

    if 'Low End ' in wb.sheetnames:
        ws_le = wb['Low End ']
        le_headers = [cell.value for cell in ws_le[1]]
        for row in ws_le.iter_rows(min_row=2, values_only=True):
            if not row[1]:
                continue
            month = row[1]
            if isinstance(month, (datetime, date)):
                month_key = month.strftime('%Y-%m')
            else:
                continue
            d = le_monthly[month_key]
            d['lot_qty']    += safe_float(row[5])
            d['acc']        += safe_float(row[7])
            d['rej']        += safe_float(row[8])
            d['total_lots'] += 1
            verdict = safe_str(row[9]).upper()
            if verdict == 'ACC':
                d['acc_lots'] += 1
            else:
                d['rej_lots'] += 1
            for ci in le_defect_cols:
                if ci < len(row) and row[ci]:
                    col_name = le_headers[ci] if ci < len(le_headers) else None
                    if col_name:
                        d['defects'][col_name] += safe_float(row[ci])

    # -- HIGH END sheet (100% inspection) --
    # For HE, FTA = accepted qty / total lot qty
    he_monthly = defaultdict(lambda: {
        'lot_qty': 0, 'acc': 0, 'rej': 0,
        'defects': defaultdict(float)
    })
    he_defect_cols = list(range(8, 34))  # 0-indexed columns 8-33

    if 'High end' in wb.sheetnames:
        ws_he = wb['High end']
        he_headers = [cell.value for cell in ws_he[1]]
        for row in ws_he.iter_rows(min_row=2, values_only=True):
            if not row[1]:
                continue
            month = row[1]
            if isinstance(month, (datetime, date)):
                month_key = month.strftime('%Y-%m')
            else:
                continue
            d = he_monthly[month_key]
            d['lot_qty'] += safe_float(row[5])
            d['acc']     += safe_float(row[6])
            d['rej']     += safe_float(row[7])
            for ci in he_defect_cols:
                if ci < len(row) and row[ci]:
                    col_name = he_headers[ci] if ci < len(he_headers) else None
                    if col_name:
                        d['defects'][col_name] += safe_float(row[ci])

    # -- CA ASSLY MOM OUTPUT (group-wise monthly FTA from the MOM sheet) --
    mom_monthly = defaultdict(list)
    if 'CA ASSLY MOM OUTPUT' in wb.sheetnames:
        ws_mom = wb['CA ASSLY MOM OUTPUT']
        current_month_mom = None
        for row in ws_mom.iter_rows(min_row=4, values_only=True):
            if not row or len(row) < 8:
                continue
            month_val = row[1]
            group_val = row[2]
            if month_val and isinstance(month_val, (datetime, date)):
                current_month_mom = month_val.strftime('%Y-%m')
            if not group_val or not current_month_mom:
                continue
            grp = safe_str(group_val)
            if grp.lower() in ('group', 'total', ''):
                continue
            output_v  = safe_float(row[3])
            shop_rej  = safe_float(row[4])
            iqf_rej   = safe_float(row[5])
            cum_rej   = safe_float(row[6])
            fta_v     = safe_float(row[7])
            mom_monthly[current_month_mom].append({
                'group':    normalize_group(grp),
                'output':   round(output_v),
                'shop_rej': round(shop_rej),
                'iqf_rej':  round(iqf_rej),
                'cum_rej':  round(cum_rej),
                'fta':      round(fta_v * 100, 2) if fta_v <= 1 else round(fta_v, 2)
            })

    # Build monthly summary
    all_months = sorted(set(list(le_monthly.keys()) + list(he_monthly.keys())))
    monthly_summary = []
    for m in all_months:
        le = le_monthly.get(m, {})
        he = he_monthly.get(m, {})

        # LE: lot-based FTA (accepted lots / total lots)
        le_total_lots = le.get('total_lots', 0)
        le_acc_lots   = le.get('acc_lots', 0)
        le_rej_lots   = le.get('rej_lots', 0)
        le_fta = round(le_acc_lots / le_total_lots * 100, 2) if le_total_lots > 0 else 0

        # HE: unit-based FTA (accepted units / total lot qty)
        he_lot = he.get('lot_qty', 0)
        he_acc = he.get('acc', 0)
        he_rej = he.get('rej', 0)
        he_fta = round(he_acc / he_lot * 100, 2) if he_lot > 0 else 0

        monthly_summary.append({
            'month':       m,
            # LE lot-based
            'le_lots':     le_total_lots,
            'le_acc_lots': le_acc_lots,
            'le_rej_lots': le_rej_lots,
            'le_lot_qty':  round(le.get('lot_qty', 0)),
            'le_fta':      le_fta,
            # HE unit-based
            'he_lot':      round(he_lot),
            'he_acc':      round(he_acc),
            'he_rej':      round(he_rej),
            'he_fta':      he_fta,
        })

    # Build combined defect summary (all months)
    combined_le_defects = defaultdict(float)
    combined_he_defects = defaultdict(float)
    for m, d in le_monthly.items():
        for k, v in d['defects'].items():
            combined_le_defects[k] += v
    for m, d in he_monthly.items():
        for k, v in d['defects'].items():
            combined_he_defects[k] += v

    # Top defects combined
    all_defect_keys = set(list(combined_le_defects.keys()) + list(combined_he_defects.keys()))
    combined_all = {k: combined_le_defects.get(k, 0) + combined_he_defects.get(k, 0)
                    for k in all_defect_keys}
    top_defects = sorted(combined_all.items(), key=lambda x: -x[1])[:12]

    result['monthly']      = monthly_summary
    result['le_defects']   = sorted(combined_le_defects.items(), key=lambda x: -x[1])[:12]
    result['he_defects']   = sorted(combined_he_defects.items(), key=lambda x: -x[1])[:12]
    result['top_defects']  = top_defects
    result['mom_monthly']  = dict(mom_monthly)

    with _lock:
        _fta_detail_cache = result
        try:
            _fta_detail_last_modified = os.path.getmtime(FTA_EXCEL_PATH)
        except:
            pass

    print(f"FTA detail loaded: {len(monthly_summary)} months, {len(top_defects)} defect types, "
          f"{sum(len(v) for v in mom_monthly.values())} MOM group rows")
    return result



def watch_file():
    global _last_modified, _inventory_last_modified, _fta_detail_last_modified
    while True:
        try:
            mtime = os.path.getmtime(EXCEL_PATH)
            if mtime != _last_modified:
                print("Excel file changed – reloading...")
                load_data()
        except:
            pass
        try:
            if os.path.exists(INVENTORY_EXCEL_PATH):
                mtime_inv = os.path.getmtime(INVENTORY_EXCEL_PATH)
                if mtime_inv != _inventory_last_modified:
                    print("Inventory Excel file changed – reloading...")
                    load_inventory_data()
        except:
            pass
        try:
            if os.path.exists(FTA_EXCEL_PATH):
                mtime_fta = os.path.getmtime(FTA_EXCEL_PATH)
                if mtime_fta != _fta_detail_last_modified:
                    print("FTA Excel file changed – reloading...")
                    load_fta_detail()
        except:
            pass
        time.sleep(3)


# ─────────────────────────────────────────────
# Filter helpers
# ─────────────────────────────────────────────

def filter_by(rows, date_f, model_f, group_f, shift_f=None,
              date_col='date', group_col='group', model_col='model', shift_col=None,
              dates_list=None):
    result = []
    for r in rows:
        # Dates list filter
        if dates_list is not None:
            if r.get(date_col, '') not in dates_list:
                continue
        # Date
        elif date_f and date_f not in ('month', 'all'):
            if len(date_f) == 7: # Month filter, e.g. '2026-05'
                if r.get('month', '') != date_f:
                    continue
            else: # Specific date filter, e.g. '2026-05-15'
                if r.get(date_col, '') != date_f:
                    continue
        # Model
        if model_f and model_f.lower() != 'all':
            if r.get(model_col, '').lower() != model_f.lower():
                continue
        # Group
        if group_f and group_f.lower() not in ('all', 'overall'):
            if r.get(group_col, '').upper() != group_f.upper():
                continue
        # Shift  (batch column: 1=Shift1, 2=Shift2, 0=General)
        if shift_f and shift_f != 'all' and shift_col:
            shift_map = {'1': 1.0, 'shift1': 1.0, 'shift 1': 1.0,
                         '2': 2.0, 'shift2': 2.0, 'shift 2': 2.0,
                         'general': 0.0}
            wanted = shift_map.get(shift_f.lower())
            if wanted is not None and r.get(shift_col) != wanted:
                continue
        result.append(r)
    return result


# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/api/meta')
def meta():
    """Return filter metadata for dashboard controls.

    Query Parameters:
        None

    Returns:
        flask.Response: JSON object with:
            - dates (list[str]): Distinct production dates in YYYY-MM-DD format.
            - months (list[str]): Distinct months in YYYY-MM format.
            - groups (list[str]): Distinct assembly groups.
            - models (list[str]): Distinct model names (first 300 values).
            - month_weeks (dict[str, list[dict[str, object]]]): Week buckets per month.
            - coverage (dict[str, list[str]]): Source-month availability for output/input.
            - last_updated (str): ISO timestamp when metadata was generated.
    """
    data = load_data()
    output = data.get('output', [])
    inp    = data.get('input', [])

    all_dates = sorted(set(r['date'] for r in output + inp if r.get('date')))
    all_months = sorted(set(r['month'] for r in output + inp if r.get('month')))
    all_groups = sorted(set(r['group'] for r in output + inp if r.get('group')))
    all_models = sorted(set(r['model'] for r in output + inp if r.get('model')))

    # Coverage info to show in UI
    out_months = sorted(set(r['month'] for r in output if r.get('month')))
    in_months  = sorted(set(r['month'] for r in inp if r.get('month')))

    # Group dates by ISO week dynamically
    month_weeks = {}
    for m in all_months:
        m_dates = [d for d in all_dates if d.startswith(m)]
        iso_to_dates = defaultdict(list)
        for d in m_dates:
            try:
                dt = datetime.strptime(d, '%Y-%m-%d')
                week_num = dt.isocalendar()[1]
                iso_to_dates[week_num].append(d)
            except:
                pass
        
        sorted_weeks = sorted(iso_to_dates.keys())
        weeks_list = []
        for idx, w in enumerate(sorted_weeks):
            w_dates = sorted(iso_to_dates[w])
            try:
                start_lbl = datetime.strptime(w_dates[0], '%Y-%m-%d').strftime('%d %b')
                end_lbl = datetime.strptime(w_dates[-1], '%Y-%m-%d').strftime('%d %b')
                label = f"Week {idx+1} ({start_lbl} - {end_lbl})"
            except:
                label = f"Week {idx+1}"
            weeks_list.append({
                'id': f"W{idx+1}", # We use dynamic IDs like W1, W2 etc.
                'label': label,
                'dates': w_dates
            })
        month_weeks[m] = weeks_list

    return jsonify({
        'dates':      all_dates,
        'months':     all_months,
        'groups':     all_groups,
        'models':     all_models[:300],
        'month_weeks': month_weeks,
        'coverage': {
            'output_months': out_months,
            'input_months':  in_months
        },
        'last_updated': datetime.now().isoformat()
    })


@app.route('/api/production')
def production():
    """Return production KPIs and charts for selected filters.

    Query Parameters:
        date (str): "month"/"all", YYYY-MM, or YYYY-MM-DD. Defaults to "month".
        shift (str): Shift selector ("all", "1", "2", "general"). Defaults to "all".
        model (str): Model filter. Defaults to "all".
        group (str): Group filter. Defaults to "all".
        dates (str): Optional comma-separated YYYY-MM-DD list for explicit date filtering.
        working_days (str): Optional numeric override for working-day calculations.

    Returns:
        flask.Response: JSON object with:
            - timeline (dict[str, object]): Date-wise input/output trend and month coverage.
            - groupwise (dict[str, object]): Group labels, actuals, targets, input, till-date.
            - flow (dict[str, float]): Target/actual aggregates and percentage metrics.
            - capacity_detail (list[dict[str, float]]): Capacity rows after filter adjustments.
    """
    data = load_data()
    date_f  = request.args.get('date',  'month')
    shift_f = request.args.get('shift', 'all')
    model_f = request.args.get('model', 'all')
    group_f = request.args.get('group', 'all')

    dates_arg = request.args.get('dates')
    dates_list = dates_arg.split(',') if dates_arg else None

    working_days_arg = request.args.get('working_days')
    manual_working_days = safe_float(working_days_arg) if working_days_arg else None

    out_rows = filter_by(data.get('output', []), date_f, model_f, group_f,
                         date_col='date', group_col='group', model_col='model',
                         dates_list=dates_list)
    in_rows  = filter_by(data.get('input', []),  date_f, model_f, group_f, shift_f,
                         date_col='date', group_col='group', model_col='model',
                         shift_col='shift', dates_list=dates_list)

    # ── Daily timeline (Production Out only — this is the reliable daily data) ──
    daily_out = defaultdict(float)
    daily_in  = defaultdict(float)
    group_out = defaultdict(float)
    group_in  = defaultdict(float)

    for r in out_rows:
        daily_out[r['date']] += r['qty']
        group_out[r['group']] += r['qty']

    for r in in_rows:
        daily_in[r['date']] += r['qty']
        group_in[r['group']] += r['qty']

    all_dates  = sorted(set(list(daily_out.keys()) + list(daily_in.keys())))
    all_groups = sorted(set(list(group_out.keys()) + list(group_in.keys())))

    # ── Capacity / targets ──
    capacity = data.get('capacity', [])
    
    # Calculate group actuals dynamically from out_rows
    group_actuals = defaultdict(float)
    for r in out_rows:
        group_actuals[r['group']] += r['qty']

    # Deep copy and update capacity rows dynamically
    import copy
    cap_f = []
    for c in capacity:
        if group_f.lower() in ('all', 'overall') or c['group'].upper() == group_f.upper():
            c_copy = copy.deepcopy(c)
            g = c_copy['group']
            c_copy['actuals_qty'] = group_actuals.get(g, 0.0)
            c_copy['balance_month'] = max(0.0, c_copy['target_month'] - c_copy['actuals_qty'])
            c_copy['month_pct'] = (c_copy['actuals_qty'] / c_copy['target_month']) if c_copy['target_month'] > 0 else 0.0
            cap_f.append(c_copy)

    # Calculate overall parameters from cap_f
    target_month = sum(c['target_month'] for c in cap_f)
    target_till  = sum(c['target_till_date'] for c in cap_f)
    actuals_qty  = sum(c['actuals_qty'] for c in cap_f)
    month_pct    = (actuals_qty / target_month) if target_month > 0 else 0
    current_pct  = (actuals_qty / target_till) if target_till > 0 else 0
    rate_day     = sum(c['rate_per_day'] for c in cap_f) * 1000
    capacity_val = sum(c['capacity'] for c in cap_f)

    # Group targets map
    group_daily_rate = {c['group']: c['rate_per_day'] * 1000 for c in capacity}

    # Calculate target and actuals based on the date filter
    if date_f and date_f not in ('all', 'month'):
        if len(date_f) == 10:  # Specific daily date
            # Daily target is the planned daily rate
            flow_in = rate_day
            # Daily actual is the sum of actual output on that day
            flow_out = sum(r['qty'] for r in out_rows)
            pct_till = round((flow_out / flow_in * 100) if flow_in > 0 else 0, 1)
            pct_month = pct_till
            
            # Use daily targets for a specific date
            targets_list = [round(group_daily_rate.get(g, 0)) for g in all_groups]
            till_date_list = targets_list
        else:  # Specific month, e.g. '2026-05' or '2026-06'
            flow_out = sum(r['qty'] for r in out_rows)
            distinct_dates = set(r['date'] for r in out_rows)
            
            if manual_working_days is not None:
                working_days = manual_working_days
            else:
                working_days = len(distinct_dates) if distinct_dates else 22
                
            flow_in = rate_day * working_days
            pct_till = round((flow_out / flow_in * 100) if flow_in > 0 else 0, 1)
            pct_month = round((flow_out / target_month * 100) if target_month > 0 else 0, 1)

            if dates_list is not None or manual_working_days is not None:
                targets_list = [round(group_daily_rate.get(g, 0) * working_days) for g in all_groups]
                till_date_list = targets_list
            else:
                # Group targets
                group_targets = {c['group']: c['target_month'] for c in cap_f}
                group_till_date = {c['group']: c['target_till_date'] for c in cap_f}
                targets_list = [round(group_targets.get(g, 0)) for g in all_groups]
                till_date_list = [round(group_till_date.get(g, 0)) for g in all_groups]
    else:  # 'all' or 'month' (default overview)
        if manual_working_days is not None:
            working_days = manual_working_days
            flow_in = rate_day * working_days
            flow_out = actuals_qty
            pct_till = round((flow_out / flow_in * 100) if flow_in > 0 else 0, 1)
            pct_month = round((flow_out / target_month * 100) if target_month > 0 else 0, 1)

            targets_list = [round(group_daily_rate.get(g, 0) * working_days) for g in all_groups]
            till_date_list = targets_list
        else:
            flow_in = target_till
            flow_out = actuals_qty
            pct_till = round(current_pct * 100, 1)
            pct_month = round(month_pct * 100, 1)

            # Group targets
            group_targets = {c['group']: c['target_month'] for c in cap_f}
            group_till_date = {c['group']: c['target_till_date'] for c in cap_f}
            targets_list = [round(group_targets.get(g, 0)) for g in all_groups]
            till_date_list = [round(group_till_date.get(g, 0)) for g in all_groups]

    return jsonify({
        'timeline': {
            'dates':          all_dates,
            'production_out': [round(daily_out.get(d, 0)) for d in all_dates],
            'production_in':  [round(daily_in.get(d, 0))  for d in all_dates],
            'out_months':     sorted(set(r['month'] for r in out_rows)),
            'in_months':      sorted(set(r['month'] for r in in_rows))
        },
        'groupwise': {
            'groups':   all_groups,
            'actuals':  [round(group_out.get(g, 0)) for g in all_groups],
            'targets':  targets_list,
            'input':    [round(group_in.get(g, 0))  for g in all_groups],
            'till_date': till_date_list
        },
        'flow': {
            # Capacity-based (authoritative Target vs Actuals comparison)
            'target_till_date': round(flow_in),
            'actuals':          round(flow_out),
            'target_month':     round(target_month),
            'pct_till_date':    pct_till,
            'pct_month':        pct_month,
            'daily_rate':       round(rate_day),
            'capacity':         capacity_val,
            # Raw totals from sheets (for reference)
            'raw_output_total': round(sum(daily_out.values())),
            'raw_input_total':  round(sum(daily_in.values()))
        },
        'capacity_detail': cap_f
    })


@app.route('/api/fta_detail')
def fta_detail():
    """Serve LE & HE FTA % and defect breakdown from the FTA Excel file."""
    data = load_data()
    fta  = load_fta_detail()
    inv  = load_inventory_data()  # for stripping

    date_f  = request.args.get('date',  'month')
    group_f = request.args.get('group', 'all')

    # Filter monthly summary by date/month
    monthly = fta.get('monthly', [])
    if date_f and date_f not in ('all', 'month'):
        target_month = date_f[:7]
        monthly = [m for m in monthly if m['month'] == target_month]

    # Compute stripping % from inventory data
    output_rows  = data.get('output', [])
    strip_data   = inv.get('stripping_log', [])

    # Apply group filter
    if group_f.lower() not in ('all', 'overall'):
        output_rows = [r for r in output_rows if r['group'].upper() == group_f.upper()]
        strip_data  = [r for r in strip_data  if r['group'].upper() == group_f.upper()]

    # Apply date/month filter to stripping and output
    if date_f and date_f not in ('all', 'month'):
        target_month = date_f[:7]
        output_rows = [r for r in output_rows if r.get('month', '') == target_month]
        strip_data  = [r for r in strip_data  if r['date'].startswith(target_month)]

    total_output   = sum(r['qty'] for r in output_rows)
    total_stripping = sum(r['qty'] for r in strip_data)
    stripping_pct  = round(total_stripping / total_output * 100, 2) if total_output > 0 else 0

    # Compute per-month stripping %
    monthly_output   = defaultdict(float)
    monthly_stripping = defaultdict(float)
    for r in data.get('output', []):
        monthly_output[r['month']] += r['qty']
    for r in inv.get('stripping_log', []):
        monthly_stripping[r['date'][:7]] += r['qty']

    strip_monthly = []
    for m in sorted(set(list(monthly_output.keys()) + list(monthly_stripping.keys()))):
        out = monthly_output.get(m, 0)
        stp = monthly_stripping.get(m, 0)
        strip_monthly.append({
            'month':         m,
            'output':        round(out),
            'stripping':     round(stp),
            'stripping_pct': round(stp / out * 100, 2) if out > 0 else 0
        })

    return jsonify({
        'monthly':         monthly,
        'le_defects':      fta.get('le_defects', []),
        'he_defects':      fta.get('he_defects', []),
        'top_defects':     fta.get('top_defects', []),
        'mom_monthly':     fta.get('mom_monthly', {}),
        'stripping': {
            'total_output':    round(total_output),
            'total_stripping': round(total_stripping),
            'stripping_pct':   stripping_pct,
            'monthly':         strip_monthly
        }
    })


@app.route('/api/quality')
def quality():
    """Return FTA quality metrics aggregated overall and by group.

    Query Parameters:
        group (str): Group filter ("all"/specific group). Defaults to "all".
        date (str): Month/date filter ("month"/"all"/YYYY-MM[/DD]). Defaults to "month".
        week (str): Optional week filter (for example "W1" or "Week 1").

    Returns:
        flask.Response: JSON object with:
            - overall (dict[str, float]): Overall FTA%, acceptance, rejection, and rework totals.
            - groupwise (dict[str, object]): Group labels and per-group quality measures.
    """
    data = load_data()
    group_f = request.args.get('group', 'all')
    date_f  = request.args.get('date',  'month')
    week_f  = request.args.get('week')

    fta_rows = data.get('fta', [])
    if group_f.lower() not in ('all', 'overall'):
        fta_rows = [r for r in fta_rows if r['group'].upper() == group_f.upper()]

    # Filter quality by month
    if date_f and date_f not in ('all', 'month'):
        target_month = date_f[:7] # extracts 'YYYY-MM' from 'YYYY-MM-DD' or keeps 'YYYY-MM'
        fta_rows = [r for r in fta_rows if r.get('month') == target_month]

    # Filter quality by week
    if week_f and week_f.lower() not in ('all', 'overall'):
        mapped_week = week_f
        if week_f.startswith('Week '):
            try:
                mapped_week = 'W' + week_f.split(' ')[1]
            except:
                pass
        fta_rows = [r for r in fta_rows if r.get('week') == mapped_week]

    group_acc     = defaultdict(float)
    group_rej     = defaultdict(float)
    group_he_rej  = defaultdict(float)
    group_he_acc  = defaultdict(float)
    group_fta_sum = defaultdict(list)

    for r in fta_rows:
        g = r['group']
        group_acc[g]    += r['acc']
        group_rej[g]    += r['rej']
        group_he_rej[g] += r['he_rej']
        group_he_acc[g] += r['he_acc']
        if r['fta_pct'] > 0:
            group_fta_sum[g].append(r['fta_pct'])

    all_groups = sorted(group_acc.keys())
    group_fta  = {g: round(sum(v)/len(v)*100, 1) if v else 0
                  for g, v in group_fta_sum.items()}

    total_acc   = sum(group_acc.values())
    total_rej   = sum(group_rej.values())
    total_rework= sum(group_he_rej.values())

    all_ftas = [r['fta_pct'] for r in fta_rows if r['fta_pct'] > 0]
    overall_fta = round(sum(all_ftas)/len(all_ftas)*100, 1) if all_ftas else 0

    return jsonify({
        'overall': {
            'fta_pct':      overall_fta,
            'total_acc':    total_acc,
            'total_rej':    total_rej,
            'total_rework': total_rework
        },
        'groupwise': {
            'groups':     all_groups,
            'fta_pct':    [group_fta.get(g, 0) for g in all_groups],
            'rejection':  [round(group_rej.get(g, 0)) for g in all_groups],
            'rework':     [round(group_he_rej.get(g, 0)) for g in all_groups],
            'acceptance': [round(group_acc.get(g, 0)) for g in all_groups]
        }
    })


@app.route('/api/inventory')
def api_inventory():
    """Return inventory, WIP flow, and material usage summaries.

    Query Parameters:
        date (str): "month"/"all", YYYY-MM, or YYYY-MM-DD. Defaults to "month".
        group (str): Group filter ("all"/specific group). Defaults to "all".
        model (str): Model filter. Defaults to "all".
        dates (str): Optional comma-separated YYYY-MM-DD list for explicit date filtering.

    Returns:
        flask.Response: JSON object with:
            - timeline (dict[str, list[int] | list[str]]): Daily input/output/WIP trend.
            - groupwise (list[dict[str, int | float | str]]): Per-group flow and material stats.
            - flow (dict[str, int | float]): Aggregated inventory flow totals.
            - materials (dict[str, object]): Material totals split by regular/add-on groups.
    """
    inv = load_inventory_data()
    if not inv:
        return jsonify({'error': 'Inventory data not loaded'}), 500

    date_f = request.args.get('date', 'month')
    group_f = request.args.get('group', 'all')
    model_f = request.args.get('model', 'all')
    
    dates_arg = request.args.get('dates')
    dates_list = dates_arg.split(',') if dates_arg else None

    # Filter dates
    if dates_list:
        target_dates = set(dates_list)
    elif date_f and date_f not in ('month', 'all'):
        if len(date_f) == 7: # YYYY-MM
            target_dates = {r['date'] for r in inv['input'] if r['date'].startswith(date_f)}
        else:
            target_dates = {date_f}
    else:
        target_dates = {r['date'] for r in inv['input']}

    def match(r):
        if r['date'] not in target_dates:
            return False
        if group_f.lower() not in ('all', 'overall'):
            if r['group'].upper() != group_f.upper():
                return False
        if model_f.lower() != 'all':
            if r.get('model', '').lower() != model_f.lower():
                return False
        return True

    def match_out(r):
        if r['date'] not in target_dates:
            return False
        if group_f.lower() not in ('all', 'overall'):
            if r['group'].upper() != group_f.upper():
                return False
        if model_f.lower() != 'all':
            if r.get('model', '').lower() != model_f.lower():
                return False
        return True

    filtered_inputs = [r for r in inv['input'] if match(r)]
    filtered_outputs = [r for r in inv['output'] if match_out(r)]
    
    def match_wip(r):
        if r['date'] not in target_dates:
            return False
        if group_f.lower() not in ('all', 'overall'):
            if r['group'].upper() != group_f.upper():
                return False
        return True
    filtered_wip = [r for r in inv['wip'] if match_wip(r)]

    filtered_strip = [r for r in inv['stripping_log'] if match_wip(r)]
    filtered_ret = [r for r in inv['returns_log'] if match_wip(r)]
    filtered_scrap = [r for r in inv['scrap_log'] if match_wip(r)]

    # Compute daily timeline
    timeline_inputs = defaultdict(float)
    timeline_outputs = defaultdict(float)
    timeline_wip = defaultdict(float)
    
    for r in filtered_inputs:
        timeline_inputs[r['date']] += r['qty']
    for r in filtered_outputs:
        timeline_outputs[r['date']] += r['qty']
    for r in filtered_wip:
        timeline_wip[r['date']] += r['qty']

    sorted_dates = sorted(list(target_dates))

    # Compute groupwise summary
    groups_list = ['A1', 'A2', 'A3', 'G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'KMC', 'PED', 'PED-AUTO']
    if group_f.lower() not in ('all', 'overall'):
        groups_list = [group_f.upper()]

    groupwise = {}
    for g in groups_list:
        g_in = sum(r['qty'] for r in filtered_inputs if r['group'] == g)
        g_out = sum(r['qty'] for r in filtered_outputs if r['group'] == g)
        
        g_wip_records = [r for r in filtered_wip if r['group'] == g]
        if g_wip_records:
            latest_wip_record = max(g_wip_records, key=lambda x: x['date'])
            g_wip = latest_wip_record['qty']
        else:
            g_wip = inv['summary'].get('avg_wip', {}).get(g, 0.0) if date_f in ('month', 'all') else 0.0
            
        g_adhe = sum(r['qty'] for r in filtered_inputs if r['group'] == g and r['mg_type'] == 'Adhe')
        g_gask = sum(r['qty'] for r in filtered_inputs if r['group'] == g and r['mg_type'] == 'Gask')
        g_sticker = sum(r['qty'] for r in filtered_inputs if r['group'] == g and r['mg_type'] == 'Sticker')

        g_strip = sum(r['qty'] for r in filtered_strip if r['group'] == g)
        g_ret = sum(r['qty'] for r in filtered_ret if r['group'] == g)
        g_scrap = sum(r['qty'] for r in filtered_scrap if r['group'] == g)

        if date_f in ('month', 'all') and model_f.lower() == 'all':
            g_open = inv['summary'].get('opening_wip', {}).get(g, 0.0)
            g_bal_input = inv['summary'].get('bal_input', {}).get(g, 0.0)
            g_avg = inv['summary'].get('avg_wip', {}).get(g, 0.0)
        else:
            first_wip_record = min(g_wip_records, key=lambda x: x['date']) if g_wip_records else None
            g_open = first_wip_record['qty'] if first_wip_record else 0.0
            g_bal_input = max(0.0, g_open + g_in - g_strip - g_ret - g_scrap - g_out)
            g_avg = sum(r['qty'] for r in g_wip_records) / len(g_wip_records) if g_wip_records else 0.0

        groupwise[g] = {
            'group': g,
            'opening_wip': round(g_open),
            'fresh_input': round(g_in),
            'total_output': round(g_out),
            'stripping': round(g_strip),
            'returns': round(g_ret),
            'scrap': round(g_scrap),
            'bal_input': round(g_bal_input),
            'wip_balance': round(g_wip),
            'avg_wip': round(g_avg, 1),
            'adhe': round(g_adhe),
            'gask': round(g_gask),
            'sticker': round(g_sticker)
        }

    # Aggregate Flow Totals
    flow = {
        'opening_wip': round(sum(gw['opening_wip'] for gw in groupwise.values())),
        'fresh_input': round(sum(gw['fresh_input'] for gw in groupwise.values())),
        'total_output': round(sum(gw['total_output'] for gw in groupwise.values())),
        'stripping': round(sum(gw['stripping'] for gw in groupwise.values())),
        'returns': round(sum(gw['returns'] for gw in groupwise.values())),
        'scrap': round(sum(gw['scrap'] for gw in groupwise.values())),
        'bal_input': round(sum(gw['bal_input'] for gw in groupwise.values())),
        'wip_balance': round(sum(gw['wip_balance'] for gw in groupwise.values())),
        'avg_wip': round(sum(gw['avg_wip'] for gw in groupwise.values()) / len(groupwise) if groupwise else 0.0, 1)
    }

    regular_groups = ['G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8']
    addon_groups = ['A1', 'A2', 'A3']

    materials = {
        'regular': {
            'adhe': sum(groupwise[g]['adhe'] for g in regular_groups if g in groupwise),
            'gask': sum(groupwise[g]['gask'] for g in regular_groups if g in groupwise)
        },
        'addon': {
            'adhe': sum(groupwise[g]['adhe'] for g in addon_groups if g in groupwise),
            'gask': sum(groupwise[g]['gask'] for g in addon_groups if g in groupwise)
        },
        'groups': [
            {
                'group': g,
                'category': 'regular' if g in regular_groups else 'addon' if g in addon_groups else 'other',
                'adhe': groupwise[g]['adhe'],
                'gask': groupwise[g]['gask']
            } for g in groups_list if g in groupwise
        ]
    }

    return jsonify({
        'timeline': {
            'dates': sorted_dates,
            'input': [round(timeline_inputs.get(d, 0)) for d in sorted_dates],
            'output': [round(timeline_outputs.get(d, 0)) for d in sorted_dates],
            'wip': [round(timeline_wip.get(d, 0)) for d in sorted_dates]
        },
        'groupwise': list(groupwise.values()),
        'flow': flow,
        'materials': materials
    })


@app.route('/api/refresh')
def refresh():
    """Clear in-memory caches and force reload of workbook-backed datasets.

    Query Parameters:
        None

    Returns:
        flask.Response: JSON object with refresh status and request-time ISO timestamp.
    """
    global _last_modified, _inventory_last_modified, _fta_detail_last_modified
    _last_modified = 0
    _inventory_last_modified = 0
    _fta_detail_last_modified = 0
    load_data()
    load_inventory_data()
    load_fta_detail()
    return jsonify({'status': 'refreshed', 'time': datetime.now().isoformat()})


if __name__ == '__main__':
    print("Loading initial data...")
    load_data()
    if os.path.exists(INVENTORY_EXCEL_PATH):
        load_inventory_data()
    if os.path.exists(FTA_EXCEL_PATH):
        load_fta_detail()
    watcher = threading.Thread(target=watch_file, daemon=True)
    watcher.start()
    print("Starting server on http://localhost:5050")
    app.run(host='0.0.0.0', port=5050, debug=False)
