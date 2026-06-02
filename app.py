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


# ─────────────────────────────────────────────
# File Watcher
# ─────────────────────────────────────────────

def watch_file():
    global _last_modified
    while True:
        try:
            mtime = os.path.getmtime(EXCEL_PATH)
            if mtime != _last_modified:
                print("Excel file changed – reloading...")
                load_data()
        except:
            pass
        time.sleep(3)


# ─────────────────────────────────────────────
# Filter helpers
# ─────────────────────────────────────────────

def filter_by(rows, date_f, model_f, group_f, shift_f=None,
              date_col='date', group_col='group', model_col='model', shift_col=None):
    result = []
    for r in rows:
        # Date
        if date_f and date_f not in ('month', 'all'):
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

    return jsonify({
        'dates':      all_dates,
        'months':     all_months,
        'groups':     all_groups,
        'models':     all_models[:300],
        'coverage': {
            'output_months': out_months,
            'input_months':  in_months
        },
        'last_updated': datetime.now().isoformat()
    })


@app.route('/api/production')
def production():
    data = load_data()
    date_f  = request.args.get('date',  'month')
    shift_f = request.args.get('shift', 'all')
    model_f = request.args.get('model', 'all')
    group_f = request.args.get('group', 'all')

    out_rows = filter_by(data.get('output', []), date_f, model_f, group_f,
                         date_col='date', group_col='group', model_col='model')
    in_rows  = filter_by(data.get('input', []),  date_f, model_f, group_f, shift_f,
                         date_col='date', group_col='group', model_col='model',
                         shift_col='shift')

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
            working_days = len(distinct_dates) if distinct_dates else 22
            flow_in = rate_day * working_days
            pct_till = round((flow_out / flow_in * 100) if flow_in > 0 else 0, 1)
            pct_month = round((flow_out / target_month * 100) if target_month > 0 else 0, 1)

            # Group targets
            group_targets = {c['group']: c['target_month'] for c in cap_f}
            group_till_date = {c['group']: c['target_till_date'] for c in cap_f}
            targets_list = [round(group_targets.get(g, 0)) for g in all_groups]
            till_date_list = [round(group_till_date.get(g, 0)) for g in all_groups]
    else:  # 'all' or 'month' (default overview)
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


@app.route('/api/quality')
def quality():
    data = load_data()
    group_f = request.args.get('group', 'all')
    date_f  = request.args.get('date',  'month')

    fta_rows = data.get('fta', [])
    if group_f.lower() not in ('all', 'overall'):
        fta_rows = [r for r in fta_rows if r['group'].upper() == group_f.upper()]

    # Filter quality by month
    if date_f and date_f not in ('all', 'month'):
        target_month = date_f[:7] # extracts 'YYYY-MM' from 'YYYY-MM-DD' or keeps 'YYYY-MM'
        fta_rows = [r for r in fta_rows if r.get('month') == target_month]

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


@app.route('/api/refresh')
def refresh():
    global _last_modified
    _last_modified = 0
    load_data()
    return jsonify({'status': 'refreshed', 'time': datetime.now().isoformat()})


if __name__ == '__main__':
    print("Loading initial data...")
    load_data()
    watcher = threading.Thread(target=watch_file, daemon=True)
    watcher.start()
    print("Starting server on http://localhost:5050")
    app.run(host='0.0.0.0', port=5050, debug=False)
