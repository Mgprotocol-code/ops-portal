"""Convert daily Excel uploads (uploads/*.xlsx) into portal data files.
Also deletes report days on request (uploads/delete-YYYY-MM-DD.txt).

Keeps the newest 7 report days in data/, deletes older ones, and records the
result of the last upload in data/last_upload.json (shown on the portal).
Run by .github/workflows/portal.yml; can also be run by hand:  python scripts/convert.py
"""
import json, os, re, glob, shutil, datetime, math
from openpyxl import load_workbook

KEEP_DAYS = 7
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
UPLOADS = os.path.join(ROOT, 'uploads')
FAILED = os.path.join(UPLOADS, 'failed')

INFO = [("Report date", "report_date", "d"), ("Fleet time", "fleet_time", "t"), ("City time", "city_time", "t"),
        ("FTTx time", "fttx_time", "t"), ("Box-Ops time", "box_time", "t")]
# sheet, dataset id, [(header, field, kind)]  kind: t text, n number, d date, a auto (ignored),
# l:<a,b> strict list, L: suggested list (any text)
SPEC = [
 ("Fleet by team", "fleet_by_team", [("Priority","priority","L:"),("Team","team","t"),("Request","request","n"),("Plan","plan","n"),("Actual","actual","n"),("Actual %","","a")]),
 ("Office cars", "office_cars", [("OU","ou","t"),("Day","day","n"),("Night","night","n"),("Total","","a")]),
 ("City cars", "city_cars", [("City","city","t"),("Request","request","n"),("Actual","actual","n"),("Actual %","","a")]),
 ("Fleet notes", "fleet_notes", [("Item","metric","t"),("Cars","value","n")]),
 ("FTTx blog", "fttx_blog", [("Region","region","t"),("Team","team","n"),("Car","car","n"),("Broken / high-loss","broken_highloss","n"),("Migration / activities","migration","n"),("Remain / pending","pending","n")]),
 ("Cable works", "cable_works", [("Stage","stage","l:Done,Plan"),("Date","date","d"),("Region","region","t"),("Township","tsp","t"),("Type","work_type","L:"),("Street and road","street","t"),("HC","hc","n"),("Labour","labour","n"),("Impact OLT","impact_olt","n"),("Impact CA1","impact_ca1","n"),("Impact CA2/OTB","impact_ca2_otb","n")]),
 ("Pending sites", "pending_sites", [("Reason","reason","t"),("Case","case","t"),("Sites","sites","n")]),
 ("Box-Ops", "box_ops", [("Region","region","t"),("Township","tsp","t"),("Ticket blog","ticket_blog","n"),("Plan tkts","plan_tkts","n"),("POI blog","poi_blog","n"),("Plan POI","plan_poi","n"),("SI team","si_team","n"),("Car","car","n"),("Labour","labour","n")]),
 ("Remarks", "remarks", [("Area","area","l:Fleet,FTTx,FTTx pending"),("Topic","topic","t"),("Date","date","d"),("Sites","sites","n"),("Remark","remark","t")]),
]
MON = {m: i for i, m in enumerate(['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'], 1)}


def txt(v):
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v).strip()


def to_date(v):
    if v is None or txt(v) == '': return None
    if isinstance(v, datetime.datetime): return v.date().isoformat()
    if isinstance(v, datetime.date): return v.isoformat()
    if isinstance(v, (int, float)): return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(round(v)))).isoformat()
    s = txt(v)
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m: return datetime.date(int(m[1]), int(m[2]), int(m[3])).isoformat()
    m = re.match(r'^(\d{1,2})[-\s/]([A-Za-z]{3})[A-Za-z]*[-\s/,]+(\d{4})$', s)
    if m and m[2].lower() in MON: return datetime.date(int(m[3]), MON[m[2].lower()], int(m[1])).isoformat()
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', s)
    if m: return datetime.date(int(m[3]), int(m[2]), int(m[1])).isoformat()
    return False


def to_time(v):
    if isinstance(v, (datetime.time, datetime.datetime)):
        h, mm = v.hour, v.minute
        return f"{h % 12 or 12}:{mm:02d} {'PM' if h >= 12 else 'AM'}"
    if isinstance(v, float) and 0 <= v < 1:
        mins = round(v * 1440); h = mins // 60
        return f"{h % 12 or 12}:{mins % 60:02d} {'PM' if h >= 12 else 'AM'}"
    return txt(v)


def find_sheet(wb, name):
    for ws in wb.worksheets:
        if ws.title.strip().lower() == name.lower(): return ws
    return None


def num(v):
    f = float(txt(v).replace(',', ''))
    return int(f) if f.is_integer() else f


def parse(path):
    wb = load_workbook(path, data_only=True)
    errors, warnings, info, raw = [], [], {}, {}
    ri = find_sheet(wb, 'Report info')
    if ri is None: errors.append('Sheet “Report info” is missing.')
    else:
        for row in ri.iter_rows(values_only=True):
            if not row: continue
            lab = txt(row[0]).lower(); val = row[1] if len(row) > 1 else None
            for l, key, k in INFO:
                if l.lower() == lab: info[key] = to_date(val) if k == 'd' else to_time(val)
        if not info.get('report_date'): errors.append('Report info: “Report date” is empty or not a date.')
    for sheet, did, cols in SPEC:
        ws = find_sheet(wb, sheet)
        if ws is None: errors.append(f'Sheet “{sheet}” is missing.'); continue
        rows = list(ws.iter_rows(values_only=True))
        hr = next((i for i, r in enumerate(rows) if any(txt(c).lower() == cols[0][0].lower() for c in r)), None)
        if hr is None: errors.append(f'{sheet}: header row not found (row 3 must start with “{cols[0][0]}”).'); continue
        pos = {}
        for i, c in enumerate(rows[hr]):
            for h, f, k in cols:
                if k != 'a' and txt(c).lower() == h.lower(): pos[f] = i
        missing = [h for h, f, k in cols if k != 'a' and f not in pos]
        if missing: errors.append(f'{sheet}: column(s) missing: {", ".join(missing)}.'); continue
        out = []
        for rn, row in enumerate(rows[hr + 1:], start=hr + 2):
            rec, anyv = {}, False
            for h, f, k in cols:
                if k == 'a': continue
                v = row[pos[f]] if pos[f] < len(row) else None
                if txt(v) != '': anyv = True
                if k == 'n':
                    if txt(v) == '': rec[f] = None
                    else:
                        try: rec[f] = num(v)
                        except ValueError: errors.append(f'{sheet} row {rn}: “{h}” must be a number (found “{txt(v)}”).'); rec[f] = None
                elif k == 'd':
                    dd = to_date(v)
                    if dd is False: errors.append(f'{sheet} row {rn}: “{h}” is not a date (found “{txt(v)}”).'); dd = None
                    rec[f] = dd
                else: rec[f] = txt(v)
            if not anyv: continue
            first = txt(row[pos[cols[0][1]]]).lower()
            if re.match(r'^(total|grand total|all cities)', first) or first.endswith('total'):
                warnings.append(f'{sheet} row {rn}: total row skipped – the portal adds totals itself.'); continue
            for h, f, k in cols:
                if k.startswith('l:') and rec[f] != '' and rec[f] not in k[2:].split(','):
                    errors.append(f'{sheet} row {rn}: “{h}” must be one of {" / ".join(k[2:].split(","))} (found “{rec[f]}”).')
            if not rec[cols[0][1]]: errors.append(f'{sheet} row {rn}: “{cols[0][0]}” is empty.')
            out.append((rn, rec))
        if not out: warnings.append(f'{sheet}: no rows filled.')
        raw[did] = out
    if errors: return info, None, errors, warnings

    N = lambda v: 0 if v is None else v
    pct = lambda a, b: (a / b) if b else None
    r1 = lambda v: None if v is None else math.floor(v * 10 + 0.5) / 10

    def dup(rows, keyf, sheet, label):
        seen = set()
        for r, x in rows:
            k = keyf(x)
            if k in seen: errors.append(f'{sheet} row {r}: {label} “{k}” appears twice.')
            seen.add(k)
    d = info['report_date']
    info['prev_date'] = (datetime.date.fromisoformat(d) - datetime.timedelta(days=1)).isoformat()
    R = {}
    dup(raw['fleet_by_team'], lambda x: x['team'], 'Fleet by team', 'Team')
    pr, R['fleet_by_team'] = '', []
    for _, x in raw['fleet_by_team']:
        pr = x['priority'] or pr
        R['fleet_by_team'].append(dict(priority=pr, team=x['team'], request=N(x['request']), plan=N(x['plan']), actual=N(x['actual']), actual_pct=pct(N(x['actual']), N(x['request']))))
    dup(raw['office_cars'], lambda x: x['ou'], 'Office cars', 'OU')
    R['office_cars'] = [dict(ou=x['ou'], day=N(x['day']), night=N(x['night']), total=N(x['day']) + N(x['night'])) for _, x in raw['office_cars']]
    dup(raw['city_cars'], lambda x: x['city'], 'City cars', 'City')
    R['city_cars'] = [dict(city=x['city'], request=N(x['request']), actual=N(x['actual']), actual_pct=pct(N(x['actual']), N(x['request']))) for _, x in raw['city_cars']]
    cr, ca = sum(x['request'] for x in R['city_cars']), sum(x['actual'] for x in R['city_cars'])
    R['city_cars'].append(dict(city='All Cities', request=cr, actual=ca, actual_pct=pct(ca, cr)))
    dup(raw['fleet_notes'], lambda x: x['metric'], 'Fleet notes', 'Item')
    R['fleet_notes'] = [dict(metric=x['metric'], value=N(x['value'])) for _, x in raw['fleet_notes']]
    if not any(x['metric'] == 'Driver leave / site cancel' for x in R['fleet_notes']):
        warnings.append('Fleet notes: item “Driver leave / site cancel” not found – keep the item names as in the template.')
    dup(raw['fttx_blog'], lambda x: x['region'], 'FTTx blog', 'Region')
    R['fttx_blog'] = [dict(region=x['region'], team=N(x['team']), car=N(x['car']), broken_highloss=N(x['broken_highloss']), migration=N(x['migration']), pending=N(x['pending'])) for _, x in raw['fttx_blog']]
    cnt = {'Done': 0, 'Plan': 0}; R['cable_works'] = []
    for _, x in raw['cable_works']:
        cnt[x['stage']] += 1
        R['cable_works'].append(dict(job=('D' if x['stage'] == 'Done' else 'P') + str(cnt[x['stage']]), stage=x['stage'],
            date=x['date'] or (info['prev_date'] if x['stage'] == 'Done' else d), region=x['region'], tsp=x['tsp'], work_type=x['work_type'],
            street=x['street'], hc=N(x['hc']), labour=N(x['labour']), impact_olt=N(x['impact_olt']), impact_ca1=N(x['impact_ca1']), impact_ca2_otb=N(x['impact_ca2_otb'])))
    dup(raw['pending_sites'], lambda x: x['reason'], 'Pending sites', 'Reason')
    R['pending_sites'] = [dict(reason=x['reason'], case=x['case'], sites=N(x['sites'])) for _, x in raw['pending_sites']]
    dup(raw['box_ops'], lambda x: x['region'] + ' / ' + x['tsp'], 'Box-Ops', 'Region / township')
    R['box_ops'] = [dict(region=x['region'], tsp=x['tsp'], ticket_blog=N(x['ticket_blog']), plan_tkts=N(x['plan_tkts']), plan_pct_tkt=pct(N(x['plan_tkts']), N(x['ticket_blog'])),
        poi_blog=N(x['poi_blog']), plan_poi=N(x['plan_poi']), plan_pct_poi=pct(N(x['plan_poi']), N(x['poi_blog'])), si_team=N(x['si_team']), car=N(x['car']), labour=N(x['labour']),
        car_per_team=r1(pct(N(x['si_team']), N(x['car']))), team_per_labour=r1(pct(N(x['labour']), N(x['si_team']))), is_total=False) for _, x in raw['box_ops']]
    R['remarks'] = [dict(id=f'R{i + 1:02d}', area=x['area'], topic=x['topic'], date=x['date'], sites=x['sites'], remark=x['remark']) for i, (_, x) in enumerate(raw['remarks'])]
    R['report_info'] = [dict(report_date=d, prev_date=info['prev_date'], fleet_time=info.get('fleet_time', ''), city_time=info.get('city_time', ''), fttx_time=info.get('fttx_time', ''), box_time=info.get('box_time', ''))]
    if errors: return info, None, errors, warnings
    return info, R, errors, warnings


def main():
    os.makedirs(DATA, exist_ok=True)
    files = [f for f in sorted(glob.glob(os.path.join(UPLOADS, '*.xlsx')), key=os.path.getmtime) if not os.path.basename(f).startswith('~$')]
    # delete requests: uploads/delete-YYYY-MM-DD.txt (made by the portal's "Delete this day" button)
    dels = [f for f in sorted(glob.glob(os.path.join(UPLOADS, '*.txt'))) if os.path.basename(f).lower().startswith('delete')]
    if not files and not dels:
        print('No new uploads.'); return
    now = datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds')
    days_path = os.path.join(DATA, 'days.json')
    days = json.load(open(days_path)) if os.path.exists(days_path) else []
    status = None
    for f in dels:
        name = os.path.basename(f)
        try: body = open(f, encoding='utf-8', errors='ignore').read()
        except Exception: body = ''
        wanted = sorted(set(re.findall(r'\d{4}-\d{2}-\d{2}', name + ' ' + body)))
        gone = [d for d in wanted if d in days]
        for d in gone:
            p = os.path.join(DATA, d + '.json')
            if os.path.exists(p): os.remove(p)
        days = [d for d in days if d not in gone]
        os.remove(f)
        status = dict(status='deleted', file=name, at=now, deleted=gone, not_found=[d for d in wanted if d not in gone])
        print('DELETE', name, '->', gone or 'nothing', '| not found:', status['not_found'] or 'none')
    for f in files:
        name = os.path.basename(f)
        try:
            info, R, errors, warnings = parse(f)
        except Exception as e:  # unreadable file
            info, R, errors, warnings = {}, None, [f'Could not read the file ({e}). Save it as Excel Workbook (.xlsx) and upload again.'], []
        d = info.get('report_date')
        if R is not None:
            all_days = sorted(set(days) | {d}, reverse=True)
            if d not in all_days[:KEEP_DAYS]:
                R, errors = None, [f'Report date {d} is older than the {KEEP_DAYS} days the portal keeps. Check the date on the Report info sheet.']
        if R is None:
            os.makedirs(FAILED, exist_ok=True)
            shutil.move(f, os.path.join(FAILED, name))
            status = dict(status='error', file=name, at=now, report_date=d or None, errors=errors[:50], warnings=warnings)
            print('FAILED', name, *errors, sep='\n  ')
            continue
        all_days = sorted(set(days) | {d}, reverse=True)
        keep, drop = all_days[:KEEP_DAYS], all_days[KEEP_DAYS:]
        replaced = d in days
        with open(os.path.join(DATA, d + '.json'), 'w', encoding='utf-8') as fh:
            json.dump(dict(report_date=d, uploaded_at=now, file=name, datasets=R), fh, ensure_ascii=False)
        for old in drop:
            p = os.path.join(DATA, old + '.json')
            if os.path.exists(p): os.remove(p)
        days = keep
        os.remove(f)
        status = dict(status='ok', file=name, at=now, report_date=d, replaced=replaced, removed=drop, warnings=warnings)
        print('OK', name, '->', d, '| removed:', drop or 'none')
    with open(days_path, 'w') as fh: json.dump(sorted(days, reverse=True), fh)
    if status:
        with open(os.path.join(DATA, 'last_upload.json'), 'w', encoding='utf-8') as fh: json.dump(status, fh, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
