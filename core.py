"""Garmin Connect -> classeur coach, semaine choisie, modification XML ciblée."""
import copy
import datetime as dt
import getpass
import io
import json
import math
import os
from pathlib import Path
import re
import statistics
import tempfile
import zipfile
from collections import defaultdict
from xml.etree import ElementTree as ET

BASE = Path(__file__).resolve().parent
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS)
CELL = re.compile(r'^([A-Z]+)(\d+)$')

def colnum(s):
    n = 0
    for x in s: n = n * 26 + ord(x) - 64
    return n

def colname(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def ref_offset(ref, dx=0, dy=0):
    m = CELL.fullmatch(ref)
    if not m: raise ValueError(ref)
    return f'{colname(colnum(m[1]) + dx)}{int(m[2]) + dy}'

def number(x):
    try:
        f = float(x)
        return f if math.isfinite(f) else None
    except (ValueError, TypeError): return None

def fmt_hms(seconds):
    s = max(0, round(seconds))
    return f'{s//3600:02d}:{s//60%60:02d}:{s%60:02d}'

def time_excel(seconds): return seconds / 86400.0

def get(d, *keys):
    for k in keys:
        if isinstance(d, dict) and d.get(k) is not None: return d[k]
    return None

def weighted(activities, field):
    vals = [(number(a.get(field)), number(a.get('duration_s'))) for a in activities]
    vals = [(v, w) for v, w in vals if v is not None and w and w > 0]
    return sum(v*w for v,w in vals)/sum(w for _,w in vals) if vals else None

def sum_known(activities, field):
    vals = [number(a.get(field)) for a in activities]
    vals = [x for x in vals if x is not None]
    return sum(vals) if vals else None

def fit_fields(msg):
    return {x.name: x.value for x in msg.fields if x.value is not None}

def parse_fit(blob):
    from fitparse import FitFile
    f = FitFile(io.BytesIO(blob))
    laps, session, records = [], {}, []
    for m in f.get_messages():
        if m.name == 'lap': laps.append(fit_fields(m))
        elif m.name == 'session': session = fit_fields(m)
        elif m.name == 'record': records.append(fit_fields(m))
    return laps, session, records

def lap_line(lap, sport, idx):
    secs = number(get(lap,'total_timer_time','total_elapsed_time'))
    if secs is None: return None
    hr = number(get(lap,'avg_heart_rate'))
    up = number(get(lap,'total_ascent'))
    down = number(get(lap,'total_descent'))
    hr_s = f'{round(hr)} bpm' if hr is not None else '—'
    up_s = f'{round(up)} m' if up is not None else '—'
    down_s = f'{round(down)} m' if down is not None else '—'
    cadence = number(get(lap,'avg_cadence','avg_running_cadence'))
    if sport == 'bike':
        pw = number(get(lap,'avg_power'))
        np = number(get(lap,'normalized_power'))
        return (f'Lap {idx} : {fmt_hms(secs)} | {hr_s} | '
                f'{round(pw)} W' if pw is not None else f'Lap {idx} : {fmt_hms(secs)} | {hr_s} | —') + (
                f' | NP {round(np)} W' if np is not None else ' | NP —') + (
                f' | {round(cadence)} rpm' if cadence is not None else ' | — rpm') + f' | D+ {up_s} | D- {down_s}'
    dist = number(get(lap,'total_distance'))
    speed = dist/secs*3.6 if dist is not None and secs else None
    pace = f'{int(round(secs/(dist/1000)))//60}:{int(round(secs/(dist/1000)))%60:02d}/km' if dist and dist>0 else '—'
    return (f'Lap {idx} : {fmt_hms(secs)} | {hr_s} | {pace} | '
            f'{speed:.2f} km/h' if speed is not None else f'Lap {idx} : {fmt_hms(secs)} | {hr_s} | {pace} | —') + (
            f' | {round(cadence)} ppm' if cadence is not None else ' | — ppm') + f' | D+ {up_s} | D- {down_s}'

def classify(activity):
    t = get(activity.get('activityType') or {},'typeKey') or activity.get('activityType') or ''
    t = str(t).lower()
    if 'swim' in t: return 'swim'
    if any(x in t for x in ('cycling','biking','virtual_ride','indoor_cycling')): return 'bike'
    if any(x in t for x in ('running','treadmill','trail_run','track_run')): return 'run'
    return 'other'

def zone_seconds(api, activity_id):
    """Garmin's activity HR time-in-zones endpoint; missing zones remain missing."""
    for method in ('get_activity_hr_in_timezones','get_activity_hr_in_time_zones'):
        fn = getattr(api, method, None)
        if fn:
            try:
                raw = fn(activity_id)
                if isinstance(raw, dict): raw = raw.get('zones') or raw.get('hrZones') or raw.get('timeInZones') or []
                out = {}
                for x in raw or []:
                    zone = number(get(x,'zoneNumber','zone','zoneIndex'))
                    sec = number(get(x,'secsInZone','seconds','timeInZone','duration'))
                    if zone is not None and sec is not None and 2 <= zone <= 5: out[int(zone)] = sec
                if out: return out
            except Exception as e: print(f'  Zones Garmin non disponibles ({activity_id}): {e}')
    return {}

def download_fit(api, activity_id):
    fmt = getattr(getattr(api, 'ActivityDownloadFormat', None), 'ORIGINAL', None)
    try:
        blob = api.download_activity(activity_id, dl_fmt=fmt) if fmt is not None else api.download_activity(activity_id)
        if isinstance(blob, str): blob = Path(blob).read_bytes()
        if blob[:2] == b'PK':
            with zipfile.ZipFile(io.BytesIO(blob)) as z:
                names = [n for n in z.namelist() if n.lower().endswith('.fit')]
                if not names: raise ValueError('ZIP sans fichier FIT')
                return z.read(names[0])
        return blob
    except Exception as e:
        print(f'  FIT indisponible pour {activity_id}: {e}')
        return None

def activity_info(api, raw):
    aid = raw['activityId']
    kind = classify(raw)
    duration = number(get(raw,'movingDuration','duration'))
    a = {'id':aid,'kind':kind,'date':str(get(raw,'startTimeLocal','startTimeGMT'))[:10],
         'duration_s':duration,'distance_km':(number(raw.get('distance')) or 0)/1000,
         'ascent':number(get(raw,'elevationGain')),'hr_avg':number(get(raw,'averageHR')),
         'hr_max':number(get(raw,'maxHR')),'power_avg':number(get(raw,'avgPower')),
         'np':number(get(raw,'normPower')),'cadence':number(get(raw,'averageBikingCadence','averageRunningCadenceInStepsPerMinute','averageRunningCadence')),
         'speed_kmh':(number(get(raw,'averageSpeed')) or 0)*3.6 if get(raw,'averageSpeed') is not None else None,
         'zones':zone_seconds(api,aid),'laps':[]}
    if kind in ('bike','run'):
        blob = download_fit(api, aid)
        if blob:
            try:
                laps, session, records = parse_fit(blob)
                a['laps'] = laps
                if a['np'] is None: a['np'] = number(session.get('normalized_power'))
                if a['cadence'] is None: a['cadence'] = number(get(session,'avg_cadence','avg_running_cadence'))
                if a['ascent'] is None: a['ascent'] = number(session.get('total_ascent'))
                if a['hr_max'] is None: a['hr_max'] = number(session.get('max_heart_rate'))
            except Exception as e: print(f'  Erreur lecture FIT {aid}: {e}')
    return a

def health_day(api, day):
    """Try common Garmin Connect health endpoints without fabricating values."""
    d = day.isoformat()
    result = {}
    try:
        stats = api.get_stats(d)
        if isinstance(stats, dict): result['resting'] = number(get(stats,'restingHeartRate','restingHR'))
    except Exception as e: print(f'  FC repos {d}: {e}')
    try:
        hrv = api.get_hrv_data(d)
        summary = hrv.get('hrvSummary',{}) if isinstance(hrv,dict) else {}
        result['hrv'] = number(get(summary,'lastNightAvg','weeklyAvg'))
    except Exception as e: print(f'  VFC {d}: {e}')
    return result

def build_updates(activities, health, config):
    week = int(config['week_number'])
    anchor = config['week_1_anchor']
    dx = (week-1)*int(config['week_column_stride'])
    base = colnum('AVN') + dx
    monday = dt.date.fromisoformat(config['week_1_monday']) + dt.timedelta(weeks=week-1)
    updates = {}
    warnings = []
    def put(col,row,val, kind='number'):
        if val is None: return
        if kind == 'number' and not math.isfinite(float(val)): return
        updates[f'{colname(col)}{row}'] = (val,kind)
    for day in range(7):
        date = monday + dt.timedelta(days=day)
        row = 24 + 8*day
        d = date.isoformat()
        h = health.get(d,{})
        put(colnum('AVI')+dx,26+8*day,h.get('resting'))
        put(colnum('AVI')+dx,28+8*day,h.get('hrv'))
        for kind, offset in (('bike',0),('run',4),('other',8)):
            arr = [a for a in activities if a['date']==d and a['kind']==kind]
            if not arr: continue
            c = base+offset
            for zone in (2,3,4,5):
                vals = [a['zones'].get(zone) for a in arr]
                if all(v is not None for v in vals):
                    # Excel template stores zone times as MINUTES (not Excel clock values)
                    put(c+zone-2,row,sum(vals)/60)
                else: warnings.append(f'{d} {kind}: zone FC {zone} absente de Garmin, non remplie')
            if kind != 'other':
                put(c,row+6,sum_known(arr,'ascent'))
                if kind == 'bike':
                    put(c+1,row+6,arr[0].get('np') if len(arr)==1 else None)
                    put(c+2,row+6,weighted(arr,'hr_avg'))
                    put(c+3,row+6,weighted(arr,'cadence'))
                else:
                    put(c+1,row+6,weighted(arr,'hr_avg'))
                    put(c+2,row+6,weighted(arr,'speed_kmh'))
                    put(c+3,row+6,weighted(arr,'cadence'))
                put(c,row+7,sum_known(arr,'distance_km'))
            duration = sum_known(arr,'duration_s')
            if duration is not None:
                # Time cell is merged horizontally over two columns; write top-left only
                put(c+1,row+7,time_excel(duration))
            if kind in ('bike','run'):
                lines=[]
                for a in arr:
                    if len(arr)>1: lines.append(f"Activité Garmin {a['id']} :")
                    for i, lap in enumerate(a['laps'],1):
                        line=lap_line(lap,kind,i)
                        if line: lines.append(line)
                    if a['hr_max'] is not None: lines.append(f"FC max : {round(a['hr_max'])} bpm")
                if lines: put(c,row+2,'\n'.join(lines),'text')
            else:
                # FC max in other sport comment cell, if available (not mapped yet)
                if any(a['hr_max'] is not None for a in arr):
                    warnings.append(f'{d} autre sport: FC max non écrite (cellule commentaire TRI non spécifiée)')
    return updates,warnings

def patch_xlsx(src,dest,updates,overwrite=False):
    """Patch only selected XML <c> elements. Other ZIP members copied byte-for-byte."""
    if src.resolve()==dest.resolve(): raise ValueError('Refus de modifier le fichier original')
    with zipfile.ZipFile(src,'r') as zin:
        sheet_path='xl/worksheets/sheet1.xml'
        root=ET.fromstring(zin.read(sheet_path))
        sd=root.find(f'{{{NS}}}sheetData')
        existing={c.get('r'):c for row in sd.findall(f'{{{NS}}}row') for c in row.findall(f'{{{NS}}}c')}
        changed=[]; skipped=[]
        for ref,(value,typ) in updates.items():
            c=existing.get(ref)
            if c is None:
                skipped.append((ref,'cellule absente : écriture refusée'))
                continue
            if c.find(f'{{{NS}}}f') is not None:
                skipped.append((ref,'formule protégée'))
                continue
            old_v=c.find(f'{{{NS}}}v')
            old_inline=c.find(f'{{{NS}}}is')
            if not overwrite and ((old_v is not None and old_v.text not in (None,'')) or old_inline is not None):
                skipped.append((ref,'déjà renseignée'))
                continue
            for child in list(c):
                if child.tag in (f'{{{NS}}}v',f'{{{NS}}}is'):
                    c.remove(child)
            if typ=='text':
                c.set('t','inlineStr')
                is_node=ET.SubElement(c,f'{{{NS}}}is')
                t=ET.SubElement(is_node,f'{{{NS}}}t')
                t.text=str(value)
            else:
                c.attrib.pop('t',None)
                ET.SubElement(c,f'{{{NS}}}v').text=str(value)
            changed.append(ref)
        if not changed: return changed,skipped
        xml=ET.tostring(root,encoding='utf-8',xml_declaration=True)
        with zipfile.ZipFile(dest,'w') as zout:
            for info in zin.infolist():
                zout.writestr(info, xml if info.filename==sheet_path else zin.read(info.filename))
    return changed,skipped

def main():
    config=json.loads((BASE/'config.json').read_text(encoding='utf-8'))
    src=BASE/config['excel_file']
    if not src.is_file(): raise FileNotFoundError(f'Fichier Excel introuvable : {src}')
    monday=dt.date.fromisoformat(config['week_1_monday'])+dt.timedelta(weeks=int(config['week_number'])-1)
    sunday=monday+dt.timedelta(days=6)
    print(f'Semaine {config["week_number"]} : {monday} -> {sunday}')
    from garminconnect import Garmin
    email=input('Email Garmin Connect : ').strip()
    password=getpass.getpass('Mot de passe Garmin (non sauvegardé) : ')
    api=Garmin(email,password)
    api.login()
    del password
    raw=api.get_activities_by_date(monday.isoformat(),sunday.isoformat())
    print(f'{len(raw)} activité(s) Garmin trouvée(s)')
    activities=[]
    for a in raw:
        if classify(a)=='swim': continue
        print(f"  {a.get('activityName','Activité')} ({a.get('activityId')})")
        activities.append(activity_info(api,a))
    health={}
    if config.get('resting_hr_and_hrv',True):
        for day in range(7):
            d=monday+dt.timedelta(days=day)
            health[d.isoformat()]=health_day(api,d)
    updates,warnings=build_updates(activities,health,config)
    preview={k:{'value':v,'type':t} for k,(v,t) in updates.items()}
    (BASE/'apercu.json').write_text(json.dumps({'cells':preview,'warnings':warnings},indent=2,ensure_ascii=False),encoding='utf-8')
    print(f'\n{len(updates)} cellules candidates ; aperçu dans apercu.json')
    for k,(v,t) in sorted(updates.items()): print(f'  {k} = {str(v)[:95]}')
    for w in sorted(set(warnings)): print('ATTENTION :',w)
    if config.get('dry_run',True):
        print('\nMODE APERÇU : aucun fichier Excel modifié. Passer dry_run=false pour écrire.')
        return
    out=src.with_name(f'{src.stem}_rempli_S{config["week_number"]:02d}.xlsx')
    changed,skipped=patch_xlsx(src,out,updates,config.get('overwrite_existing',False))
    print(f'\n{len(changed)} cellules écrites ; {len(skipped)} ignorées')
    for r,why in skipped: print(f'  IGNORÉ {r}: {why}')
    print(f'Fichier : {out if changed else "aucun changement"}')

if __name__=='__main__':
    try: main()
    except Exception as e:
        print(f'ERREUR : {type(e).__name__}: {e}')
        raise
