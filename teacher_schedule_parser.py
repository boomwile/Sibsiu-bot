# -*- coding: utf-8 -*-
"""Unified parser for SibGIU teacher search schedules.

Reads all .xlsx files under a directory, detects several common SibGIU
schedule layouts, normalizes lessons to SQLite, and supports teacher search.
"""
from __future__ import annotations

import argparse
import collections
import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

DAY_NAMES = {
    "ПОНЕДЕЛЬНИК": "Понедельник", "ПН": "Понедельник",
    "ВТОРНИК": "Вторник", "ВТ": "Вторник",
    "СРЕДА": "Среда", "СР": "Среда",
    "ЧЕТВЕРГ": "Четверг", "ЧТ": "Четверг",
    "ПЯТНИЦА": "Пятница", "ПТ": "Пятница",
    "СУББОТА": "Суббота", "СБ": "Суббота",
    "ВОСКРЕСЕНЬЕ": "Воскресенье", "ВС": "Воскресенье",
}
ROMAN = {"I":1,"II":2,"III":3,"IV":4,"V":5,"VI":6,"VII":7,"VIII":8,"IX":9,"X":10,"XI":11,"XII":12}
GROUP_RE = re.compile(r"(?<![\wА-Яа-яЁё])([А-ЯA-ZЁёа-яё*][А-ЯA-ZЁёа-яё0-9*_-]{1,14}-\d{2,4})(?!\w)")
TEACHER_RE = re.compile(r"(?P<surname>[А-ЯЁ][а-яё-]{2,})\s+(?P<i1>[А-ЯЁ])\.?\s*(?P<i2>[А-ЯЁ])\.?", re.U)
DATE_RE = re.compile(r"(?<!\d)(\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)(?!\d)")
DATE_RANGE_RE = re.compile(r"(?<!\d)(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?\s*-\s*(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?(?!\d)")
TIME_ONLY_RE = re.compile(r"^\d{1,2}[\.:]\d{2}\s*[-–—]\s*\d{1,2}[\.:]\d{2}$")
TITLE_RE = re.compile(r"\b(?:доц\.?|проф\.?|преп\.?|асс\.?|ст\.?\s*п\.?|ст\.п\.?)\b", re.I)


def norm(v) -> str:
    if v is None:
        return ""
    s = str(v).replace("\xa0", " ").replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    return s.strip()


def norm_search(v: str) -> str:
    return re.sub(r"\s+", " ", norm(v).lower().replace("ё", "е").replace(".", ""))


def teacher_normalize(m: re.Match) -> str:
    return f"{m.group('surname')} {m.group('i1')}.{m.group('i2')}."


def parse_lesson(v: str) -> Optional[int]:
    s = norm(v).upper().replace("—", "-").replace("–", "-")
    if not s:
        return None
    s2 = s.split()[0].strip(".,;:")
    if s2 in ROMAN:
        return ROMAN[s2]
    m = re.match(r"^(\d{1,2})(?:\s|/|$)", s)
    return int(m.group(1)) if m else None


def parse_day(v: str) -> tuple[Optional[str], Optional[str]]:
    s = norm(v).upper().replace("Ё", "Е")
    if not s:
        return None, None
    for key, day in DAY_NAMES.items():
        if s == key or s.startswith(key + " ") or s.startswith(key + "."):
            dates = DATE_RE.findall(s)
            return day, (dates[0] if dates else None)
    # e.g. "ВТ 06.10", "сб 07.11"
    m = re.match(r"^(ПН|ВТ|СР|ЧТ|ПТ|СБ|ВС)\.?\s*(\d{1,2}\.\d{1,2})?", s)
    if m:
        return DAY_NAMES[m.group(1)], m.group(2)
    return None, None


def merged_map(ws):
    amap = {}
    spans = {}
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(rng))
        anchor = (min_row, min_col)
        spans[anchor] = (max_row-min_row+1, max_col-min_col+1)
        for r in range(min_row, max_row+1):
            for c in range(min_col, max_col+1):
                amap[(r,c)] = anchor
    return amap, spans


def get_cell(ws, amap, r, c) -> str:
    anchor = amap.get((r,c))
    if anchor:
        return norm(ws.cell(anchor[0], anchor[1]).value)
    return norm(ws.cell(r,c).value)


def span_for(ws, amap, spans, r, c):
    a = amap.get((r,c))
    return (*a, *spans[a]) if a else (r,c,1,1)


def group_specs(ws, amap, spans):
    """Find group header row(s) and column spans."""
    candidates=[]
    for r in range(1, min(ws.max_row, 12)+1):
        for c in range(1, ws.max_column+1):
            v=get_cell(ws,amap,r,c)
            groups=GROUP_RE.findall(v)
            for g in groups:
                ar,ac,rs,cs=span_for(ws,amap,spans,r,c)
                candidates.append((r,g,ac,ac+cs-1,v))
    if not candidates:
        # Some sheets have the group label above a date-based table and no regex match due to formatting.
        return []
    # Pick the earliest row with the largest number of unique groups.
    counts={r:len({x[1] for x in candidates if x[0]==r}) for r,*_ in candidates}
    header_row=max(counts, key=lambda r:(counts[r],-r))
    specs=[]
    for r,g,c1,c2,v in candidates:
        if r!=header_row: continue
        specs.append({"group":g,"c1":c1,"c2":c2,"header_ranges":DATE_RANGE_RE.findall(v)})
    # De-dup and resolve accidental overlap.
    seen=set(); out=[]
    for s in sorted(specs,key=lambda x:x['c1']):
        if s['group'] in seen: continue
        seen.add(s['group']); out.append(s)
    return header_row,out


def row_fragments(ws, amap, spans, spec, r, protected_cols=()):
    vals=[]; seen_anchors=set()
    for c in range(spec['c1'], spec['c2']+1):
        anchor=amap.get((r,c))
        if anchor:
            if anchor in seen_anchors:
                continue
            arow,acol=anchor
            rs,cs=spans.get(anchor,(1,1))
            if acol < spec['c1'] and acol+cs-1 < spec['c1']:
                continue
            if acol in protected_cols:
                continue
            seen_anchors.add(anchor)
        v=get_cell(ws,amap,r,c)
        if v and v not in vals:
            vals.append(v)
    return vals


def room_extract(text: str) -> Optional[str]:
    s=norm(text)
    if not s: return None
    if re.search(r"\bдист(?:анционно)?\b", s, re.I):
        return "Дист"
    # Remove dates/times before room detection.
    t=DATE_RANGE_RE.sub(" ", s)
    t=DATE_RE.sub(" ", t)
    t=re.sub(r"\b\d{1,2}[\.:]\d{2}\s*[-–—]\s*\d{1,2}[\.:]\d{2}\b", " ", t)
    # Common room formats: 361ГТ, 8П, 610аМ, 515, 513 / 212, 368 Г.
    matches=list(re.finditer(r"\b(\d{1,4})(?:\s*([А-ЯA-ZЁа-яё]{1,3}))?\b", t))
    if not matches: return None
    # Ignore obvious academic years / stray dates.
    good=[]
    for m in matches:
        n=int(m.group(1))
        if n < 1 or n > 9999: continue
        token=m.group(0).strip()
        # Skip likely ordinal lesson numbers surrounded by subject text if no teacher afterwards is handled elsewhere.
        good.append(token.replace(" ",""))
    if not good: return None
    # A bare 1..9 is normally a stray course/lesson number, not an auditorium.
    if len(good[-1]) <= 1 and good[-1].isdigit(): return None
    # preserve slash-separated room pairs when present
    pairs=re.findall(r"\b\d{1,4}(?:\s*[А-ЯA-ZЁа-яё]{0,3})?\s*/\s*\d{1,4}(?:\s*[А-ЯA-ZЁа-яё]{0,3})?\b", t)
    if pairs:
        return pairs[-1].replace(" ", "")
    return good[-1]


COMMON_SUBJECT_WORDS = {
    "основы", "психология", "адаптивная", "менеджмент", "физиология",
    "биохимия", "педагогическое", "математика", "физика", "химия",
    "информатика", "история", "литература", "география", "программирование",
    "экономика", "правоведение", "социология", "философия", "культура",
    "теория", "технология", "безопасность", "инженерная", "иностранный",
    "русский", "рисунок", "электротехника", "автоматизация", "физическая",
}

def is_teacher_match(match: re.Match, raw: str) -> bool:
    surname = match.group('surname').lower().replace('ё','е')
    prefix = raw[max(0, match.start()-12):match.start()]
    has_title = bool(re.search(r"(?:доц\.?|проф\.?|преп\.?|асс\.?|ст\.?\s*п\.?|ст\.п\.?)\s*$", prefix, re.I))
    if surname in COMMON_SUBJECT_WORDS and not has_title:
        return False
    return True

DATE_QUALIFIER_RE = re.compile(r'(?i)(?:(?:до|с)\s*\d{1,2}\.\d{1,2}(?:\.\d{2,4})?|\d{1,2}\.\d{1,2}\s*[-–—]\s*\d{1,2}\.\d{1,2})')

def clean_subject(text: str) -> Optional[str]:
    s=norm(text)
    if not s: return None
    s=TITLE_RE.sub(' ',s)
    s=DATE_QUALIFIER_RE.sub(' ',s)
    # Remove trailing room token, but keep numbers that are part of subject names.
    s=re.sub(r'\s+\d{1,4}(?:\s*[А-ЯA-ZЁа-яё]{1,3})?\s*$', '', s)
    s=re.sub(r'\s+',' ',s).strip(' /-;,')
    return s or None

def extract_periods(text: str):
    s=norm(text); out=[]
    for m in re.finditer(r'(?i)(?:до|с)\s*\d{1,2}\.\d{1,2}(?:\.\d{2,4})?|\d{1,2}\.\d{1,2}\s*[-–—]\s*\d{1,2}\.\d{1,2}', s):
        v=m.group(0).replace(' ','')
        nums=re.findall(r'\d+',v)
        if len(nums)>=2 and all(1<=int(nums[i])<=31 for i in range(0,len(nums),2)) and all(1<=int(nums[i])<=12 for i in range(1,len(nums),2)):
            out.append(v)
    return out

def extract_entries(raw: str):
    raw=norm(raw)
    if not raw: return []
    teachers=[m for m in TEACHER_RE.finditer(raw) if is_teacher_match(m, raw)]
    if not teachers:
        return [{'subject':clean_subject(raw),'teacher':None,'room':room_extract(raw),
                 'date_ranges':[f'{a}-{b}' for a,b in DATE_RANGE_RE.findall(raw)],'lesson_periods':extract_periods(raw),'raw':raw}]

    # If there is one room token in the whole cell, it is normally shared by
    # all teachers in that lesson (common SibGIU multi-teacher layout).
    room_all=room_extract(raw)
    room_tokens=list(re.finditer(r'(?<!\d)\d{1,4}(?:\s*[А-ЯA-ZЁа-яё]{1,3})?(?!\d)', DATE_RANGE_RE.sub(' ',raw)))
    room_candidates=[]
    for m in room_tokens:
        token=m.group(0).strip(); suffix=re.sub(r'^\d+', '', token).replace(' ','').lower()
        if suffix in {'ст','стп','преп','доц','проф','асс'}: continue
        if re.search(r'[А-ЯA-ZЁа-яё]',token) or len(re.match(r'\d+',token).group(0))>=3: room_candidates.append(token)
    shared_room=room_all if len(room_candidates)==1 else None

    # The subject is usually everything before the first teacher. This is
    # especially important when several teachers are listed on following rows.
    base_subject=clean_subject(raw[:teachers[0].start()])
    base_dates=[f'{a}-{b}' for a,b in DATE_RANGE_RE.findall(raw)]
    out=[]
    for i,m in enumerate(teachers):
        next_start=teachers[i+1].start() if i+1<len(teachers) else len(raw)
        before=raw[teachers[i-1].end() if i else 0:m.start()].strip(' /;,\n')
        after=raw[m.end():next_start].strip(' /;,\n')
        local_subject=clean_subject(before) if i==0 else None
        subject=local_subject or base_subject
        # A room after this teacher belongs to it. Otherwise use a single shared room.
        room=room_extract(after)
        if room is None and shared_room: room=shared_room
        period_text=before + (' '+after if i==len(teachers)-1 else '')
        dates=[f'{a}-{b}' for a,b in DATE_RANGE_RE.findall(period_text)] or base_dates
        periods=extract_periods(period_text) or (extract_periods(raw) if i==len(teachers)-1 else [])
        out.append({'subject':subject,'teacher':teacher_normalize(m),'room':room,'date_ranges':dates,'lesson_periods':periods,'raw':raw})
    return out

def has_teacher(text): return bool(TEACHER_RE.search(norm(text)))

def has_subjectish(text):
    s=norm(text)
    if not s: return False
    if has_teacher(s): return True
    if TIME_ONLY_RE.fullmatch(s): return False
    return bool(re.search(r"[А-Яа-яЁёA-Za-z]{3,}", s))


def detect_day_lesson_columns(ws, amap, start_row=1, end_row=None):
    end_row=end_row or ws.max_row
    day_col=None
    for r in range(start_row, min(end_row,start_row+25)+1):
        for c in range(1,min(ws.max_column,5)+1):
            a=amap.get((r,c))
            if a and a!=(r,c): continue
            day,_=parse_day(get_cell(ws,amap,r,c))
            if day:
                day_col=c; break
        if day_col: break
    return (1,2) if day_col is None else (day_col,min(day_col+1,ws.max_column))


def find_sections(ws, amap, spans):
    rows=collections.defaultdict(list)
    for r in range(1,ws.max_row+1):
        for c in range(1,ws.max_column+1):
            a=amap.get((r,c))
            if a and a!=(r,c): continue
            v=norm(ws.cell(r,c).value)
            groups=GROUP_RE.findall(v)
            if not groups: continue
            ar,ac,rs,cs=span_for(ws,amap,spans,r,c)
            if ar!=r: continue
            for g in groups:
                rows[r].append({"group":g,"c1":ac,"c2":ac+cs-1,"header_ranges":DATE_RANGE_RE.findall(v)})
    sections=[]
    for r in sorted(rows):
        seen=set(); specs=[]
        for spec in sorted(rows[r],key=lambda x:x['c1']):
            if spec['group'] in seen: continue
            seen.add(spec['group']); specs.append(spec)
        if specs: sections.append((r,specs))
    good=[]
    for i,(r,specs) in enumerate(sections):
        nxt=sections[i+1][0]-1 if i+1<len(sections) else ws.max_row
        found=False
        for rr in range(r+1,min(nxt,r+25)+1):
            for c in range(1,min(ws.max_column,5)+1):
                a=amap.get((rr,c))
                if a and a!=(rr,c): continue
                if parse_day(get_cell(ws,amap,rr,c))[0]: found=True; break
            if found: break
        if found: good.append((r,specs))
    return good


def parse_section(ws,amap,spans,institute,source_file,header_row,specs,end_row):
    day_col,lesson_col=detect_day_lesson_columns(ws,amap,header_row+1,end_row)
    records=[]; current_day=None; current_date=None; current_lesson=None; lesson_start=None
    r=header_row+1
    while r<=end_row:
        a=amap.get((r,day_col)); day_is_anchor=(a is None or a==(r,day_col))
        day,date=parse_day(get_cell(ws,amap,r,day_col)) if day_is_anchor else (None,None)
        if day:
            current_day=day; current_date=date
            current_lesson=parse_lesson(get_cell(ws,amap,r,lesson_col))
            lesson_start=r if current_lesson is not None else None
            r+=1; continue
        # Only the anchor row of a vertically merged lesson cell should start a new lesson.
        la=amap.get((r,lesson_col))
        lesson_anchor=(la is None or la==(r,lesson_col))
        lesson=parse_lesson(get_cell(ws,amap,r,lesson_col)) if lesson_anchor else None
        if lesson is not None and (current_lesson!=lesson or lesson_start is None): current_lesson=lesson; lesson_start=r
        if not current_day or current_lesson is None or lesson_start is None: r+=1; continue
        end=r; rr=r+1
        while rr<=end_row:
            da=amap.get((rr,day_col)); d,_=parse_day(get_cell(ws,amap,rr,day_col)) if (da is None or da==(rr,day_col)) else (None,None)
            l=parse_lesson(get_cell(ws,amap,rr,lesson_col))
            if d or (l is not None and l!=current_lesson): break
            end=rr; rr+=1
        for spec in specs:
            rows=[]
            for rr in range(lesson_start,end+1):
                vals=row_fragments(ws,amap,spans,spec,rr,protected_cols=(day_col,lesson_col)); raw=' / '.join(vals)
                if raw and not TIME_ONLY_RE.fullmatch(raw): rows.append((rr,raw))
            if not rows: continue
            unique=[]
            for item in rows:
                if not any(item[1]==x[1] for x in unique): unique.append(item)
            occ=[]; i=0
            while i<len(unique):
                rr0,raw0=unique[i]
                if i+1<len(unique) and not has_teacher(raw0) and has_teacher(unique[i+1][1]):
                    occ.append((rr0,norm(raw0+' / '+unique[i+1][1]))); i+=2
                else: occ.append((rr0,raw0)); i+=1
            date_based=bool(current_date) or bool(re.search(r"\b\d{1,2}\.\d{1,2}\b",norm(get_cell(ws,amap,lesson_start,day_col))))
            if date_based: assigns=[(raw,'date') for _,raw in occ]
            elif len(occ)==1:
                rr0,raw0=occ[0]; both=False
                for c in range(spec['c1'],spec['c2']+1):
                    a=amap.get((rr0,c))
                    if a and spans.get(a,(1,1))[0]>=2: both=True; break
                assigns=[(raw0,'odd'),(raw0,'even')] if both else [(raw0,'odd')]
            else: assigns=[(raw,'odd' if idx%2==0 else 'even') for idx,(_,raw) in enumerate(occ)]
            for raw,week_type in assigns:
                for e in extract_entries(raw):
                    records.append({'institute':institute,'source_file':source_file,'sheet':ws.title,
                        'group':spec['group'],'day':current_day,'date_text':current_date,'lesson':current_lesson,
                        'week_type':week_type,'subject':e.get('subject'),'teacher':e.get('teacher'),
                        'teacher_search':norm_search(e['teacher']) if e.get('teacher') else '',
                        'room':e.get('room'),'lesson_periods':e.get('lesson_periods') or [],'date_ranges':e.get('date_ranges') or [f"{a}-{b}" for a,b in spec['header_ranges']],
                        'raw':e.get('raw',raw)})
        r=end+1
        lesson_start=None
    return records

def parse_sheet(ws,institute,source_file):
    amap,spans=merged_map(ws); sections=find_sections(ws,amap,spans)
    if not sections: return []
    out=[]
    for i,(header_row,specs) in enumerate(sections):
        end=sections[i+1][0]-1 if i+1<len(sections) else ws.max_row
        out.extend(parse_section(ws,amap,spans,institute,source_file,header_row,specs,end))
    return out


def parse_workbook(path: Path, institute: str):
    wb=load_workbook(path,data_only=True,read_only=False)
    out=[]
    for ws in wb.worksheets:
        if ws.max_row<=1 and ws.max_column<=1: continue
        try: out.extend(parse_sheet(ws,institute,path.name))
        except Exception as e: print(f"WARN {path.name} [{ws.title}]: {e}")
    return out


def build_records(root: str|Path):
    root=Path(root); records=[]; files=sorted(root.rglob('*.xlsx')); seen_files=set()
    for p in files:
        import hashlib
        h=hashlib.sha256(p.read_bytes()).hexdigest()
        if h in seen_files: continue
        seen_files.add(h)
        parts=p.parts; inst='UNKNOWN'
        for part in parts:
            if part.startswith('расписание_'): inst=part[len('расписание_'):]; break
        if inst=='UNKNOWN' and p.parent.name not in {'source_unique',''}: inst=p.parent.name
        rec=parse_workbook(p,inst)
        print(f'{inst}: {p.name} -> {len(rec)} records')
        records.extend(rec)
    unique=[]; seen=set()
    for r in records:
        key=(r['institute'],r['source_file'],r['sheet'],r['group'],r['day'],r['date_text'],r['lesson'],r['week_type'],r.get('subject'),r.get('teacher'),r.get('room'),tuple(r.get('date_ranges',[])))
        if key not in seen: seen.add(key); unique.append(r)
    return unique

def create_db(records, db_path):
    db_path=Path(db_path); db_path.parent.mkdir(parents=True,exist_ok=True)
    con=sqlite3.connect(db_path); cur=con.cursor()
    cur.execute('DROP TABLE IF EXISTS lessons')
    cur.execute("""CREATE TABLE lessons(
        id INTEGER PRIMARY KEY,
        institute TEXT NOT NULL, source_file TEXT, sheet TEXT,
        group_name TEXT NOT NULL, day TEXT NOT NULL, date_text TEXT,
        lesson INTEGER NOT NULL, week_type TEXT NOT NULL,
        subject TEXT, teacher TEXT, teacher_search TEXT NOT NULL,
        room TEXT, lesson_periods TEXT, date_ranges TEXT, raw TEXT)""")
    rows=[(r['institute'],r['source_file'],r['sheet'],r['group'],r['day'],r['date_text'],r['lesson'],r['week_type'],r.get('subject'),r.get('teacher'),r.get('teacher_search',''),r.get('room'),json.dumps(r.get('lesson_periods',[]),ensure_ascii=False),json.dumps(r.get('date_ranges',[]),ensure_ascii=False),r.get('raw','')) for r in records]
    cur.executemany("""INSERT INTO lessons(institute,source_file,sheet,group_name,day,date_text,lesson,week_type,subject,teacher,teacher_search,room,lesson_periods,date_ranges,raw)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",rows)
    cur.execute('CREATE INDEX idx_teacher ON lessons(teacher_search)')
    cur.execute('CREATE INDEX idx_teacher_day ON lessons(teacher_search,day,lesson)')
    cur.execute('CREATE INDEX idx_group ON lessons(group_name,day,lesson)')
    cur.execute('CREATE INDEX idx_group_week ON lessons(group_name,day,week_type,lesson)')
    con.commit(); con.close(); return len(records)

def _decode_rows(rows):
    out=[dict(x) for x in rows]
    for x in out:
        try: x['lesson_periods']=json.loads(x.get('lesson_periods') or '[]')
        except Exception: x['lesson_periods']=[]
        try: x['date_ranges']=json.loads(x.get('date_ranges') or '[]')
        except Exception: x['date_ranges']=[]
    return out

def period_active(periods, on_date=None):
    if not periods: return True
    d=on_date or datetime.now().date()
    for p in periods:
        m=re.fullmatch(r'до(\d{1,2})\.(\d{1,2})',p.replace(' ',''),re.I)
        if m and 1<=int(m.group(1))<=31 and 1<=int(m.group(2))<=12:
            if d <= datetime(d.year,int(m.group(2)),int(m.group(1))).date(): return True
            continue
        m=re.fullmatch(r'с(\d{1,2})\.(\d{1,2})',p.replace(' ',''),re.I)
        if m and 1<=int(m.group(1))<=31 and 1<=int(m.group(2))<=12:
            if d >= datetime(d.year,int(m.group(2)),int(m.group(1))).date(): return True
            continue
        m=re.fullmatch(r'(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})',p.replace(' ',''))
        if m:
            a=datetime(d.year,int(m.group(2)),int(m.group(1))).date(); b=datetime(d.year,int(m.group(4)),int(m.group(3))).date()
            if a <= d <= b: return True
    return False

def search_teacher(query, day=None, week_type=None, db_path='schedule.db'):
    q=norm_search(query)
    if not q: return []
    con=sqlite3.connect(db_path); con.row_factory=sqlite3.Row; cur=con.cursor()
    sql='''SELECT institute,group_name,day,date_text,lesson,week_type,subject,teacher,room,lesson_periods,date_ranges
           FROM lessons WHERE teacher_search LIKE ? AND teacher IS NOT NULL'''
    params=[f'%{q}%']
    if day: sql+=' AND day=?'; params.append(day)
    if week_type: sql+=' AND (week_type=? OR week_type="date")'; params.append(week_type)
    sql+=''' ORDER BY CASE day WHEN 'Понедельник' THEN 1 WHEN 'Вторник' THEN 2 WHEN 'Среда' THEN 3 WHEN 'Четверг' THEN 4 WHEN 'Пятница' THEN 5 WHEN 'Суббота' THEN 6 WHEN 'Воскресенье' THEN 7 ELSE 99 END, lesson, institute, group_name'''
    out=_decode_rows(cur.execute(sql,params).fetchall()); con.close(); return [x for x in out if period_active(x.get('lesson_periods',[]))]

def list_teachers(query='', db_path='schedule.db'):
    q=norm_search(query); con=sqlite3.connect(db_path); cur=con.cursor()
    if q: rows=cur.execute('SELECT teacher FROM lessons WHERE teacher IS NOT NULL AND teacher_search LIKE ? GROUP BY teacher ORDER BY teacher LIMIT 30',(f'%{q}%',)).fetchall()
    else: rows=cur.execute('SELECT teacher FROM lessons WHERE teacher IS NOT NULL GROUP BY teacher ORDER BY teacher').fetchall()
    con.close(); return [r[0] for r in rows]

def list_groups(db_path='schedule.db'):
    con=sqlite3.connect(db_path); cur=con.cursor(); rows=cur.execute('SELECT DISTINCT group_name FROM lessons ORDER BY group_name').fetchall(); con.close(); return [r[0] for r in rows]

def search_group(group, day=None, week_type=None, db_path='schedule.db'):
    group=norm(group)
    if not group: return []
    con=sqlite3.connect(db_path); con.row_factory=sqlite3.Row; cur=con.cursor()
    sql='''SELECT institute,group_name,day,date_text,lesson,week_type,subject,teacher,room,lesson_periods,date_ranges,raw
           FROM lessons WHERE group_name=?'''; params=[group]
    if day: sql+=' AND day=?'; params.append(day)
    if week_type: sql+=' AND (week_type=? OR week_type="date")'; params.append(week_type)
    sql+=''' ORDER BY CASE day WHEN 'Понедельник' THEN 1 WHEN 'Вторник' THEN 2 WHEN 'Среда' THEN 3 WHEN 'Четверг' THEN 4 WHEN 'Пятница' THEN 5 WHEN 'Суббота' THEN 6 ELSE 99 END, lesson, id'''
    out=_decode_rows(cur.execute(sql,params).fetchall()); con.close(); return [x for x in out if period_active(x.get('lesson_periods',[]))]

def get_group_schedule(group, day=None, week_type=None, db_path='schedule.db'):
    rows=search_group(group,day,week_type,db_path); merged={}
    for r in rows:
        subject=r.get('subject') or ''
        key=(r['group_name'],r['day'],r['lesson'],r['week_type'],r.get('date_text') or '',tuple(r.get('date_ranges') or []),subject,r.get('room') or '')
        if key not in merged: merged[key]={**r,'teachers':[],'_rooms':[]}
        t=r.get('teacher')
        if t and t not in merged[key]['teachers']: merged[key]['teachers'].append(t)
        room=r.get('room')
        if room and room not in merged[key]['_rooms']: merged[key]['_rooms'].append(room)
        if not merged[key].get('room') and room: merged[key]['room']=room
    out=[]
    for r in merged.values():
        r.pop('_rooms',None)
        out.append(r)
    out.sort(key=lambda x:(x['lesson'],x.get('subject') or '',x.get('room') or '')); return out

if __name__=='__main__':
    ap=argparse.ArgumentParser(); ap.add_argument('root',nargs='?',default='teacher_schedules'); ap.add_argument('--db',default='schedule.db'); args=ap.parse_args()
    rec=build_records(args.root); n=create_db(rec,args.db)
    print(f'Готово: {n} уникальных записей -> {args.db}')
    print('Преподавателей:',len({r['teacher'] for r in rec if r.get('teacher')})); print('Групп:',len({r['group'] for r in rec}))
