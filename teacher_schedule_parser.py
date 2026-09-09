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
from datetime import datetime, date
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
DATE_TO_RE = re.compile(r"(?<![А-Яа-яЁё])(?:до|по)\s*(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?", re.I)
DATE_FROM_RE = re.compile(r"(?<![А-Яа-яЁё])с\s*(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?", re.I)
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


SEMESTER_START = date(2026, 9, 1)
SEMESTER_END = date(2026, 12, 31)

def _parse_ddmm(s: str, default_year: int = 2026) -> date:
    d, m = [int(x) for x in s.split(".")[:2]]
    return date(default_year, m, d)

def extract_active_periods(text: str):
    s = norm(text)
    periods = []
    for a, b in DATE_RANGE_RE.findall(s):
        try:
            start, end = _parse_ddmm(a), _parse_ddmm(b)
            if end >= start:
                periods.append((start.isoformat(), end.isoformat()))
        except ValueError:
            pass
    for m in DATE_TO_RE.finditer(s):
        try:
            periods.append((SEMESTER_START.isoformat(), _parse_ddmm(m.group(1)).isoformat()))
        except ValueError:
            pass
    for m in DATE_FROM_RE.finditer(s):
        try:
            periods.append((_parse_ddmm(m.group(1)).isoformat(), SEMESTER_END.isoformat()))
        except ValueError:
            pass
    return list(dict.fromkeys(periods))

def room_extract(text: str) -> Optional[str]:
    s=norm(text)
    if not s: return None
    if re.search(r"\bдист(?:анционно)?\b", s, re.I):
        return "Дист"
    t=DATE_RANGE_RE.sub(" ", s)
    t=DATE_RE.sub(" ", t)
    t=re.sub(r"\b\d{1,2}[\.:]\d{2}\s*[-–—]\s*\d{1,2}[\.:]\d{2}\b", " ", t)
    # Частые варианты: 361ГТ, 8П, 610аМ, 403*Г, 119 Г, 343 / аГт.
    pairs=re.findall(
        r"\b(\d{1,4})\s*(?:[*/]\s*)?([А-ЯA-ZЁа-яё]{0,3})\b",
        t, re.I
    )
    candidates=[]
    for num, letters in pairs:
        # Не считаем номер пары/года отдельной аудиторией, если рядом нет
        # буквенного обозначения и число явно является частью текста.
        token=(num+letters).replace(" ","")
        candidates.append((num, letters, token))
    if not candidates: return None

    # Если после номера в исходной строке отдельно стоит обозначение Г/ГТ/П/М,
    # склеиваем его с номером даже при пробеле, '*' или '/'.
    m=list(re.finditer(r"\b(\d{1,4})\s*(?:[*/]\s*)?([А-ЯA-ZЁа-яё]{1,3})\b", t, re.I))
    if m:
        for mm in reversed(m):
            suffix=mm.group(2)
            if suffix.lower() == 'с':
                continue
            return mm.group(1)+suffix

    # Для записей вида "343 / аГт" или "370 / Г" достаём отдельное обозначение.
    m2=list(re.finditer(r"\b(\d{1,4})\s*/\s*([А-ЯA-ZЁа-яё]{1,3})\b", t, re.I))
    if m2:
        for mm in reversed(m2):
            if mm.group(2).lower() != 'с':
                return mm.group(1)+mm.group(2)

    # Числа без букв не считаем аудиторией: это часто номер пары,
    # номер проекта или дата, особенно в конструкциях "2 с 18.11".
    return None


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

def extract_entries(raw: str):
    raw=norm(raw)
    if not raw: return []
    teachers=[m for m in TEACHER_RE.finditer(raw) if is_teacher_match(m, raw)]
    if not teachers:
        return [{"subject": re.sub(r"\s+", " ", TITLE_RE.sub(" ", raw)).strip(" /-"), "teacher":None,
                 "room":room_extract(raw), "date_ranges":[f"{a}-{b}" for a,b in DATE_RANGE_RE.findall(raw)],
                 "active_periods":extract_active_periods(raw), "raw":raw}]
    out=[]
    for i,m in enumerate(teachers):
        prev_end=teachers[i-1].end() if i else 0
        next_start=teachers[i+1].start() if i+1<len(teachers) else len(raw)
        before=raw[prev_end:m.start()].strip(" /;,\n")
        if i and re.match(r"^\d{1,4}[А-ЯA-ZЁа-яё]{0,3}\b", before):
            before=re.sub(r"^\d{1,4}[А-ЯA-ZЁа-яё]{0,3}\b\s*", "", before, count=1).strip(" /;,\n")
        after=raw[m.end():next_start].strip(" /;,\n")
        # В части таблиц аудитория стоит ПЕРЕД преподавателем, например:
        # "РТПИДМ с ПСАП 119 Г / Преп. Лубина О.С.".
        # Сохраняем её как аудиторию и убираем из названия предмета.
        subject=TITLE_RE.sub(" ", before)
        room_before=room_extract(subject)
        subject=re.sub(r"\s+\d{1,4}(?:\s*[А-ЯA-ZЁа-яё]{1,3})?(?:\s*\.)?\s*$", "", subject)
        subject=re.sub(r"\s+", " ", subject).strip(" /-;.")
        room=room_extract(after)
        if room is None and room_before:
            room=room_before
        if room is None and i==len(teachers)-1:
            room=room_extract(after)
        local_text = before + " " + after
        out.append({"subject":subject or None,"teacher":teacher_normalize(m),"room":room,
                    "date_ranges":[f"{a}-{b}" for a,b in DATE_RANGE_RE.findall(local_text)],
                    "active_periods":extract_active_periods(before), "raw":raw})
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
        lesson=parse_lesson(get_cell(ws,amap,r,lesson_col))
        if day:
            current_day=day; current_date=date; current_lesson=None; lesson_start=None
            # В некоторых Excel-таблицах строка с названием дня одновременно
            # содержит номер первой пары и её предмет. Не теряем эту строку.
            if lesson is None:
                r+=1; continue
        if lesson is not None and (current_lesson!=lesson or lesson_start is None):
            current_lesson=lesson; lesson_start=r
        if not current_day or current_lesson is None or lesson_start is None:
            r+=1; continue
        end=r; rr=r+1
        while rr<=end_row:
            da=amap.get((rr,day_col)); d,_=parse_day(get_cell(ws,amap,rr,day_col)) if (da is None or da==(rr,day_col)) else (None,None)
            l=parse_lesson(get_cell(ws,amap,rr,lesson_col))
            if d or (l is not None and l!=current_lesson): break
            end=rr; rr+=1
        if r!=lesson_start:
            r+=1; continue
        for spec in specs:
            rows=[]
            for rr in range(lesson_start,end+1):
                vals=row_fragments(ws,amap,spans,spec,rr,protected_cols=(day_col,lesson_col)); raw=" / ".join(vals)
                # В некоторых таблицах аудитория вынесена в соседнюю колонку
                # сразу справа от блока группы (например, E -> F).
                # Добавляем её, если в текущем фрагменте преподаватель есть,
                # а аудитория ещё не распознана.
                if raw and has_teacher(raw):
                    existing_room=room_extract(raw)
                    if existing_room is None:
                        for extra_c in range(spec['c2']+1, min(ws.max_column, spec['c2']+2)+1):
                            extra_anchor=amap.get((rr,extra_c))
                            if extra_anchor and extra_anchor!=(rr,extra_c):
                                continue
                            extra=norm(ws.cell(rr,extra_c).value)
                            if extra and room_extract(extra):
                                raw=raw+" / "+extra
                                break
                if raw and not TIME_ONLY_RE.fullmatch(raw): rows.append((rr,raw))
            if not rows: continue
            unique=[]
            for item in rows:
                if not any(item[1]==x[1] for x in unique): unique.append(item)
            occ=[]; i=0
            while i<len(unique):
                rr0,raw0=unique[i]
                if i+1<len(unique) and not has_teacher(raw0) and has_teacher(unique[i+1][1]):
                    occ.append((rr0,norm(raw0+" / "+unique[i+1][1]))); i+=2
                else:
                    occ.append((rr0,raw0)); i+=1
            date_based=bool(current_date) or bool(re.search(r"\b\d{1,2}\.\d{1,2}\b",norm(get_cell(ws,amap,lesson_start,day_col))))
            if date_based:
                assigns=[(rr0,raw,"date") for rr0,raw in occ]
            elif len(occ)==1:
                rr0,raw0=occ[0]; both=False
                for c in range(spec['c1'],spec['c2']+1):
                    a=amap.get((rr0,c))
                    if a and spans.get(a,(1,1))[0]>=2: both=True; break
                assigns=[(rr0,raw0,"odd"),(rr0,raw0,"even")] if both else [(rr0,raw0,"odd")]
            else:
                assigns=[(rr0,raw,"odd" if idx%2==0 else "even") for idx,(rr0,raw) in enumerate(occ)]
            for rr0,raw,week_type in assigns:
                for e in extract_entries(raw):
                    if e['teacher'] and not e['subject']:
                        # Excel may place the subject in the row immediately above
                        # the teacher while the teacher occupies a vertically merged cell.
                        for back in range(rr0-1, max(lesson_start-1, rr0-3), -1):
                            prev_vals=row_fragments(ws,amap,spans,spec,back,protected_cols=(day_col,lesson_col))
                            prev_raw=" / ".join(prev_vals)
                            if prev_raw and not has_teacher(prev_raw) and has_subjectish(prev_raw):
                                combined=prev_raw+" / "+raw
                                ce=extract_entries(combined)
                                if ce and ce[0]['teacher']:
                                    e=ce[0]
                                break
                    # Если Excel разнесла предмет и преподавателя по соседним строкам
                    # (например, "Математика" -> "Лубина О.С. 216Г"),
                    # пробуем восстановить название предмета из предыдущей строки.
                    if e['teacher'] and not e['subject']:
                        raw_idx = next((j for j,(rraw,rrraw) in enumerate(occ) if rrraw == raw), None)
                        if raw_idx is not None and raw_idx > 0:
                            prev_raw = occ[raw_idx-1][1]
                            if not has_teacher(prev_raw) and has_subjectish(prev_raw):
                                combined = prev_raw + " / " + raw
                                candidates = extract_entries(combined)
                                if candidates and candidates[0]['teacher']:
                                    e = candidates[0]
                    if not e['teacher']: continue
                    fallback_ranges=e['date_ranges'] or [f"{a}-{b}" for a,b in spec['header_ranges']]
                    active_periods=e.get('active_periods') or []
                    if not active_periods and fallback_ranges:
                        for a,b in DATE_RANGE_RE.findall(" / ".join(fallback_ranges)):
                            try:
                                active_periods.append((_parse_ddmm(a).isoformat(), _parse_ddmm(b).isoformat()))
                            except ValueError:
                                pass
                    records.append({"institute":institute,"source_file":source_file,"sheet":ws.title,
                      "group":spec['group'],"day":current_day,"date_text":current_date,"lesson":current_lesson,
                      "week_type":week_type,"subject":e['subject'],"teacher":e['teacher'],"teacher_search":norm_search(e['teacher']),
                      "room":e['room'],"date_ranges":fallback_ranges,"active_periods":active_periods,"raw":e['raw']})
        r=end+1
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
    root=Path(root)
    records=[]
    files=sorted(root.rglob('*.xlsx'))
    for p in files:
        institute=p.parent.parent.name if p.parent.name in {'Бакалавриат','Магистратура'} else p.parent.name
        if institute in {'teacher_schedules',''}: institute=p.parent.name
        rec=parse_workbook(p,institute)
        print(f"{institute}: {p.name} -> {len(rec)} records")
        records.extend(rec)
    # Exact de-duplication.
    unique=[]; seen=set()
    for r in records:
        key=(r['institute'],r['source_file'],r['sheet'],r['group'],r['day'],r['date_text'],r['lesson'],r['week_type'],r['subject'],r['teacher'],r['room'],tuple(r['date_ranges']),tuple(r.get('active_periods',[])))
        if key not in seen:
            seen.add(key); unique.append(r)
    return unique


def create_db(records, db_path):
    db_path=Path(db_path); db_path.parent.mkdir(parents=True,exist_ok=True)
    con=sqlite3.connect(db_path); cur=con.cursor()
    cur.execute('DROP TABLE IF EXISTS lessons')
    cur.execute('''CREATE TABLE lessons(
        id INTEGER PRIMARY KEY,
        institute TEXT NOT NULL, source_file TEXT, sheet TEXT,
        group_name TEXT NOT NULL, day TEXT NOT NULL, date_text TEXT,
        lesson INTEGER NOT NULL, week_type TEXT NOT NULL,
        subject TEXT, teacher TEXT NOT NULL, teacher_search TEXT NOT NULL,
        room TEXT, date_ranges TEXT, active_periods TEXT, raw TEXT)''')
    rows=[(r['institute'],r['source_file'],r['sheet'],r['group'],r['day'],r['date_text'],r['lesson'],r['week_type'],r['subject'],r['teacher'],r['teacher_search'],r['room'],json.dumps(r['date_ranges'],ensure_ascii=False),json.dumps(r.get('active_periods',[]),ensure_ascii=False),r['raw']) for r in records]
    cur.executemany('''INSERT INTO lessons(institute,source_file,sheet,group_name,day,date_text,lesson,week_type,subject,teacher,teacher_search,room,date_ranges,active_periods,raw)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',rows)
    cur.execute('CREATE INDEX idx_teacher ON lessons(teacher_search)')
    cur.execute('CREATE INDEX idx_teacher_day ON lessons(teacher_search,day,lesson)')
    cur.execute('CREATE INDEX idx_group ON lessons(group_name,day,lesson)')
    con.commit(); con.close()
    return len(records)


def search_teacher(query, day=None, week_type=None, as_of_date=None, db_path='schedule.db'):
    q=norm_search(query)
    if not q: return []
    con=sqlite3.connect(db_path); con.row_factory=sqlite3.Row; cur=con.cursor()
    sql='''SELECT institute,group_name,day,date_text,lesson,week_type,subject,teacher,room,date_ranges,active_periods FROM lessons WHERE teacher_search LIKE ?'''
    params=[f'%{q}%']
    if day:
        sql+=' AND day=?'; params.append(day)
    if week_type:
        sql+=' AND (week_type=? OR week_type="date")'; params.append(week_type)
    if as_of_date:
        sql += " AND (active_periods IS NULL OR active_periods='[]' OR EXISTS (SELECT 1 FROM json_each(lessons.active_periods) WHERE json_extract(json_each.value, '$[0]') <= ? AND json_extract(json_each.value, '$[1]') >= ?))"
        params.extend([as_of_date, as_of_date])
    sql+=''' ORDER BY CASE day WHEN 'Понедельник' THEN 1 WHEN 'Вторник' THEN 2 WHEN 'Среда' THEN 3 WHEN 'Четверг' THEN 4 WHEN 'Пятница' THEN 5 WHEN 'Суббота' THEN 6 WHEN 'Воскресенье' THEN 7 ELSE 99 END, lesson, institute, group_name'''
    out=[dict(x) for x in cur.execute(sql,params).fetchall()]
    con.close()
    for x in out:
        try:x['date_ranges']=json.loads(x['date_ranges'])
        except: x['date_ranges']=[]
        try:x['active_periods']=json.loads(x['active_periods'])
        except: x['active_periods']=[]
    return out


def list_teachers(query, db_path='schedule.db'):
    q=norm_search(query)
    con=sqlite3.connect(db_path); cur=con.cursor()
    rows=cur.execute('SELECT teacher FROM lessons WHERE teacher_search LIKE ? GROUP BY teacher ORDER BY teacher LIMIT 30',(f'%{q}%',)).fetchall()
    con.close(); return [r[0] for r in rows]


if __name__=='__main__':
    ap=argparse.ArgumentParser()
    ap.add_argument('root', nargs='?', default='teacher_schedules')
    ap.add_argument('--db', default='schedule.db')
    args=ap.parse_args()
    rec=build_records(args.root)
    n=create_db(rec,args.db)
    print(f'Готово: {n} уникальных записей -> {args.db}')
    print('Преподавателей:', len({r["teacher"] for r in rec}))
    print('Групп:', len({r["group"] for r in rec}))
