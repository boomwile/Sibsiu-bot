# -*- coding: utf-8 -*-
"""SibGIU unified Excel schedule parser.

One source of truth for both:
  * student group schedule
  * teacher schedule search

The parser is intentionally conservative about rooms: tokens such as ``2с``
are NOT rooms. Project numbers in subjects (e.g. ``Проектная деятельность 1``)
are preserved.
"""
from __future__ import annotations
import argparse, collections, json, re, sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

DAY_NAMES={"ПОНЕДЕЛЬНИК":"Понедельник","ПН":"Понедельник","ВТОРНИК":"Вторник","ВТ":"Вторник","СРЕДА":"Среда","СР":"Среда","ЧЕТВЕРГ":"Четверг","ЧТ":"Четверг","ПЯТНИЦА":"Пятница","ПТ":"Пятница","СУББОТА":"Суббота","СБ":"Суббота","ВОСКРЕСЕНЬЕ":"Воскресенье","ВС":"Воскресенье"}
ROMAN={"I":1,"II":2,"III":3,"IV":4,"V":5,"VI":6,"VII":7,"VIII":8,"IX":9,"X":10,"XI":11,"XII":12}
GROUP_RE=re.compile(r"(?<![\wА-Яа-яЁё])([А-ЯA-ZЁёа-яё*][А-ЯA-ZЁёа-яё0-9*_-]{1,14}-\d{2,4})(?!\w)")
TEACHER_RE=re.compile(r"(?P<surname>[А-ЯЁ][а-яё-]{2,})\s+(?P<i1>[А-ЯЁ])\.?\s*(?P<i2>[А-ЯЁ])\.?",re.U)
DATE_RE=re.compile(r"(?<!\d)(\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)(?!\d)")
DATE_RANGE_RE=re.compile(r"(?<!\d)(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?\s*-\s*(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?(?!\d)")
DATE_QUALIFIER_RE=re.compile(r"(?i)(?:до|с)\s*\d{1,2}\.\d{1,2}(?:\.\d{2,4})?")
TIME_ONLY_RE=re.compile(r"^\d{1,2}[\.:]\d{2}\s*[-–—]\s*\d{1,2}[\.:]\d{2}$")
TITLE_RE=re.compile(r"\b(?:доц\.?|проф\.?|преп\.?|асс\.?|ст\.?\s*п\.?|ст\.п\.?)\b",re.I)
# Important: these suffixes are NOT classroom suffixes in the SibGIU files.
NON_ROOM_SUFFIXES={"с","ст","стп"}
ROOM_RE=re.compile(r"\b(\d{1,4})(?:\s*([А-ЯA-ZЁа-яё]{1,3}))?\b")
COMMON_SUBJECT_WORDS={"основы","психология","адаптивная","менеджмент","физиология","биохимия","педагогическое","математика","физика","химия","информатика","история","литература","география","программирование","экономика","правоведение","социология","философия","культура","теория","технология","безопасность","инженерная","иностранный","русский","рисунок","электротехника","автоматизация","физическая","организация","управление","ресурсы","надежность","техническая"}

COURSE_RE=re.compile(r"(?<!\d)([1-5])\s*курс",re.I)
MASTER_RE=re.compile(r"магистрат",re.I)


def norm(v)->str:
    if v is None:return ""
    s=str(v).replace("\xa0"," ").replace("\r","\n")
    s=re.sub(r"[ \t]+"," ",s); s=re.sub(r"\n+","\n",s)
    return s.strip()

def norm_search(v): return re.sub(r"\s+"," ",norm(v).lower().replace("ё","е").replace(".",""))
def teacher_normalize(m): return f"{m.group('surname')} {m.group('i1')}.{m.group('i2')}."

def parse_lesson(v):
    s=norm(v).upper().replace("—","-").replace("–","-")
    if not s:return None
    x=s.split()[0].strip(".,;:")
    if x in ROMAN:return ROMAN[x]
    m=re.match(r"^(\d{1,2})(?:\s|/|$)",s)
    return int(m.group(1)) if m else None

def parse_day(v):
    s=norm(v).upper().replace("Ё","Е")
    if not s:return None,None
    for k,d in DAY_NAMES.items():
        if s==k or s.startswith(k+" ") or s.startswith(k+"."):
            ds=DATE_RE.findall(s); return d,(ds[0] if ds else None)
    m=re.match(r"^(ПН|ВТ|СР|ЧТ|ПТ|СБ|ВС)\.?\s*(\d{1,2}\.\d{1,2})?",s)
    return (DAY_NAMES[m.group(1)],m.group(2)) if m else (None,None)

def merged_map(ws):
    amap={}; spans={}
    for rng in ws.merged_cells.ranges:
        a,b,c,d=range_boundaries(str(rng)); anchor=(b,a); spans[anchor]=(d-b+1,c-a+1)
        for r in range(b,d+1):
            for col in range(a,c+1): amap[(r,col)]=anchor
    return amap,spans

def get_cell(ws,amap,r,c):
    a=amap.get((r,c)); return norm(ws.cell(*(a or (r,c))).value)
def span_for(amap,spans,r,c):
    a=amap.get((r,c)); return (*a,*spans[a]) if a else (r,c,1,1)

def find_sections(ws,amap,spans):
    rows=collections.defaultdict(list)
    for r in range(1,ws.max_row+1):
        for c in range(1,ws.max_column+1):
            a=amap.get((r,c))
            if a and a!=(r,c):continue
            v=norm(ws.cell(r,c).value); groups=GROUP_RE.findall(v)
            if not groups:continue
            ar,ac,rs,cs=span_for(amap,spans,r,c)
            if ar!=r:continue
            for g in groups: rows[r].append({"group":g,"c1":ac,"c2":ac+cs-1,"header_ranges":DATE_RANGE_RE.findall(v)})
    sections=[]
    for r in sorted(rows):
        seen=set(); specs=[]
        for s in sorted(rows[r],key=lambda x:x['c1']):
            if s['group'] not in seen:seen.add(s['group']);specs.append(s)
        if specs:sections.append((r,specs))
    good=[]
    for i,(r,specs) in enumerate(sections):
        nxt=sections[i+1][0]-1 if i+1<len(sections) else ws.max_row
        found=False
        for rr in range(r+1,min(nxt,r+30)+1):
            for c in range(1,min(ws.max_column,6)+1):
                a=amap.get((rr,c))
                if a and a!=(rr,c):continue
                if parse_day(get_cell(ws,amap,rr,c))[0]:found=True;break
            if found:break
        if found:good.append((r,specs))
    return good

def detect_day_lesson_columns(ws,amap,start_row=1,end_row=None):
    end_row=end_row or ws.max_row; day_col=None
    for r in range(start_row,min(end_row,start_row+30)+1):
        for c in range(1,min(ws.max_column,6)+1):
            a=amap.get((r,c))
            if a and a!=(r,c):continue
            if parse_day(get_cell(ws,amap,r,c))[0]:day_col=c;break
        if day_col:break
    return (1,2) if day_col is None else (day_col,min(day_col+1,ws.max_column))

def row_fragments(ws,amap,spans,spec,r,protected_cols=()):
    vals=[];seen=set()
    for c in range(spec['c1'],spec['c2']+1):
        a=amap.get((r,c))
        if a:
            if a in seen:continue
            ar,ac=a; rs,cs=spans.get(a,(1,1))
            if ac<spec['c1'] and ac+cs-1<spec['c1']:continue
            if ac in protected_cols:continue
            seen.add(a)
        v=get_cell(ws,amap,r,c)
        if v and v not in vals:vals.append(v)
    return vals

def room_extract(text):
    s=norm(text)
    if not s:return None
    if re.search(r"\bдист(?:анционно)?\b",s,re.I):return "Дист"
    t=DATE_RANGE_RE.sub(" ",s); t=DATE_RE.sub(" ",t)
    t=re.sub(r"\b\d{1,2}[\.:]\d{2}\s*[-–—]\s*\d{1,2}[\.:]\d{2}\b"," ",t)
    pairs=re.findall(r"\b\d{1,4}(?:\s*[А-ЯA-ZЁа-яё]{0,3})?\s*/\s*\d{1,4}(?:\s*[А-ЯA-ZЁа-яё]{0,3})?\b",t)
    if pairs:
        tok=re.sub(r"\s+","",pairs[-1]);
        # validate both parts independently
        if all((m.group(2) or "").lower() not in NON_ROOM_SUFFIXES for m in re.finditer(r"(\d{1,4})(?:([А-ЯA-ZЁа-яё]{1,3}))?",tok)): return tok
    good=[]
    for m in ROOM_RE.finditer(t):
        n=int(m.group(1)); suf=(m.group(2) or "").lower()
        if n<10 or n>9999:continue
        if suf in NON_ROOM_SUFFIXES:continue
        # Do not turn an isolated project number into a room.
        good.append(m.group(0).replace(" ",""))
    return good[-1] if good else None

def extract_periods(text):
    return DATE_QUALIFIER_RE.findall(norm(text))


def infer_week_from_periods(periods, default):
    """Infer parity from explicit semester qualifiers used in converted Excel."""
    if not periods:
        return default
    vals=[]
    for p in periods:
        m=re.search(r"(\d{1,2})\.(\d{1,2})",p)
        if not m: continue
        day,month=int(m.group(1)),int(m.group(2))
        # Autumn 2026: the source schedules use the first half as odd and
        # the second half as even. Explicit `до` dates before late October
        # are odd; explicit `с` dates from late October onward are even.
        if re.match(r"(?i)с\s*",p): vals.append("even" if (month>10 or (month==10 and day>=27)) else "odd")
        elif re.match(r"(?i)до\s*",p): vals.append("odd" if (month<10 or (month==10 and day<=26)) else "even")
    if vals and all(v==vals[0] for v in vals): return vals[0]
    return default

def clean_subject(text):
    s=norm(text)
    if not s:return None
    # Dates/titles are metadata, not part of the subject.
    s=DATE_QUALIFIER_RE.sub(" ",s)
    # Converter sometimes glues the Russian preposition to the project number:
    # "Проек деятельность 2с 11.11" means project 2, not room "2с".
    s=re.sub(r"(?<=\d)с(?=\s*\d{1,2}\.\d{1,2})", "", s, flags=re.I)
    s=TITLE_RE.sub(" ",s)
    # Remove room only at the END. This deliberately preserves numbers in names:
    # "Проектная деятельность 1" stays intact.
    m=re.search(r"\s+(\d{1,4})(?:\s*([А-ЯA-ZЁа-яё]{1,3}))?\s*$",s)
    if m and (m.group(2) or "").lower() not in NON_ROOM_SUFFIXES and int(m.group(1)) >= 10:
        s=s[:m.start()].strip()
    s=re.sub(r"\s+"," ",s).strip(" /-;,.")
    return s or None

def has_teacher(text): return bool(TEACHER_RE.search(norm(text)))

def is_teacher_match(m,raw):
    surname=m.group('surname').lower().replace('ё','е'); prefix=raw[max(0,m.start()-16):m.start()]
    has_title=bool(re.search(r"(?:доц\.?|проф\.?|преп\.?|асс\.?|ст\.?\s*п\.?|ст\.п\.?)\s*$",prefix,re.I))
    return not (surname in COMMON_SUBJECT_WORDS and not has_title)

def extract_entries(raw):
    raw=norm(raw)
    if not raw:return []
    teachers=[m for m in TEACHER_RE.finditer(raw) if is_teacher_match(m,raw)]
    if not teachers:
        return [{"subject":clean_subject(raw),"teacher":None,"room":room_extract(raw),"date_ranges":[f"{a}-{b}" for a,b in DATE_RANGE_RE.findall(raw)],"periods":extract_periods(raw),"raw":raw}]
    out=[]
    for i,m in enumerate(teachers):
        prev_end=teachers[i-1].end() if i else 0; next_start=teachers[i+1].start() if i+1<len(teachers) else len(raw)
        before=raw[prev_end:m.start()].strip(" /;,\n")
        if i and re.match(r"^\d{1,4}[А-ЯA-ZЁа-яё]{0,3}\b",before): before=re.sub(r"^\d{1,4}[А-ЯA-ZЁа-яё]{0,3}\b\s*","",before,count=1).strip(" /;,\n")
        after=raw[m.end():next_start].strip(" /;,\n")
        subject=clean_subject(before)
        room=room_extract(after)
        # For a single teacher, room may be before the teacher; remove it from subject.
        if subject and room and subject.endswith(room): subject=subject[:-len(room)].strip(" /-") or None
        tail = after if i == len(teachers)-1 else ""
        dr=[f"{a}-{b}" for a,b in DATE_RANGE_RE.findall(before+" "+tail)]
        periods=extract_periods(before+" "+tail)
        out.append({"subject":subject,"teacher":teacher_normalize(m),"room":room,"date_ranges":dr,"periods":periods,"raw":raw})
    # A title-only fragment between teachers can make the subject None. In
    # shared multi-teacher cells inherit the first meaningful subject.
    base=next((e['subject'] for e in out if e['subject']),None)
    if base:
        for e in out:
            if not e['subject']:e['subject']=base
    # If the whole cell has one room, it is shared by all teachers unless a
    # segment has a more specific room.
    all_room=room_extract(raw)
    if all_room and len(teachers)>1:
        for e in out:
            if not e['room']:e['room']=all_room
    return out

def build_institute(path):
    parts=list(path.parts)
    for p in reversed(parts):
        m=re.search(r"(?:расписание[_-])(.+)",p,re.I)
        if m:return m.group(1).strip()
    return path.parent.name

def detect_level_course(path):
    s=str(path)
    if MASTER_RE.search(s): return "Магистратура", None
    m=COURSE_RE.search(path.name)
    return ("Бакалавриат",int(m.group(1))) if m else ("Бакалавриат",None)

def parse_section(ws,amap,spans,institute,source_file,header_row,specs,end_row):
    day_col,lesson_col=detect_day_lesson_columns(ws,amap,header_row+1,end_row)
    teacher_records=[]; group_records=[]; current_day=None; current_date=None; current_lesson=None; lesson_start=None; r=header_row+1
    level,course=detect_level_course(Path(source_file))
    while r<=end_row:
        a=amap.get((r,day_col)); day_is_anchor=(a is None or a==(r,day_col))
        day,date=parse_day(get_cell(ws,amap,r,day_col)) if day_is_anchor else (None,None)
        if day:
            current_day=day; current_date=date; current_lesson=None; lesson_start=None; r+=1; continue
        lesson=parse_lesson(get_cell(ws,amap,r,lesson_col))
        if lesson is not None and (current_lesson!=lesson or lesson_start is None):current_lesson=lesson;lesson_start=r
        if not current_day or current_lesson is None or lesson_start is None:r+=1;continue
        end=r;rr=r+1
        while rr<=end_row:
            da=amap.get((rr,day_col)); d,_=parse_day(get_cell(ws,amap,rr,day_col)) if (da is None or da==(rr,day_col)) else (None,None)
            l=parse_lesson(get_cell(ws,amap,rr,lesson_col))
            if d or (l is not None and l!=current_lesson):break
            end=rr;rr+=1
        if r!=lesson_start:r+=1;continue
        for spec in specs:
            rows=[]
            for rr in range(lesson_start,end+1):
                vals=row_fragments(ws,amap,spans,spec,rr,protected_cols=(day_col,lesson_col));raw=" / ".join(vals)
                if raw and not TIME_ONLY_RE.fullmatch(raw):rows.append((rr,raw))
            if not rows:continue
            unique=[]
            for item in rows:
                if not any(item[1]==x[1] for x in unique):unique.append(item)
            occ=[];i=0
            while i<len(unique):
                rr0,raw0=unique[i]
                if i+1<len(unique) and not has_teacher(raw0) and has_teacher(unique[i+1][1]):occ.append((rr0,norm(raw0+" / "+unique[i+1][1])));i+=2
                else:occ.append((rr0,raw0));i+=1
            date_based=bool(current_date) or bool(re.search(r"\b\d{1,2}\.\d{1,2}\b",norm(get_cell(ws,amap,lesson_start,day_col))))
            if date_based:assigns=[(raw,"date") for _,raw in occ]
            elif len(occ)==1:
                rr0,raw0=occ[0]; both=False
                for c in range(spec['c1'],spec['c2']+1):
                    a=amap.get((rr0,c))
                    if a and spans.get(a,(1,1))[0]>=2:both=True;break
                assigns=[(raw0,"odd"),(raw0,"even")] if both else [(raw0,"odd")]
            else:assigns=[(raw,"odd" if idx%2==0 else "even") for idx,(_,raw) in enumerate(occ)]
            for raw,week_type in assigns:
                entries=extract_entries(raw)
                # When a converted cell contains upper/lower-week text in one
                # string, explicit `до...` / `с...` qualifiers tell us which
                # half each teacher belongs to. Split those entries instead
                # of pretending two teachers teach simultaneously.
                entry_weeks=[infer_week_from_periods(e.get('periods',[]),week_type) for e in entries]
                for e,ew in zip(entries,entry_weeks):
                    if not e.get('teacher') and not e.get('subject'): continue
                    if e.get('teacher'):
                        teacher_records.append({"institute":institute,"source_file":source_file,"sheet":ws.title,"group":spec['group'],"level":level,"course":course,"day":current_day,"date_text":current_date,"lesson":current_lesson,"week_type":ew,"subject":e['subject'],"teacher":e['teacher'],"teacher_search":norm_search(e['teacher']),"room":e['room'],"date_ranges":e['date_ranges'] or [f"{a}-{b}" for a,b in spec['header_ranges']],"lesson_periods":e.get('periods',[]),"raw":e['raw']})
                # Student-side records: one record per parity. If multiple
                # entries belong to different halves, keep them separate.
                by_week=collections.defaultdict(list)
                for e,ew in zip(entries,entry_weeks): by_week[ew].append(e)
                for ew, es in by_week.items():
                    teachers=[];rooms=[];subject=None;dr=[];periods=[]
                    for e in es:
                        if e.get('subject') and not subject:subject=e['subject']
                        if e.get('teacher') and e['teacher'] not in teachers:teachers.append(e['teacher'])
                        if e.get('room') and e['room'] not in rooms:rooms.append(e['room'])
                        dr.extend(e.get('date_ranges') or []);periods.extend(e.get('periods') or [])
                    if subject or teachers or rooms:
                        group_records.append({"institute":institute,"source_file":source_file,"sheet":ws.title,"group":spec['group'],"level":level,"course":course,"day":current_day,"date_text":current_date,"lesson":current_lesson,"week_type":ew,"subject":subject,"teacher":"; ".join(teachers) or None,"room":"; ".join(rooms) or None,"date_ranges":sorted(set(dr)) or [f"{a}-{b}" for a,b in spec['header_ranges']],"lesson_periods":sorted(set(periods)),"raw":raw})
        r=end+1
    return teacher_records,group_records

def parse_sheet(ws,institute,source_file):
    amap,spans=merged_map(ws); sections=find_sections(ws,amap,spans)
    if not sections:return [],[]
    t=[];g=[]
    for i,(hr,specs) in enumerate(sections):
        end=sections[i+1][0]-1 if i+1<len(sections) else ws.max_row
        a,b=parse_section(ws,amap,spans,institute,source_file,hr,specs,end);t.extend(a);g.extend(b)
    return t,g

def parse_workbook(path,institute):
    wb=load_workbook(path,data_only=True,read_only=False);t=[];g=[]
    for ws in wb.worksheets:
        if ws.max_row<=1 and ws.max_column<=1:continue
        try:
            a,b=parse_sheet(ws,institute,path.name);t.extend(a);g.extend(b)
        except Exception as e:print(f"WARN {path.name} [{ws.title}]: {e}")
    return t,g

def build_records(root):
    root=Path(root);teacher=[];group=[]
    for p in sorted(root.rglob('*.xlsx')):
        institute=build_institute(p)
        t,g=parse_workbook(p,institute);teacher.extend(t);group.extend(g)
        print(f"{institute}: {p.name} -> teachers {len(t)}, group {len(g)}")
    def dedup(items,key):
        out=[];seen=set()
        for r in items:
            k=key(r)
            if k not in seen:seen.add(k);out.append(r)
        return out
    teacher=dedup(teacher,lambda r:(r['institute'],r['source_file'],r['sheet'],r['group'],r['day'],r['date_text'],r['lesson'],r['week_type'],r['subject'],r['teacher'],r['room'],tuple(r['date_ranges']),tuple(r.get('lesson_periods',[]))))
    group=dedup(group,lambda r:(r['institute'],r['source_file'],r['sheet'],r['group'],r['day'],r['date_text'],r['lesson'],r['week_type'],r['subject'],r['teacher'],r['room'],tuple(r['date_ranges']),tuple(r.get('lesson_periods',[]))))
    return teacher,group

def create_db(records,db_path):
    teachers,groups=records; db_path=Path(db_path);db_path.parent.mkdir(parents=True,exist_ok=True)
    con=sqlite3.connect(db_path);cur=con.cursor();cur.execute('DROP TABLE IF EXISTS lessons');cur.execute('DROP TABLE IF EXISTS group_lessons')
    cur.execute('''CREATE TABLE lessons(id INTEGER PRIMARY KEY,institute TEXT NOT NULL,source_file TEXT,sheet TEXT,group_name TEXT NOT NULL,level TEXT,course INTEGER,day TEXT NOT NULL,date_text TEXT,lesson INTEGER NOT NULL,week_type TEXT NOT NULL,subject TEXT,teacher TEXT NOT NULL,teacher_search TEXT NOT NULL,room TEXT,date_ranges TEXT,lesson_periods TEXT,raw TEXT)''')
    cur.execute('''CREATE TABLE group_lessons(id INTEGER PRIMARY KEY,institute TEXT NOT NULL,source_file TEXT,sheet TEXT,group_name TEXT NOT NULL,level TEXT,course INTEGER,day TEXT NOT NULL,date_text TEXT,lesson INTEGER NOT NULL,week_type TEXT NOT NULL,subject TEXT,teacher TEXT,room TEXT,date_ranges TEXT,lesson_periods TEXT,raw TEXT)''')
    tr=[(r['institute'],r['source_file'],r['sheet'],r['group'],r['level'],r['course'],r['day'],r['date_text'],r['lesson'],r['week_type'],r['subject'],r['teacher'],r['teacher_search'],r['room'],json.dumps(r['date_ranges'],ensure_ascii=False),json.dumps(r.get('lesson_periods',[]),ensure_ascii=False),r['raw']) for r in teachers]
    gr=[(r['institute'],r['source_file'],r['sheet'],r['group'],r['level'],r['course'],r['day'],r['date_text'],r['lesson'],r['week_type'],r['subject'],r['teacher'],r['room'],json.dumps(r['date_ranges'],ensure_ascii=False),json.dumps(r.get('lesson_periods',[]),ensure_ascii=False),r['raw']) for r in groups]
    cur.executemany('INSERT INTO lessons(institute,source_file,sheet,group_name,level,course,day,date_text,lesson,week_type,subject,teacher,teacher_search,room,date_ranges,lesson_periods,raw) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',tr)
    cur.executemany('INSERT INTO group_lessons(institute,source_file,sheet,group_name,level,course,day,date_text,lesson,week_type,subject,teacher,room,date_ranges,lesson_periods,raw) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',gr)
    cur.execute('CREATE INDEX idx_teacher ON lessons(teacher_search)');cur.execute('CREATE INDEX idx_group ON group_lessons(group_name,day,lesson)');cur.execute('CREATE INDEX idx_group_filter ON group_lessons(level,course,institute,group_name)')
    con.commit();con.close();return len(teachers),len(groups)

def _decode_rows(rows):
    out=[dict(x) for x in rows]
    for x in out:
        try:x['date_ranges']=json.loads(x['date_ranges'])
        except:x['date_ranges']=[]
        try:x['lesson_periods']=json.loads(x.get('lesson_periods') or '[]')
        except:x['lesson_periods']=[]
    return out

def search_teacher(query,day=None,week_type=None,db_path='schedule.db'):
    q=norm_search(query)
    if not q:return []
    con=sqlite3.connect(db_path);con.row_factory=sqlite3.Row
    sql='SELECT institute,group_name,day,date_text,lesson,week_type,subject,teacher,room,date_ranges,lesson_periods FROM lessons WHERE teacher_search LIKE ?';p=[f'%{q}%']
    if day:sql+=' AND day=?';p.append(day)
    if week_type:sql+=' AND (week_type=? OR week_type="date")';p.append(week_type)
    sql+=" ORDER BY CASE day WHEN 'Понедельник' THEN 1 WHEN 'Вторник' THEN 2 WHEN 'Среда' THEN 3 WHEN 'Четверг' THEN 4 WHEN 'Пятница' THEN 5 WHEN 'Суббота' THEN 6 ELSE 99 END,lesson,institute,group_name"
    rows=con.execute(sql,p).fetchall();con.close();return _decode_rows(rows)

def list_teachers(query,db_path='schedule.db'):
    q=norm_search(query);con=sqlite3.connect(db_path);rows=con.execute('SELECT teacher FROM lessons WHERE teacher_search LIKE ? GROUP BY teacher ORDER BY teacher LIMIT 30',(f'%{q}%',)).fetchall();con.close();return [r[0] for r in rows]

def list_schedule_filters(db_path='schedule.db'):
    con=sqlite3.connect(db_path);con.row_factory=sqlite3.Row
    rows=con.execute('SELECT DISTINCT level,course,institute,group_name FROM group_lessons ORDER BY level,course,institute,group_name').fetchall();con.close();return [dict(r) for r in rows]

def list_institutes(level,course,db_path='schedule.db'):
    con=sqlite3.connect(db_path);rows=con.execute('SELECT DISTINCT institute FROM group_lessons WHERE level=? AND course IS ? ORDER BY institute',(level,course)).fetchall();con.close();return [r[0] for r in rows]

def list_groups(level,course,institute,db_path='schedule.db'):
    con=sqlite3.connect(db_path);rows=con.execute('SELECT DISTINCT group_name FROM group_lessons WHERE level=? AND course IS ? AND institute=? ORDER BY group_name',(level,course,institute)).fetchall();con.close();return [r[0] for r in rows]

def search_group(group_name,day,week_type,db_path='schedule.db'):
    con=sqlite3.connect(db_path);con.row_factory=sqlite3.Row
    rows=con.execute('SELECT institute,group_name,day,date_text,lesson,week_type,subject,teacher,room,date_ranges,lesson_periods FROM group_lessons WHERE group_name=? AND day=? AND (week_type=? OR week_type="date") ORDER BY lesson',(group_name,day,week_type)).fetchall();con.close();return _decode_rows(rows)

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('root',nargs='?',default='teacher_schedules');ap.add_argument('--db',default='schedule.db');args=ap.parse_args()
    rec=build_records(args.root);nt,ng=create_db(rec,args.db);print(f'Готово: преподаватели {nt}, групповые занятия {ng}');print('Преподавателей:',len({r['teacher'] for r in rec[0]}));print('Групп:',len({r['group'] for r in rec[1]}))
