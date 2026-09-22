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
# Official SibGIU bell schedule (sibsiu.ru/raspisanie/): each pair's start
# time. Some sheets (notably evening/master's schedules) label pairs with
# purely local numbering ("I", "II"...) that only counts how many rows are
# printed for that particular day, not the university's actual pair number
# - but they do print the real start time next to it (e.g. "I 18.00-19-30"),
# and that time always matches exactly one of these 8 slots. Whenever a
# recognizable time is present, it is a far more reliable source of the true
# pair number than the local roman numeral.
OFFICIAL_PAIR_START = {1:(8,30), 2:(10,0), 3:(12,0), 4:(13,50), 5:(15,30), 6:(17,10), 7:(18,0), 8:(19,35)}


def resolve_pair_by_time(s: str) -> Optional[int]:
    m = re.search(r"(\d{1,2})[.:](\d{2})", s)
    if not m:
        return None
    hh, mm = int(m.group(1)), int(m.group(2))
    if hh > 23 or mm > 59:
        return None
    total = hh*60+mm
    for pair, (ph, pm) in OFFICIAL_PAIR_START.items():
        if abs(total-(ph*60+pm)) <= 10:
            return pair
    return None
GROUP_RE = re.compile(r"(?<![\wА-Яа-яЁё])([А-ЯA-ZЁёа-яё*][А-ЯA-ZЁёа-яё0-9*_-]{1,14}-\d{2,4})(?!\w)")
# NOTE on the fix below: the previous version made the periods after both
# initials fully optional ("\.?"), which let it match the first two letters
# of an unrelated ALL-CAPS subject abbreviation that happens to follow a
# capitalized word, e.g. "Автоматизированный ЭТПМиК" was misread as teacher
# "Автоматизированный Э.Т.". Requiring a period after the first initial and
# forbidding another letter right after the second initial removes those
# false positives while still matching every real "Фамилия И.О." / "Фамилия И.О"
# style occurrence found in the source files (verified against the full corpus).
TEACHER_RE = re.compile(r"(?P<surname>[А-ЯЁ][а-яё-]{2,})\s+(?P<i1>[А-ЯЁ])\.\s*(?P<i2>[А-ЯЁ])\.?(?![А-Яа-яЁёA-Za-z])", re.U)
DATE_RE = re.compile(r"(?<!\d)(\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)(?!\d)")
DATE_RANGE_RE = re.compile(r"(?<!\d)(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?\s*-\s*(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?(?!\d)")
TIME_ONLY_RE = re.compile(r"^\d{1,2}[\.:]\d{2}\s*[-–—]\s*\d{1,2}[\.:]\d{2}$")
# NOTE on the fix below: the previous version required a word boundary
# *after* the optional trailing period ("доц\.?\b"), which can never match
# right before a period followed by a space/newline (both are non-word
# characters, so there is no \b there). That made the regex quietly refuse
# to consume the period, leaving a dangling " ." in thousands of subjects
# (e.g. "Философия ." instead of "Философия"). Moving the boundary check to
# right after the bare title word fixes this and also adds the very common
# "с.п." ("старший преподаватель") abbreviation, which was missing entirely.
TITLE_RE = re.compile(r"\b(?:доц|проф|преп|асс|ст\.?\s*п|с\.?\s*п)(?![А-Яа-яЁёA-Za-z])\.?", re.I)
# Extracts the course number from filenames like "3 курс ИТУР ...", "1 Курс ...",
# "Магистратура 2 курс ...". Two spaces between the digit and "курс" (a common
# typo in the source files, e.g. "4 курс  ИИТиАС") are matched by \s*.
COURSE_RE = re.compile(r"(\d)\s*курс", re.IGNORECASE)
DEGREE_FOLDERS = {"Бакалавриат", "Магистратура"}

# "до 14.10" / "с 11.11" — a course-phase qualifier meaning "runs through this
# date" / "starts from this date". Never printed to the student; only used to
# decide whether a lesson is currently active (see resolve/valid_from/until
# below). The negative lookbehind keeps this from matching inside a longer
# word (e.g. the "с" in "часть").
DATE_QUALIFIER_RE = re.compile(r"(?<![а-яёa-z])(?P<kind>до|по|с)\s*(?P<day>\d{1,2})[.,](?P<month>\d{1,2})(?!\d)", re.IGNORECASE)
# SibGIU's "project work" subgroup marker: "Проек[тная] деятельность N [до|с DATE]".
# The number is a subgroup index, not a room, and the date is a validity
# qualifier, not free text — both are parsed out explicitly so neither one
# leaks into room detection (this was the source of "2с" being misread as a
# room number: "2" is subgroup 2, "с" starts the "с 11.11" qualifier for it).
PROJECT_RE = re.compile(
    r"проек(?:т(?:ная)?)?\s*деятельность\s*(?P<num>\d+)?"
    r"(?:\s*(?P<kind>до|с)\s*(?P<day>\d{1,2})[.,](?P<month>\d{1,2}))?",
    re.IGNORECASE,
)
# "Название – Курс в записи размещенный в СУО/Moodle СибГИУ" — an
# asynchronous, pre-recorded course. Not a live pair (no teacher, no room),
# but still worth showing as an informational note rather than silently
# dropping it or gluing several of these together into one unreadable blob.
ONLINE_COURSE_RE = re.compile(r"курс\s+(?:лекций\s+)?в\s*(?:записи)?\s*,?\s*размещенн\w*\s+в\s+(?:СУО|moodle)[^А-Яа-яA-Za-z]*(?:СибГИУ)?", re.IGNORECASE)


def _valid_date(day, month) -> bool:
    try:
        return 1 <= int(day) <= 31 and 1 <= int(month) <= 12
    except (TypeError, ValueError):
        return False


def resolve_year(month: int, filename: str) -> int:
    """Academic year for a bare 'DD.MM' date found in a file named like
    '...Осенний семестр 2026-2027...'. Autumn semester: Jul-Dec belongs to
    the first year in the filename, Jan-Jun to the second (exam period
    spilling into the new year). Spring semester: the reverse."""
    m = re.search(r"(20\d{2})\s*-\s*(20\d{2})", filename)
    if m:
        y1, y2 = int(m.group(1)), int(m.group(2))
    else:
        now = datetime.now()
        y1, y2 = now.year, now.year + 1
    autumn = "осенн" in filename.lower()
    month = int(month)
    if autumn:
        return y1 if month >= 7 else y2
    return y2 if month <= 6 else y1


def make_iso_date(day, month, filename: str) -> Optional[str]:
    if not _valid_date(day, month):
        return None
    try:
        return datetime(resolve_year(month, filename), int(month), int(day)).date().isoformat()
    except ValueError:
        return None


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
    by_time = resolve_pair_by_time(s)
    if by_time is not None:
        return by_time
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
        return "Дистанционно"
    # Remove dates/times before room detection.
    t=DATE_RANGE_RE.sub(" ", s)
    t=DATE_RE.sub(" ", t)
    t=re.sub(r"\b\d{1,2}[\.:]\d{2}\s*[-–—]\s*\d{1,2}[\.:]\d{2}\b", " ", t)
    # Common room formats: 361ГТ, 8П, 610аМ, 515, 513 / 212, 368 Г.
    matches=list(re.finditer(r"\b(\d{1,4})(?:\s*([А-ЯA-ZЁа-яё]{1,3}))?\b", t))
    good=[]
    for m in matches:
        n=int(m.group(1))
        if n < 1 or n > 9999: continue
        good.append((m.start(), m.end(), m.group(0).strip().replace(" ", "")))
    if not good: return None
    # A lesson can run in more than one room at once (parallel subgroups,
    # e.g. "ИиКГ Голодова М.А. 501Г, 506Г"). Keep every room in such a list
    # instead of only the last one: walk backwards from the last match and
    # keep chaining while consecutive matches are separated by nothing but
    # ", " / "/" / " / " (a genuine room list), stopping at the first gap
    # that contains other text (which means the earlier number is not a
    # room at all, e.g. a stray subgroup index).
    chain=[good[-1]]
    for cand in reversed(good[:-1]):
        gap=t[cand[1]:chain[0][0]]
        if re.fullmatch(r"\s*[,/]\s*", gap):
            chain.insert(0, cand)
        else:
            break
    return ", ".join(c[2] for c in chain)


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

def clean_subject(s: Optional[str]) -> Optional[str]:
    if not s: return s
    # A parenthesized list of specific calendar dates (occasional extra
    # sessions), e.g. "Основы ПД (16.09, 14.10, 11.11, 09.12)" - not shown to
    # the student, same as any other date noise.
    s=re.sub(r"\(\s*[\d.,\s]+\)", " ", s)
    # Same thing without parentheses, e.g. "09,22.09 5П Информац технологии"
    # or "15.09,29.09 225Г Химия" - a bare list of 2+ dates (with an
    # optional leading bare day-of-month before the first full date).
    s=re.sub(r"(?<![\d.])\d{0,2}\.?\d{1,2}\.\d{1,2}(?:\s*[,;]\s*\d{1,2}\.\d{1,2}){1,}", " ", s)
    # Leftover empty "()" / "( )" after a date qualifier was stripped out of
    # a bracketed note, e.g. "Основы ТЖ (с 08.09)" -> "Основы ТЖ ()".
    s=re.sub(r"\(\s*\)", " ", s)
    s=re.sub(r"\s+", " ", s).strip(" /-;.,")
    return s or None


TRAILING_ROOM_RE = re.compile(
    r"\s*[-–—]?\s*(?P<rooms>\d{1,4}(?:\s*[А-ЯA-ZЁа-яё]{1,3})?(?:\s*,\s*\d{1,4}(?:\s*[А-ЯA-ZЁа-яё]{1,3})?)*)\s*$"
)


def strip_trailing_rooms(text: str):
    """Remove a trailing room or room-list (optionally dash-separated) from a
    subject fragment, e.g. "НИР - 247Г, 504Г, 421М" -> "НИР". Applied to text
    that has already had dates/times/project markers removed, so any
    remaining trailing "number[+letters]" group(s) can only be rooms.
    Returns (cleaned_text, room_or_None) — some sheets print the room
    *before* the teacher's name with nothing after it (e.g. "Математика
    260Г доц Ионина А.В"), and that room must not simply be discarded."""
    m=TRAILING_ROOM_RE.search(text)
    if not m:
        return text, None
    room=", ".join(t.strip().replace(" ","") for t in m.group('rooms').split(','))
    return TRAILING_ROOM_RE.sub("", text), room


def extract_entries(raw: str, filename: str = ""):
    raw=norm(raw)
    if not raw: return []

    def extract_project(text):
        """Pull out 'Проек[тная] деятельность N [до|с DATE]'. Returns
        (subject_or_None, valid_from, valid_until, remaining_text)."""
        pm=PROJECT_RE.search(text)
        if not pm: return None, None, None, text
        num=pm.group('num')
        subject="Проектная деятельность"+(f" {num}" if num else "")
        vf=vu=None
        if pm.group('kind') and _valid_date(pm.group('day'), pm.group('month')):
            d=make_iso_date(pm.group('day'), pm.group('month'), filename)
            if d:
                if pm.group('kind').lower() in ('до','по'): vu=d
                else: vf=d
        remaining=text[:pm.start()]+text[pm.end():]
        return subject, vf, vu, remaining

    def extract_dates(text):
        """Pull out a bare 'A.B-C.D' range or a до/с DATE qualifier not tied
        to a project marker (e.g. "Основы ТЖ (с 08.09)"). Returns
        (valid_from, valid_until, remaining_text)."""
        rm=DATE_RANGE_RE.search(text)
        if rm:
            vf=make_iso_date(*rm.group(1).split('.'), filename)
            vu=make_iso_date(*rm.group(2).split('.'), filename)
            return vf, vu, text[:rm.start()]+text[rm.end():]
        qm=DATE_QUALIFIER_RE.search(text)
        if qm and _valid_date(qm.group('day'), qm.group('month')):
            d=make_iso_date(qm.group('day'), qm.group('month'), filename)
            if d:
                vf,vu=(None,d) if qm.group('kind').lower() in ('до','по') else (d,None)
                return vf, vu, text[:qm.start()]+text[qm.end():]
        return None, None, text

    teachers=[m for m in TEACHER_RE.finditer(raw) if is_teacher_match(m, raw)]

    if not teachers and ONLINE_COURSE_RE.search(raw):
        # "Математика – Курс в записи размещенный в СУО СибГИУ\nИнформатика –
        # Курс в записи..." — one or more pre-recorded (asynchronous) courses
        # glued together. Split into separate informational notes instead of
        # returning one unreadable blob. Only taken when nothing in the cell
        # looks like a real teacher assignment: a few cells mix an online
        # note for one subgroup with a real "title Фамилия И.О." entry for
        # another, and that entry must not be swallowed here.
        chunks=re.split(r"(?<=СибГИУ)", raw)
        out=[]
        for chunk in chunks:
            chunk=chunk.strip(" \n.;,")
            if not chunk: continue
            m=re.match(r"^(.*?)[–—-]\s*Курс", chunk, re.IGNORECASE)
            name=(m.group(1).strip() if m else chunk).strip(" \n.;,")
            if name:
                out.append({"subject":f"{name} — курс в записи (СУО СибГИУ)","teacher":None,"room":None,
                            "valid_from":None,"valid_until":None,"raw":raw})
        if out:
            return out

    if not teachers:
        proj_subject, vf, vu, cleaned=extract_project(raw)
        if vf is None and vu is None:
            vf, vu, cleaned=extract_dates(cleaned)
        subject=proj_subject or re.sub(r"\s+", " ", TITLE_RE.sub(" ", strip_trailing_rooms(cleaned)[0])).strip(" /-")
        subject=clean_subject(subject)
        room_src=PROJECT_RE.sub(" ", cleaned)
        return [{"subject": subject or None, "teacher":None,
                 "room":room_extract(room_src), "valid_from":vf, "valid_until":vu, "raw":raw}]
    out=[]
    for i,m in enumerate(teachers):
        prev_end=teachers[i-1].end() if i else 0
        next_start=teachers[i+1].start() if i+1<len(teachers) else len(raw)
        before=raw[prev_end:m.start()].strip(" /;,\n")
        after=raw[m.end():next_start].strip(" /;,\n")
        # A project marker belonging to the *next* subgroup (and its date)
        # can leak into this entry's trailing text; strip it before room
        # detection but never use its subject/date for this entry.
        after_clean=PROJECT_RE.sub(" ", after)
        proj_subject, vf, vu, before_clean=extract_project(before)
        if vf is None and vu is None:
            vf, vu, before_clean=extract_dates(before_clean)
        # Only *after* any date has been removed do we strip a leaked room
        # number from the start of `before` (a room for the *previous*
        # subgroup that had nowhere else to go, e.g. "212Г Проек деятельность
        # 2 ..."). Doing this before the date extraction would instead eat
        # the leading day digits of a date like "09.09-21.10", turning it
        # into the mangled leftover ".09-21.10".
        if i and re.match(r"^\d{1,4}[А-ЯA-ZЁа-яё]{0,3}\b", before_clean):
            before_clean=re.sub(r"^\d{1,4}[А-ЯA-ZЁа-яё]{0,3}\b\s*", "", before_clean, count=1).strip(" /;,\n")
        if vf is None and vu is None:
            vf, vu, after_clean=extract_dates(after_clean)
        else:
            _,_,after_clean=extract_dates(after_clean)
        subject=proj_subject
        leading_room=None
        if subject is None:
            subject=TITLE_RE.sub(" ", before_clean)
            # In several SibGIU sheets a room is printed before the teacher
            # (e.g. "Математика 363Г доц. Ионина А.В. 260Г", or, when
            # nothing follows the teacher's name at all, "Математика 260Г
            # доц Ионина А.В" — here that leading room *is* the room, not
            # just noise, so it is kept as a fallback below rather than
            # discarded).
            subject,leading_room=strip_trailing_rooms(subject)
            subject=re.sub(r"\s+", " ", subject).strip(" /-;")
            # A few sheets separate the discipline from the teacher with a bare
            # "." instead of a title abbreviation (e.g. "Арх-ра ГиПЗ .Матехина"),
            # or have a doubled ".." typo after a title. Either way a lone,
            # word-less "." left over after the cleanup above is noise, not part
            # of the subject name.
            subject=re.sub(r"(?:^|\s)\.(?=\s|$)", " ", subject)
            subject=re.sub(r"\s+", " ", subject).strip(" /-;.")
        subject=clean_subject(subject)
        room=room_extract(after_clean) or leading_room
        out.append({"subject":subject or None,"teacher":teacher_normalize(m),"room":room,
                    "valid_from":vf,"valid_until":vu,"raw":raw})
    # Several sheets list one shared discipline once, followed by multiple
    # "title Фамилия И.О." teachers for parallel subgroups, e.g.
    # "Проектная деятельность доц Иванов А.А, преп Петров Б.Б, ст.п Сидоров В.В".
    # Only the first teacher in such a chain gets a non-empty `before` text;
    # without this fill the other subgroup teachers show up with an empty
    # subject even though they clearly teach the same discipline.
    first_subject = next((e["subject"] for e in out if e["subject"]), None)
    if first_subject:
        for e in out:
            if not e["subject"]:
                e["subject"] = first_subject
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


def parse_section(ws,amap,spans,institute,source_file,header_row,specs,end_row,degree=None,course=None):
    day_col,lesson_col=detect_day_lesson_columns(ws,amap,header_row+1,end_row)
    records=[]; current_day=None; current_date=None; current_lesson=None; lesson_start=None
    r=header_row+1
    while r<=end_row:
        a=amap.get((r,day_col)); day_is_anchor=(a is None or a==(r,day_col))
        day,date=parse_day(get_cell(ws,amap,r,day_col)) if day_is_anchor else (None,None)
        if day:
            current_day=day; current_date=date; current_lesson=None; lesson_start=None; r+=1; continue
        lesson=parse_lesson(get_cell(ws,amap,r,lesson_col))
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
                if raw and not TIME_ONLY_RE.fullmatch(raw): rows.append((rr,raw))
            if not rows: continue
            # Track each distinct fragment's first AND last row (not just
            # first): a value living in a cell merged across several rows
            # will be returned identically for every physical row it
            # covers, and knowing how far down it truly extends is what
            # lets the block below tell "this content fills the whole
            # lesson slot -> happens every week" apart from "this content
            # only fills half of it -> only happens on one parity".
            unique=[]
            for rr,txt in rows:
                match=next((u for u in unique if u['text']==txt), None)
                if match: match['last']=rr
                else: unique.append({'text':txt,'first':rr,'last':rr})
            occ=[]; i=0
            while i<len(unique):
                cur=unique[i]
                if i+1<len(unique) and not has_teacher(cur['text']) and has_teacher(unique[i+1]['text']):
                    nxt=unique[i+1]
                    occ.append({'start':cur['first'],'end':nxt['last'],'text':norm(cur['text']+" / "+nxt['text'])})
                    i+=2
                else:
                    occ.append({'start':cur['first'],'end':cur['last'],'text':cur['text']}); i+=1
            date_based=bool(current_date) or bool(re.search(r"\b\d{1,2}\.\d{1,2}\b",norm(get_cell(ws,amap,lesson_start,day_col))))
            if date_based:
                assigns=[(o['text'],"date") for o in occ]
            elif len(occ)==1 and occ[0]['start']<=lesson_start and occ[0]['end']>=end:
                # The only content found in this slot spans the *entire*
                # physical height of the block (no other row held anything
                # different) -> this is one lesson rendered across several
                # visual lines, not a genuine odd/even split, so it recurs
                # every week.
                assigns=[(occ[0]['text'],"odd"),(occ[0]['text'],"even")]
            else:
                # The pair's cell is visually split into a top half
                # (нечётная/odd week) and a bottom half (чётная/even week).
                # When the block's total height is odd, the extra row goes
                # to the top half (matches every such block found in the
                # source files, e.g. a 3-row block splits 2 top / 1 bottom).
                height=end-lesson_start+1
                boundary=lesson_start+(height+1)//2-1
                assigns=[]
                for o in occ:
                    mid=(o['start']+o['end'])/2
                    assigns.append((o['text'], "odd" if mid<=boundary else "even"))
            for raw,week_type in assigns:
                for e in extract_entries(raw, source_file):
                    if not e['teacher'] and not e['subject'] and not e['room']:
                        continue
                    date_iso=make_iso_date(*current_date.split('.'), source_file) if (week_type=="date" and current_date) else None
                    records.append({"institute":institute,"source_file":source_file,"sheet":ws.title,
                      "degree":degree,"course":course,
                      "group":spec['group'],"day":current_day,"date_text":current_date,"date_iso":date_iso,"lesson":current_lesson,
                      "week_type":week_type,"subject":clean_subject(e['subject']),
                      "teacher":e['teacher'] or "","teacher_search":norm_search(e['teacher']) if e['teacher'] else "",
                      "room":e['room'],"valid_from":e.get('valid_from'),"valid_until":e.get('valid_until'),"raw":e['raw']})
        r=end+1
    return records


def parse_sheet(ws,institute,source_file,degree=None,course=None):
    amap,spans=merged_map(ws); sections=find_sections(ws,amap,spans)
    if not sections: return []
    out=[]
    for i,(header_row,specs) in enumerate(sections):
        end=sections[i+1][0]-1 if i+1<len(sections) else ws.max_row
        out.extend(parse_section(ws,amap,spans,institute,source_file,header_row,specs,end,degree,course))
    return out


def extract_course(filename: str) -> Optional[int]:
    """Course number (1-5) parsed from a schedule filename, e.g.
    "3 курс ИТУР Осенний семестр 2026-2027.xlsx" -> 3. Returns None when the
    filename has no "N курс" pattern (some Магистратура files are named only
    by track code, e.g. "Магистратура ИПИТ ПМКМ-26 ...").
    """
    m = COURSE_RE.search(filename)
    return int(m.group(1)) if m else None


def parse_workbook(path: Path, institute: str, degree: Optional[str] = None, course: Optional[int] = None):
    wb=load_workbook(path,data_only=True,read_only=False)
    out=[]
    for ws in wb.worksheets:
        if ws.max_row<=1 and ws.max_column<=1: continue
        try: out.extend(parse_sheet(ws,institute,path.name,degree,course))
        except Exception as e: print(f"WARN {path.name} [{ws.title}]: {e}")
    return out


def _damerau_levenshtein(a: str, b: str, cutoff: int = 2) -> int:
    """Edit distance counting adjacent-letter transpositions as 1 edit.
    Used only to spot near-identical surnames that are almost certainly the
    same person typed inconsistently across dozens of Excel files (e.g.
    "Леммермайер"/"Леммерайер"/"Лиммермайер"), never to guess who someone is.
    """
    n, m = len(a), len(b)
    if abs(n - m) > cutoff:
        return cutoff + 1
    d = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(n + 1):
        d[i][0] = i
    for j in range(m + 1):
        d[0][j] = j
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            d[i][j] = min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost)
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                d[i][j] = min(d[i][j], d[i - 2][j - 2] + 1)
    return d[n][m]


def canonicalize_teachers(records):
    """Merge teacher spellings that differ by a single typo in the source
    Excel files, e.g. "Зайленко К.С." vs "Зайлено К.С.", or "Серенков С.Ю."
    vs "Серенков Ю.С." (initials swapped). Without this, the same teacher's
    lessons are split across 2-4 near-identical names in teacher_search, so a
    student looking them up only sees part of their real schedule.

    This is intentionally conservative: it only merges names that share the
    exact same initials (or, for an identical surname, the same two initials
    in a different order) and whose surname differs by at most one edit
    (letter swap/insertion/deletion/adjacent transposition), with a minimum
    surname length so short/common name fragments are never touched. This
    keeps genuinely different people (e.g. "Серкова"/"Перова", both Т.Ю.)
    safely apart, since they differ by more than one edit.
    """
    by_teacher = collections.defaultdict(list)
    for r in records:
        by_teacher[r["teacher"]].append(r)
    names = list(by_teacher.keys())
    parsed = []
    for name in names:
        parts = name.split(" ", 1)
        if len(parts) == 2 and re.match(r"^[А-ЯЁ]\.[А-ЯЁ]\.?$", parts[1]):
            parsed.append((name, parts[0], parts[1].rstrip(".")))
        else:
            parsed.append((name, name, None))

    parent = {name: name for name in names}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(x, y):
        rx, ry = find(x), find(y)
        if rx != ry:
            parent[rx] = ry

    for i in range(len(parsed)):
        name1, surname1, ini1 = parsed[i]
        if ini1 is None:
            continue
        for j in range(i + 1, len(parsed)):
            name2, surname2, ini2 = parsed[j]
            if ini2 is None or surname1 == surname2 and ini1 == ini2:
                continue
            if surname1 == surname2 and sorted(ini1) == sorted(ini2):
                union(name1, name2)  # same surname, initials just reordered
                continue
            if ini1 != ini2:
                continue
            if min(len(surname1), len(surname2)) < 5:
                continue
            if _damerau_levenshtein(surname1, surname2, cutoff=1) <= 1:
                union(name1, name2)

    clusters = collections.defaultdict(list)
    for name in names:
        clusters[find(name)].append(name)

    # How many *distinct* initials each surname spelling is seen with across
    # the whole corpus. A spelling that recurs under other initials too
    # (e.g. "Никитин" also exists as "Никитин В.В.") is corroborated
    # independently elsewhere and is a safer canonical pick than a one-off
    # spelling nobody else shares (e.g. "Никиктин"), which is most likely a
    # single data-entry typo.
    surname_variety = collections.defaultdict(set)
    for _, surname, ini in parsed:
        if ini is not None:
            surname_variety[surname].add(ini)

    rename = {}
    search_extra = {}
    for members in clusters.values():
        if len(members) == 1:
            continue
        # Canonical spelling = the variant confirmed by the most *distinct*
        # source files (a typo repeated many times within a single
        # spreadsheet should not outvote a spelling used consistently
        # elsewhere); then by how independently corroborated the surname is
        # elsewhere in the corpus; then by row count; then alphabetically.
        def rank(n):
            files = {r["source_file"] for r in by_teacher[n]}
            surname = n.split(" ", 1)[0]
            return (-len(files), -len(surname_variety[surname]), -len(by_teacher[n]), n)
        canonical = sorted(members, key=rank)[0]
        # Keep every historical spelling searchable even though only the
        # canonical one is displayed: a student typing the "correct"
        # spelling should still find lessons filed under a typo, and
        # vice versa.
        combined_search = " ".join(sorted({norm_search(n) for n in members}))
        search_extra[canonical] = combined_search
        for n in members:
            if n != canonical:
                rename[n] = canonical

    if not rename and not search_extra:
        return records
    for r in records:
        if r["teacher"] in rename:
            r["teacher"] = rename[r["teacher"]]
        extra = search_extra.get(r["teacher"])
        r["teacher_search"] = extra if extra else norm_search(r["teacher"])
    return records


def build_records(root: str|Path):
    root=Path(root)
    records=[]
    files=sorted(root.rglob('*.xlsx'))
    for p in files:
        institute=p.parent.parent.name if p.parent.name in {'Бакалавриат','Магистратура'} else p.parent.name
        if institute in {'teacher_schedules',''}: institute=p.parent.name
        degree=p.parent.name if p.parent.name in DEGREE_FOLDERS else None
        course=extract_course(p.name)
        rec=parse_workbook(p,institute,degree,course)
        print(f"{institute}: {p.name} -> {len(rec)} records (degree={degree}, course={course})")
        records.extend(rec)
    records=canonicalize_teachers(records)
    # Exact de-duplication.
    unique=[]; seen=set()
    for r in records:
        key=(r['institute'],r['source_file'],r['sheet'],r['degree'],r['course'],r['group'],r['day'],r['date_text'],r.get('date_iso'),r['lesson'],r['week_type'],r['subject'],r['teacher'],r['room'],r.get('valid_from'),r.get('valid_until'))
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
        degree TEXT, course INTEGER,
        group_name TEXT NOT NULL, day TEXT NOT NULL, date_text TEXT, date_iso TEXT,
        lesson INTEGER NOT NULL, week_type TEXT NOT NULL,
        subject TEXT, teacher TEXT NOT NULL, teacher_search TEXT NOT NULL,
        room TEXT, valid_from TEXT, valid_until TEXT, raw TEXT)''')
    rows=[(r['institute'],r['source_file'],r['sheet'],r.get('degree'),r.get('course'),r['group'],r['day'],r['date_text'],r.get('date_iso'),r['lesson'],r['week_type'],r['subject'],r['teacher'],r['teacher_search'],r['room'],r.get('valid_from'),r.get('valid_until'),r['raw']) for r in records]
    cur.executemany('''INSERT INTO lessons(institute,source_file,sheet,degree,course,group_name,day,date_text,date_iso,lesson,week_type,subject,teacher,teacher_search,room,valid_from,valid_until,raw)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',rows)
    cur.execute('CREATE INDEX idx_teacher ON lessons(teacher_search)')
    cur.execute('CREATE INDEX idx_teacher_day ON lessons(teacher_search,day,lesson)')
    cur.execute('CREATE INDEX idx_group ON lessons(group_name,day,lesson)')
    cur.execute('CREATE INDEX idx_degree_course_institute ON lessons(degree,course,institute)')
    cur.execute('CREATE INDEX idx_degree_course_institute_group ON lessons(degree,course,institute,group_name)')
    con.commit(); con.close()
    return len(records)


def is_active(valid_from: Optional[str], valid_until: Optional[str], today: Optional[str] = None) -> bool:
    """Whether a lesson with this validity window should be shown "today".
    Both bounds are inclusive ISO dates; either (or both) may be None,
    meaning "no restriction on that side"."""
    if not valid_from and not valid_until:
        return True
    today = today or datetime.now().date().isoformat()
    if valid_from and today < valid_from:
        return False
    if valid_until and today > valid_until:
        return False
    return True


def search_teacher(query, day=None, week_type=None, db_path='schedule.db'):
    q=norm_search(query)
    if not q: return []
    con=sqlite3.connect(db_path); con.row_factory=sqlite3.Row; cur=con.cursor()
    sql='''SELECT institute,degree,course,group_name,day,date_text,date_iso,lesson,week_type,subject,teacher,room,valid_from,valid_until FROM lessons WHERE teacher_search LIKE ?'''
    params=[f'%{q}%']
    if day:
        sql+=' AND day=?'; params.append(day)
    if week_type:
        sql+=' AND (week_type=? OR week_type="date")'; params.append(week_type)
    sql+=''' ORDER BY CASE day WHEN 'Понедельник' THEN 1 WHEN 'Вторник' THEN 2 WHEN 'Среда' THEN 3 WHEN 'Четверг' THEN 4 WHEN 'Пятница' THEN 5 WHEN 'Суббота' THEN 6 WHEN 'Воскресенье' THEN 7 ELSE 99 END, lesson, institute, group_name'''
    out=[dict(x) for x in cur.execute(sql,params).fetchall()]
    con.close()
    return out


def list_teachers(query, db_path='schedule.db'):
    q=norm_search(query)
    con=sqlite3.connect(db_path); cur=con.cursor()
    rows=cur.execute('SELECT teacher FROM lessons WHERE teacher_search LIKE ? GROUP BY teacher ORDER BY teacher LIMIT 30',(f'%{q}%',)).fetchall()
    con.close(); return [r[0] for r in rows]


# =========================
# ОБЫЧНОЕ РАСПИСАНИЕ (ПО ГРУППАМ): курс/магистратура -> институт -> группа
# =========================
def list_institutes(degree, course=None, db_path='schedule.db'):
    """Distinct institutes that have at least one lesson for the given
    degree ('Бакалавриат' / 'Магистратура'). For 'Бакалавриат', course
    (1-5) narrows it down; for 'Магистратура' course is ignored, since
    master's programs are offered as a single bucket regardless of year."""
    con=sqlite3.connect(db_path); cur=con.cursor()
    if course is None:
        rows=cur.execute('SELECT DISTINCT institute FROM lessons WHERE degree=? ORDER BY institute',(degree,)).fetchall()
    else:
        rows=cur.execute('SELECT DISTINCT institute FROM lessons WHERE degree=? AND course=? ORDER BY institute',(degree,course)).fetchall()
    con.close(); return [r[0] for r in rows]


def list_groups(degree, institute, course=None, db_path='schedule.db'):
    """Distinct group names for the given degree/institute (and course, for
    Бакалавриат)."""
    con=sqlite3.connect(db_path); cur=con.cursor()
    if course is None:
        rows=cur.execute('SELECT DISTINCT group_name FROM lessons WHERE degree=? AND institute=? ORDER BY group_name',(degree,institute)).fetchall()
    else:
        rows=cur.execute('SELECT DISTINCT group_name FROM lessons WHERE degree=? AND institute=? AND course=? ORDER BY group_name',(degree,institute,course)).fetchall()
    con.close(); return [r[0] for r in rows]


def search_group(group_name, institute=None, day=None, week_type=None, db_path='schedule.db'):
    """Lessons for one group. `institute` should always be passed when known:
    a handful of group codes (e.g. "ТИС-26") are reused by two different
    institutes, so filtering by group_name alone can mix two unrelated
    schedules together."""
    con=sqlite3.connect(db_path); con.row_factory=sqlite3.Row; cur=con.cursor()
    sql='''SELECT institute,group_name,day,date_text,date_iso,lesson,week_type,subject,teacher,room,valid_from,valid_until FROM lessons WHERE group_name=?'''
    params=[group_name]
    if institute:
        sql+=' AND institute=?'; params.append(institute)
    if day:
        sql+=' AND day=?'; params.append(day)
    if week_type:
        sql+=' AND (week_type=? OR week_type="date")'; params.append(week_type)
    sql+=''' ORDER BY CASE day WHEN 'Понедельник' THEN 1 WHEN 'Вторник' THEN 2 WHEN 'Среда' THEN 3 WHEN 'Четверг' THEN 4 WHEN 'Пятница' THEN 5 WHEN 'Суббота' THEN 6 WHEN 'Воскресенье' THEN 7 ELSE 99 END, lesson'''
    out=[dict(x) for x in cur.execute(sql,params).fetchall()]
    con.close()
    return out


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
