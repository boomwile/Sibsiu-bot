# -*- coding: utf-8 -*-
import asyncio
import html
import json
import os
import re
import sqlite3
import threading
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from aiogram import Bot, Dispatcher, F, types
from aiogram.exceptions import TelegramBadRequest
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardButton
from aiogram.utils.keyboard import InlineKeyboardBuilder
from fastapi import FastAPI
import uvicorn

from teacher_schedule_parser import (
    list_teachers,
    search_teacher,
    TEACHER_RE,
    TITLE_RE,
    teacher_normalize,
    norm_search,
)

# =========================
# НАСТРОЙКИ
# =========================
TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    raise RuntimeError("Не задана переменная окружения BOT_TOKEN")

TEACHER_DB = Path(os.getenv("TEACHER_DB", "teacher_schedule.db"))

bot = Bot(token=TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# =========================
# КЛАВИАТУРЫ
# =========================
main_kb = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Расписание"), KeyboardButton(text="👨‍🏫 Найти преподавателя")],
        [KeyboardButton(text="Навигация по корпусам"), KeyboardButton(text="Куда обратиться?")],
        [KeyboardButton(text="Учёба и сессия"), KeyboardButton(text="Стипендии и соцподдержка")],
        [KeyboardButton(text="Общежития"), KeyboardButton(text="Контакты")],
        [KeyboardButton(text="Мероприятия")]
    ],
    resize_keyboard=True
)

DAYS = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
DAY_ORDER = {day: i for i, day in enumerate(DAYS, 1)}
PAIR_TIMES = {
    1: "08:30 - 09:50",
    2: "10:00 - 11:30",
    3: "12:00 - 13:30",
    4: "13:50 - 15:20",
    5: "15:30 - 17:00",
    6: "17:10 - 18:40",
}
COURSES = ["1 курс", "2 курс", "3 курс", "4 курс", "5 курс", "Магистратура"]
INSTITUTE_LABELS = {
    "расписание_АСИ": "АСИ",
    "расписание_ИГДиГ": "ИГДиГ",
    "расписание_ИПИТ": "ИПИТ",
    "расписание_ИИТиАС": "ИИТиАС",
    "расписание_ИМиМ": "ИМиМ",
    "расписание_IFKZIC": "ИФКЗиС",
    "расписание_ИПО": "ИПО",
    "расписание_ИТУР": "ИТУР",
    "расписание_СПО": "СПО",
}

stub_buttons = [
    "Навигация по корпусам", "Куда обратиться?", "Учёба и сессия",
    "Стипендии и соцподдержка", "Общежития", "Контакты", "Мероприятия"
]

# =========================
# НЕДЕЛЯ
# =========================
def get_week_parity():
    """Нечётная/чётная учебная неделя.
    На субботе и воскресенье показываем следующую неделю.
    01.09.2026 — нечётная.
    """
    today = datetime.now().date()
    week_date = today + timedelta(days=(7 - today.weekday())) if today.weekday() >= 5 else today
    semester_start = date(2026, 9, 1)
    week_monday = week_date - timedelta(days=week_date.weekday())
    start_monday = semester_start - timedelta(days=semester_start.weekday())
    week_index = (week_monday - start_monday).days // 7
    return "нечетная" if week_index % 2 == 0 else "четная"


def week_title():
    return "Чётная" if get_week_parity() == "четная" else "Нечётная"


def get_teacher_week_parity():
    return "even" if get_week_parity() == "четная" else "odd"

# =========================
# СОСТОЯНИЯ
# =========================
class TeacherSearch(StatesGroup):
    waiting_surname = State()
    choosing_teacher = State()
    choosing_day = State()


class ScheduleState(StatesGroup):
    choosing_course = State()
    choosing_institute = State()
    choosing_group = State()
    choosing_day = State()

# =========================
# DB
# =========================
def db_has_group_table() -> bool:
    if not TEACHER_DB.exists():
        return False
    try:
        con = sqlite3.connect(TEACHER_DB)
        row = con.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='group_lessons'"
        ).fetchone()
        con.close()
        return bool(row)
    except Exception:
        return False


def db_connect():
    con = sqlite3.connect(TEACHER_DB)
    con.row_factory = sqlite3.Row
    return con

# =========================
# ВСПОМОГАТЕЛЬНОЕ РАСПИСАНИЕ
# =========================
def source_course(source_file: str) -> str | None:
    s = (source_file or "").lower()
    if "магистратура" in s:
        return "Магистратура"
    m = re.search(r"(?<!\d)([1-5])\s*курс\b", s, re.I)
    if m:
        return f"{m.group(1)} курс"
    return None


def institute_label(raw: str) -> str:
    return INSTITUTE_LABELS.get(raw, raw.removeprefix("расписание_"))


def catalog_for_course(course: str):
    """Return {raw_institute: set(groups)} from the same SQLite database."""
    result = defaultdict(set)
    if not db_has_group_table():
        return result
    con = db_connect()
    rows = con.execute(
        "SELECT DISTINCT institute, source_file, group_name FROM group_lessons "
        "WHERE group_name IS NOT NULL AND group_name <> ''"
    ).fetchall()
    con.close()
    for r in rows:
        if source_course(r["source_file"]) != course:
            continue
        result[r["institute"]].add(r["group_name"])
    return result


def all_course_catalog():
    out = {course: defaultdict(set) for course in COURSES}
    if not db_has_group_table():
        return out
    con = db_connect()
    rows = con.execute(
        "SELECT DISTINCT institute, source_file, group_name FROM group_lessons "
        "WHERE group_name IS NOT NULL AND group_name <> ''"
    ).fetchall()
    con.close()
    for r in rows:
        course = source_course(r["source_file"])
        if course in out:
            out[course][r["institute"]].add(r["group_name"])
    return out

# =========================
# РАЗБОР СЫРОГО ТЕКСТА ДЛЯ ОТОБРАЖЕНИЯ
# (DB/парсер Клода не переписываем; исправляем только отображение)
# =========================
DATE_RANGE_DISPLAY_RE = re.compile(
    r"(?<!\d)(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?\s*[-–—]\s*(\d{1,2}\.\d{1,2})(?:\.\d{2,4})?(?!\d)"
)
DATE_BEFORE_RE = re.compile(r"(?:\b|(?<=\d))(?:до|по)\s*(\d{1,2}\.\d{1,2})(?!\d)", re.I)
DATE_AFTER_RE = re.compile(r"(?:\b|(?<=\d))с\s*(\d{1,2}\.\d{1,2})(?!\d)", re.I)
DATE_LIST_RE = re.compile(r"(?<!\d)(\d{1,2}\.\d{1,2})(?:\s*,\s*\d{1,2}\.\d{1,2})+(?!\d)")
ROOM_TOKEN_RE = re.compile(r"(?<!\w)(\d{1,4}\s*[А-ЯA-ZЁа-яё]{0,3})(?!\w)")
DIST_RE = re.compile(r"\bдист(?:анционно)?\b", re.I)
TITLE_ONLY_RE = re.compile(
    r"\b(?:доц|проф|преп|асс|ст\.?\s*п|с\.?\s*п)\.?\b", re.I
)


def compact(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ").replace("\r", " ").replace("\n", " ")).strip()


def current_semester_year() -> int:
    now = datetime.now().date()
    return now.year if now.month >= 8 else now.year


def month_day(value: str) -> date:
    d, m = [int(x) for x in value.split(".")[:2]]
    year = current_semester_year()
    return date(year, m, d)


def segment_dates(text: str):
    """Return date constraints explicitly written near this lesson segment."""
    s = compact(text)
    constraints = []
    for m in DATE_RANGE_DISPLAY_RE.finditer(s):
        constraints.append(("range", month_day(m.group(1)), month_day(m.group(2))))
    for m in DATE_LIST_RE.finditer(s):
        for value in re.findall(r"\d{1,2}\.\d{1,2}", m.group(0)):
            constraints.append(("exact", month_day(value), None))
    for m in DATE_BEFORE_RE.finditer(s):
        constraints.append(("before", month_day(m.group(1)), None))
    for m in DATE_AFTER_RE.finditer(s):
        constraints.append(("after", month_day(m.group(1)), None))
    return constraints


def segment_is_active(text: str, today: date | None = None) -> bool:
    constraints = segment_dates(text)
    if not constraints:
        return True
    today = today or datetime.now().date()
    for kind, a, b in constraints:
        if kind == "range" and a <= today <= b:
            return True
        if kind == "before" and today <= a:
            return True
        if kind == "exact" and today == a:
            return True
        if kind == "after" and today >= a:
            return True
    return False


def leading_after_date_part(text: str) -> str:
    """Date qualifier immediately after a teacher belongs to that teacher.
    A later qualifier after a new subject must not leak backwards.
    """
    s = compact(text)
    s = re.sub(r"^[,;/\s]+", "", s)
    m = re.match(r"^((?:\d{1,2}\.\d{1,2}(?:\s*[-–—]\s*\d{1,2}\.\d{1,2})?)|(?:до|по|с)\s*\d{1,2}\.\d{1,2})", s, re.I)
    return m.group(1) if m else ""


def clean_date_words(s: str) -> str:
    s = DATE_RANGE_DISPLAY_RE.sub(" ", s)
    s = DATE_BEFORE_RE.sub(" ", s)
    s = DATE_AFTER_RE.sub(" ", s)
    # Specific SibGIU notation: "Проектная деятельность 2с 18.11".
    s = re.sub(r"(?<=\d)\s*с(?=\s*\d{1,2}\.\d{1,2})", " ", s, flags=re.I)
    return s


def normalize_title(title: str | None) -> str:
    if not title:
        return ""
    x = re.sub(r"\s+", "", title.lower()).rstrip(".")
    return {
        "доц": "доц.",
        "проф": "проф.",
        "преп": "преп.",
        "асс": "асс.",
        "стп": "ст.п.",
        "сп": "с.п.",
    }.get(x, title.strip())


def room_tokens(text: str):
    """Conservative room extractor used only for display.
    Important: 2с/3с/1с are NOT rooms.
    """
    if not text:
        return []
    if DIST_RE.search(text):
        return ["Дист"]
    t = clean_date_words(text)
    found = []
    for m in ROOM_TOKEN_RE.finditer(t):
        token = compact(m.group(1)).replace(" ", "")
        digits = re.match(r"^(\d+)", token)
        if not digits:
            continue
        n = int(digits.group(1))
        suffix = token[len(digits.group(1)):].lower()
        # Project/date abbreviations, not classrooms.
        if suffix in {"с", "ст", "стп", "сп", "до", "по"}:
            continue
        # Bare 1-2 digit numbers are normally project/other notation.
        if not suffix and n < 100:
            continue
        # Long bare numbers can be real rooms; suffix rooms like 5П/8П are fine.
        if token not in found:
            found.append(token)
    return found


def clean_subject(text: str) -> str:
    s = clean_date_words(compact(text))
    s = TITLE_ONLY_RE.sub(" ", s)
    s = DIST_RE.sub(" ", s)
    # Remove room tokens only when they are not the meaningful final subject number.
    def repl_room(m):
        token = compact(m.group(1)).replace(" ", "")
        digits = re.match(r"^(\d+)", token)
        if not digits:
            return m.group(0)
        n = int(digits.group(1))
        suffix = token[len(digits.group(1)):].lower()
        if suffix in {"с", "ст", "стп", "сп", "до", "по"}:
            return m.group(0)
        if suffix or n >= 100:
            return " "
        return m.group(0)
    s = ROOM_TOKEN_RE.sub(repl_room, s)
    s = re.sub(r"[;,/]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" .-/;")
    return s


def reminder_text(raw: str, subject: str | None = None) -> str | None:
    low = compact(raw).lower()
    if "курс лекций" in low and ("суо сибгиу" in low or "сибгиу" in low) and "размещ" in low:
        return "Курс лекций размещён в СУО СИБГИУ"
    if "курс в записи" in low and "размещ" in low:
        if "moodle" in low:
            return "Курс в записи размещён в Moodle"
        return "Курс в записи размещён на учебной платформе"
    return None


def split_raw_entries(raw: str):
    raw = raw or ""
    raw_norm = raw.replace("\r", " ").replace("\n", " ")
    matches = list(TEACHER_RE.finditer(raw_norm))
    if not matches:
        subject = clean_subject(raw_norm)
        rooms = room_tokens(raw_norm)
        return [{
            "subject": subject,
            "teacher": None,
            "teacher_display": None,
            "room": ("Дист" if "Дист" in rooms else (", ".join(rooms) if rooms else None)),
            "active": segment_is_active(raw_norm),
            "reminder": reminder_text(raw_norm, subject),
            "raw_segment": raw_norm,
        }]

    entries = []
    prev_end = 0
    for i, m in enumerate(matches):
        next_start = matches[i + 1].start() if i + 1 < len(matches) else len(raw_norm)
        before = raw_norm[prev_end:m.start()]
        after = raw_norm[m.end():next_start]
        title_matches = list(TITLE_RE.finditer(before))
        title = normalize_title(title_matches[-1].group(0) if title_matches else None)
        subject = clean_subject(before)
        teacher = teacher_normalize(m)
        after_rooms = room_tokens(after)
        before_rooms = room_tokens(before)
        room = ", ".join(after_rooms or before_rooms) if (after_rooms or before_rooms) else None
        active_text = before + " " + leading_after_date_part(after)
        entries.append({
            "subject": subject,
            "teacher": teacher,
            "teacher_display": f"{title} {teacher}".strip() if title else teacher,
            "room": room,
            "active": segment_is_active(active_text),
            "reminder": None,
            "raw_segment": before + " " + after,
        })
        prev_end = m.end()

    # Shared subject for chains where only the first teacher has the subject.
    last_subject = None
    for e in entries:
        if e["subject"] and e["subject"] != "Занятие":
            last_subject = e["subject"]
        elif last_subject:
            e["subject"] = last_subject

    # If one room is written once after several teachers, it belongs to all.
    all_rooms = room_tokens(raw_norm)
    unique_rooms = []
    for r in all_rooms:
        if r not in unique_rooms:
            unique_rooms.append(r)
    if len(entries) > 1 and len(unique_rooms) == 1:
        for e in entries:
            if not e["room"]:
                e["room"] = unique_rooms[0]
    elif len(entries) == len(unique_rooms):
        for e, r in zip(entries, unique_rooms):
            if not e["room"]:
                e["room"] = r

    return entries


def find_teacher_entry(raw: str, teacher: str):
    wanted = norm_search(teacher)
    entries = split_raw_entries(raw)
    for e in entries:
        if e["teacher"] and norm_search(e["teacher"]) == wanted:
            return e
    # Canonicalized spelling can differ from raw source spelling. Match surname+initials loosely.
    surname = (teacher or "").split()[0].lower().replace("ё", "е")
    for e in entries:
        if e["teacher"] and e["teacher"].split()[0].lower().replace("ё", "е") == surname:
            return e
    return None

# =========================
# START / ЗАГЛУШКИ
# =========================
@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer("Привет! Выбери нужный раздел в меню:", reply_markup=main_kb)


@dp.message(F.text.in_(stub_buttons))
async def handle_stubs(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(f'Раздел "{message.text}" находится в разработке 🛠')

# =========================
# ПОИСК ПРЕПОДАВАТЕЛЯ
# =========================
@dp.message(F.text == "👨‍🏫 Найти преподавателя")
async def teacher_start(message: types.Message, state: FSMContext):
    await state.set_state(TeacherSearch.waiting_surname)
    await message.answer(
        "👨‍🏫 <b>Поиск преподавателя</b>\n\n"
        "Введите фамилию преподавателя. Можно написать только фамилию или фамилию с инициалами.",
        parse_mode="HTML"
    )


async def send_teacher_day_keyboard(message: types.Message, state: FSMContext):
    data = await state.get_data()
    teacher = data.get("teacher", "")
    builder = InlineKeyboardBuilder()
    for day in DAYS:
        builder.button(text=day, callback_data=f"teacher_day:{day}")
    builder.button(text="📅 Вся неделя", callback_data="teacher_day:ALL")
    builder.adjust(2)
    await message.answer(
        f"✅ Выбран преподаватель: <b>{html.escape(teacher)}</b>\n"
        f"Текущая неделя: <b>{week_title()}</b>\n\nВыберите день:",
        reply_markup=builder.as_markup(),
        parse_mode="HTML"
    )


@dp.message(TeacherSearch.waiting_surname)
async def teacher_surname(message: types.Message, state: FSMContext):
    query = (message.text or "").strip()
    candidates = list_teachers(query, TEACHER_DB)
    if not candidates:
        await message.answer("❌ Преподаватель не найден. Проверьте написание фамилии и попробуйте ещё раз.")
        return
    await state.update_data(query=query, candidates=candidates)
    if len(candidates) == 1:
        await state.update_data(teacher=candidates[0])
        await state.set_state(TeacherSearch.choosing_day)
        await send_teacher_day_keyboard(message, state)
        return
    builder = InlineKeyboardBuilder()
    for i, teacher in enumerate(candidates[:20]):
        builder.button(text=teacher, callback_data=f"teacher_select:{i}")
    builder.adjust(1)
    await state.set_state(TeacherSearch.choosing_teacher)
    await message.answer(
        f"Нашёл несколько совпадений по запросу <b>{html.escape(query)}</b>.\nВыберите нужного преподавателя:",
        reply_markup=builder.as_markup(), parse_mode="HTML"
    )


@dp.callback_query(TeacherSearch.choosing_teacher, F.data.startswith("teacher_select:"))
async def teacher_select(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    candidates = data.get("candidates", [])
    try:
        idx = int(callback.data.split(":", 1)[1])
    except (ValueError, IndexError):
        await callback.answer("Ошибка выбора", show_alert=True)
        return
    if idx < 0 or idx >= len(candidates):
        await callback.answer("Преподаватель не найден", show_alert=True)
        return
    teacher = candidates[idx]
    await state.update_data(teacher=teacher)
    await state.set_state(TeacherSearch.choosing_day)
    builder = InlineKeyboardBuilder()
    for day in DAYS:
        builder.button(text=day, callback_data=f"teacher_day:{day}")
    builder.button(text="📅 Вся неделя", callback_data="teacher_day:ALL")
    builder.adjust(2)
    await callback.message.edit_text(
        f"✅ Выбран преподаватель: <b>{html.escape(teacher)}</b>\n"
        f"Текущая неделя: <b>{week_title()}</b>\n\nВыберите день:",
        parse_mode="HTML", reply_markup=builder.as_markup()
    )
    await callback.answer()


def format_teacher_results(rows, teacher, selected_day):
    if not rows:
        return (
            f"❌ Для <b>{html.escape(teacher)}</b> занятий по выбранному фильтру не найдено.\n"
            f"Неделя: <b>{week_title()}</b>"
        )

    # Apply the explicit lesson date constraints from raw source text. This is
    # intentionally done at display time, leaving Claude's DB/parser untouched.
    filtered = []
    for r in rows:
        entry = find_teacher_entry(r.get("raw") or "", teacher)
        if entry:
            if not entry["active"]:
                continue
            rr = dict(r)
            rr["display"] = entry
            filtered.append(rr)
        else:
            rr = dict(r)
            rr["display"] = {
                "subject": r.get("subject") or "Занятие",
                "teacher_display": r.get("teacher") or teacher,
                "room": r.get("room"),
                "active": segment_is_active(r.get("raw") or ""),
                "reminder": reminder_text(r.get("raw") or ""),
            }
            if rr["display"]["active"]:
                filtered.append(rr)

    if not filtered:
        return (
            f"❌ Для <b>{html.escape(teacher)}</b> занятий по выбранному фильтру не найдено.\n"
            f"Неделя: <b>{week_title()}</b>"
        )

    header = f"👨‍🏫 <b>{html.escape(teacher)}</b>\nНеделя: <b>{week_title()}</b>\n"
    if selected_day != "ALL":
        header += f"День: <b>{html.escape(selected_day)}</b>\n"
    header += "\n"

    blocks = []
    last_day = None
    for r in filtered[:70]:
        day = r["day"]
        if selected_day == "ALL" and day != last_day:
            blocks.append(f"\n📅 <b>{html.escape(day)}</b>")
            last_day = day
        entry = r["display"]
        lines = [f"🕒 <b>{r['lesson']} пара</b> — {html.escape(day)}"]
        if entry.get("subject") and not reminder_text(r.get("raw") or "", entry.get("subject")):
            lines.append(f"📚 {html.escape(entry['subject'])}")
        td = entry.get("teacher_display") or teacher
        if td:
            lines.append(f"👨‍🏫 {html.escape(td)}")
        room = entry.get("room")
        if room == "Дист":
            lines.append("📍 Дистанционно")
        elif room:
            lines.append(f"🚪 {html.escape(room)}")
        lines.append(f"👥 {html.escape(r['group_name'])}")
        lines.append(f"🏛 {html.escape(institute_label(r.get('institute') or ''))}")
        blocks.append("\n".join(lines))

    if len(filtered) > 70:
        blocks.append(f"\n…и ещё {len(filtered)-70} записей")
    return header + "\n\n".join(blocks)


@dp.callback_query(TeacherSearch.choosing_day, F.data.startswith("teacher_day:"))
async def teacher_day(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    teacher = data.get("teacher")
    if not teacher:
        await callback.answer("Сначала выберите преподавателя", show_alert=True)
        return
    selected = callback.data.split(":", 1)[1]
    day = None if selected == "ALL" else selected
    rows = search_teacher(teacher, day=day, week_type=get_teacher_week_parity(), db_path=TEACHER_DB)
    text = format_teacher_results(rows, teacher, selected)
    builder = InlineKeyboardBuilder()
    builder.button(text="🔙 К выбору дня", callback_data="teacher_back_days")
    builder.button(text="🔎 Новый преподаватель", callback_data="teacher_back_search")
    builder.adjust(1)
    await callback.message.edit_text(text, parse_mode="HTML", reply_markup=builder.as_markup())
    await callback.answer()


@dp.callback_query(F.data == "teacher_back_days")
async def teacher_back_days(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    teacher = data.get("teacher")
    if not teacher:
        await callback.answer("Преподаватель не выбран", show_alert=True)
        return
    await state.set_state(TeacherSearch.choosing_day)
    builder = InlineKeyboardBuilder()
    for day in DAYS:
        builder.button(text=day, callback_data=f"teacher_day:{day}")
    builder.button(text="📅 Вся неделя", callback_data="teacher_day:ALL")
    builder.adjust(2)
    await callback.message.edit_text(
        f"✅ Преподаватель: <b>{html.escape(teacher)}</b>\n\nВыберите день:",
        reply_markup=builder.as_markup(), parse_mode="HTML"
    )
    await callback.answer()


@dp.callback_query(F.data == "teacher_back_search")
async def teacher_back_search(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(TeacherSearch.waiting_surname)
    await callback.message.edit_text(
        "👨‍🏫 <b>Поиск преподавателя</b>\n\nВведите фамилию преподавателя.",
        parse_mode="HTML"
    )
    await callback.answer()

# =========================
# ОБЫЧНОЕ РАСПИСАНИЕ ИЗ ТОЙ ЖЕ SQLite DB
# =========================
def course_keyboard():
    builder = InlineKeyboardBuilder()
    for i, course in enumerate(COURSES):
        icon = "🎓" if course == "Магистратура" else "📚"
        builder.button(text=f"{icon} {course}", callback_data=f"sch_course:{i}")
    builder.adjust(2)
    return builder.as_markup()


async def show_course_menu(message: types.Message, state: FSMContext, edit: bool = False):
    await state.set_state(ScheduleState.choosing_course)
    text = "📚 <b>Расписание</b>\n\nВыберите курс:"
    kb = course_keyboard()
    if edit:
        try:
            await message.edit_text(text, reply_markup=kb, parse_mode="HTML")
        except TelegramBadRequest:
            pass
    else:
        await message.answer(text, reply_markup=kb, parse_mode="HTML")


@dp.message(F.text == "Расписание")
async def cmd_schedule(message: types.Message, state: FSMContext):
    await state.clear()
    if not db_has_group_table():
        await message.answer(
            "⚠️ В teacher_schedule.db нет таблицы группового расписания. "
            "Нужна обновлённая база из комплекта проекта."
        )
        return
    await show_course_menu(message, state)


@dp.callback_query(ScheduleState.choosing_course, F.data.startswith("sch_course:"))
async def schedule_course(callback: types.CallbackQuery, state: FSMContext):
    try:
        idx = int(callback.data.split(":", 1)[1])
        course = COURSES[idx]
    except (ValueError, IndexError):
        await callback.answer("Курс не найден", show_alert=True)
        return
    catalog = catalog_for_course(course)
    institutes = sorted(catalog.keys(), key=lambda x: institute_label(x))
    if not institutes:
        await callback.answer("Для этого курса расписаний нет", show_alert=True)
        return
    await state.update_data(course=course, institutes=institutes)
    await state.set_state(ScheduleState.choosing_institute)
    builder = InlineKeyboardBuilder()
    for i, inst in enumerate(institutes):
        builder.button(text=institute_label(inst), callback_data=f"sch_inst:{i}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору курса", callback_data="sch_back_course"))
    await callback.message.edit_text(
        f"📚 Курс: <b>{html.escape(course)}</b>\n\nВыберите институт:",
        reply_markup=builder.as_markup(), parse_mode="HTML"
    )
    await callback.answer()


@dp.callback_query(F.data == "sch_back_course")
async def schedule_back_course(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    await show_course_menu(callback.message, state, edit=True)
    await callback.answer()


@dp.callback_query(ScheduleState.choosing_institute, F.data.startswith("sch_inst:"))
async def schedule_institute(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    course = data.get("course")
    institutes = data.get("institutes", [])
    try:
        idx = int(callback.data.split(":", 1)[1])
        institute = institutes[idx]
    except (ValueError, IndexError):
        await callback.answer("Институт не найден", show_alert=True)
        return
    groups = sorted(catalog_for_course(course).get(institute, set()))
    if not groups:
        await callback.answer("Групп не найдено", show_alert=True)
        return
    await state.update_data(institute=institute, groups=groups)
    await state.set_state(ScheduleState.choosing_group)
    builder = InlineKeyboardBuilder()
    for i, group in enumerate(groups):
        builder.button(text=group, callback_data=f"sch_group:{i}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору института", callback_data="sch_back_institute"))
    await callback.message.edit_text(
        f"📚 Курс: <b>{html.escape(course)}</b>\n🏛 Институт: <b>{html.escape(institute_label(institute))}</b>\n\n"
        "Выберите группу:",
        reply_markup=builder.as_markup(), parse_mode="HTML"
    )
    await callback.answer()


@dp.callback_query(F.data == "sch_back_institute")
async def schedule_back_institute(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    course = data.get("course")
    if not course:
        await state.clear()
        await show_course_menu(callback.message, state, edit=True)
        await callback.answer()
        return
    catalog = catalog_for_course(course)
    institutes = sorted(catalog.keys(), key=lambda x: institute_label(x))
    await state.update_data(institutes=institutes)
    await state.set_state(ScheduleState.choosing_institute)
    builder = InlineKeyboardBuilder()
    for i, inst in enumerate(institutes):
        builder.button(text=institute_label(inst), callback_data=f"sch_inst:{i}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору курса", callback_data="sch_back_course"))
    await callback.message.edit_text(
        f"📚 Курс: <b>{html.escape(course)}</b>\n\nВыберите институт:",
        reply_markup=builder.as_markup(), parse_mode="HTML"
    )
    await callback.answer()


@dp.callback_query(ScheduleState.choosing_group, F.data.startswith("sch_group:"))
async def schedule_group(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    groups = data.get("groups", [])
    try:
        idx = int(callback.data.split(":", 1)[1])
        group = groups[idx]
    except (ValueError, IndexError):
        await callback.answer("Группа не найдена", show_alert=True)
        return
    await state.update_data(group=group)
    await state.set_state(ScheduleState.choosing_day)
    builder = InlineKeyboardBuilder()
    for day in DAYS:
        builder.button(text=day, callback_data=f"sch_day:{day}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору группы", callback_data="sch_back_group"))
    data = await state.get_data()
    await callback.message.edit_text(
        f"📚 Курс: <b>{html.escape(data.get('course',''))}</b>\n"
        f"🏛 {html.escape(institute_label(data.get('institute','')))}\n"
        f"🎓 Группа: <b>{html.escape(group)}</b>\n\n"
        "Выберите день недели:",
        reply_markup=builder.as_markup(), parse_mode="HTML"
    )
    await callback.answer()


@dp.callback_query(F.data == "sch_back_group")
async def schedule_back_group(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    course = data.get("course")
    institute = data.get("institute")
    if not course or not institute:
        await state.clear()
        await show_course_menu(callback.message, state, edit=True)
        await callback.answer()
        return
    groups = sorted(catalog_for_course(course).get(institute, set()))
    await state.update_data(groups=groups)
    await state.set_state(ScheduleState.choosing_group)
    builder = InlineKeyboardBuilder()
    for i, group in enumerate(groups):
        builder.button(text=group, callback_data=f"sch_group:{i}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору института", callback_data="sch_back_institute"))
    await callback.message.edit_text(
        f"📚 Курс: <b>{html.escape(course)}</b>\n🏛 {html.escape(institute_label(institute))}\n\nВыберите группу:",
        reply_markup=builder.as_markup(), parse_mode="HTML"
    )
    await callback.answer()


def schedule_target_date(day: str):
    """Дата выбранного дня в текущей учебной неделе. На выходных
    меню уже относится к следующей неделе, поэтому и дата сдвигается.
    """
    if day not in DAY_ORDER:
        return datetime.now().date()
    today = datetime.now().date()
    monday = today - timedelta(days=today.weekday())
    if today.weekday() >= 5:
        monday += timedelta(days=7)
    return monday + timedelta(days=DAY_ORDER[day] - 1)


def parse_ddmm(value: str | None, reference_year: int):
    m = re.search(r"(?<!\d)(\d{1,2})\.(\d{1,2})(?!\d)", value or "")
    if not m:
        return None
    try:
        return date(reference_year, int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def date_range_active(range_text: str, target):
    m = re.search(r"(\d{1,2}\.\d{1,2})\s*[-–—]\s*(\d{1,2}\.\d{1,2})", range_text or "")
    if not m:
        return False
    sm, em = int(m.group(1).split('.')[1]), int(m.group(2).split('.')[1])
    start_year = target.year - 1 if sm > em and sm > target.month else target.year
    end_year = start_year + 1 if sm > em else start_year
    start = parse_ddmm(m.group(1), start_year)
    end = parse_ddmm(m.group(2), end_year)
    if not start or not end:
        return False
    return start <= target <= end


def row_is_active_for_date(row, parity: str, target):
    """Apply explicit calendar dates/ranges before odd/even. This fixes rows
    like ПИЭ-26 "08.09-17.11", which were stored with the incidental odd/even
    value even though the source says they run by calendar date.
    """
    source = row.get("source_file") or ""
    course = source_course(source)
    date_text = row.get("date_text")
    ranges_raw = row.get("date_ranges")
    try:
        ranges = json.loads(ranges_raw) if isinstance(ranges_raw, str) else (ranges_raw or [])
    except Exception:
        ranges = []
    if not isinstance(ranges, list):
        ranges = []

    if course == "Магистратура":
        exact = parse_ddmm(date_text, target.year)
        return exact == target if exact else False

    # A date range and week parity can coexist. For example, in the IPIT
    # 5th-year sheet a lesson may be marked as odd-week and limited to
    # 14.09-26.10, while the following row is the even-week lesson.
    # Therefore a range must NOT override odd/even; it only limits the
    # period in which that parity row is active.
    if ranges:
        if (row.get("week_type") or "") in ("odd", "even"):
            return row.get("week_type") == parity and any(
                date_range_active(str(r), target) for r in ranges
            )
        return any(date_range_active(str(r), target) for r in ranges)

    if (row.get("week_type") or "") == "date":
        exact = parse_ddmm(date_text, target.year)
        return exact == target if exact else True

    return row.get("week_type") == parity


def query_group_day(institute: str, group: str, day: str, course: str | None = None):
    parity = get_teacher_week_parity()
    target = schedule_target_date(day)
    con = db_connect()
    rows = con.execute(
        "SELECT id, institute, group_name, day, lesson, week_type, subject, teacher, room, "
        "date_text, date_ranges, raw, source_file, time_text "
        "FROM group_lessons "
        "WHERE institute=? AND group_name=? AND day=? "
        "ORDER BY lesson, id",
        (institute, group, day)
    ).fetchall()
    con.close()

    out = []
    for r0 in rows:
        r = dict(r0)
        if course and source_course(r.get("source_file") or "") != course:
            continue
        if row_is_active_for_date(r, parity, target):
            r["_target_date"] = target.isoformat()
            out.append(r)
    return out


def format_group_schedule(rows, group: str, institute: str, day: str, course: str | None = None):
    parity = get_week_parity()
    target = schedule_target_date(day)
    blocks = defaultdict(list)
    seen = set()
    reminders = []

    for r in rows:
        raw = r.get("raw") or ""
        # The exact same raw cell can occur for multiple teacher records; parse once.
        raw_key = (r["lesson"], r["week_type"], compact(raw), r.get("source_file"))
        if raw_key in seen:
            continue
        seen.add(raw_key)
        entries = split_raw_entries(raw)
        for e in entries:
            e["_time_text"] = r.get("time_text")
            if not e.get("active"):
                continue
            rem = e.get("reminder")
            if rem:
                if rem not in reminders:
                    reminders.append(rem)
                continue
            subject = e.get("subject") or ""
            if "самостоятельн" in subject.lower():
                continue
            blocks[r["lesson"]].append(e)

    if course == "Магистратура":
        period_line = f"Дата: <b>{target.strftime('%d.%m.%Y')}</b>"
    else:
        period_line = f"Неделя: <b>{week_title()}</b>"

    lines = [
        f"📅 <b>Расписание на {html.escape(day)}</b>",
        f"🎓 Группа: <b>{html.escape(group)}</b>",
        f"🏛 {html.escape(institute_label(institute))}",
        period_line,
        ""
    ]

    has_lessons = False
    for lesson in sorted(blocks):
        entries = blocks[lesson]
        # Collapse exact duplicates while keeping different teachers/rooms.
        uniq = []
        seen_entries = set()
        for e in entries:
            key = (e.get("subject"), e.get("teacher"), e.get("teacher_display"), e.get("room"))
            if key not in seen_entries:
                seen_entries.add(key)
                uniq.append(e)
        entries = uniq
        if not entries:
            continue
        has_lessons = True
        pair_time = next((e.get("_time_text") for e in entries if e.get("_time_text")), None)
        if not pair_time and lesson in (7, 8):
            pair_time = {7: "18:00 - 19:30", 8: "19:35 - 21:05"}.get(lesson)
        if not pair_time:
            pair_time = PAIR_TIMES.get(lesson, "")
        lines.append(f"🕒 <b>{lesson} пара</b> ({html.escape(pair_time)})")

        # Group by subject so several teachers of one conceptual lesson stay together.
        subject_groups = []
        by_subject = defaultdict(list)
        order = []
        for e in entries:
            subj = e.get("subject") or ""
            if subj not in by_subject:
                order.append(subj)
            by_subject[subj].append(e)

        for subj in order:
            if subj:
                lines.append(f"📚 <b>{html.escape(subj)}</b>")
            for e in by_subject[subj]:
                teacher = e.get("teacher_display") or e.get("teacher")
                if teacher:
                    lines.append(f"👨‍🏫 {html.escape(teacher)}")
                room = e.get("room")
                if room == "Дист":
                    lines.append("📍 Дистанционно")
                elif room:
                    lines.append(f"🚪 {html.escape(room)}")
        lines.append("")

    if reminders:
        lines.append("🔔 <b>Напоминания</b>")
        for rem in reminders:
            lines.append(f"• {html.escape(rem)}")
        lines.append("")

    if not has_lessons:
        lines.append("В этот день занятий нет")

    return "\n".join(lines).rstrip()


@dp.callback_query(ScheduleState.choosing_day, F.data.startswith("sch_day:"))
async def schedule_day(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    institute = data.get("institute")
    group = data.get("group")
    course = data.get("course")
    day = callback.data.split(":", 1)[1]
    if not institute or not group or day not in DAYS:
        await callback.answer("Не удалось определить расписание", show_alert=True)
        return
    rows = query_group_day(institute, group, day, course)
    text = format_group_schedule(rows, group, institute, day, course)
    builder = InlineKeyboardBuilder()
    builder.row(InlineKeyboardButton(text="🔙 К выбору дней", callback_data="sch_back_day"))
    builder.row(InlineKeyboardButton(text="🏛 Сменить институт", callback_data="sch_back_institute"))
    await callback.message.edit_text(text, reply_markup=builder.as_markup(), parse_mode="HTML")
    await callback.answer()


@dp.callback_query(F.data == "sch_back_day")
async def schedule_back_day(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    group = data.get("group")
    if not group:
        await callback.answer()
        return
    await state.set_state(ScheduleState.choosing_day)
    builder = InlineKeyboardBuilder()
    for day in DAYS:
        builder.button(text=day, callback_data=f"sch_day:{day}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору группы", callback_data="sch_back_group"))
    await callback.message.edit_text(
        f"🎓 Группа: <b>{html.escape(group)}</b>\n\nВыберите день недели:",
        reply_markup=builder.as_markup(), parse_mode="HTML"
    )
    await callback.answer()

# =========================
# FASTAPI / RENDER
# =========================
app = FastAPI()


@app.get("/")
def index():
    return "OK"


def teacher_record_count():
    try:
        con = sqlite3.connect(TEACHER_DB)
        n = con.execute("SELECT COUNT(*) FROM lessons").fetchone()[0]
        con.close()
        return n
    except Exception:
        return 0


async def main():
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot, handle_signals=False)


def run_bot_in_thread():
    asyncio.run(main())


if __name__ == "__main__":
    t = threading.Thread(target=run_bot_in_thread, daemon=True)
    t.start()
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", "7860")))
