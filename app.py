# -*- coding: utf-8 -*-
import asyncio
import html
import json
import os
import sqlite3
from datetime import datetime, timedelta, date
from pathlib import Path
import threading

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardButton
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.exceptions import TelegramBadRequest
from fastapi import FastAPI
import uvicorn

from teacher_schedule_parser import (
    build_records, create_db, list_teachers, search_teacher,
    list_institutes, list_groups, search_group, is_active,
)

# =========================
# НАСТРОЙКИ
# =========================
TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    raise RuntimeError("Не задана переменная окружения BOT_TOKEN")

# Единая база расписаний: используется и для поиска преподавателя, и для
# обычного расписания групп (обе функции читают один и тот же parser/DB).
TEACHER_DB = Path(os.getenv("TEACHER_DB", "teacher_schedule.db"))
TEACHER_SCHEDULES = Path(os.getenv("TEACHER_SCHEDULES", "teacher_schedules"))

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

stub_buttons = [
    "Навигация по корпусам", "Куда обратиться?", "Учёба и сессия",
    "Стипендии и соцподдержка", "Общежития", "Контакты", "Мероприятия"
]

# =========================
# НЕДЕЛЯ
# =========================
def get_week_parity():
    """Return parity for the academic week to display.
    On Saturday/Sunday, show the following week's parity so students can
    check next week's timetable during the weekend.
    """
    today = datetime.now().date()
    week_date = today + timedelta(days=(7 - today.weekday())) if today.weekday() >= 5 else today

    # 01.09.2026 is treated as an odd academic week.
    semester_start = date(2026, 9, 1)
    week_monday = week_date - timedelta(days=week_date.weekday())
    start_monday = semester_start - timedelta(days=semester_start.weekday())
    week_index = (week_monday - start_monday).days // 7
    return "нечетная" if week_index % 2 == 0 else "четная"


def week_title():
    return "Чётная" if get_week_parity() == "четная" else "Нечётная"

# Для базы преподавателей используются английские значения week_type:
# "even" / "odd". Не меняем get_week_parity(), потому что обычное
# расписание групп использует русские значения "четная" / "нечетная".
def get_teacher_week_parity():
    return "even" if get_week_parity() == "четная" else "odd"

# =========================
# БАЗА ПРЕПОДАВАТЕЛЕЙ
# =========================
def ensure_teacher_db():
    if TEACHER_DB.exists() and TEACHER_DB.stat().st_size > 0:
        return
    if not TEACHER_SCHEDULES.exists():
        raise RuntimeError(f"Не найдена папка с расписаниями: {TEACHER_SCHEDULES}")
    print("[teacher-db] строю базу из Excel...")
    records = build_records(TEACHER_SCHEDULES)
    create_db(records, TEACHER_DB)
    print(f"[teacher-db] готово: {len(records)} записей")

# Строим БД только при отсутствии готовой базы.
ensure_teacher_db()

# =========================
# FSM ПОИСКА ПРЕПОДАВАТЕЛЯ
# =========================
class TeacherSearch(StatesGroup):
    waiting_surname = State()
    choosing_teacher = State()
    choosing_day = State()


@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer("Привет! Выбери нужный раздел в меню:", reply_markup=main_kb)


@dp.message(F.text.in_(stub_buttons))
async def handle_stubs(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(f'Раздел "{message.text}" находится в разработке 🛠')


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
        await message.answer(
            "❌ Преподаватель не найден. Проверьте написание фамилии и попробуйте ещё раз."
        )
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
        reply_markup=builder.as_markup(),
        parse_mode="HTML"
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
    await callback.message.edit_text(
        f"✅ Выбран преподаватель: <b>{html.escape(teacher)}</b>\n"
        f"Текущая неделя: <b>{week_title()}</b>\n\nВыберите день:",
        parse_mode="HTML"
    )
    builder = InlineKeyboardBuilder()
    for day in DAYS:
        builder.button(text=day, callback_data=f"teacher_day:{day}")
    builder.button(text="📅 Вся неделя", callback_data="teacher_day:ALL")
    builder.adjust(2)
    await callback.message.edit_reply_markup(reply_markup=builder.as_markup())
    await callback.answer()


def format_teacher_results(rows, teacher, selected_day):
    rows=[r for r in rows if is_active(r.get('valid_from'), r.get('valid_until'))]
    if not rows:
        return (
            f"❌ Для <b>{html.escape(teacher)}</b> занятий по выбранному фильтру не найдено.\n"
            f"Неделя: <b>{week_title()}</b>"
        )

    header = (
        f"👨‍🏫 <b>{html.escape(teacher)}</b>\n"
        f"Неделя: <b>{week_title()}</b>\n"
    )
    if selected_day != "ALL":
        header += f"День: <b>{html.escape(selected_day)}</b>\n"
    header += "\n"

    blocks=[]
    last_key=None
    for r in rows[:70]:
        day=r['day']
        if selected_day == "ALL" and day != last_key:
            blocks.append(f"\n📅 <b>{html.escape(day)}</b>")
            last_key=day
        pair=f"{r['lesson']} пара"
        subject=html.escape(r.get('subject') or 'Предмет не указан')
        group=html.escape(r['group_name'])
        room=html.escape(r.get('room') or 'не указана')
        institute=html.escape(institute_title(r.get('institute') or ''))
        course=r.get('course')
        course_text=("Магистратура" if r.get('degree')=="Магистратура" else (f"{course} курс" if course else ""))
        week=r.get('week_type')
        week_text={"odd":"нечётная","even":"чётная","date":"по дате"}.get(week,week or '')
        date_text=f" | {html.escape(r['date_text'])}" if r.get('date_text') else ""
        blocks.append(
            f"🕒 <b>{pair}</b> — {html.escape(day)}{date_text}\n"
            f"📚 {subject}\n"
            f"👥 {group}" + (f" ({html.escape(course_text)})" if course_text else "") + "\n"
            f"🚪 {room}\n"
            f"🏛 {institute}\n"
            f"🔄 {html.escape(week_text)}\n"
        )
    if len(rows)>70:
        blocks.append(f"\n…и ещё {len(rows)-70} записей")
    return header + "\n".join(blocks)


@dp.callback_query(TeacherSearch.choosing_day, F.data.startswith("teacher_day:"))
async def teacher_day(callback: types.CallbackQuery, state: FSMContext):
    data=await state.get_data()
    teacher=data.get('teacher')
    if not teacher:
        await callback.answer("Сначала выберите преподавателя", show_alert=True)
        return
    selected=callback.data.split(":",1)[1]
    day=None if selected=="ALL" else selected
    rows=search_teacher(teacher, day=day, week_type=get_teacher_week_parity(), db_path=TEACHER_DB)
    text=format_teacher_results(rows,teacher,selected)
    builder=InlineKeyboardBuilder()
    builder.button(text="🔙 К выбору дня", callback_data="teacher_back_days")
    builder.button(text="🔎 Новый преподаватель", callback_data="teacher_back_search")
    builder.adjust(1)
    await callback.message.edit_text(text,parse_mode="HTML",reply_markup=builder.as_markup())
    await callback.answer()
    # Не очищаем state: он нужен для кнопки «К выбору дня».


@dp.callback_query(F.data == "teacher_back_days")
async def teacher_back_days(callback: types.CallbackQuery, state: FSMContext):
    data=await state.get_data()
    teacher=data.get("teacher")
    if not teacher:
        await callback.answer("Преподаватель не выбран", show_alert=True)
        return
    await state.set_state(TeacherSearch.choosing_day)
    builder=InlineKeyboardBuilder()
    for day in DAYS:
        builder.button(text=day, callback_data=f"teacher_day:{day}")
    builder.button(text="📅 Вся неделя", callback_data="teacher_day:ALL")
    builder.adjust(2)
    await callback.message.edit_text(
        f"✅ Преподаватель: <b>{html.escape(teacher)}</b>\n\nВыберите день:",
        reply_markup=builder.as_markup(),
        parse_mode="HTML"
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
# РАСПИСАНИЕ ГРУПП: курс/магистратура -> институт -> группа -> день
# =========================
COURSE_OPTIONS = [
    ("1 курс", "1"), ("2 курс", "2"), ("3 курс", "3"),
    ("4 курс", "4"), ("5 курс", "5"), ("🎓 Магистратура", "master"),
]


def institute_title(institute: str) -> str:
    # В базе институт хранится как имя папки-источника ("расписание_ИТУР");
    # для кнопок показываем только короткое название.
    return institute.replace("расписание_", "", 1) if institute else institute


def course_to_degree_course(course_code: str):
    if course_code == "master":
        return "Магистратура", None
    return "Бакалавриат", int(course_code)


class StudentSchedule(StatesGroup):
    choosing_institute = State()
    choosing_group = State()
    choosing_day = State()


def course_keyboard():
    builder = InlineKeyboardBuilder()
    for label, code in COURSE_OPTIONS:
        builder.button(text=label, callback_data=f"std_course:{code}")
    builder.adjust(3)
    return builder.as_markup()


@dp.message(F.text == "Расписание")
async def cmd_schedule(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer("📚 Выберите курс:", reply_markup=course_keyboard())


@dp.callback_query(F.data == "std_back_course")
async def std_back_course(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    try:
        await callback.message.edit_text("📚 Выберите курс:", reply_markup=course_keyboard())
    except TelegramBadRequest:
        pass
    await callback.answer()


@dp.callback_query(F.data.startswith("std_course:"))
async def std_course(callback: types.CallbackQuery, state: FSMContext):
    code = callback.data.split(":", 1)[1]
    degree, course = course_to_degree_course(code)
    institutes = list_institutes(degree, course, TEACHER_DB)

    back_builder = InlineKeyboardBuilder()
    back_builder.button(text="🔙 К выбору курса", callback_data="std_back_course")
    if not institutes:
        try:
            await callback.message.edit_text(
                "❌ Для выбранного курса расписание пока не найдено.",
                reply_markup=back_builder.as_markup(),
            )
        except TelegramBadRequest:
            pass
        await callback.answer()
        return

    await state.update_data(degree=degree, course=course)
    await state.set_state(StudentSchedule.choosing_institute)

    builder = InlineKeyboardBuilder()
    for inst in institutes:
        builder.button(text=institute_title(inst), callback_data=f"std_institute:{inst}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору курса", callback_data="std_back_course"))
    course_title = "Магистратура" if course is None else f"{course} курс"
    try:
        await callback.message.edit_text(
            f"🏛 Курс: <b>{html.escape(course_title)}</b>\nВыберите институт:",
            reply_markup=builder.as_markup(),
            parse_mode="HTML",
        )
    except TelegramBadRequest:
        pass
    await callback.answer()


async def send_institute_keyboard(callback_or_message, state: FSMContext, edit: bool):
    data = await state.get_data()
    degree = data.get("degree")
    course = data.get("course")
    institutes = list_institutes(degree, course, TEACHER_DB)
    builder = InlineKeyboardBuilder()
    for inst in institutes:
        builder.button(text=institute_title(inst), callback_data=f"std_institute:{inst}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору курса", callback_data="std_back_course"))
    course_title = "Магистратура" if course is None else f"{course} курс"
    text = f"🏛 Курс: <b>{html.escape(course_title)}</b>\nВыберите институт:"
    if edit:
        try:
            await callback_or_message.message.edit_text(text, reply_markup=builder.as_markup(), parse_mode="HTML")
        except TelegramBadRequest:
            pass
    else:
        await callback_or_message.answer(text, reply_markup=builder.as_markup(), parse_mode="HTML")


@dp.callback_query(F.data == "std_back_institute")
async def std_back_institute(callback: types.CallbackQuery, state: FSMContext):
    await state.set_state(StudentSchedule.choosing_institute)
    await send_institute_keyboard(callback, state, edit=True)
    await callback.answer()


@dp.callback_query(F.data.startswith("std_institute:"))
async def std_institute(callback: types.CallbackQuery, state: FSMContext):
    institute = callback.data.split(":", 1)[1]
    data = await state.get_data()
    degree = data.get("degree")
    course = data.get("course")
    if not degree:
        await callback.answer("Сначала выберите курс", show_alert=True)
        return
    groups = list_groups(degree, institute, course, TEACHER_DB)

    if not groups:
        back_builder = InlineKeyboardBuilder()
        back_builder.button(text="🔙 К выбору института", callback_data="std_back_institute")
        try:
            await callback.message.edit_text(
                "❌ Для выбранного института группы не найдены.",
                reply_markup=back_builder.as_markup(),
            )
        except TelegramBadRequest:
            pass
        await callback.answer()
        return

    await state.update_data(institute=institute)
    await state.set_state(StudentSchedule.choosing_group)

    builder = InlineKeyboardBuilder()
    for g in groups:
        builder.button(text=g, callback_data=f"std_group:{g}")
    builder.adjust(3)
    builder.row(InlineKeyboardButton(text="🔙 К выбору института", callback_data="std_back_institute"))
    try:
        await callback.message.edit_text(
            f"🏛 Институт: <b>{html.escape(institute_title(institute))}</b>\nВыберите группу:",
            reply_markup=builder.as_markup(),
            parse_mode="HTML",
        )
    except TelegramBadRequest:
        pass
    await callback.answer()


def day_keyboard():
    builder = InlineKeyboardBuilder()
    for day in DAYS:
        builder.button(text=day, callback_data=f"std_day:{day}")
    builder.button(text="📅 Вся неделя", callback_data="std_day:ALL")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору группы", callback_data="std_back_group"))
    return builder.as_markup()


@dp.callback_query(F.data == "std_back_group")
async def std_back_group(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    degree = data.get("degree")
    course = data.get("course")
    institute = data.get("institute")
    if not institute:
        await callback.answer("Сначала выберите институт", show_alert=True)
        return
    groups = list_groups(degree, institute, course, TEACHER_DB)
    await state.set_state(StudentSchedule.choosing_group)
    builder = InlineKeyboardBuilder()
    for g in groups:
        builder.button(text=g, callback_data=f"std_group:{g}")
    builder.adjust(3)
    builder.row(InlineKeyboardButton(text="🔙 К выбору института", callback_data="std_back_institute"))
    try:
        await callback.message.edit_text(
            f"🏛 Институт: <b>{html.escape(institute_title(institute))}</b>\nВыберите группу:",
            reply_markup=builder.as_markup(),
            parse_mode="HTML",
        )
    except TelegramBadRequest:
        pass
    await callback.answer()


@dp.callback_query(F.data.startswith("std_group:"))
async def std_group(callback: types.CallbackQuery, state: FSMContext):
    group_name = callback.data.split(":", 1)[1]
    await state.update_data(group=group_name)
    await state.set_state(StudentSchedule.choosing_day)
    try:
        await callback.message.edit_text(
            f"✅ Группа: <b>{html.escape(group_name)}</b>\n"
            f"Текущая неделя: <b>{week_title()}</b>\n\nВыберите день:",
            reply_markup=day_keyboard(),
            parse_mode="HTML",
        )
    except TelegramBadRequest:
        pass
    await callback.answer()


def format_group_results(rows, group_name, selected_day):
    rows=[r for r in rows if is_active(r.get('valid_from'), r.get('valid_until'))]
    if not rows:
        if selected_day != "ALL":
            return (
                f"🎓 <b>{html.escape(group_name)}</b>\n"
                f"Неделя: <b>{week_title()}</b>\n"
                f"День: <b>{html.escape(selected_day)}</b>\n\n"
                f"📭 В этот день занятий нет."
            )
        return (
            f"❌ Для группы <b>{html.escape(group_name)}</b> занятий по выбранному фильтру не найдено.\n"
            f"Неделя: <b>{week_title()}</b>"
        )

    institute = institute_title(rows[0].get("institute") or "")
    header = (
        f"🎓 <b>{html.escape(group_name)}</b>"
        + (f" ({html.escape(institute)})" if institute else "")
        + "\n"
        f"Неделя: <b>{week_title()}</b>\n"
    )
    if selected_day != "ALL":
        header += f"День: <b>{html.escape(selected_day)}</b>\n"
    header += "\n"

    # Parallel subgroups (several teachers/rooms for the exact same
    # day+pair+week) are stored as separate rows so each teacher stays
    # individually searchable, but for a group's own timetable they belong
    # together as one pair — merge them here rather than repeating the same
    # time slot two or three times in a row.
    merged=[]
    for r in rows[:200]:
        key=(r['day'], r['lesson'], r.get('week_type'), r.get('subject'))
        if merged and merged[-1]['_key']==key:
            slot=merged[-1]
            t=r.get('teacher')
            if t and t not in slot['_teachers']: slot['_teachers'].append(t)
            rm=r.get('room')
            if rm and rm not in slot['_rooms']: slot['_rooms'].append(rm)
        else:
            merged.append({'_key':key, 'day':r['day'], 'lesson':r['lesson'], 'week_type':r.get('week_type'),
                           'subject':r.get('subject'),
                           '_teachers':[r['teacher']] if r.get('teacher') else [],
                           '_rooms':[r['room']] if r.get('room') else []})

    blocks = []
    last_day = None
    for slot in merged[:70]:
        day = slot["day"]
        if selected_day == "ALL" and day != last_day:
            blocks.append(f"\n📅 <b>{html.escape(day)}</b>")
            last_day = day
        pair = f"{slot['lesson']} пара"
        subject = html.escape(slot.get("subject") or "Предмет не указан")
        week = slot.get("week_type")
        week_text = {"odd": "нечётная", "even": "чётная", "date": "по дате"}.get(week, week or "")
        lines=[f"🕒 <b>{pair}</b> — {html.escape(day)}\n📚 {subject}\n"]
        if slot['_teachers']:
            lines.append(f"👨‍🏫 {html.escape('; '.join(slot['_teachers']))}\n")
        if slot['_rooms']:
            lines.append(f"🚪 {html.escape(', '.join(slot['_rooms']))}\n")
        lines.append(f"🔄 {html.escape(week_text)}\n")
        blocks.append("".join(lines))
    if len(merged) > 70:
        blocks.append(f"\n…и ещё {len(merged) - 70} записей")
    return header + "\n".join(blocks)


@dp.callback_query(StudentSchedule.choosing_day, F.data.startswith("std_day:"))
async def std_day(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    group_name = data.get("group")
    institute = data.get("institute")
    if not group_name:
        await callback.answer("Сначала выберите группу", show_alert=True)
        return
    selected = callback.data.split(":", 1)[1]
    day = None if selected == "ALL" else selected
    rows = search_group(group_name, institute=institute, day=day, week_type=get_teacher_week_parity(), db_path=TEACHER_DB)
    text = format_group_results(rows, group_name, selected)
    builder = InlineKeyboardBuilder()
    builder.button(text="🔙 К выбору дня", callback_data="std_back_days")
    builder.button(text="🔙 К выбору группы", callback_data="std_back_group")
    builder.adjust(1)
    try:
        await callback.message.edit_text(text, parse_mode="HTML", reply_markup=builder.as_markup())
    except TelegramBadRequest:
        pass
    await callback.answer()


@dp.callback_query(F.data == "std_back_days")
async def std_back_days(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    group_name = data.get("group")
    if not group_name:
        await callback.answer("Группа не выбрана", show_alert=True)
        return
    await state.set_state(StudentSchedule.choosing_day)
    try:
        await callback.message.edit_text(
            f"✅ Группа: <b>{html.escape(group_name)}</b>\n\nВыберите день:",
            reply_markup=day_keyboard(),
            parse_mode="HTML",
        )
    except TelegramBadRequest:
        pass
    await callback.answer()


# =========================
# FASTAPI / RENDER
# =========================
app=FastAPI()

@app.get("/")
def index():
    return "OK"


def teacher_record_count():
    try:
        con=sqlite3.connect(TEACHER_DB)
        n=con.execute("SELECT COUNT(*) FROM lessons").fetchone()[0]
        con.close(); return n
    except Exception:
        return 0


async def main():
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot, handle_signals=False)


def run_bot_in_thread():
    asyncio.run(main())


if __name__ == "__main__":
    t=threading.Thread(target=run_bot_in_thread,daemon=True)
    t.start()
    uvicorn.run(app,host="0.0.0.0",port=int(os.getenv("PORT","7860")))
