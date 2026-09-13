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

from teacher_schedule_parser import build_records, create_db, list_teachers, search_teacher, list_groups, get_group_schedule

# =========================
# НАСТРОЙКИ
# =========================
TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    raise RuntimeError("Не задана переменная окружения BOT_TOKEN")

SCHEDULE_DB = Path(os.getenv("SCHEDULE_DB", os.getenv("TEACHER_DB", "schedule.db")))
TEACHER_DB = SCHEDULE_DB
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
def ensure_schedule_db():
    if SCHEDULE_DB.exists() and SCHEDULE_DB.stat().st_size > 0:
        return
    if not TEACHER_SCHEDULES.exists():
        raise RuntimeError(f"Не найдена папка с расписаниями: {TEACHER_SCHEDULES}")
    print('[schedule-db] строю единую базу из Excel...')
    records=build_records(TEACHER_SCHEDULES)
    create_db(records,SCHEDULE_DB)
    print(f'[schedule-db] готово: {len(records)} записей')

ensure_schedule_db()

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
        room=html.escape(r.get('room') or 'аудитория не указана')
        institute=html.escape(r.get('institute') or '')
        week=r.get('week_type')
        week_text={"odd":"нечётная","even":"чётная","date":"по дате"}.get(week,week or '')
        dates=r.get('date_ranges') or []
        date_text=""
        if r.get('date_text'):
            date_text=f" | {html.escape(r['date_text'])}"
        elif dates:
            date_text=f" | {html.escape(', '.join(dates))}"
        blocks.append(
            f"🕒 <b>{pair}</b> — {html.escape(day)}{date_text}\n"
            f"📚 {subject}\n"
            f"👥 {group}\n"
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
# РАСПИСАНИЕ ГРУПП — ТА ЖЕ ЕДИНАЯ БД
# =========================
PAIR_TIMES={
    1:'08:30–09:50', 2:'10:00–11:30', 3:'12:00–13:30',
    4:'13:50–15:20', 5:'15:30–17:00', 6:'17:10–18:40'
}

def group_keyboard(groups=None):
    groups=groups if groups is not None else list_groups(SCHEDULE_DB)
    if not groups: return None
    builder=InlineKeyboardBuilder()
    for group in groups:
        builder.button(text=group,callback_data=f'group_{group}')
    builder.adjust(2)
    return builder.as_markup()

@dp.message(F.text == 'Расписание')
async def cmd_schedule(message: types.Message, state: FSMContext):
    await state.clear()
    keyboard=group_keyboard()
    if not keyboard:
        await message.answer('В базе не найдено расписание групп.')
        return
    await message.answer('📚 Выберите вашу группу:',reply_markup=keyboard)

@dp.callback_query(F.data == 'back_to_groups')
async def process_back_to_groups(callback: types.CallbackQuery):
    keyboard=group_keyboard()
    try: await callback.message.edit_text('📚 Выберите вашу группу:',reply_markup=keyboard)
    except TelegramBadRequest: pass
    await callback.answer()

@dp.callback_query(F.data.startswith('group_'))
async def process_group(callback: types.CallbackQuery):
    group_name=callback.data.split('_',1)[1]
    builder=InlineKeyboardBuilder()
    for day in DAYS: builder.button(text=day,callback_data=f'day_{group_name}_{day}')
    builder.adjust(2); builder.row(InlineKeyboardButton(text='🔙 К выбору групп',callback_data='back_to_groups'))
    try: await callback.message.edit_text(f'🎓 Группа: <b>{html.escape(group_name)}</b>\nВыберите день недели:',reply_markup=builder.as_markup(),parse_mode='HTML')
    except TelegramBadRequest: pass
    await callback.answer()

@dp.callback_query(F.data.startswith('back_to_days_'))
async def process_back_to_days(callback: types.CallbackQuery):
    group_name=callback.data.split('_',3)[3]
    builder=InlineKeyboardBuilder()
    for day in DAYS: builder.button(text=day,callback_data=f'day_{group_name}_{day}')
    builder.adjust(2); builder.row(InlineKeyboardButton(text='🔙 К выбору групп',callback_data='back_to_groups'))
    try: await callback.message.edit_text(f'🎓 Группа: <b>{html.escape(group_name)}</b>\nВыберите день недели:',reply_markup=builder.as_markup(),parse_mode='HTML')
    except TelegramBadRequest: pass
    await callback.answer()

@dp.callback_query(F.data.startswith('day_'))
async def process_day(callback: types.CallbackQuery):
    parts=callback.data.split('_',2); group_name=parts[1]; day_name=parts[2]
    parity=get_week_parity(); parity_title=week_title()
    rows=get_group_schedule(group_name,day=day_name,week_type=parity,db_path=SCHEDULE_DB)
    builder=InlineKeyboardBuilder(); builder.row(InlineKeyboardButton(text='🔙 К выбору дней',callback_data=f'back_to_days_{group_name}'))
    response_text=f'📅 <b>Расписание на {html.escape(day_name)}</b>\n🎓 Группа: {html.escape(group_name)} | Неделя: <b>{parity_title}</b>\n\n'
    if not rows:
        response_text+='🎉 В этот день пар нет!'
    else:
        for r in rows:
            subject=r.get('subject') or 'Предмет не указан'
            response_text+=f"🕒 <b>{r['lesson']} пара</b> ({PAIR_TIMES.get(r['lesson'],'')})\n"
            response_text+=f"📚 <b>{html.escape(subject)}</b>\n"
            teachers=r.get('teachers') or []
            if teachers: response_text+=f"👨‍🏫 {html.escape(', '.join(teachers))}\n"
            if r.get('room'): response_text+=f"🚪 ауд. {html.escape(r['room'])}\n"
            response_text+='\n'
    try: await callback.message.edit_text(response_text,reply_markup=builder.as_markup(),parse_mode='HTML')
    except TelegramBadRequest: pass
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
