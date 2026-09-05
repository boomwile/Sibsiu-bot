# -*- coding: utf-8 -*-
import asyncio
import json
import os
from datetime import datetime
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.exceptions import TelegramBadRequest
from fastapi import FastAPI
import uvicorn
import threading

# Сюда вставь свой токен бота от BotFather
TOKEN = "8739289680:AAFSOC8mmYuoze3Q08u_jDgiBcbKTlmFxMI"

bot = Bot(token=TOKEN)
dp = Dispatcher()

# Главное меню (8 кнопок)
main_kb = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Расписание"), KeyboardButton(text="Навигация по корпусам")],
        [KeyboardButton(text="Куда обратиться?"), KeyboardButton(text="Учёба и сессия")],
        [KeyboardButton(text="Стипендии и соцподдержка"), KeyboardButton(text="Общежития")],
        [KeyboardButton(text="Контакты"), KeyboardButton(text="Мероприятия")]
    ],
    resize_keyboard=True
)

def get_week_parity():
    now = datetime.now()
    year = now.year if now.month >= 9 else now.year - 1
    sept_1 = datetime(year, 9, 1)
    delta_days = (now - sept_1).days
    week_number = (delta_days // 7) + 1
    return "четная" if week_number % 2 == 0 else "нечетная"

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    await message.answer("Привет! Выбери нужный раздел в меню:", reply_markup=main_kb)

stub_buttons = [
    "Навигация по корпусам", "Куда обратиться?", "Учёба и сессия",
    "Стипендии и соцподдержка", "Общежития", "Контакты", "Мероприятия"
]

@dp.message(F.text.in_(stub_buttons))
async def handle_stubs(message: types.Message):
    await message.answer(f'Раздел "{message.text}" находится в разработке 🛠')

async def get_groups_keyboard():
    schedules_dir = "schedules"
    if not os.path.exists(schedules_dir):
        os.makedirs(schedules_dir, exist_ok=True)
        return None

    groups = []
    for filename in os.listdir(schedules_dir):
        if filename.endswith(".json"):
            path = os.path.join(schedules_dir, filename)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for group_name in data.keys():
                        groups.append(group_name)
            except Exception as e:
                print(f"Ошибка чтения {filename}: {e}")

    if not groups:
        return None

    builder = InlineKeyboardBuilder()
    for group in groups:
        builder.button(text=group, callback_data=f"group_{group}")
    builder.adjust(2)
    return builder.as_markup()

@dp.message(F.text == "Расписание")
async def cmd_schedule(message: types.Message):
    keyboard = await get_groups_keyboard()
    if not keyboard:
        await message.answer("В папке schedules не найдено файлов с расписанием.")
        return
    await message.answer("📚 Выберите вашу группу:", reply_markup=keyboard)

@dp.callback_query(F.data == "back_to_groups")
async def process_back_to_groups(callback: types.CallbackQuery):
    keyboard = await get_groups_keyboard()
    try:
        await callback.message.edit_text("📚 Выберите вашу группу:", reply_markup=keyboard)
    except TelegramBadRequest:
        pass
    await callback.answer()

@dp.callback_query(F.data.startswith("group_"))
async def process_group(callback: types.CallbackQuery):
    group_name = callback.data.split("_", 1)[1]
    days = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
    builder = InlineKeyboardBuilder()
    for day in days:
        builder.button(text=day, callback_data=f"day_{group_name}_{day}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору групп", callback_data="back_to_groups"))

    try:
        await callback.message.edit_text(
            f"🎓 Группа: <b>{group_name}</b>\nВыберите день недели:",
            reply_markup=builder.as_markup(),
            parse_mode="HTML"
        )
    except TelegramBadRequest:
        pass
    await callback.answer()

@dp.callback_query(F.data.startswith("back_to_days_"))
async def process_back_to_days(callback: types.CallbackQuery):
    group_name = callback.data.split("_", 3)[3]
    days = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
    builder = InlineKeyboardBuilder()
    for day in days:
        builder.button(text=day, callback_data=f"day_{group_name}_{day}")
    builder.adjust(2)
    builder.row(InlineKeyboardButton(text="🔙 К выбору групп", callback_data="back_to_groups"))

    try:
        await callback.message.edit_text(
            f"🎓 Группа: <b>{group_name}</b>\nВыберите день недели:",
            reply_markup=builder.as_markup(),
            parse_mode="HTML"
        )
    except TelegramBadRequest:
        pass
    await callback.answer()

@dp.callback_query(F.data.startswith("day_"))
async def process_day(callback: types.CallbackQuery):
    parts = callback.data.split("_", 2)
    group_name = parts[1]
    day_name = parts[2]

    parity = get_week_parity()
    parity_title = "Чётная" if parity == "четная" else "Нечётная"

    schedules_dir = "schedules"
    schedule_data = None
    for filename in os.listdir(schedules_dir):
        if filename.endswith(".json"):
            path = os.path.join(schedules_dir, filename)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if group_name in data:
                        schedule_data = data[group_name]
                        break
            except:
                pass

    builder = InlineKeyboardBuilder()
    builder.row(InlineKeyboardButton(text="🔙 К выбору дней", callback_data=f"back_to_days_{group_name}"))

    if not schedule_data or day_name not in schedule_data:
        try:
            await callback.message.edit_text(
                f"❄1�7 Расписание для группы {group_name} на {day_name} не найдено.",
                reply_markup=builder.as_markup()
            )
        except TelegramBadRequest:
            pass
        await callback.answer()
        return

    day_schedule = schedule_data[day_name]
    response_text = f"📅 <b>Расписание на {day_name}</b>\n" \
                    f"🎓 Группа: {group_name} | Неделя: <b>{parity_title}</b>\n\n"

    has_lessons = False
    for pair_name, pair_info in day_schedule.items():
        time_str = pair_info.get("время", "")
        block = pair_info.get(parity)
        if not block or not block.get("предмет"):
            continue

        subject = block.get("предмет", "")
        teacher = block.get("преподаватель", "")
        auditory = block.get("аудитория", "")

        if not subject:
            continue

        has_lessons = True
        response_text += f"🕒 <b>{pair_name}</b> ({time_str})\n" \
                         f"📚 <b>{subject}</b>\n"
        if teacher:
            response_text += f"👨‍🏄1�7 {teacher}\n"
        if auditory:
            response_text += f"🚪 ауд. {auditory}\n"
        response_text += "\n"

    if not has_lessons:
        response_text += "🎉 В этот день пар нет!"

    try:
        await callback.message.edit_text(
            response_text,
            reply_markup=builder.as_markup(),
            parse_mode="HTML"
        )
    except TelegramBadRequest:
        pass
    await callback.answer()

app = FastAPI()

@app.get("/")
def index():
    return {"status": "Bot is running 24/7!"}

async def main():
    print("Бот запущен на Hugging Face!")
async def main():
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot, handle_signals=False)

def run_bot_in_thread():
    asyncio.run(main())

if __name__ == "__main__":
    t = threading.Thread(target=run_bot_in_thread, daemon=True)
    t.start()
    uvicorn.run(app, host="0.0.0.0", port=7860)
