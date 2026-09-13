# -*- coding: utf-8 -*-
import asyncio, html, os, sqlite3, threading
from datetime import datetime, timedelta, date
from pathlib import Path
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
from teacher_schedule_parser import build_records, create_db, list_teachers, search_teacher, list_institutes, list_groups, search_group

TOKEN=os.getenv('BOT_TOKEN')
if not TOKEN: raise RuntimeError('Не задана переменная окружения BOT_TOKEN')
TEACHER_DB=Path(os.getenv('TEACHER_DB','teacher_schedule.db'))
TEACHER_SCHEDULES=Path(os.getenv('TEACHER_SCHEDULES','teacher_schedules'))
bot=Bot(token=TOKEN); dp=Dispatcher(storage=MemoryStorage())
DAYS=['Понедельник','Вторник','Среда','Четверг','Пятница','Суббота']
main_kb=ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text='Расписание'),KeyboardButton(text='👨‍🏫 Найти преподавателя')],[KeyboardButton(text='Навигация по корпусам'),KeyboardButton(text='Куда обратиться?')],[KeyboardButton(text='Учёба и сессия'),KeyboardButton(text='Стипендии и соцподдержка')],[KeyboardButton(text='Общежития'),KeyboardButton(text='Контакты')],[KeyboardButton(text='Мероприятия')]],resize_keyboard=True)
stub_buttons=['Навигация по корпусам','Куда обратиться?','Учёба и сессия','Стипендии и соцподдержка','Общежития','Контакты','Мероприятия']

def get_week_parity():
    today=datetime.now().date(); week_date=today+timedelta(days=(7-today.weekday())) if today.weekday()>=5 else today
    semester_start=date(2026,9,1); week_monday=week_date-timedelta(days=week_date.weekday()); start_monday=semester_start-timedelta(days=semester_start.weekday())
    return 'нечетная' if ((week_monday-start_monday).days//7)%2==0 else 'четная'
def week_title(): return 'Чётная' if get_week_parity()=='четная' else 'Нечётная'
def get_teacher_week_parity(): return 'even' if get_week_parity()=='четная' else 'odd'

def ensure_db():
    rebuild=not TEACHER_DB.exists() or TEACHER_DB.stat().st_size==0
    if not rebuild:
        try:
            con=sqlite3.connect(TEACHER_DB); tables={r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}; con.close(); rebuild='group_lessons' not in tables or 'lessons' not in tables
        except Exception: rebuild=True
    if rebuild:
        if not TEACHER_SCHEDULES.exists(): raise RuntimeError(f'Не найдена папка с расписаниями: {TEACHER_SCHEDULES}')
        print('[schedule-db] строю базу из Excel...')
        create_db(build_records(TEACHER_SCHEDULES),TEACHER_DB)
ensure_db()

# ---------------- teacher search ----------------
class TeacherSearch(StatesGroup): waiting_surname=State(); choosing_teacher=State(); choosing_day=State()
@dp.message(Command('start'))
async def cmd_start(message,state): await state.clear(); await message.answer('Привет! Выбери нужный раздел в меню:',reply_markup=main_kb)
@dp.message(F.text.in_(stub_buttons))
async def handle_stubs(message,state): await state.clear(); await message.answer(f'Раздел "{message.text}" находится в разработке 🛠')
@dp.message(F.text=='👨‍🏫 Найти преподавателя')
async def teacher_start(message,state): await state.set_state(TeacherSearch.waiting_surname); await message.answer('👨‍🏫 <b>Поиск преподавателя</b>\n\nВведите фамилию преподавателя. Можно написать только фамилию или фамилию с инициалами.',parse_mode='HTML')
async def teacher_days(message,state):
    data=await state.get_data(); teacher=data.get('teacher',''); b=InlineKeyboardBuilder()
    for d in DAYS:b.button(text=d,callback_data=f'teacher_day:{d}')
    b.button(text='📅 Вся неделя',callback_data='teacher_day:ALL'); b.adjust(2)
    await message.answer(f'✅ Выбран преподаватель: <b>{html.escape(teacher)}</b>\nТекущая неделя: <b>{week_title()}</b>\n\nВыберите день:',reply_markup=b.as_markup(),parse_mode='HTML')
@dp.message(TeacherSearch.waiting_surname)
async def teacher_surname(message,state):
    q=(message.text or '').strip(); candidates=list_teachers(q,TEACHER_DB)
    if not candidates: await message.answer('❌ Преподаватель не найден. Проверьте написание фамилии и попробуйте ещё раз.'); return
    await state.update_data(query=q,candidates=candidates)
    if len(candidates)==1:
        await state.update_data(teacher=candidates[0]); await state.set_state(TeacherSearch.choosing_day); await teacher_days(message,state); return
    b=InlineKeyboardBuilder()
    for i,t in enumerate(candidates[:20]):b.button(text=t,callback_data=f'teacher_select:{i}')
    b.adjust(1); await state.set_state(TeacherSearch.choosing_teacher); await message.answer(f'Нашёл несколько совпадений по запросу <b>{html.escape(q)}</b>.\nВыберите нужного преподавателя:',reply_markup=b.as_markup(),parse_mode='HTML')
@dp.callback_query(TeacherSearch.choosing_teacher,F.data.startswith('teacher_select:'))
async def teacher_select(callback,state):
    data=await state.get_data(); cs=data.get('candidates',[]); idx=int(callback.data.split(':',1)[1])
    if idx>=len(cs): await callback.answer('Преподаватель не найден',show_alert=True); return
    await state.update_data(teacher=cs[idx]); await state.set_state(TeacherSearch.choosing_day); await callback.message.edit_text(f'✅ Выбран преподаватель: <b>{html.escape(cs[idx])}</b>\nТекущая неделя: <b>{week_title()}</b>\n\nВыберите день:',parse_mode='HTML')
    b=InlineKeyboardBuilder()
    for d in DAYS:b.button(text=d,callback_data=f'teacher_day:{d}')
    b.button(text='📅 Вся неделя',callback_data='teacher_day:ALL');b.adjust(2);await callback.message.edit_reply_markup(reply_markup=b.as_markup());await callback.answer()
def format_teacher(rows,teacher,selected):
    if not rows:return f'❌ Для <b>{html.escape(teacher)}</b> занятий по выбранному фильтру не найдено.\nНеделя: <b>{week_title()}</b>'
    out=f'👨‍🏫 <b>{html.escape(teacher)}</b>\nНеделя: <b>{week_title()}</b>\n'
    if selected!='ALL':out+=f'День: <b>{html.escape(selected)}</b>\n'
    out+='\n'; last=None
    for r in rows[:70]:
        if selected=='ALL' and r['day']!=last:out+=f"\n📅 <b>{html.escape(r['day'])}</b>\n";last=r['day']
        periods=r.get('lesson_periods') or []; period=(' | '+', '.join(periods)) if periods else ''
        dates=r.get('date_ranges') or []; datepart=(' | '+', '.join(dates)) if not period and dates else period
        out+=f"🕒 <b>{r['lesson']} пара</b> — {html.escape(r['day'])}{html.escape(datepart)}\n📚 {html.escape(r.get('subject') or 'Предмет не указан')}\n👥 {html.escape(r['group_name'])}\n"
        if r.get('room'):out+=f"🚪 {html.escape(r['room'])}\n"
        if r.get('institute'):out+=f"🏛 {html.escape(r['institute'])}\n"
        out+=f"🔄 { {'odd':'нечётная','even':'чётная','date':'по дате'}.get(r.get('week_type'),r.get('week_type',''))}\n\n"
    if len(rows)>70:out+=f'…и ещё {len(rows)-70} записей'
    return out
@dp.callback_query(TeacherSearch.choosing_day,F.data.startswith('teacher_day:'))
async def teacher_day(callback,state):
    data=await state.get_data();teacher=data.get('teacher');sel=callback.data.split(':',1)[1];day=None if sel=='ALL' else sel
    rows=search_teacher(teacher,day=day,week_type=get_teacher_week_parity(),db_path=TEACHER_DB);b=InlineKeyboardBuilder();b.button(text='🔙 К выбору дня',callback_data='teacher_back_days');b.button(text='🔎 Новый преподаватель',callback_data='teacher_back_search');b.adjust(1)
    await callback.message.edit_text(format_teacher(rows,teacher,sel),parse_mode='HTML',reply_markup=b.as_markup());await callback.answer()
@dp.callback_query(F.data=='teacher_back_days')
async def teacher_back_days(callback,state): await state.set_state(TeacherSearch.choosing_day);await teacher_days(callback.message,state);await callback.answer()
@dp.callback_query(F.data=='teacher_back_search')
async def teacher_back_search(callback,state): await state.clear();await state.set_state(TeacherSearch.waiting_surname);await callback.message.edit_text('👨‍🏫 <b>Поиск преподавателя</b>\n\nВведите фамилию преподавателя.',parse_mode='HTML');await callback.answer()

# ---------------- group schedule ----------------
async def schedule_course_keyboard():
    con=sqlite3.connect(TEACHER_DB); rows=con.execute('SELECT DISTINCT level,course FROM group_lessons ORDER BY level,course').fetchall();con.close();b=InlineKeyboardBuilder()
    courses=sorted({r[1] for r in rows if r[0]=='Бакалавриат' and r[1] is not None})
    for c in courses:b.button(text=f'{c} курс',callback_data=f'sch_course:{c}')
    if any(r[0]=='Магистратура' for r in rows):b.button(text='🎓 Магистратура',callback_data='sch_master')
    b.adjust(2);return b.as_markup()
@dp.message(F.text=='Расписание')
async def cmd_schedule(message,state): await state.clear();await message.answer('📚 <b>Выберите курс:</b>',reply_markup=await schedule_course_keyboard(),parse_mode='HTML')
async def institute_kb(level,course):
    ins=list_institutes(level,course,TEACHER_DB);b=InlineKeyboardBuilder()
    for i in ins:b.button(text=i,callback_data=f'sch_inst:{i}:{course if course is not None else "M"}')
    b.button(text='🔙 К выбору курса',callback_data='sch_back_courses');b.adjust(2);return b.as_markup()
@dp.callback_query(F.data.startswith('sch_course:'))
async def sch_course(callback,state):
    c=int(callback.data.split(':')[1]);await state.update_data(level='Бакалавриат',course=c);b=await institute_kb('Бакалавриат',c);await callback.message.edit_text(f'📚 <b>{c} курс</b>\n\nВыберите институт:',reply_markup=b,parse_mode='HTML');await callback.answer()
@dp.callback_query(F.data=='sch_master')
async def sch_master(callback,state): await state.update_data(level='Магистратура',course=None);b=await institute_kb('Магистратура',None);await callback.message.edit_text('🎓 <b>Магистратура</b>\n\nВыберите институт:',reply_markup=b,parse_mode='HTML');await callback.answer()
@dp.callback_query(F.data=='sch_back_courses')
async def sch_back_courses(callback,state): await callback.message.edit_text('📚 <b>Выберите курс:</b>',reply_markup=await schedule_course_keyboard(),parse_mode='HTML');await callback.answer()
@dp.callback_query(F.data.startswith('sch_inst:'))
async def sch_inst(callback,state):
    _,inst,c=callback.data.split(':',2);level=(await state.get_data()).get('level','Бакалавриат');course=None if c=='M' else int(c);await state.update_data(institute=inst,course=course)
    groups=list_groups(level,course,inst,TEACHER_DB);b=InlineKeyboardBuilder()
    for g in groups:b.button(text=g,callback_data=f'sch_group:{g}')
    b.button(text='🔙 К институтам',callback_data='sch_back_inst');b.adjust(2)
    await callback.message.edit_text(f'🏛 <b>{html.escape(inst)}</b>\n\nВыберите группу:',reply_markup=b.as_markup(),parse_mode='HTML');await callback.answer()
@dp.callback_query(F.data=='sch_back_inst')
async def sch_back_inst(callback,state):
    d=await state.get_data();b=await institute_kb(d.get('level','Бакалавриат'),d.get('course'));await callback.message.edit_text('Выберите институт:',reply_markup=b);await callback.answer()
@dp.callback_query(F.data.startswith('sch_group:'))
async def sch_group(callback,state):
    g=callback.data.split(':',1)[1];await state.update_data(group=g);b=InlineKeyboardBuilder()
    for d in DAYS:b.button(text=d,callback_data=f'sch_day:{d}')
    b.button(text='🔙 К выбору группы',callback_data='sch_back_group');b.adjust(2)
    await callback.message.edit_text(f'🎓 <b>Группа: {html.escape(g)}</b>\n\nВыберите день недели:',reply_markup=b.as_markup(),parse_mode='HTML');await callback.answer()
@dp.callback_query(F.data=='sch_back_group')
async def sch_back_group(callback,state):
    d=await state.get_data();level=d.get('level');course=d.get('course');inst=d.get('institute');groups=list_groups(level,course,inst,TEACHER_DB);b=InlineKeyboardBuilder()
    for g in groups:b.button(text=g,callback_data=f'sch_group:{g}')
    b.button(text='🔙 К институтам',callback_data='sch_back_inst');b.adjust(2);await callback.message.edit_text(f'🏛 <b>{html.escape(inst)}</b>\n\nВыберите группу:',reply_markup=b.as_markup(),parse_mode='HTML');await callback.answer()

def format_group(rows,group,day):
    if not rows:return f'🎉 У группы <b>{html.escape(group)}</b> в {html.escape(day.lower())} пар нет.\nНеделя: <b>{week_title()}</b>'
    out=f'📅 <b>Расписание на {html.escape(day)}</b>\n🎓 Группа: <b>{html.escape(group)}</b>\nНеделя: <b>{week_title()}</b>\n\n';times={1:'08:30–09:50',2:'10:00–11:30',3:'12:00–13:30',4:'13:50–15:20',5:'15:30–17:00',6:'17:10–18:40'}
    for r in rows:
        periods=r.get('lesson_periods') or []; suffix=(' | '+', '.join(periods)) if periods else ''
        out+=f"🕒 <b>{r['lesson']} пара</b> ({times.get(r['lesson'],'')}){html.escape(suffix)}\n"
        if r.get('subject'):out+=f"📚 <b>{html.escape(r['subject'])}</b>\n"
        if r.get('teacher'):out+=f"👨‍🏫 {html.escape(r['teacher'])}\n"
        if r.get('room'):out+=f"🚪 ауд. {html.escape(r['room'])}\n"
        out+='\n'
    return out
@dp.callback_query(F.data.startswith('sch_day:'))
async def sch_day(callback,state):
    d=await state.get_data();day=callback.data.split(':',1)[1];rows=search_group(d['group'],day,get_teacher_week_parity(),TEACHER_DB);b=InlineKeyboardBuilder();b.button(text='🔙 К выбору дней',callback_data='sch_back_days');b.button(text='🔙 К выбору группы',callback_data='sch_back_group');b.adjust(1);await callback.message.edit_text(format_group(rows,d['group'],day),parse_mode='HTML',reply_markup=b.as_markup());await callback.answer()
@dp.callback_query(F.data=='sch_back_days')
async def sch_back_days(callback,state):
    d=await state.get_data();b=InlineKeyboardBuilder()
    for day in DAYS:b.button(text=day,callback_data=f'sch_day:{day}')
    b.button(text='🔙 К выбору группы',callback_data='sch_back_group');b.adjust(2);await callback.message.edit_text(f'🎓 <b>Группа: {html.escape(d.get("group", ""))}</b>\n\nВыберите день недели:',reply_markup=b.as_markup(),parse_mode='HTML');await callback.answer()

app=FastAPI()
@app.get('/')
def index():return 'OK'
async def main(): await bot.delete_webhook(drop_pending_updates=True);await dp.start_polling(bot,handle_signals=False)
def run_bot_in_thread():asyncio.run(main())
if __name__=='__main__':
    threading.Thread(target=run_bot_in_thread,daemon=True).start();uvicorn.run(app,host='0.0.0.0',port=int(os.getenv('PORT','7860')))
