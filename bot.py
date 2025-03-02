#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import logging
import os
import json
import asyncio
import datetime
import re
from urllib.parse import quote
from datetime import datetime
from zoneinfo import ZoneInfo

from aiogram import Bot, Dispatcher, types, executor
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Text, Regexp
from aiogram.dispatcher.filters.state import State, StatesGroup

from aiohttp import web
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# бібліотека для форматування клітинок
from gspread_formatting import (
    format_cell_range,
    cellFormat,
    Color,
    set_frozen,
    set_column_width
)
from gspread.utils import rowcol_to_a1

############################################
# 1) ЧИТАЄМО ЗМІННІ ОТОЧЕННЯ ЗАМІСТЬ .env ТА credentials.json
############################################

logging.basicConfig(
    level=logging.DEBUG,
    format='[%(asctime)s] [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

############################################
# 2) ОГОЛОШЕННЯ ЗМІННИХ ОТОЧЕННЯ
############################################

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "")
ADMINS = [uid.strip() for uid in os.getenv("ADMINS", "").split(",") if uid.strip()]

GOOGLE_SPREADSHEET_ID = os.getenv("GOOGLE_SPREADSHEET_ID", "")
SHEET1_NAME = os.getenv("SHEET1_NAME", "Лист1")

GOOGLE_SPREADSHEET_ID2 = os.getenv("GOOGLE_SPREADSHEET_ID2", "")
SHEET2_NAME = os.getenv("SHEET2_NAME", "Лист1")

CHECK_INTERVAL = int(os.getenv("CHECK_INTERVAL_SECONDS", "60"))
API_PORT = int(os.getenv("API_PORT", "8080"))

############################################
# 3) ЗАВАНТАЖУЄМО credentials.json ЗІ ЗМІННОЇ ENV
############################################

import_base64 = False
GSPREAD_CREDENTIALS_JSON = os.getenv("GSPREAD_CREDENTIALS_JSON", "")
if not GSPREAD_CREDENTIALS_JSON:
    raise RuntimeError("Немає GSPREAD_CREDENTIALS_JSON у змінних оточення!")

if import_base64:
    import base64
    GSPREAD_CREDENTIALS_JSON = base64.b64decode(GSPREAD_CREDENTIALS_JSON).decode('utf-8')

try:
    gspread_creds_dict = json.loads(GSPREAD_CREDENTIALS_JSON)
except Exception as e:
    raise RuntimeError(f"Помилка парсингу GSPREAD_CREDENTIALS_JSON: {e}")

############################################
# 4) ФАЙЛИ users.json, applications_by_user.json, config.py
#    ЗБЕРІГАЄМО У '/data'
############################################

DATA_DIR = os.getenv("DATA_DIR", "/data")

if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR, exist_ok=True)

USERS_FILE = os.path.join(DATA_DIR, "users.json")
APPLICATIONS_FILE = os.path.join(DATA_DIR, "applications_by_user.json")
CONFIG_FILE = os.path.join(DATA_DIR, "config.py")

if not os.path.exists(USERS_FILE):
    initial_users_data = {"approved_users": {}, "blocked_users": [], "pending_users": {}}
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(initial_users_data, f, indent=2, ensure_ascii=False)

if not os.path.exists(APPLICATIONS_FILE):
    with open(APPLICATIONS_FILE, "w", encoding="utf-8") as f:
        json.dump({}, f, indent=2, ensure_ascii=False)

try:
    import importlib.util
    spec = importlib.util.spec_from_file_location("config", CONFIG_FILE)
    config_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config_module)
    CONFIG = config_module.CONFIG
except (ImportError, FileNotFoundError):
    default_config_content = '''# config.py
CONFIG = {
    "fgh_name_column": "D",
    "edrpou_column": "E",
    "region_column": "I",
    "district_column": "I",
    "city_column": "I",
    "group_column": "F",
    "culture_column": "G",
    "quantity_column": "H",
    "price_column": "M",
    "currency_column": "L",
    "payment_form_column": "K",
    "extra_fields_column": "J",
    "row_start": 2,
    "manager_price_column": "O",
    "user_id_column": "AZ",
    "phone_column": "P"
}
'''
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(default_config_content)
    spec = importlib.util.spec_from_file_location("config", CONFIG_FILE)
    config_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config_module)
    CONFIG = config_module.CONFIG

############################################
# 5) ФОРМАТИ ФОНУ (ЧЕРВОНИЙ, ЗЕЛЕНИЙ, ЖОВТИЙ, ТА ДОДАЄМО #de0000)
############################################

from gspread_formatting import cellFormat, Color

red_format = cellFormat(backgroundColor=Color(1, 0.8, 0.8))
green_format = cellFormat(backgroundColor=Color(0.8, 1, 0.8))
yellow_format = cellFormat(backgroundColor=Color(1, 1, 0.8))

# Червоний колір "#de0000" у десяткових частках (222 / 255 ~ 0.8705882353)
deep_red_format = cellFormat(backgroundColor=Color(0.8705882353, 0.0, 0.0))

############################################
# 6) ДОДАТКОВІ ПОЛЯ У FRIENDLY-ФОРМАТ
############################################

friendly_names = {
    "natura": "Натура",
    "bilok": "Білок",
    "kleikovina": "Клейковина",
    "smitteva": "Сміттєва домішка",
    "vologhist": "Вологість",
    "sazhkov": "Сажкові зерна",
    "natura_ya": "Натура",
    "vologhist_ya": "Вологість",
    "smitteva_ya": "Сміттєва домішка",
    "vologhist_k": "Вологість",
    "zernovadomishka": "Зернова домішка",
    "poshkodjeni": "Пошкоджені зерна",
    "smitteva_k": "Сміттєва домішка",
    "zipsovani": "Зіпсовані зерна",
    "olijnist_na_suhu": "Олійність на суху",
    "vologhist_son": "Вологість",
    "smitteva_son": "Сміттєва домішка",
    "kislotne": "Кислотне число",
    "olijnist_na_siru": "Олійність на сиру",
    "vologhist_ripak": "Вологість",
    "glukozinolati": "Глюкозінолати",
    "smitteva_ripak": "Сміттєва домішка",
    "bilok_na_siru": "Білок на сиру",
    "vologhist_soya": "Вологість",
    "smitteva_soya": "Сміттєва домішка",
    "olijna_domishka": "Олійна домішка",
    "ambrizia": "Амброзія"
}

############################################
# 7) ІНІЦІАЛІЗАЦІЯ БОТА
############################################

bot = Bot(token=TELEGRAM_TOKEN, parse_mode="HTML")
dp = Dispatcher(bot, storage=MemoryStorage())

def remove_keyboard():
    return types.ReplyKeyboardRemove()

############################################
# КЛАСИ СТАНІВ (FSM)
############################################

class RegistrationStates(StatesGroup):
    waiting_for_fullname = State()
    waiting_for_phone = State()

class ApplicationStates(StatesGroup):
    waiting_for_webapp_data = State()
    confirm_application = State()
    editing_application = State()
    viewing_application = State()
    proposal_reply = State()
    confirm_deletion = State()
    waiting_for_phone_confirmation = State()
    waiting_for_price_confirmation = State()

class AdminReview(StatesGroup):
    waiting_for_application_selection = State()
    waiting_for_decision = State()
    viewing_confirmed_list = State()
    viewing_confirmed_app = State()

############################################
# ФУНКЦІЇ РОБОТИ З ЛОКАЛЬНИМИ JSON-ФАЙЛАМИ
############################################

def load_users():
    with open(USERS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_users(data):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def load_applications():
    with open(APPLICATIONS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_applications(apps):
    with open(APPLICATIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(apps, f, indent=2, ensure_ascii=False)

############################################
# ДОДАЄМО/ОНОВЛЮЄМО ЗАЯВКУ
############################################

def add_application(user_id, chat_id, application_data):
    """
    Зберігає нову заявку до локального JSON, встановлюючи початковий статус active.
    """
    from datetime import datetime
    application_data['timestamp'] = datetime.now().isoformat()
    application_data['user_id'] = user_id
    application_data['chat_id'] = chat_id
    application_data["proposal_status"] = "active"
    apps = load_applications()
    uid = str(user_id)
    if uid not in apps:
        apps[uid] = []
    apps[uid].append(application_data)
    save_applications(apps)
    logging.info(f"Заявка для user_id={user_id} збережена як active.")

def update_application_status(user_id, app_index, status, proposal=None):
    """
    Оновлює статус заявки (наприклад, на Agreed, confirmed, waiting, deleted і т.ін.).
    За потреби можна встановити нову пропозицію (proposal).
    """
    apps = load_applications()
    uid = str(user_id)
    if uid in apps and 0 <= app_index < len(apps[uid]):
        apps[uid][app_index]["proposal_status"] = status
        if proposal is not None:
            apps[uid][app_index]["proposal"] = proposal
        save_applications(apps)

############################################
# БЛОКУВАННЯ/СХВАЛЕННЯ КОРИСТУВАЧІВ
############################################

def approve_user(user_id):
    data = load_users()
    uid = str(user_id)
    if uid not in data.get("approved_users", {}):
        pending = data.get("pending_users", {}).get(uid, {})
        fullname = pending.get("fullname", "")
        phone = pending.get("phone", "")
        data.setdefault("approved_users", {})[uid] = {"fullname": fullname, "phone": phone}
        data.get("pending_users", {}).pop(uid, None)
        save_users(data)
        logging.info(f"Користувач {uid} схвалений.")

def block_user(user_id):
    data = load_users()
    uid = str(user_id)
    if uid not in data.get("blocked_users", []):
        data.setdefault("blocked_users", []).append(uid)
        data.get("pending_users", {}).pop(uid, None)
        save_users(data)
        logging.info(f"Користувач {uid} заблокований.")

############################################
# ІНІЦІАЛІЗАЦІЯ GSPREAD
############################################

def init_gspread():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(gspread_creds_dict, scope)
    client = gspread.authorize(creds)
    return client

def get_worksheet1():
    client = init_gspread()
    sheet = client.open_by_key(GOOGLE_SPREADSHEET_ID)
    ws = sheet.worksheet(SHEET1_NAME)
    return ws

def get_worksheet2():
    client = init_gspread()
    sheet = client.open_by_key(GOOGLE_SPREADSHEET_ID2)
    ws = sheet.worksheet(SHEET2_NAME)
    return ws

############################################
# ЗБЕРІГАННЯ ЗАЯВКИ У ТАБЛИЦЮ 1
############################################

def ensure_columns(ws, required_col: int):
    if ws.col_count < required_col:
        ws.resize(rows=ws.row_count, cols=required_col)

def update_google_sheet(data: dict) -> int:
    """
    Додає заявку в кінець таблиці 1 (Лист1) і повертає номер нового рядка.
    """
    ws = get_worksheet1()
    ensure_columns(ws, 52)
    col_a = ws.col_values(1)
    new_row = len(col_a) + 1
    request_number = new_row - 1

    # 1-й стовпець: Номер заявки
    ws.update_cell(new_row, 1, request_number)
    # 2-й стовпець: Дата (день.місяць)
    current_date = datetime.now().strftime("%d.%m")
    ws.update_cell(new_row, 2, current_date)

    # 3-й стовпець: ПІБ у кілька рядків
    fullname = data.get("fullname", "")
    if isinstance(fullname, dict):
        fullname = fullname.get("fullname", "")
    fullname_lines = "\n".join(fullname.split())
    ws.update_cell(new_row, 3, fullname_lines)

    # 4-й: ФГ
    ws.update_cell(new_row, 4, data.get("fgh_name", ""))
    # 5-й: ЄДРПОУ
    ws.update_cell(new_row, 5, data.get("edrpou", ""))
    # 6-й: Група
    ws.update_cell(new_row, 6, data.get("group", ""))
    # 7-й: Культура
    ws.update_cell(new_row, 7, data.get("culture", ""))

    # 8-й: Кількість (з позначкою т)
    quantity = data.get("quantity", "")
    if quantity:
        quantity = f"{quantity} Т"
    ws.update_cell(new_row, 8, quantity)

    # 9-й: Область, район, місто
    region = data.get("region", "")
    district = data.get("district", "")
    city = data.get("city", "")
    location = f"Область: {region}\nРайон: {district}\nНас. пункт: {city}"
    ws.update_cell(new_row, 9, location)

    # 10-й: Додаткові параметри
    extra = data.get("extra_fields", {})
    extra_lines = []
    for key, value in extra.items():
        ukr_name = friendly_names.get(key, key.capitalize())
        extra_lines.append(f"{ukr_name}: {value}")
    ws.update_cell(new_row, 10, "\n".join(extra_lines))

    # 11-й: Форма оплати
    ws.update_cell(new_row, 11, data.get("payment_form", ""))

    # 12-й: Валюта
    currency_map = {"dollar": "Долар $", "euro": "Євро €", "uah": "Грн ₴"}
    curr = data.get("currency", "").lower()
    ws.update_cell(new_row, 12, currency_map.get(curr, data.get("currency", "")))

    # 13-й: Бажана ціна
    ws.update_cell(new_row, 13, data.get("price", ""))

    # 15-й: manager_price
    ws.update_cell(new_row, 15, data.get("manager_price", ""))

    # 16-й: Телефон
    ws.update_cell(new_row, 16, data.get("phone", ""))

    # 52-й: user_id
    ws.update_cell(new_row, 52, data.get("user_id", ""))

    return new_row

############################################
# ЗАФАРБОВУВАННЯ В ТАБЛИЦІ2 (СТОВПЧИК L = 12)
############################################

def color_price_cell_in_table2(row: int, fmt: cellFormat, col: int = 12):
    ws2 = get_worksheet2()
    cell_range = f"{rowcol_to_a1(row, col)}:{rowcol_to_a1(row, col)}"
    format_cell_range(ws2, cell_range, fmt)

def color_cell_red(row: int):
    color_price_cell_in_table2(row, red_format, 12)

def color_cell_green(row: int):
    color_price_cell_in_table2(row, green_format, 12)

def color_cell_yellow(row: int):
    color_price_cell_in_table2(row, yellow_format, 12)

############################################
# ДОДАТКОВІ ФУНКЦІЇ ФОРМАТУВАННЯ:
# 1) ЗАФАРБУВАТИ ВЕСЬ РЯДОК У ТАБЛИЦІ1
# 2) ЗАФАРБУВАТИ КЛІТИНКУ В ТАБЛИЦІ2 У #de0000
############################################

def color_entire_row_in_table1(row: int, fmt: cellFormat, max_columns: int = 52):
    """
    Зафарбовує весь рядок row (1-based index) у таблиці1 від колонки 1 до max_columns.
    """
    ws1 = get_worksheet1()
    start_cell = rowcol_to_a1(row, 1)     # наприклад, A{row}
    end_cell = rowcol_to_a1(row, max_columns)  # наприклад, AZ{row} якщо max_columns=52
    cell_range = f"{start_cell}:{end_cell}"
    format_cell_range(ws1, cell_range, fmt)

def color_cell_in_table2_red(row: int, col: int = 12):
    ws2 = get_worksheet2()
    cell_range = f"{rowcol_to_a1(row, col)}:{rowcol_to_a1(row, col)}"
    format_cell_range(ws2, cell_range, deep_red_format)


############################################
# СТВОРЕННЯ КНОПОК
############################################

def get_main_menu_keyboard():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("Подати заявку", "Переглянути мої заявки")
    return kb

def get_admin_menu_keyboard():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("Користувачі на модерацію", "Переглянути заявки культур")
    kb.add("Переглянути підтверджені заявки")
    return kb

############################################
# ОБРОБКА ДАНИХ ІЗ WEBAPP (JSON)
############################################

async def process_webapp_data_direct(user_id: int, data: dict, edit_index: int = None, sheet_row: int = None):
    """
    Формує текст попереднього перегляду заявки і пропонує кнопки:
    - Підтвердити
    - Редагувати
    - Скасувати
    """
    if not data or not any(data.values()):
        logging.warning("Отримано порожні дані, повідомлення не надсилається.")
        return

    message_lines = [
        "<b>Перевірте заявку:</b>",
        f"ФГ: {data.get('fgh_name', '')}",
        f"ЄДРПОУ: {data.get('edrpou', '')}",
        f"Область: {data.get('region', '')}",
        f"Район: {data.get('district', '')}",
        f"Місто: {data.get('city', '')}",
        f"Група: {data.get('group', '')}",
        f"Культура: {data.get('culture', '')}"
    ]
    extra = data.get("extra_fields", {})
    if extra:
        message_lines.append("Додаткові параметри:")
        for key, value in extra.items():
            ukr_name = friendly_names.get(key, key.capitalize())
            message_lines.append(f"{ukr_name}: {value}")

    message_lines.extend([
        f"Кількість: {data.get('quantity', '')} т",
        f"Форма оплати: {data.get('payment_form', '')}",
        f"Валюта: {data.get('currency', '')}",
        f"Ціна: {data.get('price', '')}"
    ])

    preview_text = "\n".join(message_lines)
    reply_kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    reply_kb.add("Підтвердити", "Редагувати", "Скасувати")

    await bot.send_message(user_id, preview_text, reply_markup=reply_kb)

    state = dp.current_state(chat=user_id, user=user_id)
    if edit_index is not None and sheet_row is not None:
        # Якщо редагуємо існуючу заявку (edit_index) і є посилання на конкретний ряд (sheet_row)
        await state.update_data(edit_index=edit_index, sheet_row=sheet_row, webapp_data=data)
        await state.set_state(ApplicationStates.editing_application.state)
    else:
        # Якщо це нова заявка
        await state.update_data(webapp_data=data)
        await state.set_state(ApplicationStates.confirm_application.state)

@dp.message_handler(lambda message: message.text and "/webapp_data" in message.text, state=ApplicationStates.waiting_for_webapp_data)
async def webapp_data_handler_text(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    try:
        prefix = "/webapp_data "
        if message.text.startswith(prefix):
            data_str = message.text[len(prefix):].strip()
        else:
            data_str = message.text.split("/webapp_data", 1)[-1].strip()

        data_dict = json.loads(data_str)
        await state.update_data(webapp_data=data_dict)
        current_data = await state.get_data()
        sheet_row = current_data.get("sheet_row")
        edit_index = current_data.get("edit_index")

        await process_webapp_data_direct(user_id, data_dict, edit_index, sheet_row)
    except Exception as e:
        logging.exception(f"Помилка обробки даних для user_id={user_id}: {e}")
        await bot.send_message(user_id, "Помилка обробки даних. Спробуйте ще раз.", reply_markup=remove_keyboard())

@dp.message_handler(content_types=types.ContentType.WEB_APP_DATA, state=ApplicationStates.waiting_for_webapp_data)
async def webapp_data_handler_web_app(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    try:
        data_str = message.web_app_data.data
        data_dict = json.loads(data_str)
        await state.update_data(webapp_data=data_dict)
        current_data = await state.get_data()
        sheet_row = current_data.get("sheet_row")
        edit_index = current_data.get("edit_index")

        await process_webapp_data_direct(user_id, data_dict, edit_index, sheet_row)
    except Exception as e:
        logging.exception(f"Помилка WEB_APP_DATA для user_id={user_id}: {e}")
        await bot.send_message(user_id, "Помилка обробки даних. Спробуйте ще раз.", reply_markup=remove_keyboard())

@dp.message_handler(Text(equals="Редагувати"), state=[ApplicationStates.confirm_application, ApplicationStates.editing_application])
async def edit_application_handler(message: types.Message, state: FSMContext):
    data = await state.get_data()
    webapp_data = data.get("webapp_data")
    if not webapp_data:
        await message.answer("Немає даних для редагування.", reply_markup=get_main_menu_keyboard())
        await state.finish()
        return

    webapp_url = "https://danza13.github.io/agro-webapp/webapp.html"
    prefill = quote(json.dumps(webapp_data))
    url_with_data = f"{webapp_url}?data={prefill}"

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add(types.KeyboardButton("Відкрити WebApp для редагування", web_app=types.WebAppInfo(url=url_with_data)))
    kb.row("Скасувати")

    await message.answer("Редагуйте заявку у WebApp:", reply_markup=kb)
    await state.set_state(ApplicationStates.waiting_for_webapp_data.state)

@dp.message_handler(Text(equals="Скасувати"), state=[ApplicationStates.waiting_for_webapp_data,
                                                     ApplicationStates.confirm_application,
                                                     ApplicationStates.editing_application])
async def cancel_process_reply(message: types.Message, state: FSMContext):
    await state.finish()
    await message.answer("Процес скасовано. Головне меню:", reply_markup=get_main_menu_keyboard())

############################################
# /start (Реєстрація користувача)
############################################

@dp.message_handler(commands=["start"], state="*")
async def cmd_start(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    await state.finish()
    users = load_users()
    uid = str(user_id)

    if uid in users.get("blocked_users", []):
        await message.answer("На жаль, у Вас немає доступу.", reply_markup=remove_keyboard())
        return

    if uid in users.get("approved_users", {}):
        await message.answer("Вітаємо! Оберіть дію:", reply_markup=get_main_menu_keyboard())
        return

    if uid in users.get("pending_users", {}):
        await message.answer("Ваша заявка на модерацію вже відправлена. Очікуйте.", reply_markup=remove_keyboard())
        return

    await message.answer("Введіть, будь ласка, своє ПІБ (повністю).", reply_markup=remove_keyboard())
    await RegistrationStates.waiting_for_fullname.set()

@dp.message_handler(commands=["menu"], state="*")
async def show_menu(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    users = load_users()
    uid = str(user_id)
    if uid not in users.get("approved_users", {}):
        await message.answer("Немає доступу. Очікуйте схвалення.", reply_markup=remove_keyboard())
        return
    await state.finish()
    await message.answer("Головне меню:", reply_markup=get_main_menu_keyboard())

@dp.message_handler(commands=["admin"], state="*")
async def admin_menu(message: types.Message, state: FSMContext):
    if str(message.from_user.id) not in ADMINS:
        await message.answer("Немає доступу.", reply_markup=remove_keyboard())
        return
    await state.finish()
    await message.answer("Адмін меню:", reply_markup=get_admin_menu_keyboard())

@dp.message_handler(state=RegistrationStates.waiting_for_fullname)
async def process_fullname(message: types.Message, state: FSMContext):
    fullname = message.text.strip()
    if not fullname:
        await message.answer("ПІБ не може бути порожнім.")
        return
    await state.update_data(fullname=fullname)

    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    keyboard.add(types.KeyboardButton("Поділитись контактом", request_contact=True))
    await message.answer("Введіть номер телефону (+380XXXXXXXXX) або поділіться контактом:", reply_markup=keyboard)
    await RegistrationStates.waiting_for_phone.set()

@dp.message_handler(content_types=types.ContentType.CONTACT, state=RegistrationStates.waiting_for_phone)
async def process_phone_contact(message: types.Message, state: FSMContext):
    phone = message.contact.phone_number if message.contact and message.contact.phone_number else ""
    phone = re.sub(r"[^\d+]", "", phone)
    await state.update_data(phone=phone)
    await complete_registration(message, state)

@dp.message_handler(state=RegistrationStates.waiting_for_phone)
async def process_phone_text(message: types.Message, state: FSMContext):
    phone = re.sub(r"[^\d+]", "", message.text.strip())
    if not re.fullmatch(r"\+380\d{9}", phone):
        await message.answer("Невірний формат. Введіть номер у форматі +380XXXXXXXXX")
        return
    await state.update_data(phone=phone)
    await complete_registration(message, state)

async def complete_registration(message: types.Message, state: FSMContext):
    data = await state.get_data()
    fullname = data.get("fullname")
    phone = data.get("phone")
    user_id = message.from_user.id
    uid = str(user_id)

    users = load_users()
    users.setdefault("pending_users", {})[uid] = {
        "fullname": fullname,
        "phone": phone,
        "timestamp": datetime.now().isoformat()
    }
    save_users(users)

    await state.finish()
    await message.answer("Ваша заявка на модерацію відправлена.", reply_markup=remove_keyboard())

    for admin in ADMINS:
        try:
            await bot.send_message(
                admin,
                f"Новий користувач на модерацію:\nПІБ: {fullname}\nНомер: {phone}\nUser ID: {user_id}",
                reply_markup=remove_keyboard()
            )
        except Exception as e:
            logging.exception(f"Не вдалося сповістити адміністратора {admin}: {e}")

############################################
# АДМІНСЬКІ КОМАНДИ: КОРИСТУВАЧІ НА МОДЕРАЦІЮ
############################################

@dp.message_handler(Text(equals="Користувачі на модерацію"), state="*")
async def admin_pending_requests(message: types.Message, state: FSMContext):
    if str(message.from_user.id) not in ADMINS:
        await message.answer("Немає доступу.", reply_markup=remove_keyboard())
        return

    users_data = load_users()
    pending = users_data.get("pending_users", {})
    if not pending:
        await message.answer("Немає заявок на модерацію.", reply_markup=remove_keyboard())
        return

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    for uid, info in pending.items():
        kb.add(info.get("fullname", "Невідомо"))
    kb.add("Назад")
    await message.answer("Оберіть заявку для перегляду:", reply_markup=kb)
    await AdminReview.waiting_for_application_selection.set()
    await state.update_data(pending_dict=pending)

@dp.message_handler(state=AdminReview.waiting_for_application_selection)
async def admin_select_application(message: types.Message, state: FSMContext):
    if message.text == "Назад":
        await state.finish()
        await message.answer("Адмін меню:", reply_markup=get_admin_menu_keyboard())
        return

    selected_fullname = message.text.strip()
    data = await state.get_data()
    pending = data.get("pending_dict", {})

    uid = None
    for k, info in pending.items():
        if info.get("fullname", "").strip() == selected_fullname:
            uid = k
            break

    if not uid:
        await message.answer("Заявку не знайдено. Спробуйте ще раз або натисніть 'Назад'.", 
                             reply_markup=remove_keyboard())
        return

    info = pending[uid]
    timestamp_str = info.get("timestamp", "")
    if timestamp_str:
        dt = datetime.fromisoformat(timestamp_str)
        dt_kyiv = dt.astimezone(ZoneInfo("Europe/Kiev"))
        formatted_timestamp = dt_kyiv.strftime("%d.%m.%Y | %H:%M:%S")
    else:
        formatted_timestamp = "Невідомо"

    text = (
        f"Користувач на модерацію:\n"
        f"User ID: {uid}\n"
        f"ПІБ: {info.get('fullname', 'Невідомо')}\n"
        f"Номер: {info.get('phone', '')}\n"
        f"Дата та час: {formatted_timestamp}"
    )

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add("Дозволити", "Заблокувати")
    kb.add("Назад")

    await state.update_data(selected_uid=uid)
    await message.answer(text, reply_markup=kb)
    await AdminReview.waiting_for_decision.set()

@dp.message_handler(lambda message: message.text in ["Дозволити", "Заблокувати"], state=AdminReview.waiting_for_decision)
async def admin_decision(message: types.Message, state: FSMContext):
    data = await state.get_data()
    uid = data.get("selected_uid")

    if not uid:
        await message.answer("Заявку не знайдено.", reply_markup=remove_keyboard())
        return

    if message.text == "Дозволити":
        approve_user(uid)
        response = "Користувача дозволено."
        try:
            await bot.send_message(uid, "Ви пройшли модерацію! Тепер можете користуватись ботом.", reply_markup=remove_keyboard())
        except Exception as e:
            logging.exception(f"Не вдалося сповістити користувача: {e}")
    else:
        block_user(uid)
        response = "Користувача заблоковано."
        try:
            await bot.send_message(uid, "На жаль, Ви не пройшли модерацію.", reply_markup=remove_keyboard())
        except Exception as e:
            logging.exception(f"Не вдалося сповістити користувача: {e}")

    users_data = load_users()
    if uid in users_data.get("pending_users", {}):
        users_data["pending_users"].pop(uid)
        save_users(users_data)

    await message.answer(response + " Натисніть 'Назад'.", reply_markup=types.ReplyKeyboardMarkup(resize_keyboard=True).add("Назад"))

@dp.message_handler(Text(equals="Назад"), state=AdminReview.waiting_for_decision)
async def admin_back_from_decision(message: types.Message, state: FSMContext):
    await state.finish()
    await message.answer("Адмін меню:", reply_markup=get_admin_menu_keyboard())

############################################
# ПЕРЕГЛЯД ПІДТВЕРДЖЕНИХ ЗАЯВОК (АДМІН)
############################################

@dp.message_handler(Text(equals="Переглянути підтверджені заявки"), state="*")
async def admin_view_confirmed_applications(message: types.Message, state: FSMContext):
    if str(message.from_user.id) not in ADMINS:
        await message.answer("Немає доступу.", reply_markup=remove_keyboard())
        return

    apps = load_applications()
    confirmed_apps = []
    for user_id, user_applications in apps.items():
        for idx, app_data in enumerate(user_applications):
            if app_data.get("proposal_status") == "confirmed":
                confirmed_apps.append({
                    "user_id": user_id,
                    "app_index": idx,
                    "app_data": app_data
                })

    if not confirmed_apps:
        await message.answer("Немає підтверджених заявок.", reply_markup=get_admin_menu_keyboard())
        return

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    row = []
    for i, entry in enumerate(confirmed_apps, start=1):
        culture = entry["app_data"].get("culture", "Невідомо")
        quantity = entry["app_data"].get("quantity", "Невідомо")
        btn_text = f"{i}. {culture} | {quantity}"
        row.append(btn_text)
        if len(row) == 2:
            kb.row(*row)
            row = []
    if row:
        kb.row(*row)
    kb.add("Назад")

    await state.update_data(confirmed_apps=confirmed_apps)
    await AdminReview.viewing_confirmed_list.set()
    await message.answer("Список підтверджених заявок:", reply_markup=kb)

@dp.message_handler(state=AdminReview.viewing_confirmed_list)
async def admin_view_confirmed_list_choice(message: types.Message, state: FSMContext):
    if message.text == "Назад":
        await state.finish()
        await message.answer("Адмін меню:", reply_markup=get_admin_menu_keyboard())
        return

    split_msg = message.text.split('.', 1)
    if len(split_msg) < 2 or not split_msg[0].isdigit():
        await message.answer("Оберіть номер заявки у форматі 'X. Культура | Кількість' або натисніть 'Назад'.")
        return

    choice = int(split_msg[0])
    data = await state.get_data()
    confirmed_apps = data.get("confirmed_apps", [])
    if choice < 1 or choice > len(confirmed_apps):
        await message.answer("Невірний вибір.", reply_markup=remove_keyboard())
        return

    selected_entry = confirmed_apps[choice - 1]
    app_data = selected_entry["app_data"]
    timestamp = app_data.get("timestamp", "")
    try:
        dt = datetime.fromisoformat(timestamp)
        formatted_date = dt.strftime("%d.%m.%Y")
    except Exception:
        formatted_date = timestamp

    details = [
        "<b>ЗАЯВКА ПІДТВЕРДЖЕНА:</b>",
        f"Дата створення: <b>{formatted_date}</b>",
        f"ФГ: <b>{app_data.get('fgh_name', '')}</b>",
        f"ЄДРПОУ: <b>{app_data.get('edrpou', '')}</b>",
        f"Область: <b>{app_data.get('region', '')}</b>",
        f"Район: <b>{app_data.get('district', '')}</b>",
        f"Місто: <b>{app_data.get('city', '')}</b>",
        f"Група: <b>{app_data.get('group', '')}</b>",
        f"Культура: <b>{app_data.get('culture', '')}</b>",
        f"Кількість: <b>{app_data.get('quantity', '')} т</b>",
        f"Форма оплати: <b>{app_data.get('payment_form', '')}</b>",
        f"Валюта: <b>{app_data.get('currency', '')}</b>",
        f"Бажана ціна: <b>{app_data.get('price', '')}</b>",
        f"Пропозиція ціни: <b>{app_data.get('proposal', '')}</b>"
    ]

    extra = app_data.get("extra_fields", {})
    if extra:
        details.append("Додаткові параметри:")
        for key, value in extra.items():
            details.append(f"{friendly_names.get(key, key.capitalize())}: {value}")

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add("Видалити", "Назад")

    await state.update_data(selected_confirmed=selected_entry, chosen_confirmed_index=choice - 1)
    await AdminReview.viewing_confirmed_app.set()
    await message.answer("\n".join(details), parse_mode="HTML", reply_markup=kb)

@dp.message_handler(Text(equals="Назад"), state=AdminReview.viewing_confirmed_app)
async def admin_back_to_confirmed_list(message: types.Message, state: FSMContext):
    data = await state.get_data()
    confirmed_apps = data.get("confirmed_apps", [])
    if not confirmed_apps:
        await state.finish()
        await message.answer("Список підтверджених заявок тепер порожній.", reply_markup=get_admin_menu_keyboard())
        return

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    row = []
    for i, entry in enumerate(confirmed_apps, start=1):
        culture = entry["app_data"].get("culture", "Невідомо")
        quantity = entry["app_data"].get("quantity", "Невідомо")
        btn_text = f"{i}. {culture} | {quantity}"
        row.append(btn_text)
        if len(row) == 2:
            kb.row(*row)
            row = []
    if row:
        kb.row(*row)
    kb.add("Назад")

    await AdminReview.viewing_confirmed_list.set()
    await message.answer("Список підтверджених заявок:", reply_markup=kb)

@dp.message_handler(Text(equals="Видалити"), state=AdminReview.viewing_confirmed_app)
async def admin_delete_confirmed_app(message: types.Message, state: FSMContext):
    """
    Адмін тисне "Видалити" підтверджену заявку:
    - Ми не видаляємо з таблиці, а лише ставимо статус deleted.
    - Зафарбовуємо весь рядок у таблиці1 в #de0000.
    - Зафарбовуємо клітинку в таблиці2 (стовпець L) так само.
    """
    data = await state.get_data()
    selected_entry = data.get("selected_confirmed")
    confirmed_apps = data.get("confirmed_apps", [])
    chosen_index = data.get("chosen_confirmed_index")

    if not selected_entry or chosen_index is None:
        await message.answer("Немає заявки для видалення.", reply_markup=get_admin_menu_keyboard())
        await state.finish()
        return

    user_id = int(selected_entry["user_id"])
    app_index = selected_entry["app_index"]
    apps = load_applications()
    uid = str(user_id)

    if uid in apps and 0 <= app_index < len(apps[uid]):
        app = apps[uid][app_index]
        # Оновлюємо статус
        app["proposal_status"] = "deleted"
        # Зафарбовуємо рядок у таблиці1
        sheet_row = app.get("sheet_row")
        if sheet_row:
            color_entire_row_in_table1(sheet_row, deep_red_format, 52)
            # Зафарбовуємо клітинку в таблиці2 (стовпець L = 12)
            color_cell_in_table2_red(sheet_row, 12)

        save_applications(apps)

    # Прибираємо заявку з локального списку confirmed_apps у стані FSM
    if 0 <= chosen_index < len(confirmed_apps):
        confirmed_apps.pop(chosen_index)

    await state.update_data(confirmed_apps=confirmed_apps, selected_confirmed=None, chosen_confirmed_index=None)

    if not confirmed_apps:
        await message.answer("Заявку позначено 'deleted'. Список підтверджених заявок тепер порожній.",
                             reply_markup=get_admin_menu_keyboard())
        await state.finish()
    else:
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        row = []
        for i, entry in enumerate(confirmed_apps, start=1):
            culture = entry["app_data"].get("culture", "Невідомо")
            quantity = entry["app_data"].get("quantity", "Невідомо")
            btn_text = f"{i}. {culture} | {quantity}"
            row.append(btn_text)
            if len(row) == 2:
                kb.row(*row)
                row = []
        if row:
            kb.row(*row)
        kb.add("Назад")

        await AdminReview.viewing_confirmed_list.set()
        await message.answer("Заявку позначено 'deleted'. Оновлений список:", reply_markup=kb)

############################################
# ОБРОБНИК «ПОДАТИ ЗАЯВКУ»
############################################

@dp.message_handler(Text(equals="Подати заявку"), state="*")
async def start_application(message: types.Message, state: FSMContext):
    await state.finish()
    webapp_url = "https://danza13.github.io/agro-webapp/webapp.html"
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add(types.KeyboardButton("Відкрити форму для заповнення", web_app=types.WebAppInfo(url=webapp_url)))
    kb.row("Скасувати")
    await message.answer("Заповніть дані заявки у WebApp:", reply_markup=kb)
    await ApplicationStates.waiting_for_webapp_data.set()

############################################
# ПЕРЕГЛЯД "МОЇХ ЗАЯВОК"
############################################

@dp.message_handler(Text(equals="Переглянути мої заявки"), state="*")
async def show_user_applications(message: types.Message):
    user_id = message.from_user.id
    uid = str(user_id)
    apps = load_applications()
    user_apps = apps.get(uid, [])

    if not user_apps:
        await message.answer("Ви не маєте заявок.", reply_markup=remove_keyboard())
        return

    buttons = []
    for i, app in enumerate(user_apps, start=1):
        culture = app.get('culture', 'Невідомо')
        quantity = app.get('quantity', 'Невідомо')
        status = app.get("proposal_status", "")
        if status == "confirmed":
            # Додаємо смайлик ✅
            btn_text = f"{i}. {culture} | {quantity} т ✅"
        else:
            btn_text = f"{i}. {culture} | {quantity} т"
        buttons.append(btn_text)

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    row = []
    for text in buttons:
        row.append(text)
        if len(row) == 2:
            kb.row(*row)
            row = []
    if row:
        kb.row(*row)
    kb.row("Назад")

    await message.answer("Ваші заявки:", reply_markup=kb)

############################################
# ДЕТАЛЬНИЙ ПЕРЕГЛЯД КОНКРЕТНОЇ ЗАЯВКИ
############################################

@dp.message_handler(Regexp(r"^(\d+)\.\s(.+)\s\|\s(.+)\sт(?:\s✅)?$"), state="*")
async def view_application_detail(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    uid = str(user_id)
    apps = load_applications()
    user_apps = apps.get(uid, [])

    match = re.match(r"^(\d+)\.\s(.+)\s\|\s(.+)\sт(?:\s✅)?$", message.text.strip())
    if not match:
        await message.answer("Невірна заявка.", reply_markup=remove_keyboard())
        return

    idx_str = match.group(1)
    idx = int(idx_str) - 1

    if idx < 0 or idx >= len(user_apps):
        await message.answer("Невірна заявка.", reply_markup=remove_keyboard())
        return

    app = user_apps[idx]
    timestamp = app.get("timestamp", "")
    try:
        dt = datetime.fromisoformat(timestamp)
        formatted_date = dt.strftime("%d.%m.%Y")
    except Exception:
        formatted_date = timestamp

    status = app.get("proposal_status", "")

    # Якщо заявка вже confirmed
    if status == "confirmed":
        details = [
            "<b>Детальна інформація по заявці:</b>",
            f"Дата створення: {formatted_date}",
            f"ФГ: {app.get('fgh_name', '')}",
            f"ЄДРПОУ: {app.get('edrpou', '')}",
            f"Область: {app.get('region', '')}",
            f"Район: {app.get('district', '')}",
            f"Місто: {app.get('city', '')}",
            f"Група: {app.get('group', '')}",
            f"Культура: {app.get('culture', '')}",
            f"Кількість: {app.get('quantity', '')}",
            f"Форма оплати: {app.get('payment_form', '')}",
            f"Валюта: {app.get('currency', '')}",
            f"Бажана ціна: {app.get('price', '')}",
            f"Пропозиція ціни: {app.get('proposal', '—')}",
            "Ціна була ухвалена, очікуйте, скоро з вами зв'яжуться"
        ]

        extra = app.get("extra_fields", {})
        if extra:
            details.append("Додаткові параметри:")
            for key, value in extra.items():
                details.append(f"{friendly_names.get(key, key.capitalize())}: {value}")

        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        kb.add("Назад")

        await state.update_data(selected_app_index=idx)
        await message.answer("\n".join(details), reply_markup=kb, parse_mode="HTML")
        await ApplicationStates.viewing_application.set()
        return

    # Якщо заявка має інші стани
    details = [
        "<b>Детальна інформація по заявці:</b>",
        f"Дата створення: {formatted_date}",
        f"ФГ: {app.get('fgh_name', '')}",
        f"ЄДРПОУ: {app.get('edrpou', '')}",
        f"Область: {app.get('region', '')}",
        f"Район: {app.get('district', '')}",
        f"Місто: {app.get('city', '')}",
        f"Група: {app.get('group', '')}",
        f"Культура: {app.get('culture', '')}",
        f"Кількість: {app.get('quantity', '')}",
        f"Форма оплати: {app.get('payment_form', '')}",
        f"Валюта: {app.get('currency', '')}",
        f"Бажана ціна: {app.get('price', '')}"
    ]

    extra = app.get("extra_fields", {})
    if extra:
        details.append("Додаткові параметри:")
        for key, value in extra.items():
            details.append(f"{friendly_names.get(key, key.capitalize())}: {value}")

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    if status == "Agreed":
        once_waited = app.get("onceWaited", False)
        details.append(f"\nПропозиція ціни: {app.get('proposal', '')}")
        if once_waited:
            kb.row("Підтвердити", "Видалити")
        else:
            kb.row("Підтвердити", "Відхилити", "Видалити")
    elif status == "active":
        kb.add("Переглянути пропозицію")
    elif status == "waiting":
        kb.add("Переглянути пропозицію")
    elif status == "rejected":
        kb.row("Видалити", "Очікувати")

    kb.add("Назад")
    await state.update_data(selected_app_index=idx)
    await message.answer("\n".join(details), reply_markup=kb, parse_mode="HTML")
    await ApplicationStates.viewing_application.set()

############################################
# ПЕРЕГЛЯНУТИ ПРОПОЗИЦІЮ
############################################

@dp.message_handler(Text(equals="Переглянути пропозицію"), state=ApplicationStates.viewing_application)
async def view_proposal(message: types.Message, state: FSMContext):
    data = await state.get_data()
    index = data.get("selected_app_index")
    if index is None:
        await message.answer("Немає даних про заявку.", reply_markup=remove_keyboard())
        return

    uid = str(message.from_user.id)
    apps = load_applications()
    user_apps = apps.get(uid, [])
    if index < 0 or index >= len(user_apps):
        await message.answer("Невірна заявка.", reply_markup=remove_keyboard())
        return

    app = user_apps[index]
    status = app.get("proposal_status", "")
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add("Назад")

    once_waited = app.get("onceWaited", False)
    proposal_text = f"Пропозиція по заявці: {app.get('proposal', 'Немає даних')}"

    if status == "confirmed":
        await message.answer("Ви вже підтвердили пропозицію, очікуйте результатів.", reply_markup=kb)
    elif status == "waiting":
        await message.answer("Очікування: як тільки менеджер оновить пропозицію, Вам прийде сповіщення.", reply_markup=kb)
    elif status == "Agreed":
        if once_waited:
            kb.row("Підтвердити", "Видалити")
        else:
            kb.row("Підтвердити", "Відхилити", "Видалити")
        await message.answer(proposal_text, reply_markup=kb)
    else:
        await message.answer("Немає актуальної пропозиції.", reply_markup=kb)

############################################
# ВІДХИЛИТИ ПРОПОЗИЦІЮ
############################################

@dp.message_handler(Text(equals="Відхилити"), state=ApplicationStates.viewing_application)
async def proposal_rejected(message: types.Message, state: FSMContext):
    data = await state.get_data()
    index = data.get("selected_app_index")
    update_application_status(message.from_user.id, index, "rejected")

    apps = load_applications()
    uid = str(message.from_user.id)
    app = apps[uid][index]
    sheet_row = app.get("sheet_row")
    if sheet_row:
        # Зафарбовуємо клітинку в табл.2 у червоний (м'який)
        color_cell_red(sheet_row)

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add("Видалити", "Очікувати")
    await message.answer("Пропозицію відхилено. Оберіть: Видалити заявку або Очікувати кращу пропозицію?", reply_markup=kb)
    await ApplicationStates.proposal_reply.set()

@dp.message_handler(Text(equals="Очікувати"), state=ApplicationStates.proposal_reply)
async def wait_after_rejection(message: types.Message, state: FSMContext):
    data = await state.get_data()
    index = data.get("selected_app_index")
    update_application_status(message.from_user.id, index, "waiting")

    apps = load_applications()
    uid = str(message.from_user.id)
    app = apps[uid][index]
    app["onceWaited"] = True
    sheet_row = app.get("sheet_row")
    if sheet_row:
        color_cell_yellow(sheet_row)

    save_applications(apps)
    await message.answer("Заявка оновлена. Ви будете повідомлені при появі кращої пропозиції.",
                         reply_markup=get_main_menu_keyboard())
    await state.finish()

############################################
# ВИДАЛЕННЯ ЗАЯВКИ (КОРИСТУВАЧЕМ)
############################################

@dp.message_handler(Text(equals="Видалити"), state=[ApplicationStates.viewing_application, ApplicationStates.proposal_reply])
async def delete_request_user(message: types.Message, state: FSMContext):
    """
    Користувач обирає "Видалити" заявку.
    Запитуємо підтвердження.
    """
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add("Так", "Ні")
    await message.answer("Ви впевнені, що хочете видалити заявку?", reply_markup=kb)
    await ApplicationStates.confirm_deletion.set()

@dp.message_handler(Text(equals="Так"), state=ApplicationStates.confirm_deletion)
async def confirm_deletion(message: types.Message, state: FSMContext):
    """
    Остаточне видалення заявки:
    - Ставимо `proposal_status = "deleted"`.
    - Зафарбовуємо рядок у таблиці1 (#de0000).
    - Зафарбовуємо клітинку стовпчика L (12) у таблиці2 (#de0000).
    """
    data = await state.get_data()
    index = data.get("selected_app_index")
    uid = str(message.from_user.id)
    apps = load_applications()
    user_apps = apps.get(uid, [])

    if index is None or index < 0 or index >= len(user_apps):
        await message.answer("Невірна заявка.", reply_markup=get_main_menu_keyboard())
        await state.finish()
        return

    app = user_apps[index]
    sheet_row = app.get("sheet_row")

    # Позначаємо статус заявки як deleted
    app["proposal_status"] = "deleted"

    # Зафарбовуємо рядок у таблиці1
    if sheet_row:
        color_entire_row_in_table1(sheet_row, deep_red_format, 52)
        # Зафарбовуємо клітинку в таблиці2 (стовпець L = 12)
        color_cell_in_table2_red(sheet_row, 12)

    # Зберігаємо зміни
    save_applications(apps)

    await message.answer("Ваша заявка позначена як 'deleted'.", reply_markup=get_main_menu_keyboard())
    await state.finish()

@dp.message_handler(Text(equals="Ні"), state=ApplicationStates.confirm_deletion)
async def cancel_deletion(message: types.Message, state: FSMContext):
    await ApplicationStates.viewing_application.set()
    await message.answer("Видалення скасовано.", reply_markup=get_main_menu_keyboard())

############################################
# КНОПКА "ПІДТВЕРДИТИ" ПІСЛЯ WEBAPP
############################################

@dp.message_handler(Text(equals="Підтвердити"), state=ApplicationStates.confirm_application)
async def confirm_application_handler(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    await message.answer("Очікуйте, зберігаємо заявку...")
    data_state = await state.get_data()
    webapp_data = data_state.get("webapp_data")

    if not webapp_data:
        await message.answer("Немає даних заявки. Спробуйте ще раз.", reply_markup=remove_keyboard())
        await state.finish()
        return

    if "fullname" not in webapp_data or not webapp_data.get("fullname"):
        users = load_users()
        approved_user_info = users.get("approved_users", {}).get(str(user_id), {})
        webapp_data["fullname"] = approved_user_info.get("fullname", "")

    webapp_data["chat_id"] = str(message.chat.id)
    webapp_data["original_manager_price"] = webapp_data.get("manager_price", "")

    try:
        sheet_row = update_google_sheet(webapp_data)
        webapp_data["sheet_row"] = sheet_row
        add_application(user_id, message.chat.id, webapp_data)

        await state.finish()
        await message.answer("Ваша заявка прийнята!", reply_markup=get_main_menu_keyboard())
    except Exception as e:
        logging.exception(f"Помилка при збереженні заявки: {e}")
        await message.answer("Сталася помилка при збереженні. Спробуйте пізніше.", reply_markup=remove_keyboard())
        await state.finish()

############################################
# ФОНОВИЙ ЦИКЛ ПЕРЕВІРКИ manager_price
############################################

async def poll_manager_proposals():
    while True:
        try:
            ws = get_worksheet1()
            rows = ws.get_all_values()
            apps = load_applications()
            for i, row in enumerate(rows[1:], start=2):
                # row - це список значень з рядка, де row[0] відповідає колонці A, row[1] - B тощо
                if len(row) < 15:
                    continue
                current_manager_price_str = row[14].strip()  # кол. O (15-й стовпець)
                if not current_manager_price_str:
                    continue
                try:
                    cur_price = float(current_manager_price_str)
                except ValueError:
                    continue

                for uid, app_list in apps.items():
                    for idx, app in enumerate(app_list):
                        if app.get("sheet_row") == i:
                            status = app.get("proposal_status", "active")
                            original_manager_price_str = app.get("original_manager_price", "").strip()
                            try:
                                orig_price = float(original_manager_price_str) if original_manager_price_str else None
                            except:
                                orig_price = None

                            if orig_price is None:
                                # Перший раз отримали manager_price
                                culture = app.get("culture", "Невідомо")
                                quantity = app.get("quantity", "Невідомо")
                                app["original_manager_price"] = current_manager_price_str
                                app["proposal"] = current_manager_price_str
                                app["proposal_status"] = "Agreed"
                                await bot.send_message(
                                    app.get("chat_id"),
                                    f"Нова пропозиція по Вашій заявці {idx+1}. {culture} | {quantity} т. Ціна: {current_manager_price_str}"
                                )
                            else:
                                previous_proposal = app.get("proposal")
                                if previous_proposal != current_manager_price_str:
                                    app["original_manager_price"] = previous_proposal
                                    app["proposal"] = current_manager_price_str
                                    app["proposal_status"] = "Agreed"

                                    if status == "waiting":
                                        culture = app.get("culture", "Невідомо")
                                        quantity = app.get("quantity", "Невідомо")
                                        await bot.send_message(
                                            app.get("chat_id"),
                                            f"Ціна по заявці {idx+1}. {culture} | {quantity} т змінилась з {previous_proposal} на {current_manager_price_str}"
                                        )
                                    else:
                                        await bot.send_message(
                                            app.get("chat_id"),
                                            f"Для Вашої заявки оновлено пропозицію: {current_manager_price_str}"
                                        )
            save_applications(apps)
        except Exception as e:
            logging.exception(f"Помилка у фоні: {e}")

        await asyncio.sleep(CHECK_INTERVAL)

############################################
# HTTP-СЕРВЕР (ОПЦІОНАЛЬНО)
############################################

async def handle_webapp_data(request: web.Request):
    try:
        data = await request.json()
        user_id = data.get("user_id")
        if not user_id:
            return web.json_response({"status": "error", "error": "user_id missing"})
        if not data or not any(data.values()):
            return web.json_response({"status": "error", "error": "empty data"})
        logging.info(f"API отримав дані для user_id={user_id}: {json.dumps(data, ensure_ascii=False)}")
        return web.json_response({"status": "preview"})
    except Exception as e:
        logging.exception(f"API: Помилка: {e}")
        return web.json_response({"status": "error", "error": str(e)})

async def start_webserver():
    app_web = web.Application()
    app_web.add_routes([web.post('/api/webapp_data', handle_webapp_data)])
    runner = web.AppRunner(app_web)
    await runner.setup()
    site = web.TCPSite(runner, '0.0.0.0', API_PORT)
    await site.start()
    logging.info(f"HTTP-сервер запущено на порті {API_PORT}.")

############################################
# on_startup
############################################

async def on_startup(dp):
    logging.info("Бот запущено. Старт фонових задач...")
    asyncio.create_task(poll_manager_proposals())
    asyncio.create_task(start_webserver())

############################################
# ТОЧКА ВХОДУ
############################################

if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True, on_startup=on_startup)
