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
from gspread_formatting import format_cell_range, cellFormat, Color, set_column_width
from gspread.utils import rowcol_to_a1

############################################
# 1) ЧИТАЄМО ЗМІННІ ОТОЧЕННЯ
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
# 5) ФОРМАТИ ФОНУ (ЧЕРВОНИЙ, ЗЕЛЕНИЙ, ЖОВТИЙ)
############################################

red_format = cellFormat(backgroundColor=Color(1, 0.8, 0.8))
green_format = cellFormat(backgroundColor=Color(0.8, 1, 0.8))
yellow_format = cellFormat(backgroundColor=Color(1, 1, 0.8))

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
# ФЛАГ ПАУЗИ ДЛЯ ПОЛІНГУ
############################################

POLLING_PAUSED = False

def pause_polling():
    global POLLING_PAUSED
    POLLING_PAUSED = True

def resume_polling():
    global POLLING_PAUSED
    POLLING_PAUSED = False

############################################
# КЛАСИ СТАНІВ (FSM)
############################################

class RegistrationStates(StatesGroup):
    waiting_for_fullname = State()
    waiting_for_phone = State()
    preview = State()             # Попередній перегляд (ПІБ + телефон)
    editing = State()            # Меню редагування
    editing_fullname = State()   # Зміна ПІБ
    editing_phone = State()      # Зміна телефону

class ApplicationStates(StatesGroup):
    waiting_for_webapp_data = State()
    confirm_application = State()
    editing_application = State()
    viewing_application = State()
    proposal_reply = State()
    confirm_deletion = State()
    waiting_for_phone_confirmation = State()
    waiting_for_price_confirmation = State()

class AdminMenuStates(StatesGroup):
    choosing_section = State()     # «Модерація» або «Заявки» або «Вийти»
    moderation_section = State()   # усередині «Модерація»
    requests_section = State()     # усередині «Заявки»

class AdminReview(StatesGroup):
    waiting_for_application_selection = State()
    waiting_for_decision = State()
    viewing_confirmed_list = State()
    viewing_confirmed_app = State()
    viewing_deleted_list = State()
    viewing_deleted_app = State()
    # **NEW**: перегляд / редагування approved-користувачів
    viewing_approved_list = State()
    viewing_approved_user = State()
    editing_approved_user = State()
    editing_approved_user_fullname = State()
    editing_approved_user_phone = State()

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

def add_application(user_id, chat_id, application_data):
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

############################################
# ФУНКЦІЇ ДЛЯ АДМІНА: APPROVE І BLOCK
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
        data.get("approved_users", {}).pop(uid, None)
        save_users(data)
        logging.info(f"Користувач {uid} заблокований.")

############################################
# ОНОВЛЕННЯ СТАТУСУ ЗАЯВКИ
############################################

def update_application_status(user_id, app_index, status, proposal=None):
    apps = load_applications()
    uid = str(user_id)
    if uid in apps and 0 <= app_index < len(apps[uid]):
        apps[uid][app_index]["proposal_status"] = status
        if proposal is not None:
            apps[uid][app_index]["proposal"] = proposal
        save_applications(apps)

def delete_application_soft(user_id, app_index):
    """
    «М'яке» видалення: тільки змінюємо status -> 'deleted', не видаляємо з файлу.
    Таким чином заявка переходить у «видалені», але ще лежить у файлі.
    """
    apps = load_applications()
    uid = str(user_id)
    if uid in apps and 0 <= app_index < len(apps[uid]):
        apps[uid][app_index]["proposal_status"] = "deleted"
        save_applications(apps)

def delete_application_from_file_entirely(user_id, app_index):
    """Повне видалення з файлу (із масиву apps[uid])."""
    apps = load_applications()
    uid = str(user_id)
    if uid in apps and 0 <= app_index < len(apps[uid]):
        del apps[uid][app_index]
        if not apps[uid]:
            apps.pop(uid, None)  # Якщо масив став порожнім
        save_applications(apps)

############################################
# ФУНКЦІЇ ВИДАЛЕННЯ РЯДКА У GSheets
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

def ensure_columns(ws, required_col: int):
    if ws.col_count < required_col:
        ws.resize(rows=ws.row_count, cols=required_col)

def delete_price_cell_in_table2(row: int, col: int = 12):
    """
    Видалити клітинку в таблиці2 (SHEET2_NAME) у вказаному рядку
    зі зсувом усіх наступних рядків вгору, а також
    прибрати колір (щоб не переносився на інші рядки).
    """
    ws2 = get_worksheet2()

    # Спочатку очищаємо форматування (колір) у стовпці починаючи з цього рядка:
    format_cell_range(
        ws2,
        f"{rowcol_to_a1(row, col)}:{rowcol_to_a1(ws2.row_count, col)}",
        cellFormat(backgroundColor=Color(1, 1, 1))
    )

    col_values = ws2.col_values(col)
    if row - 1 >= len(col_values):
        return

    # Видаляємо значення (зсув вгору)
    col_values.pop(row - 1)
    for i in range(row - 1, len(col_values)):
        ws2.update_cell(i + 1, col, col_values[i])

    # Очищаємо останню клітинку після зсуву
    last_row_to_clear = len(col_values) + 1
    ws2.update_cell(last_row_to_clear, col, "")
    
from gspread_formatting import (
    format_cell_range,
    CellFormat,
    TextFormat,
    set_column_width
)

def export_database():
    """
    Створює новий лист у таблиці 1 з назвою "База дд.мм"
    і вносить дані для кожного схваленого користувача:
      A: Телеграм‑ID
      B: ПІБ
      C: Номер телефону
      D: Остання заявка (дата у форматі дд.мм.рррр, alt+enter, час у форматі гг:хв)
      E: Загальна кількість заявок
    Використовуємо один виклик update() для завантаження всієї матриці даних,
    а потім застосовуємо форматування:
      - Вирівнювання по центру (горизонтально та вертикально)
      - Жирний шрифт для всього тексту
      - Автоматичне встановлення ширини стовпців на основі максимального вмісту
    """
    # Завантаження даних
    users_data = load_users()
    approved = users_data.get("approved_users", {})
    apps = load_applications()

    # Отримуємо доступ до таблиці 1
    client = init_gspread()
    sheet = client.open_by_key(GOOGLE_SPREADSHEET_ID)

    # Форматуємо назву листа за поточною датою (наприклад, "База 13.02")
    today = datetime.now().strftime("%d.%m")
    new_title = f"База {today}"
    new_ws = sheet.add_worksheet(title=new_title, rows="1000", cols="5")

    # Формуємо матрицю даних: перший рядок — заголовки, решта — дані користувачів
    headers = ["ID", "ПІБ", "Номер телефону", "Остання заявка", "Загальна кількість заявок"]
    data_matrix = [headers]

    for uid, info in approved.items():
        user_apps = apps.get(uid, [])
        count_apps = len(user_apps)
        last_timestamp = ""
        if count_apps > 0:
            last_app = max(user_apps, key=lambda a: a.get("timestamp", ""))
            ts = last_app.get("timestamp", "")
            try:
                dt = datetime.fromisoformat(ts)
                last_timestamp = dt.strftime("%d.%m.%Y\n%H:%M")
            except Exception:
                last_timestamp = ts
        row = [uid, info.get("fullname", ""), info.get("phone", ""), last_timestamp, count_apps]
        data_matrix.append(row)

    # Batch update всіх клітинок матриці даних
    end_row = len(data_matrix)
    cell_range = f"A1:E{end_row}"
    new_ws.update(cell_range, data_matrix, value_input_option="USER_ENTERED")

    # Створюємо форматування: відцентровування і жирний шрифт
    cell_format = CellFormat(
        horizontalAlignment='CENTER',
        verticalAlignment='MIDDLE',
        textFormat=TextFormat(bold=True)
    )
    # Форматуємо весь діапазон
    format_cell_range(new_ws, cell_range, cell_format)

    # Автоматичне встановлення ширини стовпців
    # Обчислюємо максимальну довжину тексту для кожного стовпця
    num_cols = 5
    for col in range(1, num_cols + 1):
        # Отримуємо літерне позначення стовпця, напр. "A" для 1, "B" для 2 і т.д.
        col_letter = rowcol_to_a1(1, col)[0]
        col_range = f"{col_letter}:{col_letter}"
        max_len = max(len(str(row[col-1])) for row in data_matrix)
        width = max_len * 10  # Приблизне число пікселів на символ (можна налаштувати)
        set_column_width(new_ws, col_range, width)

@dp.message_handler(Text(equals="Вивантажити базу"), state=AdminReview.viewing_approved_list)
async def handle_export_database(message: types.Message, state: FSMContext):
    """
    Хендлер для кнопки "Вивантажити базу" у розділі "База користувачів".
    При натисканні викликається функція export_database(), а адміністратору повідомляється про результат.
    """
    try:
        export_database()
        await message.answer("База успішно вивантажена до Google Sheets.", reply_markup=get_admin_moderation_menu())
    except Exception as e:
        logging.exception(f"Помилка вивантаження бази: {e}")
        await message.answer("Помилка вивантаження бази.", reply_markup=get_admin_moderation_menu())

############################################
# ПОВНЕ ВИДАЛЕННЯ ЗАЯВКИ АДМІНОМ (5 КРОКІВ)
############################################

async def admin_remove_app_permanently(user_id: int, app_index: int):
    """
    1) Зупинити poll_manager_proposals (pause_polling)
    2) Видалити заявку з applications_by_user.json (остаточно)
    3) Видалити рядок у sheets1
    4) Видалити клітинку з таблиці2 (зі зсувом тексту й очищенням кольору)
    5) Затримка 20 секунд, і відновити poll_manager_proposals (resume_polling)
    """
    pause_polling()
    try:
        apps = load_applications()
        uid = str(user_id)
        if uid not in apps or app_index < 0 or app_index >= len(apps[uid]):
            return False

        app = apps[uid][app_index]
        sheet_row = app.get("sheet_row")

        # 2) Видаляємо з файлу
        delete_application_from_file_entirely(user_id, app_index)

        # 3) Видаляємо рядок у таблиці1
        if sheet_row:
            try:
                # 4) Видаляємо клітинку в таблиці2 (колонка із ціною, за замовчуванням col=12)
                delete_price_cell_in_table2(sheet_row, 12)

                ws = get_worksheet1()
                ws.delete_rows(sheet_row)

                # Оновлюємо sheet_row у залишилихся заявках (все, що було нижче - змістилося вгору на 1)
                updated_apps = load_applications()
                for u_str, user_apps in updated_apps.items():
                    for a in user_apps:
                        old_row = a.get("sheet_row", 0)
                        if old_row and old_row > sheet_row:
                            a["sheet_row"] = old_row - 1
                save_applications(updated_apps)
            except Exception as e:
                logging.exception(f"Помилка видалення рядка в Google Sheets: {e}")

        return True
    finally:
        # Додаємо затримку 20 секунд, щоб встигли оновитися дані в таблиці
        await asyncio.sleep(20)
        resume_polling()

############################################
# ЗАПИС У ТАБЛИЦЮ 1
############################################

def update_google_sheet(data: dict) -> int:
    ws = get_worksheet1()
    ensure_columns(ws, 52)

    # Отримуємо значення першого стовпця
    col_a = ws.col_values(1)
    # Пропускаємо перший рядок (заголовки) і фільтруємо числові значення
    numeric_values = []
    for value in col_a[1:]:
        try:
            numeric_values.append(int(value))
        except ValueError:
            continue

    # Якщо є числа, беремо останнє, інакше встановлюємо 0
    last_number = numeric_values[-1] if numeric_values else 0
    new_request_number = last_number + 1

    # Обчислюємо номер нового рядка: це буде (кількість рядків + 1)
    new_row = len(col_a) + 1

    # Записуємо номер заявки в першу клітинку нового рядка
    ws.update_cell(new_row, 1, new_request_number)

    # Далі оновлюємо інші клітинки (наприклад, дата, ПІБ тощо)
    current_date = datetime.now().strftime("%d.%m")
    ws.update_cell(new_row, 2, current_date)
    fullname = data.get("fullname", "")
    if isinstance(fullname, dict):
        fullname = fullname.get("fullname", "")
    fullname_lines = "\n".join(fullname.split())
    ws.update_cell(new_row, 3, fullname_lines)
    ws.update_cell(new_row, 4, data.get("fgh_name", ""))
    ws.update_cell(new_row, 5, data.get("edrpou", ""))
    ws.update_cell(new_row, 6, data.get("group", ""))
    ws.update_cell(new_row, 7, data.get("culture", ""))

    quantity = data.get("quantity", "")
    if quantity:
        quantity = f"{quantity} Т"
    ws.update_cell(new_row, 8, quantity)

    region = data.get("region", "")
    district = data.get("district", "")
    city = data.get("city", "")
    location = f"Область: {region}\nРайон: {district}\nНас. пункт: {city}"
    ws.update_cell(new_row, 9, location)

    extra = data.get("extra_fields", {})
    extra_lines = []
    for key, value in extra.items():
        ukr_name = friendly_names.get(key, key.capitalize())
        extra_lines.append(f"{ukr_name}: {value}")
    ws.update_cell(new_row, 10, "\n".join(extra_lines))

    ws.update_cell(new_row, 11, data.get("payment_form", ""))

    currency_map = {"dollar": "Долар $", "euro": "Євро €", "uah": "Грн ₴"}
    curr = data.get("currency", "").lower()
    ws.update_cell(new_row, 12, currency_map.get(curr, data.get("currency", "")))

    ws.update_cell(new_row, 13, data.get("price", ""))
    ws.update_cell(new_row, 15, data.get("manager_price", ""))
    ws.update_cell(new_row, 16, data.get("phone", ""))
    ws.update_cell(new_row, 52, data.get("user_id", ""))

    return new_row

############################################
# ЗАФАРБОВУВАННЯ КЛІТИНОК У ТАБЛИЦІ2
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
# ГОЛОВНІ КЛАВІАТУРИ (ЮЗЕР / АДМІН)
############################################

def get_main_menu_keyboard():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("Подати заявку", "Переглянути мої заявки")
    return kb

def get_admin_root_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.row("Модерація", "Заявки")
    kb.add("Вийти з адмін-меню")
    return kb

def get_admin_moderation_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.row("Користувачі на модерацію", "База користувачів")
    kb.add("Назад")
    return kb

def get_admin_requests_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.row("Підтверджені", "Видалені")
    kb.add("Переглянути заявки культур", "Видалення заявок")
    kb.add("Назад")
    return kb
    
############################################
# ХЕНДЛЕР Видалення заявок
############################################

@dp.message_handler(Text(equals="Видалення заявок"), state=AdminMenuStates.requests_section)
async def handle_delete_applications(message: types.Message, state: FSMContext):
    """
    Обробляє кнопку «Видалення заявок»: зчитує дані з Google Sheets і формує клавіатуру,
    де для кожного рядка (з заявкою) показується номер заявки (значення клітинки A) та номер рядка.
    """
    try:
        ws = get_worksheet1()
        rows = ws.get_all_values()
        if len(rows) <= 1:
            await message.answer("У таблиці немає заявок.", reply_markup=get_admin_requests_menu())
            return

        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        # Пропускаємо заголовковий рядок
        for i, row in enumerate(rows[1:], start=2):
            if row and row[0].strip():
                request_number = row[0].strip()
                # Наприклад, "123 (рядок 5)"
                btn_text = f"{request_number} (рядок {i})"
                kb.add(btn_text)
        kb.add("Назад")
        await message.answer("Оберіть заявку для видалення:", reply_markup=kb)
    except Exception as e:
        logging.exception("Помилка отримання заявок з Google Sheets")
        await message.answer("Помилка отримання заявок.", reply_markup=get_admin_requests_menu())


@dp.message_handler(lambda message: re.match(r"^\d+\s\(рядок\s\d+\)$", message.text), state=AdminMenuStates.requests_section)
async def handle_delete_application_selection(message: types.Message, state: FSMContext):
    """
    Обробляє вибір конкретної заявки для видалення.
    Парсить номер рядка із тексту кнопки, знаходить заявку у JSON (за полем sheet_row)
    і викликає admin_remove_app_permanently для остаточного видалення.
    """
    text = message.text.strip()
    match = re.search(r"\(рядок\s(\d+)\)$", text)
    if not match:
        await message.answer("Невірний формат вибору.", reply_markup=get_admin_requests_menu())
        return
    row_number = int(match.group(1))

    apps = load_applications()
    found = False
    for uid, app_list in apps.items():
        for idx, app in enumerate(app_list):
            if app.get("sheet_row") == row_number:
                success = await admin_remove_app_permanently(int(uid), idx)
                if success:
                    await message.answer(f"Заявку з рядка {row_number} успішно видалено.", reply_markup=get_admin_requests_menu())
                else:
                    await message.answer("Помилка видалення заявки.", reply_markup=get_admin_requests_menu())
                found = True
                break
        if found:
            break
    if not found:
        await message.answer("Заявку не знайдено.", reply_markup=get_admin_requests_menu())

############################################
# ХЕНДЛЕР /admin
############################################

@dp.message_handler(commands=["admin"], state="*")
async def admin_entry_point(message: types.Message, state: FSMContext):
    if str(message.from_user.id) not in ADMINS:
        await message.answer("Немає доступу.", reply_markup=remove_keyboard())
        return

    await state.finish()
    await message.answer("Ви в адмін-меню. Оберіть розділ:", reply_markup=get_admin_root_menu())
    await AdminMenuStates.choosing_section.set()

############################################
# ХЕНДЛЕР ДЛЯ ГОЛОВНОГО АДМІН-МЕНЮ
############################################

@dp.message_handler(state=AdminMenuStates.choosing_section)
async def admin_menu_choosing_section(message: types.Message, state: FSMContext):
    text = message.text.strip()

    if text == "Модерація":
        await message.answer("Розділ 'Модерація'. Оберіть дію:", reply_markup=get_admin_moderation_menu())
        await AdminMenuStates.moderation_section.set()
    elif text == "Заявки":
        await message.answer("Розділ 'Заявки'. Оберіть дію:", reply_markup=get_admin_requests_menu())
        await AdminMenuStates.requests_section.set()
    elif text == "Вийти з адмін-меню":
        await state.finish()
        await message.answer("Вихід з адмін-меню. Повертаємось у звичайне меню:", reply_markup=get_main_menu_keyboard())
    else:
        await message.answer("Будь ласка, оберіть із меню: «Модерація», «Заявки» або «Вийти з адмін-меню».")

############################################
# 1) РОЗДІЛ «МОДЕРАЦІЯ»
############################################

@dp.message_handler(state=AdminMenuStates.moderation_section)
async def admin_moderation_section_handler(message: types.Message, state: FSMContext):
    """
    Головний хендлер для розділу «Модерація».
    """
    text = message.text.strip()

    if text == "Користувачі на модерацію":
        users_data = load_users()
        pending = users_data.get("pending_users", {})

        if not pending:
            await message.answer("Немає заявок на модерацію.", reply_markup=get_admin_moderation_menu())
            return

        # Формуємо клавіатуру зі списком користувачів, які очікують
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        for uid, info in pending.items():
            kb.add(info.get("fullname", "Невідомо"))
        kb.add("Назад")
        await message.answer("Оберіть заявку для перегляду:", reply_markup=kb)

        # Зберігаємо pending у state та виставляємо стан
        await AdminReview.waiting_for_application_selection.set()
        await state.update_data(pending_dict=pending, from_moderation_menu=True)

    elif text == "База користувачів":
        users_data = load_users()
        approved = users_data.get("approved_users", {})
        if not approved:
            await message.answer("Немає схвалених користувачів.", reply_markup=get_admin_moderation_menu())
            return

        # Відображаємо схвалених по 2 в рядку
        approved_dict = {}
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        row = []
        for u_id, info in approved.items():
            name = info.get("fullname", f"ID:{u_id}")
            approved_dict[name] = u_id  # Зберігаємо зіставлення ім'я -> ID
            row.append(name)
            if len(row) == 2:
                kb.row(*row)
                row = []
        if row:
            kb.row(*row)
        kb.add("Вивантажити базу", "Назад")
        await state.update_data(approved_dict=approved_dict, from_moderation_menu=True)
        await message.answer("Список схвалених користувачів (по два в рядку):", reply_markup=kb)
        await AdminReview.viewing_approved_list.set()

    elif text == "Назад":
        # Повертаємось у головне меню адміна
        await message.answer("Головне меню адміна:", reply_markup=get_admin_root_menu())
        await AdminMenuStates.choosing_section.set()

    else:
        await message.answer("Оберіть зі списку: «Користувачі на модерацію», «База користувачів» або «Назад».")


@dp.message_handler(state=AdminReview.waiting_for_application_selection)
async def admin_select_pending_application(message: types.Message, state: FSMContext):
    """
    Хендлер, що реагує на вибір конкретного користувача із «pending».
    """
    if message.text == "Назад":
        await message.answer("Повертаємось до розділу 'Модерація':", reply_markup=get_admin_moderation_menu())
        await AdminMenuStates.moderation_section.set()
        return

    data = await state.get_data()
    pending = data.get("pending_dict", {})
    selected_fullname = message.text.strip()

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


@dp.message_handler(lambda msg: msg.text in ["Дозволити", "Заблокувати"], state=AdminReview.waiting_for_decision)
async def admin_decision_pending_user(message: types.Message, state: FSMContext):
    """
    Обробляє натискання кнопок «Дозволити» / «Заблокувати» 
    для pending-користувачів (стан `waiting_for_decision`).
    """
    data = await state.get_data()
    uid = data.get("selected_uid", None)
    if not uid:
        await message.answer("Не знайдено користувача.", reply_markup=remove_keyboard())
        return

    if message.text == "Дозволити":
        approve_user(uid)
        response_text = "Користувача дозволено."
        # Надсилаємо йому повідомлення (uid - це рядок, тому int(uid))
        try:
            await bot.send_message(
                int(uid),
                "Вітаємо! Ви пройшли модерацію і тепер можете користуватися ботом.",
                reply_markup=remove_keyboard()
            )
        except Exception as e:
            logging.exception(f"Не вдалося надіслати повідомлення користувачу {uid}: {e}")
    else:
        block_user(uid)
        response_text = "Користувача заблоковано."
        try:
            await bot.send_message(
                int(uid),
                "На жаль, Ви не пройшли модерацію.",
                reply_markup=remove_keyboard()
            )
        except Exception as e:
            logging.exception(f"Не вдалося надіслати повідомлення користувачу {uid}: {e}")

    # Прибираємо з pending_users
    users_data = load_users()
    if uid in users_data.get("pending_users", {}):
        users_data["pending_users"].pop(uid)
        save_users(users_data)

    # Відповідаємо адміну
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("Назад")
    await message.answer(f"{response_text}\nНатисніть «Назад» для повернення в меню.", reply_markup=kb)
    await AdminMenuStates.moderation_section.set()
#
# 1.1) Перегляд Approved-користувачів
#

@dp.message_handler(state=AdminReview.viewing_approved_list)
async def admin_view_approved_users(message: types.Message, state: FSMContext):
    """
    Стан для перегляду списку схвалених користувачів (approved_dict),
    де на кнопках показується лише ім'я користувача.
    """
    text = message.text.strip()
    data = await state.get_data()
    from_moderation_menu = data.get("from_moderation_menu", False)
    approved_dict = data.get("approved_dict", {})

    if text == "Назад":
        if from_moderation_menu:
            await message.answer("Повертаємось до розділу 'Модерація':", reply_markup=get_admin_moderation_menu())
            await AdminMenuStates.moderation_section.set()
        else:
            await state.finish()
            await message.answer("Головне меню адміна:", reply_markup=get_admin_root_menu())
        return

    # Тепер очікуємо, що текст відповідає ключу (ім'ю користувача) зі словника approved_dict
    if text not in approved_dict:
        await message.answer("Оберіть користувача зі списку або натисніть «Назад».")
        return

    user_id = approved_dict[text]
    users_data = load_users()
    approved_users = users_data.get("approved_users", {})

    if str(user_id) not in approved_users:
        await message.answer("Користувача не знайдено серед схвалених.")
        return

    info = approved_users[str(user_id)]
    fullname = info.get("fullname", "—")
    phone = info.get("phone", "—")

    details = (
        f"ПІБ: {fullname}\n"
        f"Номер телефону: {phone}\n"
        f"Телеграм ID: {user_id}"
    )

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.row("Редагувати", "Видалити")
    kb.add("Назад")

    await state.update_data(selected_approved_user_id=str(user_id))
    await AdminReview.viewing_approved_user.set()
    await message.answer(details, reply_markup=kb)


#
# 1.2) Детальний перегляд 1 approved-користувача
#

@dp.message_handler(state=AdminReview.viewing_approved_user)
async def admin_view_approved_single_user(message: types.Message, state: FSMContext):
    """
    Стан для перегляду / видалення / редагування 
    конкретного схваленого користувача.
    """
    text = message.text.strip()
    data = await state.get_data()
    user_id_str = data.get("selected_approved_user_id", None)
    from_moderation_menu = data.get("from_moderation_menu", False)

    if not user_id_str:
        await message.answer("Немає вибраного користувача.", reply_markup=get_admin_moderation_menu())
        return

    if text == "Назад":
        # Повертаємось до списку схвалених
        users_data = load_users()
        approved = users_data.get("approved_users", {})
        if not approved:
            # Порожньо
            await message.answer("Наразі немає схвалених користувачів.", reply_markup=get_admin_moderation_menu())
            await AdminMenuStates.moderation_section.set()
            return

        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        row = []
        approved_items = list(approved.items())
        for i, (uid, info) in enumerate(approved_items, start=1):
            fname = info.get("fullname", f"ID:{uid}")
            btn_text = f"{fname} | {uid}"
            row.append(btn_text)
            if len(row) == 2:
                kb.row(*row)
                row = []
        if row:
            kb.row(*row)
        kb.add("Назад")

        await state.update_data(approved_dict=approved)
        await AdminReview.viewing_approved_list.set()
        await message.answer("Список схвалених користувачів:", reply_markup=kb)
        return

    elif text == "Редагувати":
        # Переходимо у стан редагування
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        kb.row("Змінити ПІБ", "Змінити номер телефону")
        kb.add("Назад")
        await message.answer("Оберіть, що бажаєте змінити:", reply_markup=kb)
        await AdminReview.editing_approved_user.set()

    elif text == "Видалити":
        # Повністю видаляємо користувача з approved_users
        users_data = load_users()
        if user_id_str in users_data.get("approved_users", {}):
            users_data["approved_users"].pop(user_id_str)
            save_users(users_data)
            await message.answer("Користувача видалено із схвалених.", reply_markup=get_admin_moderation_menu())
        else:
            await message.answer("Користувача не знайдено у схвалених.", reply_markup=get_admin_moderation_menu())

        await AdminMenuStates.moderation_section.set()

    else:
        await message.answer("Оберіть: «Редагувати», «Видалити» або «Назад».")


#
# 1.2.1) Редагування схваленого користувача
#

@dp.message_handler(state=AdminReview.editing_approved_user)
async def admin_edit_approved_user_menu(message: types.Message, state: FSMContext):
    """
    Меню вибору, що саме редагувати (ПІБ чи номер).
    """
    text = message.text.strip()

    if text == "Назад":
        # Повторно відображаємо поточного користувача
        data = await state.get_data()
        user_id_str = data.get("selected_approved_user_id", None)
        if user_id_str is None:
            await message.answer("Немає користувача. Повернення.", reply_markup=get_admin_moderation_menu())
            await AdminMenuStates.moderation_section.set()
            return

        users_data = load_users()
        user_info = users_data.get("approved_users", {}).get(user_id_str, {})
        fullname = user_info.get("fullname", "—")
        phone = user_info.get("phone", "—")
        details = (
            f"ПІБ: {fullname}\n"
            f"Номер телефону: {phone}\n"
            f"Телеграм ID: {user_id_str}"
        )

        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        kb.row("Редагувати", "Видалити")
        kb.add("Назад")
        await AdminReview.viewing_approved_user.set()
        await message.answer(details, reply_markup=kb)
        return

    elif text == "Змінити ПІБ":
        await AdminReview.editing_approved_user_fullname.set()
        await message.answer("Введіть новий ПІБ:", reply_markup=remove_keyboard())

    elif text == "Змінити номер телефону":
        await AdminReview.editing_approved_user_phone.set()
        await message.answer("Введіть новий номер телефону у форматі +380XXXXXXXXX:", reply_markup=remove_keyboard())

    else:
        await message.answer("Оберіть 'Змінити ПІБ', 'Змінити номер телефону' або 'Назад'.")


@dp.message_handler(state=AdminReview.editing_approved_user_fullname)
async def admin_edit_approved_user_fullname(message: types.Message, state: FSMContext):
    """
    Зміна ПІБ у approved_users.
    """
    new_fullname = message.text.strip()
    if not new_fullname:
        await message.answer("ПІБ не може бути порожнім. Спробуйте ще раз.")
        return

    data = await state.get_data()
    user_id_str = data.get("selected_approved_user_id", None)
    if user_id_str is None:
        await message.answer("Немає користувача для редагування.", reply_markup=get_admin_moderation_menu())
        await AdminMenuStates.moderation_section.set()
        return

    users_data = load_users()
    if user_id_str not in users_data.get("approved_users", {}):
        await message.answer("Користувача не знайдено в approved_users.", reply_markup=get_admin_moderation_menu())
        await AdminMenuStates.moderation_section.set()
        return

    # Оновлюємо ПІБ
    users_data["approved_users"][user_id_str]["fullname"] = new_fullname
    save_users(users_data)

    # Повертаємось у меню редагування
    await message.answer("ПІБ успішно змінено!", reply_markup=remove_keyboard())

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.row("Змінити ПІБ", "Змінити номер телефону")
    kb.add("Назад")

    await AdminReview.editing_approved_user.set()
    await message.answer("Оновлено! Оберіть наступну дію:", reply_markup=kb)


@dp.message_handler(state=AdminReview.editing_approved_user_phone)
async def admin_edit_approved_user_phone(message: types.Message, state: FSMContext):
    """
    Зміна номера телефону у approved_users.
    """
    new_phone = re.sub(r"[^\d+]", "", message.text.strip())
    if not re.fullmatch(r"\+380\d{9}", new_phone):
        await message.answer("Невірний формат. Введіть номер у форматі +380XXXXXXXXX або «Назад» для відміни.")
        return

    data = await state.get_data()
    user_id_str = data.get("selected_approved_user_id", None)
    if user_id_str is None:
        await message.answer("Немає користувача для редагування.", reply_markup=get_admin_moderation_menu())
        await AdminMenuStates.moderation_section.set()
        return

    users_data = load_users()
    if user_id_str not in users_data.get("approved_users", {}):
        await message.answer("Користувача не знайдено в approved_users.", reply_markup=get_admin_moderation_menu())
        await AdminMenuStates.moderation_section.set()
        return

    # Оновлюємо телефон
    users_data["approved_users"][user_id_str]["phone"] = new_phone
    save_users(users_data)

    await message.answer("Номер телефону успішно змінено!", reply_markup=remove_keyboard())

    # Повертаємось у меню редагування
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.row("Змінити ПІБ", "Змінити номер телефону")
    kb.add("Назад")

    await AdminReview.editing_approved_user.set()
    await message.answer("Оновлено! Оберіть наступну дію:", reply_markup=kb)



############################################
# 2) РОЗДІЛ «ЗАЯВКИ»
############################################

@dp.message_handler(state=AdminMenuStates.requests_section)
async def admin_requests_section_handler(message: types.Message, state: FSMContext):
    text = message.text.strip()

    if text == "Підтверджені":
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
            await message.answer("Немає підтверджених заявок.", reply_markup=get_admin_requests_menu())
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

        await state.update_data(confirmed_apps=confirmed_apps, from_requests_menu=True)
        await AdminReview.viewing_confirmed_list.set()
        await message.answer("Список підтверджених заявок:", reply_markup=kb)

    elif text == "Видалені":
        apps = load_applications()
        deleted_apps = []
        for user_id, user_applications in apps.items():
            for idx, app_data in enumerate(user_applications):
                if app_data.get("proposal_status") == "deleted":
                    deleted_apps.append({
                        "user_id": user_id,
                        "app_index": idx,
                        "app_data": app_data
                    })

        if not deleted_apps:
            await message.answer("Немає видалених заявок.", reply_markup=get_admin_requests_menu())
            return

        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        row = []
        for i, entry in enumerate(deleted_apps, start=1):
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

        await state.update_data(deleted_apps=deleted_apps, from_requests_menu=True)
        await AdminReview.viewing_deleted_list.set()
        await message.answer("Список «видалених» заявок:", reply_markup=kb)

    elif text == "Переглянути заявки культур":
        await message.answer("Функціонал «Переглянути заявки культур» ще не реалізовано.", reply_markup=get_admin_requests_menu())

    elif text == "Назад":
        await message.answer("Головне меню адміна:", reply_markup=get_admin_root_menu())
        await AdminMenuStates.choosing_section.set()
    else:
        await message.answer("Оберіть дію: «Підтверджені», «Видалені», «Переглянути заявки культур» або «Назад».")

############################################
# ПЕРЕГЛЯД «ПІДТВЕРДЖЕНИХ»
############################################

@dp.message_handler(state=AdminReview.viewing_confirmed_list)
async def admin_view_confirmed_list_choice(message: types.Message, state: FSMContext):
    data = await state.get_data()
    confirmed_apps = data.get("confirmed_apps", [])
    from_requests_menu = data.get("from_requests_menu", False)

    if message.text == "Назад":
        if from_requests_menu:
            await message.answer("Розділ 'Заявки':", reply_markup=get_admin_requests_menu())
            await AdminMenuStates.requests_section.set()
        else:
            await state.finish()
            await message.answer("Адмін меню:", reply_markup=get_admin_root_menu())
        return

    split_msg = message.text.split('.', 1)
    if len(split_msg) < 2 or not split_msg[0].isdigit():
        await message.answer("Оберіть номер заявки у форматі 'X. Культура | Кількість' або натисніть 'Назад'.")
        return

    choice = int(split_msg[0])
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
        f"Пропозиція ціни: <b>{app_data.get('proposal', '')}</b>",
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

@dp.message_handler(state=AdminReview.viewing_confirmed_app)
async def admin_view_confirmed_app_handler(message: types.Message, state: FSMContext):
    data = await state.get_data()
    selected_entry = data.get("selected_confirmed")
    confirmed_apps = data.get("confirmed_apps", [])
    chosen_index = data.get("chosen_confirmed_index")
    if not selected_entry or chosen_index is None:
        await message.answer("Немає заявки для опрацювання.", reply_markup=get_admin_requests_menu())
        await state.finish()
        return

    if message.text == "Назад":
        if not confirmed_apps:
            await state.finish()
            await message.answer("Список підтверджених заявок тепер порожній.", reply_markup=get_admin_requests_menu())
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
        return

    elif message.text == "Видалити":
        user_id = int(selected_entry["user_id"])
        app_index = selected_entry["app_index"]
        update_application_status(user_id, app_index, "deleted")
        if 0 <= chosen_index < len(confirmed_apps):
            confirmed_apps.pop(chosen_index)
        await state.update_data(confirmed_apps=confirmed_apps, selected_confirmed=None, chosen_confirmed_index=None)
        await message.answer("Заявка перенесена у 'видалені'.", reply_markup=get_admin_requests_menu())
        await AdminMenuStates.requests_section.set()

    else:
        await message.answer("Оберіть «Видалити» або «Назад».")

############################################
# ПЕРЕГЛЯД «ВИДАЛЕНИХ»
############################################

@dp.message_handler(state=AdminReview.viewing_deleted_list)
async def admin_view_deleted_list_choice(message: types.Message, state: FSMContext):
    data = await state.get_data()
    deleted_apps = data.get("deleted_apps", [])
    from_requests_menu = data.get("from_requests_menu", False)

    if message.text == "Назад":
        if from_requests_menu:
            await message.answer("Розділ 'Заявки':", reply_markup=get_admin_requests_menu())
            await AdminMenuStates.requests_section.set()
        else:
            await state.finish()
            await message.answer("Адмін меню:", reply_markup=get_admin_root_menu())
        return

    split_msg = message.text.split('.', 1)
    if len(split_msg) < 2 or not split_msg[0].isdigit():
        await message.answer("Оберіть номер заявки у форматі 'X. Культура | Кількість' або натисніть 'Назад'.")
        return

    choice = int(split_msg[0])
    if choice < 1 or choice > len(deleted_apps):
        await message.answer("Невірний вибір.", reply_markup=remove_keyboard())
        return

    selected_entry = deleted_apps[choice - 1]
    app_data = selected_entry["app_data"]
    timestamp = app_data.get("timestamp", "")
    try:
        dt = datetime.fromisoformat(timestamp)
        formatted_date = dt.strftime("%d.%m.%Y")
    except Exception:
        formatted_date = timestamp

    details = [
        "<b>«ВИДАЛЕНА» ЗАЯВКА:</b>",
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
        f"Пропозиція ціни: <b>{app_data.get('proposal', '')}</b>",
        "\nЦя заявка позначена як «deleted»."
    ]

    extra = app_data.get("extra_fields", {})
    if extra:
        details.append("Додаткові параметри:")
        for key, value in extra.items():
            details.append(f"{friendly_names.get(key, key.capitalize())}: {value}")

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add("Видалити назавжди", "Назад")

    await state.update_data(selected_deleted=selected_entry, chosen_deleted_index=choice - 1)
    await AdminReview.viewing_deleted_app.set()
    await message.answer("\n".join(details), parse_mode="HTML", reply_markup=kb)

@dp.message_handler(state=AdminReview.viewing_deleted_app)
async def admin_view_deleted_app_handler(message: types.Message, state: FSMContext):
    data = await state.get_data()
    selected_entry = data.get("selected_deleted")
    deleted_apps = data.get("deleted_apps", [])
    chosen_index = data.get("chosen_deleted_index")
    if not selected_entry or chosen_index is None:
        await message.answer("Немає заявки для опрацювання.", reply_markup=get_admin_requests_menu())
        await state.finish()
        return

    if message.text == "Назад":
        if not deleted_apps:
            await message.answer("Список видалених заявок порожній.", reply_markup=get_admin_requests_menu())
            await AdminMenuStates.requests_section.set()
            return

        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        row = []
        for i, entry in enumerate(deleted_apps, start=1):
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

        await AdminReview.viewing_deleted_list.set()
        await message.answer("Список видалених заявок:", reply_markup=kb)
        return

    elif message.text == "Видалити назавжди":
        user_id = int(selected_entry["user_id"])
        app_index = selected_entry["app_index"]

        success = await admin_remove_app_permanently(user_id, app_index)
        if success:
            if 0 <= chosen_index < len(deleted_apps):
                deleted_apps.pop(chosen_index)
            await state.update_data(deleted_apps=deleted_apps, selected_deleted=None, chosen_deleted_index=None)
            await message.answer("Заявку остаточно видалено з файлу та таблиць.", reply_markup=get_admin_requests_menu())
        else:
            await message.answer("Помилка: Заявка не знайдена або вже була видалена.", reply_markup=get_admin_requests_menu())

        await AdminMenuStates.requests_section.set()

    else:
        await message.answer("Оберіть «Видалити назавжди» або «Назад».")

############################################
# ОБРОБНИК /start (РЕЄСТРАЦІЯ КОРИСТУВАЧА)
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

@dp.message_handler(state=RegistrationStates.waiting_for_fullname)
async def process_fullname(message: types.Message, state: FSMContext):
    fullname = message.text.strip()
    if not fullname:
        await message.answer("ПІБ не може бути порожнім. Введіть коректне значення.")
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
    await show_registration_preview(message, state)

@dp.message_handler(state=RegistrationStates.waiting_for_phone)
async def process_phone_text(message: types.Message, state: FSMContext):
    phone = re.sub(r"[^\d+]", "", message.text.strip())
    if not re.fullmatch(r"\+380\d{9}", phone):
        await message.answer("Невірний формат. Введіть номер у форматі +380XXXXXXXXX")
        return
    await state.update_data(phone=phone)
    await show_registration_preview(message, state)

async def show_registration_preview(message: types.Message, state: FSMContext):
    data = await state.get_data()
    fullname = data.get("fullname", "—")
    phone = data.get("phone", "—")

    preview_text = (
        "<b>Перевірте свої дані:</b>\n\n"
        f"ПІБ: {fullname}\n"
        f"Телефон: {phone}\n\n"
        "Якщо все вірно, натисніть <b>Підтвердити</b>."
    )

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.row("Підтвердити", "Редагувати", "Скасувати")

    await message.answer(preview_text, parse_mode="HTML", reply_markup=kb)
    await RegistrationStates.preview.set()

@dp.message_handler(Text(equals="Підтвердити"), state=RegistrationStates.preview)
async def confirm_registration_preview(message: types.Message, state: FSMContext):
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

@dp.message_handler(Text(equals="Редагувати"), state=RegistrationStates.preview)
async def edit_registration_preview(message: types.Message, state: FSMContext):
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.row("Змінити ПІБ", "Змінити номер телефону")
    kb.add("Назад")
    await message.answer("Оберіть, що змінити:", reply_markup=kb)
    await RegistrationStates.editing.set()

@dp.message_handler(Text(equals="Скасувати"), state=RegistrationStates.preview)
async def cancel_registration_preview(message: types.Message, state: FSMContext):
    await state.finish()
    await message.answer("Реєстрацію скасовано. Якщо передумаєте – введіть /start заново.",
                         reply_markup=remove_keyboard())

@dp.message_handler(Text(equals="Змінити ПІБ"), state=RegistrationStates.editing)
async def editing_fullname_button(message: types.Message, state: FSMContext):
    await message.answer("Введіть нове ПІБ:", reply_markup=remove_keyboard())
    await RegistrationStates.editing_fullname.set()

@dp.message_handler(state=RegistrationStates.editing_fullname)
async def process_editing_fullname(message: types.Message, state: FSMContext):
    new_fullname = message.text.strip()
    if not new_fullname:
        await message.answer("ПІБ не може бути порожнім.")
        return
    await state.update_data(fullname=new_fullname)
    await return_to_editing_menu(message, state)

@dp.message_handler(Text(equals="Змінити номер телефону"), state=RegistrationStates.editing)
async def editing_phone_button(message: types.Message, state: FSMContext):
    await message.answer("Введіть новий номер телефону (+380XXXXXXXXX):", reply_markup=remove_keyboard())
    await RegistrationStates.editing_phone.set()

@dp.message_handler(state=RegistrationStates.editing_phone)
async def process_editing_phone(message: types.Message, state: FSMContext):
    phone = re.sub(r"[^\d+]", "", message.text.strip())
    if not re.fullmatch(r"\+380\d{9}", phone):
        await message.answer("Невірний формат. Введіть номер у форматі +380XXXXXXXXX")
        return
    await state.update_data(phone=phone)
    await return_to_editing_menu(message, state)

@dp.message_handler(Text(equals="Назад"), state=RegistrationStates.editing)
async def back_to_preview_from_editing(message: types.Message, state: FSMContext):
    await show_registration_preview(message, state)

async def return_to_editing_menu(message: types.Message, state: FSMContext):
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.row("Змінити ПІБ", "Змінити номер телефону")
    kb.add("Назад")
    await RegistrationStates.editing.set()
    await message.answer("Оновлено! Що бажаєте змінити далі?", reply_markup=kb)

############################################
# /menu
############################################

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

############################################
# /support
############################################

@dp.message_handler(commands=["support"], state="*")
async def support_command(message: types.Message, state: FSMContext):
    keyboard = types.InlineKeyboardMarkup()
    # Замініть текст кнопки, якщо потрібно, та посилання на бота
    keyboard.add(types.InlineKeyboardButton("Звернутись до підтримки", url="https://t.me/Dealeragro_bot"))
    await message.answer("Якщо вам потрібна допомога, натисніть кнопку нижче:", reply_markup=keyboard)

############################################
# Подати заявку / Переглянути мої заявки
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

@dp.message_handler(Text(equals="Переглянути мої заявки"), state="*")
async def show_user_applications(message: types.Message):
    user_id = message.from_user.id
    uid = str(user_id)
    apps = load_applications()
    user_apps = apps.get(uid, [])

    if not user_apps:
        await message.answer("Ви не маєте заявок.", reply_markup=get_main_menu_keyboard())
        return

    buttons = []
    for i, app in enumerate(user_apps, start=1):
        culture = app.get('culture', 'Невідомо')
        quantity = app.get('quantity', 'Невідомо')
        status = app.get("proposal_status", "")
        if status == "confirmed":
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
# ДЕТАЛЬНИЙ ПЕРЕГЛЯД ЗАЯВКИ
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
    elif status == "deleted":
        details.append("\nЦя заявка вже позначена як 'deleted' (видалена).")
        kb.add("Назад")

    kb.add("Назад")

    await state.update_data(selected_app_index=idx)
    await message.answer("\n".join(details), reply_markup=kb, parse_mode="HTML")
    await ApplicationStates.viewing_application.set()

############################################
# "ПЕРЕГЛЯНУТИ ПРОПОЗИЦІЮ"
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
# "ВІДХИЛИТИ"
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
        color_cell_red(sheet_row)

    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.row("Видалити", "Очікувати")
    await message.answer("Пропозицію відхилено. Оберіть: Видалити заявку або Очікувати кращу пропозицію?",
                         reply_markup=kb)
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

@dp.message_handler(Text(equals="Видалити"), state=ApplicationStates.proposal_reply)
async def delete_after_rejection(message: types.Message, state: FSMContext):
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
    if sheet_row:
        color_cell_red(sheet_row)

    delete_application_soft(message.from_user.id, index)
    await message.answer("Ваша заявка видалена (позначена як 'deleted').", reply_markup=get_main_menu_keyboard())
    await state.finish()

############################################
# "ПІДТВЕРДИТИ" (CONFIRMED)
############################################

@dp.message_handler(Text(equals="Підтвердити"), state=ApplicationStates.viewing_application)
async def confirm_proposal(message: types.Message, state: FSMContext):
    data = await state.get_data()
    index = data.get("selected_app_index")
    update_application_status(message.from_user.id, index, "confirmed")

    apps = load_applications()
    uid = str(message.from_user.id)
    app = apps[uid][index]

    sheet_row = app.get("sheet_row")
    if sheet_row:
        color_cell_green(sheet_row)

    save_applications(apps)

    timestamp = app.get("timestamp", "")
    try:
        dt = datetime.fromisoformat(timestamp)
        formatted_date = dt.strftime("%d.%m.%Y")
    except Exception:
        formatted_date = timestamp or "—"

    extra_fields = app.get("extra_fields", {})
    extra_list = []
    for key, value in extra_fields.items():
        ukr_name = friendly_names.get(key, key)
        extra_list.append(f"{ukr_name}: {value}")
    extra_part = ""
    if extra_list:
        extra_part = f"Додаткові параметри:\n<b>{chr(10).join(extra_list)}</b>\n"

    user_fullname = app.get("fullname", "")
    phone_from_app = app.get("phone", "")
    if not phone_from_app:
        users = load_users()
        phone_from_app = users.get("approved_users", {}).get(uid, {}).get("phone", "")
    if not phone_from_app:
        phone_from_app = "—"

    if not user_fullname:
        users = load_users()
        user_fullname = users.get("approved_users", {}).get(uid, {}).get("fullname", "—")

    user_fullname_line = f"Користувач: {user_fullname}"
    user_phone_line = f"Телефон: {phone_from_app}"

    admin_msg = (
        "<b>ЗАЯВКА ПІДТВЕРДЖЕНА</b>\n\n"
        "Повна інформація по заявці:\n"
        f"Дата створення: <b>{formatted_date}</b>\n\n"
        f"ФГ: <b>{app.get('fgh_name', 'Невідомо')}</b>\n"
        f"ЄДРПОУ: <b>{app.get('edrpou', 'Невідомо')}</b>\n"
        f"Область: <b>{app.get('region', 'Невідомо')}</b>\n"
        f"Район: <b>{app.get('district', 'Невідомо')}</b>\n"
        f"Місто: <b>{app.get('city', 'Невідомо')}</b>\n"
        f"Група: <b>{app.get('group', 'Невідомо')}</b>\n"
        f"Культура: <b>{app.get('culture', 'Невідомо')}</b>\n"
        f"{extra_part}"
        f"Кількість: <b>{app.get('quantity', 'Невідомо')} т</b>\n"
        f"Бажана ціна: <b>{app.get('price', 'Невідомо')}</b>\n"
        f"Валюта: <b>{app.get('currency', 'Невідомо')}</b>\n"
        f"Форма оплати: <b>{app.get('payment_form', 'Невідомо')}</b>\n"
        f"Пропозиція ціни: <b>{app.get('proposal', 'Невідомо')}</b>\n\n"
        f"{user_fullname_line}\n"
        f"{user_phone_line}"
    )

    for admin_id in ADMINS:
        try:
            await bot.send_message(admin_id, admin_msg)
        except Exception as e:
            logging.exception(f"Не вдалося відправити підтвердження адміну {admin_id}: {e}")

    await message.answer("Ви підтвердили пропозицію. Очікуйте на подальші дії від менеджера/адміністратора.",
                         reply_markup=get_main_menu_keyboard())
    await state.finish()

############################################
# ВИДАЛЕННЯ (КОРИСТУВАЧЕМ)
############################################

@dp.message_handler(Text(equals="Видалити"), state=ApplicationStates.viewing_application)
async def delete_request(message: types.Message, state: FSMContext):
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add("Так", "Ні")
    await message.answer("Ви впевнені, що хочете видалити заявку?", reply_markup=kb)
    await ApplicationStates.confirm_deletion.set()

@dp.message_handler(Text(equals="Так"), state=ApplicationStates.confirm_deletion)
async def confirm_deletion(message: types.Message, state: FSMContext):
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
    if sheet_row:
        color_cell_red(sheet_row)

    delete_application_soft(message.from_user.id, index)

    await message.answer("Ваша заявка видалена (позначена як 'deleted').", reply_markup=get_main_menu_keyboard())
    await state.finish()

@dp.message_handler(Text(equals="Ні"), state=ApplicationStates.confirm_deletion)
async def cancel_deletion(message: types.Message, state: FSMContext):
    await ApplicationStates.viewing_application.set()
    await message.answer("Видалення скасовано.", reply_markup=get_main_menu_keyboard())

############################################
# ГЛОБАЛЬНА "Назад"
############################################

@dp.message_handler(Text(equals="Назад"), state="*")
async def go_to_main_menu(message: types.Message, state: FSMContext):
    current_state = await state.get_state()

    if current_state and current_state.startswith("AdminReview:"):
        return

    if current_state and current_state.startswith("AdminMenuStates:"):
        return

    await state.finish()
    await message.answer("Головне меню:", reply_markup=get_main_menu_keyboard())

############################################
# ОБРОБНИК ДАНИХ ІЗ WEBAPP
############################################

@dp.message_handler(lambda message: message.text and "/webapp_data" in message.text, state=ApplicationStates.waiting_for_webapp_data)
async def webapp_data_handler_text(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    try:
        prefix = "/webapp_data "
        data_str = (
            message.text[len(prefix):].strip()
            if message.text.startswith(prefix)
            else message.text.split("/webapp_data", 1)[-1].strip()
        )
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

async def process_webapp_data_direct(user_id: int, data: dict, edit_index: int = None, sheet_row: int = None):
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
        await state.update_data(edit_index=edit_index, sheet_row=sheet_row, webapp_data=data)
        await state.set_state(ApplicationStates.editing_application.state)
    else:
        await state.update_data(webapp_data=data)
        await state.set_state(ApplicationStates.confirm_application.state)

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
    kb.add(types.KeyboardButton("Відкрити форму для редагування", web_app=types.WebAppInfo(url=url_with_data)))
    kb.row("Скасувати")
    await message.answer("Редагуйте заявку у WebApp:", reply_markup=kb)
    await state.set_state(ApplicationStates.waiting_for_webapp_data.state)

@dp.message_handler(Text(equals="Скасувати"), state=[
    ApplicationStates.waiting_for_webapp_data,
    ApplicationStates.confirm_application,
    ApplicationStates.editing_application
])
async def cancel_process_reply(message: types.Message, state: FSMContext):
    await state.finish()
    await message.answer("Процес скасовано. Головне меню:", reply_markup=get_main_menu_keyboard())

############################################
# ПІДТВЕРДЖЕННЯ ЗАЯВКИ КОРИСТУВАЧЕМ (ПІСЛЯ WEBAPP)
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
    global POLLING_PAUSED
    while True:
        if POLLING_PAUSED:
            await asyncio.sleep(3)
            continue

        try:
            ws = get_worksheet1()
            rows = ws.get_all_values()
            apps = load_applications()
            for i, row in enumerate(rows[1:], start=2):
                if len(row) < 15:
                    continue
                current_manager_price_str = row[13].strip()
                if not current_manager_price_str:
                    continue
                try:
                    cur_price = float(current_manager_price_str)
                except ValueError:
                    continue

                for uid, app_list in apps.items():
                    for idx, app in enumerate(app_list, start=1):
                        if app.get("sheet_row") == i:
                            status = app.get("proposal_status", "active")
                            if status in ("deleted", "confirmed"):
                                continue

                            original_manager_price_str = app.get("original_manager_price", "").strip()
                            try:
                                orig_price = float(original_manager_price_str) if original_manager_price_str else None
                            except:
                                orig_price = None

                            if orig_price is None:
                                culture = app.get("culture", "Невідомо")
                                quantity = app.get("quantity", "Невідомо")
                                app["original_manager_price"] = current_manager_price_str
                                app["proposal"] = current_manager_price_str
                                app["proposal_status"] = "Agreed"
                                await bot.send_message(
                                    app.get("chat_id"),
                                    f"Нова пропозиція по Вашій заявці {idx}. {culture} | {quantity} т. Ціна: {current_manager_price_str}"
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
                                            f"Ціна по заявці {idx}. {culture} | {quantity} т змінилась з {previous_proposal} на {current_manager_price_str}"
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
