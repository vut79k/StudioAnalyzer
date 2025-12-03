
import re
import time
from datetime import datetime, timedelta
from collections import defaultdict
from calendar import monthrange
import os 

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

import gspread
import gspread.exceptions
from oauth2client.service_account import ServiceAccountCredentials

# --- КОНСТАНТЫ ПРОЕКТА ---
STUDIO_ID = 21 
STUDIO_NAME = "Хохловка"
SELENIUM_WAIT_TIME = 5 
POPUP_WAIT_TIME = 5

FIN_ID = '1lgn068NObnej5A0J0Ya4Kxz5JY6tptxKf9dZTf2VXrI' 
KASSA_ID = '1gcM0hGf-D2s4-DhOYzpH6R-LCAe0MGRyKCUH3Fry1aQ'

# --- Google Sheets ---
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
try:
    creds = ServiceAccountCredentials.from_json_keyfile_name('google_key.json', scope)
    client = gspread.authorize(creds)
except Exception as e:
    print(f"Ошибка инициализации gspread: {e}")
    exit()

# Локализованные названия месяцев (без изменений)
months_ru = {
    'January': 'Январь','February': 'Февраль','March': 'Март','April': 'Апрель',
    'May': 'Май','June': 'Июнь','July': 'Июль','August': 'Август',
    'September': 'Сентябрь','October': 'Октябрь','November': 'Ноябрь','December': 'Декабрь'
}

months_ru_genitive = {
    'января': 'январь', 'февраля': 'февраль', 'марта': 'март', 'апреля': 'апрель',
    'мая': 'май', 'июня': 'июнь', 'июля': 'июль', 'августа': 'август',
    'сентября': 'сентябрь', 'октября': 'октябрь', 'ноября': 'ноябрь', 'декабря': 'декабрь'
}

months_ru_inv = {v.lower(): k for k, v in months_ru.items()}

def column_to_letter(n: int) -> str:
    result = []
    while n > 0:
        n -= 1
        result.append(chr(n % 26 + ord('A')))
        n //= 26
    return ''.join(reversed(result))

# --- Canonical mapping (оставляем без изменений) ---
RAW_MAPPING = {
    'фотосъемка': 'photo', 'банкет': 'banquet', 'видео съемки/мастер класс': 'video_master',
    'мероприятие': 'event', 'корпоративные клиенты': 'corporate', 'фотошкола занятия': 'school_class',
    'фотошкола домашние работы студентов': 'school_homework', 'плавающая бронь': 'floating',
    'не приехали/не приедут': 'no_show', 'тех.бронь': 'tech', 'мероприятие бланк': 'event_blank',
    'мероприятие сися и white studios': 'event_sisia_white', 'мероприятия yauza_place': 'yauza_place',
    'мероприятия crystal': 'crystal'
}
DOP_KEYS = ['парковк','цикл','циклорама','фон','улице','парк', 'отпар', 'аренда', 'улиц', 'раннее', 'позднее', 'стойк', 'источник']
ORDERED_KEYS = sorted(RAW_MAPPING.keys(), key=lambda s: len(s), reverse=True)

# Маппинг категорий и строк в Google Sheet
HOUR_ROW_MAPPING = {
    'photo': 46, 'banquet': 47, 'video_master': 48, 'event': 49, 'corporate': 50,
    'school_class': 51, 'school_homework': 52, 'floating': 53, 'no_show': 54,
    'tech': 55, 'event_blank': 56, 'event_sisia_white': 57, 'yauza_place': 58,
    'crystal': 59, 'unknown': 60 
}

# --- Parsers and classifiers (оставляем без изменений) ---
def extract_start_end(text: str):
    t = re.sub(r'\s+', ' ', text).lower()
    m = re.search(r'(?:с|c)\s*(\d{1,2}:\d{2}).*?до\s*(\d{1,2}:\d{2})', t)
    if not m: return None, None
    s, e = m.group(1), m.group(2)
    if len(s) == 4: s = '0' + s
    if len(e) == 4: e = '0' + e
    return s, e

def extract_declared_hours_and_nextline(popup_text: str):
    lines = [ln.strip() for ln in popup_text.splitlines() if ln.strip()]
    for i, ln in enumerate(lines):
        low = ln.lower()
        m = re.search(r'кол-?во\s*часов[:\s]*([\d]+)', low) or re.search(r'количество\s+часов[:\s]*([\d]+)', low)
        if m:
            hours = None
            try: hours = int(m.group(1))
            except: hours = None
            next_line = lines[i+1].strip().lower() if i+1 < len(lines) else None
            return hours, next_line
    return None, None

def extract_prepaid(popup_text: str):
    lines = [ln.strip() for ln in popup_text.splitlines() if ln.strip()]
    for ln in lines:
        low = ln.lower()
        if 'итого оплачено' in low:
            m = re.search(r'(\d+)\s*руб\.', low)
            if m: return int(m.group(1))
    return 0

def extract_booking_date(popup_text: str):
    lines = [ln.strip() for ln in popup_text.splitlines() if ln.strip()]
    for ln in lines:
        low = ln.lower()
        m = re.search(r'дата:\s*(\d{1,2})\s*(\w+)\s*(\d{4})', low)
        if m:
            day = int(m.group(1))
            month_str = m.group(2).lower()
            year = int(m.group(3))
            if month_str in months_ru_genitive: month_str = months_ru_genitive[month_str]
            try:
                month = list(months_ru_inv.keys()).index(month_str) + 1
                return datetime(year, month, day)
            except ValueError:
                return None
    return None

def classify_from_text(popup_text: str):
    low = popup_text.lower()
    _, next_line = extract_declared_hours_and_nextline(popup_text)
    if next_line:
        for k in ORDERED_KEYS:
            if k in next_line: return RAW_MAPPING[k]
    for k in ORDERED_KEYS:
        if k in low: return RAW_MAPPING[k]
    for k in DOP_KEYS:
        if k in low: return 'dop'
    if any(x in low for x in ['видео','мастер','мастеркласс']): return 'video_master'
    if any(x in low for x in ['школ','занятия','домашние']): return 'school_homework' if 'домаш' in low or 'домашние' in low else 'school_class'
    if any(x in low for x in ['корпор','корп']): return 'corporate'
    if any(x in low for x in ['мероприяти','yauza','crystal','сися','white studios']): return 'event_sisia_white' if any(y in low for y in ['сися','white studios']) else 'event'
    if 'floating' in low: return 'floating'
    if 'no_show' in low: return 'no_show'
    if 'tech' in low: return 'tech'
    return 'unknown'

# --- ФУНКЦИИ GOOGLE SHEETS API ---
def update_sheet_batch(itogi_sheet, col, daily_data_item, daily_hours_totals):
    update_requests = []
    # ИЗМЕНЕННЫЙ МАППИНГ ДЛЯ ФИН. ТАБЛИЦЫ ХОХЛОВКИ (строки 5-15)
    financial_mapping = {
        'prep_photo': 5, 'fakt_photo': 7, 'prep_video': 6, 
        'fakt_video': 8, 'school': 13, 'dop': 15
    }
    for key, row_num in financial_mapping.items():
        update_requests.append({
            'range': f"{col}{row_num}",
            'values': [[daily_data_item[key]]]
        })
    for key, row_num in HOUR_ROW_MAPPING.items():
        hours = round(daily_hours_totals.get(key, 0.0), 2)
        update_requests.append({
            'range': f"{col}{row_num}",
            'values': [[hours]]
        })
    if update_requests:
        itogi_sheet.batch_update([{
            'range': req['range'],
            'values': req['values']
        } for req in update_requests])


# --- ОСНОВНАЯ ЛОГИКА ---
try:
    # Инициализация браузера
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    # Login
    driver.get("https://whitestudios.ru/reservator/login/")
    login_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "login")))
    password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "pass")))
    login_field.send_keys("Kapr")
    password_field.send_keys("vut79k")
    driver.find_element(By.CSS_SELECTOR, 'input[type="submit"]').click()
    time.sleep(5)

    main_url = driver.current_url

    # Request period
    print(f"--- АВТОМАТИЗАЦИЯ ДЛЯ СТУДИИ: {STUDIO_NAME} ---")
    period_input = input("За какое число (формат: dd mm yyyy для дня, mm yyyy для месяца, dd-dd mm yyyy для диапазона, dd.mm.yyyy-dd.mm.yyyy для кросс-месяца): ").strip()

    # --- ЛОГИКА ПАРСИНГА ДАТ (без изменений) ---
    days = []
    if '-' in period_input and '.' in period_input:
        start_str, end_str = period_input.split('-')
        start_day, start_month, start_year = map(int, start_str.split('.'))
        end_day, end_month, end_year = map(int, end_str.split('.'))
        current_date = datetime(start_year, start_month, start_day)
        end_date = datetime(end_year, end_month, end_day)
        while current_date <= end_date:
            days.append((current_date.day, current_date.month, current_date.year))
            current_date += timedelta(days=1)
    else:
        parts = period_input.split()
        if len(parts) == 3 and '-' in parts[0]:
            start_day, end_day = map(int, parts[0].split('-'))
            month = int(parts[1])
            year = int(parts[2])
            for day in range(start_day, end_day + 1):
                days.append((day, month, year))
        elif len(parts) == 2:
            month = int(parts[0])
            year = int(parts[1])
            _, last_day = monthrange(year, month)
            for day in range(1, last_day + 1):
                days.append((day, month, year))
        elif len(parts) == 3:
            day = int(parts[0])
            month = int(parts[1])
            year = int(parts[2])
            days.append((day, month, year))
        else:
            print("Неверный формат. Выход.")
            exit()

    # Общие часы
    total_hours_totals = defaultdict(float)

    # Daily data storage for batch write
    daily_data = {}

    # Cache kassa per month
    kassa_cache = {}
    
    # Open workbook once
    fin_workbook = client.open_by_key(FIN_ID)
    kassa_workbook = client.open_by_key(KASSA_ID)

    # Process days
    for day, month, year in days:
        current_date_dt = datetime(year=year, month=month, day=day)
        current_date_str = f"{day:02d}.{month:02d}.{year}"
        print(f"Processing day {current_date_str}")
        
        # ИНИЦИАЛИЗАЦИЯ ДАННЫХ ДЛЯ ДНЯ
        itogi_sheet = None
        kassa_data = [] 
        daily_hours_totals = defaultdict(float)
        daily_counts = defaultdict(int)
        daily_total_prepayment_photo = 0
        daily_total_prepayment_video = 0
        daily_total_fakt_photo = 0
        daily_total_fakt_video = 0
        daily_total_fakt_dop = 0
        daily_parking_amount = 0
        daily_parking_count = 0
        
        # --- СЕКЦИЯ SELENIUM (ПАРСИНГ БРОНЕЙ) ---
        
        # 1. Загрузка расписания студии
        retry = 0
        success = False
        while retry < 3 and not success:
            try:
                driver.execute_script(f"showStudio({STUDIO_ID}, '{current_date_str}');")
                time.sleep(SELENIUM_WAIT_TIME) 

                studio_body = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, f"studioBody{STUDIO_ID}")))
                WebDriverWait(driver, 20).until(lambda x: len(studio_body.text.strip()) > 0)

                success = True

            except Exception as e:
                print(f"Retry {retry+1} for load day {current_date_str} on {STUDIO_NAME}: {e}")
                retry += 1
                driver.refresh()
                time.sleep(SELENIUM_WAIT_TIME)

        if not success:
            print(f"Failed to load day {current_date_str} after 3 retries. Skip processing bookings.")
            
            # Даже если не загрузилось, нужно определить itogi_sheet для потенциальной записи 0 (или чтобы не крашнулось в конце)
            month_en = current_date_dt.strftime('%B')
            sheet_name = f'{months_ru[month_en]}{str(year)[-2:]}'
            try:
                itogi_sheet = fin_workbook.worksheet(sheet_name)
            except gspread.exceptions.WorksheetNotFound:
                itogi_sheet = fin_workbook.add_worksheet(title=sheet_name, rows=100, cols=50)
                time.sleep(3) 

        else: # Если брони загружены успешно
            # Collect all bookings
            bookings = driver.find_elements(By.CSS_SELECTOR, f'#studioBody{STUDIO_ID} .reserved') 
            
            hrefs = set()
            for booking in bookings:
                try:
                    link = booking.find_element(By.TAG_NAME, "a")
                    href = link.get_attribute("href")
                    if href:
                        hrefs.add(href)
                except:
                    continue

            # Fact totals initialization (needed if scraping succeeds)
            daily_total_fakt_photo = 0
            daily_total_fakt_video = 0
            daily_total_fakt_dop = 0
            daily_parking_amount = 0
            daily_parking_count = 0

            for href in hrefs:
                try:
                    driver.get(href)
                    time.sleep(POPUP_WAIT_TIME) 

                    popup_text = driver.find_element(By.TAG_NAME, "body").text
                    
                    if STUDIO_NAME not in popup_text:
                        driver.get(main_url)
                        time.sleep(1)
                        continue

                    cls = classify_from_text(popup_text)

                    if cls == "unknown" and "Не приехали" in popup_text:
                        cls = "no_show"

                    if cls == "unknown":
                        print(f"Бронь {href} не классифицирована. Пропуск.")
                        driver.get(main_url)
                        time.sleep(1)
                        continue


                    # Hours
                    declared_hours, next_line = extract_declared_hours_and_nextline(popup_text)
                    hours = declared_hours if declared_hours else 0

                    # --- ИСПРАВЛЕННЫЙ БЛОК ДЛЯ РАСЧЕТА ЧАСОВ В ПРЕДЕЛАХ ДНЯ (как для Яузы) ---
                    start_str, end_str = extract_start_end(popup_text)
                    full_hours = 0.0
                    hours_in_day = 0.0
                    
                    if start_str and end_str:
                        start_t = datetime.strptime(start_str, '%H:%M').time()
                        end_t = datetime.strptime(end_str, '%H:%M').time()
                    
                        booking_date = extract_booking_date(popup_text)
                        
                        if not booking_date:
                            booking_date = current_date_dt 
                        
                        actual_start = datetime.combine(booking_date.date(), start_t)
                        actual_end = datetime.combine(booking_date.date(), end_t)

                        # Корректируем на ночную бронь (если время конца раньше времени начала, и это не 24-часовая бронь 00:00-00:00)
                        if actual_end <= actual_start:
                            actual_end += timedelta(days=1)
                        
                        full_hours = (actual_end - actual_start).total_seconds() / 3600

                        # Определяем интервал 24 часа для текущего обрабатываемого дня
                        day_start_dt = datetime.combine(current_date_dt.date(), datetime.min.time()) # 00:00 текущего дня
                        day_end_dt = day_start_dt + timedelta(days=1) # 00:00 следующего дня

                        # Расчет пересечения интервалов
                        overlap_start = max(actual_start, day_start_dt)
                        overlap_end = min(actual_end, day_end_dt)

                        if overlap_end > overlap_start:
                            hours_in_day = (overlap_end - overlap_start).total_seconds() / 3600
                        else:
                            hours_in_day = 0.0

                        # Округляем до целого или одного знака после запятой, как ранее
                        hours_in_day = round(hours_in_day) 
                    else:
                        full_hours = hours
                        hours_in_day = hours
                    # --- КОНЕЦ ИСПРАВЛЕННОГО БЛОКА ---


                    daily_hours_totals[cls] += hours_in_day
                    daily_counts[cls] += 1
                    total_hours_totals[cls] += hours_in_day

                    print(f"Бронь {href}: {cls}, {round(hours_in_day, 1)} ч (full: {round(full_hours, 1) if full_hours > 0 else 'N/A'} ч)")

                    # Prepayment
                    prepayment = extract_prepaid(popup_text)
                    if cls == 'photo':
                        daily_total_prepayment_photo += prepayment
                    elif cls == 'video_master':
                        daily_total_prepayment_video += prepayment

                    driver.get(main_url)
                    time.sleep(1)

                except Exception as e:
                    print(f"Ошибка при обработке брони {href}: {e}")
                    driver.get(main_url)
                    time.sleep(1)
                    continue

            # Get Itogi Sheet (для успешного парсинга)
            month_en = current_date_dt.strftime('%B')
            sheet_name = f'{months_ru[month_en]}{str(year)[-2:]}'
            try:
                itogi_sheet = fin_workbook.worksheet(sheet_name)
            except gspread.exceptions.WorksheetNotFound:
                print(f"Лист итогов '{sheet_name}' не найден. Создание.")
                itogi_sheet = fin_workbook.add_worksheet(title=sheet_name, rows=100, cols=50)
                time.sleep(3)
        
        # --- СЕКЦИЯ GOOGLE SHEETS (КАССА) ---

        month_en = current_date_dt.strftime('%B')
        month_ru_lower = months_ru[month_en].lower()
        
        # Имя листа Кассы Хохловка - 'октябрь 2025'
        kassa_sheet_name = f'{month_ru_lower} {year}' 

        # Попытка загрузить Кассу
        if kassa_sheet_name not in kassa_cache:
            try:
                kassa_sheet = kassa_workbook.worksheet(kassa_sheet_name)
                kassa_cache[kassa_sheet_name] = kassa_sheet.get_all_values()
                print(f"Лист кассы '{kassa_sheet_name}' успешно загружен.")
                time.sleep(2) 
                kassa_data = kassa_cache[kassa_sheet_name]
            except gspread.exceptions.WorksheetNotFound:
                print(f"Лист кассы '{kassa_sheet_name}' не найден. Данные по Кассе не будут учтены.")
                kassa_data = []
            except Exception as e:
                print(f"Ошибка чтения кассы '{kassa_sheet_name}': {e}. Данные по Кассе не будут учтены.")
                kassa_data = []
        else:
             kassa_data = kassa_cache[kassa_sheet_name]
             print(f"Лист кассы '{kassa_sheet_name}' загружен из кэша.")

        # Kassa processing
        for row in kassa_data[1:]:  # Skip header
            if len(row) < 4: continue
            date_cell = row[0].strip()
            if date_cell != current_date_str:
                continue

            amount_b = row[1].strip()
            amount_d = row[3].strip()
            amount_str = amount_b if amount_b.startswith('р.') else amount_d if amount_d.startswith('р.') else ''
            if amount_str:
                try:
                    amount = int(re.sub(r'[^\d]', '', amount_str))
                except:
                    continue
            else:
                continue

            if amount < 0:
                continue  # Skip negative

            desc = ' '.join(row[4:]).lower()
            if 'хохловка' not in desc:
                continue

            # Get аналитика (G = row[6])
            аналитика = row[6].lower() if len(row) > 6 else ''

            if any(k in desc for k in DOP_KEYS):
                daily_total_fakt_dop += amount
                if 'парковк' in desc or 'парк' in desc:
                    daily_parking_amount += amount
                    daily_parking_count += 1
                continue

            if 'выручка' in desc:
                if 'фото' in аналитика:
                    daily_total_fakt_photo += amount
                    continue
                if 'видео' in аналитика or 'мастер' in аналитика:
                    daily_total_fakt_video += amount
                    continue


        # School money for the day
        daily_school_money = int((daily_hours_totals.get('school_class', 0.0) + daily_hours_totals.get('school_homework', 0.0)) * 600)

        # Output for the day
        def pr(name, key):
            h = round(daily_hours_totals.get(key, 0.0), 2)
            c = daily_counts.get(key, 0)
            print(f"{name}: {h} ч (бронирований: {c})")

        print(f"Day {current_date_str}:")
        pr("Фотосъемка", "photo")
        pr("Банкет", "banquet")
        pr("Видео съемки/Мастер класс", "video_master")
        pr("Мероприятие", "event")
        pr("Корпоративные клиенты", "corporate")
        pr("фотошкола занятия", "school_class")
        pr("фотошкола домашние работы студентов", "school_homework")
        pr("Плавающая бронь", "floating")
        pr("Не приехали/не приедут", "no_show")
        pr("тех.бронь", "tech")
        pr("мероприятие бланк", "event_blank")
        pr("мероприятие Сися и White Studios", "event_sisia_white")
        pr("мероприятия yauza_place", "yauza_place")
        pr("мероприятия crystal", "crystal")
        pr("Неопределённые", "unknown")

        print(f"Предоплаты фото: {daily_total_prepayment_photo} руб.")
        print(f"Предоплаты видео: {daily_total_prepayment_video} руб.")
        print(f"По факту фото: {daily_total_fakt_photo} руб.")
        print(f"По факту видео: {daily_total_fakt_video} руб.")
        print(f"Доп. услуги: {daily_total_fakt_dop} руб.")

        print(f"Парковки сумма: {daily_parking_amount} руб.; кол-во: {daily_parking_count}")
        print(f"Школа по часам: {daily_school_money} руб.")

        col_num = day + 1  # B for 1, etc.
        col = column_to_letter(col_num)

        # Сохранение данных для пакетной записи
        # В этом месте itogi_sheet гарантированно не None
        daily_data[current_date_str] = {
            'itogi_sheet': itogi_sheet,
            'col': col,
            'daily_hours_totals': daily_hours_totals, 
            'prep_photo': daily_total_prepayment_photo,
            'fakt_photo': daily_total_fakt_photo,
            'prep_video': daily_total_prepayment_video,
            'fakt_video': daily_total_fakt_video,
            'school': daily_school_money,
            'dop': daily_total_fakt_dop
        }

        print("Обновления для дня:")
        print(f"{col}5: {daily_total_prepayment_photo}")
        print(f"{col}7: {daily_total_fakt_photo}")
        print(f"{col}6: {daily_total_prepayment_video}")
        print(f"{col}8: {daily_total_fakt_video}")
        print(f"{col}13: {daily_school_money}")
        print(f"{col}15: {daily_total_fakt_dop}")

        driver.refresh()
        time.sleep(2)  # Refresh and pause after day

    # Общие часы
    print("Общие часы за период:")
    for name, key in [
        ("Фотосъемка", "photo"),
        ("Банкет", "banquet"),
        ("Видео съемки/Мастер класс", "video_master"),
        ("Мероприятие", "event"),
        ("Корпоративные клиенты", "corporate"),
        ("фотошкола занятия", "school_class"),
        ("фотошкола домашние работы студентов", "school_homework"),
        ("Плавающая бронь", "floating"),
        ("Не приехали/не приедут", "no_show"),
        ("тех.бронь", "tech"),
        ("мероприятие бланк", "event_blank"),
        ("мероприятие Сися и White Studios", "event_sisia_white"),
        ("мероприятия yauza_place", "yauza_place"),
        ("мероприятия crystal", "crystal"),
        ("Неопределённые", "unknown")
    ]:
        h = round(total_hours_totals.get(key, 0.0), 2)
        print(f"{name}: {h} ч")

    confirm = input("Внести в таблицу? (yes/no): ").strip().lower()
    
    # --- ИСПРАВЛЕНИЕ ОШИБКИ 429: ИСПОЛЬЗОВАНИЕ BATCH_UPDATE ---
    if confirm == 'yes':
        for date_str, data in daily_data.items():
            if data['itogi_sheet'] is None:
                print(f"Пропущена запись для {date_str}, так как лист итогов не был инициализирован.")
                continue
            try:
                update_sheet_batch(data['itogi_sheet'], data['col'], data, data['daily_hours_totals'])
                print(f"Записано для дня {date_str}.")
                time.sleep(1) 
            except gspread.exceptions.APIError as e:
                if 'Quota exceeded' in str(e):
                    print(f"ОШИБКА 429 (Quota exceeded) при записи данных для {date_str}. Попробуйте повторить позже.")
                else:
                    print(f"API Ошибка при записи данных для {date_str}: {e}")
            except Exception as e:
                print(f"Неизвестная ошибка записи в таблицу для {date_str}: {e}")
    else:
        print("Отменено.")

finally:
    if 'driver' in locals() and driver:
        driver.quit()