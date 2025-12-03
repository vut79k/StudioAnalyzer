import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('google_key.json', scope)
client = gspread.authorize(creds)

# Словарь для русских имен месяцев
months_ru = {
    'January': 'Январь',
    'February': 'Февраль',
    'March': 'Март',
    'April': 'Апрель',
    'May': 'Май',
    'June': 'Июнь',
    'July': 'Июль',
    'August': 'Август',
    'September': 'Сентябрь',
    'October': 'Октябрь',
    'November': 'Ноябрь',
    'December': 'Декабрь'
}

# Получаем текущий месяц на английском и переводим на русский
current_month_en = datetime.now().strftime('%B')
current_month_ru = months_ru[current_month_en]
year_short = datetime.now().strftime('%y')  # '25'
sheet_name = f'{current_month_ru}{year_short}'  # 'Сентябрь25' для сентября 2025

sheet_itogi = client.open_by_key('1lgn068NObnej5A0J0Ya4Kxz5JY6tptxKf9dZTf2VXrI').worksheet(sheet_name)
sheet_kassa = client.open_by_key('1biAzb8vVeaTsClkozuViWdxqQ10Bo5NI93SLtvZy1bk').sheet1

print(f"Доступ ок! Лист: {sheet_name}, первая строка:", sheet_itogi.row_values(1))
print("Первая строка Кассы:", sheet_kassa.row_values(1))