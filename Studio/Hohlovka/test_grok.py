from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime

# Настройки Chrome
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Запуск браузера
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    # Вход в резерватор
    driver.get("https://whitestudios.ru/reservator/login/")
    login_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "login")))
    password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "pass")))
    login_field.send_keys("Kapr")  # Твой логин
    password_field.send_keys("vut79k")  # Твой пароль
    driver.find_element(By.CSS_SELECTOR, 'input[type="submit"]').click()
    time.sleep(5)

    # Получаем текущую дату
    current_date = datetime.now().strftime('%d.%m.%Y')  # Сегодня: '22.09.2025'
    print(f"Текущая дата: {current_date}")

    # Вызываем функцию showStudio для "Яуза" на текущий день
    driver.execute_script(f"showStudio(32, '{current_date}');")
    time.sleep(5)  # Даём время на загрузку

    # Ожидание загрузки данных в studioBody32
    studio_body = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "studioBody32"))
    )
    WebDriverWait(driver, 20).until(
        lambda x: len(studio_body.text.strip()) > 0
    )
    print(f"Данные для студии Яуза на сегодня: {studio_body.text[:200]}")

    # Собираем все уникальные href бронирований для фотосъёмок (красный фон)
    photo_bookings = driver.find_elements(By.CSS_SELECTOR, '#studioBody32 .reserved[style*="background: #FF0000"]')
    print(f"Найдено бронирований: {len(photo_bookings)}")

    hrefs = set()  # Используем set для уникальности
    for booking in photo_bookings:
        try:
            booking_link = booking.find_element(By.TAG_NAME, "a")
            href = booking_link.get_attribute("href")
            if href:
                hrefs.add(href)
        except Exception as e:
            print(f"Ошибка при извлечении ссылки: {str(e)}")
            continue

    print(f"Найдено уникальных ссылок на брони: {len(hrefs)}")

    # Общая сумма предоплат
    total_prepayment = 0

    for href in hrefs:
        try:
            # Переходим напрямую по ссылке на поп-ап брони
            driver.get(href)
            time.sleep(3)  # Ждём загрузки поп-апа

            # Ожидание изменения URL (если нужно, но поскольку get() напрямую, проверяем наличие контента)
            WebDriverWait(driver, 15).until(
                lambda d: "booking/view" in d.current_url or len(d.find_element(By.TAG_NAME, "body").text.strip()) > 0
            )
            print(f"Перешли на URL брони: {driver.current_url}")

            # Получаем полный текст поп-апа
            full_popup_text = driver.find_element(By.TAG_NAME, "body").text
            print(f"Полный текст поп-апа брони: {full_popup_text[:200]}...")  # Обрезаем для лога

            # Поиск общего времени (опционально, для отладки)
            time_line = next((line for line in full_popup_text.splitlines() if "c " in line and " до " in line), None)
            if time_line:
                print(f"Общее время брони: {time_line}")

            # Поиск предоплаты: ищем строку с "Оплата с сайта" и извлекаем число
            prepayment_line = next((line for line in full_popup_text.splitlines() if "Оплата с сайта" in line), None)
            if prepayment_line:
                # Ищем все цифры и берём первую подходящую (предполагаем, что это сумма)
                digits = [int(s) for s in prepayment_line.split() if s.isdigit()]
                prepayment = digits[0] if digits else 0
                total_prepayment += prepayment
                print(f"Найдена предоплата для брони: {prepayment} руб.")
            else:
                print("Предоплата для брони не найдена.")

        except Exception as e:
            print(f"Ошибка при обработке брони по {href}: {str(e)}")
            continue

    print(f"Общая сумма предоплат за фотосъёмки на сегодня: {total_prepayment} руб.")

except Exception as e:
    print(f"Произошла ошибка: {str(e)}")

finally:
    driver.quit()