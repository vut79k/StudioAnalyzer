from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from config import RESERVATOR_URL, USERNAME, PASSWORD
import time

# Указываем путь к Chrome
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Запускаем автоматический браузер
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    # Открываем страницу логина
    driver.get(RESERVATOR_URL)
    print(f"Открыта страница: {driver.current_url}")  # Диагностика URL

    # Ждём и заполняем поля
    login_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "login")))
    password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "pass")))

    login_field.send_keys(USERNAME)
    password_field.send_keys(PASSWORD)

    # Нажимаем кнопку
    submit_button = driver.find_element(By.CSS_SELECTOR, 'input[type="submit"]')
    submit_button.click()
    time.sleep(5)  # Ждём после логина

    # Проверяем, зашли ли
    if 'dashboard' in driver.current_url or 'reservations' in driver.current_url or '/login' not in driver.current_url:
        print("Логин успешно выполнен! Текст на странице:", driver.page_source[:200])
    else:
        print("Ошибка логина. Проверь логин/пароль. Текущий URL:", driver.current_url)

except Exception as e:
    print(f"Произошла ошибка: {str(e)}")

finally:
    driver.quit()