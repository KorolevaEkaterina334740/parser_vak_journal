import random
import time
import pandas as pd
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fake_useragent import UserAgent

def selenium_move(driver):
    pass
    # actions = ActionChains(driver)
    # actions.move_by_offset(random.randint(10, 100), random.randint(10, 100)).perform()
    # driver.execute_script("window.scrollBy(0, document.body.scrollHeight * 0.5);")

# Конфигурация авторизации
ELIBRARY_LOGIN = "LOGIN"  # <- ВВЕДИТЕ ВАШ ЛОГИН
ELIBRARY_PASSWORD = "PASSWORD"  # <- ВВЕДИТЕ ВАШ ПАРОЛЬ


def elibrary_login(driver):
    """Функция для авторизации на eLibrary.ru"""
    try:
        # Переходим на страницу авторизации
        driver.get("https://elibrary.ru/defaultx.asp")

        # Ожидаем появления формы логина
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "login")))

        selenium_move(driver)
        # Заполняем логин
        login_field = driver.find_element(By.ID, "login")
        login_field.clear()
        login_field.send_keys(ELIBRARY_LOGIN)

        selenium_move(driver)

        # Заполняем пароль
        password_field = driver.find_element(By.ID, "password")
        password_field.clear()
        password_field.send_keys(ELIBRARY_PASSWORD)

        # Нажимаем кнопку входа
        login_button = driver.find_element(By.XPATH, "//div[@class='butred' and contains(text(), 'Вход')]")
        login_button.click()

        print("Авторизация прошла успешно!")
        time.sleep(random.uniform(1, 3))

    except Exception as e:
        print(f"Ошибка авторизации: {str(e)}")
        driver.quit()
        exit()

def get_new_driver():
    user_agent = UserAgent().random
    options = webdriver.ChromeOptions()
    options.add_argument(f"user-agent={user_agent}")
    return webdriver.Chrome(options=options)


def main():
    # Создаем список для хранения данных
    data = []

    try:
        driver = get_new_driver()

        # Переход на страницу eLibrary.ru
        driver.get("https://www.elibrary.ru/titles.asp")
        time.sleep(10)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "vak")))

        # Раскрытие выпадающего списка "тематика журнала"
        theme_dropdown = driver.find_element(By.NAME, "rubriccode")
        theme_dropdown.click()

        vak_select = Select(driver.find_element(By.NAME, "rubriccode"))
        vak_select.select_by_index(63)

        # Работа с перечнем ВАК
        vak_dropdown = driver.find_element(By.NAME, "vak")
        vak_dropdown.click()

        vak_select = Select(driver.find_element(By.NAME, "vak"))
        vak_select.select_by_index(3)

        search_button = driver.find_element(By.CSS_SELECTOR, "div.butred[onclick='title_search()']")
        search_button.click()

        journals = {}
        link = "https://elibrary.ru/title_profile.asp?id="
        counter = 0

        for step in range(0, 2):
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "restab"))
            )

        results_table = driver.find_element(By.ID, "restab")
        rows = results_table.find_elements(By.TAG_NAME, "tr")

        for row in rows[3:]:
            cells = row.find_elements(By.TAG_NAME, "td")
            counter += 1
            journals[counter] = {
                "link": link + row.get_attribute("id")[1:],
                "title": cells[2].text.split('\n')[0],
                "author": cells[2].text.split('\n')[1],
                "publications": cells[3].text,
                "article": cells[4].text,
                "quotes": cells[5].text,
            }

        if step < 1:
            next_page = driver.find_element(By.XPATH, "//a[@title='Следующая страница']")
            next_page.click()
            time.sleep(random.randint(2, 4))

        # Сбор данных
        for number, journal in journals.items():
            try:
                driver.get(journal["link"])
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "footer")))
                table = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, "(//table[@width='580' and @cellspacing='0' and @cellpadding='3'])[2]")))

                rows = table.find_elements(By.TAG_NAME, "tr")

                # Извлекаем данные
                count_of_articles = rows[3].find_elements(By.TAG_NAME, "td")[-1].text
                science_index = rows[5].find_elements(By.TAG_NAME, "td")[-1].text
                index_hirsha = rows[46].find_elements(By.TAG_NAME, "td")[-1].text
                index_herfindal = rows[49].find_elements(By.TAG_NAME, "td")[-1].text
                index_jinny = rows[53].find_elements(By.TAG_NAME, "td")[-1].text
                views_per_year = rows[59].find_elements(By.TAG_NAME, "td")[-1].text

                # Рассчет показателя
                try:
                    views_per_article = int(views_per_year) / int(count_of_articles)
                except (ValueError, ZeroDivisionError):
                    views_per_article = 0

                # Добавляем данные в список
                data.append({
                    '№': number,
                    'link': journal["link"],
                    'title': journal["title"],
                    'author': journal["author"],
                    'publications': journal["publications"],
                    'article': journal["article"],
                    'quotes': journal["quotes"],
                    'science_index': science_index,
                    'index_hirsha': index_hirsha,
                    'index_herfindal': index_herfindal,
                    'index_jinny': index_jinny,
                    'views_per_year': views_per_year,
                    'count_of_articles': count_of_articles,
                    'views_per_article': views_per_article
                })
            except TimeoutException as e:
                continue
            except Exception as e:
                print(f"Ошибка при парсинге журнала {number}: {str(e)}")

            time.sleep(random.randint(5, 6))

    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        # Сохраняем результат в Excel
        if data:
            df = pd.DataFrame(data)
            try:
                with pd.ExcelWriter('journals_data.xlsx', engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Журналы')

                    # Получаем объекты для форматирования
                    workbook = writer.book
                    worksheet = writer.sheets['Журналы']

                    # Форматируем заголовки
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'top',
                        'fg_color': '#D7E4BC',
                        'border': 1
                    })

                    # Применяем форматирование к заголовкам
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                    # Автонастройка ширины столбцов
                    for i, col in enumerate(df.columns):
                        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                        worksheet.set_column(i, i, max_len)

                print("Данные успешно сохранены в файл journals_data.xlsx")
            except Exception as e:
                print(f"Ошибка при сохранении файла: {str(e)}")
        else:
            print("Нет данных для сохранения")

        if 'driver' in locals():
            driver.quit()


if __name__ == "__main__":
    main()
