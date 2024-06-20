from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl
import warnings
from colorama import init, Fore, Back, Style

#Инициализация фреймворка Colorama
init()

#Раскрашивание текста и фона при успешном выполнении и при ошибке
success_step = Back.LIGHTGREEN_EX + Fore.BLACK
fail_step = Back.RED + Fore.BLACK
reset_color = Style.RESET_ALL

#Игнорирование ошибки о Data validation extension в Excel
def excel_error_ignore():
    excel_error_ignore = warnings.simplefilter(action='ignore', category=UserWarning)
    return excel_error_ignore

#Данные входа в веб-сервис
def input_login():
    login = input(reset_color + 'Введите ваш логин от веб-сервиса >>> ')
    password = input(reset_color + 'Введите пароль >>> ')
    return login, password

#Задание диапазона обращений
def appeals_range():
    print(reset_color + '------------')
    appeal_num_start = int(input(reset_color + 'Введите номер обращения с которого хотите начать заведение >>> '))
    appeal_num_end = int(input(reset_color + 'Введите номер обращения на котором хотите закончить (включительно) >>> '))

    if appeal_num_end < appeal_num_start:
        while appeal_num_end < appeal_num_start:
            appeal_num_end = int(input(reset_color + 'Номер конечного обращения должен быть больше начального, введите еще раз >>> '))
    else:
        pass
    return appeal_num_start, appeal_num_end

#Открытие Appeals.xlsx и запись данных выбранного обращения и активности в словарь
def write_appeals_data_to_dict(appeal_start):
    book = openpyxl.open("Appeals.xlsx", data_only=True, read_only=True)
    sheet = book.active

    data = {}
    datanames = [f''
        'category',
        'contact',
        'servis',
        'work',
        'project',
        'theme',
        'description',
        'result',
        'worktime',
        'duty',
        'department'
    ]

    for var in range(2,13):
        tempvar = sheet[1 + int(appeal_start)][var].value
        data[datanames[var - 2]] = tempvar

    book.close()

    return data

# Запуск хрома в фоне
def run_headless_browser():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('--log-level=3')
    driver = webdriver.Chrome(options=options)
    return driver

# Запуск хрома не в фоне
def run_not_headless_browser():
    driver = webdriver.Chrome()
    return driver

#Выполнить ли somefunc в конце заведения обращений
def complete_easycheck_func():
    print('------------')
    complete_easycheck = input(reset_color + f'Выполнить somefunc в конце заведения обращений?'
                               f'\n1 - Да.'
                               f'\n2 - Нет.'
                               f'\n>>> ')
    return complete_easycheck

#Класс с методами для взаимодействия с веб элементами
class step_by_step():

    #Ожидание элемента
    def webdriverwait_func(self, driver, selector, get_except, web_driver_wait_sec):

        try:
            element = WebDriverWait(driver, web_driver_wait_sec)
            element.until(EC.presence_of_element_located((By.XPATH, selector)))
        except:
            print(fail_step + get_except)
            print(fail_step + 'Программа остановлена!')
            time.sleep(999999999)
        finally:
            pass

    #Заполнить поле
    def fill_the_field(self, driver, selector, for_fill, get_except, time_sleep_sec):
        try:
            time.sleep(time_sleep_sec)
            focus_at_element = driver.find_element(By.XPATH, selector)
            time.sleep(time_sleep_sec)
            focus_at_element.clear()
            time.sleep(time_sleep_sec)
            focus_at_element.send_keys(for_fill)
        except:
            print(fail_step + get_except)
            print(fail_step + 'Программа остановлена!')
            time.sleep(999999999)
        finally:
            pass

    #Клик по элементу
    def click_at_the_element(self, driver, selector, get_except, time_sleep_sec):
        try:
            time.sleep(time_sleep_sec)
            focus_at_element = driver.find_element(By.XPATH, selector)
            time.sleep(time_sleep_sec)
            focus_at_element.send_keys(Keys.RETURN)
        except:
            print(fail_step + get_except)
            print(fail_step + 'Программа остановлена!')
            time.sleep(999999999)
        finally:
            pass

    #Клик по первому элементу из всплывающего окна (при заполнении поля)
    def click_at_the_element_at_context(self, driver, selector, get_except, time_sleep_sec, time_sleep_sec_till_appears_context_menu):
        try:
            time.sleep(time_sleep_sec)
            focus_at_element = driver.find_element(By.XPATH, selector)
            time.sleep(time_sleep_sec_till_appears_context_menu)
            focus_at_element.send_keys(Keys.ARROW_DOWN)
            time.sleep(time_sleep_sec)
            focus_at_element.send_keys(Keys.RETURN)
        except:
            print(fail_step + get_except)
            print(fail_step + 'Программа остановлена!')
            time.sleep(999999999)
        finally:
            pass

    #Клик по чекбоксу
    def checkbox_click(self, driver, selector, get_except, time_sleep_sec):
        try:
            time.sleep(time_sleep_sec)
            focus_at_element = driver.find_element(By.XPATH, selector)
            if not focus_at_element.is_selected():
                focus_at_element.click()
            else:
                pass
        except:
            print(fail_step + get_except)
            print(fail_step + 'Программа остановлена!')
            time.sleep(999999999)
        finally:
            pass

