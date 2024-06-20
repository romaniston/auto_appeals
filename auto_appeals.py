import funcs
import links
import time
import openpyxl
from colorama import init, Fore, Back, Style

#Инициализация фреймворка Colorama 
init()

#Для отключения дублирования записи "Авторизация выполнена"
first_auth = True

#Раскрашивание текста и фона при успешном выполнении и при ошибке
success_step = Back.LIGHTGREEN_EX + Fore.BLACK
fail_step = Back.RED + Fore.BLACK
reset_color = Style.RESET_ALL

#Игнор ошибки Excel, ввод логина и пароля, выбор режима работы браузера
funcs.excel_error_ignore()
login, password = funcs.input_login()
headless_or_not = input(reset_color + f'------------'
                        f'\nВыберите режим работы браузера:'
                        f'\n1 - В фоне.'
                        f'\n2 - В окне.'
                        f'\n>>> ')

#Цикл программы
infinite_while = 1
while infinite_while == 1:
    appeal_num_start, appeal_num_end = funcs.appeals_range()
    complete_easycheck = funcs.complete_easycheck_func()

    #Запись режима работы браузера
    if headless_or_not == '1':
        driver = funcs.run_headless_browser()
    else:
        driver = funcs.run_not_headless_browser()

    print(reset_color + '------------')
    print(reset_color + 'Идет авторизация в веб-сервисе')

    #Открытие страницы в браузере
    driver.get(links.links['main_page'])

    #Авторизация в веб-сервисе
    funcs.step_by_step().webdriverwait_func(driver,
                                            links.links['login_field'],
                                            'Ошибка: не найден элемент "Поле для ввода логина"',
                                            60)
    funcs.step_by_step().fill_the_field(driver,
                                        links.links['login_field'],
                                        login,
                                        'Ошибка: не удалость заполнить поле для ввода логина',
                                        0.2)
    funcs.step_by_step().fill_the_field(driver,
                                        links.links['pass_field'],
                                        password,
                                        'Ошибка: не удалость заполнить поле для ввода пароля',
                                        0.2)
    funcs.step_by_step().click_at_the_element(driver,
                                              links.links['pass_field'],
                                              'Ошибка: не удалось активировать кнопку авторизации',
                                              0.2)

    while appeal_num_start <= appeal_num_end:

        data = funcs.write_appeals_data_to_dict(appeal_num_start)

        #Добавление обращения
        funcs.step_by_step().webdriverwait_func(driver,
                                                links.links['add_appeal_button'],
                                                'Ошибка: не найдена кнопка "Добавить"',
                                                60)

        #Вывод сообщения при первом выполнении программы, и отключение вывода при последующих
        if first_auth:
            print(reset_color + "Авторизация выполнена.")
            first_auth = False
        else:
            pass

        funcs.step_by_step().click_at_the_element(driver,
                                                  links.links['add_appeal_button'],
                                                  'Ошибка: не удалось нажать на кнопку "Добавить"',
                                                  0.2)

        print(reset_color + "------------")
        print(success_step + f"Выполняется создание обращения №{appeal_num_start}")

        #Выбор категории
        funcs.step_by_step().webdriverwait_func(driver,
                                                links.links['category_field'],
                                                'Ошибка: не найдено поле "Категория"',
                                                60)
        funcs.step_by_step().fill_the_field(driver,
                                            links.links['category_field'],
                                            data['category'],
                                            'Ошибка: не удалость заполнить поле "Категория"',
                                            0.2)
        funcs.step_by_step().click_at_the_element_at_context(driver,
                                                             links.links['category_field'],
                                                             'Ошибка: не удалость заполнить поле "Категория"',
                                                             0.2,
                                                             3)

        #Выбор контакта
        funcs.step_by_step().fill_the_field(driver,
                                            links.links['contact_field'],
                                            data['contact'],
                                            'Ошибка: не удалость заполнить поле "Контакт"',
                                            0.2)
        funcs.step_by_step().click_at_the_element_at_context(driver,
                                                             links.links['contact_field'],
                                                             'Ошибка: не удалость заполнить поле "Контакт"',
                                                             0.2,
                                                             3)

        #Выбор сервиса
        funcs.step_by_step().fill_the_field(driver,
                                            links.links['servis_field'],
                                            data['servis'],
                                            'Ошибка: не удалость заполнить поле "Сервис"',
                                            0.2)
        funcs.step_by_step().click_at_the_element_at_context(driver,
                                                             links.links['servis_field'],
                                                             'Ошибка: не удалость заполнить поле "Сервис"',
                                                             0.2,
                                                             3)

        #Выбор работы
        funcs.step_by_step().fill_the_field(driver,
                                            links.links['work_field'],
                                            data['work'],
                                            'Ошибка: не удалость заполнить поле "Работа"',
                                            0.2)
        funcs.step_by_step().click_at_the_element_at_context(driver,
                                                             links.links['work_field'],
                                                             'Ошибка: не удалость заполнить поле "Работа"',
                                                             0.2,
                                                             3)

        #Заполнение проекта
        if data['project'] == None:
            pass
        else:
            funcs.step_by_step().fill_the_field(driver,
                                                links.links['project_field'],
                                                data['project'],
                                                'Ошибка: не удалость заполнить поле "Проект"',
                                                0.2)
            funcs.step_by_step().click_at_the_element_at_context(driver,
                                                                 links.links['project_field'],
                                                                 'Ошибка: не удалость заполнить поле "Проект"',
                                                                 0.2,
                                                                 3)

        #Заполнение темы
        funcs.step_by_step().fill_the_field(driver,
                                            links.links['theme_field'],
                                            data['theme'],
                                            'Ошибка: не удалость заполнить поле "Тема"',
                                            0.1)

        #Заполнение описания
        funcs.step_by_step().fill_the_field(driver,
                                            links.links['description_field'],
                                            data['description'],
                                            'Ошибка: не удалость заполнить поле "Описание"',
                                            0.1)

        #Заполнение поля результата
        funcs.step_by_step().fill_the_field(driver,
                                            links.links['result_field'],
                                            data['result'],
                                            'Ошибка: не удалость заполнить поле "Решение"',
                                            0.1)

        #Заполнение поля трудозатрат
        funcs.step_by_step().fill_the_field(driver,
                                            links.links['worktime_field'],
                                            data['worktime'],
                                            'Ошибка: не удалость заполнить поле "Трудозатраты"',
                                            0.1)

        #Дежурство
        if data['duty'] == None:
            pass
        else:
            funcs.step_by_step().checkbox_click(driver,
                                                links.links['duty_checkbox'],
                                                'Ошибка: не удалось поставить галочку в чекбоксе "Дежурство"',
                                                0.1)

        #Заполнение поля отдела
        funcs.step_by_step().fill_the_field(driver,
                                            links.links['department_field'],
                                            data['department'],
                                            'Ошибка: не удалость заполнить поле "Отдел"',
                                            0.2)
        funcs.step_by_step().click_at_the_element_at_context(driver,
                                                             links.links['department_field'],
                                                             'Ошибка: не удалость заполнить поле "Отдел"',
                                                             0.2,
                                                             3)

        #Сохранение активности
        funcs.step_by_step().click_at_the_element(driver,
                                                  links.links['save_activity'],
                                                  'Ошибка: не удалость нажать на кнопку сохранения активности',
                                                  0.2)

        #Запись в AppealsCompleted.xlsx об успешном заведении обращения
        book = openpyxl.open("AppealsCompleted.xlsx")
        sheet = book.active
        sheet[1 + int(appeal_num_start)][0].value = str('Заведено')
        book.save('AppealsCompleted.xlsx')
        book.close()

        #Finished
        print(success_step + 'Обращение №' + str(appeal_num_start) + ' успешно заведено! AppealsCompleted.xlsx обновлен')
        appeal_num_start += 1

    print(reset_color + '------------')
    print(success_step + "Все обращения заведены!")

    #Выполнение EasyCheck
    if complete_easycheck == '1':
        print(reset_color + '------------')
        print(success_step + "Идёт выполнение somefunc!")
        funcs.step_by_step().webdriverwait_func(driver,
                                                links.links['somefunc_button'],
                                                'Ошибка: не найдена кнопка "somefunc"',
                                                60)
        funcs.step_by_step().click_at_the_element(driver,
                                                  links.links['somefunc_button'],
                                                  'Ошибка: не удалость нажать на кнопку "somefunc"',
                                                  1)
        time.sleep(20)
        print(success_step + "somefunc выполнен")
        driver.quit()
    else:
        driver.quit()