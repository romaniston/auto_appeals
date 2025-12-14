# auto_appeals
<p>A tool to automate filling out certain forms on a specific web service using Selenium (depersonalized).
<hr>
<p>Скрипт представляет собой инструмент автоматизации заполнения определенных форм на определенном веб-сервисе
для личного использования. Данная версия проекта является обезличеной.
<hr>
<p>• done_delete.py - скрипт для очищения статусов задач в DoneDelete.xlsx.
<p>• Appeals.xlsx - таблица для заготовок задач для последующего заведения в формы.
<p>• DoneDelete.xlsx - таблица статусов обращений.
<hr>
<p>Инструкция:
<p>1) Готовим типовые шаблоны в таблице "Appeals.xlsx", заносим нужную информацию в контекстные списки на странице "Lists".
<p>2) Заполняем страницу "Tasks" таблицы "Appeals.xlsx", используя подготовленные типовые шаблоны.
<p>3) Запускаем "auto_appeals.py", устанавливаем параметры заполнения форм, выполняем.
<p>4) В таблице "Appeals.xlsx" перед каждой задачей устанавливается статус заполнения "Заведено". Чтобы очистить статусы запускаем "DoneDelete.py"
<hr>
Screenshots:
<p>
<p><img src="screenshots/1.jpg" width=500></img>
<p><img src="screenshots/2.png" width=1920></img>
<p><img src="screenshots/3.png" width=1920></img>
