"""
localhost.py - Веб-сервер Flask для калькулятора цен OZON FBO
==============================================================

НАЗНАЧЕНИЕ:
- Предоставляет REST API для управления базами данных
- Обслуживает веб-интерфейс для запуска расчётов
- Обрабатывает загрузку файлов и скачивание результатов

ИСПОЛЬЗУЕМЫЕ БИБЛИОТЕКИ:
- Flask: микрофреймворк для создания веб-приложений
- requests: для HTTP запросов
- pandas: для работы с Excel и DataFrame
- sqlite3: для работы с базами данных
- werkzeug: для безопасной обработки имён файлов
"""

# ============== ИМПОРТЫ БИБЛИОТЕК ==============

from flask import Flask, request, jsonify, send_file, render_template, send_from_directory
# Flask - основной фреймворк для веб-приложения
# request - объект для получения данных из HTTP запроса
# jsonify - функция для преобразования dict в JSON
# send_file - отправка файлов клиенту
# render_template - рендеринг HTML шаблонов
# send_from_directory - отправка файлов из директории

import sqlite3
# sqlite3 - встроенная библиотека для работы с SQLite базами данных

import os
# os - работа с файловой системой и переменными окружения

import pandas as pd
# pandas - работа с Excel, CSV и табличными данными

import sys
# sys - системные переменные и параметры

from werkzeug.utils import secure_filename

# secure_filename - функция для безопасного преобразования имён файлов
# (удаляет опасные символы, предотвращает атаки)

# ============== ИНИЦИАЛИЗАЦИЯ ПРИЛОЖЕНИЯ ==============

# Добавляем путь к папке с calculator.py
sys.path.append(os.path.dirname(__file__))

# Создаём Flask приложение
# template_folder='templates' - папка с HTML шаблонами
# static_folder='static' - папка со статическими файлами (CSS, JS, изображения)
app = Flask(__name__, template_folder='templates', static_folder='static')

# ============== КОНФИГУРАЦИЯ ==============

# Константы - папки для хранения файлов
DB_FOLDER = "databases"  # Папка для БД (поставляются пользователем)
UPLOAD_FOLDER = "uploads"  # Папка для загруженных файлов Excel
RESULT_FOLDER = "results"  # Папка для результатов расчётов

# Создаём папки если они не существуют
for folder in [DB_FOLDER, UPLOAD_FOLDER, RESULT_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# Конфигурация Flask приложения
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER  # Папка для загрузок
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # Максимальный размер файла: 50MB


# ============== REST API ЭНДПОИНТЫ ==============

@app.route('/')
def index():
    """
    ЭНДПОИНТ: GET /
    НАЗНАЧЕНИЕ: Главная страница веб-интерфейса
    ВОЗВРАЩАЕТ: HTML страница калькулятора
    """
    return render_template('index.html')


@app.route('/api/databases', methods=['GET'])
def get_databases():
    """
    ЭНДПОИНТ: GET /api/databases
    НАЗНАЧЕНИЕ: Получить список всех загруженных баз данных

    ЛОГИКА:
    1. Сканирует папку databases
    2. Отбирает файлы с расширением .db
    3. Сортирует по убыванию (новые первыми)
    4. Возвращает JSON

    ВОЗВРАЩАЕТ: {"success": true, "databases": ["db1.db", "db2.db"]}
    """
    try:
        # Получаем список файлов в папке и фильтруем только .db файлы
        databases = [f for f in os.listdir(DB_FOLDER) if f.endswith('.db')]

        # Сортируем: от новых к старым (для быстрого поиска последней БД)
        databases.sort(reverse=True)

        return jsonify({"success": True, "databases": databases})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/upload_database', methods=['POST'])
def upload_database():
    """
    ЭНДПОИНТ: POST /api/upload_database
    НАЗНАЧЕНИЕ: Загрузить новую базу данных на сервер

    ПАРАМЕТРЫ (multipart/form-data):
    - file: SQLite БД файл (.db)

    ЛОГИКА:
    1. Проверяет наличие файла в запросе
    2. Проверяет расширение .db
    3. Сохраняет файл в папку databases
    4. Возвращает имя загруженного файла

    ВОЗВРАЩАЕТ: {"success": true, "database_name": "ozon_data.db"}
    """
    try:
        # Проверка 1: Файл был ли загружен
        if 'file' not in request.files:
            return jsonify({"success": False, "message": "Файл не найден"}), 400

        file = request.files['file']

        # Проверка 2: Имя файла не пустое
        if file.filename == '':
            return jsonify({"success": False, "message": "Имя файла пустое"}), 400

        # Проверка 3: Расширение файла .db
        if not file.filename.endswith('.db'):
            return jsonify({"success": False, "message": "Только .db файлы"}), 400

        # Путь для сохранения файла
        filepath = os.path.join(DB_FOLDER, file.filename)

        # Сохраняем файл
        file.save(filepath)

        # Возвращаем успешный ответ
        return jsonify({
            "success": True,
            "message": "БД загружена",
            "database_name": file.filename
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route('/api/delete_database/<db_name>', methods=['DELETE'])
def delete_database(db_name):
    """
    ЭНДПОИНТ: DELETE /api/delete_database/<db_name>
    НАЗНАЧЕНИЕ: Удалить базу данных с сервера

    ПАРАМЕТРЫ:
    - db_name: Имя файла БД (например, "ozon_data.db")

    ЛОГИКА:
    1. Проверяет существование файла
    2. Удаляет файл
    3. Возвращает подтверждение

    ВОЗВРАЩАЕТ: {"success": true, "message": "БД удалена"}
    """
    try:
        # Полный путь к файлу БД
        filepath = os.path.join(DB_FOLDER, db_name)

        # Проверка: Файл существует?
        if not os.path.exists(filepath):
            return jsonify({"success": False, "message": "БД не найдена"}), 404

        # Удаляем файл
        os.remove(filepath)

        return jsonify({
            "success": True,
            "message": f"БД {db_name} удалена"
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route('/api/get_data/<db_name>/<table_name>', methods=['GET'])
def get_data(db_name, table_name):
    """
    ЭНДПОИНТ: GET /api/get_data/<db_name>/<table_name>
    НАЗНАЧЕНИЕ: Получить данные из таблицы БД

    ПАРАМЕТРЫ:
    - db_name: Имя БД файла
    - table_name: Имя таблицы (catalog, logistics, new_tariffs)

    ЛОГИКА:
    1. Подключается к БД
    2. Проверяет существование таблицы
    3. Получает все строки
    4. Преобразует в JSON формат (список словарей)

    ВОЗВРАЩАЕТ: {"success": true, "data": [...]}
    """
    try:
        # Путь к БД файлу
        db_path = os.path.join(DB_FOLDER, db_name)

        # Проверка: БД существует?
        if not os.path.exists(db_path):
            return jsonify({"success": False, "error": "БД не найдена"}), 404

        # Подключаемся к SQLite БД
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Проверка: Таблица существует в БД?
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
        if not cursor.fetchone():
            conn.close()
            return jsonify({"success": False, "error": f"Таблица {table_name} не найдена"}), 404

        # Получаем ВСЕ данные из таблицы
        cursor.execute(f"SELECT * FROM {table_name}")
        rows = cursor.fetchall()

        # Получаем названия колонок из описания курсора
        # cursor.description содержит информацию о колонках
        column_names = [description[0] for description in cursor.description]

        conn.close()

        # Преобразуем строки в список словарей (JSON формат)
        # каждая строка становится dict: {"col1": val1, "col2": val2, ...}
        data = []
        for row in rows:
            data.append(dict(zip(column_names, row)))

        return jsonify({"success": True, "data": data})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/calculate', methods=['POST'])
def calculate():
    """
    ЭНДПОИНТ: POST /api/calculate
    НАЗНАЧЕНИЕ: Запустить расчёт цен для товаров

    ПАРАМЕТРЫ (multipart/form-data):
    - file: Excel файл с товарами (.xlsx)
    - target_margin: Целевая рентабельность (%)

    ЛОГИКА:
    1. Получает загруженный Excel файл
    2. Сохраняет его во временную папку
    3. Запускает функцию calculate_file() из calculator.py
    4. Возвращает имя файла с результатами

    ВОЗВРАЩАЕТ: {"success": true, "result_file": "result_products.xlsx"}
    """
    try:
        # Проверка 1: Файл был загружен?
        if 'file' not in request.files:
            return jsonify({"success": False, "message": "Файл не найден"}), 400

        # Получаем целевую рентабильность из формы (по умолчанию 20%)
        target_margin = float(request.form.get('target_margin', 20))

        # Получаем файл из запроса
        file = request.files['file']
        original_filename = file.filename

        # Проверка 2: Имя файла не пустое
        if file.filename == '':
            return jsonify({"success": False, "message": "Имя файла пустое"}), 400

        # Проверка 3: Расширение .xlsx или .xls
        if not (original_filename.endswith('.xlsx') or original_filename.endswith('.xls')):
            return jsonify({"success": False, "message": "Только .xlsx или .xls файлы"}), 400

        # БЕЗОПАСНОЕ СОХРАНЕНИЕ ИМЕНИ ФАЙЛА:
        # Разделяем имя и расширение
        file_ext = os.path.splitext(original_filename)[1]  # .xlsx
        file_base = os.path.splitext(original_filename)[0]  # имя без расширения

        # secure_filename удаляет опасные символы (/, \, ..., и т.д.)
        safe_base = secure_filename(file_base)

        # Объединяем обратно: безопасное имя + расширение
        filename = safe_base + file_ext

        # Полный путь к файлу для сохранения
        filepath = os.path.join(UPLOAD_FOLDER, filename)

        # Сохраняем файл
        file.save(filepath)

        # ОТЛАДКА: выводим информацию в консоль сервера
        print("\n" + "=" * 70)
        print("ЗАГРУЗКА ФАЙЛА")
        print("=" * 70)
        print(f"Оригинальное имя: {original_filename}")
        print(f"Сохранено как: {filename}")
        print(f"Полный путь: {os.path.abspath(filepath)}")
        print(f"Размер: {os.path.getsize(filepath) / 1024:.2f} KB")
        print(f"Файл существует: {os.path.exists(filepath)}")

        # Проверка: Файл действительно сохранился?
        if not os.path.exists(filepath):
            print("ОШИБКА: Ошибка сохранения файла")
            return jsonify({"success": False, "message": "Ошибка сохранения файла"}), 500

        # ============== ЗАПУСК РАСЧЁТА ==============

        # Импортируем функцию расчёта из calculator.py
        from calculator import calculate_file

        # Формируем имя для файла результата
        result_filename = f"result_{filename}"
        result_path = os.path.join(RESULT_FOLDER, result_filename)

        # Выводим информацию о расчёте
        print(f"\nЗАПУСК РАСЧЁТА")
        print(f"Целевая рентабильность: {target_margin}%")
        print(f"Входной файл: {filepath}")
        print(f"Выходной файл: {result_path}")
        print("=" * 70)

        # Запускаем функцию calculate_file из calculator.py
        # debug_first=False - без отладки для веб-версии (экономим время)
        calculate_file(filepath, target_margin, result_path, debug_first=False)

        # Проверка: Результат был создан?
        if not os.path.exists(result_path):
            print("ОШИБКА: Результат не создан")
            return jsonify({"success": False, "message": "Результат не создан"}), 500

        # Выводим информацию об успехе
        print(f"\nРАСЧЕТ ЗАВЕРШЁН")
        print(f"Размер результата: {os.path.getsize(result_path) / 1024:.2f} KB")
        print("=" * 70 + "\n")

        # Возвращаем успешный ответ с именем файла результата
        return jsonify({
            "success": True,
            "message": "Расчёт завершён",
            "result_file": result_filename
        })

    except Exception as e:
        # Если произошла ошибка, выводим её в консоль
        import traceback
        error_trace = traceback.format_exc()

        print("\n" + "=" * 70)
        print("ОШИБКА В РАСЧЁТЕ")
        print("=" * 70)
        print(error_trace)
        print("=" * 70 + "\n")

        # Возвращаем ошибку клиенту
        return jsonify({
            "success": False,
            "message": str(e)
        }), 500


@app.route('/api/download/<filename>')
def download(filename):
    """
    ЭНДПОИНТ: GET /api/download/<filename>
    НАЗНАЧЕНИЕ: Скачать файл результатов расчёта

    ПАРАМЕТРЫ:
    - filename: Имя файла (например, "result_products.xlsx")

    ЛОГИКА:
    1. Проверяет наличие файла в папке results
    2. Отправляет файл клиенту
    3. Браузер предлагает скачать файл

    ВОЗВРАЩАЕТ: Содержимое файла (Binary)
    """
    try:
        # send_from_directory отправляет файл из папки
        # as_attachment=True - браузер предлагает скачать файл
        return send_from_directory(RESULT_FOLDER, filename, as_attachment=True)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 404


@app.route('/health', methods=['GET'])
def health():
    """
    ЭНДПОИНТ: GET /health
    НАЗНАЧЕНИЕ: Проверка работоспособности сервера

    ЛОГИКА:
    Просто возвращает JSON подтверждение что сервер работает

    ВОЗВРАЩАЕТ: {"status": "ok", "message": "Сервер работает"}
    """
    return jsonify({"status": "ok", "message": "Сервер работает"})


# ============== ЗАПУСК ПРИЛОЖЕНИЯ ==============

# ============== ЗАПУСК ПРИЛОЖЕНИЯ ==============

if __name__ == '__main__':
    # Получаем порт из переменной окружения (для Render) или используем 5000
    port = int(os.environ.get('PORT', 5000))

    # Выводим информацию о запуске
    print("\n" + "=" * 70)
    print("СЕРВЕР КАЛЬКУЛЯТОРА ЦЕН OZON FBO")
    print("=" * 70)
    print(f"Папка БД: {os.path.abspath(DB_FOLDER)}")
    print(f"Папка загрузок: {os.path.abspath(UPLOAD_FOLDER)}")
    print(f"Папка результатов: {os.path.abspath(RESULT_FOLDER)}")
    print(f"Порт: {port}")
    print("=" * 70)

    # Для продакшена используется gunicorn (запускается через Procfile)
    # Для локальной разработки используем встроенный сервер Flask
    app.run(debug=False, port=port, host='0.0.0.0', threaded=True)
