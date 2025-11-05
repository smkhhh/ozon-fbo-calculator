"""
compilator.py - Компилятор данных для OZON FBO
================================================

НАЗНАЧЕНИЕ:
Утилита для подготовки данных:
1. Читает справочные таблицы из Excel (Catalog, Logistics, New Tariffs)
2. Создаёт SQLite БД
3. Загружает БД на сервер через API

ИСПОЛЬЗУЕМЫЕ БИБЛИОТЕКИ:
- pandas: чтение Excel файлов
- sqlite3: создание и работа с БД
- requests: HTTP запросы для загрузки на сервер
- os: работа с файловой системой
- datetime: получение текущей даты/времени для имён файлов
"""

# ============== ИМПОРТЫ ==============

import pandas as pd
# pandas - библиотека для работы с табличными данными
# read_excel() - чтение Excel файлов в DataFrame

import sqlite3
# sqlite3 - встроенная библиотека для создания и работы с SQLite БД

import requests
# requests - библиотека для HTTP запросов
# Используется для загрузки БД файла на сервер

import os
# os - работа с файловой системой
# makedirs() - создание папок
# getsize() - получение размера файла

import datetime


# datetime - работа с датой и временем
# Используется для формирования имён файлов с текущей датой/временем

# ============== КЛАСС OZONCOMPILER ==============

class OzonCompiler:
    """
    ГЛАВНЫЙ КЛАСС: Компилятор данных OZON

    НАЗНАЧЕНИЕ:
    Инкапсулирует весь процесс подготовки БД:
    - Загрузка таблиц из Excel
    - Создание SQLite БД
    - Загрузка на сервер
    """

    def __init__(self):
        """
        КОНСТРУКТОР: Инициализирует компилятор

        ПЕРЕМЕННЫЕ ЭКЗЕМПЛЯРА:
        - server: Адрес сервера для загрузки БД
        - db_name: Имя созданного БД файла
        """
        self.server = "http://localhost:5000"
        self.db_name = None

    def clear_old_databases(self):
        """
        МЕТОД: Удалить все старые БД с сервера

        ЛОГИКА:
        1. Получает список всех БД на сервере
        2. Для каждой БД отправляет DELETE запрос
        3. Выводит информацию об удалении

        НАЗНАЧЕНИЕ:
        Освобождает место на сервере перед загрузкой новой БД
        """
        print("\nОчистка старых баз данных...")
        try:
            # Получаем список БД
            response = requests.get(f"{self.server}/api/databases")
            response.raise_for_status()
            databases = response.json().get("databases", [])

            if not databases:
                print("   OK Нет старых БД")
                return

            print(f"   Найдено БД: {len(databases)}")

            # Удаляем каждую БД
            for db in databases:
                try:
                    delete_response = requests.delete(f"{self.server}/api/delete_database/{db}")
                    if delete_response.status_code == 200:
                        print(f"   OK Удалена: {db}")
                    else:
                        print(f"   ОШИБКА: Не удалось удалить {db}")
                except Exception as e:
                    print(f"   ОШИБКА при удалении {db}: {e}")
        except Exception as e:
            print(f"   ОШИБКА при получении списка БД: {e}")

    def load_excel_tables(self, excel_path: str):
        """
        МЕТОД: Загрузить таблицы из Excel файла

        НАЗНАЧЕНИЕ:
        Читает три справочные таблицы:
        1. Catalog (структура: Категория | Тип товара | Комиссии по диапазонам цен)
        2. Logistics (структура: Объём | Стоимость доставки)
        3. New Tariffs (структура: Параметры обработки и сборов)

        ЛОГИКА:
        1. Для Catalog:
           - Читает только колонки A:L (ограничиваем размер)
           - Пропускает первую строку (заголовки)
           - Переименовывает колонки в B, C, D, E, F, G, H и т.д.
        2. Для остальных таблиц: читает как есть
        3. Удаляет пустые строки
        4. Выводит информацию о загруженных данных

        ПАРАМЕТРЫ:
        - excel_path: Путь к Excel файлу

        ВОЗВРАЩАЕТ: Dict с DataFrames всех таблиц
        """
        print("\nЧтение таблиц из Excel...")

        tables = {}

        try:
            # ============== ЗАГРУЗКА CATALOG ==============
            print("   Catalog...", end=" ")

            # Читаем только колонки A:L (ограничиваем объём данных)
            df_catalog = pd.read_excel(excel_path, sheet_name="Catalog",
                                       usecols="A:L", skiprows=1)  # skiprows=1 пропускает заголовок

            # Удаляем полностью пустые строки
            df_catalog = df_catalog.dropna(how='all')

            # Переименовываем колонки в буквы (удобнее для работы)
            # Index -> индекс, B -> категория, C -> тип товара, D-H -> комиссии
            df_catalog.columns = ['Index', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']

            tables['catalog'] = df_catalog
            print(f"OK {len(df_catalog)} строк")

            # Выводим примеры категорий
            if len(df_catalog) > 0:
                print(f"\n   Примеры блоков категорий:")
                unique_blocks = df_catalog['B'].dropna().unique()[:5]
                for i, block in enumerate(unique_blocks, 1):
                    print(f"      {i}. {block}")

                # Проверяем наличие нужной категории
                search_category = "Инструменты для ремонта и строительства"
                if search_category in df_catalog['B'].values:
                    print(f"\n   OK Найдена категория: '{search_category}'")
                    matching_rows = df_catalog[df_catalog['B'] == search_category]
                    print(f"   Типов товаров: {len(matching_rows)}")
                    print(f"   Примеры:")
                    for product_type in matching_rows['C'].head(3):
                        print(f"      - {product_type}")
                else:
                    print(f"\n   ОШИБКА: Категория '{search_category}' не найдена")
                    print(f"   Попробуй поискать похожие:")
                    matching = df_catalog[df_catalog['B'].str.contains("Инструмент", case=False, na=False)]
                    if not matching.empty:
                        for cat in matching['B'].unique()[:5]:
                            print(f"      - {cat}")

            # ============== ЗАГРУЗКА LOGISTICS ==============
            print("\n   Logistics...", end=" ")
            df_logistics = pd.read_excel(excel_path, sheet_name="Logistics")
            df_logistics = df_logistics.dropna(how='all')
            tables['logistics'] = df_logistics
            print(f"OK {len(df_logistics)} строк")

            # ============== ЗАГРУЗКА NEW TARIFFS ==============
            print("   New Tariffs...", end=" ")
            try:
                # Берём лист "Новые тарифы" БЕЗ заголовков (header=None)
                # Так как нам нужны значения по адресам вроде E3, G5 (индексы)
                df_tariffs = pd.read_excel(excel_path, sheet_name="Новые тарифы", header=None)
                df_tariffs = df_tariffs.dropna(how='all')
                tables['new_tariffs'] = df_tariffs
                print(f"OK {len(df_tariffs)} строк")

                # Проверяем ключевые ячейки
                print(f"\n   Проверка ключевых ячеек:")
                try:
                    e3_val = df_tariffs.iat[2, 4]  # E3 = строка 2 (0-indexed), колонка 4
                    print(f"      E3 = {e3_val} (тип: {type(e3_val).__name__})")
                except:
                    print(f"      E3 = не найдена")

                try:
                    g5_val = df_tariffs.iat[4, 6]  # G5 = строка 4, колонка 6
                    print(f"      G5 = {g5_val} (тип: {type(g5_val).__name__})")
                except:
                    print(f"      G5 = не найдена")

                try:
                    g6_val = df_tariffs.iat[5, 6]  # G6 = строка 5, колонка 6
                    print(f"      G6 = {g6_val} (тип: {type(g6_val).__name__})")
                except:
                    print(f"      G6 = не найдена")

            except Exception as e:
                print(f"ОШИБКА: {e}")
                tables['new_tariffs'] = pd.DataFrame()

            return tables

        except Exception as e:
            print(f"\nОШИБКА: {e}")
            import traceback
            traceback.print_exc()
            return None

    def create_database(self, tables: dict):
        """
        МЕТОД: Создать SQLite БД из таблиц

        ЛОГИКА:
        1. Формирует имя БД с текущей датой/временем
        2. Создаёт файл БД
        3. Сохраняет каждую таблицу в БД
        4. Закрывает БД
        5. Выводит информацию

        ПАРАМЕТРЫ:
        - tables: Dict с DataFrames всех таблиц

        ВОЗВРАЩАЕТ: Путь к созданному БД файлу
        """
        # Формируем имя с текущей датой и временем
        # Формат: ozon_data_YYYYMMDD_HHMMSS.db
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.db_name = f"ozon_data_{timestamp}.db"

        print(f"\nСоздание БД: {self.db_name}")

        # Создаём (или открываем) SQLite БД
        conn = sqlite3.connect(self.db_name)

        # Сохраняем каждую таблицу в БД
        for table_name, df in tables.items():
            print(f"   {table_name}...", end=" ")

            # to_sql() сохраняет DataFrame в БД
            # if_exists='replace' - переписывает если таблица уже существует
            # index=False - не сохраняем индекс DataFrame
            df.to_sql(table_name, conn, if_exists='replace', index=False)

            print(f"OK {len(df)} строк")

        # Закрываем БД
        conn.close()

        print(f"\nOK БД создана: {self.db_name}")
        return self.db_name

    def upload_to_server(self, db_path: str):
        """
        МЕТОД: Загрузить БД файл на сервер

        ЛОГИКА:
        1. Открывает БД файл в бинарном режиме
        2. Отправляет POST запрос на сервер с файлом
        3. Проверяет ответ
        4. Выводит результат

        ПАРАМЕТРЫ:
        - db_path: Путь к локальному БД файлу

        ВОЗВРАЩАЕТ: True если успешно, False если ошибка
        """
        print(f"\nЗагрузка БД на сервер...")

        try:
            # Открываем БД файл в бинарном режиме для передачи
            with open(db_path, 'rb') as f:
                # Формируем данные для отправки (multipart/form-data)
                files = {'file': (os.path.basename(db_path), f, 'application/x-sqlite3')}

                # Отправляем POST запрос на сервер
                response = requests.post(f"{self.server}/api/upload_database", files=files)
                response.raise_for_status()  # Проверяем на ошибки HTTP

                # Получаем JSON ответ
                result = response.json()

                if result.get("success"):
                    print(f"OK БД загружена: {result.get('database_name')}")
                    return True
                else:
                    print(f"ОШИБКА: {result.get('message')}")
                    return False
        except Exception as e:
            print(f"ОШИБКА: {e}")
            return False

    def compile(self, excel_path: str):
        """
        МЕТОД: Полный цикл компиляции БД

        ЛОГИКА:
        1. Очищает старые БД на сервере
        2. Загружает таблицы из Excel
        3. Создаёт SQLite БД
        4. Загружает на сервер
        5. Удаляет локальный БД файл
        6. Выводит итоги

        ПАРАМЕТРЫ:
        - excel_path: Путь к Excel файлу с данными

        ВОЗВРАЩАЕТ: True если успешно, False если ошибка
        """
        print("\n" + "=" * 70)
        print("КОМПИЛЯТОР ДАННЫХ OZON")
        print("=" * 70)

        # Шаг 1: Очистка старых БД
        self.clear_old_databases()

        # Шаг 2: Загрузка таблиц из Excel
        tables = self.load_excel_tables(excel_path)
        if not tables:
            return False

        # Шаг 3: Создание БД
        db_path = self.create_database(tables)
        if not db_path:
            return False

        # Шаг 4: Загрузка на сервер
        success = self.upload_to_server(db_path)

        # Шаг 5: Удаление локального БД файла (если успешно)
        if success:
            try:
                os.remove(db_path)
                print(f"\nОчистка: Локальная БД удалена")
            except Exception as e:
                print(f"\nПредупреждение: Не удалось удалить локальную БД: {e}")

        # Выводим итоги
        print("\n" + "=" * 70)
        print("ГОТОВО!" if success else "ЗАВЕРШЕНО С ОШИБКАМИ")
        print("=" * 70)

        return success


# ============== ЗАПУСК ИЗ КОМАНДНОЙ СТРОКИ ==============

if __name__ == '__main__':
    """
    ЗАПУСК:
    Если запустить этот файл напрямую (не импортировать),
    программа попросит ввести путь к Excel файлу
    """
    compiler = OzonCompiler()

    # Запрашиваем путь к Excel файлу
    excel_file = input("Путь к Excel файлу: ").strip()

    # Запускаем компиляцию
    compiler.compile(excel_file)
