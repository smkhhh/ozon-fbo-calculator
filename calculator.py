"""
calculator.py - Калькулятор цен OZON FBO
==========================================

НАЗНАЧЕНИЕ:
Программа подбирает оптимальные цены на товары с учётом целевой рентабильности
и расчётов по формулам комиссий, логистики, обработки от платформы OZON FBO.

ИСПОЛЬЗУЕМЫЕ БИБЛИОТЕКИ:
- requests: для получения данных из API сервера
- pandas: для работы с Excel и табличными данными
- math: математические функции (ceil - округление вверх)
- os: работа с файловой системой
"""

# ============== ИМПОРТЫ ==============

import requests
# requests - библиотека для HTTP запросов (GET, POST)
# Используется для получения данных из БД через API сервера

import pandas as pd
# pandas - мощная библиотека для работы с табличными данными
# Используется для чтения/записи Excel файлов и операций с DataFrames

import math
# math - встроенная библиотека математических функций
# ceil() - округляет вверх (1.1 -> 2, используется для расчёта объёма коробки)

import os
# os - работа с файловой системой

from typing import Optional

# Optional - аннотация типов (указывает что параметр может быть None)

# ============== КОНСТАНТЫ ==============

import os

# Определяем SERVER в зависимости от окружения
if os.environ.get('RENDER'):
    # В облаке Render используем текущий домен
    SERVER = "http://89.111.163.171"
else:
    # Локально используем localhost
    SERVER = "http://localhost:5000"
# Адрес сервера (localhost = текущий компьютер, порт 5000)

API_DB_LIST = f"{SERVER}/api/databases"
# Эндпоинт для получения списка БД

API_GET = lambda db, table: f"{SERVER}/api/get_data/{db}/{table}"


# Функция для формирования URL эндпоинта получения данных
# Пример использования: API_GET("ozon_data.db", "catalog")


# ============== КЛАСС REMOTEREFTABLES ==============

class RemoteRefTables:
    """
    НАЗНАЧЕНИЕ: Загрузить справочные таблицы из БД (Catalog, Logistics, New Tariffs)

    ТАБЛИЦЫ:
    - catalog: информация о категориях, типах товаров и комиссиях
    - logistics: тарифы на доставку в зависимости от объёма
    - new_tariffs: параметры обработки и дополнительные сборы
    """

    def __init__(self, db_name: str):
        """
        КОНСТРУКТОР: Инициализирует класс и загружает все таблицы из БД

        ПАРАМЕТРЫ:
        - db_name: Имя БД файла на сервере (например, "ozon_data.db")
        """
        self.db_name = db_name

        # Загружаем каждую таблицу из БД через API
        self.catalog = self._load_table("catalog")
        self.logistics = self._load_table("logistics")
        self.new_tariffs = self._load_table("new_tariffs")

    def _load_table(self, table_name: str) -> pd.DataFrame:
        """
        МЕТОД: Загрузить таблицу из БД через API

        ПАРАМЕТРЫ:
        - table_name: Имя таблицы ("catalog", "logistics", и т.д.)

        ЛОГИКА:
        1. Отправляет GET запрос к API сервера
        2. Получает JSON с данными
        3. Преобразует в pandas DataFrame
        4. Если таблица не найдена - пробует альтернативные имена
        5. Если всё равно не найдена - возвращает пустой DataFrame

        ВОЗВРАЩАЕТ: pandas DataFrame с данными таблицы
        """
        try:
            # Формируем URL и отправляем GET запрос
            r = requests.get(API_GET(self.db_name, table_name))
            r.raise_for_status()  # Проверка на ошибки HTTP

            # Получаем JSON из ответа и извлекаем поле "data"
            data = r.json().get("data", [])

            # Преобразуем список словарей в DataFrame
            return pd.DataFrame(data)
        except Exception:
            # Если первая попытка не сработала, пробуем альтернативные имена
            alt_names = [
                table_name,  # "catalog"
                table_name.capitalize(),  # "Catalog"
                table_name.replace("_", " "),  # "new tariffs"
                table_name.replace("_", " ").capitalize()  # "New tariffs"
            ]

            # Перебираем альтернативные имена
            for alt in alt_names:
                try:
                    r = requests.get(API_GET(self.db_name, alt))
                    r.raise_for_status()
                    data = r.json().get("data", [])
                    return pd.DataFrame(data)
                except:
                    pass  # Если эта попытка не сработала - пробуем следующую

        # Если всё не сработало - возвращаем пустой DataFrame
        return pd.DataFrame()


# ============== СЛУЖЕБНЫЕ ФУНКЦИИ ==============

def get_latest_db() -> str:
    """
    ФУНКЦИЯ: Получить имя последней (самой новой) БД на сервере

    ЛОГИКА:
    1. Отправляет GET запрос к API для получения списка БД
    2. Сортирует БД по убыванию (от новых к старым)
    3. Возвращает первый (самый новый) файл

    ВОЗВРАЩАЕТ: Имя БД файла (строка)
    """
    # Получаем список всех БД с сервера
    r = requests.get(API_DB_LIST)
    r.raise_for_status()
    dbs = r.json().get("databases", [])

    if not dbs:
        raise RuntimeError("Нет БД на сервере")

    # Сортируем в обратном порядке (новые в начале)
    sorted_dbs = sorted(dbs, reverse=True)

    # Берём первую (самую новую)
    latest_db = sorted_dbs[0]
    print(f"БД: {latest_db} (из {len(dbs)})")

    return latest_db


def normalize_percent(val):
    """
    ФУНКЦИЯ: Нормализировать процентное значение в десятичную дробь

    НАЗНАЧЕНИЕ:
    - Преобразует 30% в 0.30
    - Преобразует "30%" в 0.30
    - Преобразует 0.30 в 0.30
    - Если значение > 1, разделяет на 100

    ПАРАМЕТРЫ:
    - val: Значение в любом формате (int, float, str, None)

    ВОЗВРАЩАЕТ: Десятичная дробь (0.0 - 1.0)
    """
    if val is None:
        return 0.0

    try:
        if isinstance(val, str):
            # Удаляем символ % и пробелы, заменяем запятую на точку
            s = val.replace("%", "").replace(",", ".").strip()
            num = float(s)
            # Если число > 1, это проценты (разделяем на 100)
            return num / 100.0 if num > 1 else num
        else:
            # Преобразуем в float
            num = float(val)
            return num / 100.0 if num > 1 else num
    except:
        return 0.0


def vlookup_approx(value, table_df: pd.DataFrame, ret_col_index=1):
    """
    ФУНКЦИЯ: Приблизительный поиск значения в таблице (как VLOOKUP в Excel)

    НАЗНАЧЕНИЕ:
    Ищет наибольшее значение в первой колонке, которое <= value
    и возвращает соответствующее значение из другой колонки

    ПРИМЕР:
    Таблица логистики: Объём | Цена
                       0.5  | 10
                       1.0  | 15
                       2.0  | 20
    value = 1.5 -> возвращает 15 (потому что 1.0 <= 1.5)

    ПАРАМЕТРЫ:
    - value: Значение для поиска
    - table_df: DataFrame с данными
    - ret_col_index: Индекс колонки для возврата (0, 1, 2, и т.д.)

    ВОЗВРАЩАЕТ: Найденное значение или None
    """
    if table_df is None or table_df.empty:
        return None

    try:
        # Получаем первую колонку и преобразуем в числа
        keys = pd.to_numeric(table_df.iloc[:, 0], errors='coerce')

        # Получаем нужную колонку для возврата
        vals = table_df.iloc[:, ret_col_index]

        # Отбираем только строки где первая колонка - число (не NaN)
        mask = ~keys.isna()
        if not mask.any():
            return None

        # Создаём новый DataFrame с парами ключ-значение
        df = pd.DataFrame({"k": keys[mask].astype(float).values,
                           "v": vals[mask].values})

        # Сортируем по ключам
        df = df.sort_values("k")

        # Ищем все значения где ключ <= value
        le = df[df["k"] <= float(value)]

        # Возвращаем последнее (наибольшее) значение
        return le.iloc[-1]["v"] if not le.empty else None
    except:
        return None


# ============== ДВУХУРОВНЕВЫЙ ПОИСК В CATALOG ==============

def two_level_catalog_lookup(catalog_df: pd.DataFrame, category, product_type, return_col_letter, debug=False):
    """
    ФУНКЦИЯ: Двухуровневый поиск в таблице Catalog

    НАЗНАЧЕНИЕ:
    Ищет комиссию для конкретного товара:
    1. Сначала фильтрует по Категории (колонка B)
    2. Потом ищет Тип товара (колонка C) в отфильтрованных строках
    3. Возвращает комиссию из нужной колонки (D/E/F/G/H в зависимости от цены)

    ПАРАМЕТРЫ:
    - catalog_df: DataFrame таблицы Catalog из БД
    - category: Название категории (например, "Инструменты для ремонта")
    - product_type: Тип товара (например, "Резец токарный")
    - return_col_letter: Буква колонки для возврата (D, E, F, G, H)
    - debug: Выводить ли отладочную информацию

    ВОЗВРАЩАЕТ: Значение комиссии (например, 0.33) или None если не найдено
    """
    if catalog_df is None or catalog_df.empty:
        if debug:
            print(f"        Таблица Catalog пуста")
        return None

    try:
        if debug:
            print(f"        Шаг 1: Фильтр по категории '{category}' в колонке B")

        # ВАЖНО: Используем имена колонок, а не индексы!
        # Это избегает проблем с переупорядочиванием колонок при загрузке из БД
        if 'B' not in catalog_df.columns or 'C' not in catalog_df.columns:
            if debug:
                print(f"        Ошибка: Колонки B или C отсутствуют!")
                print(f"        Доступные: {list(catalog_df.columns)}")
            return None

        # Шаг 1: Фильтруем по Категории (колонка B = блок категорий)
        col_category = catalog_df['B']

        if debug:
            print(f"        Первые значения колонки B: {col_category.head(3).tolist()}")

        # Ищем точное совпадение категории
        category_matches = catalog_df[col_category == category]

        # Если точное совпадение не найдено, ищем без учёта регистра
        if category_matches.empty:
            mask = col_category.astype(str).str.strip().str.lower() == str(category).strip().lower()
            category_matches = catalog_df[mask]

        if category_matches.empty:
            if debug:
                print(f"        Ошибка: Категория '{category}' не найдена в колонке B")
            return None

        if debug:
            print(f"        OK: Найдено {len(category_matches)} строк")
            print(f"        Шаг 2: Поиск типа '{product_type}' в колонке C")

        # Шаг 2: В отфильтрованных строках ищем Тип товара (колонка C)
        col_product_type = category_matches['C']
        final_matches = category_matches[col_product_type == product_type]

        # Если точное совпадение не найдено, ищем без учёта регистра
        if final_matches.empty:
            mask = col_product_type.astype(str).str.strip().str.lower() == str(product_type).strip().lower()
            final_matches = category_matches[mask]

        if final_matches.empty:
            if debug:
                print(f"        Ошибка: Тип товара '{product_type}' не найден")
            return None

        # Шаг 3: Возвращаем значение из нужной колонки
        if return_col_letter not in final_matches.columns:
            if debug:
                print(f"        Ошибка: Колонка {return_col_letter} не найдена!")
            return None

        result = final_matches[return_col_letter].iloc[0]

        if debug:
            print(f"        OK: Найдено значение {result}")

        return result

    except Exception as e:
        if debug:
            print(f"        Ошибка: {e}")
        return None


# ============== КЛАСС OZONPRICEFINDER ==============

class OzonPriceFinder:
    """
    ГЛАВНЫЙ КЛАСС: Калькулятор цен OZON FBO

    НАЗНАЧЕНИЕ:
    Инкапсулирует всю логику расчёта:
    - Загрузка справочных таблиц из БД
    - Расчёт компонентов FBO (комиссия, логистика, обработка)
    - Бинарный поиск оптимальной цены для целевой рентабильности
    """

    def __init__(self, db_name: Optional[str] = None):
        """
        КОНСТРУКТОР: Инициализирует класс

        ПАРАМЕТРЫ:
        - db_name: Имя БД (если None, берётся последняя загруженная)
        """
        if db_name is None:
            print("\nОпределение БД...")
            self.db = get_latest_db()
        else:
            self.db = db_name

        print("\nЗагрузка таблиц...")
        # Создаём объект для работы с таблицами БД
        self.refs = RemoteRefTables(self.db)

        # Выводим информацию о загруженных таблицах
        if not self.refs.catalog.empty:
            print(f"   OK Catalog: {len(self.refs.catalog)} строк")
        else:
            print("   ОШИБКА: Catalog не загружен")

        if not self.refs.logistics.empty:
            print(f"   OK Logistics: {len(self.refs.logistics)} строк")
        else:
            print("   ОШИБКА: Logistics не загружен")

        if not self.refs.new_tariffs.empty:
            print(f"   OK New Tariffs: {len(self.refs.new_tariffs)} строк")
        else:
            print("   ОШИБКА: New Tariffs не загружен")

    def get_new_tariff_cell(self, addr: str):
        """
        МЕТОД: Получить значение из таблицы New Tariffs по адресу (например, E3)

        ПАРАМЕТРЫ:
        - addr: Адрес ячейки (например, "E3" = колонка E, строка 3)

        ЛОГИКА:
        1. Парсит буквы (колонка) и цифры (строка)
        2. Преобразует буквы в индекс колонки
        3. Преобразует цифры в индекс строки
        4. Получает значение из DataFrame

        ВОЗВРАЩАЕТ: Значение ячейки или None если не найдена
        """
        df = self.refs.new_tariffs
        if df is None or df.empty:
            return None

        addr = addr.strip().upper()

        # Выделяем буквы (колонка) и цифры (строка)
        letters = ''.join([c for c in addr if c.isalpha()])
        digits = ''.join([c for c in addr if c.isdigit()])

        if not letters or not digits:
            return None

        # Преобразуем буквы в индекс колонки
        # A=0, B=1, ..., Z=25, AA=26, и т.д.
        idx = 0
        for ch in letters:
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        col = idx - 1

        # Преобразуем строку (от 1) в индекс (от 0)
        row = int(digits) - 1

        try:
            return df.iat[row, col]
        except:
            return None

    def calc_C46(self, price, category, product_type, debug=False):
        """
        МЕТОД: Рассчитать C46 = Комиссия OZON

        НАЗНАЧЕНИЕ:
        Получить процент комиссии из таблицы Catalog в зависимости от:
        - Категории товара (Инструменты для ремонта, и т.д.)
        - Типа товара (Резец токарный, и т.д.)
        - Диапазона цены (D/E/F/G/H в зависимости от суммы)

        ПАРАМЕТРЫ:
        - price: Цена товара (число)
        - category: Категория товара (строка)
        - product_type: Тип товара (строка)
        - debug: Выводить ли отладку

        ВОЗВРАЩАЕТ: Комиссия как десятичная дробь (0.33) или пустая строка
        """
        if not category or not product_type:
            if debug:
                print(f"      C46: категория или тип товара пустые")
            return ""

        try:
            p = float(price)
        except:
            if debug:
                print(f"      C46: ошибка парсинга цены")
            return ""

        # Определяем диапазон цены и выбираем соответствующую колонку
        # D = до 100, E = 101-300, F = 301-500, G = 501-1500, H = свыше 1500
        if p <= 100:
            col = 'D'
        elif p <= 300:
            col = 'E'
        elif p <= 500:
            col = 'F'
        elif p <= 1500:
            col = 'G'
        else:
            col = 'H'

        if debug:
            print(f"      C46: категория='{category}', тип='{product_type}', цена={p:.2f}, колонка={col}")

        # Используем двухуровневый поиск в Catalog
        val = two_level_catalog_lookup(self.refs.catalog, category, product_type, col, debug=debug)

        if val is None:
            if debug:
                print(f"      C46: не найдено в БД")
            return ""

        try:
            result = float(val)
            # Если значение > 1, это проценты (30 вместо 0.30), разделяем
            if result > 1:
                result = result / 100.0
            if debug:
                print(f"      C46: результат = {result} ({result * 100:.1f}%)")
            return result
        except:
            if debug:
                print(f"      C46: ошибка преобразования")
            return ""

    def calc_C50(self, price, C19, C34, debug=False):
        """
        МЕТОД: Рассчитать C50 = Логистика (доставка)

        НАЗНАЧЕНИЕ:
        Рассчитать стоимость доставки товара в зависимости от цены и объёма

        ФОРМУЛА:
        - Если цена > 300: берётся формула с минимумом 46 руб (до 1л), потом +15 за каждый литр
        - Если цена <= 300: берётся ВПР из таблицы Logistics по объёму

        ПАРАМЕТРЫ:
        - price: Цена товара
        - C19: Объём товара в литрах
        - C34: Переменная для проверки что расчёт активен (цена > 0)
        - debug: Отладка

        ВОЗВРАЩАЕТ: Стоимость логистики (число) или пустая строка
        """
        if C34 == "":
            return ""

        try:
            p = float(price)
            vol = float(C19) if C19 != "" else 0
        except:
            return ""

        # Для дорогих товаров (> 300 руб) используется фиксированная схема
        if p > 300:
            if vol > 0:
                # Минимальный тариф: 46 руб за 1 литр
                if vol <= 1:
                    result = 46
                # 56 руб за 2 литра
                elif vol <= 2:
                    result = 56
                # 66 руб за 3 литра
                elif vol <= 3:
                    result = 66
                # Каждый литр после 3-го: +15 руб
                elif vol <= 190:
                    result = 66 + (math.ceil(vol) - 3) * 15
                # Максимум 2871 руб
                else:
                    result = 2871

                if debug:
                    print(f"      C50: цена={p:.2f}, объём={vol:.3f}л -> {result} руб")
                return result
            else:
                return ""
        else:
            # Для дешёвых товаров (<=300 руб) используется таблица Logistics
            if self.refs.logistics is not None and not self.refs.logistics.empty:
                # ВПР (приблизительный поиск) по объёму
                result = vlookup_approx(vol, self.refs.logistics, ret_col_index=1)
                try:
                    res = float(result) if result is not None else ""
                    if debug:
                        print(f"      C50: ВПР по объёму={vol:.3f} -> {res}")
                    return res
                except:
                    return ""
            return ""

    def calc_C51(self, price, C34, debug=False):
        """
        МЕТОД: Рассчитать C51 = Обработка товара

        НАЗНАЧЕНИЕ:
        Рассчитать стоимость обработки и упаковки товара (если он продаётся через FBO)

        ФОРМУЛА:
        - Берём E3 из таблицы Новые тарифы
        - Множим на цену: E3 * price
        - Ограничиваем минимумом G5 и максимумом G6

        ПАРАМЕТРЫ:
        - price: Цена товара
        - C34: Переменная для проверки (цена > 0)
        - debug: Отладка

        ВОЗВРАЩАЕТ: Стоимость обработки (число) или пустая строка
        """
        if C34 == "" or C34 <= 0:
            return ""

        try:
            # Получаем параметры из таблицы Новые тарифы
            E3 = self.get_new_tariff_cell('E3')  # Коэффициент обработки
            G5 = self.get_new_tariff_cell('G5')  # Минимальная сумма
            G6 = self.get_new_tariff_cell('G6')  # Максимальная сумма

            if E3 is None or G5 is None or G6 is None:
                return ""

            e3 = float(E3)
            g5 = float(G5)
            g6 = float(G6)
            p = float(price)

            # Рассчитываем стоимость: коэффициент * цена
            val = e3 * p

            # Ограничиваем результат минимумом и максимумом
            if val < g5:
                result = g5
            elif val > g6:
                result = g6
            else:
                result = val

            if debug:
                print(f"      C51: E3={e3}, цена={p:.2f}, расчёт={val:.2f}, {g5}<{result}<{g6}")
            return result
        except Exception as e:
            if debug:
                print(f"      C51: ошибка {e}")
            return ""

    def calc_C55(self, C34, category, debug=False):
        """
        МЕТОД: Рассчитать C55 = Дополнительный сбор

        НАЗНАЧЕНИЕ:
        Дополнительные комиссионные сборы (если применимо для категории)

        ПАРАМЕТРЫ:
        - C34: Цена (для проверки)
        - category: Категория товара (для проверки)
        - debug: Отладка

        ВОЗВРАЩАЕТ: Размер сбора или пустая строка
        """
        if C34 == "" or category == "":
            return ""

        try:
            # Берём значение из ячейки C10 таблицы Новые тарифы
            val = self.get_new_tariff_cell('C10')
            result = float(val) if (val is not None and str(val).strip() != "") else ""
            if debug:
                print(f"      C55: C10={val}")
            return result
        except:
            return ""

    def compute_fbo_for_price(self, price, row, debug=False):
        """
        МЕТОД: Полный расчёт FBO при заданной цене

        НАЗНАЧЕНИЕ:
        Рассчитать все компоненты FBO (комиссия, логистика, обработка) для конкретной цены

        ЛОГИКА РАСЧЁТА FBO:
        1. Получаем комиссию (C46) и эквайринг (C47) -> умножаем на цену
        2. Получаем логистику (C50) и обработку (C51)
        3. Рассчитываем возвраты (C52) исходя из % выкупа
        4. Суммируем всё: (комиссия + эквайринг) * цена + логистика + возвраты
        5. Делим на цену, получаем процент FBO (C45)
        6. Умножаем процент на цену, получаем итоговый FBO (C44)

        ПАРАМЕТРЫ:
        - price: Цена товара
        - row: Строка из DataFrame с информацией о товаре
        - debug: Выводить отладку

        ВОЗВРАЩАЕТ: Dict с результатами всех вычислений
        """
        # Получаем данные из строки товара
        C9 = row.get("Категория")
        C11 = row.get("Тип товара")

        # Рассчитываем объём товара из размеров (см -> дм -> литры)
        try:
            C16 = float(row.get("Длина, мм*", 0)) / 10.0  # Длина в см
            C17 = float(row.get("Высота, мм*", 0)) / 10.0  # Высота в см
            C18 = float(row.get("Ширина, мм*", 0)) / 10.0  # Ширина в см
        except:
            C16 = C17 = C18 = 0.0

        # Объём в литрах (см³ / 1000)
        C19 = (C16 * C17 * C18) / 1000.0 if (C16 > 0 and C17 > 0 and C18 > 0) else 0.0

        # Получаем вес товара
        try:
            C20 = float(row.get("Вес, г*", 0)) / 1000.0  # Вес в кг
        except:
            C20 = 0.0

        # Процент выкупа (нормализуем в 0.0-1.0)
        C21 = normalize_percent(row.get("% Выкупа"))

        # Проверяем обязательные поля
        if not C9 or not C11:
            return {
                "C33": "", "C34": "", "C46": "", "C47": "",
                "C48": "", "C50": "", "C51": "", "C52": "",
                "C44": "", "C45": "", "FBO": 0
            }

        # C34 - цена товара для расчёта
        C34 = float(price) if price > 0 else ""

        # C46 = комиссия OZON
        C46 = self.calc_C46(price, C9, C11, debug=debug)

        # C47 = эквайринг (1.5% от суммы)
        C47 = 0.015 if C34 != "" else ""

        # C50 = логистика
        C50 = self.calc_C50(price, C19, C34, debug=debug)

        # C51 = обработка
        C51 = self.calc_C51(price, C34, debug=debug)

        # C48 = C50 + C51 (всего логистика + обработка)
        try:
            if C50 == "" and C51 == "":
                C48 = ""
            elif C50 == "":
                C48 = C51
            elif C51 == "":
                C48 = C50
            else:
                C48 = float(C50) + float(C51)
        except:
            C48 = ""

        C54 = C50
        C55 = self.calc_C55(C34, C9, debug=debug)

        # C53 = C54 + C55 (логистика + дополнительный сбор)
        if C34 != "" and C9 != "":
            if C48 != "" and C48 > 0:
                sum_val = 0
                if C54 != "":
                    sum_val += float(C54)
                if C55 != "":
                    sum_val += float(C55)
                C53 = sum_val
            else:
                C53 = ""
        else:
            C53 = ""

        # C52 = возвраты = C53 * (1 - % выкупа)
        # Это доля логистики, приходящаяся на возвращённые товары
        try:
            if C53 != "" and C53 > 0 and C21 > 0:
                C52 = float(C53) * (1 - C21)
            else:
                C52 = ""
        except:
            C52 = ""

        # C45 = Итоговый процент FBO ко всем расходам
        # Формула: ((Комиссия + Эквайринг) * Цена + Логистика + Возвраты) / Цена
        try:
            if C34 != "" and C34 > 0:
                numerator = 0

                # Комиссионная часть
                commission_part = 0
                if C46 != "" and C47 != "":
                    commission_part = (float(C46) + float(C47)) * float(C34)
                    numerator += commission_part

                # Логистика и обработка
                if C48 != "":
                    numerator += float(C48)

                # Возвраты
                if C52 != "":
                    numerator += float(C52)

                # Делим на цену, получаем процент
                C45 = numerator / float(C34)
            else:
                C45 = ""
        except:
            C45 = ""

        # C44 = итоговый FBO в рублях = C45 * Цена
        try:
            if C45 != "" and C34 != "":
                C44 = float(C45) * float(C34)
            else:
                C44 = ""
        except:
            C44 = ""

        return {
            "C33": f"{C9} / {C11}", "C34": C34, "C46": C46, "C47": C47,
            "C48": C48, "C50": C50, "C51": C51, "C52": C52,
            "C44": C44, "C45": C45,
            "FBO": float(C44) if C44 != "" else 0
        }

    def find_price_for_target_margin(self, row, target_margin: float, debug=False):
        """
        МЕТОД: Бинарный поиск оптимальной цены для целевой рентабильности

        НАЗНАЧЕНИЕ:
        Найти цену товара, при которой рентабильность = target_margin

        ФОРМУЛА РЕНТАБИЛЬНОСТИ:
        Рентабильность = (Цена - Себестоимость - FBO) / Цена

        АЛГОРИТМ:
        1. Устанавливаем начальный диапазон цен (low - min, high - max)
        2. Вычисляем среднюю цену (mid)
        3. Рассчитываем FBO и рентабильность при этой цене
        4. Если рентабильность < target -> повышаем цену (low = mid)
        5. Если рентабильность > target -> понижаем цену (high = mid)
        6. Повторяем до сходимости (разница < 0.05%)

        ПАРАМЕТРЫ:
        - row: Строка с информацией о товаре
        - target_margin: Целевая рентабильность (0.0-1.0)
        - debug: Выводить отладку

        ВОЗВРАЩАЕТ: Dict с оптимальной ценой, FBO и рентабильностью
        """
        # Получаем себестоимость товара
        себестоимость = float(row.get("Себестоимость") or 0.0)

        # Устанавливаем начальный диапазон поиска цен
        low = max(1.0, себестоимость * 1.2)  # Минимум: себестоимость * 120%
        high = max(low * 5.0, low + 5000.0)  # Максимум: либо 5x, либо +5000 руб

        # Инициализируем результаты (на случай если поиск не сходится)
        best = {"price": None, "fbo": None, "margin_diff": None,
                "margin_pct": None, "debug": None}

        # Параметры сходимости
        tol = 0.0005  # Допустимая разница: 0.05%
        max_iterations = 100  # Максимум итераций

        # Выполняем бинарный поиск
        for iteration in range(max_iterations):
            # Вычисляем среднюю цену
            mid = (low + high) / 2.0

            # Показываем отладку только на итерациях 0 и 10
            show_detail = debug and (iteration == 0 or iteration == 10)

            # Рассчитываем FBO для этой цены
            pieces = self.compute_fbo_for_price(mid, row, debug=show_detail)
            fbo_mid = pieces["FBO"]

            # Рассчитываем рентабильность: (Цена - Себестоимость - FBO) / Цена
            profit = mid - себестоимость - fbo_mid
            margin = (profit / mid) if mid > 0 else -1.0
            diff = margin - target_margin  # Разница между целевой и текущей

            # Выводим информацию каждые 10 итераций
            if debug and iteration % 10 == 0:
                print(f"  Итер {iteration}: Цена={mid:.2f}, FBO={fbo_mid:.2f}, "
                      f"Рентаб={margin * 100:.2f}%, Разница={diff * 100:.3f}%")

            # Сохраняем лучший результат (ближайший к целевой рентабильности)
            if (best["margin_diff"] is None) or (abs(diff) < abs(best["margin_diff"])):
                best["margin_diff"] = diff
                best["price"] = mid
                best["fbo"] = fbo_mid
                best["margin_pct"] = margin * 100.0
                best["debug"] = pieces

            # Проверяем сходимость: если разница <= допустимой - выходим
            if abs(diff) <= tol:
                break

            # Сужаем диапазон поиска
            # Если рентабильность меньше целевой - нужна более высокая цена
            if margin < target_margin:
                low = mid
            else:
                high = mid

            # Если диапазон стал очень узким - выходим
            if abs(high - low) < 0.01:
                break

        # Форматируем результаты
        if best["price"] is not None:
            best["price"] = round(best["price"], 2)
            best["fbo"] = round(best["fbo"], 2)
            best["margin_pct"] = round(best["margin_pct"], 2)

        return best


def calculate_file(input_excel: str, target_margin_pct: float = 20.0,
                   output_excel: str = "calculated_result.xlsx", debug_first=False):
    """
    ФУНКЦИЯ: Главная функция для расчёта цен для всех товаров из Excel файла

    НАЗНАЧЕНИЕ:
    1. Читает Excel файл с товарами
    2. Для каждого товара находит оптимальную цену
    3. Сохраняет ВСЕ исходные данные + добавляет 3 новые колонки

    ПАРАМЕТРЫ:
    - input_excel: Путь к входному Excel файлу
    - target_margin_pct: Целевая рентабильность в процентах (20.0 = 20%)
    - output_excel: Путь к выходному файлу с результатами
    - debug_first: Выводить ли отладку для первого товара

    ЛОГИКА:
    1. Проверяет файл
    2. Читает Excel
    3. Инициализирует калькулятор
    4. Для каждого товара:
       - Вызывает find_price_for_target_margin()
       - Добавляет результаты в новые колонки
    5. Записывает полный DataFrame с новыми колонками в Excel
    """
    print("=" * 70)
    print("КАЛЬКУЛЯТОР ЦЕН OZON FBO")
    print("=" * 70)

    # Преобразуем в абсолютные пути
    input_excel = os.path.abspath(input_excel)
    output_excel = os.path.abspath(output_excel)

    print(f"\nВходной файл: {input_excel}")
    print(f"Файл существует: {os.path.exists(input_excel)}")

    if not os.path.exists(input_excel):
        raise FileNotFoundError(f"Файл не найден: {input_excel}")

    print(f"Размер: {os.path.getsize(input_excel) / 1024:.2f} KB")

    # Читаем Excel файл
    try:
        print(f"\nЧтение Excel...")
        df = pd.read_excel(input_excel, engine='openpyxl')
        print(f"OK Прочитано {len(df)} товаров")
    except Exception as e:
        print(f"ОШИБКА: {e}")
        raise

    print(f"\nКолонки: {list(df.columns)}")

    # Инициализируем калькулятор
    pf = OzonPriceFinder()

    # Создаём копию исходного DataFrame для сохранения всех данных
    result_df = df.copy()

    # Добавляем новые колонки для результатов
    result_df['Итоговая цена (руб)'] = None
    result_df['FBO расходы (руб)'] = None
    result_df['Рентабельность (%)'] = None

    target_margin = float(target_margin_pct) / 100.0

    print(f"\nЦелевая рентабильность: {target_margin_pct}%")
    print(f"Товаров: {len(df)}\n")

    # Обрабатываем каждый товар
    for idx, row in df.iterrows():
        name = row.get("Название")
        if pd.isna(name) or str(name).strip() == "":
            continue

        try:
            print(f"[{idx + 1}/{len(df)}]: {name[:40]}...", end=" ")

            # Показываем отладку для первого товара
            show_debug = debug_first and idx == 0
            if show_debug:
                print("\n  ОТЛАДКА:")

            # Ищем оптимальную цену
            best = pf.find_price_for_target_margin(row, target_margin, debug=show_debug)

            if show_debug:
                print("  КОНЕЦ ОТЛАДКИ\n")

            # Выводим результат
            print(f"OK {best['price']} руб, FBO: {best['fbo']} руб, {best['margin_pct']}%")

            # Записываем результаты в соответствующую строку DataFrame
            result_df.at[idx, 'Итоговая цена (руб)'] = best["price"]
            result_df.at[idx, 'FBO расходы (руб)'] = best["fbo"]
            result_df.at[idx, 'Рентабельность (%)'] = best["margin_pct"]

        except Exception as e:
            print(f"ОШИБКА {e}")
            # Оставляем пустые значения в случае ошибки
            result_df.at[idx, 'Итоговая цена (руб)'] = None
            result_df.at[idx, 'FBO расходы (руб)'] = None
            result_df.at[idx, 'Рентабельность (%)'] = None

    # Записываем ВЕСЬ DataFrame (исходные данные + новые колонки) в Excel
    result_df.to_excel(output_excel, index=False)

    print("\n" + "=" * 70)
    print("ГОТОВО!")
    print("=" * 70)
    print(f"Обработано: {len(result_df)} товаров")
    print(f"Исходных колонок: {len(df.columns)}")
    print(f"Добавлено колонок: 3")
    print(f"Всего колонок: {len(result_df.columns)}")
    print(f"Файл результатов: {output_excel}")
    print("=" * 70)

    return result_df


# ============== ЗАПУСК ИЗ КОМАНДНОЙ СТРОКИ ==============

if __name__ == '__main__':
    # Если запустить этот файл напрямую, а не импортировать его
    path = input("Excel файл: ").strip()
    margin = float(input("Рентабильность (%): ").strip())
    debug = input("Отладка? (y/n): ").strip().lower() == 'y'
    calculate_file(path, margin, debug_first=debug)
