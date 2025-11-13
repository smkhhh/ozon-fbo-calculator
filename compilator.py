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

import pandas as pd
import os
import sqlite3
import requests
import datetime

class OzonCompiler:
    def __init__(self):
        # Жестко задан адрес вашего VPS
        self.server = "http://89.111.163.171"
        self.db_name = None

    def clear_old_databases(self):
        print("\nОчистка старых баз данных...")
        try:
            response = requests.get(f"{self.server}/api/databases")
            response.raise_for_status()
            databases = response.json().get("databases", [])
            if not databases:
                print("   OK Нет старых БД")
                return
            print(f"   Найдено БД: {len(databases)}")
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
        print("\nЧтение таблиц из Excel...")
        tables = {}
        try:
            print("   Catalog...", end=" ")
            df_catalog = pd.read_excel(excel_path, sheet_name="Catalog",
                                       usecols="A:L", skiprows=1)
            df_catalog = df_catalog.dropna(how='all')
            df_catalog.columns = ['Index', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
            tables['catalog'] = df_catalog
            print(f"OK {len(df_catalog)} строк")
            # ...Дополнительные проверки данных, аналогично вашему коду...
            print("   Logistics...", end=" ")
            df_logistics = pd.read_excel(excel_path, sheet_name="Logistics")
            df_logistics = df_logistics.dropna(how='all')
            tables['logistics'] = df_logistics
            print(f"OK {len(df_logistics)} строк")
            print("   New Tariffs...", end=" ")
            try:
                df_tariffs = pd.read_excel(excel_path, sheet_name="Новые тарифы", header=None)
                df_tariffs = df_tariffs.dropna(how='all')
                tables['new_tariffs'] = df_tariffs
                print(f"OK {len(df_tariffs)} строк")
            except Exception as e:
                print(f"ОШИБКА: {e}")
                tables['new_tariffs'] = pd.DataFrame()
            return tables
        except Exception as e:
            print(f"\nОШИБКА: {e}")
            import traceback
            traceback.print_exc()
            return None

    def create_database(self, tables):
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.db_name = f"ozon_data_{timestamp}.db"
        print(f"\nСоздание БД: {self.db_name}")
        conn = sqlite3.connect(self.db_name)
        for table_name, df in tables.items():
            print(f"   {table_name}...", end=" ")
            df.to_sql(table_name, conn, if_exists='replace', index=False)
            print(f"OK {len(df)} строк")
        conn.close()
        print(f"\nOK БД создана: {self.db_name}")
        return self.db_name

    def upload_to_server(self, db_path: str):
        print(f"\nЗагрузка БД на сервер...")
        try:
            with open(db_path, 'rb') as f:
                files = {'file': (os.path.basename(db_path), f, 'application/x-sqlite3')}
                response = requests.post(f"{self.server}/api/upload_database", files=files)
                response.raise_for_status()
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
        print("\n" + "=" * 70)
        print("КОМПИЛЯТОР ДАННЫХ OZON")
        print("=" * 70)
        self.clear_old_databases()
        tables = self.load_excel_tables(excel_path)
        if not tables:
            return False
        db_path = self.create_database(tables)
        if not db_path:
            return False
        success = self.upload_to_server(db_path)
        if success:
            try:
                os.remove(db_path)
                print(f"\nОчистка: Локальная БД удалена")
            except Exception as e:
                print(f"\nПредупреждение: Не удалось удалить локальную БД: {e}")
        print("\n" + "=" * 70)
        print("ГОТОВО!" if success else "ЗАВЕРШЕНО С ОШИБКАМИ")
        print("=" * 70)
        return success

if __name__ == '__main__':
    compiler = OzonCompiler()
    excel_file = input("Путь к Excel файлу: ").strip()
    compiler.compile(excel_file)