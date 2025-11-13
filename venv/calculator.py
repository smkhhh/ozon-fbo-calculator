# calculator.py

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

import requests
import pandas as pd
import math
import os
from typing import Optional


SERVER = "http://89.111.163.171"

API_DB_LIST = f"{SERVER}/api/databases"
API_GET = lambda db, table: f"{SERVER}/api/get_data/{db}/{table}"

class RemoteRefTables:
    def __init__(self, db_name: str):
        self.db_name = db_name
        self.catalog = self._load_table("catalog")
        self.logistics = self._load_table("logistics")
        self.new_tariffs = self._load_table("new_tariffs")

    def _load_table(self, table_name: str) -> pd.DataFrame:
        try:
            r = requests.get(API_GET(self.db_name, table_name))
            r.raise_for_status()
            data = r.json().get("data", [])
            return pd.DataFrame(data)
        except Exception:
            alt_names = [
                table_name,
                table_name.capitalize(),
                table_name.replace("_", " "),
                table_name.replace("_", " ").capitalize()
            ]
            for alt in alt_names:
                try:
                    r = requests.get(API_GET(self.db_name, alt))
                    r.raise_for_status()
                    data = r.json().get("data", [])
                    return pd.DataFrame(data)
                except:
                    continue
            return pd.DataFrame()

def get_latest_db() -> str:
    r = requests.get(API_DB_LIST)
    r.raise_for_status()
    dbs = r.json().get("databases", [])
    if not dbs:
        raise RuntimeError("Нет БД на сервере")
    sorted_dbs = sorted(dbs, reverse=True)
    latest_db = sorted_dbs[0]
    print(f"БД: {latest_db} (из {len(dbs)})")
    return latest_db

def normalize_percent(val):
    if val is None:
        return 0.0
    try:
        if isinstance(val, str):
            s = val.replace("%", "").replace(",", ".").strip()
            num = float(s)
            return num / 100.0 if num > 1 else num
        else:
            num = float(val)
            return num / 100.0 if num > 1 else num
    except:
        return 0.0

def vlookup_approx(value, table_df: pd.DataFrame, ret_col_index=1):
    if table_df is None or table_df.empty:
        return None
    try:
        keys = pd.to_numeric(table_df.iloc[:, 0], errors='coerce')
        vals = table_df.iloc[:, ret_col_index]
        mask = ~keys.isna()
        if not mask.any():
            return None
        df = pd.DataFrame({"k": keys[mask].astype(float).values,
                           "v": vals[mask].values})
        df = df.sort_values("k")
        le = df[df["k"] <= float(value)]
        return le.iloc[-1]["v"] if not le.empty else None
    except:
        return None

def two_level_catalog_lookup(catalog_df: pd.DataFrame, category, product_type, return_col_letter, debug=False):
    if catalog_df is None or catalog_df.empty:
        if debug:
            print(f"        Таблица Catalog пуста")
        return None
    try:
        if debug:
            print(f"        Шаг 1: Фильтр по категории '{category}' в колонке B")
        if 'B' not in catalog_df.columns or 'C' not in catalog_df.columns:
            if debug:
                print(f"        Ошибка: Колонки B или C отсутствуют!")
                print(f"        Доступные: {list(catalog_df.columns)}")
            return None
        col_category = catalog_df['B']
        if debug:
            print(f"        Первые значения колонки B: {col_category.head(3).tolist()}")
        category_matches = catalog_df[col_category == category]
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
        col_product_type = category_matches['C']
        final_matches = category_matches[col_product_type == product_type]
        if final_matches.empty:
            mask = col_product_type.astype(str).str.strip().str.lower() == str(product_type).strip().lower()
            final_matches = category_matches[mask]
        if final_matches.empty:
            if debug:
                print(f"        Ошибка: Тип товара '{product_type}' не найден")
            return None
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

class OzonPriceFinder:
    def __init__(self, db_name: Optional[str] = None):
        if db_name is None:
            print("\nОпределение БД...")
            self.db = get_latest_db()
        else:
            self.db = db_name

        print("\nЗагрузка таблиц...")
        self.refs = RemoteRefTables(self.db)

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
        df = self.refs.new_tariffs
        if df is None or df.empty:
            return None

        addr = addr.strip().upper()
        letters = ''.join([c for c in addr if c.isalpha()])
        digits = ''.join([c for c in addr if c.isdigit()])

        if not letters or not digits:
            return None

        idx = 0
        for ch in letters:
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        col = idx - 1

        row = int(digits) - 1

        try:
            return df.iat[row, col]
        except:
            return None

    def calc_C46(self, price, category, product_type, debug=False):
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

        val = two_level_catalog_lookup(self.refs.catalog, category, product_type, col, debug=debug)

        if val is None:
            if debug:
                print(f"      C46: не найдено в БД")
            return ""

        try:
            result = float(val)
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
        if C34 == "":
            return ""

        try:
            p = float(price)
            vol = float(C19) if C19 != "" else 0
        except:
            return ""

        if p > 300:
            if vol > 0:
                if vol <= 1:
                    result = 46
                elif vol <= 2:
                    result = 56
                elif vol <= 3:
                    result = 66
                elif vol <= 190:
                    result = 66 + (math.ceil(vol) - 3) * 15
                else:
                    result = 2871

                if debug:
                    print(f"      C50: цена={p:.2f}, объём={vol:.3f}л -> {result} руб")
                return result
            else:
                return ""
        else:
            if self.refs.logistics is not None and not self.refs.logistics.empty:
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
        if C34 == "" or C34 <= 0:
            return ""

        try:
            E3 = self.get_new_tariff_cell('E3')
            G5 = self.get_new_tariff_cell('G5')
            G6 = self.get_new_tariff_cell('G6')

            if E3 is None or G5 is None or G6 is None:
                return ""

            e3 = float(E3)
            g5 = float(G5)
            g6 = float(G6)
            p = float(price)

            val = e3 * p

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
        if C34 == "" or category == "":
            return ""

        try:
            val = self.get_new_tariff_cell('C10')
            result = float(val) if (val is not None and str(val).strip() != "") else ""
            if debug:
                print(f"      C55: C10={val}")
            return result
        except:
            return ""

    def compute_fbo_for_price(self, price, row, debug=False):
        C9 = row.get("Категория")
        C11 = row.get("Тип товара")

        try:
            C16 = float(row.get("Длина, мм*", 0)) / 10.0
            C17 = float(row.get("Высота, мм*", 0)) / 10.0
            C18 = float(row.get("Ширина, мм*", 0)) / 10.0
        except:
            C16 = C17 = C18 = 0.0

        C19 = (C16 * C17 * C18) / 1000.0 if (C16 > 0 and C17 > 0 and C18 > 0) else 0.0

        try:
            C20 = float(row.get("Вес, г*", 0)) / 1000.0
        except:
            C20 = 0.0

        C21 = normalize_percent(row.get("% Выкупа"))

        if not C9 or not C11:
            return {
                "C33": "", "C34": "", "C46": "", "C47": "",
                "C48": "", "C50": "", "C51": "", "C52": "",
                "C44": "", "C45": "", "FBO": 0
            }

        C34 = float(price) if price > 0 else ""

        C46 = self.calc_C46(price, C9, C11, debug=debug)

        C47 = 0.015 if C34 != "" else ""

        C50 = self.calc_C50(price, C19, C34, debug=debug)

        C51 = self.calc_C51(price, C34, debug=debug)

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

        try:
            if C53 != "" and C53 > 0 and C21 > 0:
                C52 = float(C53) * (1 - C21)
            else:
                C52 = ""
        except:
            C52 = ""

        try:
            if C34 != "" and C34 > 0:
                numerator = 0
                commission_part = 0
                if C46 != "" and C47 != "":
                    commission_part = (float(C46) + float(C47)) * float(C34)
                    numerator += commission_part
                if C48 != "":
                    numerator += float(C48)
                if C52 != "":
                    numerator += float(C52)
                C45 = numerator / float(C34)
            else:
                C45 = ""
        except:
            C45 = ""

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
        себестоимость = float(row.get("Себестоимость") or 0.0)

        low = max(1.0, себестоимость * 1.2)
        high = max(low * 5.0, low + 5000.0)

        best = {"price": None, "fbo": None, "margin_diff": None,
                "margin_pct": None, "debug": None}

        tol = 0.0005
        max_iterations = 100

        for iteration in range(max_iterations):
            mid = (low + high) / 2.0

            show_detail = debug and (iteration == 0 or iteration == 10)

            pieces = self.compute_fbo_for_price(mid, row, debug=show_detail)
            fbo_mid = pieces["FBO"]
            profit = mid - себестоимость - fbo_mid
            margin = (profit / mid) if mid > 0 else -1.0
            diff = margin - target_margin

            if debug and iteration % 10 == 0:
                print(f"  Итер {iteration}: Цена={mid:.2f}, FBO={fbo_mid:.2f}, "
                      f"Рентаб={margin * 100:.2f}%, Разница={diff * 100:.3f}%")

            if (best["margin_diff"] is None) or (abs(diff) < abs(best["margin_diff"])):
                best["margin_diff"] = diff
                best["price"] = mid
                best["fbo"] = fbo_mid
                best["margin_pct"] = margin * 100.0
                best["debug"] = pieces

            if abs(diff) <= tol:
                break

            if margin < target_margin:
                low = mid
            else:
                high = mid

            if abs(high - low) < 0.01:
                break

        if best["price"] is not None:
            best["price"] = round(best["price"], 2)
            best["fbo"] = round(best["fbo"], 2)
            best["margin_pct"] = round(best["margin_pct"], 2)

        return best


def calculate_file(input_excel: str, target_margin_pct: float = 20.0,
                   output_excel: str = "calculated_result.xlsx", debug_first=False):
    print("=" * 70)
    print("КАЛЬКУЛЯТОР ЦЕН OZON FBO")
    print("=" * 70)

    input_excel = os.path.abspath(input_excel)
    output_excel = os.path.abspath(output_excel)

    print(f"\nВходной файл: {input_excel}")
    print(f"Файл существует: {os.path.exists(input_excel)}")

    if not os.path.exists(input_excel):
        raise FileNotFoundError(f"Файл не найден: {input_excel}")

    print(f"Размер: {os.path.getsize(input_excel) / 1024:.2f} KB")

    try:
        print(f"\nЧтение Excel...")
        df = pd.read_excel(input_excel, engine='openpyxl')
        print(f"OK Прочитано {len(df)} товаров")
    except Exception as e:
        print(f"ОШИБКА: {e}")
        raise

    print(f"\nКолонки: {list(df.columns)}")

    pf = OzonPriceFinder()

    result_df = df.copy()

    result_df['Итоговая цена (руб)'] = None
    result_df['FBO расходы (руб)'] = None
    result_df['Рентабельность (%)'] = None

    target_margin = float(target_margin_pct) / 100.0

    print(f"\nЦелевая рентабильность: {target_margin_pct}%")
    print(f"Товаров: {len(df)}\n")

    for idx, row in df.iterrows():
        name = row.get("Название")
        if pd.isna(name) or str(name).strip() == "":
            continue

        try:
            print(f"[{idx + 1}/{len(df)}]: {name[:40]}...", end=" ")
            show_debug = debug_first and idx == 0
            if show_debug:
                print("\n  ОТЛАДКА:")

            best = pf.find_price_for_target_margin(row, target_margin, debug=show_debug)

            if show_debug:
                print("  КОНЕЦ ОТЛАДКИ\n")

            print(f"OK {best['price']} руб, FBO: {best['fbo']} руб, {best['margin_pct']}%")

            result_df.at[idx, 'Итоговая цена (руб)'] = best["price"]
            result_df.at[idx, 'FBO расходы (руб)'] = best["fbo"]
            result_df.at[idx, 'Рентабельность (%)'] = best["margin_pct"]

        except Exception as e:
            print(f"ОШИБКА {e}")
            result_df.at[idx, 'Итоговая цена (руб)'] = None
            result_df.at[idx, 'FBO расходы (руб)'] = None
            result_df.at[idx, 'Рентабельность (%)'] = None

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


if __name__ == '__main__':
    path = input("Excel файл: ").strip()
    margin = float(input("Рентабильность (%): ").strip())
    debug = input("Отладка? (y/n): ").strip().lower() == 'y'
    calculate_file(path, margin, debug_first=debug)
