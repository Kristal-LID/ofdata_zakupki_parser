import requests
import configparser
import pandas as pd
import sys
import os
from datetime import datetime

# ---------- API ----------
def get_data(params):
    url = "https://api.ofdata.ru/v2/contracts"
    clean_params = {k: v for k, v in params.items() if v}
    try:
        response = requests.get(url, params=clean_params, timeout=15)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"❌ Ошибка при запросе: {e}")
        return None

def save_to_excel(json_data):
    try:
        if not json_data or 'data' not in json_data or not json_data['data'].get('Записи'):
            print("⚠️ API не вернуло записей. Возможно, по этим фильтрам ничего не найдено.")
            return False

        records = json_data['data']['Записи']
        rows = []

        for item in records:
            objs = [o.get('Наим') for o in item.get('Объекты', []) if o.get('Наим')]
            objs_str = ", ".join(objs) if objs else "Не указано"

            sups = []
            for p in item.get('Постав', []):
                name = p.get('НаимСокр') or p.get('ФИО') or "Поставщик"
                inn = p.get('ИНН') or "нет ИНН"
                sups.append(f"{name} (ИНН: {inn})")
            sups_str = "; ".join(sups) if sups else "Не указано"

            rows.append({
                'Рег. Номер': item.get('РегНомер') or "Б/Н",
                'Дата': item.get('Дата') or "-",
                'Срок исп.': item.get('ДатаИсп') or "-",
                'Цена (руб)': item.get('Цена') or 0,
                'Заказчик': item.get('Заказ', {}).get('НаимСокр') or "Не указано",
                'Поставщики': sups_str,
                'Объекты': objs_str,
                'Ссылка ЕИС': item.get('СтрЕИС') or ""
            })

        df = pd.DataFrame(rows)
        folder = os.path.dirname(os.path.abspath(__file__))
        name = f"Результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = os.path.join(folder, name)
        df.to_excel(path, index=False)
        print(f"✅ Готово! Файл создан: {name}")
        return True
    except Exception as e:
        print(f"❌ Ошибка Excel: {e}")
        return False

# ---------- Псевдографика / меню ----------
def clear():
    os.system('cls' if os.name == 'nt' else 'clear')

def line(char='─', n=64):
    print(char * n)

def header(title):
    clear()
    line('═')
    print(f"  {title}")
    line('═')

def ask(prompt, default=None, required=False):
    suffix = f" [{default}]" if default else ""
    while True:
        val = input(f"{prompt}{suffix}: ").strip()
        if not val and default is not None:
            return default
        if not val and required:
            print("  ⚠️  Поле обязательно для заполнения.")
            continue
        return val

def choose(prompt, options, default=None):
    """options: список (ключ, описание)"""
    print(f"\n{prompt}")
    for i, (k, desc) in enumerate(options, 1):
        mark = " *" if default == k else ""
        print(f"  {i}) {desc}{mark}")
    while True:
        val = input(f"Выбор [{default}]: ").strip()
        if not val and default is not None:
            return default
        if val.isdigit() and 1 <= int(val) <= len(options):
            return options[int(val) - 1][0]
        # разрешаем вводить ключ напрямую
        keys = [k for k, _ in options]
        if val in keys:
            return val
        print("  ⚠️  Неверный выбор, попробуйте снова.")

def confirm(prompt):
    return input(f"{prompt} (y/n): ").strip().lower() in ('y', 'yes', 'д', 'да')

def interactive_menu():
    header("🔎 Парсер закупок по API ofdata.ru")
    # --- АПИ ключ ---
    print("\nШаг 0. Ваш API-ключ от ofdata.ru")
    print("  (получить можно на https://ofdata.ru/)")
    print(" *ВАЖНО! если вы НЕ хотите каждый раз вводить ключ, удалите строку 113 в файле script.py. после впишите свой ключ в ковычки вместо API_key на строке 158*")
    API_key = ask(" API-ключ", default="")

    # --- Идентификатор компании ---
    print("\nШаг 1. Идентификатор организации")
    print("  Можно указать ОГРН или ИНН (одно из двух обязательно).")
    ogrn = ask("  ОГРН", default="")
    inn  = ask("  ИНН",  default="")
    if not ogrn and not inn:
        print("  ⚠️  Нужно указать хотя бы ОГРН или ИНН!")
        return None
    kpp = ask("  КПП (необязательно, только вместе с ИНН)", default="")

    # --- ФЗ ---
    law = choose("Шаг 2. По какому ФЗ искать?",
                 [("44",  "44-ФЗ"),
                  ("223", "223-ФЗ (с 2019 не публикуется)"),
                  ("94",  "94-ФЗ")],
                 default="44")

    # --- Роль ---
    role = choose("Шаг 3. Роль участника:",
                  [("customer", "Заказчик (customer)"),
                   ("supplier", "Поставщик (supplier)")],
                  default="supplier")

    # --- Лимит / страница ---
    print("\nШаг 4. Пагинация")
    print("  Лимит — от 1 до 100. Страница — номер страницы (1, 2, 3...).")
    while True:
        limit = ask("  LIMIT", default="100")
        if limit.isdigit() and 1 <= int(limit) <= 100:
            break
        print("  ⚠️  LIMIT должен быть числом от 1 до 100.")
    page = ask("  PAGE (оставьте пустым для первой страницы)", default="")

    # --- Сортировка ---
    sort = choose("Шаг 5. Сортировка результатов:",
                  [("-date", "По дате подписания, сначала свежие (-date)"),
                   ("date",  "По дате подписания, сначала старые (date)"),
                   ("-price","По цене, сначала дорогие (-price)"),
                   ("price", "По цене, сначала дешёвые (price)"),
                   ("",      "Без сортировки")],
                  default="-date")

    api_params = {
        "key":  "API_key",   # можно вынести самостоятельно и не вводить каждый раз
        "ogrn": ogrn,
        "inn":  inn,
        "kpp":  kpp,
        "law":  law,
        "role": role,
        "limit":limit,
        "page": page,
        "sort": sort,
    }

    # --- Итог ---
    header("📋 Проверьте параметры")
    for k, v in api_params.items():
        if k == "key":
            v = "***" + v[-4:]
        print(f"  {k:6} = {v or '—'}")
    line()
    if not confirm("Запустить запрос?"):
        print("Отменено.")
        return None
    return api_params

# ---------- Запуск из .ini (как было) ----------
def load_from_ini():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    ini_path = os.path.join(base_dir, 'settings.ini')
    if not os.path.exists(ini_path):
        print(f"❌ Файл {ini_path} не найден!")
        return None
    config = configparser.ConfigParser()
    config.read(ini_path, encoding='utf-8')
    if 'НАСТРОЙКИ' not in config:
        print("❌ В файле нет секции [НАСТРОЙКИ]")
        return None
    p = config['НАСТРОЙКИ']
    return {
        "key":   p.get('API_KEY'),
        "ogrn":  p.get('OGRN'),
        "inn":   p.get('INN'),
        "kpp":   p.get('KPP'),
        "law":   p.get('LAW'),
        "role":  p.get('ROLE'),
        "limit": p.get('LIMIT'),
        "page":  p.get('PAGE'),
        "sort":  p.get('SORT'),
    }

# ---------- main ----------
if __name__ == "__main__":
    use_ini = '--ini' in sys.argv

    if use_ini:
        api_params = load_from_ini()
    else:
        api_params = interactive_menu()

    if api_params:
        print(f"\n🚀 Запуск... ИНН={api_params.get('inn') or '—'}, ОГРН={api_params.get('ogrn') or '—'}")
        data = get_data(api_params)
        if data and data.get('meta', {}).get('status') == 'ok':
            save_to_excel(data)
            print(f"💰 Остаток на балансе: {data.get('meta', {}).get('balance')} руб.")
        else:
            msg = data.get('meta', {}).get('message') if data else "Нет ответа от API"
            print(f"❌ Ошибка API: {msg}")

    input("\nНажмите Enter, чтобы выйти...")