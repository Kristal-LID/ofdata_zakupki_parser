import requests
import configparser
import pandas as pd
import sys
import os
from datetime import datetime

def get_data(params):
    """Запрос к API, используя только заполненные параметры."""
    url = "https://api.ofdata.ru/v2/contracts"
    
    # Убираем параметры, которые пользователь оставил пустыми в .ini
    clean_params = {k: v for k, v in params.items() if v}
    
    try:
        response = requests.get(url, params=clean_params, timeout=15)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"❌ Ошибка при запросе: {e}")
        return None

def save_to_excel(json_data):
    """Преобразует JSON в Excel."""
    try:
        if not json_data or 'data' not in json_data or not json_data['data'].get('Записи'):
            print("⚠️ API не вернуло записей. Возможно, по этим фильтрам ничего не найдено.")
            return False

        records = json_data['data']['Записи']
        rows = []

        for item in records:
            # Обработка объектов (защита от пустых значений)
            objs = [o.get('Наим') for o in item.get('Объекты', []) if o.get('Наим')]
            objs_str = ", ".join(objs) if objs else "Не указано"
            
            # Обработка поставщиков
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
        
        # Сохраняем файл в ту же папку, где лежит скрипт
        folder = os.path.dirname(os.path.abspath(__file__))
        name = f"Результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = os.path.join(folder, name)
        
        df.to_excel(path, index=False)
        print(f"✅ Готово! Файл создан: {name}")
        return True
    except Exception as e:
        print(f"❌ Ошибка Excel: {e}")
        return False

if __name__ == "__main__":
    # Настройка путей
    base_dir = os.path.dirname(os.path.abspath(__file__))
    ini_path = os.path.join(base_dir, 'настройки.ini')
    
    config = configparser.ConfigParser()
    
    if not os.path.exists(ini_path):
        print(f"❌ Файл {ini_path} не найден!")
    else:
        try:
            config.read(ini_path, encoding='utf-8')
            if 'НАСТРОЙКИ' not in config:
                print("❌ В файле нет секции [НАСТРОЙКИ]")
            else:
                p = config['НАСТРОЙКИ']
                
                # Собираем все параметры
                api_params = {
                    "key": p.get('API_KEY'),
                    "ogrn": p.get('OGRN'),
                    "inn": p.get('INN'),
                    "kpp": p.get('KPP'),
                    "law": p.get('LAW'),
                    "role": p.get('ROLE'),
                    "limit": p.get('LIMIT'),
                    "page": p.get('PAGE'),
                    "sort": p.get('SORT'),

                }

                print(f"🚀 Запуск... Проверяем ИНН {api_params['inn']}...")
                
                data = get_data(api_params)
                
                if data and data.get('meta', {}).get('status') == 'ok':
                    save_to_excel(data)
                    print(f"💰 Остаток на балансе: {data.get('meta', {}).get('balance')} руб.")
                else:
                    msg = data.get('meta', {}).get('message') if data else "Нет ответа от API"
                    print(f"❌ Ошибка API: {msg}")

        except Exception as e:
            print(f"❌ Критическая ошибка: {e}")

    input("\nНажмите Enter, чтобы выйти...")