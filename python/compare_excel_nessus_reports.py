import pandas as pd
import hashlib
import re
from datetime import datetime

# ========== НАСТРОЙКИ ==========
FILE_OLD = "old_report.xlsx"      # Первый (старый) файл
FILE_NEW = "new_report.xlsx"      # Второй (новый) файл
OUTPUT_FILE = "comparison_result.xlsx"

# Точные названия колонок (при необходимости отредактируйте)
COLUMN_MAPPING = {
    "vuln_name": "Наименование уязвимости",
    "vuln_criticality": "уровень критичности уязвимости",
    "host_criticality": "уровень критичности хостов",
    "ip": "ip-адрес хоста",
    "hostname": "имя хоста",
    "os": "ОС",
    "hostname2": "хостнейм",
    "ports": "список портов",
    "description": "описание уязвимости",
    "recommendation": "рекомендации по устранению",
    "links": "ссылки",
    "additional": "дополнительно",
    "system": "система(к какой ИС относится)",
    "comment": "комментарий",
    "pack": "пачка"
}
# ================================

def get_vuln_id(row):
    """Формирует уникальный идентификатор уязвимости на основе IP + названия + портов"""
    ip = str(row.get(COLUMN_MAPPING["ip"], ""))
    vuln = str(row.get(COLUMN_MAPPING["vuln_name"], ""))
    ports = str(row.get(COLUMN_MAPPING["ports"], ""))
    
    # Нормализация портов: удаляем пробелы, сортируем числа
    ports_normalized = re.sub(r'\s+', '', ports.lower())
    if ',' in ports_normalized:
        parts = [p for p in ports_normalized.split(',') if p.isdigit()]
        ports_normalized = ','.join(sorted(parts, key=int))
    
    unique_str = f"{ip}|{vuln}|{ports_normalized}"
    return hashlib.md5(unique_str.encode('utf-8')).hexdigest()

def load_excel_with_ids(file_path):
    """Загружает Excel, добавляет колонку с ID и возвращает DataFrame и множество ID"""
    df = pd.read_excel(file_path, sheet_name=0, dtype=str)
    df = df.fillna("")
    # Проверяем наличие всех необходимых колонок
    for col in COLUMN_MAPPING.values():
        if col not in df.columns:
            print(f"Внимание: колонка '{col}' не найдена в файле {file_path}")
    # Добавляем ID
    df['_id'] = df.apply(get_vuln_id, axis=1)
    return df, set(df['_id'])

def main():
    print("=== Сравнение двух отчётов Nessus ===\n")
    
    # Загружаем файлы
    print(f"Чтение старого файла: {FILE_OLD}")
    df_old, ids_old = load_excel_with_ids(FILE_OLD)
    print(f"  Записей: {len(df_old)}")
    
    print(f"Чтение нового файла: {FILE_NEW}")
    df_new, ids_new = load_excel_with_ids(FILE_NEW)
    print(f"  Записей: {len(df_new)}")
    
    # Находим пересечение и разность
    common_ids = ids_old.intersection(ids_new)      # ID, которые есть в обоих файлах
    unique_ids = ids_new.difference(ids_old)        # ID, которые есть только в новом файле
    
    # Формируем DataFrame для одинаковых записей (из старого файла)
    df_common = df_old[df_old['_id'].isin(common_ids)].copy()
    df_common = df_common.drop(columns=['_id'])
    
    # Формируем DataFrame для уникальных записей (из нового файла)
    df_unique = df_new[df_new['_id'].isin(unique_ids)].copy()
    df_unique = df_unique.drop(columns=['_id'])
    
    # Статистика
    stats_data = {
        "Показатель": [
            "Количество записей в старом файле",
            "Количество записей в новом файле",
            "Количество одинаковых записей (присутствуют в обоих файлах)",
            "Количество уникальных записей в новом файле"
        ],
        "Значение": [
            len(df_old),
            len(df_new),
            len(df_common),
            len(df_unique)
        ]
    }
    df_stats = pd.DataFrame(stats_data)
    
    # Сохраняем в Excel с тремя листами
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df_stats.to_excel(writer, sheet_name="Статистика", index=False)
        df_common.to_excel(writer, sheet_name="Одинаковые (из старого)", index=False)
        df_unique.to_excel(writer, sheet_name="Уникальные (из нового)", index=False)
    
    print(f"\n✅ Результат сохранён в файл: {OUTPUT_FILE}")
    print(f"   - Лист 'Статистика' – общие цифры")
    print(f"   - Лист 'Одинаковые (из старого)' – {len(df_common)} записей")
    print(f"   - Лист 'Уникальные (из нового)' – {len(df_unique)} записей")

if __name__ == "__main__":
    main()