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
    "vuln_criticality": "Уовень критичности уязвимости",
    "host_criticality": "Уровень критичности хостов",
    "ip": "IP-адрес хоста",
    "hostname": "Имя хоста",
    "os": "ОС",
    "hostname2": "Хостнейм",
    "ports": "Список портов",
    "description": "Описание уязвимости",
    "recommendation": "Рекомендации по устранению",
    "links": "Ссылки",
    "additional": "Дополнительно",
    "system": "Система(к какой ИС относится)",
    # "comment": "Комментарий",
    "pack": "Пачка"
}
# ================================

def get_vuln_id(row):
    """Уникальный ID на основе IP + названия уязвимости + портов"""
    ip = str(row.get(COLUMN_MAPPING["ip"], ""))
    vuln = str(row.get(COLUMN_MAPPING["vuln_name"], ""))
    ports = str(row.get(COLUMN_MAPPING["ports"], ""))
    
    ports_normalized = re.sub(r'\s+', '', ports.lower())
    if ',' in ports_normalized:
        parts = [p for p in ports_normalized.split(',') if p.isdigit()]
        ports_normalized = ','.join(sorted(parts, key=int))
    
    unique_str = f"{ip}|{vuln}|{ports_normalized}"
    return hashlib.md5(unique_str.encode('utf-8')).hexdigest()

def load_excel_with_ids(file_path):
    """Загружает Excel, добавляет колонку _id, возвращает DataFrame и словарь id -> индекс"""
    df = pd.read_excel(file_path, sheet_name=0, dtype=str)
    df = df.fillna("")
    df['_id'] = df.apply(get_vuln_id, axis=1)
    # Словарь для быстрого поиска индекса по id (нужно для подстановки комментариев)
    id_to_idx = {row['_id']: idx for idx, row in df.iterrows()}
    return df, id_to_idx

def main():
    print("=== Сравнение двух отчётов Nessus с подстановкой комментариев ===\n")
    
    # Загружаем файлы
    print(f"Чтение старого файла: {FILE_OLD}")
    df_old, old_id_to_idx = load_excel_with_ids(FILE_OLD)
    print(f"  Записей: {len(df_old)}")
    
    print(f"Чтение нового файла: {FILE_NEW}")
    df_new, new_id_to_idx = load_excel_with_ids(FILE_NEW)
    print(f"  Записей: {len(df_new)}")
    
    # Множества ID
    ids_old = set(df_old['_id'])
    ids_new = set(df_new['_id'])
    
    common_ids = ids_old.intersection(ids_new)
    unique_ids = ids_new.difference(ids_old)
    
    # 1. Одинаковые записи (из старого файла)
    df_common = df_old[df_old['_id'].isin(common_ids)].copy()
    df_common = df_common.drop(columns=['_id'])
    
    # 2. Уникальные записи (из нового файла)
    df_unique = df_new[df_new['_id'].isin(unique_ids)].copy()
    df_unique = df_unique.drop(columns=['_id'])
    
    # 3. НОВЫЙ ЛИСТ: копия нового файла с подстановкой комментариев и пачки из старого
    df_new_with_comments = df_new.copy()
    # Для каждой строки в новом файле, если её id есть в common_ids, подставляем комментарий и пачку из старого
    for idx, row in df_new_with_comments.iterrows():
        vid = row['_id']
        if vid in common_ids:
            # Находим соответствующую строку в старом файле
            old_idx = old_id_to_idx[vid]
            old_row = df_old.loc[old_idx]
            df_new_with_comments.at[idx, COLUMN_MAPPING['comment']] = old_row[COLUMN_MAPPING['comment']]
            df_new_with_comments.at[idx, COLUMN_MAPPING['pack']] = old_row[COLUMN_MAPPING['pack']]
        # else: оставляем как есть (обычно пусто)
    df_new_with_comments = df_new_with_comments.drop(columns=['_id'])
    
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
    
    # Сохраняем в Excel
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df_stats.to_excel(writer, sheet_name="Статистика", index=False)
        df_common.to_excel(writer, sheet_name="Одинаковые (из старого)", index=False)
        df_unique.to_excel(writer, sheet_name="Уникальные (из нового)", index=False)
        df_new_with_comments.to_excel(writer, sheet_name="Новый с комментариями", index=False)
    
    print(f"\n✅ Результат сохранён в файл: {OUTPUT_FILE}")
    print(f"   - Лист 'Статистика' – общие цифры")
    print(f"   - Лист 'Одинаковые (из старого)' – {len(df_common)} записей")
    print(f"   - Лист 'Уникальные (из нового)' – {len(df_unique)} записей")
    print(f"   - Лист 'Новый с комментариями' – полная копия нового файла, но для совпавших строк заполнены 'Комментарий' и 'Пачка' из старого")

if __name__ == "__main__":
    main()