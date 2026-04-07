import pandas as pd
import hashlib
import re
from datetime import datetime

# ========== НАСТРОЙКИ ==========
FILE_OLD = "old_report.xlsx"
FILE_NEW = "new_report.xlsx"
OUTPUT_FILE = "comparison_result.xlsx"

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
    df = pd.read_excel(file_path, sheet_name=0, dtype=str)
    df = df.fillna("")
    df['_id'] = df.apply(get_vuln_id, axis=1)
    return df

def main():
    print("=== Сравнение двух отчётов Nessus с подстановкой комментариев ===\n")
    
    df_old = load_excel_with_ids(FILE_OLD)
    df_new = load_excel_with_ids(FILE_NEW)
    
    print(f"Старый файл: {len(df_old)} записей")
    print(f"Новый файл: {len(df_new)} записей")
    
    ids_old = set(df_old['_id'])
    ids_new = set(df_new['_id'])
    
    common_ids = ids_old.intersection(ids_new)
    unique_ids = ids_new.difference(ids_old)
    
    # 1. Одинаковые записи (из старого)
    df_common = df_old[df_old['_id'].isin(common_ids)].drop(columns=['_id'])
    
    # 2. Уникальные записи (из нового)
    df_unique = df_new[df_new['_id'].isin(unique_ids)].drop(columns=['_id'])
    
    # 3. Новый файл с подстановкой комментариев и пачки из старого
    # Создаём словарь: id -> (comment, pack) из старого файла
    comment_pack_dict = {}
    for vid in common_ids:
        old_row = df_old[df_old['_id'] == vid].iloc[0]
        comment_pack_dict[vid] = (
            old_row[COLUMN_MAPPING['comment']],
            old_row[COLUMN_MAPPING['pack']]
        )
    
    # Копируем новый файл и подставляем значения
    df_new_with_comments = df_new.copy()
    # Для строк с совпадающим id заменяем comment и pack
    for vid, (comm, pack_val) in comment_pack_dict.items():
        mask = df_new_with_comments['_id'] == vid
        df_new_with_comments.loc[mask, COLUMN_MAPPING['comment']] = comm
        df_new_with_comments.loc[mask, COLUMN_MAPPING['pack']] = pack_val
    
    df_new_with_comments = df_new_with_comments.drop(columns=['_id'])
    
    # Статистика
    stats_data = {
        "Показатель": [
            "Количество записей в старом файле",
            "Количество записей в новом файле",
            "Количество одинаковых записей",
            "Количество уникальных записей в новом файле"
        ],
        "Значение": [
            len(df_old),
            len(df_new),
            len(common_ids),
            len(unique_ids)
        ]
    }
    df_stats = pd.DataFrame(stats_data)
    
    # Сохраняем
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df_stats.to_excel(writer, sheet_name="Статистика", index=False)
        df_common.to_excel(writer, sheet_name="Одинаковые (из старого)", index=False)
        df_unique.to_excel(writer, sheet_name="Уникальные (из нового)", index=False)
        df_new_with_comments.to_excel(writer, sheet_name="Новый с комментариями", index=False)
    
    print(f"\n✅ Результат сохранён в {OUTPUT_FILE}")
    print(f"   - Лист 'Новый с комментариями' содержит все строки нового файла, для совпавших уязвимостей проставлены 'Комментарий' и 'Пачка' из старого")

if __name__ == "__main__":
    main()