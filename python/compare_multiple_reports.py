import pandas as pd
import hashlib
import re
import os
import glob
from datetime import datetime
from pathlib import Path

# ========== НАСТРОЙКИ ==========
INPUT_FOLDER = "."                     # Каталог с файлами отчётов
DATE_PATTERN = r"(\d{4}[.-]\d{2}[.-]\d{2})"  # извлечение даты
OUTPUT_FILE = "comparison_result.xlsx"

# Отображение названий колонок (обязательные и опциональные)
COLUMN_MAPPING = {
    "vuln_name": "Наименование уязвимости",
    "vuln_criticality": "Уровень критичности уязвимости",
    "host_criticality": "Уровень критичности хостов",
    "ip": "IP-адрес хоста",
    "hostname": "Имя хоста",          # может отсутствовать
    "os": "ОС",                        # может отсутствовать
    "hostname2": "Хостнейм",
    "ports": "Список портов",
    "description": "Описание уязвимости",
    "recommendation": "Рекомендации по устранению",
    "links": "Ссылки",
    "additional": "Дополнительно",
    "system": "Система",  # может отсутствовать
    "comment": "Комментарий",
    "pack": "Пачка"
}
# ================================

def extract_date_from_filename(filepath):
    """Извлекает дату из имени файла по шаблону. Возвращает datetime или None."""
    name = os.path.basename(filepath)
    match = re.search(DATE_PATTERN, name)
    if match:
        try:
            return datetime.strptime(match.group(1), "%Y-%m-%d")
        except:
            pass
    # Если не получилось, берём дату модификации файла
    return datetime.fromtimestamp(os.path.getmtime(filepath))

def get_vuln_id(row, col_map):
    """Уникальный ID на основе IP + названия + портов."""
    ip = str(row.get(col_map["ip"], ""))
    vuln = str(row.get(col_map["vuln_name"], ""))
    ports = str(row.get(col_map["ports"], ""))
    
    ports_normalized = re.sub(r'\s+', '', ports.lower())
    if ',' in ports_normalized:
        parts = [p for p in ports_normalized.split(',') if p.isdigit()]
        ports_normalized = ','.join(sorted(parts, key=int))
    
    unique_str = f"{ip}|{vuln}|{ports_normalized}"
    return hashlib.md5(unique_str.encode('utf-8')).hexdigest()

def load_excel_with_ids(filepath, col_map):
    """
    Загружает Excel, добавляет колонку _id.
    Отсутствующие колонки из col_map создаёт пустыми.
    """
    df = pd.read_excel(filepath, sheet_name=0, dtype=str)
    df = df.fillna("")
    # Проверяем наличие всех нужных колонок, добавляем пустые если нет
    for key, col_name in col_map.items():
        if col_name not in df.columns:
            df[col_name] = ""
    df['_id'] = df.apply(lambda row: get_vuln_id(row, col_map), axis=1)
    # Сохраним исходный индекс (номер строки в файле)
    df['_original_index'] = range(1, len(df) + 1)  # человеческий номер (1-based)
    return df

def find_most_recent_comment(target_row, source_files_data):
    """
    source_files_data: список кортежей (date, filepath, df)
    Ищет в source_files_data совпадение по _id с target_row.
    Возвращает кортеж (comment, pack, source_filepath, source_row_number, source_date)
    или (None, None, None, None, None) если не найдено.
    """
    target_id = target_row['_id']
    # Идём по убыванию даты (самые свежие сначала)
    for date, filepath, df in sorted(source_files_data, key=lambda x: x[0], reverse=True):
        # Ищем строку с таким же _id
        matches = df[df['_id'] == target_id]
        if not matches.empty:
            # Берём первую (должна быть одна)
            match_row = matches.iloc[0]
            comment = match_row[COLUMN_MAPPING['comment']]
            pack = match_row[COLUMN_MAPPING['pack']]
            # Если комментарий и пачка пустые, продолжаем поиск в более старых?
            # По условию: если есть заполненный комментарий, берём его.
            # Но если комментарий пуст, возможно, стоит искать дальше.
            # Считаем, что пустая строка - не комментарий, ищем дальше.
            if comment.strip() or pack.strip():
                source_row = match_row['_original_index']
                return comment, pack, filepath, source_row, date
    return None, None, None, None, None

def main():
    print("=== Сравнение нескольких отчётов Nessus ===\n")
    
    # Находим все Excel-файлы в папке
    pattern = os.path.join(INPUT_FOLDER, "*.xlsx")
    files = glob.glob(pattern)
    if not files:
        print(f"Ошибка: не найдено файлов .xlsx в папке {INPUT_FOLDER}")
        return
    
    # Сортируем файлы по дате (из имени или модификации)
    file_dates = [(extract_date_from_filename(f), f) for f in files]
    file_dates.sort(key=lambda x: x[0])  # по возрастанию даты
    
    # Самый новый файл - последний
    newest_date, newest_file = file_dates[-1]
    old_files_data = []
    
    print(f"Целевой (новый) файл: {os.path.basename(newest_file)} (дата: {newest_date.date()})")
    print("Старые файлы (источники комментариев):")
    for date, fpath in file_dates[:-1]:
        print(f"  {os.path.basename(fpath)} (дата: {date.date()})")
        # Загружаем каждый старый файл с ID
        df_old = load_excel_with_ids(fpath, COLUMN_MAPPING)
        old_files_data.append((date, fpath, df_old))
    
    # Загружаем целевой файл
    df_target = load_excel_with_ids(newest_file, COLUMN_MAPPING)
    print(f"\nЗаписей в целевом файле: {len(df_target)}")
    
    # Для каждой строки целевого файла ищем комментарий в старых
    new_comments = []
    new_packs = []
    source_files = []
    source_rows = []
    source_dates = []
    
    for idx, row in df_target.iterrows():
        comment, pack, src_file, src_row, src_date = find_most_recent_comment(row, old_files_data)
        new_comments.append(comment if comment is not None else "")
        new_packs.append(pack if pack is not None else "")
        source_files.append(src_file if src_file is not None else "")
        source_rows.append(src_row if src_row is not None else "")
        source_dates.append(src_date.strftime("%Y-%m-%d") if src_date is not None else "")
    
    # Обновляем колонки в целевой DataFrame
    df_target[COLUMN_MAPPING['comment']] = new_comments
    df_target[COLUMN_MAPPING['pack']] = new_packs
    df_target.insert(
        df_target.columns.get_loc(COLUMN_MAPPING['pack']) + 1,
        "Ссылка на файл",
        source_files
    )
    df_target.insert(
        df_target.columns.get_loc(COLUMN_MAPPING['pack']) + 2,
        "Номер строки в файле",
        source_rows
    )
    df_target.insert(
        df_target.columns.get_loc(COLUMN_MAPPING['pack']) + 3,
        "Дата отправки",
        source_dates
    )
    df_target.insert(
        df_target.columns.get_loc(COLUMN_MAPPING['pack']) + 4,
        "Имя файла",
        [os.path.basename(f) if f else "" for f in source_files]
    )
    
    # Удаляем служебные колонки
    df_target = df_target.drop(columns=['_id', '_original_index'])
    
    # Статистика
    total_target = len(df_target)
    found_comments = sum(1 for c in new_comments if c.strip())
    
    stats_data = {
        "Показатель": [
            "Целевой файл",
            "Дата целевого файла",
            "Количество обработанных старых файлов",
            "Количество записей в целевом файле",
            "Из них найдены комментарии/пачка в старых файлах",
            "Не найдено"
        ],
        "Значение": [
            os.path.basename(newest_file),
            newest_date.strftime("%Y-%m-%d"),
            len(old_files_data),
            total_target,
            found_comments,
            total_target - found_comments
        ]
    }
    df_stats = pd.DataFrame(stats_data)
    
    # Сохраняем результат
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df_stats.to_excel(writer, sheet_name="Статистика", index=False)
        df_target.to_excel(writer, sheet_name="Новый с комментариями", index=False)
    
    print(f"\n✅ Результат сохранён в файл: {OUTPUT_FILE}")
    print(f"   - Лист 'Статистика' – общая информация")
    print(f"   - Лист 'Новый с комментариями' – все строки целевого файла с добавленными колонками")

if __name__ == "__main__":
    main()