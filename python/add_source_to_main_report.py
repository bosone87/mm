import pandas as pd
import hashlib
import re
import os
import glob
from datetime import datetime

# ========== НАСТРОЙКИ ==========
MAIN_FILE = "comparison_result.xlsx"      # Основной файл (с листом "Новый с комментариями")
REPORTS_FOLDER = "."                      # Папка с дополнительными отчётами (старыми)
OUTPUT_FILE = MAIN_FILE                   # Будем добавлять лист в тот же файл (или можно указать новый)

# Регулярка для даты в имени файла (поддерживает дефис и точку)
DATE_PATTERN = r"(\d{4}[.-]\d{2}[.-]\d{2})"

# Названия колонок (должны совпадать с теми, что в основном файле)
COLUMN_MAPPING = {
    "vuln_name": "Наименование уязвимости",
    "ip": "ip-адрес хоста",
    "ports": "список портов",
    "comment": "комментарий",
    "pack": "пачка"
}
# ================================

def extract_date_from_filename(filepath):
    """Извлекает дату из имени файла, если нет — берёт дату модификации."""
    name = os.path.basename(filepath)
    match = re.search(DATE_PATTERN, name)
    if match:
        date_str = match.group(1)
        for fmt in ("%Y-%m-%d", "%Y.%m.%d"):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
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

def load_report_with_ids(filepath, col_map):
    """Загружает Excel, добавляет колонку _id и _original_index (номер строки)."""
    df = pd.read_excel(filepath, sheet_name=0, dtype=str)
    df = df.fillna("")
    # Убедимся, что нужные колонки есть (если нет, создаём пустые)
    for key, col_name in col_map.items():
        if col_name not in df.columns:
            df[col_name] = ""
    df['_id'] = df.apply(lambda row: get_vuln_id(row, col_map), axis=1)
    df['_original_index'] = range(1, len(df) + 1)
    return df

def main():
    print("=== Добавление источников комментариев из дополнительных отчётов ===\n")
    
    # 1. Загружаем основной файл (лист "Новый с комментариями")
    if not os.path.exists(MAIN_FILE):
        print(f"Ошибка: файл {MAIN_FILE} не найден")
        return
    df_main = pd.read_excel(MAIN_FILE, sheet_name="Новый с комментариями", dtype=str)
    df_main = df_main.fillna("")
    # Добавляем ID в основной DataFrame
    df_main['_id'] = df_main.apply(lambda row: get_vuln_id(row, COLUMN_MAPPING), axis=1)
    # Запомним исходные индексы (порядок строк)
    df_main['_main_index'] = range(1, len(df_main) + 1)
    
    print(f"Основной файл: {MAIN_FILE}")
    print(f"Строк на листе 'Новый с комментариями': {len(df_main)}")
    
    # 2. Находим все Excel-файлы в папке (кроме основного)
    pattern = os.path.join(REPORTS_FOLDER, "*.xlsx")
    all_files = glob.glob(pattern)
    # Исключаем основной файл (по полному пути)
    main_full = os.path.abspath(MAIN_FILE)
    other_files = [f for f in all_files if os.path.abspath(f) != main_full]
    
    if not other_files:
        print("Не найдено дополнительных файлов отчётов.")
        return
    
    print(f"Найдено дополнительных файлов: {len(other_files)}")
    
    # 3. Загружаем все дополнительные файлы с их ID, датами, номерами строк
    sources = []  # список словарей: id, filepath, row_number, date
    for fpath in other_files:
        try:
            df = load_report_with_ids(fpath, COLUMN_MAPPING)
            date = extract_date_from_filename(fpath)
            for idx, row in df.iterrows():
                vid = row['_id']
                sources.append({
                    'id': vid,
                    'filepath': fpath,
                    'filename': os.path.basename(fpath),
                    'row_number': row['_original_index'],
                    'date': date,
                    'date_str': date.strftime("%Y-%m-%d")
                })
            print(f"  Загружен: {os.path.basename(fpath)} ({len(df)} записей)")
        except Exception as e:
            print(f"  Ошибка при загрузке {fpath}: {e}")
    
    if not sources:
        print("Не удалось загрузить данные из дополнительных файлов.")
        return
    
    # 4. Для каждой строки основного файла собираем все совпадения из sources
    # Создаём список для нового листа
    rows_for_new_sheet = []
    for idx, main_row in df_main.iterrows():
        main_id = main_row['_id']
        main_index = main_row['_main_index']
        # Ищем все источники с таким же id
        matches = [s for s in sources if s['id'] == main_id]
        if matches:
            for m in matches:
                rows_for_new_sheet.append({
                    '№ строки в основном файле': main_index,
                    'ID уязвимости': main_id,
                    'IP': main_row.get(COLUMN_MAPPING['ip'], ''),
                    'Наименование уязвимости': main_row.get(COLUMN_MAPPING['vuln_name'], ''),
                    'Порты': main_row.get(COLUMN_MAPPING['ports'], ''),
                    'Комментарий (из основного)': main_row.get(COLUMN_MAPPING['comment'], ''),
                    'Пачка (из основного)': main_row.get(COLUMN_MAPPING['pack'], ''),
                    'Имя файла-источника': m['filename'],
                    'Ссылка на файл': m['filepath'],
                    'Номер строки в файле': m['row_number'],
                    'Дата отправки (из имени файла)': m['date_str']
                })
        else:
            # Если не найдено ни одного источника, тоже добавим строку с пустыми полями
            rows_for_new_sheet.append({
                '№ строки в основном файле': main_index,
                'ID уязвимости': main_id,
                'IP': main_row.get(COLUMN_MAPPING['ip'], ''),
                'Наименование уязвимости': main_row.get(COLUMN_MAPPING['vuln_name'], ''),
                'Порты': main_row.get(COLUMN_MAPPING['ports'], ''),
                'Комментарий (из основного)': main_row.get(COLUMN_MAPPING['comment'], ''),
                'Пачка (из основного)': main_row.get(COLUMN_MAPPING['pack'], ''),
                'Имя файла-источника': '',
                'Ссылка на файл': '',
                'Номер строки в файле': '',
                'Дата отправки (из имени файла)': ''
            })
    
    df_sources = pd.DataFrame(rows_for_new_sheet)
    
    # 5. Добавляем новый лист в основной Excel-файл
    with pd.ExcelWriter(MAIN_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_sources.to_excel(writer, sheet_name='Источники комментариев', index=False)
    
    print(f"\n✅ В файл {MAIN_FILE} добавлен лист 'Источники комментариев'")
    print(f"   Всего записей на листе: {len(df_sources)}")
    print(f"   Из них с найденными источниками: {len(df_sources[df_sources['Имя файла-источника'] != ''])}")
    
    # Дополнительно: если хотите сохранить копию с новым именем, раскомментируйте:
    # backup = MAIN_FILE.replace('.xlsx', '_with_sources.xlsx')
    # df_sources.to_excel(backup, sheet_name='Источники комментариев', index=False)
    # print(f"   Также сохранена копия: {backup}")

if __name__ == "__main__":
    main()