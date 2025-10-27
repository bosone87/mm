#!/usr/bin/env python3

"""
Простой парсер Excel для генерации inventory.yml
"""

import sys
import yaml

def main():
    try:
        # Проверяем аргументы командной строки
        if len(sys.argv) < 2:
            print("Использование: python3 excel_to_inventory.py <excel-file> [output-file]")
            print("Пример: python3 excel_to_inventory.py hosts.xlsx inventory.yml")
            sys.exit(1)
        
        excel_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else 'inventory.yml'
        
        # Пытаемся импортировать openpyxl
        try:
            from openpyxl import load_workbook
        except ImportError:
            print("Ошибка: openpyxl не установлен. Выполните: pip3 install openpyxl")
            sys.exit(1)
        
        # Загружаем Excel файл
        print(f"Чтение файла: {excel_file}")
        workbook = load_workbook(excel_file)
        sheet = workbook.active
        
        # Получаем заголовки
        headers = []
        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col).value
            if cell_value:
                headers.append(str(cell_value).strip())
        
        print(f"Найдены столбцы: {headers}")
        
        # Проверяем обязательные столбцы
        required = ['Name', 'HostName', 'Ip-address']
        for col in required:
            if col not in headers:
                print(f"Ошибка: отсутствует столбец '{col}'")
                sys.exit(1)
        
        # Определяем индексы столбцов
        name_idx = headers.index('Name') + 1
        hostname_idx = headers.index('HostName') + 1
        ip_idx = headers.index('Ip-address') + 1
        
        # Создаем структуру inventory
        inventory = {'all': {'hosts': {}}}
        host_count = 0
        
        # Обрабатываем строки
        for row in range(2, sheet.max_row + 1):
            name = sheet.cell(row=row, column=name_idx).value
            if not name:
                continue
                
            hostname = sheet.cell(row=row, column=hostname_idx).value
            ip_address = sheet.cell(row=row, column=ip_idx).value
            
            if not ip_address:
                continue
            
            # Обрабатываем IP-адрес (убираем маску если есть)
            ip_str = str(ip_address).strip()
            if '/' in ip_str:
                ip_without_mask = ip_str.split('/')[0].strip()
            else:
                ip_without_mask = ip_str
            
            # Создаем запись хоста
            host_vars = {
                'ansible_host': ip_without_mask,
                'ansible_user': 'ubuntu'  # Измените на нужного пользователя
            }
            
            if hostname:
                host_vars['hostname'] = str(hostname).strip()
            
            inventory['all']['hosts'][str(name).strip()] = host_vars
            host_count += 1
            print(f"Добавлен хост: {name} -> {ip_without_mask}")
        
        # Сохраняем в YAML
        with open(output_file, 'w', encoding='utf-8') as f:
            yaml.dump(inventory, f, default_flow_style=False, allow_unicode=True)
        
        print(f"Готово! Создан файл: {output_file}")
        print(f"Обработано хостов: {host_count}")
        
    except Exception as e:
        print(f"Ошибка: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()