import yaml

def simple_excel_parser(excel_file: str, output_file: str = 'inventory.yml'):
    """
    Упрощенная версия с минимальными зависимостями
    """
    try:
        # Если pandas не работает, используем встроенные средства
        import csv
        
        # Конвертируем Excel в CSV сначала (вручную)
        # или используем openpyxl напрямую
        
        import openpyxl
        
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active
        
        inventory = {'all': {'hosts': {}}}
        
        # Пропускаем заголовок (первую строку)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row[0]:  # Пустое имя
                continue
                
            host_name = str(row[0])
            hostname = str(row[1]) if row[1] else ""
            ip_address = str(row[2])
            
            # Простая обработка IP (без валидации)
            ip_parts = ip_address.split('/')
            ip_without_mask = ip_parts[0]
            
            host_vars = {
                'ansible_host': ip_without_mask,
                'ansible_user': 'admin'
            }
            
            if hostname:
                host_vars['hostname'] = hostname
                
            inventory['all']['hosts'][host_name] = host_vars
        
        with open(output_file, 'w', encoding='utf-8') as f:
            yaml.dump(inventory, f, default_flow_style=False, allow_unicode=True)
        
        print(f"✅ Inventory создан: {output_file}")
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")

# simple_excel_parser('hosts.xlsx')