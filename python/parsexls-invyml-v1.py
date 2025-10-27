import openpyxl
import yaml
import ipaddress

def parse_excel_without_pandas(excel_file: str, output_file: str = 'inventory.yml') -> None:
    """
    Версия без использования pandas
    """
    try:
        # Загрузка Excel файла
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active
        
        inventory = {'all': {'hosts': {}}}
        
        # Получаем заголовки
        headers = [cell.value for cell in sheet[1]]
        
        # Проверяем необходимые столбцы
        required_columns = ['Name', 'HostName', 'Ip-address']
        for col in required_columns:
            if col not in headers:
                raise ValueError(f"Отсутствует столбец: {col}")
        
        # Индексы столбцов
        name_idx = headers.index('Name')
        hostname_idx = headers.index('HostName')
        ip_idx = headers.index('Ip-address')
        
        # Обработка строк
        for row_num, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not row[name_idx]:  # Пустое имя
                continue
                
            host_name = str(row[name_idx])
            hostname = str(row[hostname_idx]) if row[hostname_idx] else ""
            ip_address = str(row[ip_idx]) if row[ip_idx] else ""
            
            if not ip_address:
                print(f"Пропущена строка {row_num}: отсутствует IP-адрес")
                continue
            
            # Обработка IP
            try:
                if '/' in ip_address:
                    ip_without_mask = ip_address.split('/')[0]
                else:
                    ip_without_mask = ip_address
                
                ipaddress.ip_address(ip_without_mask)
                
            except ValueError:
                print(f"Пропущена строка {row_num}: некорректный IP-адрес '{ip_address}'")
                continue
            
            # Создание записи
            host_vars = {
                'ansible_host': ip_without_mask,
                'ansible_user': 'admin'
            }
            
            if hostname:
                host_vars['hostname'] = hostname
            
            inventory['all']['hosts'][host_name] = host_vars
        
        # Запись в YAML
        with open(output_file, 'w', encoding='utf-8') as f:
            yaml.dump(inventory, f, default_flow_style=False, allow_unicode=True)
        
        print(f"✅ Inventory файл успешно создан: {output_file}")
        print(f"📊 Обработано хостов: {len(inventory['all']['hosts'])}")
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")

# parse_excel_without_pandas('hosts.xlsx')