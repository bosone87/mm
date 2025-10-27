import pandas as pd
import yaml
import ipaddress
from typing import Dict, List, Any, Union

def parse_excel_to_inventory(excel_file: str, output_file: str = 'inventory.yml') -> None:
    """
    Парсит Excel файл и генерирует inventory.yml для Ansible
    """
    try:
        # Чтение Excel файла
        df = pd.read_excel(excel_file)
        
        # Проверка наличия необходимых столбцов
        required_columns = ['Name', 'HostName', 'Ip-address']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Отсутствуют необходимые столбцы: {', '.join(missing_columns)}")
        
        # Создание структуры inventory
        inventory = {'all': {'hosts': {}}}
        
        # Обработка каждой строки
        for index, row in df.iterrows():
            host_name = str(row['Name']).strip()
            hostname = str(row['HostName']).strip()
            ip_address = str(row['Ip-address']).strip()
            
            # Пропускаем пустые строки
            if not host_name or pd.isna(host_name):
                continue
            
            # Валидация IP-адреса с маской
            try:
                network = ipaddress.ip_network(ip_address, strict=False)
                ip_without_mask = str(network.network_address)
            except ValueError as e:
                print(f"Ошибка в строке {index + 1}: некорректный IP-адрес '{ip_address}' - {e}")
                continue
            
            # Создание записи для хоста
            host_vars = {
                'ansible_host': ip_without_mask,
                'ansible_user': 'admin'  # Замените на ваше имя пользователя
            }
            
            if hostname and not pd.isna(hostname):
                host_vars['hostname'] = hostname
            
            inventory['all']['hosts'][host_name] = host_vars
        
        # Запись в YAML файл
        with open(output_file, 'w', encoding='utf-8') as f:
            yaml.dump(inventory, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
        
        print(f"✅ Inventory файл успешно создан: {output_file}")
        print(f"📊 Обработано хостов: {len(inventory['all']['hosts'])}")
        
    except FileNotFoundError:
        print(f"❌ Ошибка: Файл '{excel_file}' не найден")
    except Exception as e:
        print(f"❌ Ошибка при обработке файла: {e}")

# Альтернативная версия без использования iterrows (более надежная)
def parse_excel_safe(excel_file: str, output_file: str = 'inventory.yml') -> None:
    """
    Альтернативная версия без использования iterrows
    """
    try:
        # Чтение Excel файла
        df = pd.read_excel(excel_file)
        
        # Проверка наличия необходимых столбцов
        required_columns = ['Name', 'HostName', 'Ip-address']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Отсутствует столбец: {col}")
        
        # Создание структуры inventory
        inventory = {'all': {'hosts': {}}}
        
        # Обработка данных
        for i in range(len(df)):
            try:
                host_name = str(df.iloc[i]['Name']).strip()
                if not host_name or host_name == 'nan':
                    continue
                    
                hostname = str(df.iloc[i]['HostName']).strip()
                ip_address = str(df.iloc[i]['Ip-address']).strip()
                
                # Валидация IP-адреса
                network = ipaddress.ip_network(ip_address, strict=False)
                ip_without_mask = str(network.network_address)
                
                # Создание записи для хоста
                host_vars = {
                    'ansible_host': ip_without_mask,
                    'ansible_user': 'admin'
                }
                
                if hostname and hostname != 'nan':
                    host_vars['hostname'] = hostname
                
                inventory['all']['hosts'][host_name] = host_vars
                
            except Exception as e:
                print(f"Ошибка в строке {i + 1}: {e}")
                continue
        
        # Запись в YAML файл
        with open(output_file, 'w', encoding='utf-8') as f:
            yaml.dump(inventory, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
        
        print(f"✅ Inventory файл успешно создан: {output_file}")
        print(f"📊 Обработано хостов: {len(inventory['all']['hosts'])}")
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")

# Создание тестового Excel файла
def create_sample_excel():
    """Создает пример Excel файла для тестирования"""
    sample_data = {
        'Name': ['web-server-01', 'web-server-02', 'db-server-01'],
        'HostName': ['web01.local', 'web02.local', 'db01.local'],
        'Ip-address': ['192.168.1.10/24', '192.168.1.11/24', '192.168.1.20/24']
    }
    
    df = pd.DataFrame(sample_data)
    df.to_excel('hosts.xlsx', index=False)
    print("✅ Создан пример файла: hosts.xlsx")

# Запуск
if __name__ == "__main__":
    # Создаем тестовый файл (если нужно)
    # create_sample_excel()
    
    # Запускаем парсер (используйте безопасную версию если первая не работает)
    parse_excel_to_inventory('hosts.xlsx', 'inventory.yml')
    
    # Или используйте альтернативную версию:
    # parse_excel_safe('hosts.xlsx', 'inventory.yml')