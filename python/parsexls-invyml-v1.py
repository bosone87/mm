import pandas as pd
import yaml
import ipaddress
from typing import Dict, List, Any

def parse_excel_to_inventory(excel_file: str, output_file: str = 'inventory.yml') -> None:
    """
    Парсит Excel файл и генерирует inventory.yml для Ansible
    
    Args:
        excel_file (str): Путь к Excel файлу
        output_file (str): Путь для сохранения inventory.yml
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
        inventory = {
            'all': {
                'hosts': {},
                'children': {}
            }
        }
        
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
                # Проверяем корректность IP-адреса с маской
                network = ipaddress.ip_network(ip_address, strict=False)
                # Извлекаем IP без маски для ansible_host
                ip_without_mask = str(network.network_address)
            except ValueError as e:
                print(f"Ошибка в строке {index + 1}: некорректный IP-адрес '{ip_address}' - {e}")
                continue
            
            # Создание записи для хоста
            host_vars = {
                'ansible_host': ip_without_mask,
                'ansible_user': 'your_username'  # Замените на нужное имя пользователя
            }
            
            # Добавляем hostname если он указан
            if hostname and not pd.isna(hostname):
                host_vars['hostname'] = hostname
            
            # Добавляем маску сети если нужно
            host_vars['network_mask'] = network.prefixlen
            
            inventory['all']['hosts'][host_name] = host_vars
        
        # Запись в YAML файл
        with open(output_file, 'w', encoding='utf-8') as f:
            yaml.dump(inventory, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
        
        print(f"Inventory файл успешно создан: {output_file}")
        print(f"Обработано хостов: {len(inventory['all']['hosts'])}")
        
    except FileNotFoundError:
        print(f"Ошибка: Файл '{excel_file}' не найден")
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")

def parse_excel_with_groups(excel_file: str, output_file: str = 'inventory.yml', group_column: str = None) -> None:
    """
    Расширенная версия с поддержкой групп
    
    Args:
        excel_file (str): Путь к Excel файлу
        output_file (str): Путь для сохранения inventory.yml
        group_column (str): Название столбца для группировки (опционально)
    """
    
    try:
        df = pd.read_excel(excel_file)
        
        required_columns = ['Name', 'HostName', 'Ip-address']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Отсутствуют необходимые столбцы: {', '.join(missing_columns)}")
        
        inventory = {
            'all': {
                'hosts': {},
                'children': {}
            }
        }
        
        for index, row in df.iterrows():
            host_name = str(row['Name']).strip()
            hostname = str(row['HostName']).strip()
            ip_address = str(row['Ip-address']).strip()
            
            if not host_name or pd.isna(host_name):
                continue
            
            try:
                network = ipaddress.ip_network(ip_address, strict=False)
                ip_without_mask = str(network.network_address)
            except ValueError as e:
                print(f"Ошибка в строке {index + 1}: некорректный IP-адрес '{ip_address}' - {e}")
                continue
            
            host_vars = {
                'ansible_host': ip_without_mask,
                'ansible_user': 'your_username',
                'network_mask': network.prefixlen
            }
            
            if hostname and not pd.isna(hostname):
                host_vars['hostname'] = hostname
            
            # Обработка групп если указана колонка для группировки
            if group_column and group_column in df.columns:
                group_name = str(row[group_column]).strip()
                if group_name and not pd.isna(group_name):
                    if group_name not in inventory['all']['children']:
                        inventory['all']['children'][group_name] = {'hosts': {}}
                    inventory['all']['children'][group_name]['hosts'][host_name] = host_vars
                else:
                    inventory['all']['hosts'][host_name] = host_vars
            else:
                inventory['all']['hosts'][host_name] = host_vars
        
        with open(output_file, 'w', encoding='utf-8') as f:
            yaml.dump(inventory, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
        
        print(f"Inventory файл успешно создан: {output_file}")
        
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")

# Пример использования
if __name__ == "__main__":
    # Базовый вариант
    parse_excel_to_inventory('hosts.xlsx', 'inventory.yml')
    
    # Вариант с группами (если есть столбец 'Group' в Excel)
    # parse_excel_with_groups('hosts.xlsx', 'inventory_with_groups.yml', 'Group')