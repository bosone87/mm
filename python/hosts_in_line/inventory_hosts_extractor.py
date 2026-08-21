#!/usr/bin/env python3
"""
Скрипт для извлечения всех групп, подгрупп и имён хостов из Ansible inventory (YAML).
Выводит строку вида: group1,group2,host1,host2,...
Опционально может обновить переменную hosts в файле inventory (секция vars).
"""

import sys
import yaml
from pathlib import Path

def collect_names(obj, group_names, host_names):
    """
    Рекурсивно обходит структуру inventory и заполняет множества group_names и host_names.
    """
    if isinstance(obj, dict):
        for key, value in obj.items():
            if key == 'hosts':
                # hosts может быть словарём (ключи = имена хостов) или списком
                if isinstance(value, dict):
                    for host in value.keys():
                        host_names.add(host)
                elif isinstance(value, list):
                    for host in value:
                        if isinstance(host, str):
                            host_names.add(host)
            elif key == 'children':
                if isinstance(value, dict):
                    for child_name, child_value in value.items():
                        # child_name — имя группы
                        group_names.add(child_name)
                        collect_names(child_value, group_names, host_names)
            elif key == 'vars':
                # Переменные групп не нужны для списка имён
                continue
            elif key == 'all':
                # Корневая группа — обрабатываем её содержимое, но имя не добавляем
                collect_names(value, group_names, host_names)
            else:
                # Возможно, это группа, определённая на верхнем уровне
                if isinstance(value, dict) and ('hosts' in value or 'children' in value):
                    group_names.add(key)
                    collect_names(value, group_names, host_names)
                # иначе это может быть переменная, тег и т.п. — игнорируем

def main():
    if len(sys.argv) < 2:
        print("Usage: python inventory_hosts_extractor.py <inventory.yml> [--update]")
        sys.exit(1)

    inventory_path = Path(sys.argv[1])
    if not inventory_path.exists():
        print(f"Файл {inventory_path} не найден")
        sys.exit(1)

    # Загружаем YAML
    with open(inventory_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)

    if not isinstance(data, dict):
        print("Ошибка: inventory должен быть словарём (YAML mapping)")
        sys.exit(1)

    group_names = set()
    host_names = set()

    # Начинаем обход с корня
    collect_names(data, group_names, host_names)

    # Объединяем имена, сохраняя порядок добавления (для детерминированности)
    # Сначала группы, потом хосты (можно изменить по желанию)
    all_names = list(dict.fromkeys(sorted(group_names))) + list(dict.fromkeys(sorted(host_names)))
    result_string = ','.join(all_names)

    print(result_string)

    # Если указан флаг --update, обновляем/добавляем переменную hosts в секции vars
    if '--update' in sys.argv:
        update_inventory_hosts_var(data, result_string, inventory_path)

def update_inventory_hosts_var(data, hosts_string, file_path):
    """
    Находит секцию vars (на верхнем уровне или внутри all) и заменяет/добавляет переменную hosts.
    """
    # Проверяем верхний уровень
    if 'vars' in data and isinstance(data['vars'], dict):
        data['vars']['hosts'] = hosts_string
    elif 'all' in data and isinstance(data['all'], dict) and 'vars' in data['all'] and isinstance(data['all']['vars'], dict):
        data['all']['vars']['hosts'] = hosts_string
    else:
        # Если секции vars нет, создаём её на верхнем уровне
        if 'vars' not in data:
            data['vars'] = {}
        data['vars']['hosts'] = hosts_string

    # Записываем обратно в файл
    with open(file_path, 'w', encoding='utf-8') as f:
        yaml.safe_dump(data, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
    print(f"Файл {file_path} обновлён: переменная hosts установлена.")

if __name__ == '__main__':
    main()