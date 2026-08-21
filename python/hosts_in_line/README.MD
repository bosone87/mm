Использование

1. Вывод списка в консоль
Для получения строки со всеми именами групп и хостов выполните:

bash
python inventory_hosts_extractor.py inventory.yml
Здесь inventory.yml — путь к вашему файлу инвентори.
Пример вывода:

text
postgres,webservers,databases,db1.example.com,web1.example.com,web2.example.com

2. Обновление переменной hosts в inventory
Если требуется автоматически заменить или добавить переменную hosts в секцию vars (на верхнем уровне или внутри all), используйте флаг --update:

bash
python inventory_hosts_extractor.py inventory.yml --update
После выполнения скрипт сохранит изменения в этом же файле. Если секция vars отсутствует, она будет создана на верхнем уровне. Если переменная hosts уже существует, её значение будет перезаписано.