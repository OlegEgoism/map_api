Программа создана для проверки адресов на количество субъектов.

Работает с https://yandex.ru/maps-api/

Для запуска нужно в личном кабинете создать свой "Новый ключ" из API Поиска по организациям.
![image](https://github.com/OlegEgoism/map_api/assets/81327146/2f268fb6-768d-49bf-a7f6-765a0c5b51dc)

Для установки всех библиотек из requirements.txt нужно запустить команду: pip freeze -> requirements.txt

Для запуска приложения нужно в файле найти mian.py найти строчку кода: if __name__ == '__main__': и нажать на зелёную стрелку.
![image](https://github.com/OlegEgoism/map_api/assets/81327146/4566df7a-1dde-4df2-a1dc-7cd47a548ff2)

Документация в помощь:
- https://yandex.ru/dev/geosearch/doc/ru/
- https://yandex.ru/dev/geosearch/doc/ru/request


Механизм работы:
- Входящий файл (первичные данные) для поиска: input.xls
- Содержимое:
Адреса объектов. (туда пишем адреса для проверки).

  
- Исходящий файл (результат): output.xls
- Содержимое:
Дата сверки (гггг-мм-дд)
Название (фильтр по поиску найденных объектов заказчика)
Адрес организации (найденный адрес объекта заказчика)
Название организации (найденная организация по указанному адресу)
Контактные данные (e-mail, телефон, время работы)
ID организации (уникальный номер в Яндекс картах организации)
