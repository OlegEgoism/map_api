import requests
import os
import json
import csv
from openpyxl import load_workbook
import time
from datetime import date
from datetime import datetime

from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

token_file = 'token.txt'
input_file = 'input.xlsx'


def check_file(name_file):
    """Проверка наличия файлов input.xlsx и token.txt"""
    if os.path.isfile(name_file):
        check = True
    else:
        check = False
    return check


def create_folder(folder):
    """Создаем папки с файлами проверок"""
    if not os.path.isdir(folder):
        os.mkdir(folder)
    else:
        pass


def show_data_now():
    """Создание даты на момент проверки"""
    current_date = date.today()
    return current_date


def get_token(name_file):
    """Проверка наличия токена в файле"""
    if (check_file(name_file)):
        with open(name_file, 'r') as file_token:
            token_read = file_token.read()
        file_token.close()
    if ('' != token_read):
        token = token_read
    else:
        token = None
    return token


def get_input_adr(name_file):
    """Получение списка адресов из столбца файла input.xlsx"""
    wb = load_workbook(name_file)
    sheet = wb.active
    input_adr = []
    for cell in sheet['A2': 'A' + str(sheet.max_row)]:
        for value in cell:
            if value.value:
                input_adr.append(str(value.value))
    return input_adr


def filter_fgbu_fkp(check_txt):
    """Фильтрация по службам и федеральным органам власти"""
    fgbu = 'ФГБУ'
    fkp = 'ФКП'
    if (check_txt.find(fgbu) == -1) & (check_txt.find(fkp) == -1):
        return check_txt
    else:
        return None


def get_num_of_str(str_adr):
    """Извлечь из строки адреса (номер дома, здания)"""
    length = len(str_adr)
    integers = []
    i = 0  # индекс текущего символа
    while i < length:
        s_int = ''  # строка для нового числа
        while i < length and '0' <= str_adr[i] <= '9':
            s_int += str_adr[i]
            i += 1
        i += 1
        if s_int != '':
            integers.append(int(s_int))
    return integers


def wrtie_info_in_file_xls_pack(result):
    """Запись данных в выходные файлы xls"""
    workbook = load_workbook("outpack.xlsx")  # ---- Запись данных в файл outpack.xlsx
    date_today = datetime.now().strftime('%Y-%m-%d')  # "Дата сверки"
    # sheet = workbook.create_sheet(date_today)  #  Если необходимо создание дополнительных страниц
    sheet = workbook.active
    sheet["A1"] = "Дата сверки"
    sheet["B1"] = "Название"
    sheet["C1"] = "Адрес организации"
    sheet["D1"] = "Название организации"
    sheet["E1"] = "Контактные данные"
    sheet["F1"] = "ID организации"
    # Установка ширины столбцов
    sheet.column_dimensions['A'].width = 12
    sheet.column_dimensions['B'].width = 60
    sheet.column_dimensions['C'].width = 60
    sheet.column_dimensions['D'].width = 60
    sheet.column_dimensions['E'].width = 100
    sheet.column_dimensions['F'].width = 16

    for column in range(1, 7):
        sheet.cell(row=1, column=column).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    index = 2
    for key, value in result.items():
        if type(value) == list:
            for v in value:
                property = v.get('properties')  # Все данные в json
                print("-------", property.get("CompanyMetaData"))
                company_metadata = property.get("CompanyMetaData")
                description = property.get("description")  # "Название"
                address = company_metadata.get('address')  # "Адрес организации"
                name = property.get('name')  # "Название организации"
                contact_email = company_metadata.get('url')  # "Контактные данные"
                contact_phone = company_metadata.get('Phones')
                phone_numbers_string = ''
                for phone in contact_phone:
                    phone_number = phone.get('formatted')
                    phone_numbers_string += f"{phone_number} | "
                phone_numbers_info = phone_numbers_string[:-2]  # "Контактные данные"
                contact_work_time = company_metadata.get('Hours').get('text')  # "Контактные данные"
                id_yandex = company_metadata.get('id')  # "ID организации"
                # Запись данных без фильтрации
                sheet.cell(row=index, column=1, value=date_today)  # "Дата сверки"
                sheet.cell(row=index, column=2, value=description)  # "Название"
                sheet.cell(row=index, column=3, value=address)  # "Адрес организации"
                sheet.cell(row=index, column=4, value=name)  # "Название организации"
                sheet.cell(row=index, column=5, value=f'{contact_email} | {phone_numbers_info} | {contact_work_time}')  # "Контактные данные"
                sheet.cell(row=index, column=6, value=id_yandex)  # "ID организации"
                index += 1

    workbook.save("outpack.xlsx")

#
# def write_in_folder_and_file_csv(result):
#     # Создаем папку с датой проверки
#     today = datetime.now().strftime('%Y-%m-%d')
#     create_folder(today)
#     num_dom_list = get_num_of_str(result['input_adr'])
#     print("Номер дома:", *num_dom_list)
#
#     for key, value in result.items():
#         if type(value) == list:
#             for v in value:
#                 property = v.get('properties')
#                 print(property)
#                 name_company = filter_fgbu_fkp(property.get('name'))
#                 if (num_dom_list == get_num_of_str(property.get('description'))) & (name_company != None):
#                     new_name_file = today + '/' + property.get('name') + '.csv'
#                     with open(new_name_file, 'w') as csvfile:
#                         fieldnames = ['Дата проверки', 'Адрес', 'Название компании', 'Контактные данные']
#                         writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
#                         writer.writeheader()
#                         writer.writerow({'Дата сверки': today,
#                                          'Адрес организации': property.get('description'),
#                                          'Название организации': property.get('name'),
#                                          'Контактные данные': property.get('url')})


def get_info_api(token, list_adr):
    ''' Получения данных с API Поиска по организациям'''
    for adr in list_adr:
        time.sleep(1)  # Задержка в 1 секунду
        url = 'https://search-maps.yandex.ru/v1/?text={}&type=biz&results=50&lang=ru_RU&apikey={}'.format(adr, token)
        # print(url)
        res = requests.get(url)
        contents = res.text
        result = json.loads(contents)
        # print(result)
        result['input_adr'] = adr
        # Запись нефильтрованных, всех данных
        wrtie_info_in_file_xls_pack(result)
        # # Запись в csv файл
        # write_in_folder_and_file_csv(result)
        # # Тестовая выдача
        # write_in_folder_and_file_csv_test(result)


if __name__ == '__main__':
    # Проверим наличии исходного файла с адресами
    check_file(input_file)
    # Проверим наличия токена в файле
    token = get_token(token_file)
    # Получим список исходных адресов
    adr_list = get_input_adr(input_file)
    # Запустим API
    get_info_api(token, adr_list)
    # wrtie_info_in_file()
    show_data_now()
