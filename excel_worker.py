from xls2xlsx import XLS2XLSX
from openpyxl import load_workbook

book = load_workbook('data/1.xlsx')


def convert_to_xlsx(path_to_file):
    x2x = XLS2XLSX(path_to_file)
    x2x.to_xlsx(path_to_file + 'x')
    print('Файл успешно конвертирован!')


def print_data_from_excel():
    for sheet_name in book.sheetnames:
        worksheet = book[sheet_name]
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    print(cell.value)
                    print('\t', end='')


def find_cell_of_beginning_table():
    beg_table = {}
    for sheet_name in book.sheetnames:
        worksheet = book[sheet_name]
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value == 'Учебная группа':
                    beg_table[sheet_name] = cell.coordinate
    return beg_table


