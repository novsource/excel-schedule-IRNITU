from xls2xlsx import XLS2XLSX
from openpyxl import load_workbook

book = load_workbook('data/3.xlsx')

def convert_to_xlsx(path_to_file):
    x2x = XLS2XLSX(path_to_file)
    x2x.to_xlsx(path_to_file + 'x')
    print('Файл успешно конвертирован!')


def print_data_from_excel():
    for sheetname in book.sheetnames:
        worksheet = book[sheetname]
        for row in worksheet.iter_rows():
                for cell in row:
                   if cell.value != None:
                       print(cell.value)
                       print('\t', end='')

