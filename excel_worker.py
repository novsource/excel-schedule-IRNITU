from xls2xlsx import XLS2XLSX
from openpyxl import load_workbook


def convert_to_xlsx(path_to_file):
    x2x = XLS2XLSX(path_to_file)
    x2x.to_xlsx(path_to_file + 'x')
    print('Файл успешно конвертирован!')
