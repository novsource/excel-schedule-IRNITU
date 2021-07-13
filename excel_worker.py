import re
from operator import attrgetter
from xls2xlsx import XLS2XLSX
from openpyxl import load_workbook

book = load_workbook('data/1.xlsx')  # Рабочий файл

pairs = {
    '8.15-9.45': 1,
    '9.55-11.25': 2,
    '12.05-13.35': 3,
    '13.45-15.15': 4,
    '15.25-16.55': 5,
    '17.05 -18.35': 6
}

days = {
    'понедельник': 1,
    'вторник': 2,
    'среда': 3,
    'четверг': 4,
    'пятница': 5,
    'суббота': 6
}


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


def get_cell_of_beginning_table(worksheet):
    beg_table = {}
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value == 'Учебная группа':
                beg_table[worksheet.title] = cell
    return beg_table


def get_students_group_from_sheet(worksheet):
    beg_table = int(get_cell_of_beginning_table(worksheet).get(worksheet.title).row)  # Получаем начало таблицы из листа

    groups_cells = {}  # Словарь аудиторий с клетками

    for row in worksheet.iter_rows(min_row=beg_table, max_row=beg_table, max_col=worksheet.max_column):
        for cell in row:
            if (cell.value != None) and (cell.value != 'Учебная группа'):
                groups_cells[cell.value] = cell
                print(groups_cells.get(cell.value))

    return groups_cells


def get_started():
    for sheet_name in book.sheetnames:
        worksheet = book[sheet_name]
        # get_cells_schedule(worksheet)
        # get_students_group_from_sheet(worksheet)
        # printDataFromGroup(worksheet)
        excel_into_json(worksheet)
        # testMerged(worksheet)


def get_cells_schedule(worksheet):
    beg_table = {}
    flag = False
    for row in worksheet.iter_rows(max_col=1):
        for cell in row:
            value = str(cell.value).replace(" ", '')
            print(value)
            if days.get(value) == 1:
                beg_table['Begin'] = cell
                print(f'Начало таблицы', cell)
            if days.get(value) == 6:
                flag = True
            if flag is True and value == None:
                beg_table['End'] = cell
                print(f'Конец таблицы', cell)
    return beg_table


def printDataFromGroup(worksheet):
    json_file_out = {'pairs': pairs, 'schedule': get_students_group_from_sheet(worksheet)}
    items = get_students_group_from_sheet(worksheet).items()
    for row in range(1, worksheet.max_row):
        for col in range(1, worksheet.max_column):
            cell = worksheet.cell(row=row, column=col)
            if (cell.value != None):
                print(cell.value)


def excel_into_json(worksheet):
    groups_students = get_students_group_from_sheet(worksheet)

    json_out = {'pairs': pairs, 'schedule': groups_students.keys()}

    for key in groups_students:
        group_cell = groups_students.get(key)
        for row in range(23, worksheet.max_row):
            for col in range(group_cell.column, group_cell.column + 1):
                time = worksheet.cell(row, 2)
                cell = worksheet.cell(row, col)
                audit = get_audit(worksheet, cell)
                pair_week = get_pair_week(worksheet, time, cell)


def get_pair_week(worksheet, time, pair):
    if is_merged(worksheet, pair) and pair.value != None:
        teachers_list = get_teachers(pair)
        return 0
    if (time.value != None) and (pair.value != None):
        return 1
    if (time.value is None) and (pair.value != None):
        return 2
    if (pair.value == time.value is None):
        return None


def get_teachers(pair):
    regex = re.compile(r"[А-Я][а-я]*\s\w[.]\w[.]")
    teacher_list = re.findall(regex, str(pair.value))
    if teacher_list != None:
        return teacher_list


def get_audit(worksheet, pair):
    is_fisk = []
    audit_list = []
    audit_cell = worksheet.cell(96, pair.column + 1)
    if audit_cell.value != None:
        is_fisk = re.findall(r'стадион ИРНИТУ', audit_cell.value)
    if audit_cell.value != None and len(is_fisk) == 0: #  Если это не физкультура
        audit = str(audit_cell.value).replace('\n', ',').replace(' ', ',')
        audit_list = str(audit).split(',')
    if len(is_fisk) != 0:
        audit_list = str(audit_cell.value).split(',')
    if len(audit_list) != 0:
        return audit_list
    else:
        return None


def is_merged(worksheet, cell):
    for merged_cell in sorted(worksheet.merged_cell_ranges, key=attrgetter('coord')):
        if cell.coordinate in merged_cell:
            return True
