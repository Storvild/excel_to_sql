"""
Создание SQL скрипта из содержимого файла Excel
    Скрипт использует модуль pyexcel. Установка:
        pip install pyexcel
        pip install pyexcel-xls
        pip install pyexcel-xlsx
        pip install pyexcel-ods
    Использование:
        python import_excel.py "Путь к файлу 1.xlsx"

  Если выполнить скрипт без параметров, то будут обработаны все файлы .xls, .xlsx, .ods находящиеся в папке скрипта

  Для подготовки скрипта, необходимо заполнить COLUMNS, где указать все поля из которых будут браться значения
  В каждом поле необходимо заполнить имя поля в БД, Имя столбца Excel, а также тип в БД
  Если тип не 'varchar', то вместо пустого значения '' ставится NULL
  Заполните в EXT_COLUMNS (todate) значение fieldvalue
  В SHEET_NAME можно записать имя листа Excel, если пустое, то берется первый лист
  В START_ROWNUM записывается  номер строки с которой будет начата обработка
  В EMPTY_BREAK_COL можно записать имя столбца пустое значение которого будет прерывать обработку

  # Пример COLUMNS и EXT_COLUMNS:
  COLUMNS = [
    {'fieldname': 'bid',     'colname': 'A', 'datatype': 'int4'},
    {'fieldname': 'account', 'colname': 'F', 'datatype': 'varchar'},
  ]
  # Данные поля будут добавлены в разделы INSERT и SELECT
  EXT_COLUMNS = [
    {'fieldname': 'creator', 'fieldvalue': 8482, 'comment': 'ФИО'},
    {'fieldname': 'todate', 'fieldvalue': "date'2021-10-31'", 'comment': ''},
  ]

"""

import os
import glob
import pyexcel
import datetime
import string
import sys

SHEET_NAME = ''        # Имя листа Excel, '' - использовать первый лист
START_ROWNUM = 2      # Обрабатывать с 2-й строки
EMPTY_BREAK_COL = 'A'  # Прерывать обработку до первого пустого значения в указанной колонке. Если передано '', то до конца

COLUMNS = [
    {'fieldname': 'a', 'colname': 'A', 'datatype': 'varchar'},
    {'fieldname': 'b', 'colname': 'B', 'datatype': 'varchar'},
    {'fieldname': 'c', 'colname': 'C', 'datatype': 'varchar'},
    {'fieldname': 'd', 'colname': 'D', 'datatype': 'varchar'},
    {'fieldname': 'e', 'colname': 'E', 'datatype': 'varchar'},
    {'fieldname': 'f', 'colname': 'F', 'datatype': 'varchar'},
    {'fieldname': 'g', 'colname': 'G', 'datatype': 'varchar'},
    {'fieldname': 'h', 'colname': 'H', 'datatype': 'varchar'},


]

EXT_COLUMNS = [
    # {'fieldname': 'creator', 'fieldvalue': 1, 'comment': ''},
]

curdir = os.path.abspath(os.path.dirname(__file__))
print('Рабочая директория: ', curdir)
WORKDIR = os.path.join(curdir, '')

list_files = glob.glob(os.path.join(WORKDIR, '*.*'))


def end_of_the_month(indate: datetime.datetime):
    import calendar
    days = calendar.monthrange(indate.year, indate.month)[1]
    res = indate.replace(day=days)
    return res


def get_filenames(workdir=WORKDIR):
    """ Получение списка файлов
        Список файлов можно передать в виде параметров скрипта
        Если в параметрах файлы не указаны, берутся все из директории скрипта
    """
    filenames = []
    if len(sys.argv) > 1:
        for filename in sys.argv[1:]:
            if os.path.exists(filename):
                filenames.append(filename)

    if not filenames:
        import glob
        filenames = glob.glob(os.path.join(workdir, '*.*'))
        filenames = [x for x in filenames if os.path.splitext(x)[1] in ('.xls', '.xlsx', '.ods')]  # Ищем только Excel файлы
        filenames = [x for x in filenames if not os.path.split(x)[1].startswith('~')]  # Убираем файлы начинающиеся с '~'
    return filenames


def get_data(filepath: str, columns: list, start_rownum: int, sheet_name=SHEET_NAME, empty_break_col=EMPTY_BREAK_COL) -> list:
    """

    :param filepath: Путь к файлу Exceld
    :param columns: Поля в виде: [{'fieldname': 'bid', 'colname': 'A', 'datatype': 'int4'},
                                  {'fieldname': 'account', 'colname': 'F', 'datatype': 'varchar'}]
    :param start_rownum: Номер строки Excel с которой начинать импорт (начиная с 1)
    :param sheet_name: Имя листа Excel. Если не указано, берется первый
    :param empty_break_col: Имя столбца, пустое значение которого будет прерывать обработку. Если '', то обрабатывать все строки
    :return: Список записей (Пример: [{'nn': 1, 'bid': 11000, 'account': '6001'},])
    """
    res = []
    try:
        fileext: str = os.path.splitext(filepath)[1]
        filename: str = os.path.split(filepath)[-1]
        if fileext in ('.ods', '.xls', '.xlsx') and not filename.startswith('~'):
            print('Обработка файла:', filepath)
            sheet = pyexcel.get_sheet(file_name=filepath, sheet_name=sheet_name)
            first_column = sheet.column[0]

            for rownum in range(start_rownum, len(first_column)+1):
                rec = {}
                # Ищем данные до первого пустого значения в первом указанном столбце
                # if str(sheet['{}{}'.format(columns[0]['colname'], rownum)]).strip() == '':
                if empty_break_col:
                    if str(sheet['{}{}'.format(empty_break_col, rownum)]).strip() == '':
                        break
                rec['row_num'] = rownum
                for column in columns:
                    #print(column['fieldname'], column['colname'])
                    try:
                        rec[column['fieldname']] = sheet['{}{}'.format(column['colname'], rownum)]
                    except:
                        pass
                res.append(rec)
    finally:
        pyexcel.free_resources()
    return res


def get_data_file_list(filepath_list: list) -> list:
    res = []
    for filepath in filepath_list:
        data = get_data(filepath, COLUMNS, START_ROWNUM)
        res.extend(data)

def get_datatype(fieldname: str, columns: list, fieldvalue: str) -> str:
    datatypes = {x['fieldname']: x['datatype'] for x in columns}
    if fieldname in datatypes:
        res = datatypes[fieldname]
    elif fieldname == 'nn':
        res = 'int4'
    else:
        res = 'varchar'
    return res

def to_sql_file(filename: str,
                data_list: list,
                columns: list,
                ext_columns: list,
                setzero: bool = True,
                prev_block='',    # Блок после WITH data
                insert_block='',  # Блок с INSERT INTO
                select_block='',  # Блок с SELECT FROM data
                end_block='',     # Блок в конце
                tablename='') -> bool:
    """
    Запись результирующего sql файла
    :param filename: Имя обрабатываемого файла Excel
    :param data_list: Данные полученные через get_data
    :param columns: список колонок. Пример в COLUMNS
    :param setzero: Устанавливать 0, если значение пустое или '#N/A' и тип поля числовой, иначе NULL
    :return:
    """
    filepath = os.path.join(WORKDIR, filename)
    content = ''
    with open(filepath, 'w', encoding='utf-8') as fw:
        # WITH data
        content = '''WITH data AS (\n  SELECT * FROM (VALUES\n'''
        for i, rec in enumerate(data_list):
            line = '    ('
            for j, field in enumerate(rec):
                fieldtype = get_datatype(field, columns, rec[field])
                if rec[field] == '#N/A':  # Неопределенное значение записывать как NULL
                    ref[field] = ''
                if j > 0:
                    line += ', '
                # Если значение пустое, но тип числовой, то ставить 0
                if setzero and rec[field] == '' and fieldtype in ('int4', 'integer', 'float8', 'btk_money', 'numeric'):
                    line += '0'
                # Если значение пустое, но тип не текстовый, то ставим NULL (например дата не может быть '')
                elif rec[field] == '' and fieldtype not in ('varchar', 'text'):
                    line += 'NULL'
                # Если значение числовое, то записывать без кавычек
                elif fieldtype in ('int4', 'integer', 'float8', 'btk_money', 'numeric'):
                    line += '{}'.format(rec[field])
                elif type(rec[field]) == bool:
                    line += 'True' if rec[field] else 'False'
                #elif type(rec[field]) in (int, float):
                #    line += '{}'.format(rec[field])
                else:
                    line += "'{}'".format(rec[field])
            line += ')'
            if i < len(data_list)-1:
                line += ','
            line += '\n'
            content += line
        if data_list:
            content += '  ) AS m ({})\n)\n\n'.format(', '.join([x for x in data_list[0].keys()]))
        else:
            content += '  ) AS m ()\n)\n\n'
        # Блок после WITH
        if prev_block:
            content += prev_block
            #content += '\n'

        # Блок INSERT
        if insert_block:
            content += insert_block
            #content += '\n'
        else:
            content += '\n\n/*\nINSERT INTO {} ('.format(tablename)
            i = 0
            for field in ext_columns:  # Доп.поля из EXT_COLUMNS
                line = '\n    '
                if i > 0:
                    line += ', '
                line += '{}'.format(field['fieldname'])
                content += line
                i += 1
            if data_list:
                for field in data_list[0]:  # Поля из первой полученной записи data_list
                    line = '\n    '
                    if i > 0:
                        line += ', '
                    line += '{}'.format(field)
                    content += line
                    i += 1
                # content += '{}'.format('\n    , '.join([x for x in data_list[0].keys()]))
                content += '\n    )\n*/'

        # Блок SELECT
        if select_block:
            content += select_block
            content += '\n'
        else:
            content += '\n\n--/*\nSELECT\n'
            i = 0
            for field in EXT_COLUMNS:  # Доп.поля из EXT_COLUMNS
                line = '\n    '
                if i > 0:
                    line += ', '
                line += '{} AS {}'.format(field['fieldvalue'], field['fieldname'])
                if field['comment']:
                    line += ' -- {}'.format(field['comment'])
                content +=line
                i += 1

            if data_list:
                for field in data_list[0]:
                    content += '\n    '
                    if i > 0:
                        content += ', '
                    content += 'CAST(m.{0} AS {1}) AS {0}'.format(field, get_datatype(field, columns, data_list[0][field]))
                    i += 1
            # content += '    {}'.format('\n    , '.join(['CAST(m.{0} AS {1}) AS {0}'.format(x, get_datatype(x, COLUMNS)) for x in data_list[0].keys()]))
            content += '\nFROM data AS m\n--*/\n'

        # Блок в конце файла
        if end_block:
            content += end_block
            content += '\n'
        # Запись результата в файл
        fw.write(content)

    return True


def main():
    filenames = get_filenames()

    if filenames:
        for filename in filenames:
            outfile = os.path.splitext(filename)[0] + '.sql'
            data_list = get_data(filename, COLUMNS, START_ROWNUM, sheet_name='')
            to_sql_file(outfile, data_list, columns=COLUMNS, ext_columns=EXT_COLUMNS, setzero=True, tablename='tablename')
            print('Результирующий файл записан: {}'.format(outfile))
        print('Обработка успешно завершена')
        # input()


if __name__ == '__main__':
    main()
