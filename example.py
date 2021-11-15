"""
Создание скрипта SQL из Excel, OpenOffice файла

    python example.py "Путь к файлу 1.xlsx"

"""
import os

SHEET_NAME = ''        # Имя листа с данными. Если не указано, то берется первый
START_ROWNUM = 2       # Начинать обработку с 2-й строки Excel
EMPTY_BREAK_COL = 'A'  # При встрече пустого значения в колонке A, прерывать обработку

COLUMNS = [
    {'fieldname': 'field1', 'colname': 'A', 'datatype': 'int4'},
    {'fieldname': 'field2', 'colname': 'B', 'datatype': 'varchar'},
    {'fieldname': 'field4', 'colname': 'D', 'datatype': 'bool'},
]

EXT_COLUMNS = [
    {'fieldname': 'creator', 'fieldvalue': 1, 'comment': ''},
]


import excel_to_sql

def get_prev_block():
    res = """
-- Дополнительный WITH
, params AS (
    SELECT date'2021-11-01' AS fromdate, date'2021-11-30' AS todate
)
-- Запрос после WITH        
-- SELECT * FROM data
    """
    return res

def get_insert_block() -> str:
    res = """"""
    return res

def get_select_block() -> str:
    res = """"""
    return res

def get_end_block() -> str:
    res = """
-- Блок в конце скрипта
"""
    return res

def to_file(filename: str, data_list: list) -> bool:
    """
    :param filename: Имя файла
    :param data_list: Данные из Excel
    :param reptype: Тип отчета Дебиторка 'Дт' или Кредиторка 'Кт'. Используется в поле js
    :return: Успешность завершения операции
    """
    res = excel_to_sql.to_sql_file(
        filename,
        data_list,
        COLUMNS,
        EXT_COLUMNS,
        setzero=True,
        prev_block=get_prev_block(),
        insert_block=get_insert_block(),
        select_block=get_select_block(),
        end_block=get_end_block(),
        tablename='mytable')
    return res


def main():
    filenames = excel_to_sql.get_filenames()
    for filename in filenames:
        outfile = os.path.splitext(filename)[0] + '.sql'
        data_list = excel_to_sql.get_data(filename, COLUMNS, START_ROWNUM, sheet_name=SHEET_NAME, empty_break_col=EMPTY_BREAK_COL)
        to_file(outfile, data_list)
        print('Обработано записей: {}. Результирующий файл ДЗ записан: {} '.format(len(data_list), outfile))
    print('Обработка успешно завершена.')


if __name__ == '__main__':
    main()
