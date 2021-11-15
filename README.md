# excel_to_sql
Преобразование данных Excel, OpenOffice в SQL запрос

Скрипт использует модуль **pyexcel**. Установка:
```
    pip install pyexcel
    pip install pyexcel-xls
    pip install pyexcel-xlsx
    pip install pyexcel-ods
```    
## Использование:
```
    python excel_to_sql.py "example.xlsx"
```
В результате файл example.xlsx с содержимым:
| №             | Наименование    | Дата       | bool   |
| ------------- | :-------------: | -----:     | --     |
| 1             | Строка 1        | 01.01.2021 | ИСТИНА |
| 2             | Строка 2        | 01.02.2021 | ЛОЖЬ   |
| 3             | Строка 3        |            |        |

Преобразуется в SQL-скрипт example.sql:
```
WITH data AS (
  SELECT * FROM (VALUES
    (2, '1', 'Строка 1', '2021-01-01', 'True'),
    (3, '2', 'Строка 2', '2021-02-01', 'False'),
    (4, '3', 'Строка 3', '', '')
  ) AS m (nn, a, b, c, d)
)

/*
INSERT INTO tablename (
    nn
    , a
    , b
    , c
    , d
    )
*/

--/*
SELECT

    CAST(m.nn AS int4) AS nn
    , CAST(m.a AS varchar) AS a
    , CAST(m.b AS varchar) AS b
    , CAST(m.c AS varchar) AS c
    , CAST(m.d AS varchar) AS d
FROM data AS m
--*/
```
## Использование в своем скрипте
Можно также использовать модуль excel_to_sql.py в своем скрипте.
Пример в файле example.py

В нем указаны следующие переменные:
```
SHEET_NAME = ''        # Имя листа с данными. Если не указано, то берется первый
START_ROWNUM = 2       # Начинать обработку с 2й строки Excel
EMPTY_BREAK_COL = 'A'  # При встрече пустого значения в колонке A, прерывать обработку

# Описание полей: 
COLUMNS = [
    {'fieldname': 'field1', 'colname': 'A', 'datatype': 'int4'},
    {'fieldname': 'field2', 'colname': 'B', 'datatype': 'varchar'},
    {'fieldname': 'field4', 'colname': 'C', 'datatype': 'bool'},
]
# Описание дополнительных полей для блока SELECT:
EXT_COLUMNS = [
    {'fieldname': 'creator', 'fieldvalue': 1, 'comment': ''},
]
```
Также при вызове ф-ции get_data можно передать свои блоки insert_block и select_block и тогда они не будут формироваться автоматически.

При выполнении
```
    python example.py "example.xlsx"
```
В результате файл example.xlsx с содержимым:
| №             | Наименование    | Дата       | bool   |
| ------------- | :-------------: | -----:     | --     |
| 1             | Строка 1        | 01.01.2021 | ИСТИНА |
| 2             | Строка 2        | 01.02.2021 | ЛОЖЬ   |
| 3             | Строка 3        |            |        |

Преобразуется в SQL-скрипт example.sql:
```
WITH data AS (
  SELECT * FROM (VALUES
    (2, 1, 'Строка 1', 'True'),
    (3, 2, 'Строка 2', 'False'),
    (4, 3, 'Строка 3', NULL)
  ) AS m (nn, field1, field2, field4)
)


-- Дополнительный WITH
, params AS (
    SELECT date'2021-11-01' AS fromdate, date'2021-11-30' AS todate
)
-- Запрос после WITH        
-- SELECT * FROM data
    

/*
INSERT INTO mytable (
    creator
    , nn
    , field1
    , field2
    , field4
    )
*/

--/*
SELECT

    CAST(m.nn AS int4) AS nn
    , CAST(m.field1 AS int4) AS field1
    , CAST(m.field2 AS varchar) AS field2
    , CAST(m.field4 AS bool) AS field4
FROM data AS m
--*/

-- Блок в конце скрипта
```

Без указания параметров, скрипт находит все Excel и OpenOffice файлы в каталоге со скриптом и создает там же sql файлы.