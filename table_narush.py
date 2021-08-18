# Из шаблонных файлов в папке dir_res собирает все файлы и значения из них суммирует в главный файл file_template_xl
# ...
# INSTALL
# pip install openpyxl
# ...

import os
import openpyxl


# функция конвертации данных
# в ячейке может быть целое, дробное число, строка или пусто
# если число целое или дробное, то выдать, иначе выдать 0
def conv_cell(cell_value):
    # вычисление координат полученой ячейки
    cell_coordinate = wb_narush_cells_range[indexR][indexC].coordinate

    if type(cell_value) == int:
        return cell_value
    elif type(cell_value) == float:
        return int(cell_value)
    elif type(cell_value) is None:
        cell_value = 0
        return cell_value
    elif type(cell_value) == str:
        try:
            cell_value = int(cell_value)
            return cell_value
        except TypeError:
            print(f'Ошибка в ячейке {cell_coordinate}: значение ячейки строка "{cell_value}"')
            cell_value = 0
            return cell_value
        except ValueError:
            print(f'Ошибка в ячейке {cell_coordinate}: значение ячейки строка "{cell_value}"')
            cell_value = 0
            return cell_value
    else:
        cell_value = 0
        return cell_value


# файл для вставления данных
file_template_xl = 'Таблица по нарушениям.xlsx'
# папка с файлами откуда берутся данные
dir_res = r'res'
# диапазон для обработки данных во всех файлах
template_cells_range = 'E7:V33'

# открывается файл "приёмник", назначается активный лист, выбирается диапазон ячеек
wb_narush = openpyxl.load_workbook(file_template_xl)
wb_narush_s = wb_narush.active
wb_narush_cells_range = wb_narush_s[template_cells_range]

# чищу (заполняю нулями) "приёмник" от старых данных
for row_in_range in wb_narush_cells_range:
    for cell_in_row in row_in_range:
        indexR = wb_narush_cells_range.index(row_in_range)
        indexC = row_in_range.index(cell_in_row)
        wb_narush_cells_range[indexR][indexC].value = 0


# прохожусь по папке dir_res, из каждого файла беру данные, вставляю в "приёмник""
for folders, dirs, files in os.walk(dir_res):
    for file in files:
        print('_'*50)
        print(f'Файл = "{file}"\n')

        # открывается файл "источник", назначается активный лист, выбирается диапазон ячеек
        wb_file = openpyxl.load_workbook(os.path.join(dir_res, file), read_only=True, data_only=True)
        wb_file_s = wb_file.active
        file_cells_range = wb_file_s[template_cells_range]

        # заполняю "приёмник" данными
        for row_in_range in file_cells_range:
            for cell_in_row in row_in_range:
                indexR = file_cells_range.index(row_in_range)
                indexC = row_in_range.index(cell_in_row)

                cell_source = conv_cell(cell_in_row.value)
                cell_recipient = conv_cell(wb_narush_cells_range[indexR][indexC].value)

                wb_narush_cells_range[indexR][indexC].value = cell_recipient + cell_source

        wb_file.close()

# сохраняю файл шаблона и закрываю его
wb_narush.save(file_template_xl)
wb_narush.close()

print('_'*50)
print('Запомните ошибки и нажмите ENTER для закрытия окна.')
print('Исправьте имеющиеся ошибки и запустите программу ещё раз.')
input()
