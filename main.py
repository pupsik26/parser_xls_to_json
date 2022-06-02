import xlrd
import json


def parser(files_name):
    files_name = files_name.replace(' ', '')
    array_file_name = files_name.split(',')
    for file_name in array_file_name:
        wb = xlrd.open_workbook("file/" + file_name + ".xls")
        point_sheets = wb.sheet_names()
        for wb_len in range(0, wb.nsheets):
            sh = wb.sheet_by_index(wb_len)
            all_array = []
            for i in range(1, sh.nrows):
                array_json = dict()
                for j in range(0, sh.ncols):
                    array_json[sh.cell_value(0, j)] = sh.cell_value(i, j)
                all_array.append(array_json)
            with open(file_name + "_" + point_sheets[wb_len] + ".json", "w", encoding="windows-1251") as writeJsonfile:
                json.dump(all_array, writeJsonfile, indent=4, default=str)


if __name__ == '__main__':
    filename = input('Введите название(названия) файлов(xls): ')
    parser(filename)


