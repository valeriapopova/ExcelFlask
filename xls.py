import xlsxwriter

from flask import Response


def create_xls():

    workbook = xlsxwriter.Workbook(f'random.xlsx')
    return workbook


def create_worksheet(workbook):
    worksheet = workbook.add_worksheet()
    return worksheet


def get_data_key(json_file):
    try:
        result = json_file['data']
        data_for_sheets = []
        for key in result:
            for k in key:
                if k not in data_for_sheets:
                    data_for_sheets.append(k)
        return [data_for_sheets]
    except KeyError:
        return Response("Данные для записи не найдены", 404)


def get_data_value(json_file):
    try:
        result = json_file['data']
        data_for_sheets = []
        for res in result:
            r = res.values()
            values = list(r)
            data_for_sheets.append(values)
        return data_for_sheets
    except KeyError:
        return Response("Данные для записи не найдены", 404)


def clear_and_append(worksheet, data_keys, data_values):
    col = 0
    for k in data_keys:
        for data in k:
            worksheet.write(0, col, data)
            col += 1
    row = 1
    col = 0
    for v in data_values:
        if type(v) == list:
            for values in v:
                if type(values) == list:
                    worksheet.write_column(row, col, values)
                    # row += 1
                    col += 1
                else:
                    worksheet.write_row(row, col, v)
                    col += 1

