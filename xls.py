import xlsxwriter
import os.path
import json
from flask import Response


def create_xls(json_file):
    name, __ = os.path.splitext(json_file)

    workbook = xlsxwriter.Workbook(f'{name}.xlsx')
    return workbook


def create_worksheet(workbook):
    worksheet = workbook.add_worksheet()
    return worksheet


def get_data_key(json_file):
    with open(json_file, 'r') as file:
        data = json.load(file)
        try:
            result = data['data']
            data_for_sheets = []
            for key in result:
                for k in key:
                    data_for_sheets.append(k)
            return data_for_sheets
        except KeyError:
            return Response("Данные для записи не найдены", 404)


def get_data_value(json_file):
    with open(json_file, 'r') as file:
        data = json.load(file)
        try:
            result = data['data']
            data_for_sheets = []
            for res in result:
                for k, v in res.items():
                    data_for_sheets.append(v)
            return data_for_sheets
        except KeyError:
            return Response("Данные для записи не найдены", 404)


def clear_and_append(worksheet, data_keys, data_values):
    col = 0
    for k in data_keys:
        worksheet.write(0, col, k)
        col += 1
    col = 0
    for k, v in enumerate(data_values):
        worksheet.write_column(1, col, v)
        col += 1


