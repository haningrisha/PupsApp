from openpyxl import load_workbook, Workbook
from reporter.utils import open_xls_as_xlsx, fit_columns_width, format_date
import os
from reporter.config import header, codes, ids
import re
import datetime

target = ["Фамилия", "Имя", "Отчество", "Дата рождения", "Номер полиса"]


def get_date(ws):
    text = ws.cell(14, 1).value
    date = re.search('([0-2]\d|3[01]).(0\d|1[012]).(\d{4})', text).group()
    date = datetime.datetime.strptime(date, '%d.%m.%Y')
    return date


def get_detach_data(ws):
    data = []
    found = False
    date = get_date(ws)
    for row in ws:
        if found:
            if row[0].value is not None and row[0].value != "":
                data.append([cell.value for cell in row[1:6]] + [date])
            else:
                return data
        for i, cell in enumerate(row[1:6]):
            if cell.value == target[i]:
                if not found:
                    found = True


def get_attach_data(ws):
    data = []
    for row in ws.iter_rows(min_row=3, max_col=7):
        data.append([cell.value for cell in row])
    return data


def create_detach_report(files: list, save_directory: str, filename="File"):
    wb = Workbook()
    ws = wb.active
    data = [header]
    for file in files:
        if file.split(".")[-1] == "xls":
            wb_tmp = open_xls_as_xlsx(file)
        else:
            wb_tmp = load_workbook(file)
        ws_name = wb_tmp.sheetnames[0]
        ws_tmp = wb_tmp[ws_name]
        data += get_detach_data(ws_tmp)
    for i, r in enumerate(data):
        if i != 0:
            r.insert(5, None)
            r.insert(5, None)
            ws.append(r + codes["renessans1"] + ids["renessans_detach"])
        else:
            ws.append(r)
    for i, r in enumerate(data):
        if i != 0:
            ws.append(r + codes["renessans2"] + ids["renessans_detach"])
    format_date(ws, 4)
    format_date(ws, 8)
    fit_columns_width(ws)
    wb.save(os.path.join(save_directory, filename + ".xlsx"))


def create_attach_report(files: list, save_directory: str, filename="FileAttach"):
    wb = Workbook()
    ws = wb.active
    data = [header]
    for file in files:
        if file.split(".")[-1] == "xls":
            wb_tmp = open_xls_as_xlsx(file)
        else:
            wb_tmp = load_workbook(file)
        ws_name = wb_tmp.sheetnames[0]
        ws_tmp = wb_tmp[ws_name]
        data += get_attach_data(ws_tmp)
    for i, r in enumerate(data):
        if i != 0:
            r.insert(8, None)
            ws.append(r + codes["renessans1"] + ids["renessans_attach"])
        else:
            ws.append(r)
    for i, r in enumerate(data):
        if i != 0:
            ws.append(r + codes["renessans2"] + ids["renessans_attach"])
    format_date(ws, 4)
    format_date(ws, 6)
    format_date(ws, 7)
    fit_columns_width(ws)
    wb.save(os.path.join(save_directory, filename + ".xlsx"))


if __name__ == '__main__':
    create_detach_report(
        [
            "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/002_1060_001???35890520?_3_????_01_02_2021.xls",
            "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/002_1060_001???35719520?_19_????_03_02_2021.xlsx"
        ],
        "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test")
    create_attach_report(
        [
            "/Users/grigorijhanin/Documents/ТеорМех/динамика пупс/angle.png",
            "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/002_1682_001???35682620?_79_?????_27_01_2021 (1).xls"
        ],
        "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test"
    )
