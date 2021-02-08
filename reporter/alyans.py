from openpyxl import load_workbook, Workbook
from reporter.utils import open_xls_as_xlsx, open_csv_as_xlsx, fit_columns_width
import os
from reporter.config import header, codes, ids
from reporter.exceptions import UnsupportedNameLength

target = ["ФИО", "Номер полиса", "Дата начала обслуживания", "Дата окончания обслуживания", "Программа обслуживания",
          "Дата рождения", "Страхователь", "Дата отмены"]
column_map_attach = {
    "Номер полиса": "POLICY", "Дата начала обслуживания": "DATE_FRM", "Дата окончания обслуживания": "DATE_TO",
    "Дата рождения": "DATE_BIRTH", "Дата отмены": "DATE_CNCL"
}
column_map_detach = {
    "Номер полиса": "POLICY",
    "Дата рождения": "DATE_BIRTH", "Дата окончания обслуживания": "DATE_CNCL"
}


def format_names(names):
    surname = []
    first_name = []
    sec_name = []
    for i, name in enumerate(names):
        name = name.split(" ")
        if len(name) == 3:
            surname.append(name[0])
            first_name.append(name[1])
            sec_name.append(name[2])
        elif len(name) == 4:
            surname.append(name[0] + name[1])
            first_name.append(name[2])
            sec_name.append(name[3])
        else:
            raise UnsupportedNameLength("Ошибка имени", "Неподдерживаемый формат имени в {0} ряду".format(i))
    return {"SURNAME": surname, "FIRST_NAME": first_name, "SEC_NAME": sec_name}


def get_data(ws, column_map):
    data = {}
    for row in ws.iter_rows(max_row=1):
        for i, cell in enumerate(row):
            column_name = cell.value
            if column_name in target:
                column = [row[0].value for row in ws.iter_rows(min_row=2, min_col=i+1, max_col=i+1)]
                data.update({column_map.get(column_name, column_name): column})
    names = data.pop("ФИО")
    data.update(format_names(names))
    formatted_data = [[cell] for cell in data.get(header[0])]
    for head in header[1:8]:
        column = data.get(head)
        for i, row in enumerate(formatted_data):
            if column is None:
                row += [None]
            else:
                row += [column[i]]
    return formatted_data


def get_files_data(files, column_map):
    data = [header]
    for file in files:
        if file.split(".")[-1] == "xls":
            wb_tmp = open_xls_as_xlsx(file)
        elif file.split(".")[-1] == "csv":
            wb_tmp = open_csv_as_xlsx(file)
        else:
            wb_tmp = load_workbook(file)
        ws_name = wb_tmp.sheetnames[0]
        ws_tmp = wb_tmp[ws_name]
        data += get_data(ws_tmp, column_map)
    return data


def create_attach_report(files: list, save_directory: str, filename="FileAlyans"):
    wb = Workbook()
    ws = wb.active
    data = get_files_data(files, column_map_attach)
    for i, r in enumerate(data):
        if i != 0:
            ws.append(r + codes["alyans1"] + ids["alyans_attach"])
        else:
            ws.append(r)
    for i, r in enumerate(data):
        if i != 0:
            ws.append(r + codes["alyans2"] + ids["alyans_attach"])
    fit_columns_width(ws)
    wb.save(os.path.join(save_directory, filename + ".xlsx"))


def create_detach_report(files: list, save_directory: str, filename="FileAlyans"):
    wb = Workbook()
    ws = wb.active
    data = get_files_data(files, column_map_detach)
    for i, r in enumerate(data):
        if i != 0:
            ws.append(r + codes["alyans1"] + ids["alyans_detach"])
        else:
            ws.append(r)
    for i, r in enumerate(data):
        if i != 0:
            ws.append(r + codes["alyans2"] + ids["alyans_detach"])
    fit_columns_width(ws)
    wb.save(os.path.join(save_directory, filename + ".xlsx"))


if __name__ == '__main__':
    create_attach_report(
        [
            "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/??_2021.02.03_1492_????????????.csv",
        ],
        "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test")
