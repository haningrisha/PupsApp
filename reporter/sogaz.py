from reporter.exceptions import UnsupportedSogazCode
from datetime import datetime
from reporter.base import Report


def str_to_date(dates):
    dates = [datetime.strptime(date, "%d.%m.%Y") for date in dates]
    return dates


def sum_list(program):
    s = []
    for i in program:
        s += i
    return s


class SogazReport(Report):
    def __init__(self, file):
        super().__init__(file)

    def get_content(self, return_header_row=False):
        data = []
        found = False
        header_row = 0
        for i, row in enumerate(self.ws):
            row_values = [cell.value for cell in row]
            if found:
                if row[0].value is not None and row[0].value != "":
                    data.append([cell.value for cell in row[1:]])
                else:
                    break
            elif {"Имя", "Фамилия", "Отчество"}.issubset(set(row_values)):
                found = True
                header_row = i - 1
        if return_header_row:
            return data, header_row
        return data

    def get_code(self, program):
        string_program = program
        program = program.split(';')
        program = sum_list([p.split(" ") for p in program])
        program = [p.lower() for p in program]
        program = sum_list([p.split("_") for p in program])
        program = sum_list([p.split("\"") for p in program])
        if {"\"аброссия\"", "аброссия", "аб"} & set(program):
            code = [["АБС СТОМАТОЛОГИЯ Прямой доступ", "ДС№9 к 0618RP137 АБРОССИЯ", 1102]]
        elif "невское" in program:
            code = [["Невское ПКБ СОГАЗ", "ДС№10 к 0618RP137", 4001439]]
        elif "мариинский" in program:
            code = [["Мариинский театр Прямой доступ ПК", "0618RР137 Мариинский театр Прямой доступ ПК", 4001438],
                    ["Мариинский театр Прямой доступ Согаз КДЦ", "0618RВ138 Мариинский театр Прямой доступ КДЦ",
                     4001438]]
        elif "арктика" in program:
            code = [["СОГАЗ Прямой доступ Арктика ПК", "0618RP137 Арктика ПК", 1102],
                    ["СОГАЗ Прямой доступ Арктика КДЦ", "0618RВ138 Арктика КДЦ", 1102]]
        elif {"\"цтсс\"", "цтсс"} & set(program):
            code = [["ЦТСС прямой контроль", "0618RP137 (ЦТСС)", 4001294]]
        elif {"амб.", "пнд", "амбулаторный"} & set(program):
            code = [["СОГАЗ ПРЯМОЙ ДОСТУП ПК", "0618RР137 ПРЯМОЙ ДОСТУП ПК", 4000971]]
            if "стоматология" in program:
                code += [["СОГАЗ СТОМАТОЛОГИЯ Прямой доступ ПК", "0618RР137 СТОМАТОЛОГИЯ ПРЯМОЙ ДОСТУП ПК", 1102]]
        elif "стоматология" in program:
            code = [["СОГАЗ СТОМАТОЛОГИЯ Прямой доступ ПК", "0618RР137 СТОМАТОЛОГИЯ ПРЯМОЙ ДОСТУП ПК", 1102]]
        else:
            raise UnsupportedSogazCode("Код Согаз не распознан", "Не удалось распознать код в файле {0} в "
                                                                 "строке {1}".format(self.file, string_program))
        return code


class SogazAttach(SogazReport):
    def get_data(self):
        data, header_row = self.get_content(True)
        data_formatted = []
        list_type = self.ws.cell(header_row, 1).value
        list_type = list_type.lower().split(" ")
        for row in data:
            if "изменение" in list_type:
                codes = self.get_code(row[10])
                for code in codes:
                    data_formatted.append(
                        row[0:3] + str_to_date([row[3]]) + [row[4]] + str_to_date(row[8:10]) + [None] +
                        code)
            elif "прикрепление" in list_type:
                codes = self.get_code(row[12])
                for code in codes:
                    data_formatted.append(
                        row[0:3] + str_to_date([row[3]]) + [row[9]] + str_to_date(row[10:12]) + [None] +
                        code)
        return data_formatted


class SogazDetach(SogazReport):
    def get_data(self):
        data = self.get_content()
        data_formatted = []
        for row in data:
            codes = self.get_code(row[6])
            for code in codes:
                data_formatted.append(
                    row[0:3] + str_to_date([row[3]]) + [row[4]] + [None, None] + str_to_date([row[5]]) +
                    code[0:2] + [2])
        return data_formatted


if __name__ == '__main__':
    attach = SogazAttach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Согаз/Прикреп/v8_7CE0_973c.xls")
    detach = SogazDetach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Согаз/Откреп/v8_7CE0_8a2f.xls")
    attach.get_data()
    detach.get_data()

