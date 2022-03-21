from reporter.exceptions import UnsupportedSogazCode
from datetime import datetime
from reporter.report import AbstractReport
from reporter import utils
from reporter import common
from reporter import sogaz_conf as conf


SOGAZ_ACCESS_TYPES = {
    common.AccessTypes.PK: 'АО "Поликлинический комплекс"',
    common.AccessTypes.SMT: 'АО "Современные медицинские технологии"'
}


def str_to_date(dates):
    dates = [datetime.strptime(date, "%d.%m.%Y") for date in dates]
    return dates


def sum_list(program):
    s = []
    for i in program:
        s += i
    return s


class SogazReport(AbstractReport):
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

    def get_code(self, program, work_place=""):
        matcher = conf.get_sogaz_code_matcher()
        code = matcher.get_code(program, work_place, self.access_type)
        if code is None:
            raise UnsupportedSogazCode("Код Согаз не распознан", "Не удалось распознать код в файле {0} в "
                                                                 "строке {1}".format(self.file, program))
        return code

    @property
    def access_type(self) -> common.AccessTypes:
        if utils.is_value_in_sheet(self.ws, SOGAZ_ACCESS_TYPES[common.AccessTypes.PK]):
            return common.AccessTypes.PK
        elif utils.is_value_in_sheet(self.ws, SOGAZ_ACCESS_TYPES[common.AccessTypes.SMT]):
            return common.AccessTypes.SMT
        else:
            raise UnsupportedSogazCode(
                "Организация Согаз не распознана", "Не удалось распознать код в файле {0}".format(self.file)
            )


class SogazAttach(SogazReport):
    def get_data(self):
        data, header_row = self.get_content(True)
        data_formatted = []
        list_type = self.ws.cell(header_row, 1).value
        list_type = list_type.lower().split(" ")
        for row in data:
            if "изменение" in list_type:
                codes = self.get_code(row[10], row[11])
                for code in codes:
                    data_formatted.append(
                        row[0:3] + str_to_date([row[3]]) + [row[4]] + str_to_date(row[8:10]) + [None] +
                        code)
            elif "прикрепление" in list_type:
                codes = self.get_code(row[12], row[13])
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
            codes = self.get_code(row[6], row[7])
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

