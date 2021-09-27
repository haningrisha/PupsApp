from reporter.config import codes, ids
from reporter.exceptions import UnsupportedRenCode, NoDateFound
import re
from reporter.base import AbstractReport
import datetime

target = ["Фамилия", "Имя", "Отчество", "Дата рождения", "Номер полиса"]


class RenessansReport(AbstractReport):
    def __init__(self, file):
        super(RenessansReport, self).__init__(file, None)


class RenessansAttach(RenessansReport):
    def get_data(self):
        data = []
        for row in self.ws.iter_rows(min_row=3, max_col=7):
            data += self.parse_attach_row(row)
        return data

    @staticmethod
    def parse_attach_row(row):
        r = [cell.value for cell in row]
        r.insert(8, None)
        return [r + codes["renessans2"] + ids["renessans_attach"],
                r + codes["renessans1"] + ids["renessans_attach"]]


class RenessansDetach(RenessansReport):
    def __init__(self, file):
        super().__init__(file)
        self.date = self.get_date()
        self.code = self.get_code()

    def get_data(self):
        data = []
        found = False
        for row in self.ws:
            if found:
                if row[0].value is not None and row[0].value != "":
                    data.append(self.parse_row(row))
                else:
                    return data
            for i, cell in enumerate(row[1:6]):
                if cell.value == target[i]:
                    if not found:
                        found = True

    def parse_row(self, row):
        r = [cell.value for cell in row[1:6]] + [self.date] + self.code + ids["renessans_detach"]
        r.insert(5, None)
        r.insert(5, None)
        return r

    def get_code(self):
        detach_type = self.ws.cell(1, 5).value
        if detach_type == "АО \"Поликлинический комплекс\"":
            return codes["renessans1"]
        elif detach_type == "АО \"Современные медицинские технологии\"":
            return codes["renessans2"]
        else:
            raise UnsupportedRenCode("Ошибка кода Ренессанс", "Неподдерживаемая комания {0}, \n "
                                                              "В файле {1}".format(detach_type, self.file))

    def get_date(self):
        text = self.ws.cell(14, 1).value
        date = re.search('([0-2]\d|3[01]).(0\d|1[012]).(\d{4})', text)
        if date is None:
            raise NoDateFound("14, 1", self.file)
        date = date.group()
        date = datetime.datetime.strptime(date, '%d.%m.%Y')
        return date


def get_renessans(file, attach=False, detach=False):
    if RenessansReport.is_reportable(file):
        if attach:
            return RenessansAttach(file)
        elif detach:
            return RenessansDetach(file)


if __name__ == '__main__':
    test_attach = RenessansAttach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Ренессанс Прикреп/002_1682_001???35682620?_79_?????_27_01_2021 (1).xls")
    test_attach.get_data()
    test_detach = RenessansDetach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Ренессанс Откреп/002_1060_001???35719520?_19_????_03_02_2021.xls")
    test_detach.get_data()
