from reporter.base import Report
from reporter.exceptions import UnsupportedNameLength, UnrecognisedType, UnsupportedDateQuantity, NoDateFound, \
    FileTypeWasNotDefined
import re


class RosgosstrahReport(Report):
    def __init__(self, file, pattern):
        super().__init__(file)
        self.pattern = pattern

    def find_data_parts(self, min_col=1):
        data_parts = []
        found = False
        data_part = {"rows": []}
        previous_row = []
        for row in self.ws.iter_rows(min_row=1, min_col=min_col, values_only=True):
            if found:
                if (not isinstance(row[0], float)) and (not isinstance(row[0], int)):
                    found = False
                    data_parts.append(data_part)
                    data_part = {"rows": []}
                else:
                    data_part["rows"].append(row)
            if self.match_pattern(row):
                found = True
                data_part.update({"date": self.parse_date(previous_row)})
            previous_row = row
        return data_parts

    def match_pattern(self, row):
        return row == self.pattern

    def parse_date(self, row):
        dates = re.findall('[0-9]?[0-9]\.[0-9][0-9]\.[0-9][0-9][0-9][0-9]', row[0])
        if len(dates) == 0:
            raise NoDateFound(row[0], self.file)
        if len(dates) == 1:
            return {"cnl": dates[0]}
        elif len(dates) == 2:
            return {"from": dates[0], "to": dates[1]}
        else:
            raise UnsupportedDateQuantity(row[0], self.file)

    def parse_type(self, type_cell: str):
        if "АО \"Поликлинический комплекс\"" in type_cell:
            return ["РГС ПРЯМОЙ ДОСТУП", "РГС Прямой доступ 665", 4001481]
        elif "АО \"Современные медицинские технологии\"" in type_cell:
            return ["ДС Росгосстрах КДЦ", "РГС Прямой доступ 0002/СК", 4013625]
        else:
            raise UnrecognisedType("Ошибка типа", f"Неподдреживаемый тип в файле {self.file}")

    def format_name(self, name):
        name_array = name.split(" ")
        if len(name_array) == 3:
            surname = name_array[0]
            first_name = name_array[1]
            sec_name = name_array[2]
        elif len(name_array) == 4:
            surname = name_array[0] + name_array[1]
            first_name = name_array[2]
            sec_name = name_array[3]
        else:
            raise UnsupportedNameLength("Ошибка имени",
                                        "Неподдерживаемый формат имени '{0}' в  в файле {1}"
                                        .format(name, self.file))
        return [surname, first_name, sec_name]


class RosgosstrahAttach(RosgosstrahReport):
    def __init__(self, file):
        super().__init__(file, ('№ п/п', 'ФИО', 'Пол', 'Дата рождения', 'Адрес проживания', 'Телефоны', 'Полис'))

    def get_data(self):
        data = []
        data_parts = self.find_data_parts(2)
        for data_part in data_parts:
            data += self.parse_data_part(data_part)
        return data

    def parse_data_part(self, data_part):
        data = []
        for row in data_part["rows"]:
            data.append(self.format_name(row[1]) +
                        [row[3],
                         row[6],
                         data_part["date"]["from"],
                         data_part["date"]["to"],
                         None] + self.get_type())
        return data

    def get_type(self):
        return self.parse_type(self.ws.cell(2, 7).value)


class RosgosstrahDetach(RosgosstrahReport):
    def __init__(self, file):
        super().__init__(file, ('№ п/п', 'ФИО', 'Пол', 'Дата рождения', 'Полис'))

    def get_data(self):
        data = []
        data_parts = self.find_data_parts(1)
        for data_part in data_parts:
            data += self.parse_data_part(data_part)
        return data

    def parse_data_part(self, data_part):
        data = []
        for row in data_part["rows"]:
            data.append(self.format_name(row[1]) +
                        [row[3],
                        row[4],
                         None,
                         None,
                         data_part["date"]["cnl"]] + self.get_type())
        return data

    def get_type(self):
        return self.parse_type(self.ws.cell(1, 5).value)


class RosgosstrahFabric:
    @staticmethod
    def create_report(file):
        report = RosgosstrahReport(file, None)
        if report.ws.cell(3, 1).value.lower() == "открепление":
            report = RosgosstrahDetach(file)
            return report
        elif report.ws.cell(4, 2).value.lower() == "прикрепление":
            report = RosgosstrahAttach(file)
            return report
        else:
            raise FileTypeWasNotDefined(file)


if __name__ == '__main__':
    attach = RosgosstrahAttach(
        "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Росгосстрах/09-04 ?? 2071-0.xls")
    attach.get_data()
    detach = RosgosstrahDetach(
        "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Росгосстрах/09-04 ???? 2071-0.xls"
    )
    detach.get_data()
