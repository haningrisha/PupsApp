from reporter.config import header, codes, ids
from reporter.exceptions import UnsupportedNameLength, UnrecognisedType
from reporter.base import Report

target = ["ФИО", "Номер полиса", "Дата начала обслуживания", "Дата окончания обслуживания", "Программа обслуживания",
          "Дата рождения", "Страхователь", "Дата отмены"]
column_map_attach = {
    "Номер полиса": "POLICY", "Дата начала обслуживания": "DATE_FRM", "Дата окончания обслуживания": "DATE_TO",
    "Дата рождения": "DATE_BIRTH", "Дата отмены": "DATE_CNCL"
}
column_map_detach = {
    "Номер полиса": "POLICY",
    "Дата рождения": "DATE_BIRTH",
    "Дата окончания обслуживания": "DATE_CNCL"
}


class AlyansReport(Report):
    def __init__(self, file, column_map, id_type):
        super().__init__(file)
        self.column_map = column_map
        self.file_name = file.split("/")[-1]
        self.code = codes[self.get_number()]
        self.id = ids[id_type]

    def get_data(self):
        data = {}
        for row in self.ws.iter_rows(max_row=1):
            for i, cell in enumerate(row):
                column_name = cell.value
                if column_name in target:
                    column = [row[0].value for row in self.ws.iter_rows(min_row=2, min_col=i + 1, max_col=i + 1)]
                    data.update({self.column_map.get(column_name, column_name): column})
        names = data.pop("ФИО")
        data.update(self.format_names(names))
        formatted_data = [[cell] for cell in data.get(header[0])]
        for head in header[1:8]:
            column = data.get(head)
            for i, row in enumerate(formatted_data):
                if column is None:
                    row += [None]
                else:
                    row += [column[i]]
        formatted_data = [row + self.code + self.id for row in formatted_data]
        return formatted_data

    def format_names(self, names):
        surname = []
        first_name = []
        sec_name = []
        for i, name in enumerate(names):
            name_array = name.split(" ")
            if len(name_array) == 3:
                surname.append(name_array[0])
                first_name.append(name_array[1])
                sec_name.append(name_array[2])
            elif len(name_array) == 4:
                surname.append(name_array[0] + name_array[1])
                first_name.append(name_array[2])
                sec_name.append(name_array[3])
            else:
                raise UnsupportedNameLength("Ошибка имени",
                                            "Неподдерживаемый формат имени '{0}' в {1} строке  в файле {2}"
                                            .format(name, i + 1, self.file))
        return {"SURNAME": surname, "FIRST_NAME": first_name, "SEC_NAME": sec_name}

    def get_number(self):
        file_name = self.file.split("/")[-1]
        number = file_name.split("_")
        if len(number) < 3:
            raise UnrecognisedType("Нераспознанный тип", "Имя файла: {0}".format(file_name))
        number = number[2]
        if number == '1492':
            return "alyans1"
        elif number == '10062':
            return "alyans2"
        else:
            raise UnrecognisedType("Нераспознанный тип", "Найденный тип: {0}\n"
                                                         "Имя файла: {1}".format(number, self.file))


class AlyansAttach(AlyansReport):
    def __init__(self, file):
        super().__init__(file, column_map_attach, "alyans_attach")


class AlyansDetach(AlyansReport):
    def __init__(self, file):
        super().__init__(file, column_map_detach, "alyans_detach")


if __name__ == '__main__':
    attach = AlyansAttach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Альянс Прикреп/??_2021.02.03_1492_????????????.csv")
    detach = AlyansDetach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Альянс Откреп/??_2021.02.08_10062_???????????.csv")
    attach.get_data()
    detach.get_data()
