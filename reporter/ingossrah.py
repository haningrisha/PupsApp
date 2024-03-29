from reporter.exceptions import UnrecognisedType
from reporter.config import header, codes, ids
from reporter.report import AbstractReport

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
column_map_attach_fullrisk = {
    "Полис": "POLICY",
    "Фамилия": "SURNAME",
    "Имя": "FIRST_NAME",
    "Отчество": "SEC_NAME",
    "Д.Рожд.": "DATE_BIRTH",
    "Дата прикрепления": "DATE_FRM",
    "Дата открепления": "DATE_TO"
}

column_map_attach_kdp4 = {
    "Фамилия": "SURNAME",
    "Имя": "FIRST_NAME",
    "Отчество": "SEC_NAME",
    "Дата рождения": "DATE_BIRTH",
    "№ полиса": "POLICY",
    "Начало обслуживания": "DATE_FRM",
    "Конец обслуживания": "DATE_TO"
}

column_map_detach_fullrisk = {
    "Полис": "POLICY",
    "Фамилия": "SURNAME",
    "Имя": "FIRST_NAME",
    "Отчество": "SEC_NAME",
    "Д.Рожд.": "DATE_BIRTH",
    "Дата открепления": "DATE_CNCL"
}

column_map_detach_kdp4 = {
    "Фамилия": "SURNAME",
    "Имя": "FIRST_NAME",
    "Отчество": "SEC_NAME",
    "Дата рождения": "DATE_BIRTH",
    "№ полиса": "POLICY",
    "Конец обслуживания": "DATE_CNCL"
}


class IngosstrahReport(AbstractReport):
    def __init__(self, file, column_maps: dict, id_type):
        super().__init__(file)
        self.file_type = self.file.split(".")[0].split("_")[-1]
        self.column_map = column_maps.get(self.file_type)
        self.codes = codes["ingosstrah"]
        self.ids = ids[id_type]

    def get_data(self):
        if self.file_type == "FULLRISK":
            return self.get_data_fullrisk()
        elif self.file_type == "KDP4":
            return self.get_data_kdp4()
        else:
            raise UnrecognisedType(expression="Нераспознанный тип файла Ингосстрах",
                                   message="Найденный тип: {0}\n"
                                           "Имя файла: {1}".format(self.file_type, self.file.split("/")[-1]))

    def get_data_fullrisk(self):
        data = {}
        max_row = self.get_max_fullrisk()
        for row in self.ws.iter_rows(min_row=13, max_row=13):
            for cell in row:
                if cell.value in self.column_map.keys():
                    column = []
                    for column_row in self.ws.iter_rows(min_row=14, min_col=cell.column, max_col=cell.column,
                                                        max_row=max_row):
                        column.append(column_row[0].value)
                    data.update({self.column_map[cell.value]: column})
        return self.format_data(data)

    def get_data_kdp4(self):
        data = {}
        for row in self.ws.iter_rows(min_row=10, max_row=10):
            for cell in row:
                if cell.value in self.column_map.keys():
                    column = []
                    for column_row in self.ws.iter_rows(min_row=11, min_col=cell.column, max_col=cell.column):
                        column.append(column_row[0].value)
                    data.update({self.column_map[cell.value]: column})
        return self.format_data(data)

    def format_data(self, data):
        formatted_data = [[cell] for cell in data.get(header[0])]
        for head in header[1:8]:
            for i, row in enumerate(formatted_data):
                column = data.get(head)
                if column is not None:
                    row += [column[i]]
                else:
                    row += [None]
        formatted_data = [row + self.codes + self.ids for row in formatted_data]
        return formatted_data

    def get_max_fullrisk(self):
        max_row = self.ws.max_row
        for i, row in enumerate(self.ws.iter_rows(min_row=14)):
            row = [cell.value for cell in row]
            for y, cell in enumerate(row):
                if cell in ["", None]:
                    row[y] = True
                else:
                    row[y] = False
            if all(row):
                max_row = 13 + i
                break
        return max_row


class IngosstrahAttach(IngosstrahReport):
    def __init__(self, file):
        super().__init__(file, {
            "FULLRISK": column_map_attach_fullrisk,
            "KDP4": column_map_attach_kdp4
        },
                         "ingosstrah_attach")


class IngosstrahDetach(IngosstrahReport):
    def __init__(self, file):
        super().__init__(file, {
            "FULLRISK": column_map_detach_fullrisk,
            "KDP4": column_map_detach_kdp4
        },
                         "ingosstrah_detach")


def get_ingosstrah(file, attach=False, detach=False):
    if IngosstrahReport.is_reportable(file):
        if attach:
            return IngosstrahAttach(file)
        elif detach:
            return IngosstrahDetach(file)


if __name__ == '__main__':
    test_attach = IngosstrahAttach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Ингосстрах Прикреп/_04_02_2021_11_17_55_111872_FULLRISK.XLS")
    test_detach = IngosstrahDetach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Ингосстрах Откреп/_04_02_2021_11_26_38_978666_KDP4.XLS")
    test_attach.get_data()
    test_detach.get_data()

