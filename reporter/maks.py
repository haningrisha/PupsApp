from reporter.base import Report


class MaksReport(Report):
    def __init__(self, file):
        super().__init__(file, encoding="windows-1251")


class MaksAttach(MaksReport):
    def get_data(self):
        data = []
        for row in self.ws.iter_rows(min_row=2, values_only=True):
            if 'судебный' in row[1].lower().split(" "):
                codes = [["МАКС КДЦ Судебный департамент", "0034/СК Судебный департамент", 1102],
                         ["МАКС ПК Судебный департамент", "(МАКС) 101725/16-57937/330/2017 Судебный департамент", 1102]]
            else:
                codes = [["МАКС КДЦ", "0034/СК", 1102],
                         ["Макс стандарт", "(МАКС) 101725/16-57937/330/2017", 1102]]
            for code in codes:
                data += [[row[4], row[5], row[6], row[8], row[3], row[9], row[10], None] + code]
        return data


class MaksDetach(MaksReport):
    def get_data(self):
        data = []
        for row in self.ws.iter_rows(min_row=2, values_only=True):
            if 'судебный' in row[1].lower().split(" "):
                codes = [["МАКС КДЦ Судебный департамент", "0034/СК Судебный департамент", 2],
                         ["МАКС ПК Судебный департамент", "(МАКС) 101725/16-57937/330/2017 Судебный департамент", 2]]
            else:
                codes = [["МАКС КДЦ", "0034/СК", 1102],
                         ["Макс стандарт", "(МАКС) 101725/16-57937/330/2017", 1102]]
            for code in codes:
                data += [[row[4], row[5], row[6], row[8], row[3], None, None, row[10]] + code]
        return data


if __name__ == '__main__':
    attach = MaksAttach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Макс/0501p210315_05322_210312.xls")
    detach = MaksDetach("/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Макс/0501p210315_05322_210312.xls")
    attach.get_data()
    detach.get_data()
