from reporter.report import AbstractReport


class MaksReport(AbstractReport):
    def __init__(self, file):
        super().__init__(file, encoding="windows-1251")


class MaksAttach(MaksReport):
    def get_data(self):
        data = []
        for row in self.ws.iter_rows(min_row=2, values_only=True):
            codes = [["СУДЕБНЫЙ КОРПУС МАКС РК", "Судебный корпус МАКС РК", 4001477],
                     ["СУДЕБНЫЙ КОРПУС МАКС ПК", "Судебный корпус МАКС ПК", 4001477]]
            for code in codes:
                data += [[row[4], row[5], row[6], row[8], row[3], row[9], row[10], None] + code]
        return data


class MaksDetach(MaksReport):
    def get_data(self):
        data = []
        for row in self.ws.iter_rows(min_row=2, values_only=True):
            codes = [["СУДЕБНЫЙ КОРПУС МАКС РК", "Судебный корпус МАКС РК", 2],
                     ["СУДЕБНЫЙ КОРПУС МАКС ПК", "Судебный корпус МАКС ПК", 2]]
            for code in codes:
                data += [[row[4], row[5], row[6], row[8], row[3], None, None, row[10]] + code]
        return data


def get_maks(file, attach=False, detach=False):
    if MaksReport.is_reportable(file):
        if attach:
            return MaksAttach(file)
        elif detach:
            return MaksDetach(file)


if __name__ == '__main__':
    test_attach = MaksAttach(
        "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Макс/0501p210315_05322_210312.xls")
    test_detach = MaksDetach(
        "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/Макс/0501p210315_05322_210312.xls")
    test_attach.get_data()
    test_detach.get_data()
