from openpyxl import load_workbook
from reporter.utils import open_xls_as_xlsx, open_csv_as_xlsx
from collections import namedtuple
from abc import ABC, abstractmethod


Target = namedtuple('Target',
                    [
                        'name',
                        'policy_number',
                        'date_start',
                        'date_end',
                        'program',
                        'birth_date',
                        'insurance',
                        'date_cancel'
                    ])


class Report(ABC):
    def __init__(self, file_path, encoding=None):
        self.file_path = file_path
        self.encoding = encoding
        self.ws = self.get_ws(file_path)
        self.allowed_types = ("xls", "XLS", "csv", "xlsx")
        self.xls_types = ("xls", "XLS")
        self.csv_types = ("csv", "CSV")

    @abstractmethod
    def get_data(self):
        """Абстрактный метод для получения данных"""

    def get_ws(self, file):
        if self.is_xls(file):
            wb_tmp = open_xls_as_xlsx(file, self.encoding)
        elif self.is_csv(file):
            wb_tmp = open_csv_as_xlsx(file)
        else:
            wb_tmp = load_workbook(file)
        ws_name = wb_tmp.sheetnames[0]
        ws_tmp = wb_tmp[ws_name]
        return ws_tmp

    @staticmethod
    def is_csv(file_name):
        return file_name.split(".")[-1].lower == "csv"

    @staticmethod
    def is_xls(file_name):
        return file_name.split(".")[-1].lower == "xls"

    def is_reportable(self, file_name):
        return file_name.split(".")[-1] in self.allowed_types

