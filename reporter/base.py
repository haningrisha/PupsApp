from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from abc import ABC, abstractmethod
from reporter.utils import open_xls_as_xlsx, open_csv_as_xlsx


class Report(ABC):
    def __init__(self, file, encoding=None):
        self.file = file
        self.encoding = encoding
        self.ws = self._get_ws(file)

    @abstractmethod
    def get_data(self):
        pass

    def _get_ws(self, file) -> Worksheet:
        if file.split(".")[-1] in ["xls", "XLS"]:
            wb_tmp = open_xls_as_xlsx(file, self.encoding)
        elif file.split(".")[-1] == "csv":
            wb_tmp = open_csv_as_xlsx(file)
        else:
            wb_tmp = load_workbook(file)
        ws_name = wb_tmp.sheetnames[0]
        ws_tmp = wb_tmp[ws_name]
        return ws_tmp

    @staticmethod
    def is_reportable(file):
        return file.split(".")[-1] in ["xls", "XLS", "csv", "xlsx"]

