from openpyxl import load_workbook
from reporter.utils import open_xls_as_xlsx, open_csv_as_xlsx


class BaseInsurance:
    def __init__(self, files, encoding=None):
        self.files = files
        self.encoding = encoding
        self.file = None

    def create_attach_report(self):
        data = []
        for file in self.files:
            self.file = file
            data += self.get_attach(self.get_ws(file))
        return data

    def create_detach_report(self, files: list):
        data = []
        for file in files:
            self.file = file
            data += self.get_detach(self.get_ws(file))
        return data

    def get_ws(self, file):
        if file.split(".")[-1] in ["xls", "XLS"]:
            wb_tmp = open_xls_as_xlsx(file, self.encoding)
        elif file.split(".")[-1] == "csv":
            wb_tmp = open_csv_as_xlsx(file)
        else:
            wb_tmp = load_workbook(file)
        ws_name = wb_tmp.sheetnames[0]
        ws_tmp = wb_tmp[ws_name]
        return ws_tmp

    def get_attach(self, ws):
        pass

    def get_detach(self, ws):
        pass
