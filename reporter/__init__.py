from reporter.config import header
from reporter import renessans, alyans, ingossrah, sogaz
from openpyxl import Workbook
import os
from reporter.utils import fit_columns_width


class ReportChain:
    def __init__(self, func, files):
        self.func = func
        self.files = files


def create_reports(report_chains: [ReportChain], save_directory, filename="test"):
    data = [header]
    wb = Workbook()
    ws = wb.active
    for report_chain in report_chains:
        data += report_chain.func(report_chain.files)
    for r in data:
        ws.append(r)
    fit_columns_width(ws)
    wb.save(os.path.join(save_directory, filename + ".xlsx"))


if __name__ == '__main__':
    reports = [
        ReportChain(renessans.create_detach_report, [
            "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/002_1060_001???35890520?_3_????_01_02_2021.xls",
            "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/002_1060_001???35719520?_19_????_03_02_2021.xlsx"
        ]),
        ReportChain(alyans.create_detach_report, [
            "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/??_2021.02.08_1492_???????????.csv",
            "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/??_2021.02.08_10062_???????????.csv"
        ]),
        ReportChain(ingossrah.create_detach_report, [
            "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test/_04_02_2021_11_26_38_978666_KDP4.XLS"
        ])
    ]
    create_reports(reports, "/Users/grigorijhanin/Documents/Работа пупс/PupsApp/test")
