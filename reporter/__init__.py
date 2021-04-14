from reporter.config import header
from openpyxl import Workbook
import os
from os import listdir
from os.path import isdir, isfile, join
from reporter.utils import fit_columns_width
from reporter.utils import open_xls_as_xlsx, open_csv_as_xlsx
from reporter.alyans import AlyansDetach, AlyansAttach
from reporter.renessans import RenessansAttach, RenessansDetach
from reporter.sogaz import SogazAttach, SogazDetach
from reporter.ingossrah import IngosstrahAttach, IngosstrahDetach
from reporter.maks import MaksAttach, MaksDetach


alyans_folders = ["альянс"]
renessans_folders = ["рен", "ренессанс"]
ingosstrah_folders = ["ингосстрах", "ингос"]
sogaz_folders = ["согаз"]
maks_folders = ["макс"]
detach_folders = ["откреп", "откр", "открепление", "открепления", "detach", "detachment", "detachments"]
attach_folders = ["прикреп", "прикр", "прикрепление", "прикрепления", "attach", "attachment", "attachments"]


class ReportChain:
    def __init__(self):
        self.reports = []

    def add_from_folder(self, folder):
        only_dirs = [f for f in listdir(folder) if isdir(join(folder, f))]
        [self.add_alyans(join(folder, f)) for f in only_dirs if f.lower() in alyans_folders]
        [self.add_renessans(join(folder, f)) for f in only_dirs if f.lower() in renessans_folders]
        [self.add_ingosstrah(join(folder, f)) for f in only_dirs if f.lower() in ingosstrah_folders]
        [self.add_sogaz(join(folder, f)) for f in only_dirs if f.lower() in sogaz_folders]
        [self.add_maks(join(folder, f)) for f in only_dirs if f.lower() in maks_folders]

    def add_alyans(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_alyans_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_alyans_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_renessans(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_renessans_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_renessans_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_ingosstrah(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_ingosstrah_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_ingosstrah_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_sogaz(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_sogaz_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_sogaz_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_maks(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_maks_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_maks_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def get_detach(self, folder):
        only_dirs = [f for f in listdir(folder) if isdir(join(folder, f))]
        return [join(folder, f) for f in only_dirs if f.lower() in detach_folders]

    def get_attach(self, folder):
        only_dirs = [f for f in listdir(folder) if isdir(join(folder, f))]
        return [join(folder, f) for f in only_dirs if f.lower() in attach_folders]

    def add_alyans_attach(self, file):
        if AlyansAttach.is_reportable(file):
            self.reports.append(AlyansAttach(file))

    def add_alyans_detach(self, file):
        if AlyansDetach.is_reportable(file):
            self.reports.append(AlyansDetach(file))

    def add_sogaz_attach(self, file):
        if SogazAttach.is_reportable(file):
            self.reports.append(SogazAttach(file))

    def add_sogaz_detach(self, file):
        if SogazDetach.is_reportable(file):
            self.reports.append(SogazDetach(file))

    def add_ingosstrah_attach(self, file):
        if IngosstrahAttach.is_reportable(file):
            self.reports.append(IngosstrahAttach(file))

    def add_ingosstrah_detach(self, file):
        if IngosstrahDetach.is_reportable(file):
            self.reports.append(IngosstrahDetach(file))

    def add_renessans_attach(self, file):
        if RenessansAttach.is_reportable(file):
            self.reports.append(RenessansAttach(file))

    def add_renessans_detach(self, file):
        if RenessansDetach.is_reportable(file):
            self.reports.append(RenessansDetach(file))

    def add_maks_attach(self, file):
        if MaksAttach.is_reportable(file):
            self.reports.append(MaksAttach(file))

    def add_maks_detach(self, file):
        if MaksDetach.is_reportable(file):
            self.reports.append(MaksDetach(file))


class ReportGenerator:
    def __init__(self, report_chain, save_directory, filename):
        self.report_chain = report_chain
        self.save_directory = save_directory
        self.filename = filename

    def generate(self):
        data = [header]
        wb = Workbook()
        ws = wb.active
        for report in self.report_chain.reports:
            data += report.get_data()
        for r in data:
            ws.append(r)
        fit_columns_width(ws)
        wb.save(os.path.join(self.save_directory, self.filename + ".xlsx"))
        return os.path.join(self.save_directory, self.filename + ".xlsx")


if __name__ == '__main__':
    pass
